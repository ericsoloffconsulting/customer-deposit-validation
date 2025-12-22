/**
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 * @NModuleScope SameAccount
 * @NAmdConfig /SuiteScripts/ericsoloffconsulting/JsLibraryConfig.json
 * 
 * Unapplied Customer Deposit Research - Kitchen Works
 * 
 * Purpose: Displays customer deposits that are not fully applied and trace back
 * to sales orders with line items within the Kitchen Retail Sales account (338)
 * 
 * This report helps identify outstanding deposits that may need to be applied,
 * refunded, or researched for the Kitchen Works department.
 */
define(['N/ui/serverWidget', 'N/query', 'N/log', 'N/runtime', 'N/url', 'N/record', 'N/search', '../ericsoloffconsulting/lib/claude_api_library'],
    /**
     * @param {serverWidget} serverWidget
     * @param {query} query
     * @param {log} log
     * @param {runtime} runtime
     * @param {url} url
     * @param {record} record
     * @param {search} search
     * @param {claudeAPI} claudeAPI
     */
    function (serverWidget, query, log, runtime, url, record, search, claudeAPI) {

        /**
         * Handles GET and POST requests to the Suitelet
         * @param {Object} context - NetSuite context object containing request/response
         */
        function onRequest(context) {
            if (context.request.method === 'GET') {
                handleGet(context);
            } else {
                // POST requests handle Next Step updates
                handlePost(context);
            }
        }

        /**
         * Handles POST requests for updating Next Step field or SO-Invoice comparison
         * @param {Object} context
         */
        function handlePost(context) {
            var response = context.response;
            try {
                var body = JSON.parse(context.request.body);
                
                // Check if this is an AI Analysis request
                if (body.action === 'aiAnalysis') {
                    var creditMemoId = body.creditMemoId;
                    var comparisonData = body.comparisonData; // Include SO/INV comparison findings
                    log.debug('AI Analysis Request', 'CM ID: ' + creditMemoId);
                    
                    var aiResult = analyzeTransactionLifecycleWithAI(creditMemoId, comparisonData);
                    
                    response.setHeader({ name: 'Content-Type', value: 'application/json' });
                    response.write(JSON.stringify({ success: true, data: aiResult }));
                    return;
                }
                
                // Check if this is a Load AI Analysis request
                if (body.action === 'loadAIAnalysis') {
                    var creditMemoId = body.creditMemoId;
                    log.debug('Load AI Analysis Request', 'CM ID: ' + creditMemoId);
                    
                    var aiResult = loadSavedAIAnalysis(creditMemoId);
                    
                    response.setHeader({ name: 'Content-Type', value: 'application/json' });
                    response.write(JSON.stringify({ success: true, aiAnalysis: aiResult }));
                    return;
                }
                
                // Check if this is an SO-Invoice comparison request
                if (body.action === 'soInvoiceComparison') {
                    var creditMemoId = body.creditMemoId;
                    log.debug('SO-Invoice Comparison Request', 'CM ID: ' + creditMemoId);
                    
                    var result = analyzeCreditMemoOverpayment(creditMemoId);
                    
                    try {
                        var jsonString = JSON.stringify({ success: true, data: result });
                        var sizeKB = (jsonString.length / 1024).toFixed(2);
                        log.debug('Response Size', sizeKB + ' KB (' + jsonString.length + ' characters)');
                        
                        response.setHeader({ name: 'Content-Type', value: 'application/json' });
                        response.write(jsonString);
                    } catch (stringifyError) {
                        log.error('JSON Stringify Error', {
                            error: stringifyError.message,
                            stack: stringifyError.stack,
                            resultKeys: Object.keys(result)
                        });
                        response.setHeader({ name: 'Content-Type', value: 'application/json' });
                        response.write(JSON.stringify({ 
                            success: false, 
                            error: 'Failed to serialize response: ' + stringifyError.message 
                        }));
                    }
                    return;
                }
                
                // Handle Cross-SO Deposit Analysis requests
                if (body.action === 'crossSOAnalysis') {
                    var creditMemoId = body.creditMemoId;
                    log.debug('Cross-SO Analysis Request', 'CM ID: ' + creditMemoId);
                    
                    var result = analyzeCrossSODeposits(creditMemoId);
                    
                    response.setHeader({ name: 'Content-Type', value: 'application/json' });
                    response.write(JSON.stringify({ success: true, data: result }));
                    return;
                }
                
                // Handle Overpayment Summary requests
                if (body.action === 'overpaymentSummary') {
                    var creditMemoId = body.creditMemoId;
                    log.debug('Overpayment Summary Request', 'CM ID: ' + creditMemoId);
                    
                    var result = getOverpaymentSummary(creditMemoId);
                    
                    response.setHeader({ name: 'Content-Type', value: 'application/json' });
                    response.write(JSON.stringify({ success: true, data: result }));
                    return;
                }
                
                // Handle Sales Order Totals requests
                if (body.action === 'salesOrderTotals') {
                    var creditMemoId = body.creditMemoId;
                    log.debug('Sales Order Totals Request', 'CM ID: ' + creditMemoId);
                    
                    var result = getCustomerSalesOrdersSummary(creditMemoId);
                    
                    response.setHeader({ name: 'Content-Type', value: 'application/json' });
                    response.write(JSON.stringify({ success: true, data: result }));
                    return;
                }
                
                // Handle deposit update requests
                var depositId = body.depositId;
                var nextStep = body.nextStep;
                var notes = body.notes || '';
                var updateNotesOnly = body.updateNotesOnly || false;

                log.debug('POST Request', 'Updating deposit ' + depositId + ' with nextStep: ' + nextStep + ', notes: ' + notes + ', notesOnly: ' + updateNotesOnly);

                // Load and update the customer deposit record
                var depositRecord = record.load({
                    type: record.Type.CUSTOMER_DEPOSIT,
                    id: depositId,
                    isDynamic: false
                });

                // Only update nextStep if not a notes-only update
                if (!updateNotesOnly && nextStep) {
                    depositRecord.setValue({
                        fieldId: 'custbody_cd_reconciliation_next_step',
                        value: nextStep
                    });
                }

                depositRecord.setValue({
                    fieldId: 'custbody_cd_reconciliation_notes',
                    value: notes
                });

                depositRecord.save();

                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ success: true, message: 'Updated successfully' }));
            } catch (e) {
                log.error('POST Error', e.message);
                response.setHeader({ name: 'Content-Type', value: 'application/json' });
                response.write(JSON.stringify({ success: false, message: e.message }));
            }
        }

        /**
         * Handles GET requests
         * @param {Object} context
         */
        function handleGet(context) {
            var request = context.request;
            var response = context.response;

            log.debug('GET Request', 'Parameters: ' + JSON.stringify(request.parameters));

            // Create NetSuite form
            var form = serverWidget.createForm({
                title: 'Customer Deposit Research - Kitchen Works'
            });

            try {
                // Build and add HTML content
                var htmlContent = buildPageHTML(request.parameters);

                var htmlField = form.addField({
                    id: 'custpage_html_content',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: 'Content'
                });

                htmlField.defaultValue = htmlContent;

            } catch (e) {
                log.error('Error Building Form', {
                    error: e.message,
                    stack: e.stack
                });

                var errorField = form.addField({
                    id: 'custpage_error',
                    type: serverWidget.FieldType.INLINEHTML,
                    label: 'Error'
                });
                errorField.defaultValue = '<p style="color:red;">Error loading portal: ' + escapeHtml(e.message) + '</p>';
            }

            context.response.writePage(form);
        }

        /**
         * Builds the main page HTML
         * @param {Object} params - URL parameters
         * @returns {string} HTML content
         */
        function buildPageHTML(params) {
            try {
                log.debug('buildPageHTML Start', 'Params: ' + JSON.stringify(params));
                
                var scriptUrl = url.resolveScript({
                    scriptId: runtime.getCurrentScript().id,
                    deploymentId: runtime.getCurrentScript().deploymentId,
                    returnExternalUrl: false
                });

                // Check if data should be loaded
                var shouldLoadData = params.loadData === 'T';
                log.debug('Load Data Decision', 'shouldLoadData: ' + shouldLoadData);

                // Get balance as of date from params, default to 12/31/2025
                var balanceAsOf = (params.balanceAsOf && params.balanceAsOf.trim()) ? params.balanceAsOf.trim() : '2025-12-31';
                log.debug('Balance As Of', balanceAsOf);

                var html = '';

            // Loading spinner overlay - FIRST with inline styles so it renders immediately
            html += '<div id="loadingOverlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(255,255,255,0.95);display:flex;flex-direction:column;justify-content:center;align-items:center;z-index:9999;display:none;">';
            html += '<div style="width:50px;height:50px;border:4px solid #e0e0e0;border-top:4px solid #4CAF50;border-radius:50%;animation:spin 1s linear infinite;"></div>';
            html += '<div style="margin-top:15px;font-size:16px;color:#333;font-weight:500;">Loading data...</div>';
            html += '<style>@keyframes spin{0%{transform:rotate(0deg)}100%{transform:rotate(360deg)}}</style>';
            html += '</div>';

            // If data not requested, show initial page with Load button
            if (!shouldLoadData) {
                var separator = scriptUrl.indexOf('?') > -1 ? '&' : '?';
                var loadDataUrl = scriptUrl + separator + 'loadData=T&balanceAsOf=' + balanceAsOf;
                
                html += '<style>' + getStyles() + '</style>';
                html += '<div class="portal-container">';
                html += '<div style="text-align:center;padding:80px 20px;">';
                html += '<h1 style="color:#1a237e;font-size:32px;margin-bottom:20px;">Unapplied Customer Deposit Research - Kitchen Works</h1>';
                html += '<p style="color:#666;font-size:16px;margin-bottom:30px;">Click the button below to load customer deposit data</p>';
                html += '<button onclick="loadCustomerData()" id="loadDataBtn" style="padding:15px 40px;font-size:18px;font-weight:bold;color:#fff;background:#4CAF50;border:none;border-radius:6px;cursor:pointer;box-shadow:0 2px 8px rgba(76,175,80,0.3);transition:all 0.2s;">Load Customer Deposit Data</button>';
                html += '</div>';
                html += '</div>';
                html += '<script>';
                html += 'function loadCustomerData() {';
                html += '  var btn = document.getElementById("loadDataBtn");';
                html += '  btn.disabled = true;';
                html += '  btn.textContent = "Loading...";';
                html += '  var overlay = document.getElementById("loadingOverlay");';
                html += '  if (overlay) overlay.style.display = "flex";';
                html += '  window.location.href = "' + loadDataUrl + '";';
                html += '}';
                html += '</script>';
                return html;
            }

            // Data requested - load everything
            log.debug('Data Loading', 'Starting to load deposit data for ' + balanceAsOf);
            var depositResult = searchUnappliedDeposits(balanceAsOf);
            var deposits = depositResult.deposits;
            var depositsIsTruncated = depositResult.isTruncated;
            log.debug('Deposits Loaded', 'Count: ' + deposits.length);

            // Calculate totals - use aggregate if truncated, otherwise sum from line items
            var totalDeposits = depositResult.actualCount;
            var totalDepositAmount = 0;
            var totalAppliedAmount = 0;
            var totalUnappliedAmount = 0;

            if (depositsIsTruncated) {
                // Use pre-calculated aggregate total for unapplied amount
                totalUnappliedAmount = depositResult.actualTotalUnapplied;
                // Sum deposit and applied amounts from displayed records (best effort)
                for (var i = 0; i < deposits.length; i++) {
                    totalDepositAmount += deposits[i].depositAmount || 0;
                    totalAppliedAmount += deposits[i].amountApplied || 0;
                }
            } else {
                // Calculate all totals from line items
                for (var i = 0; i < deposits.length; i++) {
                    totalDepositAmount += deposits[i].depositAmount || 0;
                    totalAppliedAmount += deposits[i].amountApplied || 0;
                    totalUnappliedAmount += deposits[i].amountUnapplied || 0;
                }
            }

            // Get unapplied credit memo data (needed for summary)
            var creditMemos = searchUnappliedCreditMemos(balanceAsOf);

            // Get AI analysis lookup for icon display (only if we have CMs)
            var aiAnalysisLookup = {};
            if (creditMemos.length > 0) {
                var cmIds = creditMemos.map(function(cm) { return cm.cmId; });
                aiAnalysisLookup = getAIAnalysisLookup(cmIds);
            }

            // Calculate CM totals
            var totalCMs = creditMemos.length;
            var totalCMAmount = 0;
            var totalCMApplied = 0;
            var totalCMUnapplied = 0;

            for (var j = 0; j < creditMemos.length; j++) {
                totalCMAmount += creditMemos[j].cmAmount || 0;
                totalCMApplied += creditMemos[j].amountApplied || 0;
                totalCMUnapplied += creditMemos[j].amountUnapplied || 0;
            }

            // Add styles
            html += '<style>' + getStyles() + '</style>';

            // Main container
            html += '<div class="portal-container">';

            // Customer Balance Tooltip
            html += '<div id="customerBalanceTooltip" class="customer-balance-tooltip">';
            html += '<div class="customer-balance-tooltip-header">Customer Financial Status</div>';
            html += '<div id="customerBalanceContent" class="customer-balance-tooltip-content"></div>';
            html += '</div>';

            // Action Popup (shared by Next Step and Notes)
            html += '<div id="actionPopup" class="action-popup">';
            html += '<div class="action-popup-header"><span id="actionPopupTitle">Select Action</span><span class="action-popup-close" onclick="hideActionPopup()">&times;</span></div>';
            html += '<div id="actionPopupContent" class="action-popup-content"></div>';
            html += '</div>';
            html += '<div id="actionPopupOverlay" class="action-popup-overlay" onclick="hideActionPopup()"></div>';

            // Unified Explain Modal with Tabs
            html += '<div id="explainModal" class="comparison-modal">';
            html += '<div class="comparison-modal-content explain-modal-content">';
            html += '<div class="comparison-modal-header">';
            html += '<span class="comparison-modal-title">Transaction Analysis</span>';
            html += '<span class="comparison-modal-close" onclick="hideExplainModal()">&times;</span>';
            html += '</div>';
            
            // Customer Information Section (populated by JavaScript)
            html += '<div id="customerInfoSection" class="customer-info-section" style="padding:15px;background:#f8f9fa;border-bottom:2px solid #dee2e6;display:none;">';
            html += '</div>';
            
            // Tab navigation
            html += '<div class="tab-navigation">';
            html += '<button type="button" class="tab-button active" onclick="switchExplainTab(1)" id="tab-btn-1">SO↔INV Comparison</button>';
            html += '<button type="button" class="tab-button" onclick="switchExplainTab(2)" id="tab-btn-2">CD Cross-SO Analysis</button>';
            html += '<button type="button" class="tab-button" onclick="switchExplainTab(3)" id="tab-btn-3">Overpayment Summary</button>';
            html += '<button type="button" class="tab-button" onclick="switchExplainTab(4)" id="tab-btn-4">AI Generated SO Price Changes</button>';
            html += '<button type="button" class="tab-button" onclick="switchExplainTab(5)" id="tab-btn-5">Sales Order Totals (All)</button>';
            html += '</div>';
            
            // Tab content containers
            html += '<div id="explainModalBody" class="comparison-modal-body">';
            
            // Tab 1: SO to Invoice Comparison
            html += '<div id="tab-content-1" class="tab-content active">';
            html += '<div class="comparison-loading">Loading SO↔INV comparison...</div>';
            html += '</div>';
            
            // Tab 2: Cross-SO Analysis
            html += '<div id="tab-content-2" class="tab-content">';
            html += '<div class="comparison-loading">Loading Cross-SO analysis...</div>';
            html += '</div>';
            
            // Tab 3: Overpayment Summary
            html += '<div id="tab-content-3" class="tab-content">';
            html += '<div class="comparison-loading">Loading Overpayment Summary...</div>';
            html += '</div>';
            
            // Tab 4: AI Generated SO Price Changes
            html += '<div id="tab-content-4" class="tab-content">';
            html += '<div class="comparison-loading">Loading AI Generated SO Price Changes...</div>';
            html += '</div>';
            
            // Tab 5: Sales Order Totals (All)
            html += '<div id="tab-content-5" class="tab-content">';
            html += '<div class="comparison-loading">Loading Sales Order Totals...</div>';
            html += '</div>';
            
            html += '</div>';
            html += '</div>';
            html += '</div>';
            html += '<div id="explainModalOverlay" class="comparison-modal-overlay" onclick="hideExplainModal()"></div>';

            // Balance As Of Date Control
            html += '<div class="balance-as-of-section">';
            html += '<label class="balance-as-of-label">Balance As Of:</label>';
            html += '<input type="date" id="balanceAsOfDate" class="balance-as-of-input" value="' + balanceAsOf + '">';
            html += '<button type="button" id="loadResultsBtn" class="load-results-btn">Load Results</button>';
            html += '</div>';

            // Combined Summary Row - both sections side by side
            html += '<div class="summary-row">';

            // Customer Deposits Summary Section
            html += '<div class="summary-section">';
            html += '<h2 class="summary-title">True Customer Deposits</h2>';
            html += '<div class="summary-totals-row">';
            html += '<div class="summary-total">';
            html += '<span class="summary-total-label">Total Unapplied Amount:</span>';
            html += '<span class="summary-total-amount">' + formatCurrency(totalUnappliedAmount) + '</span>';
            html += '</div>';
            html += '<div class="summary-total">';
            html += '<div class="prior-period-header">';
            html += '<span class="summary-total-label">Received Prior To:</span>';
            html += '<input type="date" id="priorPeriodDate" class="prior-period-date-input" value="2024-12-31">';
            html += '</div>';
            html += '<span class="summary-total-amount" id="priorPeriodAmount">$0.00</span>';
            html += '<span class="prior-period-helper">Calculate Customer Deposits 1+ Years Old for Tax Implications</span>';
            html += '</div>';
            html += '</div>';
            html += '</div>';

            // Credit Memo Overpayments Summary Section
            html += '<div class="summary-section">';
            html += '<h2 class="summary-title">Credit Memo Overpayments</h2>';
            html += '<div class="summary-totals-row">';
            html += '<div class="summary-total">';
            html += '<span class="summary-total-label">Total Unapplied Amount:</span>';
            html += '<span class="summary-total-amount">' + formatCurrency(totalCMUnapplied) + '</span>';
            html += '<span class="prior-period-helper">Credit Memos Converted from Customer Deposits Via Automated Process Began December 2025 and Continues Daily as Overpayment is Recognized</span>';
            html += '</div>';
            html += '<div class="summary-total">';
            html += '<div class="prior-period-header">';
            html += '<span class="summary-total-label">Overpayment Date Prior To:</span>';
            html += '<input type="date" id="cmPriorPeriodDate" class="prior-period-date-input" value="2024-12-31">';
            html += '</div>';
            html += '<span class="summary-total-amount" id="cmPriorPeriodAmount">$0.00</span>';
            html += '<span class="prior-period-helper">Overpayment Date is the Date of the SO\'s Final Fulfillment ("Actual Ship Date") or the Date the SO is Closed. Once the Sales Order\'s Final Invoice is Generated (SO Status: Billed) the CD is Converted and Booked Officially as an A/R Credit.</span>';
            html += '</div>';
            html += '</div>';
            html += '</div>';

            html += '</div>'; // Close summary-row

            // Data Section - Customer Deposits
            html += buildDataSection('deposits', 'True Customer Deposits', 
                'Customer deposits linked to Kitchen Retail Sales orders that have not been fully applied', 
                deposits, scriptUrl, depositsIsTruncated, totalDeposits);

            // Data Section - Credit Memos
            html += buildCreditMemoDataSection('creditmemos', 'Credit Memo Overpayments from Customer Deposits', 
                'Unapplied credit memos created from overpayment customer deposits linked to Kitchen Retail Sales orders', 
                creditMemos, scriptUrl, aiAnalysisLookup);

            html += '</div>'; // Close portal-container

            // Add SheetJS library for Excel export
            html += '<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>';

            // Add JavaScript
            html += '<script>' + getJavaScript(scriptUrl) + '</script>';

            log.debug('buildPageHTML Complete', 'HTML generated successfully');
            return html;
            
            } catch (e) {
                log.error('Error in buildPageHTML', {
                    error: e.message,
                    stack: e.stack,
                    params: JSON.stringify(params)
                });
                // Return error page
                var errorHtml = '<style>body{font-family:Arial,sans-serif;padding:40px;}</style>';
                errorHtml += '<div style="max-width:600px;margin:0 auto;">';
                errorHtml += '<h1 style="color:#d32f2f;">Error Loading Portal</h1>';
                errorHtml += '<p><strong>Error:</strong> ' + escapeHtml(e.message) + '</p>';
                errorHtml += '<p><strong>Details:</strong> Check execution log for details.</p>';
                errorHtml += '<pre style="background:#f5f5f5;padding:15px;overflow:auto;">' + escapeHtml(e.stack || 'No stack trace available') + '</pre>';
                errorHtml += '</div>';
                return errorHtml;
            }
        }

        /**
         * Builds a summary card
         * @param {string} title - Card title
         * @param {number} count - Number of records
         * @param {number} amount - Total amount
         * @returns {string} HTML for summary card
         */
        function buildSummaryCard(title, count, amount) {
            var html = '';
            html += '<div class="summary-card">';
            html += '<div class="summary-card-title">' + escapeHtml(title) + '</div>';
            html += '<div class="summary-card-count">' + count + ' record' + (count !== 1 ? 's' : '') + '</div>';
            html += '<div class="summary-card-amount">' + formatCurrency(amount) + '</div>';
            html += '</div>';
            return html;
        }

        /**
         * Builds a collapsible data section
         * @param {string} sectionId - Section identifier
         * @param {string} title - Section title
         * @param {string} description - Section description
         * @param {Array} data - Data array
         * @param {string} scriptUrl - Suitelet URL
         * @returns {string} HTML for data section
         */
        function buildDataSection(sectionId, title, description, data, scriptUrl, isTruncated, actualCount) {
            var displayedRecords = data.length;
            var totalRecords = isTruncated ? actualCount : displayedRecords;
            var countDisplay = isTruncated 
                ? 'Displaying ' + displayedRecords.toLocaleString() + ' of ' + totalRecords.toLocaleString() + ' records'
                : totalRecords.toLocaleString();
            
            var html = '';
            html += '<div class="search-section" id="section-' + sectionId + '">';
            html += '<div class="search-title collapsible" data-section-id="' + sectionId + '">';
            html += '<span>' + escapeHtml(title) + ' (' + countDisplay + ')' + (isTruncated ? ' <span style="color: #4CAF50; font-size: 11px;">⚠ Totals calculated from all records</span>' : '') + '</span>';
            html += '<span class="toggle-icon" id="toggle-' + sectionId + '">−</span>';
            html += '</div>';
            html += '<div class="search-content" id="content-' + sectionId + '">';
            html += '<div class="search-count">' + escapeHtml(description) + '</div>';
            
            if (totalRecords === 0) {
                html += '<p class="no-results">No unapplied customer deposits found for Kitchen Works orders.</p>';
            } else {
                html += '<div class="search-box-container">';
                html += '<div class="search-row">';
                html += '<input type="text" id="searchBox-' + sectionId + '" class="search-box" placeholder="Search this table..." onkeyup="filterTable(\'' + sectionId + '\')">'; 
                html += '<button type="button" class="export-btn" onclick="exportToExcel(\'' + sectionId + '\')">📥 Export to Excel</button>';
                html += '</div>';
                html += '<span class="search-results-count" id="searchCount-' + sectionId + '"></span>';
                html += '</div>';
                html += buildDepositTable(data, scriptUrl, sectionId);
            }
            
            html += '</div>';
            html += '</div>';
            return html;
        }

        /**
         * Builds the deposit data table
         * @param {Array} deposits - Deposit data
         * @param {string} scriptUrl - Suitelet URL
         * @param {string} sectionId - Section identifier
         * @returns {string} HTML table
         */
        function buildDepositTable(deposits, scriptUrl, sectionId) {
            var html = '';

            html += '<div class="table-container">';
            html += '<table class="data-table" id="table-' + sectionId + '">';
            html += '<thead>';
            html += '<tr>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 0)">Next Step</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 1)">Notes</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 2)">Deposit #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 3)" title="Received before cutoff date">⏰</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 4)">Deposit Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 5)">Customer</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 6)">Deposit Amount</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 7)">Amount Applied</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 8)">Amount Unapplied</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 9)">Status</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 10)">Sales Order #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 11)">SO Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 12)">SO Status</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 13)">Selling Location</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 14)">Sales Rep</th>';
            html += '</tr>';
            html += '</thead>';
            html += '<tbody>';

            for (var i = 0; i < deposits.length; i++) {
                var dep = deposits[i];
                var rowClass = (i % 2 === 0) ? 'even-row' : 'odd-row';
                var customerBalanceData = 'data-customer-balance="' + (dep.customerBalance || 0) + '" data-customer-deposit-balance="' + (dep.customerDepositBalance || 0) + '" data-customer-unbilled-orders="' + (dep.customerUnbilledOrders || 0) + '" data-customer-name="' + escapeHtml(dep.customerName || '') + '" data-deposit-id="' + dep.depositId + '" data-deposit-number="' + escapeHtml(dep.depositNumber || '') + '" data-so-number="' + escapeHtml(dep.soNumber || '') + '" data-unapplied-amount="' + (dep.amountUnapplied || 0) + '" data-next-step="' + escapeHtml(dep.nextStep || '') + '" data-reconciliation-notes="' + escapeHtml(dep.reconciliationNotes || '') + '"';

                html += '<tr class="' + rowClass + ' customer-balance-row" id="dep-row-' + dep.depositId + '" ' + customerBalanceData + ' onmouseenter="showCustomerBalanceTooltip(this);" onmouseleave="hideCustomerBalanceTooltip();">';

                // Next Step - span trigger with popup
                var nextStepValue = dep.nextStep || '';
                var nextStepIcon = getNextStepIcon(nextStepValue);
                var displayIcon = nextStepValue ? nextStepIcon : '+';
                html += '<td class="next-step-cell">';
                html += '<span class="next-step-trigger" id="nextstep-' + dep.depositId + '" data-deposit-id="' + dep.depositId + '" data-current-value="' + nextStepValue + '" title="' + (nextStepValue ? 'Change Next Step' : 'Set Next Step') + '" onclick="showNextStepPopup(this)">' + displayIcon + '</span>';
                html += '</td>';

                // Notes - clickable icon
                var notesValue = dep.reconciliationNotes || '';
                var notesIcon = notesValue ? 'ℹ️' : '+';
                html += '<td class="next-step-cell">';
                html += '<span class="notes-trigger" id="notes-' + dep.depositId + '" data-deposit-id="' + dep.depositId + '" data-current-notes="' + escapeHtml(notesValue) + '" title="' + (notesValue ? 'View/Edit Notes' : 'Add Notes') + '" onclick="showNotesPopup(this)">' + notesIcon + '</span>';
                html += '</td>';

                // Deposit # with link
                html += '<td><a href="/app/accounting/transactions/custdep.nl?id=' + dep.depositId + '" target="_blank">' + escapeHtml(dep.depositNumber) + '</a></td>';

                // Aged icon (placeholder - will be updated by client-side script)
                html += '<td class="aged-icon-cell" data-date="' + (dep.depositDate || '') + '"></td>';

                // Deposit Date
                html += '<td data-date="' + (dep.depositDate || '') + '">' + formatDate(dep.depositDate) + '</td>';

                // Customer
                html += '<td>' + escapeHtml(dep.customerName || '-') + '</td>';

                // Deposit Amount
                html += '<td class="amount">' + formatCurrency(dep.depositAmount) + '</td>';

                // Amount Applied
                html += '<td class="amount">' + formatCurrency(dep.amountApplied) + '</td>';

                // Amount Unapplied
                html += '<td class="amount unapplied">' + formatCurrency(dep.amountUnapplied) + '</td>';

                // Status
                html += '<td>' + escapeHtml(translateStatus(dep.depositStatus)) + '</td>';

                // Sales Order # with link
                if (dep.soId) {
                    html += '<td><a href="/app/accounting/transactions/salesord.nl?id=' + dep.soId + '" target="_blank">' + escapeHtml(dep.soNumber) + '</a></td>';
                } else {
                    html += '<td>-</td>';
                }

                // SO Date
                html += '<td data-date="' + (dep.soDate || '') + '">' + formatDate(dep.soDate) + '</td>';

                // SO Status
                html += '<td>' + escapeHtml(translateSOStatus(dep.soStatus)) + '</td>';

                // Department
                html += '<td>' + escapeHtml(dep.soDepartment || '-') + '</td>';

                // Sales Rep
                html += '<td>' + escapeHtml(dep.salesrepName || '-') + '</td>';

                html += '</tr>';
            }

            html += '</tbody>';
            html += '</table>';
            html += '</div>';

            return html;
        }

        /**
         * Builds a collapsible data section for Credit Memos
         * @param {string} sectionId - Section identifier
         * @param {string} title - Section title
         * @param {string} description - Section description
         * @param {Array} data - Data array
         * @param {string} scriptUrl - Suitelet URL
         * @param {Object} aiAnalysisLookup - Lookup object for CMs with AI analysis
         * @returns {string} HTML for data section
         */
        function buildCreditMemoDataSection(sectionId, title, description, data, scriptUrl, aiAnalysisLookup) {
            var totalRecords = data.length;
            
            var html = '';
            html += '<div class="search-section" id="section-' + sectionId + '">';
            html += '<div class="search-title collapsible" data-section-id="' + sectionId + '">';
            html += '<span>' + escapeHtml(title) + ' (' + totalRecords + ')</span>';
            html += '<span class="toggle-icon" id="toggle-' + sectionId + '">−</span>';
            html += '</div>';
            html += '<div class="search-content" id="content-' + sectionId + '">';
            html += '<div class="search-count">' + escapeHtml(description) + '</div>';
            
            if (totalRecords === 0) {
                html += '<p class="no-results">No unapplied credit memo overpayments found for Kitchen Works orders.</p>';
            } else {
                html += '<div class="search-box-container">';
                html += '<div class="search-row">';
                html += '<input type="text" id="searchBox-' + sectionId + '" class="search-box" placeholder="Search this table..." onkeyup="filterTable(\'' + sectionId + '\')">';
                html += '<button type="button" class="export-btn" onclick="exportToExcel(\'' + sectionId + '\')">📥 Export to Excel</button>';
                html += '</div>';
                html += '<span class="search-results-count" id="searchCount-' + sectionId + '"></span>';
                html += '</div>';
                html += buildCreditMemoTable(data, scriptUrl, sectionId, aiAnalysisLookup);
            }
            
            html += '</div>';
            html += '</div>';
            return html;
        }

        /**
         * Builds the credit memo data table
         * @param {Array} creditMemos - Credit memo data
         * @param {string} scriptUrl - Suitelet URL
         * @param {string} sectionId - Section identifier
         * @param {Object} aiAnalysisLookup - Lookup object for CMs with AI analysis
         * @returns {string} HTML table
         */
        function buildCreditMemoTable(creditMemos, scriptUrl, sectionId, aiAnalysisLookup) {
            var html = '';

            html += '<div class="table-container">';
            html += '<table class="data-table" id="table-' + sectionId + '">';
            html += '<thead>';
            html += '<tr>';
            html += '<th>Compare</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 1)">Credit Memo #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 2)">CM Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 3)">Customer</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 4)">CM Amount</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 5)">Amount Applied</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 6)">Amount Unapplied</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 7)">Status</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 8)">Linked CD #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 9)">Overpayment Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 10)">CD Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 11)">Sales Order #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 12)">SO Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 13)">Selling Location</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 14)">Unbilled Orders</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 15)">Deposit Balance</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 16)">A/R Balance</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 17)">Sales Rep</th>';
            html += '</tr>';
            html += '</thead>';
            html += '<tbody>';

            for (var i = 0; i < creditMemos.length; i++) {
                var cm = creditMemos[i];
                var rowClass = (i % 2 === 0) ? 'even-row' : 'odd-row';

                html += '<tr class="' + rowClass + '" id="cm-row-' + cm.cmId + '">';

                // Unified EXPLAIN button
                var hasAI = aiAnalysisLookup && aiAnalysisLookup[cm.cmId];
                html += '<td class="action-btn-cell">';
                html += '<button type="button" class="explain-btn" onclick="showExplainModal(' + cm.cmId + ', \'' + escapeHtml(cm.customerName).replace(/'/g, '\\\'') + '\', \'' + escapeHtml(cm.cmNumber) + '\', ' + cm.cmAmount + ', \'' + escapeHtml(cm.linkedCD || '') + '\', \'' + escapeHtml(cm.salesOrderNumber || '') + '\')" title="Explain Transaction: SO/INV Comparison, Cross-SO Analysis, AI Insights">EXPLAIN</button>';
                if (hasAI) {
                    html += '<br><span class="ai-badge" title="AI analysis available">AI</span>';
                }
                html += '</td>';

                // Credit Memo # with link
                html += '<td><a href="/app/accounting/transactions/custcred.nl?id=' + cm.cmId + '" target="_blank">' + escapeHtml(cm.cmNumber) + '</a></td>';

                // CM Date
                html += '<td data-date="' + (cm.cmDate || '') + '">' + formatDate(cm.cmDate) + '</td>';

                // Customer
                html += '<td>' + escapeHtml(cm.customerName || '-') + '</td>';

                // CM Amount
                html += '<td class="amount">' + formatCurrency(cm.cmAmount) + '</td>';

                // Amount Applied
                html += '<td class="amount">' + formatCurrency(cm.amountApplied) + '</td>';

                // Amount Unapplied
                html += '<td class="amount unapplied">' + formatCurrency(cm.amountUnapplied) + '</td>';

                // Status
                html += '<td>' + escapeHtml(translateCMStatus(cm.cmStatus)) + '</td>';

                // Linked CD # with link
                if (cm.cdId) {
                    html += '<td><a href="/app/accounting/transactions/custdep.nl?id=' + cm.cdId + '" target="_blank">' + escapeHtml(cm.cdNumber) + '</a></td>';
                } else {
                    html += '<td>-</td>';
                }

                // Overpayment Date
                html += '<td data-date="' + (cm.overpaymentDate || '') + '">' + formatDate(cm.overpaymentDate) + '</td>';

                // CD Date
                html += '<td data-date="' + (cm.cdDate || '') + '">' + formatDate(cm.cdDate) + '</td>';

                // Sales Order # with link
                if (cm.soId) {
                    html += '<td><a href="/app/accounting/transactions/salesord.nl?id=' + cm.soId + '" target="_blank">' + escapeHtml(cm.soNumber) + '</a></td>';
                } else {
                    html += '<td>-</td>';
                }

                // SO Date
                html += '<td data-date="' + (cm.soDate || '') + '">' + formatDate(cm.soDate) + '</td>';

                // Department/Selling Location
                html += '<td>' + escapeHtml(cm.soDepartment || '-') + '</td>';

                // Unbilled Orders
                html += '<td class="amount">' + formatCurrency(cm.unbilledOrders) + '</td>';

                // Deposit Balance
                html += '<td class="amount">' + formatCurrency(cm.depositBalance) + '</td>';

                // A/R Balance
                html += '<td class="amount">' + formatCurrencyWithSign(cm.arBalance) + '</td>';

                // Sales Rep
                html += '<td>' + escapeHtml(cm.salesrepName || '-') + '</td>';

                html += '</tr>';
            }

            html += '</tbody>';
            html += '</table>';
            html += '</div>';

            return html;
        }

        /**
         * Searches for unapplied customer deposits linked to Kitchen Works sales orders
         * @param {string} balanceAsOf - Date to filter transactions (YYYY-MM-DD format)
         * @returns {Array} Array of deposit objects
         */
        function searchUnappliedDeposits(balanceAsOf) {
            var deposits = [];
            var result = {
                deposits: [],
                isTruncated: false,
                actualCount: 0,
                actualTotalUnapplied: 0
            };

            try {
                var sql = `
                    SELECT 
                        t.id AS deposit_id,
                        t.tranid AS deposit_number,
                        CASE WHEN t.trandate <= TO_DATE('2024-04-30', 'YYYY-MM-DD') THEN so.trandate ELSE t.trandate END AS deposit_date,
                        t.foreigntotal AS deposit_amount,
                        t.status AS deposit_status,
                        (SELECT COALESCE(SUM(depa2.foreigntotal), 0)
                         FROM previousTransactionLineLink ptll2
                         LEFT JOIN transaction depa2 ON ptll2.nextdoc = depa2.id
                         WHERE ptll2.previousdoc = t.id
                           AND ptll2.linktype = 'DepAppl'
                           AND depa2.trandate <= TO_DATE('` + balanceAsOf + `', 'YYYY-MM-DD')) AS amount_applied,
                        (t.foreigntotal - (SELECT COALESCE(SUM(depa2.foreigntotal), 0)
                                           FROM previousTransactionLineLink ptll2
                                           LEFT JOIN transaction depa2 ON ptll2.nextdoc = depa2.id
                                           WHERE ptll2.previousdoc = t.id
                                             AND ptll2.linktype = 'DepAppl'
                                             AND depa2.trandate <= TO_DATE('` + balanceAsOf + `', 'YYYY-MM-DD'))) AS amount_unapplied,
                        tl_dep.createdfrom AS so_id,
                        so.tranid AS so_number,
                        so.trandate AS so_date,
                        so.status AS so_status,
                        c.altname AS customer_name,
                        c.balanceSearch AS customer_balance,
                        c.depositBalanceSearch AS customer_deposit_balance,
                        c.unbilledOrdersSearch AS customer_unbilled_orders,
                        d.name AS so_department,
                        emp.firstname || ' ' || emp.lastname AS salesrep_name,
                        t.custbody_cd_reconciliation_next_step AS next_step,
                        t.custbody_cd_reconciliation_notes AS reconciliation_notes
                    FROM transaction t
                    INNER JOIN transactionline tl_dep
                            ON t.id = tl_dep.transaction
                           AND tl_dep.mainline = 'T'
                    INNER JOIN transaction so
                            ON tl_dep.createdfrom = so.id
                    INNER JOIN customer c
                            ON so.entity = c.id
                    INNER JOIN transactionline tl_so
                            ON so.id = tl_so.transaction
                           AND tl_so.mainline = 'F'
                    INNER JOIN item i
                            ON tl_so.item = i.id
                    LEFT JOIN department d
                            ON tl_so.department = d.id
                    LEFT JOIN employee emp
                            ON so.employee = emp.id
                    WHERE t.type = 'CustDep'
                      AND i.incomeaccount = 338
                      AND t.trandate <= TO_DATE('` + balanceAsOf + `', 'YYYY-MM-DD')
                    GROUP BY t.id,
                             t.tranid,
                             t.trandate,
                             t.foreigntotal,
                             t.status,
                             tl_dep.createdfrom,
                             so.tranid,
                             so.trandate,
                             so.status,
                             c.altname,
                             c.balanceSearch,
                             c.depositBalanceSearch,
                             c.unbilledOrdersSearch,
                             d.name,
                             emp.firstname,
                             emp.lastname,
                             t.custbody_cd_reconciliation_next_step,
                             t.custbody_cd_reconciliation_notes
                    HAVING (t.foreigntotal - (SELECT COALESCE(SUM(depa2.foreigntotal), 0)
                                              FROM previousTransactionLineLink ptll2
                                              LEFT JOIN transaction depa2 ON ptll2.nextdoc = depa2.id
                                              WHERE ptll2.previousdoc = t.id
                                                AND ptll2.linktype = 'DepAppl'
                                                AND depa2.trandate <= TO_DATE('` + balanceAsOf + `', 'YYYY-MM-DD'))) > 0
                    ORDER BY t.trandate DESC
                `;

                var results = query.runSuiteQL({ query: sql }).asMappedResults();

                for (var i = 0; i < results.length; i++) {
                    var row = results[i];
                    deposits.push({
                        depositId: row.deposit_id,
                        depositNumber: row.deposit_number,
                        depositDate: row.deposit_date,
                        depositAmount: parseFloat(row.deposit_amount) || 0,
                        depositStatus: row.deposit_status,
                        amountApplied: parseFloat(row.amount_applied) || 0,
                        amountUnapplied: parseFloat(row.amount_unapplied) || 0,
                        soId: row.so_id,
                        soNumber: row.so_number,
                        soDate: row.so_date,
                        soStatus: row.so_status,
                        customerName: row.customer_name,
                        customerBalance: parseFloat(row.customer_balance) || 0,
                        customerDepositBalance: parseFloat(row.customer_deposit_balance) || 0,
                        customerUnbilledOrders: parseFloat(row.customer_unbilled_orders) || 0,
                        soDepartment: row.so_department,
                        salesrepName: row.salesrep_name,
                        nextStep: row.next_step || '',
                        reconciliationNotes: row.reconciliation_notes || ''
                    });
                }

                result.deposits = deposits;
                result.actualCount = deposits.length;

                // If we hit exactly 5000, results are likely truncated - run aggregate query for true totals
                if (deposits.length === 5000) {
                    result.isTruncated = true;
                    log.debug('Results Truncated', 'Hit 5000 limit, running aggregate query for accurate totals');

                    var aggregateSql = `
                        SELECT 
                            COUNT(*) AS total_count,
                            SUM(amount_unapplied) AS total_unapplied
                        FROM (
                            SELECT 
                                t.id,
                                (t.foreigntotal - (SELECT COALESCE(SUM(depa2.foreigntotal), 0)
                                                 FROM previousTransactionLineLink ptll2
                                                 LEFT JOIN transaction depa2 ON ptll2.nextdoc = depa2.id
                                                 WHERE ptll2.previousdoc = t.id
                                                   AND ptll2.linktype = 'DepAppl'
                                                   AND depa2.trandate <= TO_DATE('` + balanceAsOf + `', 'YYYY-MM-DD'))) AS amount_unapplied
                            FROM transaction t
                            INNER JOIN transactionline tl_dep
                                    ON t.id = tl_dep.transaction
                                   AND tl_dep.mainline = 'T'
                            INNER JOIN transaction so
                                    ON tl_dep.createdfrom = so.id
                            INNER JOIN transactionline tl_so
                                    ON so.id = tl_so.transaction
                                   AND tl_so.mainline = 'F'
                            INNER JOIN item i
                                    ON tl_so.item = i.id
                            WHERE t.type = 'CustDep'
                              AND i.incomeaccount = 338
                              AND t.trandate <= TO_DATE('` + balanceAsOf + `', 'YYYY-MM-DD')
                            GROUP BY t.id, t.foreigntotal
                            HAVING (t.foreigntotal - (SELECT COALESCE(SUM(depa2.foreigntotal), 0)
                                                      FROM previousTransactionLineLink ptll2
                                                      LEFT JOIN transaction depa2 ON ptll2.nextdoc = depa2.id
                                                      WHERE ptll2.previousdoc = t.id
                                                        AND ptll2.linktype = 'DepAppl'
                                                        AND depa2.trandate <= TO_DATE('` + balanceAsOf + `', 'YYYY-MM-DD'))) > 0
                        )
                    `;

                    var aggregateResults = query.runSuiteQL({ query: aggregateSql }).asMappedResults();
                    if (aggregateResults.length > 0) {
                        result.actualCount = parseInt(aggregateResults[0].total_count) || deposits.length;
                        result.actualTotalUnapplied = parseFloat(aggregateResults[0].total_unapplied) || 0;
                    }

                    log.debug('Aggregate Results', 'Actual count: ' + result.actualCount + ', Actual total unapplied: ' + result.actualTotalUnapplied);
                }

                log.debug('Search Results', 'Found ' + deposits.length + ' unapplied deposits' + (result.isTruncated ? ' (truncated, actual: ' + result.actualCount + ')' : ''));

            } catch (e) {
                log.error('Error Searching Deposits', {
                    error: e.message,
                    stack: e.stack
                });
            }

            return result;
        }

        /**
         * Searches for unapplied credit memos from overpayment deposits linked to Kitchen Works sales orders
         * @param {string} balanceAsOf - Date to filter transactions (YYYY-MM-DD format)
         * @returns {Array} Array of credit memo objects
         */
        function searchUnappliedCreditMemos(balanceAsOf) {
            var creditMemos = [];

            try {
                var sql = `
                    SELECT 
                        cm.id AS cm_id,
                        cm.tranid AS cm_number,
                        cm.trandate AS cm_date,
                        ABS(cm.foreigntotal) AS cm_amount,
                        cm.status AS cm_status,
                        cm.custbody_overpayment_tran AS cd_id,
                        cd.tranid AS cd_number,
                        cm.custbody_overpayment_date AS overpayment_date,
                        cm.custbody_overpayment_cd_date AS cd_date,
                        tl_cd.createdfrom AS so_id,
                        so.tranid AS so_number,
                        so.trandate AS so_date,
                        c.altname AS customer_name,
                        d.name AS so_department,
                        COALESCE(tal.amountlinked, 0) AS amount_linked,
                        (ABS(tal.amount) - COALESCE(tal.amountlinked, 0)) AS amount_unapplied,
                        c.unbilledOrdersSearch AS unbilled_orders,
                        c.depositBalanceSearch AS deposit_balance,
                        c.balanceSearch AS ar_balance,
                        emp.firstname || ' ' || emp.lastname AS salesrep_name
                    FROM transaction cm
                    INNER JOIN transactionaccountingline tal ON cm.id = tal.transaction AND tal.credit IS NOT NULL
                    INNER JOIN transaction cd ON cm.custbody_overpayment_tran = cd.id
                    INNER JOIN transactionline tl_cd ON cd.id = tl_cd.transaction AND tl_cd.mainline = 'T'
                    INNER JOIN transaction so ON tl_cd.createdfrom = so.id
                    INNER JOIN customer c ON cm.entity = c.id
                    INNER JOIN transactionline tl_so ON so.id = tl_so.transaction AND tl_so.mainline = 'F'
                    INNER JOIN item i ON tl_so.item = i.id
                    LEFT JOIN department d ON tl_so.department = d.id
                    LEFT JOIN employee emp ON so.employee = emp.id
                    WHERE cm.type = 'CustCred'
                      AND cm.status = 'A'
                      AND cm.custbody_overpayment_tran IS NOT NULL
                      AND i.incomeaccount = 338
                      AND cm.trandate <= TO_DATE('` + balanceAsOf + `', 'YYYY-MM-DD')
                      AND (ABS(tal.amount) - COALESCE(tal.amountlinked, 0)) > 0
                    GROUP BY cm.id, cm.tranid, cm.trandate, cm.foreigntotal, cm.status,
                             cm.custbody_overpayment_tran, cd.tranid, tl_cd.createdfrom,
                             cm.custbody_overpayment_date, cm.custbody_overpayment_cd_date,
                             so.tranid, so.trandate, c.altname, d.name, tal.amount, tal.amountlinked,
                             c.unbilledOrdersSearch, c.depositBalanceSearch, c.balanceSearch,
                             emp.firstname, emp.lastname
                    ORDER BY cm.trandate DESC
                `;

                var results = query.runSuiteQL({ query: sql }).asMappedResults();

                for (var i = 0; i < results.length; i++) {
                    var row = results[i];
                    var cmAmount = parseFloat(row.cm_amount) || 0;
                    var amountUnapplied = parseFloat(row.amount_unapplied) || 0;
                    var amountApplied = cmAmount - amountUnapplied;

                    creditMemos.push({
                        cmId: row.cm_id,
                        cmNumber: row.cm_number,
                        cmDate: row.cm_date,
                        cmAmount: cmAmount,
                        cmStatus: row.cm_status,
                        amountApplied: amountApplied,
                        amountUnapplied: amountUnapplied,
                        cdId: row.cd_id,
                        cdNumber: row.cd_number,
                        overpaymentDate: row.overpayment_date,
                        cdDate: row.cd_date,
                        soId: row.so_id,
                        soNumber: row.so_number,
                        soDate: row.so_date,
                        customerName: row.customer_name,
                        soDepartment: row.so_department,
                        unbilledOrders: parseFloat(row.unbilled_orders) || 0,
                        depositBalance: parseFloat(row.deposit_balance) || 0,
                        arBalance: parseFloat(row.ar_balance) || 0,
                        salesrepName: row.salesrep_name
                    });
                }

                log.debug('CM Search Results', 'Found ' + creditMemos.length + ' unapplied credit memos');

            } catch (e) {
                log.error('Error Searching Credit Memos', {
                    error: e.message,
                    stack: e.stack
                });
            }

            return creditMemos;
        }

        /**
         * Translates deposit status code to display text
         * @param {string} status - Status code
         * @returns {string} Display text
         */
        function translateStatus(status) {
            var statusMap = {
                'A': 'Not Deposited',
                'B': 'Deposited',
                'C': 'Fully Applied'
            };
            return statusMap[status] || status || '-';
        }

        /**
         * Translates credit memo status code to display text
         * @param {string} status - Status code
         * @returns {string} Display text
         */
        function translateCMStatus(status) {
            var statusMap = {
                'A': 'Open',
                'B': 'Fully Applied'
            };
            return statusMap[status] || status || '-';
        }

        /**
         * Translates sales order status code to display text
         * @param {string} status - Status code
         * @returns {string} Display text
         */
        function translateSOStatus(status) {
            var statusMap = {
                'A': 'Pending Approval',
                'B': 'Pending Fulfillment',
                'C': 'Cancelled',
                'D': 'Partially Fulfilled',
                'E': 'Pending Billing/Partially Fulfilled',
                'F': 'Pending Billing',
                'G': 'Billed',
                'H': 'Closed'
            };
            return statusMap[status] || status || '-';
        }

        /**
         * Gets the icon for a Next Step value
         * @param {string} value - Next Step value (1-4)
         * @returns {string} Icon character
         */
        function getNextStepIcon(value) {
            var iconMap = {
                '1': '↗️',
                '2': '↔️',
                '3': '↩️',
                '4': '🔄'
            };
            return iconMap[value] || '';
        }

        /**
         * Gets the label for a Next Step value
         * @param {string} value - Next Step value (1-4)
         * @returns {string} Label text
         */
        function getNextStepLabel(value) {
            var labelMap = {
                '1': 'Apply to Invoice',
                '2': 'Transfer to Another Customer',
                '3': 'Refund to Customer',
                '4': 'Pending Review'
            };
            return labelMap[value] || '';
        }

        /**
         * Formats a currency value
         * @param {number} value - Currency value
         * @returns {string} Formatted currency
         */
        function formatCurrency(value) {
            if (!value && value !== 0) return '-';
            return '$' + Math.abs(value).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
        }

        /**
         * Formats a currency value preserving the sign (for A/R balance)
         * @param {number} value - Currency value
         * @returns {string} Formatted currency with sign
         */
        function formatCurrencyWithSign(value) {
            if (!value && value !== 0) return '-';
            var prefix = value < 0 ? '-$' : '$';
            return prefix + Math.abs(value).toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
        }

        /**
         * Formats a date value
         * @param {string} dateValue - Date string
         * @returns {string} Formatted date
         */
        function formatDate(dateValue) {
            if (!dateValue) return '-';

            try {
                var date = new Date(dateValue);
                var month = (date.getMonth() + 1).toString().padStart(2, '0');
                var day = date.getDate().toString().padStart(2, '0');
                var year = date.getFullYear();
                return month + '/' + day + '/' + year;
            } catch (e) {
                return dateValue;
            }
        }

        /**
         * Escapes HTML special characters
         * @param {string} text - Text to escape
         * @returns {string} Escaped text
         */
        function escapeHtml(text) {
            if (!text) return '';
            var map = {
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#039;'
            };
            return String(text).replace(/[&<>"']/g, function (m) { return map[m]; });
        }

        /**
         * Returns CSS styles for the page
         * @returns {string} CSS content
         */
        function getStyles() {
            return '' +
                /* Remove NetSuite default borders */
                '.uir-page-title-secondline { border: none !important; margin: 0 !important; padding: 0 !important; }' +
                '.uir-record-type { border: none !important; }' +
                '.bglt { border: none !important; }' +
                '.smalltextnolink { border: none !important; }' +

                /* Loading Spinner Overlay */
                '.loading-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(255, 255, 255, 0.9); display: flex; flex-direction: column; justify-content: center; align-items: center; z-index: 9999; }' +
                '.loading-overlay.hidden { display: none; }' +
                '.loading-spinner { width: 50px; height: 50px; border: 4px solid #e0e0e0; border-top: 4px solid #4CAF50; border-radius: 50%; animation: spin 1s linear infinite; }' +
                '@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }' +
                '.loading-text { margin-top: 15px; font-size: 16px; color: #333; font-weight: 500; }' +

                /* Main container - avoid targeting global td */
                '.portal-container { margin: 0; padding: 20px; border: none; background: transparent; position: relative; }' +

                /* Balance As Of Section */
                '.balance-as-of-section { background: linear-gradient(135deg, #1a237e 0%, #3949ab 100%); border-radius: 8px; padding: 12px 20px; margin-bottom: 20px; display: flex; align-items: center; justify-content: center; gap: 12px; box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15); }' +
                '.balance-as-of-label { color: white; font-size: 15px; font-weight: bold; margin: 0; }' +
                '.balance-as-of-input { padding: 6px 10px; border: 2px solid #fff; border-radius: 4px; font-size: 14px; font-weight: 600; color: #1a237e; background: #fff; cursor: pointer; }' +
                '.balance-as-of-input:focus { outline: none; box-shadow: 0 0 0 3px rgba(255, 255, 255, 0.5); }' +
                '.load-results-btn { padding: 6px 16px; border: 2px solid #fff; border-radius: 4px; font-size: 14px; font-weight: 600; color: #1a237e; background: #fff; cursor: pointer; transition: background 0.2s, color 0.2s; }' +
                '.load-results-btn:hover { background: #c5cae9; }' +
                '.load-results-btn:active { background: #9fa8da; }' +

                /* Summary Row - side by side sections */
                '.summary-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0 30px 0; }' +

                /* Summary Section */
                '.summary-section { background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%); border: 2px solid #dee2e6; border-radius: 8px; padding: 20px; margin: 0; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); }' +
                '.summary-title { margin: 0 0 15px 0; font-size: 24px; font-weight: bold; color: #333; text-align: center; }' +
                '.summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }' +
                '.summary-card { background: white; border: 1px solid #dee2e6; border-radius: 6px; padding: 15px; text-align: center; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08); transition: transform 0.2s, box-shadow 0.2s; }' +
                '.summary-card:hover { transform: translateY(-2px); box-shadow: 0 4px 8px rgba(0, 0, 0, 0.12); }' +
                '.summary-card-title { font-size: 12px; color: #666; margin-bottom: 8px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }' +
                '.summary-card-count { font-size: 14px; color: #333; margin-bottom: 8px; }' +
                '.summary-card-amount { font-size: 18px; font-weight: bold; color: #4CAF50; }' +
                '.summary-totals-row { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }' +
                '.summary-total { background: #fff; border: 2px solid #4CAF50; border-radius: 6px; padding: 12px 10px; text-align: center; font-size: 18px; font-weight: bold; }' +
                '.summary-total-label { color: #333; font-size: 12px; font-weight: 600; }' +
                '.summary-total-amount { color: #4CAF50; font-size: 24px; display: block; margin-top: 8px; }' +
                '.prior-period-header { display: flex; align-items: center; justify-content: center; flex-wrap: nowrap; gap: 6px; }' +
                '.prior-period-date-input { padding: 4px 6px; border: 1px solid #4CAF50; border-radius: 4px; font-size: 12px; color: #333; background: #fff; cursor: pointer; }' +
                '.prior-period-date-input:focus { outline: none; box-shadow: 0 0 0 2px rgba(76, 175, 80, 0.3); }' +
                '.prior-period-helper { display: block; font-size: 12px; font-weight: normal; color: #666; font-style: italic; margin-top: 6px; }' +
                '.aged-header { font-size: 14px; cursor: pointer; width: 30px; min-width: 30px; text-align: center; }' +
                '.aged-icon-cell { text-align: center; font-size: 12px; width: 30px; min-width: 30px; }' +
                '.aged-icon { color: #F57C00; opacity: 0.7; }' +
                'table.data-table td.next-step-cell { width: 40px; min-width: 40px; max-width: 40px; padding: 0 !important; vertical-align: middle; }' +
                '.next-step-trigger, .notes-trigger { display: block; font-size: 18px; cursor: pointer; color: #666; text-align: center; padding: 8px 0; margin: 0; width: 100%; }' +
                '.next-step-trigger:hover, .notes-trigger:hover { color: #4CAF50; }' +

                /* Action Popup Styles */
                '.action-popup-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.3); z-index: 99998; }' +
                '.action-popup-overlay.visible { display: block; }' +
                '.action-popup { display: none; position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.25); z-index: 99999; min-width: 280px; max-width: 400px; }' +
                '.action-popup.visible { display: block; }' +
                '.action-popup-header { display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; background: #4CAF50; color: white; border-radius: 8px 8px 0 0; font-weight: 600; font-size: 14px; }' +
                '.action-popup-close { cursor: pointer; font-size: 20px; line-height: 1; opacity: 0.8; }' +
                '.action-popup-close:hover { opacity: 1; }' +
                '.action-popup-content { padding: 8px 0; }' +
                '.action-popup-option { padding: 12px 16px; cursor: pointer; font-size: 14px; border-bottom: 1px solid #eee; transition: background 0.15s; }' +
                '.action-popup-option:last-child { border-bottom: none; }' +
                '.action-popup-option:hover { background: #e8f5e9; }' +
                '.action-popup-option.selected { background: #c8e6c9; font-weight: 600; }' +
                '.action-popup-textarea { width: 100%; min-height: 100px; padding: 10px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; font-family: inherit; resize: vertical; box-sizing: border-box; margin-bottom: 12px; }' +
                '.action-popup-textarea:focus { outline: none; border-color: #4CAF50; box-shadow: 0 0 0 2px rgba(76,175,80,0.2); }' +
                '.action-popup-buttons { display: flex; gap: 10px; justify-content: flex-end; }' +
                '.action-popup-btn { padding: 8px 16px; border-radius: 4px; font-size: 14px; font-weight: 600; cursor: pointer; transition: background 0.15s; }' +
                '.action-popup-btn-cancel { background: #f5f5f5; border: 1px solid #ddd; color: #666; }' +
                '.action-popup-btn-cancel:hover { background: #eee; }' +
                '.action-popup-btn-save { background: #4CAF50; border: none; color: white; }' +
                '.action-popup-btn-save:hover { background: #45a049; }' +
                '.action-popup-form { padding: 16px; }' +

                /* Search/Data Sections */
                '.search-section { margin-bottom: 30px; }' +
                '.search-title { font-size: 16px; font-weight: bold; margin: 25px 0 0 0; color: #333; padding: 15px 10px 15px 10px; border-bottom: 2px solid #4CAF50; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center; position: -webkit-sticky; position: sticky; top: 0; background: white; z-index: 103; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); }' +
                '.search-title:hover { background-color: #f8f9fa; }' +
                '.search-title.collapsible { padding-left: 10px; padding-right: 10px; }' +
                '.toggle-icon { font-size: 20px; font-weight: bold; color: #4CAF50; transition: transform 0.3s ease; }' +
                '.search-content { transition: max-height 0.3s ease; }' +
                '.search-content.collapsed { display: none; }' +
                '.search-count { font-style: italic; color: #666; margin: 0; font-size: 12px; padding: 10px 10px; background: white; position: -webkit-sticky; position: sticky; top: 51px; z-index: 102; border-bottom: 1px solid #e9ecef; }' +

                /* No results message */
                '.no-results { text-align: center; color: #999; padding: 40px 20px; font-style: italic; }' +

                /* Search Box */
                '.search-box-container { margin: 0; padding: 12px 10px 15px 10px; background: white; position: -webkit-sticky; position: sticky; top: 86px; z-index: 102; border-bottom: 5px solid #4CAF50; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); }' +
                '.search-row { display: flex; gap: 10px; align-items: center; }' +
                '.search-box { flex: 1; padding: 10px 12px; border: 1px solid #cbd5e1; border-radius: 4px; font-size: 14px; box-sizing: border-box; }' +
                '.search-box:focus { outline: none; border-color: #4CAF50; box-shadow: 0 0 0 2px rgba(76, 175, 80, 0.15); }' +
                '.search-results-count { display: none; margin-left: 10px; color: #6c757d; font-size: 13px; font-style: italic; }' +
                '.export-btn { padding: 10px 16px; background: #4CAF50; color: white; border: none; border-radius: 4px; font-size: 14px; font-weight: 600; cursor: pointer; white-space: nowrap; transition: background 0.2s; }' +
                '.export-btn:hover { background: #45a049; }' +
                '.export-btn:active { background: #3d8b40; }' +

                /* Table Container */
                '.table-container { overflow: visible; }' +

                /* Data Table - scoped to .data-table to avoid global td targeting */
                'table.data-table { border-collapse: separate; border-spacing: 0; width: 100%; margin: 0; margin-top: 0 !important; border-left: 1px solid #ddd; border-right: 1px solid #ddd; border-bottom: 1px solid #ddd; background: white; }' +
                'table.data-table thead th { position: -webkit-sticky; position: sticky; top: 157px; z-index: 101; background-color: #f8f9fa; border: 1px solid #ddd; border-top: none; padding: 10px 8px; text-align: left; vertical-align: top; font-weight: bold; color: #333; font-size: 12px; cursor: pointer; user-select: none; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); margin-top: 0; }' +
                'table.data-table thead th:hover { background-color: #e9ecef; }' +
                'table.data-table th, table.data-table td { border: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; color: #000; }' +
                'table.data-table tbody tr:nth-child(even) td { background-color: #f9f9f9; }' +
                'table.data-table tbody tr:hover td { background-color: #e8f4f8; }' +
                'table.data-table a { color: #0c5460; text-decoration: none; }' +
                'table.data-table a:hover { text-decoration: underline; }' +
                'table.data-table td.amount { text-align: right !important; white-space: nowrap; }' +
                'table.data-table td.unapplied { color: #d9534f; font-weight: bold; }' +

                /* Customer Balance Tooltip */
                '.customer-balance-row { cursor: help; }' +
                '.customer-balance-tooltip { display: none; position: fixed; bottom: 20px; right: 20px; background: white; border: 2px solid #4CAF50; border-radius: 6px; padding: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.25); z-index: 100000; min-width: 300px; max-width: 400px; }' +
                '.customer-balance-tooltip.visible { display: block; }' +
                '.customer-balance-tooltip-header { font-weight: 700; color: #1a237e; margin-bottom: 12px; font-size: 16px; border-bottom: 2px solid #4CAF50; padding-bottom: 8px; }' +
                '.tooltip-header { font-weight: 600; color: #333; margin-bottom: 8px; font-size: 14px; border-bottom: 1px solid #cbd5e1; padding-bottom: 4px; }' +
                '.tooltip-detail { padding: 6px 8px; border-bottom: 1px solid #e2e8f0; color: #1a2e1f; display: flex; justify-content: space-between; font-size: 13px; }' +
                '.tooltip-detail:last-child { border-bottom: none; }' +
                '.tooltip-detail:hover { background: #f8faf9; }' +
                '.tooltip-detail-label { font-weight: 500; color: #666; font-size: 13px; }' +
                '.tooltip-detail-value { font-weight: 600; color: #333; text-align: right; font-size: 13px; }' +
                '.tooltip-detail-value.positive { color: #4CAF50; }' +
                '.tooltip-detail-value.negative { color: #d9534f; }' +

                /* Success/Error Messages */
                '.success-msg { background-color: #d4edda; color: #155724; padding: 12px; border: 1px solid #c3e6cb; border-radius: 6px; margin: 15px 0; font-size: 13px; }' +
                '.error-msg { background-color: #f8d7da; color: #721c24; padding: 12px; border: 1px solid #f5c6cb; border-radius: 6px; margin: 15px 0; font-size: 13px; }' +

                /* EXPLAIN Button */
                '.explain-btn { padding: 8px 16px; font-size: 12px; font-weight: 700; color: #fff; background: linear-gradient(135deg, #6366f1, #4f46e5); border: none; border-radius: 6px; cursor: pointer; white-space: nowrap; transition: all 0.2s; box-shadow: 0 2px 4px rgba(99,102,241,0.3); display: block; margin: 0 auto; text-transform: uppercase; letter-spacing: 0.5px; }' +
                '.explain-btn:hover { background: linear-gradient(135deg, #4f46e5, #4338ca); transform: translateY(-1px); box-shadow: 0 3px 8px rgba(99,102,241,0.4); }' +
                
                '.action-btn-cell { text-align: center !important; padding: 4px !important; }' +
                '.ai-badge { display: block; margin: 6px auto 0; padding: 4px 8px; font-size: 11px; font-weight: 600; color: #fff; background: linear-gradient(135deg, #7c4dff, #651fff); border-radius: 4px; text-transform: uppercase; letter-spacing: 0.5px; box-shadow: 0 1px 3px rgba(124,77,255,0.3); cursor: help; transition: all 0.2s; }' +
                '.ai-badge:hover { background: linear-gradient(135deg, #651fff, #6200ea); box-shadow: 0 2px 4px rgba(124,77,255,0.4); transform: translateY(-1px); }' +

                /* Unified Explain Modal */
                '.comparison-modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 99998; }' +
                '.comparison-modal-overlay.visible { display: block; }' +
                '.comparison-modal { display: none; position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; border-radius: 10px; box-shadow: 0 8px 32px rgba(0,0,0,0.3); z-index: 99999; width: 95%; max-width: 1400px; max-height: 90vh; overflow: hidden; }' +
                '.comparison-modal.visible { display: block; }' +
                '.comparison-modal-content { display: flex; flex-direction: column; height: 100%; max-height: 90vh; }' +
                '.explain-modal-content { max-width: 1400px; }' +
                '.comparison-modal-header { display: flex; justify-content: space-between; align-items: center; padding: 16px 20px; background: linear-gradient(135deg, #6366f1, #4f46e5); color: white; border-radius: 10px 10px 0 0; }' +
                '.comparison-modal-title { font-size: 18px; font-weight: 700; }' +
                '.comparison-modal-close { cursor: pointer; font-size: 28px; line-height: 1; opacity: 0.8; padding: 0 8px; }' +
                '.comparison-modal-close:hover { opacity: 1; }' +
                
                /* Tab Navigation */
                '.tab-navigation { display: flex; background: #f8f9fa; border-bottom: 2px solid #dee2e6; padding: 0 20px; }' +
                '.tab-button { flex: 1; padding: 12px 16px; font-size: 14px; font-weight: 600; color: #666; background: transparent; border: none; border-bottom: 3px solid transparent; cursor: pointer; transition: all 0.2s; }' +
                '.tab-button:hover { color: #333; background: #e9ecef; }' +
                '.tab-button.active { color: #6366f1; border-bottom-color: #6366f1; background: white; }' +
                
                /* Tab Content */
                '.comparison-modal-body { padding: 20px; overflow-y: auto; flex: 1; }' +
                '.tab-content { display: none; }' +
                '.tab-content.active { display: block; }' +
                '.comparison-loading { text-align: center; padding: 60px 20px; color: #666; font-size: 16px; }' +

                /* Comparison Summary Cards */
                '.comparison-summary { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 25px; }' +
                '.comparison-card { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; padding: 12px 10px; text-align: center; }' +
                '.comparison-card.highlight { background: linear-gradient(135deg, #e3f2fd, #bbdefb); border-color: #1976d2; }' +
                '.comparison-card.warning { background: linear-gradient(135deg, #fff3e0, #ffe0b2); border-color: #f57c00; }' +
                '.comparison-card.info { background: linear-gradient(135deg, #e3f2fd, #bbdefb); border-color: #2196f3; }' +
                '.comparison-card.error { background: linear-gradient(135deg, #ffebee, #ffcdd2); border-color: #d32f2f; }' +
                '.comparison-card.success { background: linear-gradient(135deg, #e8f5e9, #c8e6c9); border-color: #4caf50; }' +
                '.comparison-card-label { font-size: 10px; color: #666; text-transform: uppercase; font-weight: 600; letter-spacing: 0.3px; margin-bottom: 5px; }' +
                '.comparison-card-value { font-size: 17px; font-weight: 700; color: #333; }' +
                '.comparison-card-sub { font-size: 11px; color: #666; margin-top: 3px; }' +

                /* Comparison Transaction Links */
                '.comparison-transactions { display: flex; gap: 15px; margin-bottom: 25px; flex-wrap: wrap; }' +
                '.comparison-tran-grid { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 20px; }' +
                '.comparison-tran-link { display: flex; align-items: center; gap: 8px; padding: 10px 16px; background: #fff; border: 1px solid #dee2e6; border-radius: 6px; text-decoration: none; color: #333; transition: all 0.2s; min-width: 140px; flex: 0 0 auto; }' +
                '.comparison-tran-link:hover { border-color: #1976d2; background: #e3f2fd; }' +
                '.comparison-tran-label { font-size: 10px; color: #666; text-transform: uppercase; letter-spacing: 0.3px; }' +
                '.comparison-tran-value { font-size: 14px; font-weight: 600; color: #1976d2; }' +
                '.comparison-tran-amount { font-size: 12px; color: #333; font-weight: 500; margin-top: 2px; }' +
                '.cm-aggregate { background: linear-gradient(135deg, #fff3e0, #ffe0b2); border-color: #f57c00; border-width: 2px; }' +
                '.cm-aggregate .comparison-tran-label { color: #e65100; }' +
                '.cm-aggregate .comparison-tran-value { color: #e65100; }' +
                '.cm-aggregate .comparison-tran-amount { color: #bf360c; font-weight: 700; }' +
                '.cm-individual { background: #fff8e1; border-style: dashed; font-size: 90%; }' +
                '.cm-individual .comparison-tran-amount { color: #bf360c; }' +

                /* Comparison Table */
                '.comparison-table-container { overflow-x: auto; }' +
                '.comparison-table { width: 100%; border-collapse: collapse; font-size: 13px; }' +
                '.comparison-table th { background: #f8f9fa; padding: 10px 8px; text-align: left; font-weight: 600; border-bottom: 2px solid #dee2e6; position: sticky; top: 0; }' +
                '.comparison-table td { padding: 8px; border-bottom: 1px solid #eee; }' +
                '.comparison-table tr:hover td { background: #f5f5f5; }' +
                '.comparison-table .amount { text-align: right; font-family: monospace; }' +

                /* Sales Order Totals Visual Column Grouping */
                '.so-totals-table td.group-total-start, .so-totals-table td.group-total-end { background: rgba(76, 175, 80, 0.1) !important; }' +
                '.so-totals-table td.group-nontax-start, .so-totals-table td.group-nontax-end { background: rgba(33, 150, 243, 0.1) !important; }' +
                '.so-totals-table td.group-tax-start, .so-totals-table td.group-tax-end { background: rgba(255, 235, 59, 0.15) !important; }' +
                '.so-totals-table tr.bill-variance-row td { border-top: 3px solid #9c27b0 !important; border-bottom: 3px solid #9c27b0 !important; }' +
                '.so-totals-table .unbilled-amount { color: #9c27b0; font-weight: 600; }' +
                '.comparison-table .match { color: #4caf50; font-weight: 600; }' +
                '.comparison-table .mismatch { color: #d32f2f; font-weight: 600; }' +
                '.comparison-table .not-invoiced { color: #daa520; font-weight: 600; }' +
                '.comparison-table .not-on-so { color: #d32f2f; font-weight: 600; }' +
                '.comparison-table .discount { color: #9c27b0; font-weight: 600; }' +
                '.comparison-table .variance-positive { color: #4caf50; }' +
                '.comparison-table .variance-negative { color: #d32f2f; }' +
                '.match-row { background: #f1f8f4; }' +
                '.no-invoice-row { background: #fff8e1; }' +
                '.overpayment-row { background: #fff3cd; border: 2px solid #ffc107 !important; }' +
                '.status-badge { display: inline-block; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }' +
                '.status-success { background: #4caf50; color: white; }' +
                '.status-error { background: #dc3545; color: white; }' +
                '.status-warning { background: #ff9800; color: white; }' +
                '.status-overpayment { background: #ffc107; color: #000; }' +
                
                /* Multi-SO Comparison Styling */
                '.customer-analysis-header { background: #fff; border: 2px solid #1976d2; border-radius: 8px; padding: 20px; margin-bottom: 20px; }' +
                '.customer-analysis-header h2 { margin: 0 0 10px 0; font-size: 20px; color: #1976d2; font-weight: 600; }' +
                '.customer-analysis-header > div { color: #495057; font-size: 14px; margin-bottom: 15px; }' +
                '.customer-overview-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-top: 15px; }' +
                '.customer-overview-card { background: rgba(255, 255, 255, 0.15); padding: 12px; border-radius: 6px; backdrop-filter: blur(10px); }' +
                '.customer-overview-label { font-size: 11px; opacity: 0.9; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 5px; }' +
                '.customer-overview-value { font-size: 20px; font-weight: 700; }' +
                '.customer-overview-sub { font-size: 12px; opacity: 0.85; margin-top: 4px; }' +
                
                '.summary-table-wrapper { background: white; border: 2px solid #e0e0e0; border-radius: 8px; padding: 15px; margin-bottom: 30px; }' +
                '.summary-table-wrapper h3 { margin: 0 0 15px 0; font-size: 16px; color: #333; }' +
                '.summary-table { width: 100%; border-collapse: collapse; font-size: 13px; }' +
                '.summary-table thead th { background: #f8f9fa; padding: 10px; text-align: left; font-weight: 600; border-bottom: 2px solid #dee2e6; }' +
                '.summary-table tbody td { padding: 10px; border-bottom: 1px solid #eee; }' +
                '.summary-table tbody tr:hover { background: #f5f5f5; }' +
                '.summary-table .amount { text-align: right; font-family: monospace; }' +
                '.summary-table .status-icon { font-size: 16px; }' +
                
                '.so-section { margin-bottom: 40px; border: 2px solid #dee2e6; border-radius: 8px; padding: 20px; background: white; }' +
                '.so-section.source-so { border-color: #9c27b0; border-width: 3px; background: rgba(156, 39, 176, 0.02); }' +
                '.so-section.has-mismatch { border-color: #ff9800; border-width: 2px; background: rgba(255, 152, 0, 0.02); }' +
                '.so-section.has-unbilled { border-color: #ffc107; border-width: 2px; background: rgba(255, 193, 7, 0.02); }' +
                '.so-section.no-issues { border-color: #4caf50; background: rgba(76, 175, 80, 0.02); }' +
                
                '.so-section-header { display: flex; align-items: center; gap: 12px; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 2px solid #e0e0e0; }' +
                '.so-section-icon { font-size: 24px; }' +
                '.so-section-title { flex: 1; }' +
                '.so-section-title h3 { margin: 0; font-size: 18px; color: #333; }' +
                '.so-section-title .so-status { font-size: 12px; color: #666; margin-top: 4px; }' +
                '.so-section-badges { display: flex; gap: 8px; }' +
                '.so-badge { padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: 600; text-transform: uppercase; }' +
                '.so-badge.source { background: #9c27b0; color: white; }' +
                '.so-badge.mismatch { background: #ff9800; color: white; }' +
                '.so-badge.unbilled { background: #ffc107; color: #000; }' +
                '.so-badge.clean { background: #4caf50; color: white; }' +
                
                '.detailed-comparison-divider { margin: 40px 0; border-top: 3px dashed #dee2e6; position: relative; }' +
                '.detailed-comparison-divider::after { content: "DETAILED COMPARISON RESULTS"; position: absolute; top: -10px; left: 50%; transform: translateX(-50%); background: white; padding: 0 15px; font-size: 11px; font-weight: 700; color: #666; letter-spacing: 1px; }' +

                /* Conclusion Box */
                '.comparison-conclusion { margin-top: 25px; padding: 15px 20px; border-radius: 8px; font-size: 14px; font-weight: 600; }' +
                '.comparison-conclusion.match { background: #e8f5e9; border: 2px solid #4caf50; color: #2e7d32; }' +
                '.comparison-conclusion.mismatch { background: #ffebee; border: 2px solid #d32f2f; color: #c62828; }' +

                /* Totals Row */
                '.comparison-totals { background: #f8f9fa; border-top: 2px solid #dee2e6; }' +
                '.comparison-totals td { font-weight: 700; padding: 12px 8px !important; }';
        }

        /**
         * Returns JavaScript for the page
         * @param {string} scriptUrl - Suitelet URL
         * @returns {string} JavaScript content
         */
        function getJavaScript(scriptUrl) {
            return '' +
                /* Define scriptUrl for fetch calls */
                'var scriptUrl = "' + scriptUrl + '";' +
                '' +
                /* Event delegation for collapsible sections */
                '(function() {' +
                '    document.addEventListener(\'click\', function(e) {' +
                '        var target = e.target.closest(\'.search-title.collapsible\');' +
                '        if (target) {' +
                '            var sectionId = target.getAttribute(\'data-section-id\');' +
                '            if (sectionId) {' +
                '                toggleSection(sectionId);' +
                '            }' +
                '        }' +
                '    });' +
                '})();' +

                /* Toggle section visibility */
                'function toggleSection(sectionId) {' +
                '    var content = document.getElementById(\'content-\' + sectionId);' +
                '    var icon = document.getElementById(\'toggle-\' + sectionId);' +
                '    if (content && icon) {' +
                '        if (content.classList.contains(\'collapsed\')) {' +
                '            content.classList.remove(\'collapsed\');' +
                '            icon.textContent = String.fromCharCode(8722);' +
                '            saveExpandedState(sectionId, true);' +
                '        } else {' +
                '            content.classList.add(\'collapsed\');' +
                '            icon.textContent = \'+\';' +
                '            saveExpandedState(sectionId, false);' +
                '        }' +
                '    }' +
                '}' +

                /* Save expanded state to localStorage */
                'function saveExpandedState(sectionId, isExpanded) {' +
                '    try {' +
                '        var state = JSON.parse(localStorage.getItem(\'ucd_kw_expanded\') || \'{}\');' +
                '        state[sectionId] = isExpanded;' +
                '        localStorage.setItem(\'ucd_kw_expanded\', JSON.stringify(state));' +
                '    } catch (e) {}' +
                '}' +

                /* Customer Balance Tooltip Functions */
                'function showCustomerBalanceTooltip(row) {' +
                '  var tooltip = document.getElementById("customerBalanceTooltip");' +
                '  var tooltipContent = document.getElementById("customerBalanceContent");' +
                '  if (!tooltip || !tooltipContent) return;' +
                '  var customerBalance = parseFloat(row.getAttribute("data-customer-balance") || 0);' +
                '  var depositBalance = parseFloat(row.getAttribute("data-customer-deposit-balance") || 0);' +
                '  var unbilledOrders = parseFloat(row.getAttribute("data-customer-unbilled-orders") || 0);' +
                '  var customerName = row.getAttribute("data-customer-name") || "";' +
                '  var nextStep = row.getAttribute("data-next-step") || "";' +
                '  var reconciliationNotes = row.getAttribute("data-reconciliation-notes") || "";' +
                '  var balanceClass = customerBalance > 0 ? "positive" : customerBalance < 0 ? "negative" : "";' +
                '  var depositClass = depositBalance > 0 ? "positive" : depositBalance < 0 ? "negative" : "";' +
                '  var unbilledClass = unbilledOrders > 0 ? "positive" : unbilledOrders < 0 ? "negative" : "";' +
                '  var nextStepLabels = { "1": "↗️ Fulfill & Bill for CD Application", "2": "↔️ Move to Different Sales Order", "3": "↩️ Refund Customer", "4": "🔄 Update Sales Order from Lead Tracker", "5": "🆗 Old CD Approved to Remain on Account" };' +
                '  var html = "";' +
                '  if (nextStep) {' +
                '    html += "<div class=\\"tooltip-detail\\" style=\\"margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid #ddd;\\"><span class=\\"tooltip-detail-label\\">Next Step:</span><span class=\\"tooltip-detail-value\\">" + (nextStepLabels[nextStep] || nextStep) + "</span></div>";' +
                '  }' +
                '  if (reconciliationNotes) {' +
                '    html += "<div class=\\"tooltip-detail\\" style=\\"margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid #ddd;\\"><span class=\\"tooltip-detail-label\\">Notes:</span><span class=\\"tooltip-detail-value\\" style=\\"font-style:italic;\\">" + reconciliationNotes + "</span></div>";' +
                '  }' +
                '  html += "<div class=\\"tooltip-detail\\"><span class=\\"tooltip-detail-label\\">A/R Balance:</span><span class=\\"tooltip-detail-value " + balanceClass + "\\">$" + customerBalance.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",") + "</span></div>";' +
                '  html += "<div class=\\"tooltip-detail\\"><span class=\\"tooltip-detail-label\\">Deposit Balance:</span><span class=\\"tooltip-detail-value " + depositClass + "\\">$" + depositBalance.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",") + "</span></div>";' +
                '  html += "<div class=\\"tooltip-detail\\"><span class=\\"tooltip-detail-label\\">Unbilled Orders:</span><span class=\\"tooltip-detail-value " + unbilledClass + "\\">$" + unbilledOrders.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",") + "</span></div>";' +
                '  var allRows = document.querySelectorAll(".customer-balance-row");' +
                '  var customerDeposits = [];' +
                '  for (var i = 0; i < allRows.length; i++) {' +
                '    if (allRows[i].getAttribute("data-customer-name") === customerName) {' +
                '      customerDeposits.push({' +
                '        depositId: allRows[i].getAttribute("data-deposit-id"),' +
                '        depositNumber: allRows[i].getAttribute("data-deposit-number"),' +
                '        soNumber: allRows[i].getAttribute("data-so-number"),' +
                '        unappliedAmount: parseFloat(allRows[i].getAttribute("data-unapplied-amount") || 0)' +
                '      });' +
                '    }' +
                '  }' +
                '  if (customerDeposits.length > 1) {' +
                '    html += "<div class=\\"tooltip-detail\\" style=\\"margin-top:10px;padding-top:10px;border-top:1px solid #ddd;\\"><strong>Multiple Deposits ("+customerDeposits.length+"):</strong></div>";' +
                '    var totalUnapplied = 0;' +
                '    for (var i = 0; i < customerDeposits.length; i++) {' +
                '      var cd = customerDeposits[i];' +
                '      totalUnapplied += cd.unappliedAmount;' +
                '      html += "<div class=\\"tooltip-detail\\" style=\\"padding-left:10px;\\"><span>"+cd.depositNumber+" / "+cd.soNumber+":</span><span style=\\"float:right;\\">$"+cd.unappliedAmount.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",")+"</span></div>";' +
                '    }' +
                '    html += "<div class=\\"tooltip-detail\\" style=\\"padding-left:10px;font-weight:bold;margin-top:5px;padding-top:5px;border-top:1px solid #eee;\\"><span>Total Unapplied:</span><span style=\\"float:right;\\">$"+totalUnapplied.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",")+"</span></div>";' +
                '  }' +
                '  tooltipContent.innerHTML = html;' +
                '  tooltip.className = "customer-balance-tooltip visible";' +
                '}' +
                '' +
                'function hideCustomerBalanceTooltip() {' +
                '  var tooltip = document.getElementById("customerBalanceTooltip");' +
                '  if (tooltip) tooltip.className = "customer-balance-tooltip";' +
                '}' +
                '' +
                /* Restore expanded state from localStorage */
                'function restoreExpandedState() {' +
                '    try {' +
                '        var state = JSON.parse(localStorage.getItem(\'ucd_kw_expanded\') || \'{}\');' +
                '        for (var sectionId in state) {' +
                '            if (state.hasOwnProperty(sectionId)) {' +
                '                var content = document.getElementById(\'content-\' + sectionId);' +
                '                var icon = document.getElementById(\'toggle-\' + sectionId);' +
                '                if (content && icon) {' +
                '                    if (state[sectionId]) {' +
                '                        content.classList.remove(\'collapsed\');' +
                '                        icon.textContent = String.fromCharCode(8722);' +
                '                    } else {' +
                '                        content.classList.add(\'collapsed\');' +
                '                        icon.textContent = \'+\';' +
                '                    }' +
                '                }' +
                '            }' +
                '        }' +
                '    } catch (e) {}' +
                '}' +

                /* Restore state on page load */
                '/* Show loading spinner */' +
                'function showLoading(message) {' +
                '    var overlay = document.getElementById(\'loadingOverlay\');' +
                '    if (overlay) {' +
                '        var textEl = overlay.querySelector(\'div:last-child\');' +
                '        if (textEl && message) textEl.textContent = message;' +
                '        overlay.style.display = \'flex\';' +
                '    }' +
                '}' +
                '' +
                '/* Hide loading spinner */' +
                'function hideLoading() {' +
                '    var overlay = document.getElementById(\'loadingOverlay\');' +
                '    if (overlay) overlay.style.display = \'none\';' +
                '}' +
                '' +
                '/* Action Popup - shared state */' +
                'var currentPopupDepositId = null;' +
                'var currentPopupType = null;' +
                '' +
                '/* Show Next Step popup */' +
                'function showNextStepPopup(trigger) {' +
                '    var depositId = trigger.getAttribute("data-deposit-id");' +
                '    var currentValue = trigger.getAttribute("data-current-value") || "";' +
                '    currentPopupDepositId = depositId;' +
                '    currentPopupType = "nextstep";' +
                '    ' +
                '    var options = [' +
                '        { value: "1", icon: "↗️", label: "Fulfill & Bill for CD Application" },' +
                '        { value: "2", icon: "↔️", label: "Move to Different Sales Order" },' +
                '        { value: "3", icon: "↩️", label: "Refund Customer" },' +
                '        { value: "4", icon: "🔄", label: "Update Sales Order from Lead Tracker" },' +
                '        { value: "5", icon: "🆗", label: "Old CD Approved to Remain on Account" }' +
                '    ];' +
                '    ' +
                '    var content = "";' +
                '    for (var i = 0; i < options.length; i++) {' +
                '        var opt = options[i];' +
                '        var selectedClass = (opt.value === currentValue) ? " selected" : "";' +
                '        content += "<div class=\\"action-popup-option" + selectedClass + "\\" onclick=\\"selectNextStep(\'" + opt.value + "\', \'" + opt.icon + "\')\\">";' +
                '        content += opt.icon + " " + opt.label;' +
                '        content += "</div>";' +
                '    }' +
                '    ' +
                '    document.getElementById("actionPopupTitle").textContent = "Select Next Step";' +
                '    document.getElementById("actionPopupContent").innerHTML = content;' +
                '    document.getElementById("actionPopup").classList.add("visible");' +
                '    document.getElementById("actionPopupOverlay").classList.add("visible");' +
                '}' +
                '' +
                '/* Hide action popup */' +
                'function hideActionPopup() {' +
                '    document.getElementById("actionPopup").classList.remove("visible");' +
                '    document.getElementById("actionPopupOverlay").classList.remove("visible");' +
                '    currentPopupDepositId = null;' +
                '    currentPopupType = null;' +
                '}' +
                '' +
                '/* Select Next Step option */' +
                'function selectNextStep(value, icon) {' +
                '    var depositId = currentPopupDepositId;' +
                '    hideActionPopup();' +
                '    ' +
                '    var trigger = document.getElementById("nextstep-" + depositId);' +
                '    var row = document.getElementById("dep-row-" + depositId);' +
                '    var originalIcon = trigger.textContent;' +
                '    trigger.textContent = "⏳";' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ depositId: depositId, nextStep: value, notes: "" })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(data) {' +
                '        if (data.success) {' +
                '            trigger.textContent = icon;' +
                '            trigger.setAttribute("data-current-value", value);' +
                '            trigger.title = "Change Next Step";' +
                '            if (row) {' +
                '                row.setAttribute("data-next-step", value);' +
                '            }' +
                '        } else {' +
                '            alert("Error saving: " + data.message);' +
                '            trigger.textContent = originalIcon;' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        alert("Error saving: " + err.message);' +
                '        trigger.textContent = originalIcon;' +
                '    });' +
                '}' +
                '' +
                '/* Show Notes popup */' +
                'function showNotesPopup(trigger) {' +
                '    var depositId = trigger.getAttribute("data-deposit-id");' +
                '    var currentNotes = trigger.getAttribute("data-current-notes") || "";' +
                '    currentPopupDepositId = depositId;' +
                '    currentPopupType = "notes";' +
                '    ' +
                '    var content = "<div class=\\"action-popup-form\\">";' +
                '    content += "<textarea id=\\"notesTextarea\\" class=\\"action-popup-textarea\\" placeholder=\\"Enter notes for this deposit...\\">" + currentNotes + "</textarea>";' +
                '    content += "<div class=\\"action-popup-buttons\\">";' +
                '    content += "<button type=\\"button\\" class=\\"action-popup-btn action-popup-btn-cancel\\" onclick=\\"hideActionPopup()\\">Cancel</button>";' +
                '    content += "<button type=\\"button\\" class=\\"action-popup-btn action-popup-btn-save\\" onclick=\\"saveNotes()\\">Save Notes</button>";' +
                '    content += "</div>";' +
                '    content += "</div>";' +
                '    ' +
                '    document.getElementById("actionPopupTitle").textContent = "Edit Notes";' +
                '    document.getElementById("actionPopupContent").innerHTML = content;' +
                '    document.getElementById("actionPopup").classList.add("visible");' +
                '    document.getElementById("actionPopupOverlay").classList.add("visible");' +
                '    ' +
                '    setTimeout(function() {' +
                '        var textarea = document.getElementById("notesTextarea");' +
                '        if (textarea) { textarea.focus(); textarea.select(); }' +
                '    }, 100);' +
                '}' +
                '' +
                '/* Save Notes */' +
                'function saveNotes() {' +
                '    var depositId = currentPopupDepositId;' +
                '    var textarea = document.getElementById("notesTextarea");' +
                '    var newNotes = textarea ? textarea.value.trim() : "";' +
                '    hideActionPopup();' +
                '    ' +
                '    var trigger = document.getElementById("notes-" + depositId);' +
                '    var row = document.getElementById("dep-row-" + depositId);' +
                '    var originalIcon = trigger.textContent;' +
                '    trigger.textContent = "⏳";' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ depositId: depositId, notes: newNotes, updateNotesOnly: true })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(data) {' +
                '        if (data.success) {' +
                '            trigger.textContent = newNotes ? "ℹ️" : "+";' +
                '            trigger.setAttribute("data-current-notes", newNotes);' +
                '            trigger.title = newNotes ? "View/Edit Notes" : "Add Notes";' +
                '            if (row) {' +
                '                row.setAttribute("data-reconciliation-notes", newNotes);' +
                '            }' +
                '        } else {' +
                '            alert("Error saving: " + data.message);' +
                '            trigger.textContent = originalIcon;' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        alert("Error saving: " + err.message);' +
                '        trigger.textContent = originalIcon;' +
                '    });' +
                '}' +
                '' +
                'document.addEventListener(\'DOMContentLoaded\', function() {' +
                '    /* Hide loading spinner once page is ready */' +
                '    hideLoading();' +
                '    restoreExpandedState();' +
                '    calculatePriorPeriodAmount();' +
                '    calculateCMPriorPeriodAmount();' +
                '    var dateInput = document.getElementById(\'priorPeriodDate\');' +
                '    if (dateInput) {' +
                '        dateInput.addEventListener(\'change\', calculatePriorPeriodAmount);' +
                '    }' +
                '    var cmDateInput = document.getElementById(\'cmPriorPeriodDate\');' +
                '    if (cmDateInput) {' +
                '        cmDateInput.addEventListener(\'change\', calculateCMPriorPeriodAmount);' +
                '    }' +
                '    /* Balance As Of Load Results button handler */' +
                '    var loadResultsBtn = document.getElementById(\'loadResultsBtn\');' +
                '    if (loadResultsBtn) {' +
                '        loadResultsBtn.addEventListener(\'click\', function() {' +
                '            var balanceAsOfInput = document.getElementById(\'balanceAsOfDate\');' +
                '            var newDate = balanceAsOfInput ? balanceAsOfInput.value : null;' +
                '            if (newDate) {' +
                '                showLoading(\'Loading results for \' + newDate + \'...\');' +
                '                var baseUrl = \'' + scriptUrl + '\';' +
                '                var separator = baseUrl.indexOf(\'?\') > -1 ? \'&\' : \'?\';' +
                '                window.location.href = baseUrl + separator + \'loadData=T&balanceAsOf=\' + newDate;' +
                '            }' +
                '        });' +
                '    }' +
                '});' +
                '' +
                '/* Calculate unapplied amount for deposits prior to selected date and update aged icons */' +
                'function calculatePriorPeriodAmount() {' +
                '    var dateInput = document.getElementById(\'priorPeriodDate\');' +
                '    var amountSpan = document.getElementById(\'priorPeriodAmount\');' +
                '    if (!dateInput || !amountSpan) return;' +
                '    ' +
                '    var cutoffDate = new Date(dateInput.value + \'T23:59:59\');' +
                '    var table = document.getElementById(\'table-deposits\');' +
                '    if (!table) { amountSpan.textContent = \'$0.00\'; return; }' +
                '    ' +
                '    var rows = table.querySelectorAll(\'tbody tr\');' +
                '    var total = 0;' +
                '    ' +
                '    for (var i = 0; i < rows.length; i++) {' +
                '        var agedCell = rows[i].cells[3];' +
                '        if (!agedCell || !agedCell.hasAttribute(\'data-date\')) continue;' +
                '        var dateCell = rows[i].cells[4];' +
                '        var unappliedCell = rows[i].cells[8];' +
                '        var dateStr = agedCell.getAttribute(\'data-date\');' +
                '        ' +
                '        if (dateStr) {' +
                '            var rowDate = new Date(dateStr);' +
                '            if (rowDate <= cutoffDate) {' +
                '                var amountText = unappliedCell.textContent.replace(/[^0-9.-]/g, \'\');' +
                '                total += parseFloat(amountText) || 0;' +
                '                agedCell.innerHTML = \'<span class="aged-icon" title="Received before \' + dateInput.value + \'">\u23f0</span>\';' +
                '            } else {' +
                '                agedCell.innerHTML = \'\';' +
                '            }' +
                '        } else {' +
                '            agedCell.innerHTML = \'\';' +
                '        }' +
                '    }' +
                '    ' +
                '    amountSpan.textContent = \'$\' + total.toFixed(2).replace(/\\d(?=(\\d{3})+\\.)/g, \'$&,\');' +
                '}' +
                '' +
                '/* Calculate unapplied amount for credit memos prior to selected overpayment date */' +
                'function calculateCMPriorPeriodAmount() {' +
                '    var dateInput = document.getElementById(\'cmPriorPeriodDate\');' +
                '    var amountSpan = document.getElementById(\'cmPriorPeriodAmount\');' +
                '    if (!dateInput || !amountSpan) return;' +
                '    ' +
                '    var cutoffDate = new Date(dateInput.value + \'T23:59:59\');' +
                '    var table = document.getElementById(\'table-creditmemos\');' +
                '    if (!table) { amountSpan.textContent = \'$0.00\'; return; }' +
                '    ' +
                '    var rows = table.querySelectorAll(\'tbody tr\');' +
                '    var total = 0;' +
                '    ' +
                '    for (var i = 0; i < rows.length; i++) {' +
                '        var overpaymentDateCell = rows[i].cells[8];' +
                '        var unappliedCell = rows[i].cells[5];' +
                '        var dateStr = overpaymentDateCell.getAttribute(\'data-date\');' +
                '        ' +
                '        if (dateStr) {' +
                '            var rowDate = new Date(dateStr);' +
                '            if (rowDate <= cutoffDate) {' +
                '                var amountText = unappliedCell.textContent.replace(/[^0-9.-]/g, \'\');' +
                '                total += parseFloat(amountText) || 0;' +
                '            }' +
                '        }' +
                '    }' +
                '    ' +
                '    amountSpan.textContent = \'$\' + total.toFixed(2).replace(/\\d(?=(\\d{3})+\\.)/g, \'$&,\');' +
                '}' +

                /* Sort table by column */
                'function sortTable(sectionId, columnIndex) {' +
                '    var table = document.getElementById(\'table-\' + sectionId);' +
                '    var tbody = table.querySelector(\'tbody\');' +
                '    var rows = Array.from(tbody.querySelectorAll(\'tr\'));' +
                '    var currentSort = table.getAttribute(\'data-sort-col\');' +
                '    var currentDir = table.getAttribute(\'data-sort-dir\') || \'asc\';' +
                '    var newDir = (currentSort == columnIndex && currentDir == \'asc\') ? \'desc\' : \'asc\';' +
                '    ' +
                '    var headerCell = table.querySelectorAll(\'th\')[columnIndex];' +
                '    var originalText = headerCell.textContent.replace(/ [▲▼]/g, \'\');' +
                '    headerCell.textContent = \'⏳ Sorting...\';' +
                '    headerCell.style.pointerEvents = \'none\';' +
                '    ' +
                '    setTimeout(function() {' +
                '        rows.sort(function(a, b) {' +
                '            var aCell = a.cells[columnIndex];' +
                '            var bCell = b.cells[columnIndex];' +
                '            var aVal = aCell.getAttribute(\'data-date\') || aCell.textContent.trim();' +
                '            var bVal = bCell.getAttribute(\'data-date\') || bCell.textContent.trim();' +
                '            ' +
                '            if (aCell.classList.contains(\'amount\')) {' +
                '                aVal = parseFloat(aVal.replace(/[^0-9.-]/g, \'\')) || 0;' +
                '                bVal = parseFloat(bVal.replace(/[^0-9.-]/g, \'\')) || 0;' +
                '            } else if (aCell.hasAttribute(\'data-date\')) {' +
                '                var parseDate = function(d) {' +
                '                    if (!d || d === \'-\') return 0;' +
                '                    if (d.indexOf(\'/\') > 0) {' +
                '                        var parts = d.split(\'/\');' +
                '                        return parseInt(parts[2]) * 10000 + parseInt(parts[0]) * 100 + parseInt(parts[1]);' +
                '                    }' +
                '                    return parseInt(d.replace(/-/g, \'\'));' +
                '                };' +
                '                aVal = parseDate(aVal);' +
                '                bVal = parseDate(bVal);' +
                '            } else {' +
                '                aVal = aVal.toLowerCase();' +
                '                bVal = bVal.toLowerCase();' +
                '            }' +
                '            ' +
                '            if (aVal < bVal) return newDir === \'asc\' ? -1 : 1;' +
                '            if (aVal > bVal) return newDir === \'asc\' ? 1 : -1;' +
                '            return 0;' +
                '        });' +
                '        ' +
                '        rows.forEach(function(row) { tbody.appendChild(row); });' +
                '        table.setAttribute(\'data-sort-col\', columnIndex);' +
                '        table.setAttribute(\'data-sort-dir\', newDir);' +
                '        ' +
                '        var allHeaders = table.querySelectorAll(\'th\');' +
                '        for (var i = 0; i < allHeaders.length; i++) {' +
                '            var header = allHeaders[i];' +
                '            if (i == columnIndex) {' +
                '                header.textContent = originalText + (newDir === \'asc\' ? \' ▲\' : \' ▼\');' +
                '            } else {' +
                '                var text = header.textContent.replace(/ [▲▼]/g, \'\').trim();' +
                '                header.textContent = text;' +
                '            }' +
                '        }' +
                '        headerCell.style.pointerEvents = \'\';' +
                '    }, 10);' +
                '}' +

                /* Filter table rows based on search input */
                'function filterTable(sectionId) {' +
                '    var input = document.getElementById(\'searchBox-\' + sectionId);' +
                '    var filter = input.value.toUpperCase();' +
                '    var tbody = document.querySelector(\'#table-\' + sectionId + \' tbody\');' +
                '    var rows = tbody.querySelectorAll(\'tr\');' +
                '    var visibleCount = 0;' +
                '    ' +
                '    for (var i = 0; i < rows.length; i++) {' +
                '        var row = rows[i];' +
                '        var text = row.textContent || row.innerText;' +
                '        if (text.toUpperCase().indexOf(filter) > -1) {' +
                '            row.style.display = \'\';' +
                '            visibleCount++;' +
                '        } else {' +
                '            row.style.display = \'none\';' +
                '        }' +
                '    }' +
                '    ' +
                '    var countSpan = document.getElementById(\'searchCount-\' + sectionId);' +
                '    if (filter) {' +
                '        countSpan.textContent = \'Showing \' + visibleCount + \' of \' + rows.length + \' results\';' +
                '        countSpan.style.display = \'inline\';' +
                '    } else {' +
                '        countSpan.style.display = \'none\';' +
                '    }' +
                '}' +
                '' +
                '/* Export table to Excel using SheetJS */' +
                'function exportToExcel(sectionId) {' +
                '    var table = document.getElementById(\'table-\' + sectionId);' +
                '    if (!table) { alert(\'No data to export\'); return; }' +
                '    ' +
                '    var headers = [];' +
                '    var headerCells = table.querySelectorAll(\'thead th\');' +
                '    for (var i = 0; i < headerCells.length; i++) {' +
                '        headers.push(headerCells[i].textContent.replace(/ [▲▼]/g, \'\').trim());' +
                '    }' +
                '    ' +
                '    var data = [headers];' +
                '    var rows = table.querySelectorAll(\'tbody tr\');' +
                '    for (var i = 0; i < rows.length; i++) {' +
                '        var row = rows[i];' +
                '        var rowData = [];' +
                '        var cells = row.querySelectorAll(\'td\');' +
                '        for (var j = 0; j < cells.length; j++) {' +
                '            var cell = cells[j];' +
                '            var val = cell.textContent.trim();' +
                '            if (cell.classList.contains(\'amount\')) {' +
                '                val = parseFloat(val.replace(/[\\$,]/g, \'\')) || 0;' +
                '            }' +
                '            rowData.push(val);' +
                '        }' +
                '        data.push(rowData);' +
                '    }' +
                '    ' +
                '    var ws = XLSX.utils.aoa_to_sheet(data);' +
                '    ' +
                '    /* Apply currency format to amount columns (D, E, F = columns 3, 4, 5) */' +
                '    var range = XLSX.utils.decode_range(ws["!ref"]);' +
                '    for (var R = 1; R <= range.e.r; R++) {' +
                '        for (var C = 3; C <= 5; C++) {' +
                '            var addr = XLSX.utils.encode_cell({r: R, c: C});' +
                '            if (ws[addr]) { ws[addr].z = "\\\"$\\\"#,##0.00"; }' +
                '        }' +
                '    }' +
                '    ' +
                '    var wb = XLSX.utils.book_new();' +
                '    var sheetName = sectionId === \'creditmemos\' ? \'Credit Memos\' : \'Customer Deposits\';' +
                '    XLSX.utils.book_append_sheet(wb, ws, sheetName);' +
                '    ' +
                '    var today = new Date();' +
                '    var dateStr = (today.getMonth()+1) + \'-\' + today.getDate() + \'-\' + today.getFullYear();' +
                '    var fileName = sectionId === \'creditmemos\' ? \'Kitchen_Works_Credit_Memos_\' : \'Kitchen_Works_Deposits_\';' +
                '    XLSX.writeFile(wb, fileName + dateStr + \'.xlsx\');' +
                '}' +
                '' +
                '/* =========================================== */' +
                '/* UNIFIED EXPLAIN MODAL FUNCTIONS            */' +
                '/* =========================================== */' +
                '' +
                'var currentCreditMemoId = null;' +
                'var currentTab = 1;' +
                'var tab1Loaded = false;' +
                'var tab2Loaded = false;' +
                'var tab3Loaded = false;' +
                '' +
                '/* Show Unified Explain Modal */' +
                'function showExplainModal(creditMemoId, customerName, cmNumber, cmAmount, linkedCD, salesOrder) {' +
                '    currentCreditMemoId = creditMemoId;' +
                '    currentTab = 1;' +
                '    tab1Loaded = false;' +
                '    tab2Loaded = false;' +
                '    tab3Loaded = false;' +
                '    tab4Loaded = false;' +
                '    tab5Loaded = false;' +
                '    tab4HasExistingRecord = false;' +
                '    ' +
                '    var modal = document.getElementById("explainModal");' +
                '    var overlay = document.getElementById("explainModalOverlay");' +
                '    ' +
                '    /* Populate Customer Information Section */' +
                '    var customerInfo = document.getElementById("customerInfoSection");' +
                '    ' +
                '    /* Find all CMs for this customer from the table */' +
                '    var customerCMs = [];' +
                '    var totalCMAmount = 0;' +
                '    var cmTable = document.getElementById("table-creditmemos");' +
                '    if (cmTable) {' +
                '        var rows = cmTable.querySelectorAll("tbody tr");' +
                '        for (var i = 0; i < rows.length; i++) {' +
                '            var cells = rows[i].querySelectorAll("td");' +
                '            if (cells.length > 3) {' +
                '                var rowCustomer = cells[3].textContent.trim();' +
                '                if (rowCustomer === customerName) {' +
                '                    var cmNum = cells[1].textContent.trim();' +
                '                    var cmAmtText = cells[4].textContent.trim().replace("$", "").replace(",", "");' +
                '                    var cmAmt = parseFloat(cmAmtText) || 0;' +
                '                    customerCMs.push({ number: cmNum, amount: cmAmt });' +
                '                    totalCMAmount += cmAmt;' +
                '                }' +
                '            }' +
                '        }' +
                '    }' +
                '    ' +
                '    var infoHtml = "<div style=\\"display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:15px;\\\">";' +
                '    infoHtml += "<div><strong style=\\"color:#495057;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;\\">Customer</strong><div style=\\"font-size:16px;font-weight:600;color:#212529;margin-top:4px;\\">" + customerName + "</div></div>";' +
                '    ' +
                '    /* Show combined total if multiple CMs */' +
                '    if (customerCMs.length > 1) {' +
                '        infoHtml += "<div><strong style=\\"color:#495057;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;\\">Combined Overpayments</strong><div style=\\"font-size:16px;font-weight:600;color:#dc3545;margin-top:4px;\\">-" + formatCompCurrency(totalCMAmount) + "</div></div>";' +
                '    }' +
                '    ' +
                '    /* Show individual CMs */' +
                '    for (var i = 0; i < customerCMs.length; i++) {' +
                '        var cm = customerCMs[i];' +
                '        var label = customerCMs.length > 1 ? "CM " + (i + 1) : "Credit Memo";' +
                '        infoHtml += "<div><strong style=\\"color:#495057;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;\\">" + label + "</strong><div style=\\"font-size:16px;font-weight:600;color:#dc3545;margin-top:4px;\\">" + cm.number + " (-" + formatCompCurrency(cm.amount) + ")</div></div>";' +
                '    }' +
                '    if (linkedCD) {' +
                '        infoHtml += "<div><strong style=\\"color:#495057;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;\\">CD #</strong><div style=\\"font-size:16px;font-weight:600;color:#28a745;margin-top:4px;\\">" + linkedCD + "</div></div>";' +
                '    }' +
                '    if (salesOrder) {' +
                '        infoHtml += "<div><strong style=\\"color:#495057;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;\\">Sales Order #</strong><div style=\\"font-size:16px;font-weight:600;color:#007bff;margin-top:4px;\\">" + salesOrder + "</div></div>";' +
                '    }' +
                '    infoHtml += "</div>";' +
                '    customerInfo.innerHTML = infoHtml;' +
                '    customerInfo.style.display = "block";' +
                '    ' +
                '    /* Reset tabs */' +
                '    document.getElementById("tab-btn-1").classList.add("active");' +
                '    document.getElementById("tab-btn-2").classList.remove("active");' +
                '    document.getElementById("tab-btn-3").classList.remove("active");' +
                '    document.getElementById("tab-btn-4").classList.remove("active");' +
                '    document.getElementById("tab-btn-5").classList.remove("active");' +
                '    ' +
                '    document.getElementById("tab-content-1").classList.add("active");' +
                '    document.getElementById("tab-content-2").classList.remove("active");' +
                '    document.getElementById("tab-content-3").classList.remove("active");' +
                '    document.getElementById("tab-content-4").classList.remove("active");' +
                '    document.getElementById("tab-content-5").classList.remove("active");' +
                '    document.getElementById("tab-content-4").innerHTML = "<div class=\\"comparison-loading\\">Loading AI Generated SO Price Changes...</div>";' +
                '    ' +
                '    /* Show modal */' +
                '    modal.classList.add("visible");' +
                '    overlay.classList.add("visible");' +
                '    ' +
                '    /* Load Tab 1 immediately */' +
                '    loadTab1Data();' +
                '}' +
                '' +
                '/* Hide Explain Modal */' +
                'function hideExplainModal() {' +
                '    document.getElementById("explainModal").classList.remove("visible");' +
                '    document.getElementById("explainModalOverlay").classList.remove("visible");' +
                '    currentCreditMemoId = null;' +
                '}' +
                '' +
                '/* Switch Between Tabs */' +
                'function switchExplainTab(tabNumber) {' +
                '    currentTab = tabNumber;' +
                '    ' +
                '    /* Update tab buttons */' +
                '    for (var i = 1; i <= 5; i++) {' +
                '        var btn = document.getElementById("tab-btn-" + i);' +
                '        var content = document.getElementById("tab-content-" + i);' +
                '        if (i === tabNumber) {' +
                '            btn.classList.add("active");' +
                '            content.classList.add("active");' +
                '        } else {' +
                '            btn.classList.remove("active");' +
                '            content.classList.remove("active");' +
                '        }' +
                '    }' +
                '    ' +
                '    /* Load tab data if not already loaded */' +
                '    if (tabNumber === 1 && !tab1Loaded) {' +
                '        loadTab1Data();' +
                '    } else if (tabNumber === 2 && !tab2Loaded) {' +
                '        loadTab2Data();' +
                '    } else if (tabNumber === 3 && !tab3Loaded) {' +
                '        loadTab3Data();' +
                '    } else if (tabNumber === 4 && !tab4Loaded) {' +
                '        loadTab4Data();' +
                '    } else if (tabNumber === 5 && !tab5Loaded) {' +
                '        loadTab5Data();' +
                '    }' +
                '}' +
                '' +
                '/* Load Tab 1: SO↔INV Comparison */' +
                'function loadTab1Data() {' +
                '    var body = document.getElementById("tab-content-1");' +
                '    body.innerHTML = "<div class=\\"comparison-loading\\">Loading SO\u2194INV comparison...</div>";' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { ' +
                '            "Content-Type": "application/json",' +
                '            "X-Requested-With": "XMLHttpRequest"' +
                '        },' +
                '        credentials: "same-origin",' +
                '        body: JSON.stringify({ action: "soInvoiceComparison", creditMemoId: currentCreditMemoId })' +
                '    })' +
                '    .then(function(response) {' +
                '        if (!response.ok) {' +
                '            throw new Error("HTTP " + response.status + ": " + response.statusText);' +
                '        }' +
                '        return response.text();' +
                '    })' +
                '    .then(function(text) {' +
                '        try {' +
                '            var result = JSON.parse(text);' +
                '            if (result.success && result.data) {' +
                '                renderTab1Result(result.data);' +
                '                tab1Loaded = true;' +
                '            } else {' +
                '                body.innerHTML = "<div class=\\"error-msg\\">Error: " + (result.data && result.data.error ? result.data.error : "Unknown error") + "</div>";' +
                '            }' +
                '        } catch (parseErr) {' +
                '            console.error("JSON parse error:", parseErr);' +
                '            console.log("Response text:", text.substring(0, 500));' +
                '            body.innerHTML = "<div class=\\"error-msg\\">Invalid JSON response. Check browser console for details.</div>";' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        console.error("Fetch error:", err);' +
                '        body.innerHTML = "<div class=\\"error-msg\\">Error loading comparison: " + err.message + "</div>";' +
                '    });' +
                '}' +
                '' +
                '/* Render SO details - reusable function */' +
                'function renderSODetails(data) {' +
                '    var html = "", totals = data.totals || {}, soData = data.salesOrder || {}, relatedCMs = data.relatedCreditMemos || [], totalCMAmount = data.totalCMAmount || 0;' +
                '    html += "<div class=\\"comparison-transactions\\">";' +
                '    if (relatedCMs.length > 1) { html += "<div class=\\"comparison-tran-link cm-aggregate\\"><div><div class=\\"comparison-tran-label\\">Credit Memos (" + relatedCMs.length + ")</div><div class=\\"comparison-tran-value\\">Combined</div><div class=\\"comparison-tran-amount\\">-" + formatCompCurrency(totalCMAmount) + "</div></div></div>"; for (var c = 0; c < relatedCMs.length; c++) { html += "<a href=\\"/app/accounting/transactions/custcred.nl?id=" + relatedCMs[c].id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link cm-individual\\"><div><div class=\\"comparison-tran-label\\">CM " + (c+1) + "</div><div class=\\"comparison-tran-value\\">" + relatedCMs[c].tranid + "</div><div class=\\"comparison-tran-amount\\">-" + formatCompCurrency(relatedCMs[c].amount) + "</div></div></a>"; } } else if (relatedCMs.length === 1) { html += "<a href=\\"/app/accounting/transactions/custcred.nl?id=" + relatedCMs[0].id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link\\"><div><div class=\\"comparison-tran-label\\">Credit Memo</div><div class=\\"comparison-tran-value\\">" + relatedCMs[0].tranid + "</div><div class=\\"comparison-tran-amount\\">-" + formatCompCurrency(relatedCMs[0].amount) + "</div></div></a>"; }' +
                '    html += "<a href=\\"/app/accounting/transactions/salesord.nl?id=" + soData.id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link\\"><div><div class=\\"comparison-tran-label\\">Sales Order</div><div class=\\"comparison-tran-value\\">" + (soData.tranid || "-") + "</div><div class=\\"comparison-tran-amount\\">" + formatCompCurrency(Math.abs(soData.foreigntotal || 0)) + "</div></div></a>";' +
                '    var invoices = data.invoices || []; for (var i = 0; i < invoices.length; i++) { html += "<a href=\\"/app/accounting/transactions/custinvc.nl?id=" + invoices[i].id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link\\"><div><div class=\\"comparison-tran-label\\">Invoice " + (i+1) + "</div><div class=\\"comparison-tran-value\\">" + invoices[i].tranid + "</div><div class=\\"comparison-tran-amount\\">" + formatCompCurrency(Math.abs(invoices[i].amount)) + "</div></div></a>"; }' +
                '    html += "</div>"; var mismatchVar = totals.mismatchVariance || 0, unbilledVar = totals.unbilledVariance || 0; html += "<div class=\\"comparison-summary\\"><div class=\\"comparison-card\\"><div class=\\"comparison-card-label\\">SO Grand Total</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(totals.soGrandTotal || 0) + "</div><div class=\\"comparison-card-sub\\">Lines: " + formatCompCurrency(totals.soLineTotal || 0) + " | Tax: " + formatCompCurrency(totals.soTaxTotal || 0) + "</div></div><div class=\\"comparison-card\\"><div class=\\"comparison-card-label\\">Invoice Grand Total</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(totals.invoiceGrandTotal || 0) + "</div><div class=\\"comparison-card-sub\\">Lines: " + formatCompCurrency(totals.invoiceLineTotal || 0) + " | Tax: " + formatCompCurrency(totals.invoiceTaxTotal || 0) + "</div></div><div class=\\"comparison-card " + (Math.abs(mismatchVar) < 0.01 ? "success" : "error") + "\\"><div class=\\"comparison-card-label\\">Mismatch Variance</div><div class=\\"comparison-card-value\\">" + formatCompCurrencyWithSign(mismatchVar) + "</div><div class=\\"comparison-card-sub\\">Unbilled: " + formatCompCurrencyWithSign(unbilledVar) + "</div></div>";' +
                '    if (relatedCMs.length > 0) { html += "<div class=\\"comparison-card " + (Math.abs(Math.abs(mismatchVar) - totalCMAmount) < 0.005 ? "success" : "warning") + "\\"><div class=\\"comparison-card-label\\">CM Total</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(totalCMAmount) + "</div></div>"; }' +
                '    html += "</div><div class=\\"comparison-conclusion " + (Math.abs(mismatchVar) < 0.01 ? "match" : "mismatch") + "\\">" + (data.conclusion || "") + "</div>"; var problemItems = data.problemItems || []; if (problemItems.length > 0) { html += "<div style=\\"margin-top:20px;padding:15px;background:#fff3e0;border:1px solid #f57c00;border-radius:6px;\\"><strong style=\\"color:#e65100;\\">\u26A0\uFE0F " + problemItems.length + " Problem Item(s):</strong><ul style=\\"margin:10px 0 0 20px;\\">"; for (var pi = 0; pi < problemItems.length; pi++) { html += "<li><strong>" + problemItems[pi].itemName + "</strong> - " + problemItems[pi].status + "</li>"; } html += "</ul></div>"; }' +
                '    html += "<div class=\\"comparison-table-container\\"><table class=\\"comparison-table\\"><thead><tr><th>Item</th><th>Description</th><th>SO #</th><th class=\\"amount\\">SO Qty</th><th class=\\"amount\\">SO Rate</th><th class=\\"amount\\">SO Amount</th><th>Invoice(s)</th><th class=\\"amount\\">Inv Qty</th><th class=\\"amount\\">Inv Rate</th><th class=\\"amount\\">Inv Amount</th><th class=\\"amount\\">Mismatch</th><th class=\\"amount\\">Unbilled</th><th>Status</th></tr></thead><tbody>";' +
                '    var compTable = data.comparisonTable || []; for (var i = 0; i < compTable.length; i++) { var row = compTable[i], statusClass = row.status === "Match" ? "match" : (row.status === "NOT INVOICED" ? "not-invoiced" : (row.status === "DISCOUNT" ? "discount" : "mismatch")); html += "<tr><td>" + (row.itemName || "-") + "</td><td>" + (row.lineDescription || "-") + "</td><td>" + (soData.tranid || "-") + "</td><td class=\\"amount\\">" + (row.soQty || 0) + "</td><td class=\\"amount\\">" + formatCompCurrencySigned(row.soRate || 0) + "</td><td class=\\"amount\\">" + formatCompCurrencySigned(row.soAmount || 0) + "</td><td>" + (row.invoiceNumbers || "-") + "</td><td class=\\"amount\\">" + (row.invoiceQty || 0) + "</td><td class=\\"amount\\">" + formatCompCurrencySigned(row.invoiceRate || 0) + "</td><td class=\\"amount\\">" + formatCompCurrencySigned(row.invoiceAmount || 0) + "</td><td class=\\"amount\\">" + formatCompCurrencyWithSign(row.mismatchVariance || 0) + "</td><td class=\\"amount\\">" + formatCompCurrencyWithSign(row.unbilledVariance || 0) + "</td><td class=\\"" + statusClass + "\\">" + row.status + "</td></tr>"; }' +
                '    html += "<tr class=\\"comparison-totals\\"><td colspan=\\"5\\"><strong>TOTALS</strong></td><td class=\\"amount\\">" + formatCompCurrency(totals.soLineTotal || 0) + "</td><td></td><td></td><td></td><td class=\\"amount\\">" + formatCompCurrency(totals.invoiceLineTotal || 0) + "</td><td class=\\"amount\\">" + formatCompCurrencyWithSign(mismatchVar) + "</td><td class=\\"amount\\">" + formatCompCurrencyWithSign(unbilledVar) + "</td><td></td></tr></tbody></table></div>";' +
                '    return html;' +
                '}' +
                '' +
                '/* Original single-SO view */' +
                'function renderOriginalSingleSOView(data) {' +
                '    window.currentComparisonData = { mismatchVariance: data.totals ? data.totals.mismatchVariance : 0, unbilledVariance: data.totals ? data.totals.unbilledVariance : 0, totalVariance: data.totals ? data.totals.totalVariance : 0 };' +
                '    return "<div style=\\"margin:0 0 20px;padding:15px;background:#e3f2fd;border-left:4px solid #1976d2;border-radius:4px;\\"><strong>What This Analysis Shows:</strong> Line-by-line SO vs Invoice comparison.</div>" + renderSODetails(data);' +
                '}' +
                '' +
                '/* Load Tab 2: Cross-SO Analysis */' +
                'function loadTab2Data() {' +
                '    var body = document.getElementById("tab-content-2");' +
                '    body.innerHTML = "<div class=\\"comparison-loading\\">Loading Cross-SO analysis...</div>";' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ action: "crossSOAnalysis", creditMemoId: currentCreditMemoId })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(result) {' +
                '        if (result.success && result.data) {' +
                '            renderTab2Result(result.data);' +
                '            tab2Loaded = true;' +
                '        } else {' +
                '            body.innerHTML = "<div class=\\"error-msg\\">Error: " + (result.data && result.data.error ? result.data.error : "Unknown error") + "</div>";' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\">Error loading analysis: " + err.message + "</div>";' +
                '    });' +
                '}' +
                '' +
                '/* Load Tab 3: Overpayment Summary */' +
                'function loadTab3Data() {' +
                '    var body = document.getElementById("tab-content-3");' +
                '    body.innerHTML = "<div class=\\"comparison-loading\\">Loading Overpayment Summary...</div>";' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ action: "overpaymentSummary", creditMemoId: currentCreditMemoId })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(result) {' +
                '        if (result.success && result.data) {' +
                '            renderTab3Result(result.data);' +
                '            tab3Loaded = true;' +
                '        } else {' +
                '            body.innerHTML = "<div class=\\"error-msg\\">Error: " + (result.data && result.data.error ? result.data.error : "Unknown error") + "</div>";' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\">Error loading summary: " + err.message + "</div>";' +
                '    });' +
                '}' +
                '' +
                '/* Trigger AI Analysis */' +
                'var tab4Loaded = false;' +
                'var tab4HasExistingRecord = false;' +
                '' +
                '/* Load Tab 4: AI Generated SO Price Changes */' +
                'function loadTab4Data() {' +
                '    var body = document.getElementById("tab-content-4");' +
                '    body.innerHTML = "<div class=\\"comparison-loading\\">Loading AI Generated SO Price Changes...</div>";' +
                '    ' +
                '    /* First check if there is an existing AI record */' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ action: "loadAIAnalysis", creditMemoId: currentCreditMemoId })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(result) {' +
                '        if (result.success && result.aiAnalysis && result.aiAnalysis.found) {' +
                '            /* Existing record found - display it with Re-Run button */' +
                '            tab4HasExistingRecord = true;' +
                '            tab4Loaded = true;' +
                '            var data = result.aiAnalysis;' +
                '            renderTab4ResultWithRerun({' +
                '                haikuResponse: data.haikuResponse,' +
                '                savedRecordId: data.recordId,' +
                '                createdDate: data.createdDate' +
                '            });' +
                '        } else {' +
                '            /* No existing record - auto-run AI analysis */' +
                '            tab4HasExistingRecord = false;' +
                '            runTab4AIAnalysis();' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\">Error checking for existing analysis: " + err.message + "</div>";' +
                '    });' +
                '}' +
                '' +
                '/* Run AI Analysis for Tab 4 */' +
                'function runTab4AIAnalysis() {' +
                '    var body = document.getElementById("tab-content-4");' +
                '    body.innerHTML = "<div class=\\"comparison-loading\\">Running AI Analysis... This may take 10-30 seconds...</div>";' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ action: "aiAnalysis", creditMemoId: currentCreditMemoId, comparisonData: window.currentComparisonData })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(result) {' +
                '        if (result.success && result.data) {' +
                '            tab4Loaded = true;' +
                '            tab4HasExistingRecord = true;' +
                '            renderTab4ResultWithRerun(result.data);' +
                '        } else {' +
                '            body.innerHTML = "<div class=\\"error-msg\\">Error: " + (result.data && result.data.error ? result.data.error : "Unknown error") + "</div>";' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\">Error running AI analysis: " + err.message + "</div>";' +
                '    });' +
                '}' +
                '' +
                '/* Legacy trigger function - now calls loadTab4Data */' +
                'function triggerAIAnalysis() {' +
                '    switchExplainTab(4);' +
                '}' +
                '' +
                '/* Load Tab 5: Sales Order Totals (All) */' +
                'function loadTab5Data() {' +
                '    var body = document.getElementById("tab-content-5");' +
                '    body.innerHTML = "<div class=\\"comparison-loading\\">Loading Sales Order Totals...</div>";' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ action: "salesOrderTotals", creditMemoId: currentCreditMemoId })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(result) {' +
                '        if (result.success && result.data) {' +
                '            renderTab5Result(result.data);' +
                '            tab5Loaded = true;' +
                '        } else {' +
                '            body.innerHTML = "<div class=\\"error-msg\\">Error: " + (result.data && result.data.error ? result.data.error : "Unknown error") + "</div>";' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\">Error loading sales order totals: " + err.message + "</div>";' +
                '    });' +
                '}' +
                '' +
                '/* Render Tab 5: Sales Order Totals (All) Result */' +
                'function renderTab5Result(data) {' +
                '    var body = document.getElementById("tab-content-5");' +
                '    ' +
                '    if (data.error) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\"><strong>Error:</strong> " + data.error + "</div>";' +
                '        return;' +
                '    }' +
                '    ' +
                '    var html = "";' +
                '    ' +
                '    /* Helper Text */' +
                '    html += "<div style=\\"margin:0 0 20px;padding:15px;background:#e8f5e9;border-left:4px solid #4CAF50;border-radius:4px;\\">";' +
                '    html += "<strong>Customer Sales Orders Summary:</strong> This view shows all sales orders for <strong>" + (data.customerName || "this customer") + "</strong>. ";' +
                '    html += "Review tax discrepancies and billing variances to identify potential issues with invoicing.";' +
                '    html += "</div>";' +
                '    ' +
                '    /* Summary Cards */' +
                '    var summary = data.summary || {};' +
                '    var totalSalesOrders = summary.totalSalesOrders || 0;' +
                '    var totalOrderAmount = summary.totalOrderAmount || 0;' +
                '    var totalBilledAmount = summary.totalBilledAmount || 0;' +
                '    var totalUnbilledAmount = summary.totalUnbilledAmount || 0;' +
                '    var totalOrderNonTaxAmount = summary.totalOrderNonTaxAmount || 0;' +
                '    var totalOrderTaxAmount = summary.totalOrderTaxAmount || 0;' +
                '    var totalBilledNonTaxAmount = summary.totalBilledNonTaxAmount || 0;' +
                '    var totalBilledTaxAmount = summary.totalBilledTaxAmount || 0;' +
                '    var taxDiscrepancyCount = summary.taxDiscrepancyCount || 0;' +
                '    var billingVarianceCount = summary.billingVarianceCount || 0;' +
                '    ' +
                '    html += "<div class=\\"comparison-summary\\">";' +
                '    html += "<div class=\\"comparison-card info\\"><div class=\\"comparison-card-label\\">Total Sales Orders</div><div class=\\"comparison-card-value\\">" + totalSalesOrders + "</div></div>";' +
                '    html += "<div class=\\"comparison-card\\"><div class=\\"comparison-card-label\\">SO Total (with Tax)</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(totalOrderAmount) + "</div></div>";' +
                '    html += "<div class=\\"comparison-card\\"><div class=\\"comparison-card-label\\">Invoiced (with Tax)</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(totalBilledAmount) + "</div></div>";' +
                '    html += "<div class=\\"comparison-card " + (totalUnbilledAmount > 0.01 ? "warning" : "success") + "\\"><div class=\\"comparison-card-label\\">Unbilled Amount</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(totalUnbilledAmount) + "</div></div>";' +
                '    html += "</div>";' +
                '    ' +
                '    /* Variance Alerts */' +
                '    if (taxDiscrepancyCount > 0 || billingVarianceCount > 0) {' +
                '        html += "<div style=\\"margin:20px 0;padding:15px;background:#fff3e0;border:1px solid #f57c00;border-radius:6px;\\">";' +
                '        html += "<strong style=\\"color:#e65100;\\">⚠️ Issues Detected:</strong>";' +
                '        html += "<ul style=\\"margin:10px 0 0 20px;padding:0;\\">";' +
                '        if (taxDiscrepancyCount > 0) {' +
                '            html += "<li><strong>" + taxDiscrepancyCount + " Tax Discrepanc" + (taxDiscrepancyCount === 1 ? "y" : "ies") + "</strong>: SO tax amounts differ from invoiced tax amounts</li>";' +
                '        }' +
                '        if (billingVarianceCount > 0) {' +
                '            html += "<li><strong>" + billingVarianceCount + " Billing Variance" + (billingVarianceCount === 1 ? "" : "s") + "</strong>: Fully billed orders with amount mismatches (Status G)</li>";' +
                '        }' +
                '        html += "</ul></div>";' +
                '    }' +
                '    ' +
                '    /* Sales Orders Table */' +
                '    html += "<div class=\\"comparison-table-container\\">";' +
                '    html += "<table class=\\"comparison-table so-totals-table\\">";' +
                '    html += "<thead><tr>";' +
                '    html += "<th>SO #</th>";' +
                '    html += "<th>SO Date</th>";' +
                '    html += "<th>Status</th>";' +
                '    html += "<th class=\\"amount group-total-start\\">SO Total</th>";' +
                '    html += "<th class=\\"amount group-total-end\\">INV Total</th>";' +
                '    html += "<th class=\\"amount\\">SO Unbilled</th>";' +
                '    html += "<th class=\\"amount group-nontax-start\\">SO Non-Tax</th>";' +
                '    html += "<th class=\\"amount group-nontax-end\\">INV Non-Tax</th>";' +
                '    html += "<th class=\\"amount group-tax-start\\">SO Tax</th>";' +
                '    html += "<th class=\\"amount group-tax-end\\">INV Tax</th>";' +
                '    html += "<th>Created Date</th>";' +
                '    html += "<th>Created By</th>";' +
                '    html += "<th>Issues</th>";' +
                '    html += "</tr></thead>";' +
                '    html += "<tbody>";' +
                '    ' +
                '    var salesOrders = data.salesOrders || [];' +
                '    for (var i = 0; i < salesOrders.length; i++) {' +
                '        var so = salesOrders[i];' +
                '        var soTranid = so.soTranId || "-";' +
                '        var soId = so.soId || "";' +
                '        var trandate = so.soDate ? so.soDate.split(" ")[0] : "-";' +
                '        var createdDate = so.soCreatedDate ? so.soCreatedDate.split(" ")[0] : "-";' +
                '        var status = so.soStatusText || "-";' +
                '        var creator = so.soCreatedBy || "-";' +
                '        var soNontax = so.soNontaxAmount || 0;' +
                '        var soTax = so.soTaxAmount || 0;' +
                '        var soTotal = so.soTotalAmount || 0;' +
                '        var invNontax = so.invNontaxBilled || 0;' +
                '        var invTax = so.invTaxBilled || 0;' +
                '        var invTotal = so.invTotalBilled || 0;' +
                '        var unbilled = so.soTotalUnbilled || 0;' +
                '        var soTaxPct = so.soTaxPct || 0;' +
                '        var invTaxPct = so.invTaxPct || 0;' +
                '        var hasTaxDisc = so.hasTaxDiscrepancy || false;' +
                '        var hasBillVar = so.hasBillingVariance || false;' +
                '        ' +
                '        var issuesHtml = "";' +
                '        if (hasBillVar) {' +
                '            issuesHtml += "<span class=\\"status-badge\\" style=\\"background:#9c27b0;color:#fff;\\" title=\\"Fully billed but amounts don\'t match\\">BILL VAR</span> ";' +
                '        }' +
                '        if (hasTaxDisc) {' +
                '            issuesHtml += "<span class=\\"status-badge status-warning\\" title=\\"Tax amounts differ between SO and invoices\\">TAX DISC</span>";' +
                '        }' +
                '        if (!hasBillVar && !hasTaxDisc) {' +
                '            issuesHtml = "<span style=\\"color:#4caf50;\\">✓</span>";' +
                '        }' +
                '        ' +
                '        var rowClass = hasBillVar ? "bill-variance-row" : "";' +
                '        ' +
                '        html += "<tr class=\\"" + rowClass + "\\">";' +
                '        html += "<td><a href=\\"/app/accounting/transactions/salesord.nl?id=" + soId + "\\" target=\\"_blank\\">" + soTranid + "</a></td>";' +
                '        html += "<td>" + trandate + "</td>";' +
                '        html += "<td>" + status + "</td>";' +
                '        html += "<td class=\\"amount group-total-start\\">" + formatCompCurrency(soTotal) + "</td>";' +
                '        html += "<td class=\\"amount group-total-end\\">" + formatCompCurrency(invTotal) + "</td>";' +
                '        html += "<td class=\\"amount " + (unbilled > 0.01 ? "unbilled-amount" : "") + "\\">" + formatCompCurrency(unbilled) + "</td>";' +
                '        html += "<td class=\\"amount group-nontax-start\\">" + formatCompCurrency(soNontax) + "</td>";' +
                '        html += "<td class=\\"amount group-nontax-end\\">" + formatCompCurrency(invNontax) + "</td>";' +
                '        html += "<td class=\\"amount group-tax-start\\">" + formatCompCurrency(soTax) + "<br><span style=\\"color:#666;font-size:11px;\\">(" + soTaxPct + "%)</span></td>";' +
                '        html += "<td class=\\"amount group-tax-end\\">" + formatCompCurrency(invTax) + "<br><span style=\\"color:#666;font-size:11px;\\">(" + invTaxPct + "%)</span></td>";' +
                '        html += "<td>" + createdDate + "</td>";' +
                '        html += "<td>" + creator + "</td>";' +
                '        html += "<td>" + issuesHtml + "</td>";' +
                '        html += "</tr>";' +
                '    }' +
                '    ' +
                '    /* Totals Row */' +
                '    html += "<tr class=\\"comparison-totals\\">";' +
                '    html += "<td colspan=\\"3\\"><strong>TOTALS</strong></td>";' +
                '    html += "<td class=\\"amount group-total-start\\">" + formatCompCurrency(totalOrderAmount) + "</td>";' +
                '    html += "<td class=\\"amount group-total-end\\">" + formatCompCurrency(totalBilledAmount) + "</td>";' +
                '    html += "<td class=\\"amount\\">" + formatCompCurrency(totalUnbilledAmount) + "</td>";' +
                '    html += "<td class=\\"amount group-nontax-start\\">" + formatCompCurrency(totalOrderNonTaxAmount) + "</td>";' +
                '    html += "<td class=\\"amount group-nontax-end\\">" + formatCompCurrency(totalBilledNonTaxAmount) + "</td>";' +
                '    html += "<td class=\\"amount group-tax-start\\">" + formatCompCurrency(totalOrderTaxAmount) + "</td>";' +
                '    html += "<td class=\\"amount group-tax-end\\">" + formatCompCurrency(totalBilledTaxAmount) + "</td>";' +
                '    html += "<td colspan=\\"3\\"></td>";' +
                '    html += "</tr>";' +
                '    ' +
                '    html += "</tbody></table></div>";' +
                '    ' +
                '    body.innerHTML = html;' +
                '}' +
                '' +
                '/* Render Tab 1: SO↔INV Comparison Result */' +
                'function renderTab1Result(data) {' +
                '    console.log("renderTab1Result called with data:", data);' +
                '    var body = document.getElementById("tab-content-1");' +
                '    console.log("tab-content-1 element:", body);' +
                '    ' +
                '    if (!body) {' +
                '        console.error("ERROR: tab-content-1 element not found!");' +
                '        alert("Modal element not found. Please refresh and try again.");' +
                '        return;' +
                '    }' +
                '    ' +
                '    if (data.error) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\"><strong>Error:</strong> " + data.error + "</div>";' +
                '        return;' +
                '    }' +
                '    ' +
                '    var html = "";' +
                '    ' +
                '    /* Check if we have multi-SO data */' +
                '    var hasMultiSOData = data.soComparisons && data.soComparisons.length > 0;' +
                '    ' +
                '    if (hasMultiSOData) {' +
                '        /* NEW MULTI-SO VIEW */' +
                '        var customerTotals = data.customerTotals || {};' +
                '        var soComparisons = data.soComparisons || [];' +
                '        ' +
                '        /* Summary Table */' +
                '        html += "<div class=\\"summary-table-wrapper\\">";' +
                '        html += "<h3>\u{1F4CA} SUMMARY TABLE - All Sales Orders at a Glance</h3>";' +
                '        html += "<table class=\\"summary-table\\">";' +
                '        html += "<thead><tr>";' +
                '        html += "<th>SO Number</th>";' +
                '        html += "<th>Status</th>";' +
                '        html += "<th class=\\"amount\\">Mismatch</th>";' +
                '        html += "<th class=\\"amount\\">Unbilled</th>";' +
                '        html += "<th class=\\"amount\\">Total Variance</th>";' +
                '        html += "<th>Status</th>";' +
                '        html += "</tr></thead><tbody>";' +
                '        ' +
                '        for (var sumIdx = 0; sumIdx < soComparisons.length; sumIdx++) {' +
                '            var sumComp = soComparisons[sumIdx];' +
                '            if (sumComp.error) continue;' +
                '            ' +
                '            var sumTotals = sumComp.totals || {};' +
                '            var sumMismatch = sumTotals.mismatchVariance || 0;' +
                '            var sumUnbilled = sumTotals.unbilledVariance || 0;' +
                '            var sumTotal = sumMismatch + sumUnbilled;' +
                '            var sumSO = sumComp.salesOrder || {};' +
                '            ' +
                '            var sumStatusIcon = "";' +
                '            var sumStatusText = "";' +
                '            if (Math.abs(sumMismatch) > 0.01) {' +
                '                sumStatusIcon = "\u26A0\uFE0F";' +
                '                sumStatusText = "HAS MISMATCH";' +
                '            } else if (Math.abs(sumUnbilled) > 0.01) {' +
                '                sumStatusIcon = "\u{1F4CB}";' +
                '                sumStatusText = "HAS UNBILLED";' +
                '            } else {' +
                '                sumStatusIcon = "\u2705";' +
                '                sumStatusText = "NO ISSUES";' +
                '            }' +
                '            ' +
                '            var sumSourceMarker = sumComp.isSource ? " \u2B50" : "";' +
                '            ' +
                '            html += "<tr>";' +
                '            html += "<td><strong>" + (sumSO.tranid || "-") + sumSourceMarker + "</strong></td>";' +
                '            html += "<td>" + (sumComp.statusDisplay || sumComp.status || "-") + "</td>";' +
                '            html += "<td class=\\"amount\\">" + formatCompCurrencyWithSign(sumMismatch) + "</td>";' +
                '            html += "<td class=\\"amount\\">" + formatCompCurrencyWithSign(sumUnbilled) + "</td>";' +
                '            html += "<td class=\\"amount\\">" + formatCompCurrencyWithSign(sumTotal) + "</td>";' +
                '            html += "<td><span class=\\"status-icon\\">" + sumStatusIcon + "</span> " + sumStatusText + "</td>";' +
                '            html += "</tr>";' +
                '        }' +
                '        ' +
                '        /* Add totals row */' +
                '        html += "<tr style=\\"font-weight:bold;background:#f8f9fa;border-top:2px solid #dee2e6;\\">";' +
                '        html += "<td colspan=\\"2\\">TOTALS</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencyWithSign(customerTotals.totalMismatchVariance) + "</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencyWithSign(customerTotals.totalUnbilledVariance) + "</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencyWithSign(customerTotals.grandTotalVariance) + "</td>";' +
                '        html += "<td></td>";' +
                '        html += "</tr>";' +
                '        ' +
                '        html += "</tbody></table></div>";' +
                '        ' +
                '        /* Detailed Comparison Divider */' +
                '        html += "<div class=\\"detailed-comparison-divider\\"></div>";' +
                '        ' +
                '        /* Render each SO section */' +
                '        for (var soIdx = 0; soIdx < soComparisons.length; soIdx++) {' +
                '            var comp = soComparisons[soIdx];' +
                '            if (comp.error) {' +
                '                html += "<div class=\\"so-section\\"><div class=\\"error-msg\\">Error loading SO: " + comp.error + "</div></div>";' +
                '                continue;' +
                '            }' +
                '            ' +
                '            html += renderSOSection(comp, soIdx);' +
                '        }' +
                '        ' +
                '    } else {' +
                '        /* FALLBACK: Original single-SO view (for backward compatibility) */' +
                '        html += renderOriginalSingleSOView(data);' +
                '    }' +
                '    ' +
                '    console.log("About to set body.innerHTML, html length:", html.length);' +
                '    body.innerHTML = html;' +
                '    console.log("Successfully set body.innerHTML");' +
                '}' +
                '' +
                '/* Render individual SO section */' +
                'function renderSOSection(comp, index) {' +
                '    var html = "";' +
                '    var totals = comp.totals || {};' +
                '    var soData = comp.salesOrder || {};' +
                '    var mismatchVar = totals.mismatchVariance || 0;' +
                '    var unbilledVar = totals.unbilledVariance || 0;' +
                '    ' +
                '    /* Determine section class */' +
                '    var sectionClass = "so-section";' +
                '    if (comp.isSource) {' +
                '        sectionClass += " source-so";' +
                '    } else if (Math.abs(mismatchVar) > 0.01) {' +
                '        sectionClass += " has-mismatch";' +
                '    } else if (Math.abs(unbilledVar) > 0.01) {' +
                '        sectionClass += " has-unbilled";' +
                '    } else {' +
                '        sectionClass += " no-issues";' +
                '    }' +
                '    ' +
                '    /* Section header */' +
                '    html += "<div class=\\"" + sectionClass + "\\">";' +
                '    html += "<div class=\\"so-section-header\\">";' +
                '    ' +
                '    /* Icon */' +
                '    var icon = comp.isSource ? "\u2B50" : (Math.abs(mismatchVar) > 0.01 ? "\u26A0\uFE0F" : (Math.abs(unbilledVar) > 0.01 ? "\u{1F4CB}" : "\u2705"));' +
                '    html += "<div class=\\"so-section-icon\\">" + icon + "</div>";' +
                '    ' +
                '    /* Title */' +
                '    html += "<div class=\\"so-section-title\\">";' +
                '    var titlePrefix = comp.isSource ? "SOURCE SALES ORDER: " : "SALES ORDER: ";' +
                '    html += "<h3>" + titlePrefix + (soData.tranid || "-") + "</h3>";' +
                '    html += "<div class=\\"so-status\\">" + (comp.statusDisplay || comp.status || "-") + "</div>";' +
                '    html += "</div>";' +
                '    ' +
                '    /* Badges */' +
                '    html += "<div class=\\"so-section-badges\\">";' +
                '    if (comp.isSource) {' +
                '        html += "<span class=\\"so-badge source\\">SOURCE</span>";' +
                '    }' +
                '    if (Math.abs(mismatchVar) > 0.01) {' +
                '        html += "<span class=\\"so-badge mismatch\\">MISMATCH: " + formatCompCurrency(Math.abs(mismatchVar)) + "</span>";' +
                '    }' +
                '    if (Math.abs(unbilledVar) > 0.01) {' +
                '        html += "<span class=\\"so-badge unbilled\\">UNBILLED: " + formatCompCurrency(Math.abs(unbilledVar)) + "</span>";' +
                '    }' +
                '    if (Math.abs(mismatchVar) < 0.01 && Math.abs(unbilledVar) < 0.01) {' +
                '        html += "<span class=\\"so-badge clean\\">\u2705 NO ISSUES</span>";' +
                '    }' +
                '    html += "</div>";' +
                '    html += "</div>";' +
                '    ' +
                '    /* Use original rendering logic for the details */' +
                '    html += renderSODetails(comp);' +
                '    ' +
                '    html += "</div>";' +
                '    return html;' +
                '}' +
                '' +
                '/* Render SO comparison details (reusable) */' +
                'function renderSODetails(data) {' +
                '    var html = "";' +
                '    ' +
                '    /* Transaction Links Container */' +
                '    html += "<div class=\\"comparison-tran-grid\\">"' + ';' +
                '    ' +
                '    var totals = data.totals || {};' +
                '    var allCMs = data.relatedCreditMemos || [];' +
                '    var totalCMAmount = data.totalCMAmount || 0;' +
                '    if (allCMs.length > 1) {' +
                '        /* Multiple CMs - show aggregate card first */' +
                '        html += "<div class=\\"comparison-tran-link cm-aggregate\\">";' +
                '        html += "<div><div class=\\"comparison-tran-label\\">Credit Memos (" + allCMs.length + " total)</div><div class=\\"comparison-tran-value\\">Combined Overpayments</div><div class=\\"comparison-tran-amount\\">-" + formatCompCurrency(totalCMAmount) + "</div></div>";' +
                '        html += "</div>";' +
                '        /* Individual CM links */' +
                '        for (var c = 0; c < allCMs.length; c++) {' +
                '            var cm = allCMs[c];' +
                '            html += "<a href=\\"/app/accounting/transactions/custcred.nl?id=" + cm.id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link cm-individual\\">";' +
                '            html += "<div><div class=\\"comparison-tran-label\\">CM " + (c+1) + "</div><div class=\\"comparison-tran-value\\">" + cm.tranid + "</div><div class=\\"comparison-tran-amount\\">-" + formatCompCurrency(cm.amount) + "</div></div>";' +
                '            html += "</a>";' +
                '        }' +
                '    } else if (allCMs.length === 1) {' +
                '        /* Single CM */' +
                '        var singleCM = allCMs[0];' +
                '        html += "<a href=\\"/app/accounting/transactions/custcred.nl?id=" + singleCM.id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link\\">";' +
                '        html += "<div><div class=\\"comparison-tran-label\\">Credit Memo</div><div class=\\"comparison-tran-value\\">" + singleCM.tranid + "</div><div class=\\"comparison-tran-amount\\">-" + formatCompCurrency(singleCM.amount) + "</div></div>";' +
                '        html += "</a>";' +
                '    }' +
                '    html += "</div>";' +
                '    ' +
                '    /* Summary Cards - Totals are already calculated with correct signs from server */' +
                '    var totals = data.totals || {};' +
                '    var soGrandPos = totals.soGrandTotal || 0;' +
                '    var soLinePos = totals.soLineTotal || 0;' +
                '    var soTaxPos = totals.soTaxTotal || 0;' +
                '    var invGrandPos = totals.invoiceGrandTotal || 0;' +
                '    var invLinePos = totals.invoiceLineTotal || 0;' +
                '    var invTaxPos = totals.invoiceTaxTotal || 0;' +
                '    var totalVar = totals.totalVariance || 0;' +
                '    var mismatchVar = totals.mismatchVariance || 0;' +
                '    var unbilledVar = totals.unbilledVariance || 0;' +
                '    ' +
                '    /* Calculate CM total and compare to MISMATCH variance (not total) */' +
                '    var cmTotalAmt = data.totalCMAmount || 0;' +
                '    var mismatchVsCm = Math.abs(Math.abs(mismatchVar) - cmTotalAmt);' +
                '    var cmMatchClass = "warning";' +
                '    var cmMatchNote = "";' +
                '    ' +
                '    if (mismatchVsCm < 0.005) {' +
                '        /* Exact match (no tolerance) */' +
                '        cmMatchClass = "success";' +
                '        cmMatchNote = "✓ Matches Mismatch Variance";' +
                '    } else if (Math.abs(mismatchVar) > 0.01) {' +
                '        /* Check if difference is approximately 6% tax on the mismatch */' +
                '        var expectedTax = Math.abs(mismatchVar) * 0.06;' +
                '        var taxDiff = Math.abs(mismatchVsCm - expectedTax);' +
                '        ' +
                '        if (taxDiff <= 0.02) {' +
                '            /* Difference is exactly 6% of mismatch - likely tax issue */' +
                '            cmMatchClass = "info";' +
                '            cmMatchNote = "⚠️ Diff " + formatCompCurrency(mismatchVsCm) + " = 6% tax on mismatch. Check tax change.";' +
                '        } else {' +
                '            cmMatchNote = "Diff from Mismatch: " + formatCompCurrency(mismatchVsCm);' +
                '        }' +
                '    } else {' +
                '        cmMatchNote = "Diff from Mismatch: " + formatCompCurrency(mismatchVsCm);' +
                '    }' +
                '    ' +
                '    /* Variance card class based on mismatch (the error portion) */' +
                '    var varianceClass = Math.abs(mismatchVar) < 0.01 ? "success" : "error";' +
                '    ' +
                '    html += "<div class=\\"comparison-summary\\">";' +
                '    html += "<div class=\\"comparison-card\\"><div class=\\"comparison-card-label\\">SO Grand Total</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(soGrandPos) + "</div><div class=\\"comparison-card-sub\\">Lines: " + formatCompCurrency(soLinePos) + " | Tax: " + formatCompCurrency(soTaxPos) + "</div></div>";' +
                '    html += "<div class=\\"comparison-card\\"><div class=\\"comparison-card-label\\">Invoice Grand Total</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(invGrandPos) + "</div><div class=\\"comparison-card-sub\\">Lines: " + formatCompCurrency(invLinePos) + " | Tax: " + formatCompCurrency(invTaxPos) + "</div></div>";' +
                '    html += "<div class=\\"comparison-card " + varianceClass + "\\"><div class=\\"comparison-card-label\\">Mismatch Variance</div><div class=\\"comparison-card-value\\">" + formatCompCurrencyWithSign(mismatchVar) + "</div><div class=\\"comparison-card-sub\\">Unbilled: " + formatCompCurrencyWithSign(unbilledVar) + " | Total: " + formatCompCurrencyWithSign(totalVar) + "</div></div>";' +
                '    html += "<div class=\\"comparison-card " + cmMatchClass + "\\"><div class=\\"comparison-card-label\\">CM Total (Overpayments)</div><div class=\\"comparison-card-value\\">" + formatCompCurrency(cmTotalAmt) + "</div><div class=\\"comparison-card-sub\\">" + cmMatchNote + "</div></div>";' +
                '    html += "</div>";' +
                '    ' +
                '    /* Conclusion - moved above table */' +
                '    var conclusionClass = Math.abs(mismatchVar) < 0.01 ? "match" : "mismatch";' +
                '    html += "<div class=\\"comparison-conclusion " + conclusionClass + "\\">" + (data.conclusion || "Analysis complete") + "</div>";' +
                '    ' +
                '    /* Problem Items Summary - moved above table */' +
                '    if (data.problemItems && data.problemItems.length > 0) {' +
                '        html += "<div style=\\"margin-top:20px;padding:15px;background:#fff3e0;border:1px solid #f57c00;border-radius:6px;\\">" ;' +
                '        html += "<strong style=\\"color:#e65100;\\">⚠️ " + data.problemItems.length + " Problem Item(s) Found:</strong>";' +
                '        html += "<ul style=\\"margin:10px 0 0 20px;padding:0;\\">";' +
                '        for (var pi = 0; pi < data.problemItems.length; pi++) {' +
                '            var prob = data.problemItems[pi];' +
                '            var probSoAmt = Math.abs(prob.soAmount || 0);' +
                '            var probInvAmt = Math.abs(prob.invoiceAmount || 0);' +
                '            var probVar = probInvAmt - probSoAmt;' +
                '            html += "<li><strong>" + prob.itemName + "</strong> - " + prob.status + ": SO=" + formatCompCurrency(probSoAmt) + ", INV=" + formatCompCurrency(probInvAmt) + ", Variance=" + formatCompCurrencyWithSign(probVar) + "</li>";' +
                '        }' +
                '        html += "</ul></div>";' +
                '    }' +
                '    ' +
                '    /* Comparison Table */' +
                '    var soTranid = data.salesOrder ? data.salesOrder.tranid : "-";' +
                '    html += "<div class=\\"comparison-table-container\\">";' +
                '    html += "<table class=\\"comparison-table\\">";' +
                '    html += "<thead><tr>";' +
                '    html += "<th>Item</th>";' +
                '    html += "<th>Description</th>";' +
                '    html += "<th>SO #</th>";' +
                '    html += "<th class=\\"amount\\">SO Qty</th>";' +
                '    html += "<th class=\\"amount\\">SO Rate</th>";' +
                '    html += "<th class=\\"amount\\">SO Amount</th>";' +
                '    html += "<th>Invoice(s)</th>";' +
                '    html += "<th class=\\"amount\\">Inv Qty</th>";' +
                '    html += "<th class=\\"amount\\">Inv Rate</th>";' +
                '    html += "<th class=\\"amount\\">Inv Amount</th>";' +
                '    html += "<th class=\\"amount\\">Mismatch</th>";' +
                '    html += "<th class=\\"amount\\">Unbilled</th>";' +
                '    html += "<th>Status</th>";' +
                '    html += "</tr></thead>";' +
                '    html += "<tbody>";' +
                '    ' +
                '    var compTable = data.comparisonTable || [];' +
                '    for (var i = 0; i < compTable.length; i++) {' +
                '        var row = compTable[i];' +
                '        var statusClass = row.status === "Match" ? "match" : (row.status === "NOT INVOICED" ? "not-invoiced" : (row.status === "NOT ON SO" ? "not-on-so" : (row.status === "DISCOUNT" ? "discount" : "mismatch")));' +
                '        ' +
                '        /* Values already have correct sign: */' +
                '        /* Normal items = positive, Discounts = negative */' +
                '        var soAmt = row.soAmount || 0;' +
                '        var invAmt = row.invoiceAmount || 0;' +
                '        var soQty = row.soQty || 0;' +
                '        var invQty = row.invoiceQty || 0;' +
                '        ' +
                '        /* Rates: positive for normal items, negative for discounts */' +
                '        var soRate = row.soRate || 0;' +
                '        var invRate = row.invoiceRate || 0;' +
                '        ' +
                '        /* Use server-calculated variances: mismatch (errors) vs unbilled (pending) */' +
                '        var rowMismatch = row.mismatchVariance || 0;' +
                '        var rowUnbilled = row.unbilledVariance || 0;' +
                '        var mismatchClass = rowMismatch > 0.01 ? "variance-positive" : (rowMismatch < -0.01 ? "variance-negative" : "");' +
                '        var unbilledClass = rowUnbilled > 0.01 ? "variance-positive" : (rowUnbilled < -0.01 ? "variance-negative" : "");' +
                '        ' +
                '        html += "<tr>";' +
                '        html += "<td>" + (row.itemName || "-") + "</td>";' +
                '        html += "<td>" + (row.lineDescription || "-") + "</td>";' +
                '        html += "<td><a href=\\"/app/accounting/transactions/salesord.nl?id=" + data.salesOrder.id + "\\" target=\\"_blank\\" style=\\"color:#1976d2;text-decoration:none;\\">" + soTranid + "</a></td>";' +
                '        html += "<td class=\\"amount\\">" + soQty + "</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencySigned(soRate) + "</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencySigned(soAmt) + "</td>";' +
                '        var invDisplay = (row.invoiceNumbers || "-");' +
                '        if (data.invoices && data.invoices.length > 0 && invDisplay !== "-") {' +
                '            /* Match specific invoices for this line item */' +
                '            var invNumbers = invDisplay.split(",").map(function(s) { return s.trim(); });' +
                '            var invLinks = invNumbers.map(function(invNum) {' +
                '                var matchedInv = null;' +
                '                for (var j = 0; j < data.invoices.length; j++) {' +
                '                    if (data.invoices[j].tranid === invNum) {' +
                '                        matchedInv = data.invoices[j];' +
                '                        break;' +
                '                    }' +
                '                }' +
                '                if (matchedInv) {' +
                '                    return "<a href=\\"/app/accounting/transactions/custinvc.nl?id=" + matchedInv.id + "\\" target=\\"_blank\\" style=\\"color:#1976d2;text-decoration:none;\\">" + invNum + "</a>";' +
                '                }' +
                '                return invNum;' +
                '            }).join(", ");' +
                '            html += "<td>" + invLinks + "</td>";' +
                '        } else {' +
                '            html += "<td>" + invDisplay + "</td>";' +
                '        }' +
                '        html += "<td class=\\"amount\\">" + invQty + "</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencySigned(invRate) + "</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencySigned(invAmt) + "</td>";' +
                '        html += "<td class=\\"amount " + mismatchClass + "\\">" + formatCompCurrencyWithSign(rowMismatch) + "</td>";' +
                '        html += "<td class=\\"amount " + unbilledClass + "\\">" + formatCompCurrencyWithSign(rowUnbilled) + "</td>";' +
                '        html += "<td class=\\"" + statusClass + "\\">" + row.status + "</td>";' +
                '        html += "</tr>";' +
                '    }' +
                '    ' +
                '    /* Totals Row */' +
                '    html += "<tr class=\\"comparison-totals\\">";' +
                '    html += "<td colspan=\\"5\\"><strong>TOTALS</strong></td>";' +
                '    html += "<td class=\\"amount\\">" + formatCompCurrency(soLinePos) + "</td>";' +
                '    html += "<td></td>";' +
                '    html += "<td></td>";' +
                '    html += "<td></td>";' +
                '    html += "<td class=\\"amount\\">" + formatCompCurrency(invLinePos) + "</td>";' +
                '    html += "<td class=\\"amount " + (Math.abs(mismatchVar) > 0.01 ? "variance-negative" : "") + "\\">" + formatCompCurrencyWithSign(mismatchVar) + "</td>";' +
                '    html += "<td class=\\"amount " + (Math.abs(unbilledVar) > 0.01 ? "variance-negative" : "") + "\\">" + formatCompCurrencyWithSign(unbilledVar) + "</td>";' +
                '    html += "<td></td>";' +
                '    html += "</tr>";' +
                '    ' +
                '    html += "</tbody></table></div>";' +
                '    ' +
                '    return html;' +
                '}' +
                '' +
                '/* Render Tab 2: Cross-SO Analysis Result (rename from renderCrossSOResult) */' +
                'function renderTab2Result(data) {' +
                '    var body = document.getElementById("tab-content-2");' +
                '    ' +
                '    if (data.error) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\"><strong>Error:</strong> " + data.error + "</div>";' +
                '        if (data.memo) {' +
                '            body.innerHTML += "<div style=\\"margin-top:10px;padding:10px;background:#f5f5f5;border-radius:4px;\\"><strong>Memo field:</strong> " + data.memo + "</div>";' +
                '        }' +
                '        return;' +
                '    }' +
                '    ' +
                '    var html = "";' +
                '    ' +
                '    /* Helper Text */' +
                '    html += "<div style=\\"margin:0 0 20px;padding:15px;background:#e3f2fd;border-left:4px solid #1976d2;border-radius:4px;\\">";' +
                '    html += "<strong>What This Analysis Shows:</strong> This compares the \\"Created From\\" Source Sales Order from each Customer Deposit to the \\"Created From\\" Source Sales Order from the Invoice it was applied to. ";' +
                '    html += "A <span style=\\"color:#28a745;font-weight:bold;\\">✓ MATCH</span> (green checkmark) confirms the CD was applied to an invoice in the proper SO→Invoice trail. ";' +
                '    html += "A <span style=\\"color:#dc3545;font-weight:bold;\\">✗ CROSS-SO</span> mismatch (red X) indicates the CD and Invoice came from different sales orders.";' +
                '    html += "</div>";' +
                '    ' +
                '    /* Overall Summary Based on Mismatches */' +
                '    var mismatchSummaryClass = data.mismatches === 0 ? "match" : "mismatch";' +
                '    var mismatchSummaryIcon = data.mismatches === 0 ? "✓" : "⚠️";' +
                '    var mismatchSummaryText = data.mismatches === 0 ' +
                '        ? "<strong>" + mismatchSummaryIcon + " No Cross-SO Mismatches Found:</strong> Cross-SO deposit applications are NOT the source of this overpayment. The deposits were correctly applied to invoices from their originating sales orders."' +
                '        : "<strong>" + mismatchSummaryIcon + " Cross-SO Mismatches Detected:</strong> Found " + data.mismatches + " deposit application(s) where the CD and Invoice came from different sales orders. This may be a contributor to the overpayment.";' +
                '    html += "<div class=\\"comparison-conclusion " + mismatchSummaryClass + "\\" style=\\"margin-bottom:20px;\\">" + mismatchSummaryText + "</div>";' +
                '    ' +
                '    /* Summary Cards */' +
                '    html += "<div class=\\"comparison-summary\\">";' +
                '    html += "<div class=\\"comparison-card\\">";' +
                '    html += "<div class=\\"comparison-card-label\\">Total Applications</div>";' +
                '    html += "<div class=\\"comparison-card-value\\">" + data.totalApplications + "</div>";' +
                '    html += "</div>";' +
                '    html += "<div class=\\"comparison-card\\">";' +
                '    html += "<div class=\\"comparison-card-label\\">Same-SO Matches</div>";' +
                '    html += "<div class=\\"comparison-card-value\\" style=\\"color:#28a745;\\">" + data.matches + "</div>";' +
                '    html += "</div>";' +
                '    html += "<div class=\\"comparison-card\\">";' +
                '    html += "<div class=\\"comparison-card-label\\">Cross-SO Mismatches</div>";' +
                '    html += "<div class=\\"comparison-card-value\\" style=\\"color:#dc3545;\\">" + data.mismatches + "</div>";' +
                '    html += "</div>";' +
                '    html += "<div class=\\"comparison-card\\">";' +
                '    html += "<div class=\\"comparison-card-label\\">Non-Invoice Apps</div>";' +
                '    html += "<div class=\\"comparison-card-value\\" style=\\"color:#ff9800;\\">" + data.noInvoice + "</div>";' +
                '    html += "</div>";' +
                '    html += "</div>";' +
                '    ' +
                '    /* Overpayment Source Info */' +
                '    if (data.overpaymentCDTranId) {' +
                '        html += "<div style=\\"margin:15px 0;padding:10px;background:#fff3cd;border:1px solid #ffc107;border-radius:4px;\\"><strong>🎯 Overpayment Source CD:</strong> " + data.overpaymentCDTranId + " (overpayment invoice highlighted below)</div>";' +
                '    }' +
                '    ' +
                '    /* Applications Table */' +
                '    if (data.applications && data.applications.length > 0) {' +
                '        /* Check if we need the Alternate Applications column */' +
                '        var hasAlternateApps = false;' +
                '        for (var j = 0; j < data.applications.length; j++) {' +
                '            if (data.applications[j].appliedTranType && data.applications[j].appliedTranType !== "Invoice") {' +
                '                hasAlternateApps = true;' +
                '                break;' +
                '            }' +
                '        }' +
                '        ' +
                '        html += "<h3 style=\\"margin:20px 0 10px;\\">Deposit Applications</h3>";' +
                '        html += "<table class=\\"comparison-table\\">";' +
                '        html += "<thead><tr>";' +
                '        html += "<th>Status</th>";' +
                '        html += "<th>Customer Deposit</th>";' +
                '        html += "<th>CD Source SO</th>";' +
                '        html += "<th>Invoice</th>";' +
                '        html += "<th>INV Source SO</th>";' +
                '        html += "<th>Amount Applied</th>";' +
                '        if (hasAlternateApps) {' +
                '            html += "<th>Alternate Applications</th>";' +
                '        }' +
                '        html += "</tr></thead><tbody>";' +
                '        ' +
                '        for (var i = 0; i < data.applications.length; i++) {' +
                '            var app = data.applications[i];' +
                '            ' +
                '            /* Determine row class */' +
                '            var rowClass = "";' +
                '            if (app.isOverpaymentInv) {' +
                '                rowClass = "overpayment-row";' +
                '            } else if (app.status === "cross-so") {' +
                '                rowClass = "mismatch-row";' +
                '            } else if (app.status === "no-invoice") {' +
                '                rowClass = "no-invoice-row";' +
                '            } else {' +
                '                rowClass = "match-row";' +
                '            }' +
                '            ' +
                '            /* Determine status badge */' +
                '            var statusBadge = "";' +
                '            if (app.status === "cross-so") {' +
                '                statusBadge = "<span class=\\"status-badge status-error\\">CROSS-SO</span>";' +
                '            } else if (app.status === "no-invoice") {' +
                '                statusBadge = "<span class=\\"status-badge status-warning\\">NO INVOICE</span>";' +
                '            } else {' +
                '                statusBadge = "<span class=\\"status-badge status-success\\">MATCH</span>";' +
                '            }' +
                '            ' +
                '            /* Add overpayment indicator ONLY for overpayment invoice */' +
                '            if (app.isOverpaymentInv) {' +
                '                statusBadge += " <span class=\\"status-badge status-overpayment\\">🎯 OVERPAYMENT</span>";' +
                '            }' +
                '            ' +
                '            /* Add checkmarks/X for SO source matching */' +
                '            var cdSourceSODisplay = app.cdSourceSO || "N/A";' +
                '            var invSourceSODisplay = app.invSourceSO || "N/A";' +
                '            if (app.status === "match" && app.cdSourceSO && app.cdSourceSO !== "N/A" && app.invSourceSO && app.invSourceSO !== "N/A") {' +
                '                cdSourceSODisplay = "<span style=\\"color:#28a745;font-weight:bold;\\">✓</span> " + app.cdSourceSO;' +
                '                invSourceSODisplay = "<span style=\\"color:#28a745;font-weight:bold;\\">✓</span> " + app.invSourceSO;' +
                '            } else if (app.status === "cross-so") {' +
                '                cdSourceSODisplay = "<span style=\\"color:#dc3545;font-weight:bold;\\">✗</span> " + app.cdSourceSO;' +
                '                invSourceSODisplay = "<span style=\\"color:#dc3545;font-weight:bold;\\">✗</span> " + app.invSourceSO;' +
                '            }' +
                '            ' +
                '            /* Build Alternate Applications column (only show non-invoice applications) */' +
                '            var alternateAppCell = "";' +
                '            if (hasAlternateApps) {' +
                '                if (app.appliedTranType && app.appliedTranType !== "Invoice" && app.appliedTranNumber && app.appliedTranId) {' +
                '                    var tranUrl = "";' +
                '                    if (app.appliedTranType === "Refund") {' +
                '                        tranUrl = "/app/accounting/transactions/custref.nl?id=" + app.appliedTranId;' +
                '                    } else if (app.appliedTranType === "Credit Memo") {' +
                '                        tranUrl = "/app/accounting/transactions/custcred.nl?id=" + app.appliedTranId;' +
                '                    } else {' +
                '                        tranUrl = "/app/common/entity/custjob.nl?id=" + app.appliedTranId;' +
                '                    }' +
                '                    alternateAppCell = "<a href=\\"" + tranUrl + "\\" target=\\"_blank\\">" + app.appliedTranType + " " + app.appliedTranNumber + "</a>";' +
                '                } else {' +
                '                    alternateAppCell = "<em style=\\"color:#999;\\">-</em>";' +
                '                }' +
                '            }' +
                '            ' +
                '            html += "<tr class=\\"" + rowClass + "\\">";' +
                '            html += "<td>" + statusBadge + "</td>";' +
                '            html += "<td><a href=\\"/app/accounting/transactions/custdep.nl?id=" + app.cdId + "\\" target=\\"_blank\\">" + app.cdTranId + "</a></td>";' +
                '            html += "<td>" + cdSourceSODisplay + "</td>";' +
                '            html += "<td>" + (app.invTranId ? "<a href=\\"/app/accounting/transactions/custinvc.nl?id=" + app.invId + "\\" target=\\"_blank\\">" + app.invTranId + "</a>" : "<em style=\\"color:#999;\\">N/A</em>") + "</td>";' +
                '            html += "<td>" + invSourceSODisplay + "</td>";' +
                '            html += "<td>" + formatCompCurrency(app.amount) + "</td>";' +
                '            if (hasAlternateApps) {' +
                '                html += "<td>" + alternateAppCell + "</td>";' +
                '            }' +
                '            html += "</tr>";' +
                '        }' +
                '        ' +
                '        html += "</tbody></table>";' +
                '    }' +
                '    ' +
                '    body.innerHTML = html;' +
                '}' +
                '' +
                '/* Render Tab 3: Overpayment Summary Result */' +
                'function renderTab3Result(data) {' +
                '    var body = document.getElementById("tab-content-3");' +
                '    ' +
                '    if (data.error) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\"><strong>Error:</strong> " + data.error + "</div>";' +
                '        return;' +
                '    }' +
                '    ' +
                '    var html = "";' +
                '    ' +
                '    /* Helper Text */' +
                '    html += "<div style=\\"margin:0 0 20px;padding:15px;background:#e3f2fd;border-left:4px solid #1976d2;border-radius:4px;\\">";' +
                '    html += "<strong>What This Analysis Shows:</strong> Overpayments are calculated per Sales Order. ";' +
                '    html += "Each section below shows a Sales Order, its deposit vs invoice calculation, and the Credit Memo(s) created from that overpayment.";' +
                '    html += "</div>";' +
                '    ' +
                '    /* Overall Summary - Match CD Cross-SO tab style with grey background */' +
                '    if (data.salesOrders && data.salesOrders.length > 0) {' +
                '        html += "<div style=\\"display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:25px;\\">";' +
                '        html += "<div style=\\"padding:16px;background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;text-align:center;\\">";' +
                '        html += "<div style=\\"font-size:11px;color:#6c757d;margin-bottom:6px;text-transform:uppercase;font-weight:600;letter-spacing:0.5px;\\">SALES ORDERS WITH OVERPAYMENTS</div>";' +
                '        html += "<div style=\\"font-size:28px;font-weight:700;color:#212529;\\">" + data.salesOrders.length + "</div>";' +
                '        html += "</div>";' +
                '        html += "<div style=\\"padding:16px;background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;text-align:center;\\">";' +
                '        html += "<div style=\\"font-size:11px;color:#6c757d;margin-bottom:6px;text-transform:uppercase;font-weight:600;letter-spacing:0.5px;\\">TOTAL CREDIT MEMOS</div>";' +
                '        html += "<div style=\\"font-size:28px;font-weight:700;color:#212529;\\">" + data.totalCMCount + "</div>";' +
                '        html += "</div>";' +
                '        html += "<div style=\\"padding:16px;background:#f8f9fa;border:1px solid #dee2e6;border-radius:4px;text-align:center;\\">";' +
                '        html += "<div style=\\"font-size:11px;color:#6c757d;margin-bottom:6px;text-transform:uppercase;font-weight:600;letter-spacing:0.5px;\\">COMBINED CM AMOUNT</div>";' +
                '        html += "<div style=\\"font-size:28px;font-weight:700;color:#dc3545;\\">-" + formatCompCurrency(data.totalCMAmount) + "</div>";' +
                '        html += "</div>";' +
                '        html += "</div>";' +
                '        ' +
                '        /* Loop through each Sales Order */' +
                '        for (var i = 0; i < data.salesOrders.length; i++) {' +
                '            var so = data.salesOrders[i];' +
                '            var totalCMsForSO = 0;' +
                '            for (var j = 0; j < so.creditMemos.length; j++) {' +
                '                totalCMsForSO += so.creditMemos[j].cmAmount;' +
                '            }' +
                '            var varianceMatchesCMs = Math.abs(so.overpaymentVariance - totalCMsForSO) < 0.01;' +
                '            ' +
                '            /* SO Card Container - Simple border like other tabs */' +
                '            html += "<div style=\\"margin-bottom:30px;padding:20px;background:#fff;border:1px solid #dee2e6;border-radius:6px;\\">";' +
                '            ' +
                '            /* SO Header */' +
                '            html += "<div style=\\"margin-bottom:20px;padding-bottom:15px;border-bottom:2px solid #e9ecef;\\">";' +
                '            html += "<h3 style=\\"margin:0 0 8px;font-size:18px;color:#212529;\\">Sales Order: <a href=\\"/app/accounting/transactions/salesord.nl?id=" + so.soId + "\\" target=\\"_blank\\" style=\\"color:#007bff;\\">" + so.soNumber + "</a></h3>";' +
                '            html += "<div style=\\"font-size:14px;color:#6c757d;\\">SO Total Value: " + formatCompCurrency(so.soTotal) + "</div>";' +
                '            html += "</div>";' +
                '            ' +
                '            /* Overpayment Calculation */' +
                '            html += "<div style=\\"background:#f8f9fa;padding:20px;border-radius:6px;margin-bottom:20px;\\">";' +
                '            html += "<h4 style=\\"margin:0 0 15px;color:#495057;font-size:14px;font-weight:700;text-transform:uppercase;\\">OVERPAYMENT CALCULATION</h4>";' +
                '            ' +
                '            html += "<table style=\\"width:100%;border-collapse:collapse;\\">";' +
                '            html += "<tr style=\\"border-bottom:1px solid #dee2e6;\\">";' +
                '            html += "<td style=\\"padding:10px 0;font-size:14px;\\">Total Customer Deposits Collected</td>";' +
                '            html += "<td style=\\"padding:10px 0;text-align:right;font-size:16px;font-weight:700;color:#28a745;\\">" + formatCompCurrency(so.totalDeposits) + "</td>";' +
                '            html += "</tr>";' +
                '            html += "<tr style=\\"border-bottom:1px solid #dee2e6;\\">";' +
                '            html += "<td style=\\"padding:10px 0;font-size:14px;\\">Minus Total Invoiced Value</td>";' +
                '            html += "<td style=\\"padding:10px 0;text-align:right;font-size:16px;font-weight:700;color:#dc3545;\\">-" + formatCompCurrency(so.totalInvoiced) + "</td>";' +
                '            html += "</tr>";' +
                '            html += "<tr style=\\"border-top:2px solid #495057;background:#fff;\\">";' +
                '            html += "<td style=\\"padding:12px 0;font-size:15px;font-weight:700;\\">= Overpayment Variance</td>";' +
                '            var varianceIcon = varianceMatchesCMs ? "\u2713" : "\u26a0\ufe0f";' +
                '            var varianceColor = varianceMatchesCMs ? "#28a745" : "#ff6b6b";' +
                '            html += "<td style=\\"padding:12px 0;text-align:right;font-size:18px;font-weight:700;color:" + varianceColor + ";\\">" + varianceIcon + " " + formatCompCurrency(so.overpaymentVariance) + "</td>";' +
                '            html += "</tr>";' +
                '            html += "</table>";' +
                '            html += "</div>";' +
                '            ' +
                '            /* Credit Memos from this Overpayment */' +
                '            html += "<div style=\\"background:#fff3cd;padding:16px;border-left:4px solid #ffc107;border-radius:4px;\\">";' +
                '            html += "<h4 style=\\"margin:0 0 12px;color:#856404;font-size:13px;font-weight:700;text-transform:uppercase;\\">CREDIT MEMOS FROM THIS OVERPAYMENT</h4>";' +
                '            ' +
                '            for (var k = 0; k < so.creditMemos.length; k++) {' +
                '                var cm = so.creditMemos[k];' +
                '                html += "<div style=\\"margin-bottom:8px;padding:12px;background:#fff;border-radius:4px;border:1px solid #ffc107;\\">";' +
                '                html += "<div style=\\"display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;\\">";' +
                '                html += "<div style=\\"font-size:14px;\\"><strong>Credit Memo:</strong> <a href=\\"/app/accounting/transactions/custcred.nl?id=" + cm.cmId + "\\" target=\\"_blank\\" style=\\"color:#dc3545;font-weight:700;\\">" + cm.cmNumber + "</a></div>";' +
                '                html += "<div style=\\"font-size:14px;\\"><strong>Amount:</strong> <span style=\\"color:#dc3545;font-weight:700;font-size:15px;\\">-" + formatCompCurrency(cm.cmAmount) + "</span></div>";' +
                '                html += "<div style=\\"font-size:14px;\\"><strong>Customer Deposit:</strong> <a href=\\"/app/accounting/transactions/custdep.nl?id=" + cm.cdId + "\\" target=\\"_blank\\" style=\\"color:#28a745;font-weight:600;\\">" + cm.cdNumber + "</a></div>";' +
                '                html += "<div style=\\"font-size:14px;\\"><strong>Date:</strong> " + cm.overpaymentDate + "</div>";' +
                '                html += "</div>";' +
                '                html += "</div>";' +
                '            }' +
                '            ' +
                '            /* Variance Check Message */' +
                '            if (varianceMatchesCMs) {' +
                '                html += "<div style=\\"margin-top:12px;padding:10px;background:#d4edda;border:1px solid #c3e6cb;border-radius:4px;color:#155724;text-align:center;font-size:14px;\\"><span style=\\"font-size:16px;\\">✓</span> Overpayment variance matches credit memo total</div>";' +
                '            } else {' +
                '                var difference = Math.abs(so.overpaymentVariance - totalCMsForSO);' +
                '                html += "<div style=\\"margin-top:12px;padding:10px;background:#fff3cd;border:1px solid #ffc107;border-radius:4px;color:#856404;text-align:center;font-size:14px;\\"><strong>⚠️ Note:</strong> CM total (" + formatCompCurrency(totalCMsForSO) + ") differs from variance by " + formatCompCurrency(difference) + "</div>";' +
                '            }' +
                '            ' +
                '            html += "</div>";' +
                '            html += "</div>";' +
                '        }' +
                '    } else {' +
                '        html += "<div style=\\"padding:20px;text-align:center;color:#6c757d;\\">No overpayment credit memos found.</div>";' +
                '    }' +
                '    ' +
                '    body.innerHTML = html;' +
                '}' +
                '' +
                '/* Format currency for comparison modal - always positive with $ */' +
                'function formatCompCurrency(value) {' +
                '    if (value === null || value === undefined) return "$0.00";' +
                '    var num = parseFloat(value) || 0;' +
                '    return "$" + Math.abs(num).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '}' +
                '' +
                '/* Format currency preserving sign - negative shows -$, positive shows $ (no +), zero shows $0.00 */' +
                'function formatCompCurrencySigned(value) {' +
                '    if (value === null || value === undefined) return "$0.00";' +
                '    var num = parseFloat(value) || 0;' +
                '    if (Math.abs(num) < 0.01) return "$0.00";' +
                '    if (num < 0) return "-$" + Math.abs(num).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '    return "$" + num.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '}' +
                '' +
                '/* Format currency with explicit sign for variance display - always shows + or - */' +
                'function formatCompCurrencyWithSign(value) {' +
                '    if (value === null || value === undefined) return "$0.00";' +
                '    var num = parseFloat(value) || 0;' +
                '    if (Math.abs(num) < 0.01) return "$0.00";' +
                '    var prefix = num < 0 ? "-$" : "+$";' +
                '    return prefix + Math.abs(num).toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");' +
                '}' +
                '' +
                '/* =========================================== */' +
                '/* AI TRANSACTION LIFECYCLE ANALYSIS          */' +
                '/* =========================================== */' +
                '' +
                '/* Client-side cache for AI analysis results */' +
                'var aiAnalysisCache = {};' +
                '' +
                '/* Determine if AI button should be hidden (mismatch fully explains CM) */' +
                'function shouldHideAIButton(mismatchVariance, cmTotal) {' +
                '    var mismatchAbs = Math.abs(mismatchVariance || 0);' +
                '    var cmAbs = Math.abs(cmTotal || 0);' +
                '    var difference = Math.abs(mismatchAbs - cmAbs);' +
                '    return difference < 0.10;' +
                '}' +
                '' +
                '/* Load existing AI analysis for a CM */' +
                'function loadExistingAIAnalysis(creditMemoId) {' +
                '    if (aiAnalysisCache[creditMemoId]) {' +
                '        renderAIAnalysisResult(aiAnalysisCache[creditMemoId]);' +
                '        updateAIButtonForExistingAnalysis();' +
                '        return;' +
                '    }' +
                '    ' +
                '    var resultsContent = document.getElementById(\"aiResultsContent\");' +
                '    if (resultsContent) {' +
                '        resultsContent.innerHTML = \"<div style=\\\"padding:15px;color:#777;font-style:italic;\\\"><em>Loading previous AI analysis...</em></div>\";' +
                '    }' +
                '    ' +
                '    var xhr = new XMLHttpRequest();' +
                '    xhr.open(\"POST\", window.location.href, true);' +
                '    xhr.setRequestHeader(\"Content-Type\", \"application/json\");' +
                '    xhr.onload = function() {' +
                '        if (xhr.status === 200) {' +
                '            var response = JSON.parse(xhr.responseText);' +
                '            if (response.success && response.aiAnalysis && response.aiAnalysis.found) {' +
                '                var data = response.aiAnalysis;' +
                '                var displayData = {' +
                '                    haikuResponse: data.haikuResponse,' +
                '                    savedRecordId: data.recordId,' +
                '                    createdDate: data.createdDate' +
                '                };' +
                '                aiAnalysisCache[creditMemoId] = displayData;' +
                '                renderAIAnalysisResult(displayData);' +
                '                updateAIButtonForExistingAnalysis();' +
                '            } else if (resultsContent) {' +
                '                resultsContent.innerHTML = \"\";' +
                '            }' +
                '        }' +
                '    };' +
                '    xhr.send(JSON.stringify({ action: \"loadAIAnalysis\", creditMemoId: creditMemoId }));' +
                '}' +
                '' +
                '/* Update AI button text to show Re-run when analysis exists */' +
                'function updateAIButtonForExistingAnalysis() {' +
                '    var btn = document.getElementById(\"aiAnalysisBtn\");' +
                '    if (btn) {' +
                '        btn.textContent = \"🤖 Re-run AI Analysis\";' +
                '    }' +
                '}' +
                '' +
                '/* Run AI Analysis on Transaction Lifecycle */' +
                'function runAIAnalysis(creditMemoId) {' +
                '    var btn = document.getElementById("aiAnalysisBtn");' +
                '    var resultsContainer = document.getElementById("aiResultsContainer");' +
                '    var resultsContent = document.getElementById("aiResultsContent");' +
                '    ' +
                '    if (!btn || !resultsContainer || !resultsContent) return;' +
                '    ' +
                '    btn.disabled = true;' +
                '    btn.textContent = "⏳ Analyzing Transaction Lifecycle...";' +
                '    btn.style.background = "#999";' +
                '    ' +
                '    resultsContainer.style.display = "block";' +
                '    resultsContent.innerHTML = "<div style=\\"text-align:center;padding:20px;\\"><div style=\\"font-size:18px;\\">⏳</div><div style=\\"margin-top:10px;\\">Analyzing sales order transaction history...</div></div>";' +
                '    ' +
                '    var comparisonData = window.currentComparisonData || {};' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ ' +
                '            action: "aiAnalysis", ' +
                '            creditMemoId: creditMemoId, ' +
                '            comparisonData: comparisonData ' +
                '        })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(result) {' +
                '        if (result.success && result.data) {' +
                '            renderAIAnalysisResult(result.data);' +
                '        } else {' +
                '            resultsContent.innerHTML = "<div style=\\"color:#d32f2f;padding:15px;\\">Error: " + (result.data && result.data.error ? result.data.error : "Unknown error") + "</div>";' +
                '        }' +
                '        btn.disabled = false;' +
                '        btn.textContent = "🤖 Re-run AI Analysis";' +
                '        btn.style.background = "#1976d2";' +
                '    })' +
                '    .catch(function(err) {' +
                '        resultsContent.innerHTML = "<div style=\\"color:#d32f2f;padding:15px;\\">Error: " + err.message + "</div>";' +
                '        btn.disabled = false;' +
                '        btn.textContent = "🤖 Retry AI Analysis";' +
                '        btn.style.background = "#1976d2";' +
                '    });' +
                '}' +
                '' +
                '/* Render AI Analysis Result */' +
                'function renderAIAnalysisResult(data) {' +
                '    /* Legacy function - now redirects to Tab 3 */' +
                '    renderTab4Result(data);' +
                '}' +
                '' +
                '/* Render Tab 4: AI Analysis Result */' +
                'function renderTab4Result(data) {' +
                '    var body = document.getElementById("tab-content-4");' +
                '    if (!body) return;' +
                '    ' +
                '    if (data.error) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\"><strong>Error:</strong> " + data.error + "</div>";' +
                '        return;' +
                '    }' +
                '    ' +
                '    var html = "";' +
                '    ' +
                '    html += "<div style=\\"background:white;padding:20px;border:1px solid #ddd;border-radius:6px;\\">" ;' +
                '    html += "<h3 style=\\"margin:0 0 15px 0;font-size:16px;color:#7c4dff;\\">\ud83e\udd16 AI Generated SO Price Changes</h3>";' +
                '    ' +
                '    var responseText = data.haikuResponse || "No response received";' +
                '    ' +
                '    html += "<div style=\\"white-space:pre-wrap;font-family:Arial,sans-serif;font-size:13px;line-height:1.6;color:#333;\\">" + escapeHtmlClient(responseText) + "</div>";' +
                '    ' +
                '    html += "<div style=\\"margin-top:15px;padding-top:15px;border-top:1px solid #ddd;font-size:12px;color:#777;\\">" ;' +
                '    if (data.systemNotesCount !== undefined) {' +
                '        html += "System Notes Analyzed: " + data.systemNotesCount + " | ";' +
                '        html += "Financial: " + (data.organizedNotes && data.organizedNotes.financial ? data.organizedNotes.financial.length : 0) + " | ";' +
                '        html += "Lifecycle: " + (data.organizedNotes && data.organizedNotes.lifecycle ? data.organizedNotes.lifecycle.length : 0);' +
                '    }' +
                '    if (data.savedRecordId) {' +
                '        html += (data.systemNotesCount !== undefined ? " | " : "") + "<span style=\\"color:#7c4dff;\\">Record ID: " + data.savedRecordId + "</span>";' +
                '    }' +
                '    if (data.createdDate) {' +
                '        html += " | <span style=\\"color:#777;\\">Created: " + data.createdDate + "</span>";' +
                '    }' +
                '    html += "</div>";' +
                '    html += "</div>";' +
                '    ' +
                '    body.innerHTML = html;' +
                '}' +
                '' +
                '/* Render Tab 4 with Re-Run Button */' +
                'function renderTab4ResultWithRerun(data) {' +
                '    var body = document.getElementById("tab-content-4");' +
                '    if (!body) return;' +
                '    ' +
                '    if (data.error) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\"><strong>Error:</strong> " + data.error + "</div>";' +
                '        return;' +
                '    }' +
                '    ' +
                '    var html = "";' +
                '    ' +
                '    /* Re-Run AI Generation Button at top */' +
                '    html += "<div style=\\"margin-bottom:15px;\\\">";' +
                '    html += "<button type=\\"button\\" id=\\"rerunAIBtn\\" onclick=\\"rerunAIGeneration()\\" style=\\"padding:10px 20px;font-size:14px;font-weight:700;color:#fff;background:linear-gradient(135deg,#7c4dff,#651fff);border:none;border-radius:6px;cursor:pointer;transition:all 0.2s;box-shadow:0 2px 4px rgba(124,77,255,0.3);\\">";' +
                '    html += "🤖 Re-Run AI Generation</button>";' +
                '    if (data.createdDate) {' +
                '        html += "<span style=\\"margin-left:15px;font-size:12px;color:#777;\\">Last generated: " + data.createdDate + "</span>";' +
                '    }' +
                '    html += "</div>";' +
                '    ' +
                '    html += "<div style=\\"background:white;padding:20px;border:1px solid #ddd;border-radius:6px;\\">" ;' +
                '    html += "<h3 style=\\"margin:0 0 15px 0;font-size:16px;color:#7c4dff;\\">\ud83e\udd16 AI Generated SO Price Changes</h3>";' +
                '    ' +
                '    var responseText = data.haikuResponse || "No response received";' +
                '    ' +
                '    html += "<div style=\\"white-space:pre-wrap;font-family:Arial,sans-serif;font-size:13px;line-height:1.6;color:#333;\\">" + escapeHtmlClient(responseText) + "</div>";' +
                '    ' +
                '    html += "<div style=\\"margin-top:15px;padding-top:15px;border-top:1px solid #ddd;font-size:12px;color:#777;\\">" ;' +
                '    if (data.systemNotesCount !== undefined) {' +
                '        html += "System Notes Analyzed: " + data.systemNotesCount + " | ";' +
                '        html += "Financial: " + (data.organizedNotes && data.organizedNotes.financial ? data.organizedNotes.financial.length : 0) + " | ";' +
                '        html += "Lifecycle: " + (data.organizedNotes && data.organizedNotes.lifecycle ? data.organizedNotes.lifecycle.length : 0);' +
                '    }' +
                '    if (data.savedRecordId) {' +
                '        html += (data.systemNotesCount !== undefined ? " | " : "") + "<span style=\\"color:#7c4dff;\\">Record ID: " + data.savedRecordId + "</span>";' +
                '    }' +
                '    html += "</div>";' +
                '    html += "</div>";' +
                '    ' +
                '    body.innerHTML = html;' +
                '}' +
                '' +
                '/* Re-Run AI Generation */' +
                'function rerunAIGeneration() {' +
                '    var btn = document.getElementById("rerunAIBtn");' +
                '    if (btn) {' +
                '        btn.disabled = true;' +
                '        btn.innerHTML = "⏳ Running AI Analysis...";' +
                '        btn.style.background = "#999";' +
                '    }' +
                '    runTab4AIAnalysis();' +
                '}' +
                '' +
                '/* Client-side HTML escape */' +
                'function escapeHtmlClient(text) {' +
                '    if (!text) return "";' +
                '    return text' +
                '        .replace(/&/g, "&amp;")' +
                '        .replace(/</g, "&lt;")' +
                '        .replace(/>/g, "&gt;")' +
                '        .replace(/"/g, "&quot;")' +
                '        .replace(/\'/g, "&#039;");' +
                '}';
        }

        // ============================================================================
        // AI TRANSACTION LIFECYCLE ANALYSIS FUNCTIONS
        // ============================================================================

        /**
         * Query current line items on Sales Order for context
         * @param {number} soInternalId - Sales Order internal ID
         * @returns {Array} Current line items with item names and amounts
         */
        function querySalesOrderLineItems(soInternalId) {
            try {
                // Note: NetSuite stores SO line quantities/amounts as negative in transactionLine table
                // Use ABS() to normalize for display and AI analysis
                // Use netamount instead of amount (amount is not exposed in SuiteQL)
                var sql = `
                    SELECT 
                        tl.id,
                        tl.linesequencenumber,
                        i.itemid,
                        i.displayname,
                        ABS(tl.quantity) AS quantity,
                        tl.rate,
                        ABS(tl.netamount) AS amount
                    FROM transactionLine tl
                    LEFT JOIN item i ON tl.item = i.id
                    WHERE tl.transaction = ` + soInternalId + `
                    AND tl.mainline = 'F'
                    ORDER BY tl.linesequencenumber
                `;

                var results = query.runSuiteQL({ query: sql }).asMappedResults();
                log.debug('SO Line Items Query', {
                    soId: soInternalId,
                    lineCount: results.length
                });
                return results;

            } catch (e) {
                log.error('Error querying SO line items', { error: e.message, soId: soInternalId });
                return [];
            }
        }

        /**
         * Query System Notes for a Sales Order to track transaction changes
         * @param {number} soInternalId - Sales Order internal ID
         * @param {string} soCreationDate - Sales Order creation date (to capture initial "Set" values)
         * @param {string} overpaymentDate - Overpayment recognition date (YYYY-MM-DD)
         * @returns {Array} System Notes records
         */
        function querySystemNotesForSO(soInternalId, soCreationDate, overpaymentDate) {
            try {
                // Convert dates from M/D/YYYY to YYYY-MM-DD for SQL
                function formatDateForSQL(dateStr) {
                    if (!dateStr) return null;
                    var parts = dateStr.split('/');
                    if (parts.length === 3) {
                        var month = parts[0].padStart(2, '0');
                        var day = parts[1].padStart(2, '0');
                        var year = parts[2];
                        return year + '-' + month + '-' + day;
                    }
                    return null;
                }

                var soDateSQL = formatDateForSQL(soCreationDate);
                var overpaymentDateSQL = formatDateForSQL(overpaymentDate);

                if (!soDateSQL || !overpaymentDateSQL) {
                    log.error('Invalid Dates for SQL', {
                        soCreationDate: soCreationDate,
                        overpaymentDate: overpaymentDate
                    });
                    return [];
                }

                log.debug('System Notes Query Params', {
                    soInternalId: soInternalId,
                    soCreationDate: soDateSQL,
                    overpaymentDate: overpaymentDateSQL
                });

                // OPTIMIZED: Query only fields we care about with date filtering in SQL
                // Focus on: Amount changes (MAMOUNT, RUNITPRICE, RQTY), Header total (MAMOUNTMAIN)
                // Include WHO (employee), WHEN (date), WHAT (item via lineid + item name + description)
                // Added RQTYSHIPRECV to track line-level fulfillment dates
                // Added sn.type to distinguish "Set" (initial value) vs "Change" (modification)
                var sql = `
                    SELECT 
                        sn.id,
                        sn.date,
                        sn.field,
                        sn.type,
                        sn.oldvalue,
                        sn.newvalue,
                        sn.name AS employee_id,
                        sn.lineid,
                        e.firstname || ' ' || e.lastname AS user_name,
                        i.itemid,
                        i.displayname AS item_name,
                        i.description AS item_description
                    FROM systemNote sn
                    LEFT JOIN employee e ON sn.name = e.id
                    LEFT JOIN transactionLine tl ON sn.recordid = tl.transaction AND sn.lineid = tl.id
                    LEFT JOIN item i ON tl.item = i.id
                    WHERE sn.recordid = ` + soInternalId + `
                        AND sn.field IN (
                            'TRANLINE.MAMOUNT',
                            'TRANLINE.RUNITPRICE', 
                            'TRANLINE.RQTY',
                            'TRANLINE.RQTYSHIPRECV',
                            'TRANDOC.MAMOUNTMAIN',
                            'TRANDOC.KSTATUS'
                        )
                        AND sn.date BETWEEN TO_DATE('` + soDateSQL + `', 'YYYY-MM-DD') 
                                        AND TO_DATE('` + overpaymentDateSQL + `', 'YYYY-MM-DD')
                    ORDER BY sn.date, sn.id
                `;

                var results = query.runSuiteQL({ query: sql }).asMappedResults();
                
                log.debug('System Notes Query - Optimized Results', {
                    soId: soInternalId,
                    totalFound: results.length,
                    dateRange: soDateSQL + ' to ' + overpaymentDateSQL,
                    sampleFields: results.length > 0 ? results.slice(0, 5).map(function(r) { 
                        return { date: r.date, field: r.field, user: r.user_name }; 
                    }) : []
                });
                
                return results;

            } catch (e) {
                log.error('Error querying system notes', { error: e.message, soId: soInternalId });
                return [];
            }
        }

        /**
         * Organize and filter System Notes by tier (financial, lifecycle)
         * OPTIMIZED: Since query now only returns relevant fields, categorization is simpler
         * @param {Array} rawSystemNotes - Raw system notes from optimized query
         * @returns {Object} Organized system notes by tier
         */
        function organizeSystemNotesByTier(rawSystemNotes) {
            var organized = {
                financial: [],
                lifecycle: []
            };

            // Temporary array for lifecycle notes before consolidation
            var rawLifecycleNotes = [];

            // Categorize the pre-filtered system notes
            for (var i = 0; i < rawSystemNotes.length; i++) {
                var note = rawSystemNotes[i];
                var field = note.field;

                // Financial fields (includes fulfillment for timing context)
                if (field === 'TRANLINE.MAMOUNT' || field === 'TRANLINE.RUNITPRICE' || 
                    field === 'TRANLINE.RQTY' || field === 'TRANLINE.RQTYSHIPRECV' ||
                    field === 'TRANDOC.MAMOUNTMAIN') {
                    organized.financial.push(note);
                }
                // Lifecycle fields - collect for consolidation
                else if (field === 'TRANDOC.KSTATUS') {
                    rawLifecycleNotes.push(note);
                }
            }

            // Consolidate lifecycle notes: combine sequential status changes on same date by same employee
            // Example: Pending Fulfillment → Pending Billing → Partially Fulfilled (all on same date)
            // becomes one entry showing the full chain
            organized.lifecycle = consolidateStatusChanges(rawLifecycleNotes);

            log.debug('System Notes Organized', 'Financial: ' + organized.financial.length + 
                      ', Lifecycle: ' + organized.lifecycle.length + ' (from ' + rawLifecycleNotes.length + ' raw)' +
                      ' (Total: ' + rawSystemNotes.length + ')');

            return organized;
        }

        /**
         * Consolidate sequential status changes on the same date by the same employee
         * into a single entry showing the full status progression chain
         * @param {Array} statusNotes - Raw status change notes
         * @returns {Array} Consolidated status notes with chain property
         */
        function consolidateStatusChanges(statusNotes) {
            if (!statusNotes || statusNotes.length === 0) return [];

            // Group by date + employee
            var groups = {};
            for (var i = 0; i < statusNotes.length; i++) {
                var note = statusNotes[i];
                var key = note.date + '|' + (note.user_name || note.employee_id || 'Unknown');
                
                if (!groups[key]) {
                    groups[key] = [];
                }
                groups[key].push(note);
            }

            var consolidated = [];

            // Process each group
            for (var groupKey in groups) {
                var groupNotes = groups[groupKey];
                
                if (groupNotes.length === 1) {
                    // Single note, no consolidation needed
                    consolidated.push(groupNotes[0]);
                } else {
                    // Multiple notes on same date by same employee - build chain
                    // Sort by id to maintain chronological order within the day
                    groupNotes.sort(function(a, b) {
                        return parseInt(a.id, 10) - parseInt(b.id, 10);
                    });

                    // Build the status chain: first oldvalue → intermediate newvalues → final newvalue
                    var chain = [];
                    chain.push(groupNotes[0].oldvalue || 'Unknown');
                    
                    for (var j = 0; j < groupNotes.length; j++) {
                        chain.push(groupNotes[j].newvalue);
                    }

                    // Create consolidated note using first note as base
                    var consolidatedNote = {
                        id: groupNotes[0].id,
                        date: groupNotes[0].date,
                        field: groupNotes[0].field,
                        type: groupNotes[0].type,
                        oldvalue: groupNotes[0].oldvalue,
                        newvalue: groupNotes[groupNotes.length - 1].newvalue,
                        employee_id: groupNotes[0].employee_id,
                        user_name: groupNotes[0].user_name,
                        // Custom property: full status chain for display
                        statusChain: chain,
                        consolidatedCount: groupNotes.length
                    };

                    consolidated.push(consolidatedNote);
                }
            }

            // Sort consolidated results by date, then by id
            consolidated.sort(function(a, b) {
                if (a.date !== b.date) {
                    return new Date(a.date) - new Date(b.date);
                }
                return parseInt(a.id, 10) - parseInt(b.id, 10);
            });

            return consolidated;
        }

        /**
         * Build system prompt for Claude AI forensic analyst
         * @returns {string} System prompt
         */
        function buildTransactionLifecycleSystemPrompt() {
            return 'You are a Sales Order change analyst. Analyze and narrate order modifications clearly based on observed facts.\n\n' +
                'DATA YOU WILL RECEIVE:\n' +
                '- TRANLINE.MAMOUNT: Line item dollar amount\n' +
                '- TRANLINE.RUNITPRICE: Line item unit price\n' +
                '- TRANLINE.RQTY: Line item quantity\n' +
                '- TRANLINE.RQTYSHIPRECV: Line fulfillment (0→1 = shipped)\n' +
                '- TRANDOC.MAMOUNTMAIN: Header-level order total\n' +
                '- TRANDOC.KSTATUS: Order status changes\n' +
                '- Each record includes: type, date, user_name, itemid, item_name, item_description\n\n' +
                'TYPE FIELD VALUES:\n' +
                '- SET = INITIAL VALUE when order/line was created (no oldvalue)\n' +
                '- EDIT = intermediate value entry\n' +
                '- CHANGE = documented modification from old to new value\n\n' +
                'CRITICAL PRINCIPLES:\n\n' +
                '1. REPORT OBSERVED FACTS ONLY\n' +
                '   - State what the audit trail shows\n' +
                '   - Do NOT infer causes (change orders, phase additions, deletions)\n' +
                '   - Do NOT say "probably," "likely," or "may indicate"\n\n' +
                '2. FLAG DISCREPANCIES, DON\'T EXPLAIN THEM\n' +
                '   - Point out mismatches (fulfillment count vs value changes)\n' +
                '   - Report timing gaps (fulfilled before priced)\n' +
                '   - Do NOT propose explanations for why\n\n' +
                '3. POST-FULFILLMENT CHANGES (CRITICAL)\n' +
                '   - If item was fulfilled THEN value changed = ⚠️ FLAG IMMEDIATELY\n' +
                '   - Invoice was created at fulfillment; current SO may differ\n' +
                '   - This MUST be in SUMMARY, not buried\n\n' +
                'FORMATTING RULES:\n' +
                '- ALWAYS include item_description with itemid: "Item 101954723-03 (Omega Renner Maple Beach house)"\n' +
                '- ALWAYS use user_name, never employee_id\n' +
                '- CONSOLIDATE: When MAMOUNT and RUNITPRICE change same amount on same date, report ONCE\n' +
                '- STATUS CHANGES must include DATE and WHO\n\n' +
                'OUTPUT FORMAT:\n\n' +
                'SUMMARY:\n' +
                '2-3 sentences. If post-fulfillment changes detected, lead with:\n' +
                '"⚠️ POST-FULFILLMENT VALUE CHANGE: [itemid] ([description]) was fulfilled [date] but value changed [later date]. Invoice may not match current SO."\n' +
                'If none: "✓ No post-fulfillment value changes detected. [Brief order summary]."\n\n' +
                'ORDER TOTAL CHANGES:\n' +
                '• [date] ([employee]): Initial $X (SET)\n' +
                '• [date] ([employee]): $X → $Y\n' +
                '• Current Total: $Z\n\n' +
                'LINE ITEM CHANGES:\n' +
                'Only items with value changes:\n' +
                '• [itemid] ([description]):\n' +
                '  - Fulfilled: [date] by [employee]\n' +
                '  - Value: $X → $Y on [date] by [employee]\n' +
                '  - ⚠️ [FLAG if fulfilled before value change]\n\n' +
                'STATUS CHANGES:\n' +
                '• [date] ([employee]): [old] → [new]\n\n' +
                'OBSERVATIONS:\n' +
                'Report discrepancies without interpretation:\n' +
                '• [X] items fulfilled, [Y] items with documented value changes\n' +
                '• Any timing gaps between fulfillment and pricing\n' +
                '• Order total changes without corresponding line changes\n\n' +
                'IMPORTANT: Do NOT include any decision or recommendation. Report only the facts and observations.';
        }

        /**
         * Build user prompt with transaction details and system notes
         * @param {Object} context - Analysis context with CM, CD, SO, system notes, comparison data
         * @returns {string} User prompt
         */
        function buildTransactionLifecycleUserPrompt(context) {
            var cm = context.creditMemo;
            var cd = context.customerDeposit;
            var so = context.salesOrder;
            var organizedNotes = context.organizedNotes;

            var prompt = 'SALES ORDER: ' + so.tranid + '\n';
            prompt += 'Order Date: ' + so.trandate + '\n';
            prompt += 'Current Total: $' + so.amount.toFixed(2) + '\n';
            prompt += 'Status: ' + so.statusText + '\n\n';

            // Format system notes in readable text format (not JSON)
            prompt += 'SYSTEM NOTES:\n';
            prompt += 'Below are the meaningful changes to this order. Each note includes:\n';
            prompt += '- type: "Set" = INITIAL VALUE when order created, "Change" = MODIFICATION from old to new\n';
            prompt += '- item_name/itemid: The specific item (for line-level changes)\n';
            prompt += '- user_name: The employee who made the change\n';
            prompt += '- date, field, oldvalue, newvalue\n';
            prompt += 'IMPORTANT: "Set" entries show the STARTING values. No oldvalue means this was the original value.\n\n';

            // Helper function to map NetSuite type IDs to text
            // Type 1 = Set (initial value), Type 4 = Change, Type 2 = Edit/intermediate
            function mapTypeToText(typeId) {
                var typeNum = parseInt(typeId, 10);
                if (typeNum === 1) return 'SET';
                if (typeNum === 4) return 'CHANGE';
                if (typeNum === 2) return 'EDIT';
                return 'CHANGE'; // Default to CHANGE for unknown types
            }

            // Separate and format by type
            var financialNotes = organizedNotes.financial || [];
            var lifecycleNotes = organizedNotes.lifecycle || [];

            // Group financial notes by type for clarity
            var headerChanges = [];
            var lineAmountChanges = [];
            var fulfillmentChanges = [];

            for (var i = 0; i < financialNotes.length; i++) {
                var note = financialNotes[i];
                if (note.field === 'TRANDOC.MAMOUNTMAIN') {
                    headerChanges.push(note);
                } else if (note.field === 'TRANLINE.RQTYSHIPRECV') {
                    fulfillmentChanges.push(note);
                } else {
                    lineAmountChanges.push(note);
                }
            }

            // Header total changes
            if (headerChanges.length > 0) {
                prompt += '=== ORDER TOTAL CHANGES ===\n';
                for (var h = 0; h < headerChanges.length; h++) {
                    var hdr = headerChanges[h];
                    var changeType = mapTypeToText(hdr.type);
                    prompt += hdr.date + ' | ' + changeType + ' | ' + (hdr.user_name || 'Unknown') + ' | ';
                    if (changeType === 'SET') {
                        prompt += 'INITIAL VALUE: $' + hdr.newvalue + '\n';
                    } else {
                        prompt += '$' + (hdr.oldvalue || '0') + ' → $' + hdr.newvalue + '\n';
                    }
                }
                prompt += '\n';
            }

            // Line fulfillments (important for timing context)
            if (fulfillmentChanges.length > 0) {
                prompt += '=== LINE FULFILLMENTS ===\n';
                prompt += '(RQTYSHIPRECV 0→1 means item was shipped/fulfilled)\n';
                for (var f = 0; f < fulfillmentChanges.length; f++) {
                    var ful = fulfillmentChanges[f];
                    var itemRef = ful.item_name || ful.itemid || 'Unknown';
                    var itemDesc = ful.item_description ? ' (' + ful.item_description + ')' : '';
                    var changeType = mapTypeToText(ful.type);
                    prompt += ful.date + ' | ' + changeType + ' | ' + (ful.user_name || 'Unknown') + ' | ';
                    prompt += 'Item: ' + itemRef + itemDesc + ' | ';
                    prompt += 'Qty Shipped: ' + (ful.oldvalue || '0') + ' → ' + ful.newvalue + '\n';
                }
                prompt += '\n';
            }

            // Line amount/price/qty changes (most critical)
            if (lineAmountChanges.length > 0) {
                prompt += '=== LINE ITEM VALUE CHANGES ===\n';
                prompt += '(Amount, Price, or Quantity changes on specific items)\n';
                prompt += '(SET = initial value when line added, CHANGE = modification)\n';
                for (var l = 0; l < lineAmountChanges.length; l++) {
                    var line = lineAmountChanges[l];
                    var itemName = line.item_name || line.itemid || 'Unknown';
                    var itemDesc = line.item_description ? ' (' + line.item_description + ')' : '';
                    var fieldType = line.field.replace('TRANLINE.', '');
                    var changeType = mapTypeToText(line.type);
                    prompt += line.date + ' | ' + changeType + ' | ' + (line.user_name || 'Unknown') + ' | ';
                    prompt += 'Item: ' + itemName + itemDesc + ' | ';
                    if (changeType === 'SET') {
                        prompt += fieldType + ' INITIAL: $' + line.newvalue + '\n';
                    } else {
                        prompt += fieldType + ': $' + (line.oldvalue || '0') + ' → $' + line.newvalue + '\n';
                    }
                }
                prompt += '\n';
            }

            // Status changes (consolidated - multiple changes on same date shown as chain)
            if (lifecycleNotes.length > 0) {
                prompt += '=== STATUS CHANGES ===\n';
                prompt += '(Status transitions on same date by same employee shown as full chain)\n';
                for (var s = 0; s < lifecycleNotes.length; s++) {
                    var stat = lifecycleNotes[s];
                    var changeType = mapTypeToText(stat.type);
                    prompt += stat.date + ' | ' + changeType + ' | ' + (stat.user_name || 'Unknown') + ' | ';
                    
                    // Check if this is a consolidated entry with a status chain
                    if (stat.statusChain && stat.statusChain.length > 2) {
                        // Full chain: Pending Fulfillment → Pending Billing → Partially Fulfilled
                        prompt += stat.statusChain.join(' → ') + '\n';
                    } else if (changeType === 'SET') {
                        prompt += 'INITIAL STATUS: ' + stat.newvalue + '\n';
                    } else {
                        prompt += (stat.oldvalue || 'Unknown') + ' → ' + stat.newvalue + '\n';
                    }
                }
                prompt += '\n';
            }

            prompt += 'TOTAL NOTES: ' + (financialNotes.length + lifecycleNotes.length) + '\n\n';
            prompt += 'Please analyze these changes and provide your summary following the output format.';

            return prompt;
        }

        // parseAIDecision function removed - we now provide narrative only, no decision

        /**
         * Save AI analysis to custom record
         * @param {Object} params - Save parameters
         * @returns {number} Custom record internal ID
         */
        function saveAIAnalysisToRecord(params) {
            try {
                var aiRecord = record.create({
                    type: 'customrecord_ai_analysis_stored_response',
                    isDynamic: false
                });

                aiRecord.setValue({
                    fieldId: 'custrecord_linked_transaction',
                    value: params.creditMemoId
                });

                aiRecord.setValue({
                    fieldId: 'custrecord_haiku_response',
                    value: params.haikuResponse ? params.haikuResponse.substring(0, 9999) : '' // Field limit
                });

                aiRecord.setValue({
                    fieldId: 'custrecord_user_prompt',
                    value: params.userPrompt ? params.userPrompt.substring(0, 9999) : ''
                });

                aiRecord.setValue({
                    fieldId: 'custrecord_system_prompt',
                    value: params.systemPrompt ? params.systemPrompt.substring(0, 9999) : ''
                });

                // aiDecision field no longer saved - we provide narrative only

                var recordId = aiRecord.save();
                log.debug('AI Analysis Saved', 'Custom record ID: ' + recordId);
                return recordId;

            } catch (e) {
                log.error('Error saving AI analysis', { error: e.message, stack: e.stack });
                return null;
            }
        }

        /**
         * Load saved AI analysis from custom record
         * @param {number} creditMemoId - Credit Memo internal ID
         * @returns {Object} Saved AI analysis or {found: false}
         */
        function loadSavedAIAnalysis(creditMemoId) {
            try {
                var aiSearch = search.create({
                    type: 'customrecord_ai_analysis_stored_response',
                    filters: [
                        ['custrecord_linked_transaction', 'anyof', creditMemoId]
                    ],
                    columns: [
                        'internalid',
                        'custrecord_haiku_response',
                        'custrecord_user_prompt',
                        'custrecord_system_prompt',
                        search.createColumn({ name: 'created', sort: search.Sort.DESC })
                    ]
                });

                // Get the most recent record (sorted by created date DESC)
                var results = aiSearch.run().getRange({ start: 0, end: 1 });

                if (results && results.length > 0) {
                    var result = results[0];
                    log.debug('AI Analysis Found', 'Record ID: ' + result.id);
                    
                    return {
                        found: true,
                        recordId: result.id,
                        haikuResponse: result.getValue('custrecord_haiku_response'),
                        userPrompt: result.getValue('custrecord_user_prompt'),
                        systemPrompt: result.getValue('custrecord_system_prompt'),
                        createdDate: result.getValue('created')
                    };
                } else {
                    log.debug('No AI Analysis Found', 'CM ID: ' + creditMemoId);
                    return { found: false };
                }

            } catch (e) {
                log.error('Error loading AI analysis', { error: e.message, stack: e.stack, creditMemoId: creditMemoId });
                return { found: false, error: e.message };
            }
        }

        /**
         * Get AI analysis status for multiple credit memos (for icon display)
         * @param {Array} cmIds - Array of Credit Memo internal IDs
         * @returns {Object} Lookup object { cmId: true, ... } for CMs with AI analysis
         */
        function getAIAnalysisLookup(cmIds) {
            try {
                if (!cmIds || cmIds.length === 0) {
                    return {};
                }

                var aiSearch = search.create({
                    type: 'customrecord_ai_analysis_stored_response',
                    filters: [
                        ['custrecord_linked_transaction', 'anyof', cmIds]
                    ],
                    columns: ['custrecord_linked_transaction']
                });

                var results = aiSearch.run().getRange({ start: 0, end: 1000 });
                var lookup = {};

                for (var i = 0; i < results.length; i++) {
                    var cmId = results[i].getValue('custrecord_linked_transaction');
                    if (cmId) {
                        lookup[cmId] = true;
                    }
                }

                log.debug('AI Analysis Lookup', 'Found ' + Object.keys(lookup).length + ' CMs with AI analysis out of ' + cmIds.length + ' total CMs');
                return lookup;

            } catch (e) {
                log.error('Error getting AI analysis lookup', { error: e.message, stack: e.stack });
                return {};
            }
        }

        /**
         * Main function: Analyze Transaction Lifecycle with Claude AI
         * @param {number} creditMemoId - Credit Memo internal ID
         * @param {Object} comparisonData - SO/INV comparison results (mismatch variance, etc.)
         * @returns {Object} AI analysis result
         */
        function analyzeTransactionLifecycleWithAI(creditMemoId, comparisonData) {
            try {
                log.debug('AI Analysis Start', 'CM ID: ' + creditMemoId);

                // Get Credit Memo details
                var cmData = getCreditMemoDetails(creditMemoId);
                if (!cmData) {
                    return { error: 'Credit Memo not found', creditMemoId: creditMemoId };
                }

                // Get Customer Deposit details
                var cdTranid = extractCDFromMemo(cmData.memo);
                if (!cdTranid) {
                    return { error: 'Could not extract Customer Deposit from memo', creditMemo: cmData };
                }

                var cdData = getCustomerDepositDetails(cdTranid);
                if (!cdData || !cdData.soId) {
                    return { error: 'Customer Deposit or Sales Order not found', creditMemo: cmData };
                }

                // Get Sales Order details
                var soData = getSalesOrderDetails(cdData.soId);
                if (!soData) {
                    return { error: 'Sales Order not found', creditMemo: cmData };
                }

                // Query System Notes for SO - use SO creation date to capture initial "Set" values
                var rawSystemNotes = querySystemNotesForSO(cdData.soId, soData.trandate, cmData.trandate);
                
                // Query current SO line items for context
                var currentLineItems = querySalesOrderLineItems(cdData.soId);
                
                // Organize and filter system notes
                var organizedNotes = organizeSystemNotesByTier(rawSystemNotes);

                // Build prompts
                var systemPrompt = buildTransactionLifecycleSystemPrompt();
                var userPrompt = buildTransactionLifecycleUserPrompt({
                    creditMemo: cmData,
                    customerDeposit: cdData,
                    salesOrder: soData,
                    organizedNotes: organizedNotes,
                    currentLineItems: currentLineItems,
                    comparisonData: comparisonData
                });

                // Get API key from script parameters directly
                var currentScript = runtime.getCurrentScript();
                var apiKey = currentScript.getParameter({ name: 'custscript_claude_api_key_kitchenoverpay' });
                if (!apiKey) {
                    return { error: 'Claude API key not found in script parameters. Please set custscript_claude_api_key_kitchenoverpay.' };
                }

                // Call Claude API (Haiku model for cost efficiency)
                log.debug('Calling Claude API', 'Model: haiku, System Prompt cached: yes');
                var claudeResponse = claudeAPI.callClaude({
                    apiKey: apiKey,
                    systemPrompt: systemPrompt,
                    userPrompt: userPrompt,
                    model: 'haiku',
                    maxTokens: 2000,
                    cacheSystemPrompt: true
                });

                if (!claudeResponse.success) {
                    return { error: 'Claude API error: ' + claudeResponse.error };
                }

                var haikuResponse = claudeResponse.analysis;

                log.debug('AI Analysis Complete', 'Response length: ' + (haikuResponse ? haikuResponse.length : 0));

                // Save to custom record (no decision parsing - narrative only)
                var savedRecordId = saveAIAnalysisToRecord({
                    creditMemoId: creditMemoId,
                    haikuResponse: haikuResponse,
                    userPrompt: userPrompt,
                    systemPrompt: systemPrompt
                });

                return {
                    success: true,
                    haikuResponse: haikuResponse,
                    systemNotesCount: rawSystemNotes.length,
                    organizedNotes: organizedNotes,
                    savedRecordId: savedRecordId,
                    creditMemo: cmData,
                    customerDeposit: cdData,
                    salesOrder: soData
                };

            } catch (e) {
                log.error('Error in AI Analysis', { error: e.message, stack: e.stack, creditMemoId: creditMemoId });
                return { error: e.message, creditMemoId: creditMemoId };
            }
        }

        // ============================================================================
        // SO TO INVOICE COMPARISON FUNCTIONS
        // ============================================================================

        /**
         * Analyzes a Credit Memo to compare SO vs Invoice line items
         * @param {number} creditMemoId - Internal ID of the Credit Memo
         * @returns {Object} Complete analysis result
         */
        function analyzeCreditMemoOverpayment(creditMemoId) {
            try {
                log.debug('Analyzing CM', 'Credit Memo ID: ' + creditMemoId);

                // STEP 1: Get Credit Memo details
                var cmData = getCreditMemoDetails(creditMemoId);
                if (!cmData) {
                    return { error: 'Credit Memo not found', creditMemoId: creditMemoId };
                }
                log.debug('CM Data', JSON.stringify(cmData));

                // STEP 2: Parse CD from memo
                var cdTranid = extractCDFromMemo(cmData.memo);
                if (!cdTranid) {
                    return { 
                        error: 'Could not extract Customer Deposit reference from Credit Memo memo field', 
                        creditMemo: cmData,
                        memo: cmData.memo
                    };
                }
                log.debug('Extracted CD', cdTranid);

                // STEP 3: Get Customer Deposit & trace to SO
                var cdData = getCustomerDepositDetails(cdTranid);
                if (!cdData) {
                    return { 
                        error: 'Customer Deposit not found: ' + cdTranid, 
                        creditMemo: cmData 
                    };
                }
                log.debug('CD Data', JSON.stringify(cdData));

                var soId = cdData.soId;
                if (!soId) {
                    return { 
                        error: 'No Sales Order linked to Customer Deposit', 
                        creditMemo: cmData,
                        customerDeposit: cdData
                    };
                }

                // STEP 4: Get customer ID from SO to query all customer SOs
                var soData = getSalesOrderDetails(soId);
                log.debug('SO Data', JSON.stringify(soData));
                
                var customerId = soData.customerId;
                log.debug('Customer ID', customerId);

                // STEP 5: Get ALL Sales Orders for this customer (filtered by status)
                var allCustomerSOs = getCustomerSalesOrders(customerId, soId);
                log.debug('Customer SOs', 'Found ' + allCustomerSOs.length + ' sales orders for customer');

                // STEP 6: Run comparison for each SO
                var soComparisons = [];
                for (var soIdx = 0; soIdx < allCustomerSOs.length; soIdx++) {
                    var currentSO = allCustomerSOs[soIdx];
                    log.debug('Processing SO', 'ID: ' + currentSO.id + ', tranid: ' + currentSO.tranid);
                    
                    var soComparison = runSingleSOComparison(currentSO.id);
                    soComparison.isSource = currentSO.isSource;
                    soComparison.status = currentSO.status;
                    soComparison.statusDisplay = currentSO.statusDisplay;
                    soComparisons.push(soComparison);
                }

                // STEP 7: Calculate customer-level totals
                var customerTotals = calculateCustomerTotals(soComparisons);
                log.debug('Customer Totals', JSON.stringify(customerTotals));

                // STEP 8: Find the source SO comparison for backward compatibility
                var sourceSoComparison = null;
                for (var i = 0; i < soComparisons.length; i++) {
                    if (soComparisons[i].isSource) {
                        sourceSoComparison = soComparisons[i];
                        break;
                    }
                }

                // STEP 9: Build result object (maintains backward compatibility + new multi-SO data)
                var result = {
                    // Original single-SO fields (for backward compatibility)
                    creditMemo: cmData,
                    customerDeposit: cdData,
                    salesOrder: sourceSoComparison ? sourceSoComparison.salesOrder : soData,
                    invoices: sourceSoComparison ? sourceSoComparison.invoices : [],
                    comparisonTable: sourceSoComparison ? sourceSoComparison.comparisonTable : [],
                    totals: sourceSoComparison ? sourceSoComparison.totals : {},
                    problemItems: sourceSoComparison ? sourceSoComparison.problemItems : [],
                    allRelatedCreditMemos: sourceSoComparison ? sourceSoComparison.relatedCreditMemos : [],
                    totalCMAmount: sourceSoComparison ? sourceSoComparison.totalCMAmount : 0,
                    varianceMatchesCMs: sourceSoComparison ? sourceSoComparison.varianceMatchesCMs : false,
                    conclusion: sourceSoComparison ? sourceSoComparison.conclusion : '',
                    
                    // NEW: Multi-SO data
                    customerId: customerId,
                    customerName: soData.customer_name,
                    allSalesOrders: allCustomerSOs,
                    soComparisons: soComparisons,
                    customerTotals: customerTotals
                };

                log.debug('Analysis Complete', 'Total SOs: ' + allCustomerSOs.length);
                return result;

            } catch (e) {
                log.error('Error in analyzeCreditMemoOverpayment', {
                    error: e.message,
                    stack: e.stack,
                    creditMemoId: creditMemoId
                });
                return { error: e.message, creditMemoId: creditMemoId };
            }
        }

        // ============================================================================
        // SO LIFECYCLE SYSTEM NOTES FUNCTIONS
        // ============================================================================

        /**
         * Gets Sales Order lifecycle data (system notes) for display in modal
         * @param {number} creditMemoId - Internal ID of the Credit Memo
         * @returns {Object} SO info and system notes
         */
        function getSOLifecycleData(creditMemoId) {
            try {
                log.debug('getSOLifecycleData', 'Starting for CM ID: ' + creditMemoId);

                // Get Credit Memo details using existing helper
                var cmData = getCreditMemoDetails(creditMemoId);
                if (!cmData) {
                    return { error: 'Credit Memo not found' };
                }

                // Extract CD tranid from memo using existing helper
                var cdTranid = extractCDFromMemo(cmData.memo);
                if (!cdTranid) {
                    return { error: 'Could not extract Customer Deposit from Credit Memo memo field' };
                }

                log.debug('Extracted CD', cdTranid);

                // Get Customer Deposit and linked SO using existing helper
                var cdData = getCustomerDepositDetails(cdTranid);
                if (!cdData || !cdData.soId) {
                    return { error: 'Customer Deposit or linked Sales Order not found for ' + cdTranid };
                }

                log.debug('CD Found', {
                    cdId: cdData.id,
                    cdTranid: cdData.tranid,
                    soId: cdData.soId,
                    soTranid: cdData.soTranid
                });

                // Get Sales Order details using existing helper
                var soData = getSalesOrderDetails(cdData.soId);
                if (!soData) {
                    return { error: 'Sales Order not found for ID: ' + cdData.soId };
                }

                // Query system notes for this SO (all notes, no date filter)
                var sql = `
                    SELECT 
                        sn.id,
                        sn.date,
                        sn.field,
                        sn.oldvalue,
                        sn.newvalue,
                        sn.name,
                        e.firstname || ' ' || e.lastname AS user_name
                    FROM systemNote sn
                    LEFT JOIN employee e ON sn.name = e.id
                    WHERE sn.recordid = ` + cdData.soId + `
                    ORDER BY sn.date DESC, sn.id DESC
                `;

                var systemNotesResults = query.runSuiteQL({ query: sql }).asMappedResults();
                
                log.debug('System Notes Query', {
                    soId: cdData.soId,
                    notesFound: systemNotesResults.length
                });

                // Format system notes for display
                var formattedNotes = [];
                for (var i = 0; i < systemNotesResults.length; i++) {
                    var note = systemNotesResults[i];
                    formattedNotes.push({
                        date: note.date || '',
                        field: note.field || '',
                        oldValue: note.oldvalue || '',
                        newValue: note.newvalue || '',
                        setBy: note.user_name || 'System'
                    });
                }

                return {
                    soInfo: {
                        soNumber: soData.tranid,
                        soDate: soData.trandate,
                        total: soData.foreigntotal,
                        customerName: soData.customer_name,
                        soId: cdData.soId
                    },
                    systemNotes: formattedNotes
                };

            } catch (e) {
                log.error('Error in getSOLifecycleData', {
                    error: e.message,
                    stack: e.stack,
                    creditMemoId: creditMemoId
                });
                return { error: e.message };
            }
        }

        // ============================================================================
        // CROSS-SO DEPOSIT ANALYSIS FUNCTIONS
        // ============================================================================

        /**
         * Analyzes Customer Deposit applications for cross-SO mismatches
         * @param {number} creditMemoId - Internal ID of the Credit Memo
         * @returns {Object} Analysis result with matches and mismatches
         */
        function analyzeCrossSODeposits(creditMemoId) {
            try {
                log.debug('Cross-SO Analysis', 'Starting for CM ID: ' + creditMemoId);

                // Step 1: Get customer from Credit Memo
                var cmRecord = record.load({
                    type: record.Type.CREDIT_MEMO,
                    id: creditMemoId
                });
                var customerId = cmRecord.getValue({ fieldId: 'entity' });
                var customerName = cmRecord.getText({ fieldId: 'entity' });
                
                log.debug('Customer Found', 'ID: ' + customerId + ', Name: ' + customerName);

                // Step 2: Query all deposit applications for this customer
                // Uses previousTransactionLineLink table to find relationships:
                //   - linktype='DepAppl': CD (previousdoc) -> DEPA (nextdoc)
                //   - linktype='Payment': INV (previousdoc) -> DEPA (nextdoc)
                var sql = `
                    SELECT 
                        depa.id as depaId,
                        depa.tranid as depaTranId,
                        cd.id as cdId,
                        cd.tranid as cdTranId,
                        cd_tl.createdfrom as cdSourceSOId,
                        cd_so.tranid as cdSourceSO,
                        inv.id as invId,
                        inv.tranid as invTranId,
                        inv_tl.createdfrom as invSourceSOId,
                        inv_so.tranid as invSourceSO,
                        inv.custbody_overpayment_tran as invOverpaymentCD,
                        depa_tl.netamount as amount,
                        -- Get the applied-to transaction (could be invoice, refund, etc)
                        applied_tran.id as appliedTranId,
                        applied_tran.tranid as appliedTranNumber,
                        applied_tran.type as appliedTranType
                    FROM 
                        transaction depa
                        -- Link DEPA to CD via DepAppl linktype
                        INNER JOIN previousTransactionLineLink ptll_cd 
                            ON depa.id = ptll_cd.nextdoc 
                            AND ptll_cd.linktype = 'DepAppl'
                        INNER JOIN transaction cd 
                            ON ptll_cd.previousdoc = cd.id
                        -- Get CD's source SO from transactionLine.createdfrom
                        INNER JOIN transactionLine cd_tl 
                            ON cd.id = cd_tl.transaction 
                            AND cd_tl.mainline = 'T'
                        LEFT JOIN transaction cd_so 
                            ON cd_tl.createdfrom = cd_so.id
                        -- Link DEPA to the applied-to transaction via Payment linktype
                        LEFT JOIN previousTransactionLineLink ptll_applied 
                            ON depa.id = ptll_applied.nextdoc 
                            AND ptll_applied.linktype = 'Payment'
                        LEFT JOIN transaction applied_tran 
                            ON ptll_applied.previousdoc = applied_tran.id
                        -- Link DEPA to INV via Payment linktype (LEFT JOIN to include non-invoice applications)
                        LEFT JOIN previousTransactionLineLink ptll_inv 
                            ON depa.id = ptll_inv.nextdoc 
                            AND ptll_inv.linktype = 'Payment'
                        LEFT JOIN transaction inv 
                            ON ptll_inv.previousdoc = inv.id
                            AND inv.type = 'CustInvc'
                        -- Get INV's source SO from transactionLine.createdfrom
                        LEFT JOIN transactionLine inv_tl 
                            ON inv.id = inv_tl.transaction 
                            AND inv_tl.mainline = 'T'
                        LEFT JOIN transaction inv_so 
                            ON inv_tl.createdfrom = inv_so.id
                        -- Get DEPA amount from mainline
                        INNER JOIN transactionLine depa_tl 
                            ON depa.id = depa_tl.transaction 
                            AND depa_tl.mainline = 'T'
                    WHERE 
                        depa.entity = ${customerId}
                        AND depa.type = 'DepAppl'
                        AND cd.type = 'CustDep'
                    ORDER BY 
                        cd.tranid ASC, depa.id ASC
                `;

                log.debug('SQL Query', sql);

                var resultSet = query.runSuiteQL({ query: sql });
                var results = resultSet.asMappedResults();

                log.debug('Query Results Count', results.length);

                // Step 3: Get the Credit Memo memo field to find the overpayment CD
                var cmMemo = cmRecord.getValue({ fieldId: 'memo' }) || '';
                var overpaymentCDMatch = cmMemo.match(/CD\d+/);
                var overpaymentCDTranId = overpaymentCDMatch ? overpaymentCDMatch[0] : null;
                
                log.debug('Overpayment CD from memo', overpaymentCDTranId);

                // Step 4: Process results
                var applications = [];
                var matches = 0;
                var mismatches = 0;
                var noInvoice = 0;
                var crossSOAmount = 0;

                for (var i = 0; i < results.length; i++) {
                    var row = results[i];
                    
                    var cdSourceSO = row.cdsourceso || 'N/A';
                    var invSourceSO = row.invsourceso || null;
                    var invOverpaymentCD = row.invoverpaymentcd ? String(row.invoverpaymentcd) : null;
                    var cdId = String(row.cdid);
                    var isOverpaymentCDTranId = (row.cdtranid === overpaymentCDTranId);
                    
                    // Determine status
                    var status = 'match';
                    var isMismatch = false;
                    var isNoInvoice = false;
                    // Highlight ALL overpayment invoices: any invoice where custbody_overpayment_tran points to this CD
                    var isOverpaymentInv = (invOverpaymentCD === cdId);
                    
                    if (!invSourceSO) {
                        // No invoice (applied to refund or other non-invoice transaction)
                        status = 'no-invoice';
                        isNoInvoice = true;
                        noInvoice++;
                    } else if (cdSourceSO !== 'N/A' && invSourceSO !== 'N/A' && cdSourceSO !== invSourceSO) {
                        // True cross-SO mismatch
                        status = 'cross-so';
                        isMismatch = true;
                        mismatches++;
                        crossSOAmount += parseFloat(row.amount || 0);
                    } else {
                        // Match
                        matches++;
                    }

                    // Get applied-to transaction info
                    var appliedTranType = row.appliedtrantype || null;
                    var appliedTranNumber = row.appliedtrannumber || null;
                    var appliedTranId = row.appliedtranid || null;
                    
                    // Format transaction type for display
                    var appliedTranTypeDisplay = '';
                    if (appliedTranType === 'CustInvc') {
                        appliedTranTypeDisplay = 'Invoice';
                    } else if (appliedTranType === 'CustRfnd') {
                        appliedTranTypeDisplay = 'Refund';
                    } else if (appliedTranType === 'CustCred') {
                        appliedTranTypeDisplay = 'Credit Memo';
                    } else if (appliedTranType) {
                        appliedTranTypeDisplay = appliedTranType;
                    }

                    applications.push({
                        depaId: row.depaid,
                        depaTranId: row.depatranid,
                        cdId: cdId,
                        cdTranId: row.cdtranid,
                        cdSourceSO: cdSourceSO,
                        invId: row.invid,
                        invTranId: row.invtranid,
                        invSourceSO: invSourceSO || 'N/A',
                        amount: parseFloat(row.amount || 0),
                        status: status,
                        isMismatch: isMismatch,
                        isNoInvoice: isNoInvoice,
                        isOverpaymentInv: isOverpaymentInv,
                        appliedTranType: appliedTranTypeDisplay,
                        appliedTranNumber: appliedTranNumber,
                        appliedTranId: appliedTranId
                    });
                }

                // Step 5: Sort applications - Cross-SO CDs first, then by CD tranid
                applications.sort(function(a, b) {
                    // First, check if either CD has a cross-SO mismatch
                    var aCDHasMismatch = false;
                    var bCDHasMismatch = false;
                    
                    // Check if this CD has ANY cross-SO mismatches
                    for (var j = 0; j < applications.length; j++) {
                        if (applications[j].cdTranId === a.cdTranId && applications[j].isMismatch) {
                            aCDHasMismatch = true;
                        }
                        if (applications[j].cdTranId === b.cdTranId && applications[j].isMismatch) {
                            bCDHasMismatch = true;
                        }
                    }
                    
                    // CDs with cross-SO mismatches come first
                    if (aCDHasMismatch && !bCDHasMismatch) return -1;
                    if (!aCDHasMismatch && bCDHasMismatch) return 1;
                    
                    // Then sort by CD tranid
                    if (a.cdTranId < b.cdTranId) return -1;
                    if (a.cdTranId > b.cdTranId) return 1;
                    
                    // Within same CD, put mismatches first, then matches, then no-invoice
                    var statusOrder = { 'cross-so': 1, 'match': 2, 'no-invoice': 3 };
                    return (statusOrder[a.status] || 99) - (statusOrder[b.status] || 99);
                });

                log.debug('Analysis Complete', 'Total: ' + applications.length + ', Matches: ' + matches + ', Mismatches: ' + mismatches + ', No Invoice: ' + noInvoice);

                return {
                    customerId: customerId,
                    customerName: customerName,
                    totalApplications: applications.length,
                    matches: matches,
                    mismatches: mismatches,
                    noInvoice: noInvoice,
                    crossSOAmount: crossSOAmount,
                    applications: applications,
                    overpaymentCDTranId: overpaymentCDTranId
                };

            } catch (e) {
                log.error('Cross-SO Analysis Error', {
                    error: e.message,
                    stack: e.stack,
                    creditMemoId: creditMemoId
                });
                return { error: e.message, creditMemoId: creditMemoId };
            }
        }

        /**
         * Gets Overpayment Summary for all overpayment CMs on customer record
         * @param {number} creditMemoId - Internal ID of source Credit Memo
         * @returns {Object} Overpayment summary with calculations
         */
        function getOverpaymentSummary(creditMemoId) {
            try {
                log.debug('Overpayment Summary', 'Starting for CM ID: ' + creditMemoId);

                // Step 1: Get customer from source Credit Memo
                var cmRecord = record.load({
                    type: record.Type.CREDIT_MEMO,
                    id: creditMemoId
                });
                var customerId = cmRecord.getValue({ fieldId: 'entity' });
                
                log.debug('Customer Found', 'ID: ' + customerId);

                // Step 2: Query all overpayment CMs for this customer
                // Get SO via CD's createdfrom field since custbody_overpayment_linked_sales_orde isn't exposed for SuiteQL
                var cmSql = `
                    SELECT 
                        cm.id as cm_id,
                        cm.tranid as cm_number,
                        ABS(cm.foreigntotal) as cm_amount,
                        cm.custbody_overpayment_date as overpayment_date,
                        cm.custbody_overpayment_tran as cd_id,
                        cd.tranid as cd_number,
                        tl_cd.createdfrom as so_id,
                        so.tranid as so_number,
                        ABS(so.foreigntotal) as so_total
                    FROM transaction cm
                    LEFT JOIN transaction cd ON cm.custbody_overpayment_tran = cd.id
                    LEFT JOIN transactionline tl_cd ON cd.id = tl_cd.transaction AND tl_cd.mainline = 'T'
                    LEFT JOIN transaction so ON tl_cd.createdfrom = so.id
                    WHERE cm.type = 'CustCred'
                      AND cm.entity = ` + customerId + `
                      AND cm.custbody_overpayment_tran IS NOT NULL
                    ORDER BY 
                      CASE WHEN cm.id = ` + creditMemoId + ` THEN 0 ELSE 1 END,
                      cm.trandate DESC
                `;

                var cmResults = query.runSuiteQL({ query: cmSql }).asMappedResults();
                log.debug('Found CMs', cmResults.length + ' overpayment credit memos');

                // Step 3: Group CMs by Sales Order and calculate totals
                var soMap = {}; // Map of SO ID to SO data with CMs
                var totalCMAmount = 0;

                for (var i = 0; i < cmResults.length; i++) {
                    var cm = cmResults[i];
                    var soId = cm.so_id;
                    var cmAmount = parseFloat(cm.cm_amount || 0);
                    totalCMAmount += cmAmount;

                    // Initialize SO in map if not exists
                    if (!soMap[soId]) {
                        // Get total deposits collected on this SO
                        var depositsSql = `
                            SELECT COALESCE(SUM(ABS(cd.foreigntotal)), 0) as total_deposits
                            FROM transaction cd
                            INNER JOIN transactionline tl ON cd.id = tl.transaction
                            WHERE tl.createdfrom = ` + soId + `
                              AND cd.type = 'CustDep'
                              AND tl.mainline = 'T'
                        `;
                        var depositsResult = query.runSuiteQL({ query: depositsSql }).asMappedResults();
                        var totalDeposits = parseFloat(depositsResult[0].total_deposits || 0);

                        // Get total invoiced value from this SO
                        var invoicesSql = `
                            SELECT COALESCE(SUM(ABS(inv.foreigntotal)), 0) as total_invoiced
                            FROM transaction inv
                            INNER JOIN transactionline tl ON inv.id = tl.transaction
                            WHERE tl.createdfrom = ` + soId + `
                              AND inv.type = 'CustInvc'
                              AND tl.mainline = 'T'
                        `;
                        var invoicesResult = query.runSuiteQL({ query: invoicesSql }).asMappedResults();
                        var totalInvoiced = parseFloat(invoicesResult[0].total_invoiced || 0);

                        soMap[soId] = {
                            soId: soId,
                            soNumber: cm.so_number,
                            soTotal: parseFloat(cm.so_total || 0),
                            totalDeposits: totalDeposits,
                            totalInvoiced: totalInvoiced,
                            overpaymentVariance: totalDeposits - totalInvoiced,
                            creditMemos: []
                        };
                    }

                    // Add CM to this SO's list
                    soMap[soId].creditMemos.push({
                        cmId: cm.cm_id,
                        cmNumber: cm.cm_number,
                        cmAmount: cmAmount,
                        overpaymentDate: cm.overpayment_date || 'N/A',
                        cdId: cm.cd_id,
                        cdNumber: cm.cd_number,
                        isSource: (cm.cm_id === creditMemoId)
                    });
                }

                // Convert map to array
                var salesOrders = [];
                for (var soId in soMap) {
                    if (soMap.hasOwnProperty(soId)) {
                        salesOrders.push(soMap[soId]);
                    }
                }

                // Sort: source CM's SO first
                salesOrders.sort(function(a, b) {
                    var aHasSource = a.creditMemos.some(function(cm) { return cm.isSource; });
                    var bHasSource = b.creditMemos.some(function(cm) { return cm.isSource; });
                    if (aHasSource && !bHasSource) return -1;
                    if (!aHasSource && bHasSource) return 1;
                    return 0;
                });

                return {
                    salesOrders: salesOrders,
                    totalCMCount: cmResults.length,
                    totalCMAmount: totalCMAmount
                };

            } catch (e) {
                log.error('Overpayment Summary Error', {
                    error: e.message,
                    stack: e.stack,
                    creditMemoId: creditMemoId
                });
                return { error: e.message };
            }
        }

        /**
         * Gets Credit Memo details
         * @param {number} cmId - Credit Memo internal ID
         * @returns {Object} Credit Memo data
         */
        function getCreditMemoDetails(cmId) {
            var sql = `
                SELECT t.id,
                       t.tranid,
                       t.trandate,
                       t.foreigntotal,
                       t.entity,
                       BUILTIN.DF(t.entity) as customer_name,
                       t.memo
                FROM transaction t
                WHERE t.id = ` + cmId + `
                AND t.type = 'CustCred'
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            if (results.length === 0) return null;

            var row = results[0];
            return {
                id: row.id,
                tranid: row.tranid,
                trandate: row.trandate,
                amount: Math.abs(parseFloat(row.foreigntotal) || 0),
                customerId: row.entity,
                customerName: row.customer_name,
                memo: row.memo || ''
            };
        }

        /**
         * Extracts CD tranid from Credit Memo memo field
         * @param {string} memo - Memo field content
         * @returns {string|null} CD tranid or null
         */
        function extractCDFromMemo(memo) {
            if (!memo) return null;
            
            // Pattern: "Overpayment from Customer Deposit: CD5787" or similar variations
            var patterns = [
                /Customer Deposit:\s*(CD\d+)/i,
                /CD:\s*(CD\d+)/i,
                /(CD\d+)/i  // fallback - just find CDxxxx
            ];

            for (var i = 0; i < patterns.length; i++) {
                var match = memo.match(patterns[i]);
                if (match && match[1]) {
                    return match[1];
                }
            }

            return null;
        }

        /**
         * Gets Customer Deposit details and traces to Sales Order
         * @param {string} cdTranid - Customer Deposit tranid (e.g., "CD5787")
         * @returns {Object} Customer Deposit data
         */
        function getCustomerDepositDetails(cdTranid) {
            var sql = `
                SELECT t.id,
                       t.tranid,
                       t.trandate,
                       t.foreigntotal,
                       t.entity,
                       BUILTIN.DF(t.entity) as customer_name,
                       tl.createdfrom,
                       t2.tranid as source_so,
                       t2.foreigntotal as so_amount,
                       t2.id as so_id
                FROM transaction t
                INNER JOIN transactionline tl ON t.id = tl.transaction
                LEFT JOIN transaction t2 ON tl.createdfrom = t2.id
                WHERE t.tranid = '` + cdTranid + `'
                AND t.type = 'CustDep'
                AND tl.mainline = 'T'
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            if (results.length === 0) return null;

            var row = results[0];
            return {
                id: row.id,
                tranid: row.tranid,
                trandate: row.trandate,
                amount: parseFloat(row.foreigntotal) || 0,
                customerId: row.entity,
                customerName: row.customer_name,
                soId: row.so_id,
                soTranid: row.source_so,
                soAmount: parseFloat(row.so_amount) || 0
            };
        }

        /**
         * Gets Sales Order details
         * @param {number} soId - Sales Order internal ID
         * @returns {Object} Sales Order data
         */
        function getSalesOrderDetails(soId) {
            var sql = `
                SELECT t.id,
                       t.tranid,
                       t.trandate,
                       t.foreigntotal,
                       t.status,
                       BUILTIN.DF(t.status) as status_text,
                       t.entity,
                       BUILTIN.DF(t.entity) as customer_name
                FROM transaction t
                WHERE t.id = ` + soId + `
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            if (results.length === 0) return null;

            var row = results[0];
            return {
                id: row.id,
                tranid: row.tranid,
                trandate: row.trandate,
                amount: parseFloat(row.foreigntotal) || 0,
                status: row.status,
                statusText: row.status_text,
                customerId: row.entity,
                customerName: row.customer_name
            };
        }

        /**
         * Gets ALL line items from Sales Order
         * Handles discount items specially (uses amount field, not rate which may be %)
         * For normal items, uses rate * qty to get FULL amount before discount
         * @param {number} soId - Sales Order internal ID
         * @returns {Array} Array of line items
         */
        function getSOLineItems(soId) {
            var sql = `
                SELECT tl.item,
                       BUILTIN.DF(tl.item) as item_name,
                       i.itemtype,
                       tl.memo as line_description,
                       tl.quantity,
                       tl.rate as so_rate,
                       tl.foreignamount as so_amount,
                       tl.netamount as so_netamount,
                       tl.linesequencenumber,
                       tl.uniquekey
                FROM transactionline tl
                INNER JOIN item i ON tl.item = i.id
                WHERE tl.transaction = ` + soId + `
                AND tl.mainline = 'F'
                AND tl.item IS NOT NULL
                AND tl.item > 0
                ORDER BY tl.linesequencenumber
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            var lineItems = [];

            for (var i = 0; i < results.length; i++) {
                var row = results[i];
                var qty = parseFloat(row.quantity) || 0;
                var rate = parseFloat(row.so_rate) || 0;
                var foreignAmount = parseFloat(row.so_amount) || 0;
                var itemType = row.itemtype || '';
                var isDiscount = (itemType === 'Discount');
                
                // Quantity: NetSuite stores as negative (-1), flip to positive for display
                // Discounts typically have null/0 qty, keep as-is
                var displayQty = isDiscount ? qty : -qty;
                
                // Amount calculation:
                // - Normal items: rate * displayQty (positive rate * positive qty = positive amount)
                // - Discounts: -foreignamount (foreignamount is positive, prefix with - to make negative)
                //   We use foreignamount for discounts because rate could be a percentage
                var soAmount;
                if (isDiscount) {
                    soAmount = -foreignAmount;  // 48.4 becomes -48.4
                } else {
                    soAmount = rate * displayQty;  // 968 * 1 = 968
                }
                
                // Rate display:
                // - Normal items: rate as-is (positive)
                // - Discounts: show the dollar amount (negative), not the percentage rate
                var displayRate = isDiscount ? soAmount : rate;
                
                lineItems.push({
                    item: row.item,
                    itemName: row.item_name,
                    itemType: itemType,
                    isDiscount: isDiscount,
                    lineDescription: row.line_description || '',
                    quantity: displayQty,
                    soRate: displayRate,
                    soAmount: soAmount,
                    lineSequence: row.linesequencenumber,
                    uniqueKey: row.uniquekey
                });
            }

            return lineItems;
        }

        /**
         * Gets tax total from Sales Order
         * @param {number} soId - Sales Order internal ID
         * @returns {number} Tax total
         */
        function getSOTaxTotal(soId) {
            var sql = `
                SELECT SUM(tl.netamount) as tax_total
                FROM transactionline tl
                WHERE tl.transaction = ` + soId + `
                AND tl.mainline = 'F'
                AND tl.taxline = 'T'
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            if (results.length === 0) return 0;

            return parseFloat(results[0].tax_total) || 0;
        }

        /**
         * Finds ALL invoices created from a Sales Order
         * @param {number} soId - Sales Order internal ID
         * @returns {Array} Array of invoice internal IDs
         */
        function findInvoicesFromSO(soId) {
            var sql = `
                SELECT DISTINCT tl.transaction as invoice_id
                FROM transactionline tl
                INNER JOIN transaction t ON tl.transaction = t.id
                WHERE tl.createdfrom = ` + soId + `
                AND t.type = 'CustInvc'
                ORDER BY tl.transaction
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            var invoiceIds = [];

            for (var i = 0; i < results.length; i++) {
                invoiceIds.push(results[i].invoice_id);
            }

            return invoiceIds;
        }

        /**
         * Gets invoice details
         * @param {Array} invoiceIds - Array of invoice internal IDs
         * @returns {Array} Array of invoice details
         */
        function getInvoiceDetails(invoiceIds) {
            if (!invoiceIds || invoiceIds.length === 0) return [];

            var sql = `
                SELECT t.id,
                       t.tranid,
                       t.trandate,
                       t.foreigntotal
                FROM transaction t
                WHERE t.id IN (` + invoiceIds.join(',') + `)
                ORDER BY t.trandate
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            var invoices = [];

            for (var i = 0; i < results.length; i++) {
                var row = results[i];
                invoices.push({
                    id: row.id,
                    tranid: row.tranid,
                    trandate: row.trandate,
                    amount: parseFloat(row.foreigntotal) || 0
                });
            }

            return invoices;
        }

        /**
         * Gets ALL line items from ALL invoices
         * Handles discount items specially (uses amount field, not rate which can be %)
         * @param {Array} invoiceIds - Array of invoice internal IDs
         * @returns {Array} Array of invoice line items
         */
        function getInvoiceLineItems(invoiceIds) {
            if (!invoiceIds || invoiceIds.length === 0) return [];

            var sql = `
                SELECT t.id as invoice_id,
                       t.tranid as invoice_number,
                       t.trandate,
                       tl.item,
                       BUILTIN.DF(tl.item) as item_name,
                       i.itemtype,
                       tl.memo as line_description,
                       tl.quantity,
                       tl.rate as invoice_rate,
                       tl.foreignamount as inv_amount,
                       tl.netamount as inv_netamount,
                       tl.linesequencenumber
                FROM transaction t
                INNER JOIN transactionline tl ON t.id = tl.transaction
                INNER JOIN item i ON tl.item = i.id
                WHERE t.id IN (` + invoiceIds.join(',') + `)
                AND tl.mainline = 'F'
                AND tl.item IS NOT NULL
                AND tl.item > 0
                ORDER BY t.trandate, tl.linesequencenumber
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            var lineItems = [];

            for (var i = 0; i < results.length; i++) {
                var row = results[i];
                var qty = parseFloat(row.quantity) || 0;
                var rate = parseFloat(row.invoice_rate) || 0;
                var foreignAmount = parseFloat(row.inv_amount) || 0;
                var itemType = row.itemtype || '';
                var isDiscount = (itemType === 'Discount');
                
                // Quantity: NetSuite stores as negative (-1), flip to positive for display
                // Discounts typically have null/0 qty, keep as-is
                var displayQty = isDiscount ? qty : -qty;
                
                // Amount calculation:
                // - Normal items: rate * displayQty (positive rate * positive qty = positive amount)
                // - Discounts: -foreignamount (foreignamount is positive, prefix with - to make negative)
                //   We use foreignamount for discounts because rate could be a percentage
                var invoiceAmount;
                if (isDiscount) {
                    invoiceAmount = -foreignAmount;  // 48.4 becomes -48.4
                } else {
                    invoiceAmount = rate * displayQty;  // 968 * 1 = 968
                }
                
                // Rate display:
                // - Normal items: rate as-is (positive)
                // - Discounts: show the dollar amount (negative), not the percentage rate
                var displayRate = isDiscount ? invoiceAmount : rate;
                
                lineItems.push({
                    invoiceId: row.invoice_id,
                    invoiceNumber: row.invoice_number,
                    invoiceDate: row.trandate,
                    item: row.item,
                    itemName: row.item_name,
                    itemType: itemType,
                    isDiscount: isDiscount,
                    lineDescription: row.line_description || '',
                    quantity: displayQty,
                    invoiceRate: displayRate,
                    invoiceAmount: invoiceAmount,
                    lineSequence: row.linesequencenumber
                });
            }

            return lineItems;
        }

        /**
         * Gets tax total from ALL invoices
         * @param {Array} invoiceIds - Array of invoice internal IDs
         * @returns {number} Total tax amount
         */
        function getInvoiceTaxTotal(invoiceIds) {
            if (!invoiceIds || invoiceIds.length === 0) return 0;

            var sql = `
                SELECT SUM(tl.netamount) as tax_total
                FROM transaction t
                INNER JOIN transactionline tl ON t.id = tl.transaction
                WHERE t.id IN (` + invoiceIds.join(',') + `)
                AND tl.mainline = 'F'
                AND tl.taxline = 'T'
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            if (results.length === 0) return 0;

            return parseFloat(results[0].tax_total) || 0;
        }

        /**
         * Matches SO line items to Invoice line items
         * Aggregates BOTH SO and Invoice items by item ID to avoid duplicate rows
         * @param {Array} soLineItems - SO line items
         * @param {Array} invoiceLineItems - Invoice line items
         * @returns {Array} Comparison table
         */
        function matchLineItems(soLineItems, invoiceLineItems) {
            var comparisonTable = [];
            
            // Aggregate SO line items by item ID (in case item appears multiple times on SO)
            var soItemMap = {};
            for (var i = 0; i < soLineItems.length; i++) {
                var soItem = soLineItems[i];
                var itemKey = soItem.item;
                
                if (!soItemMap[itemKey]) {
                    soItemMap[itemKey] = {
                        item: soItem.item,
                        itemName: soItem.itemName,
                        lineDescription: soItem.lineDescription,
                        isDiscount: soItem.isDiscount || false,
                        totalQuantity: 0,
                        totalAmount: 0,
                        rates: []
                    };
                }
                
                soItemMap[itemKey].totalQuantity += soItem.quantity;
                soItemMap[itemKey].totalAmount += soItem.soAmount;
                soItemMap[itemKey].rates.push(soItem.soRate);
                // If any line for this item is a discount, mark it
                if (soItem.isDiscount) soItemMap[itemKey].isDiscount = true;
            }
            
            // Aggregate invoice line items by item ID (in case item appears on multiple invoices)
            var invoiceItemMap = {};
            for (var j = 0; j < invoiceLineItems.length; j++) {
                var invItem = invoiceLineItems[j];
                var itemKey = invItem.item;
                
                if (!invoiceItemMap[itemKey]) {
                    invoiceItemMap[itemKey] = {
                        item: invItem.item,
                        itemName: invItem.itemName,
                        lineDescription: invItem.lineDescription,
                        isDiscount: invItem.isDiscount || false,
                        invoices: [],
                        totalQuantity: 0,
                        totalAmount: 0,
                        rates: []
                    };
                }
                
                invoiceItemMap[itemKey].invoices.push(invItem.invoiceNumber);
                invoiceItemMap[itemKey].totalQuantity += invItem.quantity;
                invoiceItemMap[itemKey].totalAmount += invItem.invoiceAmount;
                invoiceItemMap[itemKey].rates.push(invItem.invoiceRate);
                // If any line for this item is a discount, mark it
                if (invItem.isDiscount) invoiceItemMap[itemKey].isDiscount = true;
            }

            // Create a set of all unique item IDs from both SO and Invoice
            var allItemIds = {};
            for (var soKey in soItemMap) {
                allItemIds[soKey] = true;
            }
            for (var invKey in invoiceItemMap) {
                allItemIds[invKey] = true;
            }

            // Loop through each unique item ID and compare aggregated totals
            for (var itemId in allItemIds) {
                var soData = soItemMap[itemId];
                var invData = invoiceItemMap[itemId];
                
                var itemName = soData ? soData.itemName : (invData ? invData.itemName : '');
                var lineDescription = soData ? soData.lineDescription : (invData ? invData.lineDescription : '');
                
                var soQty = soData ? soData.totalQuantity : 0;
                var soRate = soData && soData.rates.length > 0 ? soData.rates[0] : 0;
                var soAmount = soData ? soData.totalAmount : 0;
                
                var invoiceNumbers = '-';
                var invoiceQty = 0;
                var invoiceRate = 0;
                var invoiceAmount = 0;
                
                if (invData) {
                    // Get unique invoice numbers
                    invoiceNumbers = invData.invoices.filter(function(v, i, a) { return a.indexOf(v) === i; }).join(', ');
                    invoiceQty = invData.totalQuantity;
                    invoiceRate = invData.rates.length > 0 ? invData.rates[0] : 0;
                    invoiceAmount = invData.totalAmount;
                }

                // Determine if this item is a discount (from either SO or Invoice data)
                var isDiscount = (soData && soData.isDiscount) || (invData && invData.isDiscount) || false;
                
                // Status logic and variance split:
                // - mismatchVariance: variance from pricing errors (causes overpayments)
                // - unbilledVariance: variance from items not yet invoiced (just pending)
                var status;
                var mismatchVariance = 0;
                var unbilledVariance = 0;
                var totalRowVariance = invoiceAmount - soAmount;
                
                if (isDiscount) {
                    status = 'DISCOUNT';
                    // Discounts always contribute to mismatch (missing discount = overpayment)
                    mismatchVariance = totalRowVariance;
                } else if (!soData) {
                    status = 'NOT ON SO';
                    // Item on invoice but not SO is a mismatch problem
                    mismatchVariance = totalRowVariance;
                } else if (!invData) {
                    status = 'NOT INVOICED';
                    // Item on SO but not invoiced - this is unbilled, not a mismatch
                    unbilledVariance = totalRowVariance;  // Will be negative (0 - soAmount)
                } else if (Math.abs(totalRowVariance) > 0.01) {
                    status = 'MISMATCH';
                    mismatchVariance = totalRowVariance;
                } else {
                    status = 'Match';
                    // Both variances stay 0
                }
                
                var comparisonRow = {
                    itemId: itemId,
                    itemName: itemName,
                    lineDescription: lineDescription,
                    isDiscount: isDiscount,
                    soQty: soQty,
                    soRate: soRate,
                    soAmount: soAmount,
                    invoiceNumbers: invoiceNumbers,
                    invoiceQty: invoiceQty,
                    invoiceRate: invoiceRate,
                    invoiceAmount: invoiceAmount,
                    mismatchVariance: mismatchVariance,
                    unbilledVariance: unbilledVariance,
                    status: status
                };

                comparisonTable.push(comparisonRow);
            }

            return comparisonTable;
        }

        /**
         * Finds ALL Credit Memos related to a Sales Order through Customer Deposits
         * This traces: SO -> Customer Deposits (applied to SO) -> Credit Memos (overpayments)
         * @param {number} soId - Sales Order internal ID
         * @returns {Array} Array of related Credit Memos with details
         */
        function findAllRelatedCreditMemos(soId) {
            // First, find all Customer Deposits linked to this SO
            // Then find all Credit Memos that reference those deposits in their memo field
            var sql = `
                SELECT DISTINCT cm.id,
                       cm.tranid,
                       cm.trandate,
                       cm.foreigntotal as amount,
                       cm.memo,
                       cd.tranid as source_deposit
                FROM transaction cd
                INNER JOIN transactionline cdl ON cd.id = cdl.transaction
                INNER JOIN transaction cm ON cm.type = 'CustCred'
                WHERE cdl.createdfrom = ` + soId + `
                AND cd.type = 'CustDep'
                AND cdl.mainline = 'T'
                AND cm.memo LIKE '%' || cd.tranid || '%'
                ORDER BY cm.trandate, cm.id
            `;

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            var creditMemos = [];
            var seenIds = {};

            for (var i = 0; i < results.length; i++) {
                var row = results[i];
                // Avoid duplicates
                if (seenIds[row.id]) continue;
                seenIds[row.id] = true;
                
                creditMemos.push({
                    id: row.id,
                    tranid: row.tranid,
                    trandate: row.trandate,
                    amount: Math.abs(parseFloat(row.foreigntotal) || parseFloat(row.amount) || 0),
                    memo: row.memo || '',
                    sourceDeposit: row.source_deposit
                });
            }

            return creditMemos;
        }

        /**
         * Gets all Sales Orders for a customer filtered by status
         * @param {number} customerId - Customer internal ID
         * @param {number} sourceSoId - Source SO ID to mark as primary
         * @returns {Array} Array of SO objects with basic info
         */
        function getCustomerSalesOrders(customerId, sourceSoId) {
            // Query all SOs for customer with status D, E, F, G, H
            // Status D = Partially Fulfilled
            // Status E = Pending Billing/Partially Fulfilled
            // Status F = Pending Billing
            // Status G = Billed
            // Status H = Closed
            var sql = 
                'SELECT t.id, ' +
                '       t.tranid, ' +
                '       t.status, ' +
                '       BUILTIN.DF(t.status) as status_display, ' +
                '       t.trandate, ' +
                '       t.foreigntotal as total ' +
                'FROM transaction t ' +
                'WHERE t.entity = ' + customerId + ' ' +
                '  AND t.type = \'SalesOrd\' ' +
                '  AND t.status IN (\'SalesOrd:D\', \'SalesOrd:E\', \'SalesOrd:F\', \'SalesOrd:G\', \'SalesOrd:H\') ' +
                'ORDER BY ' +
                '  CASE ' +
                '    WHEN t.id = ' + sourceSoId + ' THEN 0 ' +
                '    WHEN t.status IN (\'SalesOrd:E\', \'SalesOrd:F\') THEN 1 ' +
                '    WHEN t.status = \'SalesOrd:D\' THEN 2 ' +
                '    WHEN t.status = \'SalesOrd:G\' THEN 3 ' +
                '    WHEN t.status = \'SalesOrd:H\' THEN 4 ' +
                '  END, ' +
                '  t.trandate DESC';

            var results = query.runSuiteQL({ query: sql }).asMappedResults();
            var salesOrders = [];

            for (var i = 0; i < results.length; i++) {
                var row = results[i];
                salesOrders.push({
                    id: row.id,
                    tranid: row.tranid,
                    status: row.status,
                    statusDisplay: row.status_display || '',
                    trandate: row.trandate,
                    total: parseFloat(row.total) || 0,
                    isSource: (row.id == sourceSoId)
                });
            }

            return salesOrders;
        }

        /**
         * Runs SO-Invoice comparison for a single SO
         * (Reusable version of the original analysis logic)
         * @param {number} soId - Sales Order internal ID
         * @returns {Object} Comparison result for this SO
         */
        function runSingleSOComparison(soId) {
            try {
                // Get SO details
                var soData = getSalesOrderDetails(soId);
                if (!soData) {
                    return { error: 'Sales Order not found', soId: soId };
                }

                // Get SO line items
                var soLineItems = getSOLineItems(soId);
                var soTaxTotal = getSOTaxTotal(soId);

                // Find all invoices from SO
                var invoiceIds = findInvoicesFromSO(soId);

                // Get all invoice line items
                var invoiceLineItems = [];
                var invoiceTaxTotal = 0;
                var invoiceDetails = [];

                if (invoiceIds.length > 0) {
                    invoiceLineItems = getInvoiceLineItems(invoiceIds);
                    invoiceTaxTotal = getInvoiceTaxTotal(invoiceIds);
                    invoiceDetails = getInvoiceDetails(invoiceIds);
                }

                // Match and compare
                var comparisonTable = matchLineItems(soLineItems, invoiceLineItems);

                // Calculate totals
                var totals = calculateTotals(comparisonTable, soTaxTotal, invoiceTaxTotal);

                // Identify problems
                var problemItems = identifyProblemItems(comparisonTable);

                // Find related CMs for this SO
                var relatedCMs = findAllRelatedCreditMemos(soId);
                var totalCMAmount = 0;
                for (var i = 0; i < relatedCMs.length; i++) {
                    totalCMAmount += relatedCMs[i].amount;
                }

                // Build conclusion
                var varianceMatchesCMs = Math.abs(Math.abs(totals.mismatchVariance) - totalCMAmount) < 0.005;
                var conclusion = '';
                if (Math.abs(totals.mismatchVariance) < 0.01) {
                    conclusion = 'No Mismatch Detected';
                } else if (varianceMatchesCMs && relatedCMs.length > 0) {
                    conclusion = 'Mismatch of ' + Math.abs(totals.mismatchVariance).toFixed(2) + ' Matches ' + relatedCMs.length + ' CM(s) Totaling ' + totalCMAmount.toFixed(2);
                } else {
                    conclusion = 'Mismatch Found - May Have Caused Overpayment';
                }

                return {
                    salesOrder: soData,
                    invoices: invoiceDetails,
                    comparisonTable: comparisonTable,
                    totals: totals,
                    problemItems: problemItems,
                    relatedCreditMemos: relatedCMs,
                    totalCMAmount: totalCMAmount,
                    varianceMatchesCMs: varianceMatchesCMs,
                    conclusion: conclusion
                };

            } catch (e) {
                log.error('Error in runSingleSOComparison', {
                    error: e.message,
                    stack: e.stack,
                    soId: soId
                });
                return { error: e.message, soId: soId };
            }
        }

        /**
         * Calculates customer-level totals across all SO comparisons
         * @param {Array} soComparisons - Array of SO comparison results
         * @returns {Object} Customer-level totals
         */
        function calculateCustomerTotals(soComparisons) {
            var customerTotals = {
                totalSOs: soComparisons.length,
                sosWithMismatch: 0,
                sosWithUnbilled: 0,
                sosWithNoIssues: 0,
                totalMismatchVariance: 0,
                totalUnbilledVariance: 0,
                grandTotalVariance: 0
            };

            for (var i = 0; i < soComparisons.length; i++) {
                var comp = soComparisons[i];
                if (comp.error) continue; // Skip errored SOs

                var totals = comp.totals || {};
                var mismatch = totals.mismatchVariance || 0;
                var unbilled = totals.unbilledVariance || 0;

                customerTotals.totalMismatchVariance += mismatch;
                customerTotals.totalUnbilledVariance += unbilled;

                if (Math.abs(mismatch) > 0.01) {
                    customerTotals.sosWithMismatch++;
                }
                if (Math.abs(unbilled) > 0.01) {
                    customerTotals.sosWithUnbilled++;
                }
                if (Math.abs(mismatch) < 0.01 && Math.abs(unbilled) < 0.01) {
                    customerTotals.sosWithNoIssues++;
                }
            }

            customerTotals.grandTotalVariance = customerTotals.totalMismatchVariance + customerTotals.totalUnbilledVariance;

            return customerTotals;
        }

        /**
         * Calculates totals from comparison table
         * Values are already in correct sign: positive for normal items, negative for discounts
         * @param {Array} comparisonTable - Comparison results
         * @param {number} soTaxTotal - SO tax total
         * @param {number} invoiceTaxTotal - Invoice tax total
         * @returns {Object} Totals object
         */
        function calculateTotals(comparisonTable, soTaxTotal, invoiceTaxTotal) {
            var soLineTotal = 0;
            var invoiceLineTotal = 0;
            var mismatchVarianceTotal = 0;
            var unbilledVarianceTotal = 0;

            for (var i = 0; i < comparisonTable.length; i++) {
                var row = comparisonTable[i];
                // Values already have correct sign:
                // Normal items = positive, Discounts = negative
                soLineTotal += (row.soAmount || 0);
                invoiceLineTotal += (row.invoiceAmount || 0);
                mismatchVarianceTotal += (row.mismatchVariance || 0);
                unbilledVarianceTotal += (row.unbilledVariance || 0);
            }

            // Tax totals also need sign flip (they come in negative)
            var soTaxDisplay = Math.abs(soTaxTotal);
            var invTaxDisplay = Math.abs(invoiceTaxTotal);

            var totals = {
                soLineTotal: soLineTotal,
                soTaxTotal: soTaxDisplay,
                soGrandTotal: soLineTotal + soTaxDisplay,

                invoiceLineTotal: invoiceLineTotal,
                invoiceTaxTotal: invTaxDisplay,
                invoiceGrandTotal: invoiceLineTotal + invTaxDisplay,

                lineVariance: invoiceLineTotal - soLineTotal,
                taxVariance: invTaxDisplay - soTaxDisplay,
                totalVariance: (invoiceLineTotal + invTaxDisplay) - (soLineTotal + soTaxDisplay),
                
                // Split variance into mismatch (errors) vs unbilled (pending)
                mismatchVariance: mismatchVarianceTotal,
                unbilledVariance: unbilledVarianceTotal
            };

            return totals;
        }

        /**
         * Identifies problem items with variance
         * @param {Array} comparisonTable - Comparison results
         * @returns {Array} Problem items
         */
        function identifyProblemItems(comparisonTable) {
            var problemItems = [];

            for (var i = 0; i < comparisonTable.length; i++) {
                var row = comparisonTable[i];
                
                // Skip Match status
                if (row.status === 'Match') continue;
                
                // Skip NOT INVOICED items with $0 amounts (qty=0)
                if (row.status === 'NOT INVOICED' && Math.abs(row.soAmount || 0) < 0.01) continue;
                
                // Skip DISCOUNT items with $0 mismatch variance
                if (row.status === 'DISCOUNT' && Math.abs(row.mismatchVariance || 0) < 0.01) continue;

                problemItems.push({
                    itemName: row.itemName,
                    lineDescription: row.lineDescription,
                    soRate: row.soRate,
                    invoiceRate: row.invoiceRate,
                    soAmount: row.soAmount,
                    invoiceAmount: row.invoiceAmount,
                    mismatchVariance: row.mismatchVariance,
                    unbilledVariance: row.unbilledVariance,
                    invoiceNumbers: row.invoiceNumbers,
                    status: row.status
                });
            }

            return problemItems;
        }

        /**
         * Gets comprehensive sales order summary for a customer from a Credit Memo
         * @param {number} creditMemoId - Credit Memo internal ID
         * @returns {Object} Sales order summary with customer name and all SOs
         */
        function getCustomerSalesOrdersSummary(creditMemoId) {
            try {
                // Step 1: Get customer ID from credit memo
                var cmDetails = getCreditMemoDetails(creditMemoId);
                if (!cmDetails || !cmDetails.customerId) {
                    return { 
                        error: 'Could not find customer for Credit Memo ID: ' + creditMemoId,
                        customerName: null,
                        salesOrders: [],
                        summary: {}
                    };
                }
                
                var customerId = cmDetails.customerId;
                var customerName = cmDetails.customerName;
                
                log.debug('SO Totals Query', 'Customer ID: ' + customerId + ', Name: ' + customerName);
                
                // Step 2: Query all sales orders for this customer using SuiteQL (identical to AI tool query)
                var soQuerySQL = 
                    "SELECT " +
                    "    so.id, " +
                    "    so.tranid, " +
                    "    so.trandate, " +
                    "    cust.altname as customer_name, " +
                    "    so.foreigntotal as so_total_amount, " +
                    "    COALESCE(inv_summary.total_billed, 0) as inv_total_billed, " +
                    "    so.foreigntotal - COALESCE(inv_summary.total_billed, 0) as so_total_unbilled, " +
                    "    (SELECT ABS(SUM(CASE WHEN taxline = 'F' THEN netamount ELSE 0 END)) " +
                    "     FROM transactionline " +
                    "     WHERE transaction = so.id AND mainline = 'F') as so_nontax_amount, " +
                    "    COALESCE(inv_summary.nontax_billed, 0) as inv_nontax_billed, " +
                    "    (SELECT ABS(SUM(CASE WHEN taxline = 'T' THEN netamount ELSE 0 END)) " +
                    "     FROM transactionline " +
                    "     WHERE transaction = so.id AND mainline = 'F') as so_tax_amount, " +
                    "    COALESCE(inv_summary.tax_billed, 0) as inv_tax_billed, " +
                    "    so.status as so_status_id, " +
                    "    BUILTIN.DF(so.status) as so_status_text, " +
                    "    so.createddate as so_created_date, " +
                    "    CASE " +
                    "        WHEN emp.firstname IS NOT NULL AND emp.lastname IS NOT NULL " +
                    "        THEN emp.firstname || ' ' || emp.lastname " +
                    "        WHEN emp.firstname IS NOT NULL " +
                    "        THEN emp.firstname " +
                    "        WHEN emp.lastname IS NOT NULL " +
                    "        THEN emp.lastname " +
                    "        ELSE BUILTIN.DF(so.createdby) " +
                    "    END as so_created_by " +
                    "FROM transaction so " +
                    "INNER JOIN customer cust ON so.entity = cust.id " +
                    "LEFT JOIN employee emp ON so.createdby = emp.id " +
                    "LEFT JOIN ( " +
                    "    SELECT " +
                    "        inv_data.createdfrom, " +
                    "        SUM(inv_data.inv_total) as total_billed, " +
                    "        SUM(inv_data.inv_nontax) as nontax_billed, " +
                    "        SUM(inv_data.inv_tax) as tax_billed " +
                    "    FROM ( " +
                    "        SELECT DISTINCT " +
                    "            tl.createdfrom, " +
                    "            inv.id, " +
                    "            inv.foreigntotal as inv_total, " +
                    "            (SELECT ABS(SUM(CASE WHEN taxline = 'F' THEN netamount ELSE 0 END)) " +
                    "             FROM transactionline " +
                    "             WHERE transaction = inv.id AND mainline = 'F') as inv_nontax, " +
                    "            (SELECT ABS(SUM(CASE WHEN taxline = 'T' THEN netamount ELSE 0 END)) " +
                    "             FROM transactionline " +
                    "             WHERE transaction = inv.id AND mainline = 'F') as inv_tax " +
                    "        FROM transactionline tl " +
                    "        INNER JOIN transaction inv ON tl.transaction = inv.id " +
                    "        WHERE inv.type = 'CustInvc' " +
                    "        GROUP BY tl.createdfrom, inv.id, inv.foreigntotal " +
                    "    ) inv_data " +
                    "    GROUP BY inv_data.createdfrom " +
                    ") inv_summary ON so.id = inv_summary.createdfrom " +
                    "WHERE so.entity = ? " +
                    "  AND so.type = 'SalesOrd' " +
                    "ORDER BY so.trandate ASC";
                
                var soResults = query.runSuiteQL({
                    query: soQuerySQL,
                    params: [customerId]
                });
                
                var soResultsArray = soResults.asMappedResults();
                
                log.debug('SO Query Results', 'Found ' + soResultsArray.length + ' sales orders');
                
                var salesOrders = [];
                var taxDiscrepancies = [];
                var billingVariances = [];
                
                // Summary totals
                var totalSOs = 0;
                var totalSOAmount = 0;
                var totalSOTax = 0;
                var totalInvBilled = 0;
                var totalInvTax = 0;
                
                for (var i = 0; i < soResultsArray.length; i++) {
                    var row = soResultsArray[i];
                    var soId = row.id;
                    var tranid = row.tranid;
                    var trandate = row.trandate;
                    var createdDate = row.so_created_date;
                    var status = row.so_status_id;
                    var statusText = row.so_status_text;
                    var employeeName = row.so_created_by || '-';
                    var soNontaxAmount = parseFloat(row.so_nontax_amount || 0);
                    var soTaxAmount = parseFloat(row.so_tax_amount || 0);
                    var soTotalAmount = parseFloat(row.so_total_amount || 0);
                    var invNontaxBilled = parseFloat(row.inv_nontax_billed || 0);
                    var invTaxBilled = parseFloat(row.inv_tax_billed || 0);
                    var invTotalBilled = parseFloat(row.inv_total_billed || 0);
                    
                    // Calculate tax percentages
                    var soTaxPct = soNontaxAmount > 0 ? Math.round((soTaxAmount / soNontaxAmount) * 100) : 0;
                    var invTaxPct = invNontaxBilled > 0 ? Math.round((invTaxBilled / invNontaxBilled) * 100) : 0;
                    
                    // Tax discrepancy check (ONLY for fully billed orders - Status G)
                    var hasTaxDiscrepancy = false;
                    if (status === 'G' && invTotalBilled > 0.01) {
                        var taxDiff = Math.abs(soTaxAmount - invTaxBilled);
                        if (taxDiff > 0.01) {
                            hasTaxDiscrepancy = true;
                            taxDiscrepancies.push({
                                soId: soId,
                                tranid: tranid,
                                soTax: soTaxAmount,
                                invTax: invTaxBilled,
                                difference: soTaxAmount - invTaxBilled,
                                soTaxPct: soTaxPct,
                                invTaxPct: invTaxPct
                            });
                        }
                    }
                    
                    // Billing variance check (only status G = Billed)
                    var hasBillingVariance = false;
                    if (status === 'G') {
                        var amountDiff = Math.abs(soTotalAmount - invTotalBilled);
                        if (amountDiff > 0.01) {
                            hasBillingVariance = true;
                            billingVariances.push({
                                soId: soId,
                                tranid: tranid,
                                status: status,
                                statusText: statusText,
                                soTotal: soTotalAmount,
                                invBilled: invTotalBilled,
                                difference: soTotalAmount - invTotalBilled
                            });
                        }
                    }
                    
                    salesOrders.push({
                        soTranId: tranid,
                        soDate: trandate,
                        soCustomerName: customerName,
                        soTotalAmount: soTotalAmount,
                        invTotalBilled: invTotalBilled,
                        soTotalUnbilled: soTotalAmount - invTotalBilled,
                        soNontaxAmount: soNontaxAmount,
                        invNontaxBilled: invNontaxBilled,
                        soTaxAmount: soTaxAmount,
                        invTaxBilled: invTaxBilled,
                        soTaxPct: soTaxPct,
                        invTaxPct: invTaxPct,
                        soStatusId: status,
                        soStatusText: statusText,
                        soCreatedDate: createdDate,
                        soCreatedBy: employeeName,
                        soId: soId,
                        hasTaxDiscrepancy: hasTaxDiscrepancy,
                        hasBillingVariance: hasBillingVariance
                    });
                    
                    totalSOs++;
                    totalSOAmount += soNontaxAmount;
                    totalSOTax += soTaxAmount;
                    totalInvBilled += invNontaxBilled;
                    totalInvTax += invTaxBilled;
                }
                
                return {
                    customerName: customerName,
                    customerId: customerId,
                    salesOrders: salesOrders,
                    summary: {
                        totalSalesOrders: totalSOs,
                        totalOrderAmount: totalSOAmount + totalSOTax,
                        totalBilledAmount: totalInvBilled + totalInvTax,
                        totalUnbilledAmount: (totalSOAmount + totalSOTax) - (totalInvBilled + totalInvTax),
                        totalOrderNonTaxAmount: totalSOAmount,
                        totalOrderTaxAmount: totalSOTax,
                        totalBilledNonTaxAmount: totalInvBilled,
                        totalBilledTaxAmount: totalInvTax,
                        taxDiscrepancyCount: taxDiscrepancies.length,
                        taxDiscrepancies: taxDiscrepancies,
                        billingVarianceCount: billingVariances.length,
                        billingVariances: billingVariances
                    }
                };
                
            } catch (e) {
                log.error('Error in getCustomerSalesOrdersSummary', e.toString());
                return { 
                    error: e.toString(),
                    customerName: null,
                    salesOrders: [],
                    summary: {}
                };
            }
        }
        
        /**
         * Helper: Map status text to status code
         * @param {string} statusText - Status text (e.g., "Billed", "Pending Fulfillment")
         * @returns {string} Status code (e.g., "G", "B")
         */
        function mapStatusTextToCode(statusText) {
            if (!statusText) return '';
            
            var statusMap = {
                'Pending Approval': 'A',
                'Pending Fulfillment': 'B',
                'Cancelled': 'C',
                'Partially Fulfilled': 'D',
                'Pending Billing/Partially Fulfilled': 'E',
                'Pending Billing': 'F',
                'Billed': 'G',
                'Closed': 'H'
            };
            
            return statusMap[statusText] || '';
        }

        return {
            onRequest: onRequest
        };
    });
