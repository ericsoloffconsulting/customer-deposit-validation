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

            // SO to Invoice Comparison Modal
            html += '<div id="soInvComparisonModal" class="comparison-modal">';
            html += '<div class="comparison-modal-content">';
            html += '<div class="comparison-modal-header">';
            html += '<span class="comparison-modal-title">SO to Invoice Line Item Comparison</span>';
            html += '<span class="comparison-modal-close" onclick="hideComparisonModal()">&times;</span>';
            html += '</div>';
            html += '<div id="comparisonModalBody" class="comparison-modal-body">';
            html += '<div class="comparison-loading">Loading comparison data...</div>';
            html += '</div>';
            html += '</div>';
            html += '</div>';
            html += '<div id="comparisonModalOverlay" class="comparison-modal-overlay" onclick="hideComparisonModal()"></div>';

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
                creditMemos, scriptUrl);

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
         * @returns {string} HTML for data section
         */
        function buildCreditMemoDataSection(sectionId, title, description, data, scriptUrl) {
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
                html += buildCreditMemoTable(data, scriptUrl, sectionId);
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
         * @returns {string} HTML table
         */
        function buildCreditMemoTable(creditMemos, scriptUrl, sectionId) {
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

                // SO to Invoice Comparison action button
                html += '<td class="action-btn-cell"><button type="button" class="so-inv-compare-btn" onclick="showSOInvoiceComparison(' + cm.cmId + ')" title="Compare SO vs Invoice Line Items">SO↔INV</button></td>';

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

                /* SO to Invoice Comparison Button */
                '.so-inv-compare-btn { padding: 4px 8px; font-size: 11px; font-weight: 600; color: #fff; background: linear-gradient(135deg, #1976d2, #1565c0); border: none; border-radius: 4px; cursor: pointer; white-space: nowrap; transition: all 0.2s; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }' +
                '.so-inv-compare-btn:hover { background: linear-gradient(135deg, #1565c0, #0d47a1); transform: translateY(-1px); box-shadow: 0 2px 5px rgba(0,0,0,0.25); }' +
                '.action-btn-cell { text-align: center !important; padding: 4px !important; }' +

                /* SO to Invoice Comparison Modal */
                '.comparison-modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 99998; }' +
                '.comparison-modal-overlay.visible { display: block; }' +
                '.comparison-modal { display: none; position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; border-radius: 10px; box-shadow: 0 8px 32px rgba(0,0,0,0.3); z-index: 99999; width: 95%; max-width: 1400px; max-height: 90vh; overflow: hidden; }' +
                '.comparison-modal.visible { display: block; }' +
                '.comparison-modal-content { display: flex; flex-direction: column; height: 100%; max-height: 90vh; }' +
                '.comparison-modal-header { display: flex; justify-content: space-between; align-items: center; padding: 16px 20px; background: linear-gradient(135deg, #1976d2, #1565c0); color: white; border-radius: 10px 10px 0 0; }' +
                '.comparison-modal-title { font-size: 18px; font-weight: 700; }' +
                '.comparison-modal-close { cursor: pointer; font-size: 28px; line-height: 1; opacity: 0.8; padding: 0 8px; }' +
                '.comparison-modal-close:hover { opacity: 1; }' +
                '.comparison-modal-body { padding: 20px; overflow-y: auto; flex: 1; }' +
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
                '.comparison-tran-link { display: flex; align-items: center; gap: 8px; padding: 10px 16px; background: #fff; border: 1px solid #dee2e6; border-radius: 6px; text-decoration: none; color: #333; transition: all 0.2s; min-width: 140px; }' +
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
                '.comparison-table .match { color: #4caf50; font-weight: 600; }' +
                '.comparison-table .mismatch { color: #d32f2f; font-weight: 600; }' +
                '.comparison-table .not-invoiced { color: #daa520; font-weight: 600; }' +
                '.comparison-table .not-on-so { color: #d32f2f; font-weight: 600; }' +
                '.comparison-table .discount { color: #9c27b0; font-weight: 600; }' +
                '.comparison-table .variance-positive { color: #4caf50; }' +
                '.comparison-table .variance-negative { color: #d32f2f; }' +

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
                '/* SO TO INVOICE COMPARISON FUNCTIONS         */' +
                '/* =========================================== */' +
                '' +
                '/* Show SO to Invoice Comparison Modal */' +
                'function showSOInvoiceComparison(creditMemoId) {' +
                '    var modal = document.getElementById("soInvComparisonModal");' +
                '    var overlay = document.getElementById("comparisonModalOverlay");' +
                '    var body = document.getElementById("comparisonModalBody");' +
                '    ' +
                '    body.innerHTML = "<div class=\\"comparison-loading\\">Loading comparison data...</div>";' +
                '    modal.classList.add("visible");' +
                '    overlay.classList.add("visible");' +
                '    ' +
                '    fetch("' + scriptUrl + '", {' +
                '        method: "POST",' +
                '        headers: { "Content-Type": "application/json" },' +
                '        body: JSON.stringify({ action: "soInvoiceComparison", creditMemoId: creditMemoId })' +
                '    })' +
                '    .then(function(response) { return response.json(); })' +
                '    .then(function(result) {' +
                '        if (result.success && result.data) {' +
                '            renderComparisonResult(result.data);' +
                '        } else {' +
                '            body.innerHTML = "<div class=\\"error-msg\\">Error: " + (result.data && result.data.error ? result.data.error : "Unknown error") + "</div>";' +
                '        }' +
                '    })' +
                '    .catch(function(err) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\">Error loading comparison: " + err.message + "</div>";' +
                '    });' +
                '}' +
                '' +
                '/* Hide Comparison Modal */' +
                'function hideComparisonModal() {' +
                '    document.getElementById("soInvComparisonModal").classList.remove("visible");' +
                '    document.getElementById("comparisonModalOverlay").classList.remove("visible");' +
                '}' +
                '' +
                '/* Render Comparison Result */' +
                'function renderComparisonResult(data) {' +
                '    var body = document.getElementById("comparisonModalBody");' +
                '    ' +
                '    if (data.error) {' +
                '        body.innerHTML = "<div class=\\"error-msg\\"><strong>Error:</strong> " + data.error + "</div>";' +
                '        if (data.memo) {' +
                '            body.innerHTML += "<div style=\\"margin-top:10px;padding:10px;background:#f5f5f5;border-radius:4px;\\"><strong>Memo field:</strong> " + data.memo + "</div>";' +
                '        }' +
                '        return;' +
                '    }' +
                '    ' +
                '    /* Store comparison data globally for AI analysis */' +
                '    window.currentComparisonData = {' +
                '        mismatchVariance: data.totals ? data.totals.mismatchVariance : 0,' +
                '        unbilledVariance: data.totals ? data.totals.unbilledVariance : 0,' +
                '        totalVariance: data.totals ? data.totals.totalVariance : 0' +
                '    };' +
                '    ' +
                '    var html = "";' +
                '    ' +
                '    /* Transaction Links with Amounts */' +
                '    html += "<div class=\\"comparison-transactions\\">";' +
                '    ' +
                '    /* Show ALL related Credit Memos collectively */' +
                '    var allCMs = data.allRelatedCreditMemos || (data.creditMemo ? [data.creditMemo] : []);' +
                '    var totalCMAmount = data.totalCMAmount || (data.creditMemo ? data.creditMemo.amount : 0);' +
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
                '    ' +
                '    if (data.customerDeposit) {' +
                '        html += "<a href=\\"/app/accounting/transactions/custdep.nl?id=" + data.customerDeposit.id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link\\">";' +
                '        html += "<div><div class=\\"comparison-tran-label\\">Customer Deposit</div><div class=\\"comparison-tran-value\\">" + data.customerDeposit.tranid + "</div><div class=\\"comparison-tran-amount\\">" + formatCompCurrency(Math.abs(data.customerDeposit.amount)) + "</div></div>";' +
                '        html += "</a>";' +
                '    }' +
                '    if (data.salesOrder) {' +
                '        html += "<a href=\\"/app/accounting/transactions/salesord.nl?id=" + data.salesOrder.id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link\\">";' +
                '        html += "<div><div class=\\"comparison-tran-label\\">Sales Order (" + data.salesOrder.statusText + ")</div><div class=\\"comparison-tran-value\\">" + data.salesOrder.tranid + "</div><div class=\\"comparison-tran-amount\\">" + formatCompCurrency(Math.abs(data.salesOrder.amount)) + "</div></div>";' +
                '        html += "</a>";' +
                '    }' +
                '    if (data.invoices && data.invoices.length > 0) {' +
                '        for (var i = 0; i < data.invoices.length; i++) {' +
                '            var inv = data.invoices[i];' +
                '            html += "<a href=\\"/app/accounting/transactions/custinvc.nl?id=" + inv.id + "\\" target=\\"_blank\\" class=\\"comparison-tran-link\\">";' +
                '            html += "<div><div class=\\"comparison-tran-label\\">Invoice " + (i+1) + "</div><div class=\\"comparison-tran-value\\">" + inv.tranid + "</div><div class=\\"comparison-tran-amount\\">" + formatCompCurrency(Math.abs(inv.amount)) + "</div></div>";' +
                '            html += "</a>";' +
                '        }' +
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
                '    /* AI Analysis Button - Only show if mismatch does NOT fully explain CM */' +
                '    var showAIButton = !shouldHideAIButton(mismatchVar, cmTotalAmt);' +
                '    if (showAIButton) {' +
                '        html += "<div style=\\"margin-top:20px;padding:15px;background:#e3f2fd;border:1px solid #1976d2;border-radius:6px;\\">";' +
                '        html += "<button type=\\"button\\" id=\\"aiAnalysisBtn\\" class=\\"ai-analysis-btn\\" onclick=\\"runAIAnalysis(" + data.creditMemo.id + ")\\" style=\\"padding:10px 20px;background:#1976d2;color:white;border:none;border-radius:4px;cursor:pointer;font-size:14px;font-weight:600;\\">🤖 Analyze Transaction Lifecycle with AI</button>";' +
                '        html += "<div style=\\"margin-top:8px;font-size:12px;color:#555;\\">";' +
                '        html += "The AI will analyze the sales order&#39;s transaction lifecycle to determine if changes to the order explain this overpayment credit memo.";' +
                '        html += "</div>";' +
                '        html += "</div>";' +
                '        ' +
                '        /* AI Results Container */' +
                '        html += "<div id=\\"aiResultsContainer\\" style=\\"margin-top:20px;\\">";' +
                '        html += "<div id=\\"aiResultsContent\\" style=\\"padding:20px;background:#f5f5f5;border:1px solid #ccc;border-radius:6px;\\">";' +
                '        html += "</div>";' +
                '        html += "</div>";' +
                '    }' +
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
                '        html += "<td>" + soTranid + "</td>";' +
                '        html += "<td class=\\"amount\\">" + soQty + "</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencySigned(soRate) + "</td>";' +
                '        html += "<td class=\\"amount\\">" + formatCompCurrencySigned(soAmt) + "</td>";' +
                '        html += "<td>" + (row.invoiceNumbers || "-") + "</td>";' +
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
                '    body.innerHTML = html;' +
                '    ' +
                '    if (showAIButton && data.creditMemo && data.creditMemo.id) {' +
                '        setTimeout(function() {' +
                '            loadExistingAIAnalysis(data.creditMemo.id);' +
                '        }, 100);' +
                '    }' +
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
                '                    aiDecision: data.aiDecision,' +
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
                '    var resultsContent = document.getElementById("aiResultsContent");' +
                '    if (!resultsContent) return;' +
                '    ' +
                '    if (data.error) {' +
                '        resultsContent.innerHTML = "<div style=\\"color:#d32f2f;padding:15px;\\">" + data.error + "</div>";' +
                '        return;' +
                '    }' +
                '    ' +
                '    var html = "";' +
                '    var decision = data.aiDecision || "UNKNOWN";' +
                '    var decisionClass = decision === "CONFIRMED" ? "success" : (decision === "NOT CONFIRMED" ? "error" : "warning");' +
                '    var decisionBg = decision === "CONFIRMED" ? "#e8f5e9" : (decision === "NOT CONFIRMED" ? "#ffebee" : "#fff3e0");' +
                '    var decisionBorder = decision === "CONFIRMED" ? "#4caf50" : (decision === "NOT CONFIRMED" ? "#d32f2f" : "#f57c00");' +
                '    var decisionIcon = decision === "CONFIRMED" ? "✓" : (decision === "NOT CONFIRMED" ? "✗" : "?");' +
                '    ' +
                '    html += "<div style=\\"padding:15px;background:" + decisionBg + ";border:2px solid " + decisionBorder + ";border-radius:6px;margin-bottom:15px;\\">" ;' +
                '    html += "<div style=\\"font-size:16px;font-weight:700;color:#333;\\">" + decisionIcon + " AI DECISION: " + decision + "</div>";' +
                '    if (decision === "CONFIRMED") {' +
                '        html += "<div style=\\"margin-top:5px;font-size:13px;color:#555;\\">Sales Order changes explain the overpayment credit memo.</div>";' +
                '    } else if (decision === "NOT CONFIRMED") {' +
                '        html += "<div style=\\"margin-top:5px;font-size:13px;color:#555;\\">Sales Order changes do NOT explain the overpayment. Investigate other sources.</div>";' +
                '    }' +
                '    html += "</div>";' +
                '    ' +
                '    html += "<div style=\\"background:white;padding:20px;border:1px solid #ddd;border-radius:6px;\\">" ;' +
                '    html += "<h3 style=\\"margin:0 0 15px 0;font-size:16px;color:#1976d2;\\">🤖 AI Forensic Analysis</h3>";' +
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
                '        html += (data.systemNotesCount !== undefined ? " | " : "") + "<span style=\\"color:#1976d2;\\">Record ID: " + data.savedRecordId + "</span>";' +
                '    }' +
                '    if (data.createdDate) {' +
                '        html += " | <span style=\\"color:#777;\\">Created: " + data.createdDate + "</span>";' +
                '    }' +
                '    html += "</div>";' +
                '    html += "</div>";' +
                '    ' +
                '    resultsContent.innerHTML = html;' +
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
                var sql = `
                    SELECT 
                        tl.id,
                        tl.linesequencenumber,
                        i.itemid,
                        i.displayname,
                        tl.quantity,
                        tl.rate,
                        tl.amount
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
         * @param {string} depositDate - Customer Deposit date (YYYY-MM-DD)
         * @param {string} overpaymentDate - Overpayment recognition date (YYYY-MM-DD)
         * @returns {Array} System Notes records
         */
        function querySystemNotesForSO(soInternalId, depositDate, overpaymentDate) {
            try {
                // Convert M/D/YYYY to Date object for comparison
                function parseNetSuiteDate(dateStr) {
                    if (!dateStr) return null;
                    var parts = dateStr.split('/');
                    if (parts.length === 3) {
                        // parts[0] = month, parts[1] = day, parts[2] = year
                        return new Date(parts[2], parts[0] - 1, parts[1]); // year, month (0-indexed), day
                    }
                    return null;
                }

                var depositDateObj = parseNetSuiteDate(depositDate);
                var overpaymentDateObj = parseNetSuiteDate(overpaymentDate);

                if (!depositDateObj || !overpaymentDateObj) {
                    log.error('Invalid Dates', {
                        depositDate: depositDate,
                        overpaymentDate: overpaymentDate
                    });
                    return [];
                }

                log.debug('System Notes Query Params', {
                    soInternalId: soInternalId,
                    depositDate: depositDate,
                    overpaymentDate: overpaymentDate,
                    depositDateObj: depositDateObj.toISOString(),
                    overpaymentDateObj: overpaymentDateObj.toISOString()
                });

                // Query ALL system notes for this SO (no date filter in SQL - we'll filter in JS)
                var sql = `
                    SELECT 
                        sn.id,
                        sn.date,
                        sn.field,
                        sn.oldvalue,
                        sn.newvalue,
                        sn.name,
                        sn.lineid,
                        sn.context,
                        sn.recordtypeid,
                        e.firstname || ' ' || e.lastname AS user_name
                    FROM systemNote sn
                    LEFT JOIN employee e ON sn.name = e.id
                    WHERE sn.recordid = ` + soInternalId + `
                    ORDER BY sn.date, sn.id
                `;

                var results = query.runSuiteQL({ query: sql }).asMappedResults();
                
                log.debug('System Notes Query - Total Found', {
                    soId: soInternalId,
                    totalFound: results.length
                });

                // Filter by date in JavaScript
                var filteredResults = [];
                for (var i = 0; i < results.length; i++) {
                    var note = results[i];
                    var noteDateObj = parseNetSuiteDate(note.date);
                    
                    if (noteDateObj && noteDateObj >= depositDateObj && noteDateObj <= overpaymentDateObj) {
                        filteredResults.push(note);
                    }
                }
                
                log.debug('System Notes Query - Date Filtered', {
                    beforeFilter: results.length,
                    afterFilter: filteredResults.length,
                    dateRange: depositDate + ' to ' + overpaymentDate,
                    sampleFields: filteredResults.length > 0 ? filteredResults.slice(0, 10).map(function(r) { 
                        return { date: r.date, field: r.field }; 
                    }) : []
                });
                
                return filteredResults;

            } catch (e) {
                log.error('Error querying system notes', { error: e.message, soId: soInternalId });
                return [];
            }
        }

        /**
         * Organize and filter System Notes by tier (financial, lifecycle, scheduling, business, people)
         * Excludes operational noise like fulfillment status, allocations, picks, packs
         * @param {Array} rawSystemNotes - Raw system notes from query
         * @returns {Object} Organized system notes by tier with meaningful changes only
         */
        function organizeSystemNotesByTier(rawSystemNotes) {
            var organized = {
                financial: [],
                lifecycle: [],
                scheduling: [],
                business: [],
                people: [],
                excluded: []
            };

            // Define tier classifications
            var tiers = {
                financial: ['TRANLINE.MAMOUNT', 'TRANLINE.RUNITPRICE', 'TRANLINE.RQTY', 'TRANDOC.MAMOUNTMAIN', 'TRANLINE.FXAMOUNT'],
                lifecycle: ['TRANDOC.KSTATUS', 'TRANDOC.KREVENUESTATUS'],
                scheduling: ['TRANDOC.DSHIP', 'TRANLINE.DSHIP', 'TRANLINE.DREQUESTEDDATE', 'TRANDOC.DEXPECTEDSHIPDATE'],
                business: ['TRANDOC.ENTITY', 'TRANDOC.KTERMS', 'TRANDOC.SADDR', 'TRANDOC.SSHIPADDR'],
                // Exclude operational noise
                exclude: ['TRANLINE.RQTYSHIPRECV', 'TRANLINE.RQTYPICKED', 'TRANLINE.RQTYPACKED', 'TRANLINE.RCOMMITTED', 
                         'TRANLINE.RALLOCATED', 'TRANLINE.KLOCATION', 'TRANDOC.KSHIPPINGSTATUS', 'TRANDOC.KCLOSED',
                         'TRANLINE.KCLOSED', 'TRANLINE.KCOMMITTINGSTATUS', 'TRANLINE.KFULFILLMENTSTATUS']
            };

            for (var i = 0; i < rawSystemNotes.length; i++) {
                var note = rawSystemNotes[i];
                var field = note.field;
                var placed = false;

                // Check if excluded
                for (var e = 0; e < tiers.exclude.length; e++) {
                    if (field === tiers.exclude[e]) {
                        organized.excluded.push(note);
                        placed = true;
                        break;
                    }
                }
                if (placed) continue;

                // Check financial tier
                for (var f = 0; f < tiers.financial.length; f++) {
                    if (field === tiers.financial[f]) {
                        organized.financial.push(note);
                        placed = true;
                        break;
                    }
                }
                if (placed) continue;

                // Check lifecycle tier
                for (var l = 0; l < tiers.lifecycle.length; l++) {
                    if (field === tiers.lifecycle[l]) {
                        organized.lifecycle.push(note);
                        placed = true;
                        break;
                    }
                }
                if (placed) continue;

                // Check scheduling tier
                for (var s = 0; s < tiers.scheduling.length; s++) {
                    if (field === tiers.scheduling[s]) {
                        organized.scheduling.push(note);
                        placed = true;
                        break;
                    }
                }
                if (placed) continue;

                // Check business tier
                for (var b = 0; b < tiers.business.length; b++) {
                    if (field === tiers.business[b]) {
                        organized.business.push(note);
                        placed = true;
                        break;
                    }
                }
                if (placed) continue;

                // If not classified, include in people/other
                organized.people.push(note);
            }

            log.debug('System Notes Organized', 'Financial: ' + organized.financial.length + ', Lifecycle: ' + 
                organized.lifecycle.length + ', Scheduling: ' + organized.scheduling.length + ', Business: ' + 
                organized.business.length + ', People: ' + organized.people.length + ', Excluded: ' + organized.excluded.length);

            return organized;
        }

        /**
         * Build system prompt for Claude AI forensic analyst
         * @returns {string} System prompt
         */
        function buildTransactionLifecycleSystemPrompt() {
            return 'You are a forensic accounting analyst examining a NetSuite sales order transaction lifecycle to determine if changes to the order explain an overpayment credit memo.\n\n' +
                'BUSINESS CONTEXT:\n' +
                '- Customer: Kitchen Works (custom kitchen cabinets, complex multi-SO projects)\n' +
                '- Typical Issue: Customer deposits paid upfront, then Sales Order changes occur (price adjustments, quantity changes, cancellations)\n' +
                '- When SO total decreases after deposit paid, system recognizes overpayment and creates Credit Memo\n' +
                '- Migration Date: 4/30/2024 - Data before this date may contain migration artifacts and should be considered less reliable\n\n' +
                'YOUR TASK:\n' +
                'Analyze the provided Sales Order System Notes to determine if transaction changes explain the Credit Memo overpayment.\n\n' +
                'ANALYSIS APPROACH:\n' +
                '1. Focus on FINANCIAL changes: line amounts (TRANLINE.MAMOUNT), unit prices (TRANLINE.RUNITPRICE), quantities (TRANLINE.RQTY), header totals (TRANDOC.MAMOUNTMAIN)\n' +
                '2. Consider LIFECYCLE changes: order status (TRANDOC.KSTATUS), revenue recognition status (TRANDOC.KREVENUESTATUS)\n' +
                '3. Note SCHEDULING changes if relevant: ship dates (DSHIP), requested dates (DREQUESTEDDATE)\n' +
                '4. Evaluate BUSINESS changes: customer (ENTITY), payment terms (KTERMS), shipping address (SADDR)\n' +
                '5. Track WHO made changes and WHEN - cite specific employee names and dates\n' +
                '6. Calculate net financial impact: Did SO total decrease by approximately the CM amount?\n\n' +
                'CRITICAL REQUIREMENTS:\n' +
                '- ALWAYS start with a narrative paragraph explaining your findings in plain English\n' +
                '- For line-level changes (TRANLINE.MAMOUNT, TRANLINE.RUNITPRICE, TRANLINE.RQTY), describe WHAT changed even if item names are not available\n' +
                '- Track EVERY financial change chronologically with: Date, Person, Field, Old Value → New Value, Net Impact\n' +
                '- Calculate the EXACT net change from all TRANDOC.MAMOUNTMAIN entries\n' +
                '- Compare net change to CM amount - if difference > $100, explain the discrepancy\n' +
                '- Be FACTUAL and DETAILED - cite specific dollar amounts, dates, and names\n' +
                '- Do NOT dismiss large discrepancies as "small" - quantify and explain\n\n' +
                'OUTPUT FORMAT:\n' +
                'Start with: EXECUTIVE SUMMARY - A 2-3 sentence narrative explaining what you found and whether SO changes explain the CM.\n\n' +
                'Then provide these sections:\n' +
                '1. DECISION: ---DECISION--- CONFIRMED ---END--- or ---DECISION--- NOT CONFIRMED ---END---\n' +
                '   STRICT CRITERIA:\n' +
                '   - CONFIRMED = Net SO decrease matches CM amount within $0.01 tolerance (rounding only)\n' +
                '   - NOT CONFIRMED = Variance > $0.01, or no decrease found, or any unexplained discrepancy\n' +
                '   - The SO changes must EXACTLY explain the CM amount - even $1+ variance means NOT CONFIRMED\n' +
                '   - If variance exists, the data may still be helpful but the answer is NOT CONFIRMED\n\n' +
                '2. FINANCIAL RECONCILIATION:\n' +
                '   - Initial SO Total: $X (date, by whom)\n' +
                '   - Final SO Total: $Y (date, by whom)\n' +
                '   - Net Change: $(X-Y) [DECREASE] or $(Y-X) [INCREASE]\n' +
                '   - Credit Memo Amount: $Z\n' +
                '   - Variance: $(|Net Change - CM|)\n' +
                '   - VARIANCE EXPLANATION: If variance > $0.01, explain the discrepancy. Provide theories: multiple invoices? partial billing? contract adjustments? other transactions?\n\n' +
                '3. LINE-LEVEL CHANGES (if any TRANLINE changes):\n' +
                '   IMPORTANT: Use the lineid field to identify which item changed. Match lineid to Line ID in current items list.\n' +
                '   Format: Date, User, Line ID X (Item Name), Field, Old → New, Impact\n' +
                '   Example: "7/12/2024: Michael Strange changed Line ID 3 (Installation Labor) price $295.00 → $395.80 (+$100.80)"\n' +
                '   \n' +
                '   CRITICAL PATTERN TO DETECT - LINE DELETIONS:\n' +
                '   If TRANDOC.MAMOUNTMAIN decreased significantly but you see NO or MINIMAL TRANLINE changes:\n' +
                '   - This indicates LINE ITEMS WERE DELETED from the Sales Order\n' +
                '   - NetSuite does NOT record individual line deletions in system notes\n' +
                '   - State: "LINE DELETION DETECTED: SO total decreased by $X but only $Y in line changes found. Missing $Z suggests items were removed."\n' +
                '   - Mark this as a DATA QUALITY LIMITATION in your analysis\n\n' +
                '4. CHRONOLOGICAL TIMELINE:\n' +
                '   CRITICAL: ALWAYS include WHO made each change. Every entry must show the person responsible.\n' +
                '   Format: Date: Action by [User Name]\n' +
                '   Example:\n' +
                '   - 5/15/2024: Sales Order Created by Michael Strange\n' +
                '   - 6/21/2024: Customer Deposit ($7,114.36) by Automated System\n' +
                '   - 7/12/2024: SO Total Reduced to $25,078.64 by Michael Strange\n' +
                '   - 10/04/2024: Shipping Dates Adjusted by Lizabeth Lopez\n' +
                '   - 12/15/2024: Order Marked "Billed" by Lizabeth Lopez\n' +
                '   - 12/13/2025: Credit Memo Generated ($704.25) by System\n\n' +
                '5. KEY CONTACTS:\n' +
                '   - Name: Role/Actions (summarize their involvement)\n\n' +
                '6. DATA QUALITY NOTES:\n' +
                '   - Migration date issues, missing data, anomalies\n\n' +
                'Remember: Start with the executive summary narrative, then provide structured detail.';
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
            var currentLineItems = context.currentLineItems || [];
            var comparisonData = context.comparisonData || {};

            var prompt = 'TRANSACTION DETAILS:\n';
            prompt += 'Credit Memo: ' + cm.tranid + ' (' + cm.trandate + ') Amount: $' + cm.amount.toFixed(2) + '\n';
            prompt += 'Customer Deposit: ' + cd.tranid + ' (' + cd.trandate + ') Amount: $' + cd.amount.toFixed(2) + '\n';
            prompt += 'Sales Order: ' + so.tranid + ' (' + so.trandate + ') Current Total: $' + so.amount.toFixed(2) + ' Status: ' + so.statusText + '\n';
            prompt += 'Deposit Date: ' + cd.trandate + '\n';
            prompt += 'Overpayment Recognition Date: ' + (cm.trandate || 'Unknown') + '\n\n';

            // Add current line items for context
            if (currentLineItems.length > 0) {
                prompt += 'CURRENT SALES ORDER LINE ITEMS:\n';
                prompt += 'Use this to identify WHICH ITEM changed when you see a system note with a lineid:\n';
                for (var i = 0; i < currentLineItems.length; i++) {
                    var line = currentLineItems[i];
                    var itemName = line.displayname || line.itemid || 'Unknown Item';
                    prompt += 'Line ID ' + line.id + ': Item ' + itemName;
                    prompt += ' (Seq: ' + line.linesequencenumber + ')';
                    prompt += ' | Qty: ' + (line.quantity || 0) + ' | Rate: $' + (line.rate || 0);
                    prompt += ' | Amount: $' + (line.amount || 0) + '\n';
                }
                prompt += '\nCRITICAL MATCHING RULE:\n';
                prompt += 'When a system note has field "TRANLINE.*" and includes a "lineid" value, match that lineid to the Line ID above.\n';
                prompt += 'Example: System note with lineid="3" means it changed Line ID 3 above.\n';
                prompt += 'ALWAYS reference the item name/number in your LINE-LEVEL CHANGES section.\n';
                prompt += 'Format: "Line ID 3 (Item 00229)" or "Line ID 5 (Item Kitchen Cabinet Door)"\n\n';
            }

            // Add comparison tool findings context
            if (comparisonData.mismatchVariance !== undefined) {
                var mismatchAbs = Math.abs(comparisonData.mismatchVariance || 0);
                var cmAmount = cm.amount;
                var difference = Math.abs(mismatchAbs - cmAmount);
                
                prompt += 'SO-TO-INVOICE COMPARISON TOOL FINDINGS:\n';
                prompt += 'Mismatch Variance (SO vs Invoice line item errors): $' + mismatchAbs.toFixed(2) + '\n';
                prompt += 'Credit Memo Amount: $' + cmAmount.toFixed(2) + '\n';
                
                if (difference < 0.10) {
                    // Perfect match - shouldn't see AI button, but if here somehow...
                    prompt += 'Analysis: Mismatch EXACTLY explains CM (variance matches within $0.10)\n';
                    prompt += 'AI GUIDANCE: This case is already resolved by comparison tool. Confirm SO changes align with mismatch findings.\n\n';
                } else if (mismatchAbs > 0.01 && mismatchAbs < cmAmount) {
                    // Partial match
                    var unexplained = cmAmount - mismatchAbs;
                    prompt += 'Analysis: PARTIAL MATCH - Mismatch explains $' + mismatchAbs.toFixed(2) + ' but $' + unexplained.toFixed(2) + ' remains unexplained\n';
                    prompt += 'AI GUIDANCE: Focus on explaining the $' + unexplained.toFixed(2) + ' unexplained amount. Look for SO changes NOT captured in line item mismatch.\n\n';
                } else if (mismatchAbs < 0.01) {
                    // No mismatch
                    prompt += 'Analysis: NO MISMATCH found in SO vs Invoice line items ($' + mismatchAbs.toFixed(2) + ')\n';
                    prompt += 'AI GUIDANCE: Analyze full CM amount ($' + cmAmount.toFixed(2) + '). Look for header-level changes, cancellations, or other factors not visible in line items.\n\n';
                } else {
                    // Excess mismatch
                    var excess = mismatchAbs - cmAmount;
                    prompt += 'Analysis: EXCESS MISMATCH - Variance ($' + mismatchAbs.toFixed(2) + ') exceeds CM ($' + cmAmount.toFixed(2) + ') by $' + excess.toFixed(2) + '\n';
                    prompt += 'AI GUIDANCE: Complex scenario. Determine if SO changes explain CM, or if mismatch includes unrelated billing errors.\n\n';
                }
            }

            prompt += 'SALES ORDER SYSTEM NOTES (FILTERED FOR MEANINGFUL CHANGES):\n\n';

            // Financial changes (highest priority)
            if (organizedNotes.financial.length > 0) {
                prompt += '=== FINANCIAL CHANGES (Tier 1 - Critical) ===\n';
                prompt += JSON.stringify(organizedNotes.financial, null, 2) + '\n\n';
            } else {
                prompt += '=== FINANCIAL CHANGES (Tier 1 - Critical) ===\nNone found.\n\n';
            }

            // Lifecycle changes
            if (organizedNotes.lifecycle.length > 0) {
                prompt += '=== LIFECYCLE CHANGES (Tier 2) ===\n';
                prompt += JSON.stringify(organizedNotes.lifecycle, null, 2) + '\n\n';
            } else {
                prompt += '=== LIFECYCLE CHANGES (Tier 2) ===\nNone found.\n\n';
            }

            // Scheduling changes
            if (organizedNotes.scheduling.length > 0) {
                prompt += '=== SCHEDULING CHANGES (Tier 3) ===\n';
                prompt += JSON.stringify(organizedNotes.scheduling, null, 2) + '\n\n';
            }

            // Business changes
            if (organizedNotes.business.length > 0) {
                prompt += '=== BUSINESS CHANGES (Tier 4) ===\n';
                prompt += JSON.stringify(organizedNotes.business, null, 2) + '\n\n';
            }

            // People/other
            if (organizedNotes.people.length > 0) {
                prompt += '=== OTHER CHANGES (Tier 5) ===\n';
                prompt += JSON.stringify(organizedNotes.people, null, 2) + '\n\n';
            }

            prompt += 'TOTAL SYSTEM NOTES ANALYZED: ' + (organizedNotes.financial.length + organizedNotes.lifecycle.length + 
                organizedNotes.scheduling.length + organizedNotes.business.length + organizedNotes.people.length) + '\n';
            prompt += 'EXCLUDED (operational noise): ' + organizedNotes.excluded.length + '\n\n';

            prompt += 'Based on this data, provide your forensic analysis following the output format specified in your system prompt.';

            return prompt;
        }

        /**
         * Parse AI decision from Haiku response
         * @param {string} haikuResponse - Claude Haiku response text
         * @returns {string} "CONFIRMED" or "NOT CONFIRMED" or "UNKNOWN"
         */
        function parseAIDecision(haikuResponse) {
            if (!haikuResponse) return 'UNKNOWN';
            
            var decisionMatch = haikuResponse.match(/---DECISION---\s*(CONFIRMED|NOT CONFIRMED)\s*---END---/i);
            if (decisionMatch && decisionMatch[1]) {
                return decisionMatch[1].toUpperCase().trim();
            }
            
            return 'UNKNOWN';
        }

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

                aiRecord.setValue({
                    fieldId: 'custrecord_ai_decision',
                    value: params.aiDecision
                });

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
                        'custrecord_ai_decision',
                        'custrecord_user_prompt',
                        'custrecord_system_prompt',
                        'created'
                    ]
                });

                var results = aiSearch.run().getRange({ start: 0, end: 1 });

                if (results && results.length > 0) {
                    var result = results[0];
                    log.debug('AI Analysis Found', 'Record ID: ' + result.id);
                    
                    return {
                        found: true,
                        recordId: result.id,
                        haikuResponse: result.getValue('custrecord_haiku_response'),
                        aiDecision: result.getValue('custrecord_ai_decision'),
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

                // Query System Notes for SO
                var rawSystemNotes = querySystemNotesForSO(cdData.soId, cdData.trandate, cmData.trandate);
                
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
                var aiDecision = parseAIDecision(haikuResponse);

                log.debug('AI Analysis Complete', 'Decision: ' + aiDecision);

                // Save to custom record
                var savedRecordId = saveAIAnalysisToRecord({
                    creditMemoId: creditMemoId,
                    haikuResponse: haikuResponse,
                    userPrompt: userPrompt,
                    systemPrompt: systemPrompt,
                    aiDecision: aiDecision
                });

                return {
                    success: true,
                    aiDecision: aiDecision,
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

                // STEP 4: Get SO details
                var soData = getSalesOrderDetails(soId);
                log.debug('SO Data', JSON.stringify(soData));

                // STEP 5: Get SO line items
                var soLineItems = getSOLineItems(soId);
                var soTaxTotal = getSOTaxTotal(soId);
                log.debug('SO Lines', 'Count: ' + soLineItems.length + ', Tax: ' + soTaxTotal);

                // STEP 6: Find all invoices from SO
                var invoiceIds = findInvoicesFromSO(soId);
                log.debug('Invoice IDs', JSON.stringify(invoiceIds));

                // STEP 7: Get all invoice line items
                var invoiceLineItems = [];
                var invoiceTaxTotal = 0;
                var invoiceDetails = [];

                if (invoiceIds.length > 0) {
                    invoiceLineItems = getInvoiceLineItems(invoiceIds);
                    invoiceTaxTotal = getInvoiceTaxTotal(invoiceIds);
                    invoiceDetails = getInvoiceDetails(invoiceIds);
                }
                log.debug('Invoice Lines', 'Count: ' + invoiceLineItems.length + ', Tax: ' + invoiceTaxTotal);

                // STEP 8: Match and compare
                var comparisonTable = matchLineItems(soLineItems, invoiceLineItems);

                // STEP 9: Calculate totals
                var totals = calculateTotals(comparisonTable, soTaxTotal, invoiceTaxTotal);

                // STEP 10: Identify problems
                var problemItems = identifyProblemItems(comparisonTable);

                // STEP 10.5: Find ALL related Credit Memos from this SO
                var allRelatedCMs = findAllRelatedCreditMemos(soId);
                log.debug('Related CMs', 'Found ' + allRelatedCMs.length + ' Credit Memos related to SO ' + soData.tranid);
                
                // Calculate total CM amount for comparison to variance
                var totalCMAmount = 0;
                for (var cmIdx = 0; cmIdx < allRelatedCMs.length; cmIdx++) {
                    totalCMAmount += allRelatedCMs[cmIdx].amount;
                }
                log.debug('Total CM Amount', totalCMAmount);

                // STEP 11: Build result object
                var varianceMatchesCMs = Math.abs(Math.abs(totals.mismatchVariance) - totalCMAmount) < 0.005; // Exact match only
                var conclusion = '';
                if (Math.abs(totals.mismatchVariance) < 0.01) {
                    conclusion = 'NO SO vs Invoice MISMATCH - overpayment from other source';
                } else if (varianceMatchesCMs && allRelatedCMs.length > 0) {
                    conclusion = 'SO vs Invoice MISMATCH of ' + Math.abs(totals.mismatchVariance).toFixed(2) + ' MATCHES the ' + allRelatedCMs.length + ' Credit Memo(s) totaling ' + totalCMAmount.toFixed(2) + ' - this variance caused the overpayment(s)';
                } else {
                    conclusion = 'SO vs Invoice MISMATCH found - this may have caused overpayment';
                }

                var result = {
                    creditMemo: cmData,
                    allRelatedCreditMemos: allRelatedCMs,
                    totalCMAmount: totalCMAmount,
                    customerDeposit: cdData,
                    salesOrder: soData,
                    invoices: invoiceDetails,
                    comparisonTable: comparisonTable,
                    totals: totals,
                    problemItems: problemItems,
                    varianceMatchesCMs: varianceMatchesCMs,
                    conclusion: conclusion
                };

                log.debug('Analysis Complete', 'Variance: ' + totals.totalVariance);
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

        return {
            onRequest: onRequest
        };
    });
