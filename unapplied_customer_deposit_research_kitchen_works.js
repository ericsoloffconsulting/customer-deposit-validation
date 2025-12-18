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
define(['N/ui/serverWidget', 'N/query', 'N/log', 'N/runtime', 'N/url'],
    /**
     * @param {serverWidget} serverWidget
     * @param {query} query
     * @param {log} log
     * @param {runtime} runtime
     * @param {url} url
     */
    function (serverWidget, query, log, runtime, url) {

        /**
         * Handles GET and POST requests to the Suitelet
         * @param {Object} context - NetSuite context object containing request/response
         */
        function onRequest(context) {
            if (context.request.method === 'GET') {
                handleGet(context);
            } else {
                // POST requests just redirect back to GET
                handleGet(context);
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
            var scriptUrl = url.resolveScript({
                scriptId: runtime.getCurrentScript().id,
                deploymentId: runtime.getCurrentScript().deploymentId,
                returnExternalUrl: false
            });

            // Get unapplied deposit data
            var deposits = searchUnappliedDeposits();

            // Calculate totals
            var totalDeposits = deposits.length;
            var totalDepositAmount = 0;
            var totalAppliedAmount = 0;
            var totalUnappliedAmount = 0;

            for (var i = 0; i < deposits.length; i++) {
                totalDepositAmount += deposits[i].depositAmount || 0;
                totalAppliedAmount += deposits[i].amountApplied || 0;
                totalUnappliedAmount += deposits[i].amountUnapplied || 0;
            }

            // Get unapplied credit memo data (needed for summary)
            var creditMemos = searchUnappliedCreditMemos();

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

            var html = '';

            // Add styles
            html += '<style>' + getStyles() + '</style>';

            // Main container
            html += '<div class="portal-container">';

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
            html += '<span class="summary-total-label">Unapplied Prior To:</span>';
            html += '<input type="date" id="priorPeriodDate" class="prior-period-date-input" value="2024-12-31">';
            html += '</div>';
            html += '<span class="summary-total-amount" id="priorPeriodAmount">$0.00</span>';
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
            html += '</div>';
            html += '<div class="summary-total">';
            html += '<div class="prior-period-header">';
            html += '<span class="summary-total-label">Overpayment Date Prior To:</span>';
            html += '<input type="date" id="cmPriorPeriodDate" class="prior-period-date-input" value="2024-12-31">';
            html += '</div>';
            html += '<span class="summary-total-amount" id="cmPriorPeriodAmount">$0.00</span>';
            html += '</div>';
            html += '</div>';
            html += '</div>';

            html += '</div>'; // Close summary-row

            // Data Section - Customer Deposits
            html += buildDataSection('deposits', 'True Customer Deposits', 
                'Customer deposits linked to Kitchen Retail Sales orders that have not been fully applied', 
                deposits, scriptUrl);

            // Data Section - Credit Memos
            html += buildCreditMemoDataSection('creditmemos', 'Credit Memo Overpayments from Customer Deposits', 
                'Unapplied credit memos created from overpayment customer deposits linked to Kitchen Retail Sales orders', 
                creditMemos, scriptUrl);

            html += '</div>'; // Close portal-container

            // Add SheetJS library for Excel export
            html += '<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>';

            // Add JavaScript
            html += '<script>' + getJavaScript(scriptUrl) + '</script>';

            return html;
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
        function buildDataSection(sectionId, title, description, data, scriptUrl) {
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
            html += '<th onclick="sortTable(\'' + sectionId + '\', 0)">Deposit #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 1)">Deposit Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 2)">Customer</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 3)">Deposit Amount</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 4)">Amount Applied</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 5)">Amount Unapplied</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 6)">Status</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 7)">Sales Order #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 8)">SO Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 9)">SO Status</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 10)">Selling Location</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 11)">Sales Rep</th>';
            html += '</tr>';
            html += '</thead>';
            html += '<tbody>';

            for (var i = 0; i < deposits.length; i++) {
                var dep = deposits[i];
                var rowClass = (i % 2 === 0) ? 'even-row' : 'odd-row';

                html += '<tr class="' + rowClass + '" id="dep-row-' + dep.depositId + '">';

                // Deposit # with link
                html += '<td><a href="/app/accounting/transactions/custdep.nl?id=' + dep.depositId + '" target="_blank">' + escapeHtml(dep.depositNumber) + '</a></td>';

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
            html += '<th onclick="sortTable(\'' + sectionId + '\', 0)">Credit Memo #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 1)">CM Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 2)">Customer</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 3)">CM Amount</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 4)">Amount Applied</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 5)">Amount Unapplied</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 6)">Status</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 7)">Linked CD #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 8)">Overpayment Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 9)">CD Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 10)">Sales Order #</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 11)">SO Date</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 12)">Selling Location</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 13)">Unbilled Orders</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 14)">Deposit Balance</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 15)">A/R Balance</th>';
            html += '<th onclick="sortTable(\'' + sectionId + '\', 16)">Sales Rep</th>';
            html += '</tr>';
            html += '</thead>';
            html += '<tbody>';

            for (var i = 0; i < creditMemos.length; i++) {
                var cm = creditMemos[i];
                var rowClass = (i % 2 === 0) ? 'even-row' : 'odd-row';

                html += '<tr class="' + rowClass + '" id="cm-row-' + cm.cmId + '">';

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
         * @returns {Array} Array of deposit objects
         */
        function searchUnappliedDeposits() {
            var deposits = [];

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
                           AND ptll2.linktype = 'DepAppl') AS amount_applied,
                        (t.foreigntotal - (SELECT COALESCE(SUM(depa2.foreigntotal), 0)
                                           FROM previousTransactionLineLink ptll2
                                           LEFT JOIN transaction depa2 ON ptll2.nextdoc = depa2.id
                                           WHERE ptll2.previousdoc = t.id
                                             AND ptll2.linktype = 'DepAppl')) AS amount_unapplied,
                        tl_dep.createdfrom AS so_id,
                        so.tranid AS so_number,
                        so.trandate AS so_date,
                        so.status AS so_status,
                        c.altname AS customer_name,
                        d.name AS so_department,
                        emp.firstname || ' ' || emp.lastname AS salesrep_name
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
                      AND t.status != 'C'
                      AND i.incomeaccount = 338
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
                             d.name,
                             emp.firstname,
                             emp.lastname
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
                        soDepartment: row.so_department,
                        salesrepName: row.salesrep_name
                    });
                }

                log.debug('Search Results', 'Found ' + deposits.length + ' unapplied deposits');

            } catch (e) {
                log.error('Error Searching Deposits', {
                    error: e.message,
                    stack: e.stack
                });
            }

            return deposits;
        }

        /**
         * Searches for unapplied credit memos from overpayment deposits linked to Kitchen Works sales orders
         * @returns {Array} Array of credit memo objects
         */
        function searchUnappliedCreditMemos() {
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

                /* Main container - avoid targeting global td */
                '.portal-container { margin: 0; padding: 20px; border: none; background: transparent; position: relative; }' +

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

                /* Success/Error Messages */
                '.success-msg { background-color: #d4edda; color: #155724; padding: 12px; border: 1px solid #c3e6cb; border-radius: 6px; margin: 15px 0; font-size: 13px; }' +
                '.error-msg { background-color: #f8d7da; color: #721c24; padding: 12px; border: 1px solid #f5c6cb; border-radius: 6px; margin: 15px 0; font-size: 13px; }';
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
                'document.addEventListener(\'DOMContentLoaded\', function() {' +
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
                '});' +
                '' +
                '/* Calculate unapplied amount for deposits prior to selected date */' +
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
                '        var dateCell = rows[i].cells[1];' +
                '        var unappliedCell = rows[i].cells[5];' +
                '        var dateStr = dateCell.getAttribute(\'data-date\');' +
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
                '}';
        }

        return {
            onRequest: onRequest
        };
    });
