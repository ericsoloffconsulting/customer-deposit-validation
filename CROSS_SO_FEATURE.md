# Cross-SO Deposit Application Analysis Feature

## Overview
This feature detects when Customer Deposits from one Sales Order are mistakenly applied to Invoices from a different Sales Order, which is a root cause of credit memo overpayments.

## Root Cause Example (SO3735)
- **SO3735** had deposits: CD4759 ($8,169.05) + CD13271 ($7,669.05) = $15,838.10
- **INV33988** (from SO3735) was paid by:
  - CD13271: $7,669.05 ✓ (correct - from SO3735)
  - CD4759: $619.91 ✓ (correct - from SO3735)
  - **CD4765: $424.24 ✗ (WRONG - from SO3729, not SO3735!)**
- Result: CD4759 had $424.24 that couldn't apply, creating overpayment **CM8065**

## Implementation Details

### 1. UI Components Added

#### Button in Credit Memo Table
- **Location**: Line ~658
- **HTML**: `<button class="cross-so-btn" onclick="showCrossSOAnalysis(...)">CD Cross-SO</button>`
- **Style**: Pink/magenta gradient (`#e91e63` to `#c2185b`)
- **Position**: Stacked vertically below SO↔INV button with 6px top margin

#### Modal HTML Structure
- **Location**: Lines ~316-330
- **Elements**:
  - `#crossSOModal` - Main modal container
  - `#crossSOModalBody` - Content area
  - `#crossSOModalOverlay` - Backdrop overlay
  - `.cross-so-header` - Pink gradient header

### 2. CSS Styling

#### Button Styling (Line ~1289-1298)
```css
.cross-so-btn {
  padding: 4px 8px;
  font-size: 11px;
  font-weight: 600;
  color: #fff;
  background: linear-gradient(135deg, #e91e63, #c2185b);
  border: none;
  border-radius: 4px;
  cursor: pointer;
  margin: 6px auto 0;
}
```

#### Modal Styling (Lines ~1322-1375)
- `.cross-so-header` - Pink gradient header (overrides default blue)
- `.match-row` - Light green background (#f1f8f4)
- `.mismatch-row` - Light red background (#fef5f5)
- `.status-badge` - Pill-style status indicators
- `.status-success` - Green badge for matches
- `.status-error` - Red badge for cross-SO mismatches

### 3. Client-Side JavaScript Functions (Lines ~1931-2037)

#### `showCrossSOAnalysis(creditMemoId)`
- Opens modal with loading state
- Makes POST request with `action: 'crossSOAnalysis'`
- Calls `renderCrossSOResult()` with response data

#### `hideCrossSOModal()`
- Hides modal and overlay by removing `.visible` class

#### `renderCrossSOResult(data)`
- Displays 4 summary cards:
  - Total Applications
  - Same-SO Matches (green)
  - Cross-SO Mismatches (red)
  - Cross-SO Amount (red)
- Renders table with columns:
  - Status (MATCH/CROSS-SO badge)
  - Customer Deposit (with link)
  - CD Source SO
  - Invoice (with link)
  - INV Source SO
  - Amount Applied

### 4. Backend POST Handler (Lines ~87-97)
```javascript
if (body.action === 'crossSOAnalysis') {
    var creditMemoId = body.creditMemoId;
    var result = analyzeCrossSODeposits(creditMemoId);
    response.write(JSON.stringify({ success: true, data: result }));
    return;
}
```

### 5. Backend Analysis Function (Lines ~3295-3407)

#### `analyzeCrossSODeposits(creditMemoId)`

**Logic Flow**:
1. Load Credit Memo record to get `customerId`
2. Execute SuiteQL query to get all deposit applications for customer
3. Join tables:
   - `Transaction depa` (Deposit Application)
   - `TransactionLine depa_line` (to get CD reference)
   - `Transaction cd` (Customer Deposit)
   - `Transaction cd_so` (CD's source SO via `createdfrom`)
   - `Transaction inv` (Invoice from DEPA)
   - `Transaction inv_so` (Invoice's source SO via `createdfrom`)
4. Compare `cd_so.tranid` vs `inv_so.tranid` for each application
5. Flag mismatches where CD source SO ≠ Invoice source SO

**SuiteQL Query**:
```sql
SELECT 
    depa.id, depa.tranid,
    cd.id, cd.tranid, cd_so.tranid as cdSourceSO,
    depa.amount,
    inv.id, inv.tranid, inv_so.tranid as invSourceSO
FROM Transaction depa
    INNER JOIN TransactionLine depa_line ON depa.id = depa_line.transaction
    INNER JOIN Transaction cd ON depa_line.custcol_deposit_id = cd.id
    LEFT JOIN Transaction cd_so ON cd.createdfrom = cd_so.id
    INNER JOIN Transaction inv ON depa.custbody_deposit_application_invoice = inv.id
    LEFT JOIN Transaction inv_so ON inv.createdfrom = inv_so.id
WHERE 
    depa.type = 'DepAppl'
    AND depa.entity = ${customerId}
    AND cd.type = 'CustDep'
    AND inv.type = 'CustInvc'
ORDER BY depa.trandate DESC, depa.id DESC
```

**Return Object**:
```javascript
{
    customerId: number,
    customerName: string,
    totalApplications: number,
    matches: number,
    mismatches: number,
    crossSOAmount: number,
    applications: [
        {
            depaId, depaTranId,
            cdId, cdTranId, cdSourceSO,
            invId, invTranId, invSourceSO,
            amount, isMismatch
        },
        ...
    ]
}
```

## Testing Instructions

1. **Test with SO3735** (known cross-SO case):
   - Find CM8065 in credit memo table
   - Click "CD Cross-SO" button
   - Verify modal shows:
     - Cross-SO Mismatches: > 0
     - Cross-SO Amount: Includes $424.24
     - Table row showing CD4765 (SO3729) → INV33988 (SO3735) mismatch

2. **Test with normal CM** (no cross-SO issues):
   - Click "CD Cross-SO" on different CM
   - Verify all rows show "MATCH" status (green)
   - Verify Cross-SO Mismatches = 0

3. **Visual Verification**:
   - Button appears below SO↔INV button
   - Pink/magenta color matches button gradient
   - Modal header is pink (not blue)
   - Mismatch rows have light red background
   - Match rows have light green background

## Key Custom Fields Required
**NONE!** This feature uses only standard NetSuite fields:
- `previousTransactionLineLink` table (standard NetSuite relationship table)
- `transactionLine.createdfrom` (standard field containing source SO)
- `transactionLine.mainline` (standard field identifying main transaction line)
- `transactionLine.netamount` (standard field for transaction amount)

## Integration Points
- Uses same modal styling as SO↔INV comparison tool
- Follows same client/server architecture pattern
- Uses same POST request pattern with JSON response
- Reuses `formatCompCurrency()` function for amounts

## Future Enhancements
1. Add "Fix Cross-SO" button to automatically create correcting transactions
2. Show drill-down details for each mismatch (full payment chain)
3. Export cross-SO report to Excel
4. Add preventive validation on Deposit Application creation
5. Dashboard widget showing top customers with cross-SO issues
