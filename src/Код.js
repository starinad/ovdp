// ============================================================
// OVDP INVESTMENT MANAGER — Google Apps Script
// ============================================================
// 
// SETUP INSTRUCTIONS:
// 1. Create a new Google Sheet
// 2. Go to Extensions → Apps Script
// 3. Paste this entire script
// 4. Save and run setupSheet() once (it will ask for permissions)
// 5. A custom menu "OVDP Manager" will appear in the sheet
//
// SHEET STRUCTURE (auto-created by setupSheet):
//   • Bonds       — your bond holdings
//   • Coupons     — auto-generated coupon schedule
//   • Cashflow    — monthly aggregation
//   • Analytics   — portfolio summary & metrics
//   • Ladder      — maturity distribution
//   • Config      — settings (tax rates, day count defaults)
// ============================================================

// ── CONSTANTS ──────────────────────────────────────────────

const SHEET_NAMES = {
  BONDS: 'Bonds',
  COUPONS: 'Coupons',
  CASHFLOW: 'Cashflow',
  ANALYTICS: 'Analytics',
  LADDER: 'Ladder',
  CONFIG: 'Config',
};

const BOND_HEADERS = [
  'ID', 'ISIN', 'Name', 'Status',
  'Face Value (UAH)', 'Quantity', 'Purchase Price (UAH)',
  'Accrued Interest at Purchase (UAH)', 'Interest Rate (%)',
  'Tax Rate (%)', 'Currency', 'Purchase Date', 'Maturity Date',
  'First Coupon Date', 'Coupon Frequency', 'Day Count Convention',
  'Fixed Coupon (UAH/unit)', 'Total Invested', 'Total Face Value', 'Notes',
  'Coupons Generated', 'Last Updated',
];

const COUPON_HEADERS = [
  'Bond ID', 'ISIN', 'Bond Name', 'Seq #',
  'Payment Date', 'Period Start', 'Period End', 'Accrued Days',
  'Gross Amount (UAH)', 'Tax (UAH)', 'Net Amount (UAH)',
  'Day Count', 'Status', 'Is First', 'Is Last',
];

const CASHFLOW_HEADERS = [
  'Month', 'Gross Coupon Income', 'Tax on Coupons',
  'Net Coupon Income', 'Maturity Payments',
  'Total Gross Cashflow', 'Total Net Cashflow',
  'Coupon Count', 'Maturity Count',
];

const ANALYTICS_HEADERS = ['Metric', 'Value'];

const LADDER_HEADERS = [
  'Maturity Bucket', 'Bond Count', 'Face Value (UAH)',
  'Percentage (%)', 'Bonds (ISINs)',
];

const FREQUENCIES = {
  'Monthly': 1,
  'Quarterly': 3,
  'Semi-Annual': 6,
  'Annual': 12,
  'Zero-Coupon': 0,
};

const DAY_COUNTS = ['ACT/365', 'ACT/ACT', '30/360'];


// ── MENU & SETUP ───────────────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 OVDP Manager')
    .addItem('📋 Setup Sheet (first time)', 'setupSheet')
    .addSeparator()
    .addItem('➕ Add Bond...', 'showAddBondDialog')
    .addItem('🔄 Regenerate All Coupons', 'regenerateAllCoupons')
    .addSeparator()
    .addItem('📊 Refresh Cashflow', 'refreshCashflow')
    .addItem('📈 Refresh Analytics', 'refreshAnalytics')
    .addItem('🔁 Refresh Everything', 'refreshAll')
    .addSeparator()
    .addItem('🗑️ Delete Bond...', 'showDeleteBondDialog')
    .addToUi();
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or get each sheet
  createOrGetSheet(ss, SHEET_NAMES.BONDS, BOND_HEADERS, [
    { col: 1, width: 80 },   // ID
    { col: 2, width: 140 },  // ISIN
    { col: 3, width: 200 },  // Name
    { col: 4, width: 90 },   // Status
    { col: 5, width: 130 },  // Face Value
    { col: 6, width: 80 },   // Qty
    { col: 7, width: 150 },  // Purchase Price
    { col: 8, width: 200 },  // Accrued Interest
    { col: 9, width: 120 },  // Rate
    { col: 10, width: 100 }, // Tax Rate
    { col: 11, width: 80 },  // Currency
    { col: 12, width: 120 }, // Purchase Date
    { col: 13, width: 120 }, // Maturity Date
    { col: 14, width: 130 }, // First Coupon
    { col: 15, width: 130 }, // Frequency
    { col: 16, width: 150 }, // Day Count
    { col: 17, width: 130 }, // Total Invested
    { col: 18, width: 130 }, // Total Face
    { col: 19, width: 200 }, // Notes
    { col: 20, width: 120 }, // Coupons Gen
    { col: 21, width: 150 }, // Last Updated
  ]);

  createOrGetSheet(ss, SHEET_NAMES.COUPONS, COUPON_HEADERS, [
    { col: 1, width: 80 },
    { col: 2, width: 140 },
    { col: 3, width: 200 },
    { col: 5, width: 120 },
    { col: 6, width: 120 },
    { col: 7, width: 120 },
    { col: 9, width: 140 },
    { col: 10, width: 100 },
    { col: 11, width: 140 },
  ]);

  createOrGetSheet(ss, SHEET_NAMES.CASHFLOW, CASHFLOW_HEADERS, [
    { col: 1, width: 120 },
    { col: 2, width: 160 },
    { col: 3, width: 130 },
    { col: 4, width: 150 },
    { col: 5, width: 150 },
    { col: 6, width: 160 },
    { col: 7, width: 160 },
  ]);

  createOrGetSheet(ss, SHEET_NAMES.ANALYTICS, ANALYTICS_HEADERS, [
    { col: 1, width: 280 },
    { col: 2, width: 200 },
  ]);

  createOrGetSheet(ss, SHEET_NAMES.LADDER, LADDER_HEADERS, [
    { col: 1, width: 140 },
    { col: 2, width: 110 },
    { col: 3, width: 140 },
    { col: 4, width: 120 },
    { col: 5, width: 300 },
  ]);

  // Config sheet with defaults
  const configSheet = createOrGetSheet(ss, SHEET_NAMES.CONFIG,
    ['Setting', 'Value', 'Description'], [
      { col: 1, width: 200 },
      { col: 2, width: 150 },
      { col: 3, width: 400 },
    ]);

  const configData = configSheet.getDataRange().getValues();
  if (configData.length <= 1) {
    configSheet.getRange(2, 1, 4, 3).setValues([
      ['Default Tax Rate (%)', 0, 'Applied to new bonds if not specified (0 = tax-exempt)'],
      ['Default Day Count', 'ACT/365', 'ACT/365, ACT/ACT, or 30/360'],
      ['Default Coupon Frequency', 'Semi-Annual', 'Monthly, Quarterly, Semi-Annual, Annual'],
      ['Default Currency', 'UAH', 'Currency code'],
    ]);
  }

  // Activate bonds sheet
  ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.BONDS));

  SpreadsheetApp.getUi().alert(
    '✅ Setup Complete',
    'OVDP Manager is ready. Use the "💰 OVDP Manager" menu to add bonds.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createOrGetSheet(ss, name, headers, widths) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  // Set headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Set column widths
  if (widths) {
    widths.forEach(w => sheet.setColumnWidth(w.col, w.width));
  }

  return sheet;
}


// ── ADD BOND DIALOG ────────────────────────────────────────

function showAddBondDialog() {
  const config = getConfig();
  const html = HtmlService.createHtmlOutput(getAddBondHtml(config))
    .setWidth(520)
    .setHeight(680)
    .setTitle('Add New Bond');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Bond (OVDP)');
}

function getAddBondHtml(config) {
  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; background: #f8f9fa; }
    h2 { color: #1a73e8; margin-bottom: 16px; font-size: 18px; }
    .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .form-group { display: flex; flex-direction: column; }
    .form-group.full { grid-column: 1 / -1; }
    label { font-size: 12px; font-weight: 500; color: #5f6368; margin-bottom: 4px; }
    input, select { padding: 8px 12px; border: 1px solid #dadce0; border-radius: 6px;
                    font-size: 14px; outline: none; transition: border 0.2s; }
    input:focus, select:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px #1a73e820; }
    .btn { padding: 10px 24px; border: none; border-radius: 6px; font-size: 14px;
           cursor: pointer; font-weight: 500; margin-top: 16px; }
    .btn-primary { background: #1a73e8; color: white; }
    .btn-primary:hover { background: #1557b0; }
    .btn-secondary { background: #e8eaed; color: #3c4043; margin-right: 8px; }
    .actions { display: flex; justify-content: flex-end; margin-top: 20px; }
    .hint { font-size: 11px; color: #80868b; margin-top: 2px; }
    .section-label { grid-column: 1 / -1; font-size: 13px; font-weight: 600;
                     color: #1a73e8; margin-top: 8px; padding-bottom: 4px;
                     border-bottom: 1px solid #e8eaed; }
    .info-box { grid-column: 1 / -1; background: #e8f0fe; border-radius: 8px;
                padding: 10px 14px; font-size: 12px; color: #174ea6; }
  </style>
</head>
<body>
  <h2>🏦 Add OVDP Bond</h2>
  <div class="form-grid">
    <div class="form-group">
      <label>ISIN *</label>
      <input id="isin" placeholder="UA4000204746" maxlength="12">
    </div>
    <div class="form-group">
      <label>Name</label>
      <input id="name" placeholder="OVDP 15.05.2026 19.5%">
    </div>
    <div class="form-group">
      <label>Face Value (UAH) *</label>
      <input id="faceValue" type="number" step="0.01" value="1000">
    </div>
    <div class="form-group">
      <label>Quantity *</label>
      <input id="quantity" type="number" min="1" value="1">
    </div>
    <div class="form-group">
      <label>Purchase Price (UAH) *</label>
      <input id="purchasePrice" type="number" step="0.01" value="1000">
      <span class="hint">Clean price per unit</span>
    </div>
    <div class="form-group">
      <label>Accrued Interest (UAH)</label>
      <input id="accruedInterest" type="number" step="0.01" value="0">
      <span class="hint">AI paid to seller at settlement</span>
    </div>
    <div class="form-group">
      <label>Interest Rate (%) *</label>
      <input id="interestRate" type="number" step="0.01" value="19.5">
    </div>
    <div class="form-group">
      <label>Tax Rate (%)</label>
      <input id="taxRate" type="number" step="0.01" value="${config.defaultTaxRate}">
      <span class="hint">0 = tax-exempt</span>
    </div>
    <div class="form-group">
      <label>Purchase Date *</label>
      <input id="purchaseDate" type="date">
    </div>
    <div class="form-group">
      <label>Maturity Date *</label>
      <input id="maturityDate" type="date">
      <span class="hint">Actual maturity (last payment date)</span>
    </div>
    <div class="form-group">
      <label>First Coupon Date</label>
      <input id="firstCouponDate" type="date">
      <span class="hint">Leave empty to auto-calculate</span>
    </div>
    <div class="form-group">
      <label>Coupon Frequency</label>
      <select id="couponFrequency">
        <option ${config.defaultFrequency === 'Monthly' ? 'selected' : ''}>Monthly</option>
        <option ${config.defaultFrequency === 'Quarterly' ? 'selected' : ''}>Quarterly</option>
        <option ${config.defaultFrequency === 'Semi-Annual' ? 'selected' : ''}>Semi-Annual</option>
        <option ${config.defaultFrequency === 'Annual' ? 'selected' : ''}>Annual</option>
        <option ${config.defaultFrequency === 'Zero-Coupon' ? 'selected' : ''}>Zero-Coupon</option>
      </select>
    </div>
    <div class="form-group">
      <label>Day Count Convention</label>
      <select id="dayCount">
        <option ${config.defaultDayCount === 'ACT/365' ? 'selected' : ''}>ACT/365</option>
        <option ${config.defaultDayCount === 'ACT/ACT' ? 'selected' : ''}>ACT/ACT</option>
        <option ${config.defaultDayCount === '30/360' ? 'selected' : ''}>30/360</option>
      </select>
    </div>

    <div class="section-label">💰 Coupon Amount Mode</div>
    <div class="info-box">
      <strong>Fixed coupon:</strong> enter the exact coupon per unit from your broker statement.<br>
      <strong>Leave empty (0):</strong> system will calculate from interest rate × days × day count convention.
    </div>
    <div class="form-group">
      <label>Fixed Coupon per Unit (UAH)</label>
      <input id="fixedCoupon" type="number" step="0.01" value="0">
      <span class="hint">From broker: coupon amount per 1 unit per period. 0 = auto-calculate.</span>
    </div>
    <div class="form-group">
      <label>Currency</label>
      <input id="currency" value="${config.defaultCurrency}" maxlength="3">
    </div>

    <div class="form-group full">
      <label>Notes</label>
      <input id="notes" placeholder="Optional notes...">
    </div>
  </div>
  <div class="actions">
    <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
    <button class="btn btn-primary" onclick="submitBond()">Add Bond</button>
  </div>
  <script>
    // Set default dates
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('purchaseDate').value = today;

    function submitBond() {
      const bond = {
        isin: document.getElementById('isin').value.trim(),
        name: document.getElementById('name').value.trim(),
        faceValue: parseFloat(document.getElementById('faceValue').value) || 0,
        quantity: parseInt(document.getElementById('quantity').value) || 1,
        purchasePrice: parseFloat(document.getElementById('purchasePrice').value) || 0,
        accruedInterest: parseFloat(document.getElementById('accruedInterest').value) || 0,
        interestRate: parseFloat(document.getElementById('interestRate').value) || 0,
        taxRate: parseFloat(document.getElementById('taxRate').value) || 0,
        fixedCoupon: parseFloat(document.getElementById('fixedCoupon').value) || 0,
        currency: document.getElementById('currency').value.trim() || 'UAH',
        purchaseDate: document.getElementById('purchaseDate').value,
        maturityDate: document.getElementById('maturityDate').value,
        firstCouponDate: document.getElementById('firstCouponDate').value,
        couponFrequency: document.getElementById('couponFrequency').value,
        dayCount: document.getElementById('dayCount').value,
        notes: document.getElementById('notes').value.trim(),
      };

      if (!bond.isin || !bond.purchaseDate || !bond.maturityDate) {
        alert('Please fill in all required fields (*)');
        return;
      }
      if (bond.faceValue <= 0 || bond.purchasePrice <= 0 || bond.interestRate < 0) {
        alert('Face value, purchase price must be > 0');
        return;
      }
      if (bond.maturityDate <= bond.purchaseDate) {
        alert('Maturity date must be after purchase date');
        return;
      }

      google.script.run
        .withSuccessHandler(() => {
          google.script.host.close();
        })
        .withFailureHandler(err => alert('Error: ' + err.message))
        .addBond(bond);
    }
  </script>
</body>
</html>`;
}


// ── BOND CRUD ──────────────────────────────────────────────

function addBond(bond) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bondsSheet = ss.getSheetByName(SHEET_NAMES.BONDS);

  const id = getNextBondId(bondsSheet);
  const purchaseDate = new Date(bond.purchaseDate);
  const maturityDate = new Date(bond.maturityDate);
  const firstCouponDate = bond.firstCouponDate ? new Date(bond.firstCouponDate) : '';
  const totalInvested = (bond.purchasePrice + bond.accruedInterest) * bond.quantity;
  const totalFace = bond.faceValue * bond.quantity;
  const fixedCoupon = bond.fixedCoupon || 0;

  // Generate coupons
  const coupons = generateCouponSchedule({
    faceValue: bond.faceValue,
    quantity: bond.quantity,
    interestRate: bond.interestRate,
    taxRate: bond.taxRate,
    fixedCouponPerUnit: fixedCoupon,
    purchaseDate: purchaseDate,
    maturityDate: maturityDate,
    firstCouponDate: firstCouponDate || null,
    couponFrequency: bond.couponFrequency,
    dayCountConvention: bond.dayCount,
    accruedInterest: bond.accruedInterest,
  });

  // Write bond row
  const bondRow = [
    id, bond.isin, bond.name || '', 'ACTIVE',
    bond.faceValue, bond.quantity, bond.purchasePrice,
    bond.accruedInterest, bond.interestRate,
    bond.taxRate, bond.currency,
    purchaseDate, maturityDate,
    firstCouponDate, bond.couponFrequency, bond.dayCount,
    fixedCoupon, totalInvested, totalFace, bond.notes || '',
    coupons.length, new Date(),
  ];

  bondsSheet.appendRow(bondRow);

  // Format the new row
  const lastRow = bondsSheet.getLastRow();
  formatBondRow(bondsSheet, lastRow);

  // Write coupons
  if (coupons.length > 0) {
    writeCoupons(ss, id, bond.isin, bond.name || '', coupons);
  }

  // Refresh computed sheets
  refreshCashflow();
  refreshAnalytics();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Added ${bond.isin} with ${coupons.length} coupons`,
    '✅ Bond Added', 5
  );
}

function getNextBondId(bondsSheet) {
  const data = bondsSheet.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = parseInt(data[i][0]);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  return maxId + 1;
}

function formatBondRow(sheet, row) {
  // Currency format for money columns
  const moneyFormat = '#,##0.00';
  [5, 7, 8, 17, 18, 19].forEach(col => {
    sheet.getRange(row, col).setNumberFormat(moneyFormat);
  });
  // Percentage format
  sheet.getRange(row, 9).setNumberFormat('0.00');
  sheet.getRange(row, 10).setNumberFormat('0.00');
  // Date format
  [12, 13, 14].forEach(col => {
    sheet.getRange(row, col).setNumberFormat('yyyy-mm-dd');
  });
  sheet.getRange(row, 22).setNumberFormat('yyyy-mm-dd hh:mm');
}


// ── COUPON GENERATION ENGINE ───────────────────────────────

/**
 * Generate coupon schedule.
 *
 * Two modes:
 *   1. FIXED COUPON: if fixedCouponPerUnit > 0, every regular coupon pays that exact amount × quantity.
 *      First/last stub coupons are pro-rated by (actual days / standard period days).
 *   2. CALCULATED: if fixedCouponPerUnit === 0, each coupon = faceValue × rate × days / divisor.
 *
 * Schedule logic:
 *   Coupon dates are determined by the bond's own schedule (first coupon date + frequency),
 *   NOT by the purchase date. The purchase date only affects which coupons the holder receives
 *   (coupons on or after purchase date) and whether the first received coupon is a stub.
 */
function generateCouponSchedule(input) {
  const freq = FREQUENCIES[input.couponFrequency];
  if (!freq || freq === 0) return [];

  // Step 1: Build the bond's FULL coupon date schedule (independent of purchase)
  const allCouponDates = buildBondCouponDates(
    input.maturityDate,
    input.firstCouponDate,
    freq
  );

  if (allCouponDates.length === 0) return [];

  // Step 2: Determine the standard period length (for pro-rating fixed coupons)
  const standardPeriodDays = freq === 6 ? 182.5 : freq === 3 ? 91.25 : freq === 12 ? 365 : 30.4;

  // Step 3: Filter to coupons the holder receives (payment date > purchase date)
  // and build coupon objects
  const coupons = [];

  for (let i = 0; i < allCouponDates.length; i++) {
    const paymentDate = allCouponDates[i];

    // Skip coupons that paid before or on purchase date
    if (paymentDate <= input.purchaseDate) continue;

    // Period start = previous coupon date (or bond issue anchor if first)
    const periodStart = i > 0 ? allCouponDates[i - 1] : inferPeriodStartBeforeFirst(allCouponDates[0], freq);
    const periodEnd = paymentDate;

    // Actual accrued period for this holder
    // For the first coupon the holder receives, period starts at purchase date
    const holderPeriodStart = coupons.length === 0 ? input.purchaseDate : periodStart;

    const yf = yearFraction(periodStart, periodEnd, input.dayCountConvention);
    const holderYf = yearFraction(holderPeriodStart, periodEnd, input.dayCountConvention);

    if (yf.days <= 0) continue;

    let grossAmount;

    if (input.fixedCouponPerUnit > 0) {
      // FIXED COUPON MODE
      // Regular coupon = fixedCouponPerUnit × quantity
      // For the FULL period (regardless of when holder bought — the holder
      // gets the full coupon; accrued interest at purchase compensates the seller)
      grossAmount = bankersRound(input.fixedCouponPerUnit * input.quantity * 100) / 100;
    } else {
      // CALCULATED MODE
      grossAmount = calculateCouponAmount(
        input.faceValue, input.quantity, input.interestRate, yf.days, yf.divisor
      );
    }

    const taxAmount = calculateTax(grossAmount, input.taxRate);
    const netAmount = grossAmount - taxAmount;

    const isFirst = coupons.length === 0;
    const isLast = i === allCouponDates.length - 1;

    coupons.push({
      paymentDate: paymentDate,
      periodStart: periodStart,
      periodEnd: periodEnd,
      accruedDays: yf.days,
      grossAmount: grossAmount,
      taxAmount: taxAmount,
      netAmount: netAmount,
      dayCount: input.dayCountConvention,
      isFirst: isFirst,
      isLast: isLast,
      sequenceNumber: coupons.length + 1,
      accruedAdjustment: isFirst ? input.accruedInterest || 0 : 0,
    });
  }

  return coupons;
}

/**
 * Build the bond's complete coupon date schedule from issuance to maturity.
 *
 * The key insight for OVDP: the coupon schedule belongs to the BOND, not to the holder.
 * We build the full schedule, then filter for the holder's purchase date separately.
 *
 * If firstCouponDate is provided, we build forward from it.
 * The maturity date is always the last date (may coincide with a regular coupon or be a stub).
 */
function buildBondCouponDates(maturityDate, firstCouponDate, monthsBetween) {
  if (!firstCouponDate) {
    // No first coupon date: walk backwards from maturity to build regular schedule
    const dates = [new Date(maturityDate)];
    let current = maturityDate;
    while (true) {
      current = addMonthsSafe(current, -monthsBetween);
      if (current.getTime() <= new Date(2000, 0, 1).getTime()) break; // sanity limit
      dates.unshift(new Date(current));
    }
    return dates;
  }

  // First coupon date provided: build forward from it
  const dates = [];
  let i = 0;
  while (true) {
    const date = addMonthsSafe(firstCouponDate, monthsBetween * i);
    if (date > maturityDate) break;
    dates.push(date);
    i++;
  }

  // If last generated date doesn't match maturity, add maturity as final payment
  // BUT: suppress tiny stubs (≤7 days gap). Many OVDPs have maturity 1 day after
  // the last regular coupon — the broker pays both on the maturity date, not as
  // a separate stub coupon.
  const lastDate = dates.length > 0 ? dates[dates.length - 1] : null;
  if (!lastDate) {
    dates.push(new Date(maturityDate));
  } else {
    const gapDays = daysBetween(lastDate, maturityDate);
    if (gapDays > 7) {
      // Real stub period — add it
      dates.push(new Date(maturityDate));
    } else if (gapDays > 0) {
      // Tiny gap (1-7 days) — move last coupon to maturity date instead
      dates[dates.length - 1] = new Date(maturityDate);
    }
  }

  return dates;
}

/**
 * Infer the period start date before the first coupon date.
 * This is the "theoretical previous coupon date" — one period before the first coupon.
 * Used to calculate the full first period length.
 */
function inferPeriodStartBeforeFirst(firstCouponDate, monthsBetween) {
  return addMonthsSafe(firstCouponDate, -monthsBetween);
}


// ── DAY COUNT CONVENTIONS ──────────────────────────────────

function yearFraction(periodStart, periodEnd, convention) {
  switch (convention) {
    case 'ACT/365': {
      const days = daysBetween(periodStart, periodEnd);
      return { days, divisor: 365 };
    }

    case 'ACT/ACT': {
      const days = daysBetween(periodStart, periodEnd);
      const startYear = periodStart.getFullYear();
      const endYear = periodEnd.getFullYear();

      if (startYear === endYear) {
        return { days, divisor: isLeapYear(startYear) ? 366 : 365 };
      }

      const yearBoundary = new Date(endYear, 0, 1);
      const daysInStart = daysBetween(periodStart, yearBoundary);
      const daysInEnd = daysBetween(yearBoundary, periodEnd);
      const startDenom = isLeapYear(startYear) ? 366 : 365;
      const endDenom = isLeapYear(endYear) ? 366 : 365;
      const weightedDivisor = Math.round(
        days / (daysInStart / startDenom + daysInEnd / endDenom)
      );
      return { days, divisor: weightedDivisor };
    }

    case '30/360': {
      let d1 = periodStart.getDate();
      let d2 = periodEnd.getDate();
      const m1 = periodStart.getMonth() + 1;
      const m2 = periodEnd.getMonth() + 1;
      const y1 = periodStart.getFullYear();
      const y2 = periodEnd.getFullYear();

      if (d1 === 31) d1 = 30;
      if (d2 === 31 && d1 >= 30) d2 = 30;

      const days = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
      return { days, divisor: 360 };
    }

    default:
      return yearFraction(periodStart, periodEnd, 'ACT/365');
  }
}


// ── FINANCIAL MATH ─────────────────────────────────────────

/**
 * Banker's rounding — round half to even to avoid systematic bias.
 */
function bankersRound(value) {
  const floor = Math.floor(value);
  const decimal = value - floor;
  if (Math.abs(decimal - 0.5) < 1e-10) {
    return floor % 2 === 0 ? floor : floor + 1;
  }
  return Math.round(value);
}

/**
 * Calculate coupon amount in UAH (2 decimal places).
 * faceValue and purchasePrice are in UAH (not kopecks — this is Sheets, not the backend).
 */
function calculateCouponAmount(faceValue, quantity, ratePercent, days, divisor) {
  const amount = (faceValue * quantity * ratePercent * days) / (divisor * 100);
  return bankersRound(amount * 100) / 100; // round to 2 decimals
}

function calculateTax(grossAmount, taxRatePercent) {
  if (taxRatePercent <= 0) return 0;
  const tax = (grossAmount * taxRatePercent) / 100;
  return bankersRound(tax * 100) / 100;
}


// ── DATE UTILITIES ─────────────────────────────────────────

function daysBetween(a, b) {
  const msPerDay = 86400000;
  const utcA = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const utcB = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((utcB - utcA) / msPerDay);
}

function isLeapYear(year) {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

function addMonthsSafe(date, months) {
  const result = new Date(date);
  const targetMonth = result.getMonth() + months;
  const targetDay = result.getDate();
  result.setMonth(targetMonth);
  // Handle overflow (e.g., Jan 31 + 1 month should be Feb 28, not Mar 3)
  if (result.getDate() !== targetDay) {
    result.setDate(0); // last day of previous month
  }
  return result;
}

function monthDifference(from, to) {
  return (to.getFullYear() - from.getFullYear()) * 12
    + (to.getMonth() - from.getMonth());
}

function formatMonth(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  return `${y}-${m}`;
}


// ── WRITE COUPONS ──────────────────────────────────────────

function writeCoupons(ss, bondId, isin, bondName, coupons) {
  const sheet = ss.getSheetByName(SHEET_NAMES.COUPONS);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const rows = coupons.map(c => {
    const paymentDate = new Date(c.paymentDate);
    paymentDate.setHours(0, 0, 0, 0);

    const status = paymentDate <= today ? 'PAID' : 'SCHEDULED';

    return [
      bondId, isin, bondName, c.sequenceNumber,
      c.paymentDate, c.periodStart, c.periodEnd, c.accruedDays,
      c.grossAmount, c.taxAmount, c.netAmount,
      c.dayCount, status,
      c.isFirst ? 'YES' : '', c.isLast ? 'YES' : '',
    ];
  });

  if (rows.length > 0) {
    const startRow = sheet.getLastRow() + 1;

    const range = sheet.getRange(startRow, 1, rows.length, rows[0].length);
    range.setValues(rows);

    // Batch formatting 🚀
    range.offset(0, 4, rows.length, 3)
      .setNumberFormat('yyyy-mm-dd');

    range.offset(0, 8, rows.length, 3)
      .setNumberFormat('#,##0.00');
  }
}


// ── REGENERATE ALL COUPONS ─────────────────────────────────

function regenerateAllCoupons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bondsSheet = ss.getSheetByName(SHEET_NAMES.BONDS);
  const couponsSheet = ss.getSheetByName(SHEET_NAMES.COUPONS);

  // Clear existing coupons (keep header)
  if (couponsSheet.getLastRow() > 1) {
    couponsSheet.getRange(2, 1, couponsSheet.getLastRow() - 1, COUPON_HEADERS.length).clearContent();
  }

  const bonds = getBondsData(bondsSheet);

  let totalCoupons = 0;

  bonds.forEach(bond => {
    if (bond.status === 'SOLD') return;

    const coupons = generateCouponSchedule({
      faceValue: bond.faceValue,
      quantity: bond.quantity,
      interestRate: bond.interestRate,
      taxRate: bond.taxRate,
      fixedCouponPerUnit: bond.fixedCoupon || 0,
      purchaseDate: bond.purchaseDate,
      maturityDate: bond.maturityDate,
      firstCouponDate: bond.firstCouponDate || null,
      couponFrequency: bond.couponFrequency,
      dayCountConvention: bond.dayCount,
    });

    if (coupons.length > 0) {
      writeCoupons(ss, bond.id, bond.isin, bond.name, coupons);
      totalCoupons += coupons.length;
    }

    // Update coupon count on bond row
    bondsSheet.getRange(bond.rowIndex, 21).setValue(coupons.length);
    bondsSheet.getRange(bond.rowIndex, 22).setValue(new Date());
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Generated ${totalCoupons} coupons for ${bonds.length} bonds`,
    '✅ Coupons Regenerated', 5
  );
}


// ── CASHFLOW AGGREGATION ───────────────────────────────────

function refreshCashflow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const couponsSheet = ss.getSheetByName(SHEET_NAMES.COUPONS);
  const bondsSheet = ss.getSheetByName(SHEET_NAMES.BONDS);
  const cashflowSheet = ss.getSheetByName(SHEET_NAMES.CASHFLOW);

  // Clear existing cashflow data
  if (cashflowSheet.getLastRow() > 1) {
    cashflowSheet.getRange(2, 1, cashflowSheet.getLastRow() - 1, CASHFLOW_HEADERS.length).clearContent();
  }

  const couponsData = couponsSheet.getDataRange().getValues();
  const bondsData = bondsSheet.getDataRange().getValues();

  // Aggregate coupons by month
  const monthlyMap = {};

  // Process coupons (skip cancelled)
  for (let i = 1; i < couponsData.length; i++) {
    const row = couponsData[i];
    const status = row[12];
    if (status === 'CANCELLED' || !row[4]) continue;

    const paymentDate = new Date(row[4]);
    const month = formatMonth(paymentDate);

    const gross = parseFloat(row[8]) || 0;
    const tax = parseFloat(row[9]) || 0;
    const net = parseFloat(row[10]) || 0;

    if (!monthlyMap[month]) {
      monthlyMap[month] = {
        grossCoupon: 0, tax: 0, netCoupon: 0,
        maturity: 0, couponCount: 0, maturityCount: 0,
      };
    }

    monthlyMap[month].grossCoupon += gross;
    monthlyMap[month].tax += tax;
    monthlyMap[month].netCoupon += net;
    monthlyMap[month].couponCount++;
  }

  // Process maturity payments
  for (let i = 1; i < bondsData.length; i++) {
    const row = bondsData[i];
    const status = row[3];
    if (status === 'SOLD' || !row[12]) continue;

    const maturityDate = new Date(row[12]);
    const month = formatMonth(maturityDate);
    const faceValue = parseFloat(row[4]) || 0;
    const quantity = parseInt(row[5]) || 0;
    const maturityAmount = faceValue * quantity;

    if (!monthlyMap[month]) {
      monthlyMap[month] = {
        grossCoupon: 0, tax: 0, netCoupon: 0,
        maturity: 0, couponCount: 0, maturityCount: 0,
      };
    }

    monthlyMap[month].maturity += maturityAmount;
    monthlyMap[month].maturityCount++;
  }

  // Sort by month and write
  const sortedMonths = Object.keys(monthlyMap).sort();

  const rows = sortedMonths.map(month => {
    const m = monthlyMap[month];
    const r = v => bankersRound(v * 100) / 100;
    return [
      month,
      r(m.grossCoupon), r(m.tax), r(m.netCoupon),
      r(m.maturity),
      r(m.grossCoupon + m.maturity),
      r(m.netCoupon + m.maturity),
      m.couponCount, m.maturityCount,
    ];
  });

  if (rows.length > 0) {
        // Format
    cashflowSheet
      .getRange(2, 2, rows.length, 6)
      .setNumberFormat('#,##0.00');
    cashflowSheet
      .getRange(2, 1, rows.length, 1)
      .setNumberFormat('@');

    cashflowSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

    // Add summary row
    const summaryRow = rows.length + 3;

    const formulas = [[
      'TOTAL',
      `=SUM(B2:B${rows.length + 1})`,
      `=SUM(C2:C${rows.length + 1})`,
      `=SUM(D2:D${rows.length + 1})`,
      `=SUM(E2:E${rows.length + 1})`,
      `=SUM(F2:F${rows.length + 1})`,
      `=SUM(G2:G${rows.length + 1})`,
      `=SUM(H2:H${rows.length + 1})`,
      `=SUM(I2:I${rows.length + 1})`,
    ]];

    const summaryRange = cashflowSheet.getRange(summaryRow, 1, 1, 9);
    summaryRange.setValues(formulas);
    summaryRange.setFontWeight('bold');

    summaryRange.offset(0, 1, 1, 6)
      .setNumberFormat('#,##0.00');
  }

  applyHeatmap(cashflowSheet, 2, rows.length);
}

function applyHeatmap(sheet, startRow, numRows) {
  if (numRows === 0) return;

  const col = 7; // Total Net Cashflow (G)

  const range = sheet.getRange(startRow, col, numRows, 1);
  const values = range.getValues().map(r => r[0]);

  const positive = values.filter(v => v > 0);
  const min = positive.length ? Math.min(...positive) : 0;
  const max = positive.length ? Math.max(...values) : 0;

  const backgrounds = values.map(v => {
    if (max === min) return ['#fff7cc']; // fallback

    const ratio = (v - min) / (max - min);

    // 🎨 Yellow → Red gradient
    let r = 255;
    let g = Math.round(255 - ratio * 180); // уменьшаем зелёный
    let b = Math.round(200 - ratio * 200); // уменьшаем синий

    return [`rgb(${r},${g},${Math.max(b, 0)})`];
  });

  range.setBackgrounds(backgrounds);
}


// ── ANALYTICS ──────────────────────────────────────────────

function refreshAnalytics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bondsSheet = ss.getSheetByName(SHEET_NAMES.BONDS);
  const couponsSheet = ss.getSheetByName(SHEET_NAMES.COUPONS);
  const analyticsSheet = ss.getSheetByName(SHEET_NAMES.ANALYTICS);
  const ladderSheet = ss.getSheetByName(SHEET_NAMES.LADDER);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let bonds = getBondsData(bondsSheet);
  // set MATURED status for expired bonds
  const statusRange = bondsSheet.getRange(2, 4, bonds.length, 1);
  const statuses = statusRange.getValues();
  let hasChanges = false;
  bonds.forEach((bond, i) => {
    if (bond.status === 'ACTIVE' && bond.maturityDate <= today) {
      bond.status = 'MATURED';
      statuses[i][0] = 'MATURED';
      hasChanges = true;
    }
  });
  if (hasChanges) {
    statusRange.setValues(statuses);
  }

  const activeBonds = bonds.filter(b => b.status === 'ACTIVE');
  const couponsData = couponsSheet.getDataRange().getValues();



  // ── Compute metrics ──

  let totalInvested = 0;
  let totalFaceValue = 0;
  let weightedRateSum = 0;
  let totalGrossCouponIncome = 0;
  let totalNetCouponIncome = 0;
  let totalScheduledGross = 0;
  let totalScheduledNet = 0;
  const upcomingCouponsThisMonth = [];
  let totalCouponsThisMonth = 0;
  const upcomingMaturitiesThisMonth = []
  let totalMaturitiesThisMonth = 0;

  const bondMap = {};
  activeBonds.forEach(bond => {
    bondMap[bond.id] = bond;
    const invested = (bond.purchasePrice + bond.accruedInterest) * bond.quantity;
    const face = bond.faceValue * bond.quantity;
    totalInvested += invested;
    totalFaceValue += face;
    weightedRateSum += invested * bond.interestRate;
   
    const maturityDate = normalizeDate(bond.maturityDate);
    const sameMonth =
        maturityDate.getFullYear() === today.getFullYear() &&
        maturityDate.getMonth() === today.getMonth();

      if (sameMonth) {
        totalMaturitiesThisMonth += bond.faceValue * bond.quantity;
        if (maturityDate >= today) {
          upcomingMaturitiesThisMonth.push({
            date: maturityDate,
            net: bond.faceValue * bond.quantity,
          });
        }
      }
  });

  const weightedAvgYield = totalInvested > 0
    ? bankersRound((weightedRateSum / totalInvested) * 100) / 100
    : 0;

  // Process coupons for income calculation
  for (let i = 1; i < couponsData.length; i++) {
    const row = couponsData[i];
    if (!row[4]) continue;
    const status = row[12];
    const payDate = normalizeDate(row[4]);
    const gross = parseFloat(row[8]) || 0;
    const net = parseFloat(row[10]) || 0;
    const bondId = row[0];

    if (status === 'CANCELLED') continue;

    if (status === 'PAID') {
      const bond = bondMap[bondId];

      let adjustedNet = net;

      // если это первый купон — вычитаем accrued interest
      const isFirst = row[13] === 'YES';

      if (isFirst && bond && bond.accruedInterest > 0) {
        adjustedNet -= bond.accruedInterest;
      }

      totalGrossCouponIncome += gross;
      totalNetCouponIncome += adjustedNet;
    }

    const sameMonth =
        payDate.getFullYear() === today.getFullYear() &&
        payDate.getMonth() === today.getMonth();
      if (sameMonth) {
        totalCouponsThisMonth += net;
      }

    if (status === 'SCHEDULED') {
      totalScheduledGross += gross;
      let adjustedNet = net;

      const isFirst = row[13] === 'YES';
      const bond = bondMap[bondId];

      if (isFirst && bond && bond.accruedInterest > 0) {
        adjustedNet -= bond.accruedInterest;
      }

      totalScheduledNet += adjustedNet;

      if (sameMonth && payDate >= today) {
        upcomingCouponsThisMonth.push({
          date: payDate,
          net: net,
        });
      }
    }
  }
  const r = v => bankersRound(v * 100) / 100;
  upcomingCouponsThisMonth.sort((a, b) => a.date - b.date);
  const upcomingCouponsText = upcomingCouponsThisMonth.length > 0
  ? upcomingCouponsThisMonth.map(c =>
      Utilities.formatDate(c.date, Session.getScriptTimeZone(), 'yyyy-MM-dd') +
      ' → ' + formatUAH(r(c.net))
    ).join('\n')
  : 'N/A';

  upcomingMaturitiesThisMonth.sort((a, b) => a.date - b.date)
  const upcomingMaturitiesText = upcomingMaturitiesThisMonth.length > 0
  ? upcomingMaturitiesThisMonth.map(c =>
      Utilities.formatDate(c.date, Session.getScriptTimeZone(), 'yyyy-MM-dd') +
      ' → ' + formatUAH(r(c.net))
    ).join('\n')
  : 'N/A';

  // Yearly projected income (from scheduled coupons)
  // Annualized: total scheduled income / years from TODAY to last maturity
  // This answers: "how much income per year does my current portfolio generate
  // over its remaining lifetime?"
  let yearlyGross = 0;
  let yearlyNet = 0;

  if (activeBonds.length > 0 && (totalScheduledGross > 0)) {
    // Find the latest maturity across all active bonds
    let latestMaturity = activeBonds[0].maturityDate;
    activeBonds.forEach(b => {
      if (b.maturityDate > latestMaturity) latestMaturity = b.maturityDate;
    });

    const yearsRemaining = daysBetween(today, latestMaturity) / 365;

    if (yearsRemaining > 0) {
      yearlyGross = totalScheduledGross / yearsRemaining;
      yearlyNet = totalScheduledNet / yearsRemaining;
    }
  }

  // Compute remaining time and total return
  let latestMaturity = null;
  let yearsRemaining = 0;
  if (activeBonds.length > 0) {
    latestMaturity = activeBonds[0].maturityDate;
    activeBonds.forEach(b => {
      if (b.maturityDate > latestMaturity) latestMaturity = b.maturityDate;
    });
    yearsRemaining = daysBetween(today, latestMaturity) / 365;
  }

  const totalReturn = totalScheduledGross + (totalFaceValue - totalInvested);

  const cashflows = buildCashflows(bonds, couponsData);

  let portfolioXirr = null;

  if (cashflows.length > 1) {
    try {
      portfolioXirr = xirr(cashflows);
    } catch (e) {
        portfolioXirr = null;
    }
  }

  // ── Write analytics ──

  if (analyticsSheet.getLastRow() > 1) {
    analyticsSheet.getRange(2, 1, analyticsSheet.getLastRow() - 1, 2).clearContent();
  }

  const metrics = [
    ['', ''],
    ['── PORTFOLIO OVERVIEW ──', ''],
    ['Active Bonds', activeBonds.length],
    ['Matured Bonds', bonds.filter(b => b.status === 'MATURED').length],
    ['Total Invested Capital (incl. accrued)', formatUAH(r(totalInvested))],
    ['Total Face Value', formatUAH(r(totalFaceValue))],
    ['Unrealized Capital Gain/Loss', formatUAH(r(totalFaceValue - totalInvested))],
    ['Weighted Average Yield (%)', weightedAvgYield.toFixed(2) + '%'],
    ['Portfolio XIRR (%)', portfolioXirr !== null ? (portfolioXirr * 100).toFixed(2) + '%' : 'N/A'],
    ['Portfolio Horizon (to last maturity)', r(yearsRemaining).toFixed(2) + ' years'],
    ['', ''],
    ['── INCOME (REALIZED) ──', ''],
    ['Total Gross Coupon Income Received', formatUAH(r(totalGrossCouponIncome))],
    ['Total Tax Paid', formatUAH(r(totalGrossCouponIncome - totalNetCouponIncome))],
    ['Total Net Coupon Income Received', formatUAH(r(totalNetCouponIncome))],
    ['', ''],
    ['── INCOME (PROJECTED) ──', ''],
    ['Total Scheduled Gross Income', formatUAH(r(totalScheduledGross))],
    ['Total Scheduled Net Income', formatUAH(r(totalScheduledNet))],
    ['Annualized Gross Income *', yearsRemaining >= 0.5
      ? formatUAH(r(yearlyGross))
      : 'N/A (horizon < 6 months)'],
    ['Annualized Net Income *', yearsRemaining >= 0.5
      ? formatUAH(r(yearlyNet))
      : 'N/A (horizon < 6 months)'],
    ['Monthly Avg Gross Income *', yearsRemaining >= 0.5
      ? formatUAH(r(yearlyGross / 12))
      : 'N/A (horizon < 6 months)'],
    ['Monthly Avg Net Income *', yearsRemaining >= 0.5
      ? formatUAH(r(yearlyNet / 12))
      : 'N/A (horizon < 6 months)'],
    ['', '* Annualized = total scheduled income ÷ years to last maturity.'],
    ['', '  Not shown when horizon < 6 months (misleading on short bonds).'],
    ['', '  Add longer-dated bonds for this metric to be useful.'],
    ['', ''],
    ['── TOTAL RETURN (PROJECTED) ──', ''],
    ['Total Coupon Income (scheduled)', formatUAH(r(totalScheduledGross))],
    ['Capital Gain/Loss at Maturity', formatUAH(r(totalFaceValue - totalInvested))],
    ['Total Projected Return', formatUAH(r(totalReturn))],
    ['Return on Investment (%)', totalInvested > 0 ? (r(totalReturn / totalInvested * 100)).toFixed(2) + '%' : 'N/A'],
    ['', ''],
    ['── UPCOMING ──', ''],
    ['Coupons This Month', upcomingCouponsText],
    ['Total Coupons This Month', formatUAH(r(totalCouponsThisMonth))],
    ['Maturities This Month', upcomingMaturitiesText],
    ['Total Maturities This Month', formatUAH(r(totalMaturitiesThisMonth))],
    ['', ''],
    ['Last Updated', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')],
  ];

  analyticsSheet.getRange(2, 1, metrics.length, 2).setValues(metrics);

  // Bold section headers
  metrics.forEach((row, i) => {
    if (row[0].startsWith('──')) {
      analyticsSheet.getRange(i + 2, 1).setFontWeight('bold').setFontColor('#1a73e8');
    }
  });

  // ── Maturity Ladder ──
  refreshLadder(ladderSheet, activeBonds, today);
}

function refreshLadder(ladderSheet, activeBonds, today) {
  if (ladderSheet.getLastRow() > 1) {
    ladderSheet.getRange(2, 1, ladderSheet.getLastRow() - 1, LADDER_HEADERS.length).clearContent();
  }

  const buckets = [
    { label: '0-3 months', maxDays: 91, bonds: [] },
    { label: '3-6 months', maxDays: 182, bonds: [] },
    { label: '6-12 months', maxDays: 365, bonds: [] },
    { label: '1-2 years', maxDays: 730, bonds: [] },
    { label: '2-5 years', maxDays: 1825, bonds: [] },
    { label: '5+ years', maxDays: Infinity, bonds: [] },
  ];

  let totalFace = 0;

  activeBonds.forEach(bond => {
    const daysToMaturity = daysBetween(today, bond.maturityDate);
    const faceTotal = bond.faceValue * bond.quantity;
    totalFace += faceTotal;

    for (const bucket of buckets) {
      if (daysToMaturity <= bucket.maxDays) {
        bucket.bonds.push({ isin: bond.isin, face: faceTotal });
        break;
      }
    }
  });

  const rows = buckets.map(b => {
    const bucketFace = b.bonds.reduce((sum, x) => sum + x.face, 0);
    const pct = totalFace > 0 ? bankersRound((bucketFace / totalFace) * 10000) / 100 : 0;
    return [
      b.label,
      b.bonds.length,
      bankersRound(bucketFace * 100) / 100,
      pct,
      b.bonds.map(x => x.isin).join(', '),
    ];
  });

  if (rows.length > 0) {
    ladderSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    for (let i = 0; i < rows.length; i++) {
      ladderSheet.getRange(i + 2, 3).setNumberFormat('#,##0.00');
      ladderSheet.getRange(i + 2, 4).setNumberFormat('0.00');
    }
  }
}


// ── DELETE BOND ─────────────────────────────────────────────

function showDeleteBondDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bondsSheet = ss.getSheetByName(SHEET_NAMES.BONDS);
  const bonds = getBondsData(bondsSheet);

  if (bonds.length === 0) {
    SpreadsheetApp.getUi().alert('No bonds to delete.');
    return;
  }

  const bondList = bonds.map(b => `${b.id}: ${b.isin} — ${b.name} (${b.status})`).join('\\n');
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Delete Bond',
    `Enter the Bond ID to delete:\\n\\n${bondList}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const bondId = parseInt(response.getResponseText().trim());
    if (isNaN(bondId)) {
      ui.alert('Invalid Bond ID.');
      return;
    }
    deleteBond(bondId);
  }
}

function deleteBond(bondId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bondsSheet = ss.getSheetByName(SHEET_NAMES.BONDS);
  const couponsSheet = ss.getSheetByName(SHEET_NAMES.COUPONS);

  // Find and remove bond row
  const bondsData = bondsSheet.getDataRange().getValues();
  let deleted = false;
  for (let i = bondsData.length - 1; i >= 1; i--) {
    if (parseInt(bondsData[i][0]) === bondId) {
      bondsSheet.deleteRow(i + 1);
      deleted = true;
      break;
    }
  }

  if (!deleted) {
    SpreadsheetApp.getUi().alert(`Bond ID ${bondId} not found.`);
    return;
  }

  // Remove associated coupons
  const couponsData = couponsSheet.getDataRange().getValues();
  for (let i = couponsData.length - 1; i >= 1; i--) {
    if (parseInt(couponsData[i][0]) === bondId) {
      couponsSheet.deleteRow(i + 1);
    }
  }

  refreshCashflow();
  refreshAnalytics();

  ss.toast(`Bond ${bondId} and its coupons deleted.`, '🗑️ Deleted', 5);
}


// ── HELPERS ────────────────────────────────────────────────
function normalizeDate(d) {
  if (!d) return null;
  const date = new Date(d);
  date.setHours(0, 0, 0, 0);
  return date;
}

function getBondsData(bondsSheet) {
  const data = bondsSheet.getDataRange().getValues();
  const bonds = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue; // skip empty rows

    bonds.push({
      rowIndex: i + 1,
      id: parseInt(row[0]),
      isin: row[1],
      name: row[2],
      status: row[3],
      faceValue: parseFloat(row[4]) || 0,
      quantity: parseInt(row[5]) || 0,
      purchasePrice: parseFloat(row[6]) || 0,
      accruedInterest: parseFloat(row[7]) || 0,
      interestRate: parseFloat(row[8]) || 0,
      taxRate: parseFloat(row[9]) || 0,
      currency: row[10],
      purchaseDate: normalizeDate(row[11]),
      maturityDate: normalizeDate(row[12]),
      firstCouponDate: normalizeDate(row[13]),
      couponFrequency: row[14],
      dayCount: row[15],
      fixedCoupon: parseFloat(row[16]) || 0,
    });
  }

  return bonds;
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);

  if (!sheet) {
    return {
      defaultTaxRate: 0,
      defaultDayCount: 'ACT/365',
      defaultFrequency: 'Semi-Annual',
      defaultCurrency: 'UAH',
    };
  }

  const data = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    if (key.includes('Tax Rate')) config.defaultTaxRate = parseFloat(value) || 0;
    if (key.includes('Day Count')) config.defaultDayCount = value || 'ACT/365';
    if (key.includes('Coupon Frequency')) config.defaultFrequency = value || 'Semi-Annual';
    if (key.includes('Currency')) config.defaultCurrency = value || 'UAH';
  }

  return config;
}

function formatUAH(amount) {
  return amount.toLocaleString('uk-UA', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }) + ' UAH';
}

function refreshAll() {
  regenerateAllCoupons();
  refreshCashflow();
  refreshAnalytics();
  SpreadsheetApp.getActiveSpreadsheet().toast('All data refreshed!', '✅ Done', 5);
}

function buildCashflows(bonds, couponsData) {
  const today = normalizeDate(new Date());
  const flows = [];

  // ❌ покупки
  bonds.forEach(bond => {
    const invested = (bond.purchasePrice + bond.accruedInterest) * bond.quantity;

    flows.push({
      date: bond.purchaseDate,
      amount: -invested,
    });

    if (normalizeDate(bond.maturityDate) >= today || bond.status !== 'SOLD') {
      flows.push({
        date: bond.maturityDate,
        amount: bond.faceValue * bond.quantity,
      });
    }
  });

  // ✅ купоны
  for (let i = 1; i < couponsData.length; i++) {
    const row = couponsData[i];
    const status = row[12];
    if (status === 'CANCELLED') continue;

    const date = normalizeDate(row[4]);
    const net = parseFloat(row[10]) || 0;

    flows.push({
      date,
      amount: net,
    });
  }

  return flows.sort((a, b) => a.date - b.date);
}

function xirr(cashflows, guess = 0.1) {
  const maxIter = 100;
  const tol = 1e-6;

  const t0 = cashflows[0].date;

  function npv(rate) {
    return cashflows.reduce((sum, cf) => {
      const days = (cf.date - t0) / 86400000;
      return sum + cf.amount / Math.pow(1 + rate, days / 365);
    }, 0);
  }

  function dnpv(rate) {
    return cashflows.reduce((sum, cf) => {
      const days = (cf.date - t0) / 86400000;
      const t = days / 365;
      return sum - (t * cf.amount) / Math.pow(1 + rate, t + 1);
    }, 0);
  }

  let rate = guess;

  for (let i = 0; i < maxIter; i++) {
    const f = npv(rate);
    const df = dnpv(rate);

    if (Math.abs(df) < 1e-10) break;

    const newRate = rate - f / df;

    if (Math.abs(newRate - rate) < tol) return newRate;

    rate = newRate;
  }

  return rate;
}