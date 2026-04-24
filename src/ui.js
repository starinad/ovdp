// eslint-disable-next-line no-unused-vars
const UI = {
    setupMenu() {
        const ui = SpreadsheetApp.getUi();
        ui.createMenu('💰 OVDP Manager')
            .addItem('📋 Setup Sheets (first time)', 'setupSheet')
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
    },

    showAddBondDialog() {
        const config = this._getConfig();
        const html = HtmlService.createHtmlOutput(this._getAddBondHtml(config))
            .setWidth(520)
            .setHeight(680)
            .setTitle('Add New Bond');
        SpreadsheetApp.getUi().showModalDialog(html, 'Add New Bond (OVDP)');
    },

    showDeleteBondDialog() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bondsSheet = ss.getSheetByName(Config.SHEET_NAMES.BONDS);
        const bonds = Bonds.getBondsData(bondsSheet);

        if (bonds.length === 0) {
            SpreadsheetApp.getUi().alert('No bonds to delete.');
            return;
        }

        const bondList = bonds
            .map((b) => `${b.id}: ${b.isin} — ${b.name} (${b.status})`)
            .join('\\n');
        const ui = SpreadsheetApp.getUi();
        const response = ui.prompt(
            'Delete Bond',
            `Enter the Bond ID to delete:\\n\\n${bondList}`,
            ui.ButtonSet.OK_CANCEL,
        );

        if (response.getSelectedButton() === ui.Button.OK) {
            const bondId = parseInt(response.getResponseText().trim());
            if (isNaN(bondId)) {
                ui.alert('Invalid Bond ID.');
                return;
            }
            Bonds.deleteBond(bondId);
        }
    },

    _getConfig() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(Config.SHEET_NAMES.CONFIG);

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
            if (key.includes('Tax Rate'))
                config.defaultTaxRate = parseFloat(value) || 0;
            if (key.includes('Day Count'))
                config.defaultDayCount = value || 'ACT/365';
            if (key.includes('Coupon Frequency'))
                config.defaultFrequency = value || 'Semi-Annual';
            if (key.includes('Currency'))
                config.defaultCurrency = value || 'UAH';
        }

        return config;
    },

    _getAddBondHtml(config) {
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
    },
};
