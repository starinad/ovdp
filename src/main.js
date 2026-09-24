function onOpen() {
    UI.setupMenu();
}

function onSelectionChange(e) {
    Cashflow.highlightCashflowMonth(e);
}

function openBondDialogFromCheckbox(e) {
    const range = e.range;
    if (
        range.getSheet().getName() !== Config.SHEET_NAMES.CASHFLOW ||
        range.getColumn() !== 14 ||
        range.getRow() < 2 ||
        e.value !== 'TRUE'
    ) {
        return;
    }

    const sheet = range.getSheet();
    const isin = String(sheet.getRange(range.getRow(), 13).getValue()).trim();
    range.uncheck();
    SpreadsheetApp.flush();
    if (isin) UI.showCouponsDialog(isin);
}

function setupSheet() {
    Sheets.setupSheet();
}

function showAddBondDialog() {
    UI.showAddBondDialog();
}

function showDeleteBondDialog() {
    UI.showDeleteBondDialog();
}

function addBond(bond) {
    Bonds.addBond(bond);
}

function regenerateAllCoupons() {
    Coupons.regenerateAllCoupons();
}

function refreshCashflow() {
    Cashflow.refreshCashflow('ALL');
}

function refreshCashflowFuture() {
    Cashflow.refreshCashflow('FUTURE');
}

function refreshCashflowRealized() {
    Cashflow.refreshCashflow('REALIZED');
}

function refreshAnalytics() {
    Analytics.refreshAnalytics();
}

function showCouponsDialog() {
    UI.showCouponsDialog();
}

function refreshAll() {
    Coupons.regenerateAllCoupons();
    Cashflow.refreshCashflow('FUTURE');
    Analytics.refreshAnalytics();
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'All data refreshed!',
        '✅ Done',
        5,
    );
}
