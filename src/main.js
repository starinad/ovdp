function onOpen() {
    UI.setupMenu();
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
    Cashflow.refreshCashflow();
}

function refreshAnalytics() {
    Analytics.refreshAnalytics();
}

function refreshAll() {
    Coupons.regenerateAllCoupons();
    Cashflow.refreshCashflow();
    Analytics.refreshAnalytics();
    SpreadsheetApp.getActiveSpreadsheet().toast(
        'All data refreshed!',
        '✅ Done',
        5,
    );
}
