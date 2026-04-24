// eslint-disable-next-line no-unused-vars
const Cashflow = {
    refreshCashflow() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const couponsSheet = ss.getSheetByName(Config.SHEET_NAMES.COUPONS);
        const bondsSheet = ss.getSheetByName(Config.SHEET_NAMES.BONDS);
        const cashflowSheet = ss.getSheetByName(Config.SHEET_NAMES.CASHFLOW);

        // Clear existing cashflow data
        if (cashflowSheet.getLastRow() > 1) {
            cashflowSheet
                .getRange(
                    2,
                    1,
                    cashflowSheet.getLastRow() - 1,
                    Config.CASHFLOW_HEADERS.length,
                )
                .clearContent();
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
            const month = Utils.formatMonth(paymentDate);

            const gross = parseFloat(row[8]) || 0;
            const tax = parseFloat(row[9]) || 0;
            const net = parseFloat(row[10]) || 0;

            if (!monthlyMap[month]) {
                monthlyMap[month] = {
                    grossCoupon: 0,
                    tax: 0,
                    netCoupon: 0,
                    maturity: 0,
                    couponCount: 0,
                    maturityCount: 0,
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
            const month = Utils.formatMonth(maturityDate);
            const faceValue = parseFloat(row[4]) || 0;
            const quantity = parseInt(row[5]) || 0;
            const maturityAmount = faceValue * quantity;

            if (!monthlyMap[month]) {
                monthlyMap[month] = {
                    grossCoupon: 0,
                    tax: 0,
                    netCoupon: 0,
                    maturity: 0,
                    couponCount: 0,
                    maturityCount: 0,
                };
            }

            monthlyMap[month].maturity += maturityAmount;
            monthlyMap[month].maturityCount++;
        }

        // Sort by month and write
        const sortedMonths = Object.keys(monthlyMap).sort();

        const rows = sortedMonths.map((month) => {
            const m = monthlyMap[month];
            const r = (v) => Utils.bankersRound(v * 100) / 100;
            return [
                month,
                r(m.grossCoupon),
                r(m.tax),
                r(m.netCoupon),
                r(m.maturity),
                r(m.grossCoupon + m.maturity),
                r(m.netCoupon + m.maturity),
                m.couponCount,
                m.maturityCount,
            ];
        });

        if (rows.length > 0) {
            // Format
            cashflowSheet
                .getRange(2, 2, rows.length, 6)
                .setNumberFormat('#,##0.00');
            cashflowSheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');

            cashflowSheet
                .getRange(2, 1, rows.length, rows[0].length)
                .setValues(rows);

            // Add summary row
            const summaryRow = rows.length + 3;

            const formulas = [
                [
                    'TOTAL',
                    `=SUM(B2:B${rows.length + 1})`,
                    `=SUM(C2:C${rows.length + 1})`,
                    `=SUM(D2:D${rows.length + 1})`,
                    `=SUM(E2:E${rows.length + 1})`,
                    `=SUM(F2:F${rows.length + 1})`,
                    `=SUM(G2:G${rows.length + 1})`,
                    `=SUM(H2:H${rows.length + 1})`,
                    `=SUM(I2:I${rows.length + 1})`,
                ],
            ];

            const summaryRange = cashflowSheet.getRange(summaryRow, 1, 1, 9);
            summaryRange.setValues(formulas);
            summaryRange.setFontWeight('bold');

            summaryRange.offset(0, 1, 1, 6).setNumberFormat('#,##0.00');
        }

        this._applyHeatmap(cashflowSheet, 2, rows.length);
    },

    _applyHeatmap(sheet, startRow, numRows) {
        if (numRows === 0) return;

        const col = 7; // Total Net Cashflow (G)

        const range = sheet.getRange(startRow, col, numRows, 1);
        const values = range.getValues().map((r) => r[0]);

        const positive = values.filter((v) => v > 0);
        const min = positive.length ? Math.min(...positive) : 0;
        const max = positive.length ? Math.max(...values) : 0;

        const backgrounds = values.map((v) => {
            if (max === min) return ['#fff7cc']; // fallback

            const ratio = (v - min) / (max - min);

            // 🎨 Yellow → Red gradient
            let r = 255;
            let g = Math.round(255 - ratio * 180); // уменьшаем зелёный
            let b = Math.round(200 - ratio * 200); // уменьшаем синий

            return [`rgb(${r},${g},${Math.max(b, 0)})`];
        });

        range.setBackgrounds(backgrounds);
    },
};
