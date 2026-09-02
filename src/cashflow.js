// eslint-disable-next-line no-unused-vars
const Cashflow = {
    refreshCashflow(mode = 'ALL') {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const couponsSheet = ss.getSheetByName(Config.SHEET_NAMES.COUPONS);
        const bondsSheet = ss.getSheetByName(Config.SHEET_NAMES.BONDS);
        const cashflowSheet = ss.getSheetByName(Config.SHEET_NAMES.CASHFLOW);

        // Clear existing cashflow data
        if (cashflowSheet.getMaxRows() > 1) {
            cashflowSheet
                .getRange(
                    2,
                    1,
                    cashflowSheet.getMaxRows() - 1,
                    Config.CASHFLOW_HEADERS.length,
                )
                .clearContent()
                .setFontWeight('normal')
                .setFontStyle('normal')
                .setBackground(null);
        }

        const couponsData = couponsSheet.getDataRange().getValues();
        const bondsData = bondsSheet.getDataRange().getValues();
        const today = Utils.normalizeDate(new Date());

        // Aggregate coupons by month
        const monthlyMap = {};

        // Process coupons (skip cancelled)
        for (let i = 1; i < couponsData.length; i++) {
            const row = couponsData[i];
            const status = row[12];
            if (status === 'CANCELLED' || !row[4]) continue;

            const paymentDate = Utils.normalizeDate(new Date(row[4]));
            const month = Utils.formatMonth(paymentDate);

            const showCashflow =
                mode === 'ALL' ||
                (mode === 'FUTURE' && paymentDate >= today) ||
                (mode === 'REALIZED' && paymentDate < today);

            if (!showCashflow) continue;

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

            const maturityDate = Utils.normalizeDate(new Date(row[12]));

            const showCashflow =
                mode === 'ALL' ||
                (mode === 'FUTURE' && maturityDate >= today) ||
                (mode === 'REALIZED' && maturityDate < today);

            if (!showCashflow) continue;

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

        this._applyHeatmap(cashflowSheet, 4, 2, rows.length);
        this._applyHeatmap(cashflowSheet, 7, 2, rows.length);

        // Build the available-bonds coupon opportunity table in columns L, M, N
        this._refreshAvailableCouponsTable(cashflowSheet);
    },

    _refreshAvailableCouponsTable(cashflowSheet) {
        const COL_START = 12; // column L

        // Clear previous data in columns L:N (keep row 1 for header)
        const maxRows = cashflowSheet.getMaxRows();
        if (maxRows > 1) {
            cashflowSheet
                .getRange(1, COL_START, maxRows, 3)
                .clearContent()
                .setFontWeight('normal')
                .setBackground(null)
                .setNumberFormat('@');
        }

        // Write header
        const headerRange = cashflowSheet.getRange(1, COL_START, 1, 3);
        headerRange.setValues([['Month', 'ISIN', 'Maturity']]);
        headerRange.setFontWeight('bold');

        // Load bond catalogue from config
        let bondsJson;
        try {
            bondsJson = Config.getConfig().bondsJson;
            if (typeof bondsJson === 'string') {
                bondsJson = JSON.parse(bondsJson);
            }
        } catch (e) {
            Logger.log(
                '_refreshAvailableCouponsTable: failed to parse bondsJson – ' +
                    e,
            );
            return;
        }

        // bondsJson may be a wrapper object with a `data` array (matches the
        // example payload) or a plain array of bond objects.
        const bonds = Array.isArray(bondsJson)
            ? bondsJson
            : bondsJson.data || [];

        if (!bonds.length) return;

        // Helper: parse "DD.MM.YYYY" → Date
        const parseDMY = (str) => {
            if (!str) return null;
            const parts = str.split('.');
            if (parts.length !== 3) return null;
            return new Date(
                parseInt(parts[2], 10),
                parseInt(parts[1], 10) - 1,
                parseInt(parts[0], 10),
            );
        };

        // Helper: format Date → "YYYY-MM"
        const toYearMonth = (date) => {
            const y = date.getFullYear();
            const m = String(date.getMonth() + 1).padStart(2, '0');
            return `${y}-${m}`;
        };

        // Helper: format Date → "YYYY-MM-DD"
        const toYMD = (date) => {
            const y = date.getFullYear();
            const m = String(date.getMonth() + 1).padStart(2, '0');
            const d = String(date.getDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
        };

        // Build one row per (coupon month, bond) combination.
        // Each bond can appear multiple times — once per distinct coupon month.
        const tableRows = []; // [ [month, isin, maturityYMD], ... ]

        for (const bond of bonds) {
            const isin = bond.isin;
            const maturityDate = parseDMY(bond.maturity);
            if (!maturityDate || !isin || bond.currency !== 'UAH') continue;

            const maturityYMD = toYMD(maturityDate);

            // Collect distinct coupon months for this bond (exclude Погашення)
            const couponMonthSet = new Set();
            for (const coupon of bond.coupons || []) {
                if (coupon.type === 'Погашення') continue;
                const pd = parseDMY(coupon.paymentDate);
                if (!pd) continue;
                couponMonthSet.add(toYearMonth(pd));
            }

            for (const month of couponMonthSet) {
                tableRows.push([
                    month,
                    isin,
                    maturityYMD,
                    bond.sellYield ? bond.sellYield + '%' : 'n/a ',
                ]);
            }
        }

        if (!tableRows.length) return;

        // Sort: primary = Month (lexicographic YYYY-MM), secondary = Maturity (YYYY-MM-DD)
        tableRows.sort((a, b) => {
            if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
            return a[2] < b[2] ? -1 : 1;
        });

        // Write to sheet starting at L2
        cashflowSheet
            .getRange(2, COL_START, tableRows.length, 4)
            .setValues(tableRows)
            .setNumberFormat('@'); // force text so dates are not auto-converted
    },

    _applyHeatmap(sheet, col, startRow, numRows) {
        if (numRows === 0) return;

        const range = sheet.getRange(startRow, col, numRows, 1);
        const values = range.getValues().map((r) => r[0]);

        const positive = values.filter((v) => v > 0);

        const min = positive.length ? Math.min(...positive) : 0;
        const max = positive.length ? Math.max(...positive) : 0;
        const lg = (x) => Math.log(x + 1);
        const backgrounds = values.map((v) => {
            if (max === min) return ['#fff7cc'];

            const ratio =
                max === min ? 0 : (lg(v) - lg(min)) / (lg(max) - lg(min));

            let r = 255;
            let g = Math.round(255 - ratio * 180);
            let b = Math.round(200 - ratio * 200);

            return [`rgb(${r},${g},${Math.max(b, 0)})`];
        });

        range.setBackgrounds(backgrounds);
    },
};
