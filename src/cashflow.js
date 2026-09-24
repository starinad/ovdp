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

        // FUTURE mode always shows the current month, even when its coupons
        // are already paid and nothing else remains in it.
        if (mode === 'FUTURE') {
            const currentMonth = Utils.formatMonth(today);
            if (!monthlyMap[currentMonth]) {
                monthlyMap[currentMonth] = {
                    grossCoupon: 0,
                    tax: 0,
                    netCoupon: 0,
                    maturity: 0,
                    couponCount: 0,
                    maturityCount: 0,
                };
            }
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

    // Highlights rows in the available-bonds table (L–P) whose coupon month
    // matches the month of the currently selected Cashflow data cell (A–I).
    highlightCashflowMonth(e) {
        const sheet = e.range.getSheet();
        if (sheet.getName() !== Config.SHEET_NAMES.CASHFLOW) return;
        if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

        const row = e.range.getRow();
        const col = e.range.getColumn();
        const inDataTable =
            col >= 1 && col <= Config.CASHFLOW_HEADERS.length && row >= 2;
        const value = inDataTable ? sheet.getRange(row, 1).getValue() : '';
        const selected = /^\d{4}-\d{2}$/.test(value) ? value : '';

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;

        // Column L is the first column (Coupon) of the available-bonds table.
        const bondsRange = sheet.getRange(2, 12, lastRow - 1, 5);
        const couponMonths = bondsRange.getValues().map((r) => r[0]);
        const highlight = [
            '#b6d7a8',
            '#b6d7a8',
            '#b6d7a8',
            '#b6d7a8',
            '#b6d7a8',
        ];
        const plain = [null, null, null, null, null];

        bondsRange.setBackgrounds(
            couponMonths.map((m) =>
                selected && m === selected ? highlight : plain,
            ),
        );
    },

    CF_MATURITY: 'Погашення',
    BONDS_TTL_MS: 12 * 60 * 60 * 1000, // 12h

    _refreshAvailableCouponsTable(cashflowSheet) {
        const COL_START = 12; // column L

        // Clear previous data in columns L:P (keep row 1 for header)
        const maxRows = cashflowSheet.getMaxRows();
        if (maxRows > 1) {
            cashflowSheet
                .getRange(1, COL_START, maxRows, 5)
                .clearContent()
                .setFontWeight('normal')
                .setBackground(null)
                .setNumberFormat('@')
                .clearNote();
        }

        // Write header
        const headerRange = cashflowSheet.getRange(1, COL_START, 1, 5);
        headerRange.setValues([
            ['Coupon', 'ISIN', 'Maturity', 'Rate', 'Price'],
        ]);
        headerRange.setFontWeight('bold');

        const bonds = this.getLiveBonds();

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

        // Build one row per (coupon month, bond) combination.
        // Each bond can appear multiple times — once per distinct coupon month.
        const tableRows = []; // [ [month, isin, maturityYMD], ... ]

        // Per-ISIN coupon schedule (date, value, type), source for the cell
        // note and the quantity dialog. `parsed` is the sortable Date.
        const couponsByIsin = {};
        const bondByIsin = {};
        for (const bond of bonds) {
            const isin = bond.isin;
            const maturityDate = parseDMY(bond.maturity);
            if (!maturityDate || !isin || bond.currency !== 'UAH') continue;
            if (!bond.sellPrice) continue;

            const maturityYM = toYearMonth(maturityDate);

            // Collect distinct coupon months for this bond (exclude Погашення)
            const couponMonthSet = new Set();
            for (const coupon of bond.coupons || []) {
                if (coupon.type === 'Погашення') continue;
                const pd = parseDMY(coupon.paymentDate);
                if (!pd) continue;
                couponMonthSet.add(toYearMonth(pd));
            }

            bondByIsin[isin] = bond;
            couponsByIsin[isin] = this._sortedCoupons(bond.coupons || []);

            for (const month of couponMonthSet) {
                tableRows.push([
                    month,
                    isin,
                    maturityYM,
                    bond.sellYield ? bond.sellYield + '%' : 'n/a ',
                    bond.sellPrice,
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
            .getRange(2, COL_START, tableRows.length, 5)
            .setValues(tableRows)
            .setNumberFormat('@'); // force text so dates are not auto-converted

        // Attach the coupon-schedule popup to each ISIN cell
        cashflowSheet
            .getRange(2, COL_START + 1, tableRows.length, 1)
            .setNotes(
                tableRows.map((r) => [
                    this._couponNote(bondByIsin[r[1]], couponsByIsin[r[1]]),
                ]),
            );
    },

    _parseDMY(str) {
        if (!str) return null;
        const parts = str.split('.');
        if (parts.length !== 3) return null;
        return new Date(
            parseInt(parts[2], 10),
            parseInt(parts[1], 10) - 1,
            parseInt(parts[0], 10),
        );
    },

    // Coupons sorted by date; on the same date the maturity payment goes
    // last so it reads as the final payment.
    _sortedCoupons(couponsList) {
        return (couponsList || [])
            .map((c) => ({ ...c, parsed: this._parseDMY(c.paymentDate) }))
            .filter((c) => c.parsed)
            .sort((a, b) => {
                const aT = a.parsed.getTime();
                const bT = b.parsed.getTime();
                if (aT !== bT) return aT - bT;
                const aM = a.type === this.CF_MATURITY ? 1 : 0;
                const bM = b.type === this.CF_MATURITY ? 1 : 0;
                return aM - bM;
            });
    },

    // Static note text for an ISIN cell (notes can't hold input fields).
    _couponNote(bond, coupons) {
        const lines = [
            `ISIN: ${bond.isin}`,
            `Maturity: ${bond.maturity}`,
            `Sell price: ${Utils.formatUAH(bond.sellPrice)}`,
        ];
        if (bond.sellYield) {
            lines.push(`Sell yield: ${bond.sellYield}%`);
        }

        if (coupons.length) {
            lines.push('', 'Coupons:');
            for (const c of coupons) {
                const kind = c.type === this.CF_MATURITY ? 'Погашення ' : '';
                lines.push(
                    `${c.paymentDate} — ${kind}${Utils.formatUAH(c.value)}`,
                );
            }
        } else {
            lines.push('', 'No coupons');
        }

        return lines.join('\n');
    },

    // Bond + sorted coupon schedule for the dialog.
    getBondData(isin) {
        const bond = this.getLiveBonds().find((b) => b.isin === isin);
        return {
            bond,
            coupons: bond ? this._sortedCoupons(bond.coupons || []) : [],
        };
    },

    // Live Privat24 bond catalogue: serves the Config-sheet snapshot until it
    // is older than `BONDS_TTL_MS`, then refetches and refreshes the snapshot
    // (Bonds JSON + Bonds Snapshot Timestamp cells).
    getLiveBonds() {
        const config = Config.getConfig();
        const cached = this._parseBonds(config.bondsJson);

        if (
            cached.length &&
            config.bondsSnapshot &&
            Date.now() - config.bondsSnapshot < this.BONDS_TTL_MS
        ) {
            return cached;
        }

        const fresh = this._fetchLiveBonds();
        if (fresh.length) {
            Config.setBondsSnapshot(fresh, new Date());
            return fresh;
        }

        // Fetch failed: serve whatever cached data we have rather than nothing
        return cached;
    },

    // Parses the Config snapshot (wrapper or plain array) into a plain array.
    _parseBonds(bondsJson) {
        if (typeof bondsJson === 'string') {
            try {
                bondsJson = JSON.parse(bondsJson);
            } catch {
                return [];
            }
        }
        return Array.isArray(bondsJson)
            ? bondsJson
            : (bondsJson && bondsJson.data) || [];
    },

    _fetchLiveBonds() {
        let bondsJson;
        try {
            const initResponse = UrlFetchApp.fetch(
                'https://next.privat24.ua/api/p24/init',
                {
                    method: 'post',
                    contentType: 'application/json',
                    payload: JSON.stringify({}),
                },
            );
            const initJson = JSON.parse(initResponse.getContentText());
            const xref = initJson.data && initJson.data.xref;
            if (!xref) {
                throw new Error('init returned no xref');
            }
            const allHeaders = initResponse.getAllHeaders();
            const setCookieKey = Object.keys(allHeaders).find(
                (k) => k.toLowerCase() === 'set-cookie',
            );
            const pubkey = String(
                setCookieKey ? allHeaders[setCookieKey] : '',
            ).split(';')[0];

            const response = UrlFetchApp.fetch(
                'https://next.privat24.ua/api/p24/pub/bonds',
                {
                    method: 'post',
                    contentType: 'application/json',
                    headers: { Cookie: pubkey },
                    payload: JSON.stringify({
                        action: 'bargaining',
                        xref,
                        _: Date.now(),
                    }),
                },
            );
            bondsJson = JSON.parse(response.getContentText());
        } catch (e) {
            Logger.log('getLiveBonds: failed to fetch bonds – ' + e);
            return [];
        }

        return this._parseBonds(bondsJson);
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
