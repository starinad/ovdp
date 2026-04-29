// eslint-disable-next-line no-unused-vars
const Analytics = {
    refreshAnalytics() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bondsSheet = ss.getSheetByName(Config.SHEET_NAMES.BONDS);
        const couponsSheet = ss.getSheetByName(Config.SHEET_NAMES.COUPONS);
        const analyticsSheet = ss.getSheetByName(Config.SHEET_NAMES.ANALYTICS);
        const ladderSheet = ss.getSheetByName(Config.SHEET_NAMES.LADDER);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const maturitiesThisMonth = [];
        let totalMaturitiesThisMonth = 0;
        let bonds = Bonds.getBondsData(bondsSheet);
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

            if (bond.status === 'ACTIVE' || bond.status === 'MATURED') {
                const maturityDate = Utils.normalizeDate(bond.maturityDate);
                const sameMonth =
                    maturityDate.getFullYear() === today.getFullYear() &&
                    maturityDate.getMonth() === today.getMonth();

                if (sameMonth) {
                    totalMaturitiesThisMonth += bond.faceValue * bond.quantity;

                    maturitiesThisMonth.push({
                        date: maturityDate,
                        net: bond.faceValue * bond.quantity,
                    });
                }
            }
        });
        if (hasChanges) {
            statusRange.setValues(statuses);
        }

        const activeBonds = bonds.filter((b) => b.status === 'ACTIVE');
        const couponsData = couponsSheet.getDataRange().getValues();

        // ── Compute metrics ──

        let totalInvested = 0;
        let totalFaceValue = 0;
        let weightedRateSum = 0;
        let totalGrossCouponIncome = 0;
        let totalNetCouponIncome = 0;
        let totalScheduledGross = 0;
        let totalScheduledNet = 0;
        const сouponsThisMonth = [];
        let totalCouponsThisMonth = 0;

        const bondMap = {};
        activeBonds.forEach((bond) => {
            bondMap[bond.id] = bond;
            const invested =
                (bond.purchasePrice + bond.accruedInterest) * bond.quantity;
            const face = bond.faceValue * bond.quantity;
            totalInvested += invested;
            totalFaceValue += face;
            weightedRateSum += invested * bond.interestRate;
        });

        const weightedAvgYield =
            totalInvested > 0
                ? Utils.bankersRound((weightedRateSum / totalInvested) * 100) /
                  100
                : 0;

        // Process coupons for income calculation
        for (let i = 1; i < couponsData.length; i++) {
            const row = couponsData[i];
            if (!row[4]) continue;
            const status = row[12];
            const payDate = Utils.normalizeDate(row[4]);
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

                сouponsThisMonth.push({
                    date: payDate,
                    net: net,
                });
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
            }
        }
        const r = (v) => Utils.bankersRound(v * 100) / 100;
        сouponsThisMonth.sort((a, b) => a.date - b.date);
        const сouponsThisMonthText =
            сouponsThisMonth.length > 0
                ? сouponsThisMonth
                      .map(
                          (c) =>
                              Utilities.formatDate(
                                  c.date,
                                  Session.getScriptTimeZone(),
                                  'yyyy-MM-dd',
                              ) +
                              ' → ' +
                              Utils.formatUAH(r(c.net)),
                      )
                      .join('\n')
                : 'N/A';

        maturitiesThisMonth.sort((a, b) => a.date - b.date);
        const maturitiesThisMonthText =
            maturitiesThisMonth.length > 0
                ? maturitiesThisMonth
                      .map(
                          (c) =>
                              Utilities.formatDate(
                                  c.date,
                                  Session.getScriptTimeZone(),
                                  'yyyy-MM-dd',
                              ) +
                              ' → ' +
                              Utils.formatUAH(r(c.net)),
                      )
                      .join('\n')
                : 'N/A';

        // Yearly projected income (from scheduled coupons)
        // Annualized: total scheduled income / years from TODAY to last maturity
        // This answers: "how much income per year does my current portfolio generate
        // over its remaining lifetime?"
        let yearlyGross = 0;
        let yearlyNet = 0;

        if (activeBonds.length > 0 && totalScheduledGross > 0) {
            // Find the latest maturity across all active bonds
            let latestMaturity = activeBonds[0].maturityDate;
            activeBonds.forEach((b) => {
                if (b.maturityDate > latestMaturity)
                    latestMaturity = b.maturityDate;
            });

            const yearsRemaining =
                Utils.daysBetween(today, latestMaturity) / 365;

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
            activeBonds.forEach((b) => {
                if (b.maturityDate > latestMaturity)
                    latestMaturity = b.maturityDate;
            });
            yearsRemaining = Utils.daysBetween(today, latestMaturity) / 365;
        }

        const totalReturn =
            totalScheduledGross + (totalFaceValue - totalInvested);

        const cashflows = this._buildCashflows(bonds, couponsData);

        let portfolioXirr = null;

        if (cashflows.length > 1) {
            try {
                portfolioXirr = this._xirr(cashflows);
                // eslint-disable-next-line no-unused-vars
            } catch (e) {
                portfolioXirr = null;
            }
        }

        // ── Write analytics ──

        if (analyticsSheet.getLastRow() > 1) {
            analyticsSheet
                .getRange(2, 1, analyticsSheet.getLastRow() - 1, 2)
                .clearContent();
        }

        const metrics = [
            ['', ''],
            ['── PORTFOLIO OVERVIEW ──', ''],
            ['Active Bonds', activeBonds.length],
            [
                'Matured Bonds',
                bonds.filter((b) => b.status === 'MATURED').length,
            ],
            [
                'Total Invested Capital (incl. accrued)',
                Utils.formatUAH(r(totalInvested)),
            ],
            ['Total Face Value', Utils.formatUAH(r(totalFaceValue))],
            [
                'Unrealized Capital Gain/Loss',
                Utils.formatUAH(r(totalFaceValue - totalInvested)),
            ],
            ['Weighted Average Yield (%)', weightedAvgYield.toFixed(2) + '%'],
            [
                'Portfolio XIRR (%)',
                portfolioXirr !== null
                    ? (portfolioXirr * 100).toFixed(2) + '%'
                    : 'N/A',
            ],
            [
                'Portfolio Horizon (to last maturity)',
                r(yearsRemaining).toFixed(2) + ' years',
            ],
            ['', ''],
            ['── INCOME (REALIZED) ──', ''],
            [
                'Total Gross Coupon Income Received',
                Utils.formatUAH(r(totalGrossCouponIncome)),
            ],
            [
                'Total Tax Paid',
                Utils.formatUAH(
                    r(totalGrossCouponIncome - totalNetCouponIncome),
                ),
            ],
            [
                'Total Net Coupon Income Received',
                Utils.formatUAH(r(totalNetCouponIncome)),
            ],
            ['', ''],
            ['── INCOME (PROJECTED) ──', ''],
            [
                'Total Scheduled Gross Income',
                Utils.formatUAH(r(totalScheduledGross)),
            ],
            [
                'Total Scheduled Net Income',
                Utils.formatUAH(r(totalScheduledNet)),
            ],
            [
                'Annualized Gross Income *',
                yearsRemaining >= 0.5
                    ? Utils.formatUAH(r(yearlyGross))
                    : 'N/A (horizon < 6 months)',
            ],
            [
                'Annualized Net Income *',
                yearsRemaining >= 0.5
                    ? Utils.formatUAH(r(yearlyNet))
                    : 'N/A (horizon < 6 months)',
            ],
            [
                'Monthly Avg Gross Income *',
                yearsRemaining >= 0.5
                    ? Utils.formatUAH(r(yearlyGross / 12))
                    : 'N/A (horizon < 6 months)',
            ],
            [
                'Monthly Avg Net Income *',
                yearsRemaining >= 0.5
                    ? Utils.formatUAH(r(yearlyNet / 12))
                    : 'N/A (horizon < 6 months)',
            ],
            [
                '',
                '* Annualized = total scheduled income ÷ years to last maturity.',
            ],
            [
                '',
                '  Not shown when horizon < 6 months (misleading on short bonds).',
            ],
            ['', '  Add longer-dated bonds for this metric to be useful.'],
            ['', ''],
            ['── TOTAL RETURN (PROJECTED) ──', ''],
            [
                'Total Coupon Income (scheduled)',
                Utils.formatUAH(r(totalScheduledGross)),
            ],
            [
                'Capital Gain/Loss at Maturity',
                Utils.formatUAH(r(totalFaceValue - totalInvested)),
            ],
            ['Total Projected Return', Utils.formatUAH(r(totalReturn))],
            [
                'Return on Investment (%)',
                totalInvested > 0
                    ? r((totalReturn / totalInvested) * 100).toFixed(2) + '%'
                    : 'N/A',
            ],
            ['', ''],
            ['── UPCOMING ──', ''],
            ['Coupons This Month', сouponsThisMonthText],
            [
                'Total Coupons This Month',
                Utils.formatUAH(r(totalCouponsThisMonth)),
            ],
            ['Maturities This Month', maturitiesThisMonthText],
            [
                'Total Maturities This Month',
                Utils.formatUAH(r(totalMaturitiesThisMonth)),
            ],
            ['', ''],
            [
                'Last Updated',
                Utilities.formatDate(
                    new Date(),
                    Session.getScriptTimeZone(),
                    'yyyy-MM-dd HH:mm:ss',
                ),
            ],
        ];

        analyticsSheet.getRange(2, 1, metrics.length, 2).setValues(metrics);

        // Bold section headers
        metrics.forEach((row, i) => {
            if (row[0].startsWith('──')) {
                analyticsSheet
                    .getRange(i + 2, 1)
                    .setFontWeight('bold')
                    .setFontColor('#1a73e8');
            }
        });

        // ── Maturity Ladder ──
        Ladder.refreshLadder(ladderSheet, activeBonds, today);
    },

    _buildCashflows(bonds, couponsData) {
        const today = Utils.normalizeDate(new Date());
        const flows = [];

        // ❌ sell
        bonds.forEach((bond) => {
            const invested =
                (bond.purchasePrice + bond.accruedInterest) * bond.quantity;

            flows.push({
                date: bond.purchaseDate,
                amount: -invested,
            });

            if (
                Utils.normalizeDate(bond.maturityDate) >= today ||
                bond.status !== 'SOLD'
            ) {
                flows.push({
                    date: bond.maturityDate,
                    amount: bond.faceValue * bond.quantity,
                });
            }
        });

        // ✅ buy
        for (let i = 1; i < couponsData.length; i++) {
            const row = couponsData[i];
            const status = row[12];
            if (status === 'CANCELLED') continue;

            const date = Utils.normalizeDate(row[4]);
            const net = parseFloat(row[10]) || 0;

            flows.push({
                date,
                amount: net,
            });
        }

        return flows.sort((a, b) => a.date - b.date);
    },

    _xirr(cashflows, guess = 0.1) {
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
    },
};
