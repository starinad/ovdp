// eslint-disable-next-line no-unused-vars
const Coupons = {
    FREQUENCIES: {
        Monthly: 1,
        Quarterly: 3,
        'Semi-Annual': 6,
        Annual: 12,
        'Zero-Coupon': 0,
    },

    generateCouponSchedule(input) {
        const freq = this.FREQUENCIES[input.couponFrequency];
        if (!freq || freq === 0) return [];

        // Build the bond's FULL coupon date schedule (independent of purchase)
        const allCouponDates = this._buildBondCouponDates(
            input.maturityDate,
            input.firstCouponDate,
            freq,
        );

        if (allCouponDates.length === 0) return [];

        // Filter to coupons the holder receives (payment date > purchase date)
        // and build coupon objects
        const coupons = [];

        for (let i = 0; i < allCouponDates.length; i++) {
            const paymentDate = allCouponDates[i];

            // Skip coupons that paid before or on purchase date
            if (paymentDate <= input.purchaseDate) continue;

            // Period start = previous coupon date (or bond issue anchor if first)
            const periodStart =
                i > 0
                    ? allCouponDates[i - 1]
                    : Utils.addMonthsSafe(allCouponDates[0], -freq);
            const periodEnd = paymentDate;

            const yf = Utils.yearFraction(
                periodStart,
                periodEnd,
                input.dayCountConvention,
            );

            if (yf.days <= 0) continue;

            let grossAmount;

            if (input.fixedCouponPerUnit > 0) {
                // FIXED COUPON MODE
                // Regular coupon = fixedCouponPerUnit × quantity
                // For the FULL period (regardless of when holder bought — the holder
                // gets the full coupon; accrued interest at purchase compensates the seller)
                grossAmount =
                    Utils.bankersRound(
                        input.fixedCouponPerUnit * input.quantity * 100,
                    ) / 100;
            } else {
                // CALCULATED MODE
                grossAmount = this._calculateCouponAmount(
                    input.faceValue,
                    input.quantity,
                    input.interestRate,
                    yf.days,
                    yf.divisor,
                );
            }

            const taxAmount = this._calculateTax(grossAmount, input.taxRate);
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
    },

    writeCoupons(ss, bondId, isin, bondName, coupons) {
        const sheet = ss.getSheetByName(Config.SHEET_NAMES.COUPONS);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const rows = coupons.map((c) => {
            const paymentDate = new Date(c.paymentDate);
            paymentDate.setHours(0, 0, 0, 0);

            const status = paymentDate <= today ? 'PAID' : 'SCHEDULED';

            return [
                bondId,
                isin,
                bondName,
                c.sequenceNumber,
                c.paymentDate,
                c.periodStart,
                c.periodEnd,
                c.accruedDays,
                c.grossAmount,
                c.taxAmount,
                c.netAmount,
                c.dayCount,
                status,
                c.isFirst ? 'YES' : '',
                c.isLast ? 'YES' : '',
            ];
        });

        if (rows.length > 0) {
            const startRow = sheet.getLastRow() + 1;

            const range = sheet.getRange(
                startRow,
                1,
                rows.length,
                rows[0].length,
            );
            range.setValues(rows);

            // Batch formatting 🚀
            range.offset(0, 4, rows.length, 3).setNumberFormat('yyyy-mm-dd');

            range.offset(0, 8, rows.length, 3).setNumberFormat('#,##0.00');
        }
    },

    regenerateAllCoupons() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bondsSheet = ss.getSheetByName(Config.SHEET_NAMES.BONDS);
        const couponsSheet = ss.getSheetByName(Config.SHEET_NAMES.COUPONS);

        // Clear existing coupons (keep header)
        if (couponsSheet.getLastRow() > 1) {
            couponsSheet
                .getRange(
                    2,
                    1,
                    couponsSheet.getLastRow() - 1,
                    Config.COUPON_HEADERS.length,
                )
                .clearContent();
        }

        const bonds = Bonds.getBondsData(bondsSheet);

        let totalCoupons = 0;

        bonds.forEach((bond) => {
            if (bond.status === 'SOLD') return;

            const coupons = this.generateCouponSchedule({
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
                this.writeCoupons(ss, bond.id, bond.isin, bond.name, coupons);
                totalCoupons += coupons.length;
            }

            // Update coupon count on bond row
            bondsSheet
                .getRange(
                    bond.rowIndex,
                    Utils.getColumnIndex(
                        Config.BOND_HEADERS,
                        'Coupons Generated',
                    ),
                )
                .setValue(coupons.length);
            bondsSheet
                .getRange(
                    bond.rowIndex,
                    Utils.getColumnIndex(Config.BOND_HEADERS, 'Last Updated'),
                )
                .setValue(new Date());
        });

        SpreadsheetApp.getActiveSpreadsheet().toast(
            `Generated ${totalCoupons} coupons for ${bonds.length} bonds`,
            '✅ Coupons Regenerated',
            5,
        );
    },

    /**
     * Build the bond's complete coupon date schedule from issuance to maturity.
     *
     * The key insight for OVDP: the coupon schedule belongs to the BOND, not to the holder.
     * We build the full schedule, then filter for the holder's purchase date separately.
     *
     * If firstCouponDate is provided, we build forward from it.
     * The maturity date is always the last date (may coincide with a regular coupon or be a stub).
     */
    _buildBondCouponDates(maturityDate, firstCouponDate, monthsBetween) {
        if (!firstCouponDate) {
            // No first coupon date: walk backwards from maturity to build regular schedule
            const dates = [new Date(maturityDate)];
            let current = maturityDate;
            while (true) {
                current = Utils.addMonthsSafe(current, -monthsBetween);
                if (current.getTime() <= new Date(2000, 0, 1).getTime()) break; // sanity limit
                dates.unshift(new Date(current));
            }
            return dates;
        }

        // First coupon date provided: build forward from it
        const dates = [];
        let i = 0;
        while (true) {
            const date = Utils.addMonthsSafe(
                firstCouponDate,
                monthsBetween * i,
            );
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
            const gapDays = Utils.daysBetween(lastDate, maturityDate);
            if (gapDays > 7) {
                // Real stub period — add it
                dates.push(new Date(maturityDate));
            } else if (gapDays > 0) {
                // Tiny gap (1-7 days) — move last coupon to maturity date instead
                dates[dates.length - 1] = new Date(maturityDate);
            }
        }

        return dates;
    },

    /**
     * Calculate coupon amount in UAH (2 decimal places).
     * faceValue and purchasePrice are in UAH (not kopecks — this is Sheets, not the backend).
     */
    _calculateCouponAmount(faceValue, quantity, ratePercent, days, divisor) {
        const amount =
            (faceValue * quantity * ratePercent * days) / (divisor * 100);
        return Utils.bankersRound(amount * 100) / 100; // round to 2 decimals
    },

    _calculateTax(grossAmount, taxRatePercent) {
        if (taxRatePercent <= 0) return 0;
        const tax = (grossAmount * taxRatePercent) / 100;
        return Utils.bankersRound(tax * 100) / 100;
    },
};
