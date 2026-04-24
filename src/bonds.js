// eslint-disable-next-line no-unused-vars
const Bonds = {
    addBond(bond) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bondsSheet = ss.getSheetByName(Config.SHEET_NAMES.BONDS);

        const id = this._getNextBondId(bondsSheet);
        const purchaseDate = new Date(bond.purchaseDate);
        const maturityDate = new Date(bond.maturityDate);
        const firstCouponDate = bond.firstCouponDate
            ? new Date(bond.firstCouponDate)
            : '';
        const totalInvested =
            (bond.purchasePrice + bond.accruedInterest) * bond.quantity;
        const totalFace = bond.faceValue * bond.quantity;
        const fixedCoupon = bond.fixedCoupon || 0;

        // Generate coupons
        const coupons = Coupons.generateCouponSchedule({
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
            id,
            bond.isin,
            bond.name || '',
            'ACTIVE',
            bond.faceValue,
            bond.quantity,
            bond.purchasePrice,
            bond.accruedInterest,
            bond.interestRate,
            bond.taxRate,
            bond.currency,
            purchaseDate,
            maturityDate,
            firstCouponDate,
            bond.couponFrequency,
            bond.dayCount,
            fixedCoupon,
            totalInvested,
            totalFace,
            bond.notes || '',
            coupons.length,
            new Date(),
        ];

        bondsSheet.appendRow(bondRow);

        // Format the new row
        const lastRow = bondsSheet.getLastRow();
        this._formatBondRow(bondsSheet, lastRow);

        // Write coupons
        if (coupons.length > 0) {
            Coupons.writeCoupons(ss, id, bond.isin, bond.name || '', coupons);
        }

        // Refresh computed sheets
        Cashflow.refreshCashflow();
        Analytics.refreshAnalytics();

        SpreadsheetApp.getActiveSpreadsheet().toast(
            `Added ${bond.isin} with ${coupons.length} coupons`,
            '✅ Bond Added',
            5,
        );
    },

    deleteBond(bondId) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bondsSheet = ss.getSheetByName(Config.SHEET_NAMES.BONDS);
        const couponsSheet = ss.getSheetByName(Config.SHEET_NAMES.COUPONS);

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

        Cashflow.refreshCashflow();
        Analytics.refreshAnalytics();

        ss.toast(`Bond ${bondId} and its coupons deleted.`, '🗑️ Deleted', 5);
    },

    getBondsData(bondsSheet) {
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
                purchaseDate: Utils.normalizeDate(row[11]),
                maturityDate: Utils.normalizeDate(row[12]),
                firstCouponDate: Utils.normalizeDate(row[13]),
                couponFrequency: row[14],
                dayCount: row[15],
                fixedCoupon: parseFloat(row[16]) || 0,
            });
        }

        return bonds;
    },

    _getNextBondId(bondsSheet) {
        const data = bondsSheet.getDataRange().getValues();
        let maxId = 0;
        for (let i = 1; i < data.length; i++) {
            const id = parseInt(data[i][0]);
            if (!isNaN(id) && id > maxId) maxId = id;
        }
        return maxId + 1;
    },

    // TODO: format should be defined in config and applied to whole column, not just new row
    _formatBondRow(sheet, row) {
        // Currency format for money columns
        const moneyFormat = '#,##0.00';
        [5, 7, 8, 17, 18, 19].forEach((col) => {
            sheet.getRange(row, col).setNumberFormat(moneyFormat);
        });
        // Percentage format
        sheet.getRange(row, 9).setNumberFormat('0.00');
        sheet.getRange(row, 10).setNumberFormat('0.00');
        // Date format
        [12, 13, 14].forEach((col) => {
            sheet.getRange(row, col).setNumberFormat('yyyy-mm-dd');
        });
        sheet.getRange(row, 22).setNumberFormat('yyyy-mm-dd hh:mm');
    },
};
