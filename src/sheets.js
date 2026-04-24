// eslint-disable-next-line no-unused-vars
const Sheets = {
    setupSheet() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        this._createOrGetSheet(
            ss,
            Config.SHEET_NAMES.BONDS,
            Config.BOND_HEADERS,
        );

        this._createOrGetSheet(
            ss,
            Config.SHEET_NAMES.COUPONS,
            Config.COUPON_HEADERS,
        );

        this._createOrGetSheet(
            ss,
            Config.SHEET_NAMES.CASHFLOW,
            Config.CASHFLOW_HEADERS,
        );

        this._createOrGetSheet(
            ss,
            Config.SHEET_NAMES.ANALYTICS,
            Config.ANALYTICS_HEADERS,
        );

        this._createOrGetSheet(
            ss,
            Config.SHEET_NAMES.LADDER,
            Config.LADDER_HEADERS,
        );

        const configSheet = this._createOrGetSheet(
            ss,
            Config.SHEET_NAMES.CONFIG,
            Config.CONFIG_HEADERS,
        );

        const configData = configSheet.getDataRange().getValues();
        if (configData.length <= 1) {
            configSheet.getRange(2, 1, 4, 3).setValues([
                [
                    'Default Tax Rate (%)',
                    0,
                    'Applied to new bonds if not specified (0 = tax-exempt)',
                ],
                ['Default Day Count', 'ACT/365', 'ACT/365, ACT/ACT, or 30/360'],
                [
                    'Default Coupon Frequency',
                    'Semi-Annual',
                    'Monthly, Quarterly, Semi-Annual, Annual',
                ],
                ['Default Currency', 'UAH', 'Currency code'],
            ]);
        }

        // Activate bonds sheet
        ss.setActiveSheet(ss.getSheetByName(Config.SHEET_NAMES.BONDS));

        SpreadsheetApp.getUi().alert(
            '✅ Setup Complete',
            'OVDP Manager is ready. Use the "💰 OVDP Manager" menu to add bonds.',
            SpreadsheetApp.getUi().ButtonSet.OK,
        );
    },

    _createOrGetSheet(ss, name, headers) {
        let sheet = ss.getSheetByName(name);
        if (!sheet) {
            sheet = ss.insertSheet(name);
        }

        // Set headers
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setValues([headers.map(({ header }) => header)]);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#1a73e8');
        headerRange.setFontColor('#ffffff');
        headerRange.setHorizontalAlignment('center');
        sheet.setFrozenRows(1);

        // Set column widths
        headers.forEach((h, i) => {
            if (h.width) {
                sheet.setColumnWidth(i + 1, h.width);
            }
        });

        return sheet;
    },
};
