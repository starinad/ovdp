const Config = {
    SHEET_NAMES: {
        BONDS: 'Bonds',
        COUPONS: 'Coupons',
        CASHFLOW: 'Cashflow',
        ANALYTICS: 'Analytics',
        LADDER: 'Ladder',
        CONFIG: 'Config',
    },

    BOND_HEADERS: [
        { header: 'ID', width: 80 },
        { header: 'ISIN', width: 140 },
        { header: 'Name', width: 200 },
        { header: 'Status', width: 90 },
        { header: 'Face Value (UAH)', width: 130 },
        { header: 'Quantity', width: 80 },
        { header: 'Purchase Price (UAH)', width: 150 },
        { header: 'Accrued Interest at Purchase (UAH)', width: 200 },
        { header: 'Interest Rate (%)', width: 120 },
        { header: 'Tax Rate (%)', width: 100 },
        { header: 'Currency', width: 80 },
        { header: 'Purchase Date', width: 120 },
        { header: 'Maturity Date', width: 120 },
        { header: 'First Coupon Date', width: 130 },
        { header: 'Coupon Frequency', width: 130 },
        { header: 'Day Count Convention', width: 150 },
        { header: 'Fixed Coupon (UAH/unit)', width: 130 },
        { header: 'Total Invested', width: 130 },
        { header: 'Total Face Value', width: 130 },
        { header: 'Notes', width: 200 },
        { header: 'Coupons Generated', width: 120 },
        { header: 'Last Updated', width: 150 },
    ],

    COUPON_HEADERS: [
        { header: 'Bond ID', width: 80 },
        { header: 'ISIN', width: 140 },
        { header: 'Bond Name', width: 200 },
        { header: 'Seq #', width: 80 },
        { header: 'Payment Date', width: 120 },
        { header: 'Period Start', width: 120 },
        { header: 'Period End', width: 120 },
        { header: 'Accrued Days', width: 120 },
        { header: 'Gross Amount (UAH)', width: 140 },
        { header: 'Tax (UAH)', width: 100 },
        { header: 'Net Amount (UAH)', width: 140 },
        { header: 'Day Count', width: 120 },
        { header: 'Status', width: 120 },
        { header: 'Is First', width: 120 },
        { header: 'Is Last', width: 120 },
    ],

    CASHFLOW_HEADERS: [
        { header: 'Month', width: 120 },
        { header: 'Gross Coupon Income', width: 160 },
        { header: 'Tax on Coupons', width: 130 },
        { header: 'Net Coupon Income', width: 150 },
        { header: 'Maturity Payments', width: 150 },
        { header: 'Total Gross Cashflow', width: 160 },
        { header: 'Total Net Cashflow', width: 160 },
        { header: 'Coupon Count', width: 120 },
        { header: 'Maturity Count', width: 120 },
    ],
    ANALYTICS_HEADERS: [
        { header: 'Metric', width: 280 },
        { header: 'Value', width: 200 },
    ],
    LADDER_HEADERS: [
        { header: 'Maturity Bucket', width: 140 },
        { header: 'Bond Count', width: 110 },
        { header: 'Face Value (UAH)', width: 140 },
        { header: 'Percentage (%)', width: 120 },
        { header: 'Bonds (ISINs)', width: 300 },
    ],
    CONFIG_HEADERS: [
        { header: 'Setting', width: 200 },
        { header: 'Value', width: 150 },
        { header: 'Description', width: 400 },
    ],

    getConfig() {
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
            if (key.includes('Bonds JSON'))
                config.bondsJson = JSON.parse(value) || { data: [] };
        }

        return config;
    },
};
