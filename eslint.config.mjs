import baseConfig from 'eslint-config-sdm';

export default [
    ...baseConfig,
    {
        languageOptions: {
            globals: {
                SpreadsheetApp: 'readonly',
                HtmlService: 'readonly',
                Utilities: 'readonly',
                Session: 'readonly',
                UI: 'readonly',
                Sheets: 'readonly',
                Config: 'readonly',
                Bonds: 'readonly',
                Analytics: 'readonly',
                Cashflow: 'readonly',
                Coupons: 'readonly',
                Utils: 'readonly',
                Ladder: 'readonly',
            },
        },
    },
    {
        files: ['src/main.js'],
        rules: {
            'no-unused-vars': 'off',
        },
    },
];
