import baseConfig from 'eslint-config-sdm';
import gas from 'eslint-plugin-googleappsscript';

export default [
    ...baseConfig,
    {
        languageOptions: {
            globals: {
                SpreadsheetApp: 'readonly',
                Logger: 'readonly',
                Utilities: 'readonly',
            },
        },
    },
];
