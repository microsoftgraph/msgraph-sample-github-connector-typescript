// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import globals from 'globals';
import js from '@eslint/js';
import tsParser from '@typescript-eslint/parser';
import eslintTypeScript from '@typescript-eslint/eslint-plugin';
import eslintPrettierRecommended from 'eslint-plugin-prettier/recommended';
import eslintTsDocPlugin from 'eslint-plugin-tsdoc';
import header from 'eslint-plugin-header';
header.rules.header.meta.schema = false;

export default [
  {
    ignores: ['**/out'],
  },
  js.configs.recommended,
  eslintPrettierRecommended,
  {
    files: ['**/*.{ts,mjs}'],

    languageOptions: {
      globals: {
        ...globals.node,
      },

      parser: tsParser,
      ecmaVersion: 'latest',
      sourceType: 'module',
    },

    plugins: {
      typeScript: eslintTypeScript,
      header,
      tsdoc: eslintTsDocPlugin,
    },

    rules: {
      'header/header': [
        'error',
        'line',
        [
          ' Copyright (c) Microsoft Corporation.',
          ' Licensed under the MIT license.',
        ],
      ],
      'prettier/prettier': [
        'error',
        {
          singleQuote: true,
          endOfLine: 'auto',
          printWidth: 80,
        },
      ],
      'no-unused-vars': 'off',
      'typeScript/no-unused-vars': [
        'warn',
        {
          argsIgnorePattern: '^_',
        },
      ],
      'tsdoc/syntax': 'warn',
    },
  },
];
