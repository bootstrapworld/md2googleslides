const gts = require('gts');

module.exports = [
  {ignores: ['lib/', 'build/', '**/node_modules/', 'test/fixtures/']},
  ...gts,
  {
    rules: {
      '@typescript-eslint/ban-ts-comment': 'off',
    },
  },
  {
    // plain JS entry points (not covered by the TypeScript globals)
    files: ['bin/*.js', 'index.js'],
    languageOptions: {
      globals: {
        console: 'readonly',
        process: 'readonly',
        require: 'readonly',
        module: 'writable',
      },
    },
  },
];
