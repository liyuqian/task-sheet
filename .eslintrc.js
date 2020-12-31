module.exports = {
  parser: '@typescript-eslint/parser',
  extends: [
    'airbnb-typescript/base',
  ],
  parserOptions: {
    project: './tsconfig.json',
  },
  rules: {
    'max-len': ['error', { code: 80 }],
    '@typescript-eslint/no-use-before-define': [
      2, 
      { 'functions': false, 'classes': false },
    ]
  }
};
