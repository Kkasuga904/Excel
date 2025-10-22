/** @type {import("jest").Config} */
module.exports = {
  testEnvironment: 'node',
  roots: ['<rootDir>/tests'],
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],
  collectCoverageFrom: ['src/server/**/*.{ts,tsx}'],
  coverageDirectory: 'coverage/server',
  setupFilesAfterEnv: [],
  verbose: false,
  transform: {
    '^.+\\.(ts|tsx)$': ['<rootDir>/node_modules/ts-jest/dist/index.js', { tsconfig: 'tsconfig.json' }]
  },
  moduleNameMapper: {
    '^\\.\\./src/server/(.*)$': '<rootDir>/src/server/$1.ts'
  }
};
