/** @type {import('jest').Config} */
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/tests'],
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],
  collectCoverageFrom: ['src/server/**/*.{ts,tsx}'],
  coverageDirectory: 'coverage/server',
  setupFilesAfterEnv: [],
  verbose: false
};
