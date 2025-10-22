/** @type {import("jest").Config} */
module.exports = {
  testEnvironment: 'node',
  roots: ['<rootDir>/tests'],
  moduleFileExtensions: ['js', 'json', 'node'],
  collectCoverageFrom: ['dist/server/**/*.{js}'],
  coverageDirectory: 'coverage/server',
  setupFilesAfterEnv: [],
  verbose: false
};
