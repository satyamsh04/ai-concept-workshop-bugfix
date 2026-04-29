module.exports = {
  testEnvironment: 'jsdom',
  setupFilesAfterEnv: ['<rootDir>/tests/setup.js'],
  testMatch: ['**/tests/**/*.test.js'],
  collectCoverageFrom: [
    'src/**/*.js',
    '!node_modules/**'
  ],
  moduleNameMapper: {
    '\.(css|less|scss|sass)$': '<rootDir>/tests/__mocks__/styleMock.js'
  }
};
