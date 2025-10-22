const { createTransformer } = require('ts-jest');
module.exports = createTransformer({ tsconfig: 'tsconfig.json' });
