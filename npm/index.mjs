import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const { Excel } = require('./index.js');

export { Excel };
export default Excel;
