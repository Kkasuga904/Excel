const fs = require('fs');
const paths = ['dev.log', 'dev.err'];
for (const file of paths) {
  try {
    fs.unlinkSync(file);
  } catch (error) {
    // ignore missing or locked files
  }
}
