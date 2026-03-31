const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const pdfjsRoot = path.dirname(require.resolve('pdfjs-dist/package.json'));
const buildDir = path.join(pdfjsRoot, 'build');
const destDir = path.join(root, 'public', 'pdfjs');

const files = ['pdf.mjs', 'pdf.worker.mjs'];

try {
  fs.mkdirSync(destDir, { recursive: true });
  for (const f of files) {
    const src = path.join(buildDir, f);
    if (!fs.existsSync(src)) {
      console.warn('copy-pdfjs: missing', src);
      continue;
    }
    fs.copyFileSync(src, path.join(destDir, f));
  }
  console.log('pdfjs copied to public/pdfjs');
} catch (e) {
  console.warn('copy-pdfjs:', e.message);
}
