const {
  app,
  BrowserWindow,
  ipcMain,
  dialog,
  protocol,
  net,
} = require('electron');

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'app-preview',
    privileges: {
      bypassCSP: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: true,
      standard: true,
    },
  },
]);
const path = require('path');
const fs = require('fs');
const fsPromises = require('fs').promises;
const { pathToFileURL } = require('url');
const { randomUUID } = require('crypto');
const { spawn } = require('child_process');
const { PDFDocument } = require('pdf-lib');
const sharp = require('sharp');

const previewTokens = new Map();

const IMAGE_EXT = new Set([
  '.png',
  '.jpg',
  '.jpeg',
  '.gif',
  '.webp',
  '.bmp',
  '.tif',
  '.tiff',
]);

const OFFICE_EXT = new Set([
  '.doc',
  '.docx',
  '.xls',
  '.xlsx',
  '.ppt',
  '.pptx',
  '.odt',
  '.ods',
  '.odp',
  '.rtf',
]);

function findSoffice() {
  const envPath = process.env.LIBREOFFICE_PATH;
  if (envPath && fs.existsSync(envPath)) return envPath;

  const candidates = [
    path.join(
      process.env['PROGRAMFILES'] || 'C:\\Program Files',
      'LibreOffice',
      'program',
      'soffice.exe'
    ),
    path.join(
      process.env['PROGRAMFILES(X86)'] || 'C:\\Program Files (x86)',
      'LibreOffice',
      'program',
      'soffice.exe'
    ),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

function extOf(filePath) {
  return path.extname(filePath).toLowerCase();
}

async function imageToPdf(inputPath, outputPath) {
  const imgBuffer = await fsPromises.readFile(inputPath);
  const pngBuffer = await sharp(imgBuffer).png().toBuffer();
  const pdfDoc = await PDFDocument.create();
  const pngImage = await pdfDoc.embedPng(pngBuffer);
  const { width, height } = pngImage.scale(1);
  const page = pdfDoc.addPage([width, height]);
  page.drawImage(pngImage, {
    x: 0,
    y: 0,
    width,
    height,
  });
  const pdfBytes = await pdfDoc.save();
  await fsPromises.writeFile(outputPath, pdfBytes);
}

function convertOfficeToPdf(inputPath, outDir) {
  const soffice = findSoffice();
  if (!soffice) {
    return Promise.reject(
      new Error(
        '未找到 LibreOffice。请安装 LibreOffice，或设置环境变量 LIBREOFFICE_PATH 指向 soffice.exe'
      )
    );
  }
  return new Promise((resolve, reject) => {
    const proc = spawn(
      soffice,
      ['--headless', '--convert-to', 'pdf', '--outdir', outDir, inputPath],
      { windowsHide: true }
    );
    let stderr = '';
    proc.stderr.on('data', (d) => {
      stderr += d.toString();
    });
    proc.on('error', reject);
    proc.on('close', (code) => {
      if (code !== 0) {
        reject(new Error(stderr.trim() || `LibreOffice 退出码 ${code}`));
        return;
      }
      const base = path.basename(inputPath, path.extname(inputPath)) + '.pdf';
      const out = path.join(outDir, base);
      if (!fs.existsSync(out)) {
        reject(new Error('未找到输出 PDF：' + out));
        return;
      }
      resolve(out);
    });
  });
}

function registerPreviewProtocol() {
  protocol.handle('app-preview', async (request) => {
    try {
      const host = new URL(request.url).hostname;
      const filePath = previewTokens.get(host);
      if (!filePath || !fs.existsSync(filePath)) {
        return new Response(null, { status: 404 });
      }
      return net.fetch(pathToFileURL(filePath).href);
    } catch {
      return new Response(null, { status: 500 });
    }
  });
}

function tokenUrlForPath(absPath) {
  const token = randomUUID();
  previewTokens.set(token, absPath);
  return `app-preview://${token}/`;
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 760,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  win.loadFile(path.join(__dirname, 'public', 'index.html'));
}

app.whenReady().then(() => {
  registerPreviewProtocol();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

ipcMain.handle('libreoffice-status', () => {
  const p = findSoffice();
  return { ok: !!p, path: p || null };
});

ipcMain.handle('select-file', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      {
        name: '支持的文件',
        extensions: [
          'png',
          'jpg',
          'jpeg',
          'gif',
          'webp',
          'bmp',
          'tif',
          'tiff',
          'doc',
          'docx',
          'xls',
          'xlsx',
          'ppt',
          'pptx',
          'odt',
          'ods',
          'odp',
          'rtf',
        ],
      },
      { name: '所有文件', extensions: ['*'] },
    ],
  });
  if (canceled || !filePaths[0]) return null;
  return filePaths[0];
});

ipcMain.handle('source-preview-url', async (_e, filePath) => {
  if (!filePath || !fs.existsSync(filePath)) return null;
  const ext = extOf(filePath);
  if (!IMAGE_EXT.has(ext)) return { type: 'office', name: path.basename(filePath) };
  return { type: 'image', url: tokenUrlForPath(filePath) };
});

ipcMain.handle('convert-to-pdf', async (_e, inputPath) => {
  if (!inputPath || !fs.existsSync(inputPath)) {
    throw new Error('文件不存在');
  }
  const ext = extOf(inputPath);
  const tmpDir = app.getPath('temp');
  const outName = `${path.basename(inputPath, ext)}-${randomUUID().slice(0, 8)}.pdf`;
  const outputPath = path.join(tmpDir, outName);

  if (IMAGE_EXT.has(ext)) {
    await imageToPdf(inputPath, outputPath);
  } else if (OFFICE_EXT.has(ext)) {
    const jobDir = path.join(tmpDir, `pdf-job-${randomUUID()}`);
    await fsPromises.mkdir(jobDir, { recursive: true });
    try {
      const pdfPath = await convertOfficeToPdf(inputPath, jobDir);
      await fsPromises.copyFile(pdfPath, outputPath);
    } finally {
      try {
        await fsPromises.rm(jobDir, { recursive: true, force: true });
      } catch {
        /* ignore */
      }
    }
  } else {
    throw new Error('不支持的文件类型：' + ext);
  }

  return { outputPath, previewUrl: tokenUrlForPath(outputPath) };
});

ipcMain.handle('export-pdf', async (_e, sourcePdfPath, defaultFileName) => {
  if (!sourcePdfPath || !fs.existsSync(sourcePdfPath)) {
    throw new Error('没有可导出的 PDF，请先完成转换');
  }
  const baseName =
    defaultFileName && String(defaultFileName).trim()
      ? path.basename(defaultFileName)
      : path.basename(sourcePdfPath);
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: '导出 PDF',
    defaultPath: baseName.endsWith('.pdf') ? baseName : `${baseName}.pdf`,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  });
  if (canceled || !filePath) return { canceled: true };
  let dest = filePath;
  if (!dest.toLowerCase().endsWith('.pdf')) dest += '.pdf';
  await fsPromises.copyFile(sourcePdfPath, dest);
  return { canceled: false, path: dest };
});
