const api = window.pdfTransformer;

const $ = (id) => document.getElementById(id);

let selectedPath = null;
let pdfDoc = null;
let pdfPageNum = 1;
let pdfScale = 1.2;

async function refreshLoStatus() {
  const el = $('loStatus');
  try {
    const { ok, path } = await api.libreOfficeStatus();
    if (ok) {
      el.textContent = 'LibreOffice：已检测到';
      el.className = 'status ok';
    } else {
      el.textContent = 'LibreOffice：未检测到（Office 文档转换需要安装）';
      el.className = 'status bad';
    }
  } catch {
    el.textContent = '';
    el.className = 'status';
  }
}

function setMsg(text, kind) {
  const el = $('msg');
  el.textContent = text || '';
  el.className = 'msg' + (kind ? ` ${kind}` : '');
}

function clearSourcePreview() {
  const box = $('sourcePreview');
  box.innerHTML =
    '<p class="placeholder">选择图片可在此预览；Office 文档转换后在右侧查看 PDF。</p>';
}

function clearPdfPreview() {
  pdfDoc = null;
  pdfPageNum = 1;
  $('pageInfo').textContent = '';
  $('btnPrev').disabled = true;
  $('btnNext').disabled = true;
  const canvas = $('pdfCanvas');
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);
}

async function renderPdfPage() {
  if (!pdfDoc) return;
  const canvas = $('pdfCanvas');
  const page = await pdfDoc.getPage(pdfPageNum);
  const vp = page.getViewport({ scale: pdfScale });
  canvas.width = vp.width;
  canvas.height = vp.height;
  await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp })
    .promise;
  $('pageInfo').textContent = `${pdfPageNum} / ${pdfDoc.numPages}`;
  $('btnPrev').disabled = pdfPageNum <= 1;
  $('btnNext').disabled = pdfPageNum >= pdfDoc.numPages;
}

async function loadPdfFromUrl(url) {
  const pdfjs = await import('./pdfjs/pdf.mjs');
  pdfjs.GlobalWorkerOptions.workerSrc = new URL(
    './pdfjs/pdf.worker.mjs',
    import.meta.url
  ).href;
  const loadingTask = pdfjs.getDocument({ url });
  pdfDoc = await loadingTask.promise;
  pdfPageNum = 1;
  await renderPdfPage();
}

$('btnPick').addEventListener('click', async () => {
  setMsg('');
  const p = await api.selectFile();
  if (!p) return;
  selectedPath = p;
  $('fileLabel').textContent = p;
  $('btnConvert').disabled = false;
  clearPdfPreview();

  const prev = await api.sourcePreviewUrl(p);
  const box = $('sourcePreview');
  box.innerHTML = '';
  if (prev && prev.type === 'image' && prev.url) {
    const img = document.createElement('img');
    img.alt = '源文件预览';
    img.src = prev.url;
    box.appendChild(img);
  } else if (prev && prev.type === 'office') {
    const pEl = document.createElement('p');
    pEl.className = 'placeholder';
    pEl.textContent = `已选择 Office 文档：${prev.name}\n转换后将在此侧显示 PDF 预览（右侧）。`;
    box.appendChild(pEl);
  } else {
    clearSourcePreview();
  }
});

$('btnConvert').addEventListener('click', async () => {
  if (!selectedPath) return;
  setMsg('正在转换…');
  $('btnConvert').disabled = true;
  try {
    const { previewUrl } = await api.convertToPdf(selectedPath);
    await loadPdfFromUrl(previewUrl);
    setMsg('转换完成', 'ok');
  } catch (e) {
    setMsg(e.message || String(e), 'error');
  } finally {
    $('btnConvert').disabled = false;
  }
});

$('btnPrev').addEventListener('click', async () => {
  if (pdfPageNum > 1) {
    pdfPageNum--;
    await renderPdfPage();
  }
});

$('btnNext').addEventListener('click', async () => {
  if (pdfDoc && pdfPageNum < pdfDoc.numPages) {
    pdfPageNum++;
    await renderPdfPage();
  }
});

clearSourcePreview();
refreshLoStatus();
