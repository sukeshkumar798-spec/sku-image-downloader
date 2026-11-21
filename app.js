/* SKU Image Downloader
 * Reads "Product SKU" and "Image" columns from an Excel file, downloads each
 * image and saves it to the chosen local folder as <Product SKU>.<ext>.
 * Failed rows can be exported as an Excel file in the same format.
 *
 * Requirements: Chrome / Edge 86+ (File System Access API). Must run inside the
 * extension (host_permissions allow cross-origin image fetches).
 */
(function () {
  'use strict';

  const CONCURRENCY = 2;              // parallel downloads (kept low to avoid server 503s)
  const MAX_ATTEMPTS = 4;             // tries per image on transient errors
  const RETRY_BASE_DELAY_MS = 2000;   // wait between retries (grows with each attempt)
  const FETCH_TIMEOUT_MS = 30000;     // max wait per image
  const IMAGE_EXTS = ['jpg', 'png', 'webp', 'gif', 'avif', 'svg', 'bmp', 'tif', 'tiff', 'heic'];
  const MIME_EXT = {
    'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/pjpeg': 'jpg',
    'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif',
    'image/avif': 'avif', 'image/svg+xml': 'svg', 'image/bmp': 'bmp',
    'image/tiff': 'tif', 'image/heic': 'heic'
  };
  // Reserved file names on Windows
  const RESERVED_NAME = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i;

  const $ = (id) => document.getElementById(id);
  const el = {
    excel: $('excelFile'),
    excelInfo: $('excelInfo'),
    pick: $('pickFolder'),
    folderName: $('folderName'),
    folderError: $('folderError'),
    start: $('start'),
    overwrite: $('overwrite'),
    bar: $('bar'),
    summary: $('summary'),
    status: $('status'),
    rows: $('rows'),
    downloadFailed: $('downloadFailed'),
    failInfo: $('failInfo')
  };

  let excelFile = null;   // Excel file chosen by the user
  let headers = null;     // original header text: { sku, image }
  let dirHandle = null;   // FileSystemDirectoryHandle (save folder)
  let failedItems = [];   // rows that failed in the last run
  let running = false;

  // ---------- UI helpers ----------

  function setStatus(msg, isError) {
    el.status.textContent = msg || '';
    el.status.className = isError ? 'err' : '';
  }

  function setFolderError(msg) {
    el.folderError.textContent = msg || '';
  }

  function updateStartButton() {
    el.start.disabled = running || !excelFile || !dirHandle;
  }

  function updateFailedButton() {
    el.downloadFailed.disabled = running || failedItems.length === 0;
    el.failInfo.textContent = failedItems.length
      ? `${failedItems.length} failed row(s) ready to export`
      : '';
  }

  // ---------- Excel parsing ----------

  /**
   * Reads the first sheet and returns the Product SKU + Image rows.
   * The header row is auto-detected (first row that has both an "sku" and an
   * "image" column), so blank rows or title rows above it are ignored.
   */
  async function parseExcel(file) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buffer), { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws || !ws['!ref']) throw new Error('No data found in the Excel file.');

    const range = XLSX.utils.decode_range(ws['!ref']);
    const grid = XLSX.utils.sheet_to_json(ws, {
      header: 1, defval: '', raw: false, blankrows: true
    });

    let headerIdx = -1, skuCol = -1, imgCol = -1;
    const scanLimit = Math.min(grid.length, 50);
    for (let r = 0; r < scanLimit && headerIdx < 0; r++) {
      const cells = (grid[r] || []).map((c) => String(c).trim().toLowerCase());
      const s = cells.findIndex((c) => c.includes('sku'));
      const i = cells.findIndex((c) => /image|img|photo|picture/.test(c));
      if (s >= 0 && i >= 0) { headerIdx = r; skuCol = s; imgCol = i; }
    }
    if (headerIdx < 0) {
      throw new Error('Header not found. The Excel file must have "Product SKU" and "Image" columns.');
    }

    const headerRow = grid[headerIdx] || [];
    const hdr = {
      sku: String(headerRow[skuCol] ?? '').trim() || 'Product SKU',
      image: String(headerRow[imgCol] ?? '').trim() || 'Image'
    };

    const items = [];
    for (let r = headerIdx + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const sku = String(row[skuCol] ?? '').trim();
      let url = String(row[imgCol] ?? '').trim();

      // If the cell has no text but holds a hyperlink, use the link target
      if (!url) {
        const addr = XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + imgCol });
        const cell = ws[addr];
        if (cell && cell.l && cell.l.Target) url = String(cell.l.Target).trim();
      }

      if (!sku && !url) continue; // completely empty row
      items.push({
        excelRow: range.s.r + r + 1,
        sku,
        url,
        base: '',        // final file name without extension (set in prepare())
        fileName: '',
        state: 'queued', // queued | saved | skipped | failed
        message: '',
        tr: null
      });
    }
    return { items, headers: hdr };
  }

  // ---------- Naming ----------

  /** Makes a Windows-safe file name from the SKU. */
  function safeFileBase(sku) {
    let name = sku
      .replace(/[<>:"\/\\|?*\x00-\x1F]/g, '_') // illegal characters
      .replace(/[. ]+$/g, '')                  // Windows: names cannot end with dot or space
      .trim();
    if (!name) return '';
    if (RESERVED_NAME.test(name)) name = '_' + name;
    return name.slice(0, 200);
  }

  /**
   * Duplicate handling:
   *  - same SKU + same image URL  => skipped (duplicate row)
   *  - same SKU + different URL   => NAME_2, NAME_3, ...
   */
  function prepare(items) {
    const seen = new Map(); // lower-case base -> { url, count }
    for (const it of items) {
      if (!it.sku) { it.state = 'failed'; it.message = 'Product SKU is empty'; continue; }

      const base = safeFileBase(it.sku);
      if (!base) { it.state = 'failed'; it.message = 'Product SKU cannot be turned into a valid file name'; continue; }

      const key = base.toLowerCase();
      const prev = seen.get(key);
      if (prev && prev.url === it.url) {
        it.state = 'skipped';
        it.message = 'Duplicate row (same SKU and same image)';
        continue;
      }
      if (prev) {
        prev.count += 1;
        it.base = `${base}_${prev.count}`;
      } else {
        seen.set(key, { url: it.url, count: 1 });
        it.base = base;
      }
    }
  }

  // ---------- Download ----------

  function extFromUrlPath(url) {
    try {
      const m = new URL(url).pathname.match(/\.([a-z0-9]{2,5})$/i);
      if (!m) return null;
      const ext = m[1].toLowerCase() === 'jpeg' ? 'jpg' : m[1].toLowerCase();
      return IMAGE_EXTS.includes(ext) ? ext : null;
    } catch (e) {
      return null;
    }
  }

  /** Fetches an image from the URL and returns { blob, ext }. */
  async function downloadImage(url) {
    if (!/^https?:\/\//i.test(url)) throw new Error('Image URL is missing or invalid');

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
    try {
      const res = await fetch(url, {
        signal: controller.signal,
        cache: 'no-store',
        redirect: 'follow',
        referrerPolicy: 'no-referrer'
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);

      const ct = (res.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
      const looksGeneric = !ct || ct === 'application/octet-stream' || ct === 'binary/octet-stream';
      if (!looksGeneric && !ct.startsWith('image/')) {
        throw new Error(`Not an image (content-type: ${ct})`);
      }

      const ext = MIME_EXT[ct] || extFromUrlPath(url) ||
        (ct.startsWith('image/') ? ct.split('/')[1].replace(/[^a-z0-9]/g, '') : null);
      if (!ext) throw new Error(`Could not detect image type (${ct || 'unknown'})`);

      const blob = await res.blob();
      if (!blob.size) throw new Error('Empty file received');
      return { blob, ext };
    } catch (e) {
      if (e.name === 'AbortError') throw new Error('Timeout (30 sec)');
      throw e;
    } finally {
      clearTimeout(timer);
    }
  }

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  /** 408 / 425 / 429 / 5xx, timeouts and network errors are treated as temporary. */
  function isTransientError(err) {
    return /^HTTP (408|425|429|5\d\d)$/.test(err.message) ||
      /Timeout|Failed to fetch|NetworkError/i.test(err.message);
  }

  async function downloadWithRetry(url, it) {
    for (let attempt = 1; ; attempt++) {
      try {
        return await downloadImage(url);
      } catch (e) {
        if (attempt >= MAX_ATTEMPTS || !isTransientError(e)) throw e;
        it.message = `Retry ${attempt}/${MAX_ATTEMPTS - 1}: ${e.message}`;
        renderRow(it);
        await sleep(RETRY_BASE_DELAY_MS * attempt);
      }
    }
  }

  async function fileExists(name) {
    try {
      await dirHandle.getFileHandle(name);
      return true;
    } catch (e) {
      if (e.name === 'NotFoundError') return false;
      throw e;
    }
  }

  async function saveFile(name, blob) {
    const fh = await dirHandle.getFileHandle(name, { create: true });
    const writable = await fh.createWritable();
    await writable.write(blob);
    await writable.close();
  }

  async function processItem(it) {
    try {
      const { blob, ext } = await downloadWithRetry(it.url, it);
      const fileName = `${it.base}.${ext}`;
      it.fileName = fileName;

      if (!el.overwrite.checked && await fileExists(fileName)) {
        it.state = 'skipped';
        it.message = 'File already exists in folder (overwrite is off)';
        return;
      }

      await saveFile(fileName, blob);
      it.state = 'saved';
      it.message = `${(blob.size / 1024).toFixed(1)} KB`;
    } catch (e) {
      it.state = 'failed';
      it.message = e.message || String(e);
    }
  }

  // ---------- Rendering ----------

  function renderRow(it) {
    const tr = it.tr || document.createElement('tr');
    tr.textContent = '';
    const cells = [
      String(it.excelRow),
      it.sku || '(blank)',
      it.fileName || (it.base ? it.base + '.?' : '-'),
      it.state.charAt(0).toUpperCase() + it.state.slice(1),
      it.message || ''
    ];
    cells.forEach((text, idx) => {
      const td = document.createElement('td');
      td.textContent = text;
      if (idx === 1) td.className = 'sku';
      if (idx === 3) td.className = 'st-' + it.state;
      tr.appendChild(td);
    });
    it.tr = tr;
    if (!tr.parentNode) el.rows.appendChild(tr);
  }

  function renderSummary(items, done) {
    const count = (s) => items.filter((it) => it.state === s).length;
    const total = items.length;
    el.summary.textContent =
      `${done === total ? 'Complete' : 'Progress'}: ${done}/${total}  |  Saved: ${count('saved')}  |  Skipped: ${count('skipped')}  |  Failed: ${count('failed')}`;
    el.bar.style.width = total ? `${Math.round((done / total) * 100)}%` : '0%';
  }

  // ---------- Failed list export ----------

  /**
   * Builds an Excel file with the same two columns as the uploaded file
   * (original header text, then Product SKU + Image URL for each failed row).
   */
  function downloadFailedList() {
    if (!failedItems.length || !headers) return;

    const aoa = [[headers.sku, headers.image]];
    for (const it of failedItems) aoa.push([it.sku, it.url]);

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 32 }, { wch: 60 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    const origName = excelFile ? excelFile.name.replace(/\.[^.]+$/, '') : 'images';
    XLSX.writeFile(wb, `failed_${origName}.xlsx`);
  }

  // ---------- Main flow ----------

  async function pickFolder() {
    setFolderError('');
    if (!window.showDirectoryPicker) {
      setFolderError(
        'The folder picker is not available on this page. Open the app from the extension icon ' +
        '(chrome://extensions > Load unpacked). Do not double-click app.html. ' +
        `Current protocol: ${location.protocol}`
      );
      return;
    }
    if (!window.isSecureContext) {
      setFolderError('This page is not a secure context, so the folder picker is blocked. Open it from the extension tab.');
      return;
    }
    try {
      dirHandle = await window.showDirectoryPicker({ mode: 'readwrite', id: 'sku-image-downloader' });
      el.folderName.textContent = `Folder: ${dirHandle.name}`;
      setStatus('');
    } catch (e) {
      if (e.name === 'AbortError') {
        setFolderError('No folder selected. Pick a folder in the dialog and click "Select Folder".');
      } else {
        setFolderError(`Folder picker error (${e.name}): ${e.message}`);
      }
    }
    updateStartButton();
  }

  async function onExcelChosen() {
    excelFile = el.excel.files[0] || null;
    headers = null;
    failedItems = [];
    el.rows.textContent = '';
    el.summary.textContent = '';
    el.bar.style.width = '0%';
    updateFailedButton();
    if (!excelFile) { el.excelInfo.textContent = ''; updateStartButton(); return; }

    try {
      const parsed = await parseExcel(excelFile);
      headers = parsed.headers;
      el.excelInfo.textContent = `${excelFile.name} - ${parsed.items.length} rows found`;
      setStatus(parsed.items.length ? '' : 'No data rows found in the Excel file.', !parsed.items.length);
    } catch (e) {
      excelFile = null;
      el.excel.value = '';
      el.excelInfo.textContent = '';
      setStatus(e.message, true);
    }
    updateStartButton();
  }

  async function run() {
    if (running || !excelFile || !dirHandle) return;
    running = true;
    failedItems = [];
    updateStartButton();
    updateFailedButton();
    el.rows.textContent = '';
    setStatus('Starting download...');

    let items;
    try {
      ({ items, headers } = await parseExcel(excelFile)); // fresh parse
    } catch (e) {
      running = false;
      updateStartButton();
      setStatus(e.message, true);
      return;
    }

    prepare(items);
    items.forEach(renderRow);

    // Rows already in a final state (duplicate / invalid) are not queued again
    const pending = items.filter((it) => it.state === 'queued');
    let nextIndex = 0;
    let done = items.length - pending.length;
    renderSummary(items, done);

    async function worker() {
      while (nextIndex < pending.length) {
        const it = pending[nextIndex++];
        await processItem(it);
        renderRow(it);
        done += 1;
        renderSummary(items, done);
      }
    }

    await Promise.all(Array.from({ length: Math.min(CONCURRENCY, pending.length) }, worker));

    items.forEach(renderRow);
    renderSummary(items, items.length);

    failedItems = items.filter((it) => it.state === 'failed');
    running = false;
    updateStartButton();
    updateFailedButton();
    setStatus(
      failedItems.length
        ? `Finished, but ${failedItems.length} row(s) failed. Download the failed list from step 4.`
        : 'All done.',
      failedItems.length > 0
    );
  }

  el.excel.addEventListener('change', onExcelChosen);
  el.pick.addEventListener('click', pickFolder);
  el.start.addEventListener('click', run);
  el.downloadFailed.addEventListener('click', downloadFailedList);
  updateStartButton();
  updateFailedButton();
})();
