/* ========================================
   Table Transformer v2 — App Logic
   ======================================== */

// ==================== State ====================
const S = {
  sourceData: null,
  sourceHeaders: [],
  filterData: null,
  filterHeaders: [],
  batchConfigs: [],
  extractedHeaders: null,
  extractedData: null,
  columnOrder: [],
};

// ==================== Stats (localStorage) ====================
const KEYS = {
  pageViews: 'tt2_pageViews',
  localUses: 'tt2_localUses',
  tasksDone: 'tt2_tasksDone',
};

function initStats() {
  // Page views: increment every load
  let pv = parseInt(localStorage.getItem(KEYS.pageViews)) || 0;
  pv++;
  localStorage.setItem(KEYS.pageViews, pv);

  const lu = parseInt(localStorage.getItem(KEYS.localUses)) || 0;
  const td = parseInt(localStorage.getItem(KEYS.tasksDone)) || 0;

  document.getElementById('stat-local-uses').textContent = lu.toLocaleString();
  document.getElementById('stat-tasks-done').textContent = td.toLocaleString();
}

function bumpStat(key) {
  let v = parseInt(localStorage.getItem(key)) || 0;
  v++;
  localStorage.setItem(key, v);

  if (key === KEYS.localUses) {
    document.getElementById('stat-local-uses').textContent = v.toLocaleString();
  } else if (key === KEYS.tasksDone) {
    document.getElementById('stat-tasks-done').textContent = v.toLocaleString();
  }
}

// ==================== Init ====================
document.addEventListener('DOMContentLoaded', () => {
  initStats();
  bindEvents();
  setupDragDrop();
});

function bindEvents() {
  // File inputs
  $('file-source').addEventListener('change', handleSourceFile);
  $('file-filter').addEventListener('change', handleFilterFile);
  $('file-batch').addEventListener('change', handleBatchFile);

  // Buttons
  $('btn-extract').addEventListener('click', extractData);
  $('btn-download').addEventListener('click', downloadResult);
  $('btn-reset').addEventListener('click', resetApp);
  $('btn-batch-extract').addEventListener('click', batchExtract);
  $('btn-batch-reset').addEventListener('click', resetApp);

  // Templates
  $('tpl-source').addEventListener('click', () => downloadTemplate('source'));
  $('tpl-filter').addEventListener('click', () => downloadTemplate('filter'));
  $('tpl-batch').addEventListener('click', () => downloadTemplate('batch'));

  // Config
  $('sel-filter-col').addEventListener('change', updateColumnPreview);
}

// ==================== Drag & Drop ====================
function setupDragDrop() {
  ['drop-source', 'drop-filter', 'drop-batch'].forEach(id => {
    const zone = $(id);
    if (!zone) return;
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('dragover');
      const file = e.dataTransfer.files[0];
      if (!file) return;
      const inputMap = { 'drop-source': 'file-source', 'drop-filter': 'file-filter', 'drop-batch': 'file-batch' };      const inputId = inputMap[id];
      const dt = new DataTransfer();
      dt.items.add(file);
      $(inputId).files = dt.files;
      $(inputId).dispatchEvent(new Event('change'));
    });
  });
}

// ==================== File Handlers ====================
function handleSourceFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  readExcel(file, (rows, headers) => {
    S.sourceData = rows;
    S.sourceHeaders = headers;

    $('meta-source').innerHTML = `✓ 已加载 <strong>${rows.length}</strong> 行, <strong>${headers.length}</strong> 列`;
    $('meta-source').classList.add('show');
    $('meta-source').classList.remove('error');
    $('card-source').classList.add('active');

    let preview = `<strong>可用列名:</strong> ${headers.join(', ')}<br>`;
    preview += `前2行预览:<br>`;
    rows.slice(0, 2).forEach(r => { preview += `&nbsp;&nbsp;• ${r.join(' | ')}<br>`; });
    $('preview-source').innerHTML = preview;

    checkReady();
    toast(`总表已加载: ${rows.length} 行 × ${headers.length} 列`, 'success');
  }, err => {
    $('meta-source').textContent = err;
    $('meta-source').classList.add('show', 'error');
    toast(err, 'error');
  });
}

function handleFilterFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  readExcel(file, (rows, headers) => {
    S.filterData = rows;
    S.filterHeaders = headers;

    $('meta-filter').innerHTML = `✓ 已加载 <strong>${rows.length}</strong> 行`;
    $('meta-filter').classList.add('show');
    $('meta-filter').classList.remove('error');
    $('card-filter').classList.add('active');

    let preview = `<strong>列名:</strong> ${headers.join(', ')}<br>`;
    rows.slice(0, 10).forEach(r => { preview += `&nbsp;&nbsp;• ${r.join(' | ')}<br>`; });
    if (rows.length > 10) preview += `...还有 ${rows.length - 10} 行`;
    $('preview-filter').innerHTML = preview;

    checkReady();
    toast(`列名表已加载: ${rows.length} 行`, 'success');
  }, err => {
    $('meta-filter').textContent = err;
    $('meta-filter').classList.add('show', 'error');
    toast(err, 'error');
  });
}

function handleBatchFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  readExcel(file, (rows, headers) => {
    S.batchConfigs = rows.map(r => ({
      filename: r[0] || '',
      columns: (r[1] || '').split(';').map(c => c.trim()).filter(Boolean)
    }));

    $('meta-batch').innerHTML = `✓ 已加载 <strong>${S.batchConfigs.length}</strong> 个任务`;
    $('meta-batch').classList.add('show');
    $('meta-batch').classList.remove('error');
    $('card-batch').classList.add('active');

    showBatchPreview();
    checkReady(); // 检查是否可以显示执行按钮
    toast(`批量配置已加载: ${S.batchConfigs.length} 个任务`, 'success');
  }, err => {
    $('meta-batch').textContent = err;
    $('meta-batch').classList.add('show', 'error');
    toast(err, 'error');
  });
}

// ==================== Readiness Check ====================
function checkReady() {
  if (S.sourceData && S.filterData) {
    populateFilterColSelect();
    $('config-panel').style.display = 'block';
    $('reorder-panel').style.display = 'none';
    $('result-panel').style.display = 'none';
  }
  // 批量模式随时可用，只需总表
  if (S.sourceData && S.batchConfigs.length > 0) {
    $('batch-execute-section').style.display = 'block';
  }
}

function populateFilterColSelect() {
  const sel = $('sel-filter-col');
  sel.innerHTML = '';
  S.filterHeaders.forEach((h, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = h || `列 ${i + 1}`;
    sel.appendChild(opt);
  });
  sel.value = 0;
  updateColumnPreview();
}

// ==================== Column Preview ====================
function updateColumnPreview() {
  const colIdx = parseInt($('sel-filter-col').value);
  const names = S.filterData
    .map(r => r[colIdx])
    .filter(v => v && String(v).trim())
    .map(v => String(v).trim());
  const unique = [...new Set(names)];

  let html = `<div class="col-preview-title">共 <strong>${unique.length}</strong> 个列名</div><div class="col-tags">`;
  unique.forEach(name => {
    const found = S.sourceHeaders.includes(name);
    html += `<span class="col-tag ${found ? 'found' : 'not-found'}">${found ? '✓' : '✗'} ${esc(name)}</span>`;
  });
  html += '</div>';
  $('col-preview').innerHTML = html;

  $('badge-rows').textContent = `${S.sourceData.length} 行`;
  $('badge-cols').textContent = `${unique.length} 列`;
}

// ==================== Single Extract ====================
function extractData() {
  const colIdx = parseInt($('sel-filter-col').value);
  const names = S.filterData
    .map(r => r[colIdx])
    .filter(v => v && String(v).trim())
    .map(v => String(v).trim());
  const unique = [...new Set(names)];

  const indices = [];
  const headers = [];
  const missing = [];

  unique.forEach(name => {
    const idx = S.sourceHeaders.indexOf(name);
    if (idx !== -1) { indices.push(idx); headers.push(name); }
    else { missing.push(name); }
  });

  const rows = S.sourceData.map(r => indices.map(i => r[i]));

  S.extractedHeaders = headers;
  S.extractedData = rows;
  S.columnOrder = headers.map((_, i) => i);

  // Show reorder
  if (headers.length > 1) {
    renderReorder(headers);
    $('reorder-panel').style.display = 'block';
  }

  showResult(rows, headers, missing);

  // Stats
  bumpStat(KEYS.localUses);
  bumpStat(KEYS.tasksDone);

  if (missing.length > 0) {
    toast(`提取完成，${missing.length} 个列未找到`, 'warning');
  } else {
    toast(`成功提取 ${headers.length} 列数据`, 'success');
  }
}

function showResult(rows, headers, missing) {
  const panel = $('result-panel');
  panel.style.display = 'block';

  let statsText = `提取了 <strong>${headers.length}</strong> 列, <strong>${rows.length}</strong> 行数据`;
  if (missing.length > 0) {
    statsText += `<br><span style="color:var(--error)">⚠️ ${missing.length} 列未找到: ${missing.join(', ')}</span>`;
  }
  $('result-stats-text').innerHTML = statsText;

  if (rows.length === 0 || headers.length === 0) {
    $('result-table').innerHTML = `
      <tr><td style="text-align:center;padding:40px;color:var(--text-muted)">没有提取到数据</td></tr>`;
    return;
  }

  // Render table
  const orderedH = S.columnOrder.map(i => headers[i]);
  const orderedR = rows.map(r => S.columnOrder.map(i => r[i]));

  let html = '<thead><tr>';
  orderedH.forEach(h => { html += `<th>${esc(h)}</th>`; });
  html += '</tr></thead><tbody>';
  orderedR.slice(0, 100).forEach(row => {
    html += '<tr>';
    row.forEach(cell => { html += `<td>${esc(cell)}</td>`; });
    html += '</tr>';
  });
  html += '</tbody>';
  if (orderedR.length > 100) {
    html += `<tfoot><tr><td colspan="${orderedH.length}" style="text-align:center;padding:12px;color:var(--text-muted)">显示前100行，共 ${orderedR.length} 行</td></tr></tfoot>`;
  }
  $('result-table').innerHTML = html;
  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ==================== Reorder ====================
function renderReorder(headers) {
  const track = $('reorder-track');
  track.innerHTML = '';

  S.columnOrder.forEach((origIdx, curIdx) => {
    const item = document.createElement('div');
    item.className = 'reorder-item';
    item.draggable = true;
    item.dataset.curIdx = curIdx;
    item.textContent = headers[origIdx];

    // Controls
    const controls = document.createElement('div');
    controls.className = 'reorder-controls';

    const left = document.createElement('button');
    left.className = 'reorder-btn';
    left.innerHTML = '←';
    left.addEventListener('click', e => { e.stopPropagation(); moveCol(curIdx, -1); });

    const right = document.createElement('button');
    right.className = 'reorder-btn';
    right.innerHTML = '→';
    right.addEventListener('click', e => { e.stopPropagation(); moveCol(curIdx, 1); });

    controls.appendChild(left);
    controls.appendChild(right);
    item.appendChild(controls);

    // Drag events
    item.addEventListener('dragstart', function () {
      this._dragIdx = parseInt(this.dataset.curIdx);
      this.classList.add('dragging');
    });
    item.addEventListener('dragover', function (e) {
      e.preventDefault();
      this.classList.add('drag-over');
    });
    item.addEventListener('drop', function (e) {
      e.preventDefault();
      const from = this._dragIdx || 0;
      const to = parseInt(this.dataset.curIdx);
      if (from !== to) {
        const temp = S.columnOrder.splice(from, 1)[0];
        S.columnOrder.splice(to, 0, temp);
        renderReorder(S.extractedHeaders);
        showResult(S.extractedData, S.extractedHeaders, []);
      }
    });
    item.addEventListener('dragend', function () {
      this.classList.remove('dragging');
      track.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));
    });

    track.appendChild(item);
  });
}

function moveCol(idx, dir) {
  const newIdx = idx + dir;
  if (newIdx < 0 || newIdx >= S.columnOrder.length) return;
  const temp = S.columnOrder[idx];
  S.columnOrder[idx] = S.columnOrder[newIdx];
  S.columnOrder[newIdx] = temp;
  renderReorder(S.extractedHeaders);
  showResult(S.extractedData, S.extractedHeaders, []);
}

// ==================== Batch Extract ====================
async function batchExtract() {
  if (!S.sourceData || S.batchConfigs.length === 0) {
    toast('请先上传总表和批量配置表', 'error');
    return;
  }

  let success = 0, fail = 0;
  const results = [];

  for (let i = 0; i < S.batchConfigs.length; i++) {
    const cfg = S.batchConfigs[i];
    const missing = cfg.columns.filter(c => !S.sourceHeaders.includes(c));

    if (missing.length > 0 || cfg.columns.length === 0) {
      fail++;
      results.push({ filename: cfg.filename, ok: false, reason: missing.length ? `缺失列: ${missing.join(', ')}` : '无指定列' });
      continue;
    }

    const indices = cfg.columns.map(c => S.sourceHeaders.indexOf(c));
    const rows = S.sourceData.map(r => indices.map(idx => r[idx]));

    try {
      const safeName = (cfg.filename || `提取结果_${i + 1}`).replace(/[^a-zA-Z0-9\u4e00-\u9fa5_-]/g, '_');
      await downloadExcelData([cfg.columns, ...rows], `${safeName}.xlsx`);
      success++;
      results.push({ filename: safeName, ok: true, cols: cfg.columns.length, rows: rows.length });
    } catch (err) {
      fail++;
      results.push({ filename: cfg.filename, ok: false, reason: err.message });
    }
  }

  // Show results
  showBatchResults(results, success, fail);

  if (success > 0) {
    bumpStat(KEYS.localUses);
    bumpStat(KEYS.tasksDone);
  }

  toast(fail === 0 ? `批量完成，${success} 个文件已下载` : `完成：${success} 成功 / ${fail} 失败`, fail > 0 ? 'warning' : 'success');
}

function showBatchResults(results, success, fail) {
  const panel = $('batch-result-panel');
  panel.style.display = 'block';

  $('batch-result-stats-text').innerHTML =
    `成功 <strong style="color:var(--success)">${success}</strong> · 失败 <strong style="color:var(--error)">${fail}</strong>`;

  const list = $('batch-result-list');
  list.innerHTML = results.map(r => `
    <div class="batch-result-item ${r.ok ? '' : 'failed'}">
      <span class="batch-result-icon">${r.ok ? '✅' : '❌'}</span>
      <div>
        <div class="batch-result-name">${esc(r.filename)}</div>
        <div class="batch-result-meta">${r.ok ? `${r.cols} 列 × ${r.rows} 行` : r.reason}</div>
      </div>
    </div>
  `).join('');

  panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function showBatchPreview() {
  // 使用 preview-batch 而不是 batch-preview-area
  const area = $('preview-batch');
  if (!area) return;

  let html = '<div class="batch-config-list">';
  S.batchConfigs.slice(0, 6).forEach((cfg, i) => {
    const missing = cfg.columns.filter(c => S.sourceHeaders && !S.sourceHeaders.includes(c));
    const ok = missing.length === 0 && cfg.columns.length > 0;
    html += `<div class="batch-config-item ${ok ? 'ok' : 'err'}">
      <span class="batch-config-idx">${i + 1}</span>
      <span class="batch-config-name">${esc(cfg.filename) || '(未命名)'}</span>
      <span class="batch-config-cols">${cfg.columns.length} 列</span>
    </div>`;
  });
  if (S.batchConfigs.length > 6) {
    html += `<div class="batch-config-more">...还有 ${S.batchConfigs.length - 6} 个</div>`;
  }
  html += '</div>';
  area.innerHTML = html;
}

// ==================== Download ====================
function downloadResult() {
  if (!S.extractedData || !S.extractedHeaders) {
    toast('没有数据可下载', 'error');
    return;
  }

  const fmt = $('sel-format').value;
  const ts = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
  const orderedH = S.columnOrder.map(i => S.extractedHeaders[i]);
  const orderedR = S.extractedData.map(r => S.columnOrder.map(i => r[i]));
  const data = [orderedH, ...orderedR];

  try {
    if (fmt === 'xlsx') {
      downloadExcelData(data, `提取结果_${ts}.xlsx`).then(() => toast('下载已开始', 'success'));
    } else {
      const ws = XLSX.utils.aoa_to_sheet(data);
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `提取结果_${ts}.csv`;
      link.click();
      URL.revokeObjectURL(link.href);
      toast('下载已开始', 'success');
    }
  } catch (err) {
    toast('下载失败: ' + err.message, 'error');
  }
}

function downloadExcelData(data, filename) {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return new Promise(resolve => {
    setTimeout(() => {
      XLSX.writeFile(wb, filename);
      resolve();
    }, 100);
  });
}

// ==================== Templates ====================
function downloadTemplate(type) {
  const templates = {
    source: {
      data: [
        ['姓名', '部门', '职位', '入职日期', '薪资', '城市'],
        ['张三', '研发部', '高级工程师', '2020-03-15', '35000', '北京'],
        ['李四', '销售部', '销售经理', '2019-06-20', '28000', '上海'],
        ['王五', '研发部', '工程师', '2021-01-10', '22000', '深圳'],
      ],
      name: '总表模板.xlsx'
    },
    filter: {
      data: [
        ['需要提取的列名'],
        ['姓名'],
        ['部门'],
        ['薪资'],
      ],
      name: '列名表模板.xlsx'
    },
    batch: {
      data: [
        ['文件名', '列名'],
        ['基本信息', '姓名;部门;城市'],
        ['薪资信息', '姓名;部门;薪资'],
        ['职位信息', '姓名;职位;入职日期'],
      ],
      name: '批量配置表模板.xlsx'
    }
  };

  const t = templates[type];
  downloadExcelData(t.data, t.name);
  toast(`${t.name} 已下载`, 'success');
}

// ==================== Reset ====================
function resetApp() {
  S.sourceData = null;
  S.sourceHeaders = [];
  S.filterData = null;
  S.filterHeaders = [];
  S.batchConfigs = [];
  S.extractedHeaders = null;
  S.extractedData = null;
  S.columnOrder = [];

  // Reset source
  $('file-source').value = '';
  $('meta-source').textContent = '';
  $('meta-source').classList.remove('show', 'error');
  $('preview-source').innerHTML = '';
  $('card-source').classList.remove('active');

  // Reset filter
  $('file-filter').value = '';
  $('meta-filter').textContent = '';
  $('meta-filter').classList.remove('show', 'error');
  $('preview-filter').innerHTML = '';
  $('card-filter').classList.remove('active');

  // Reset batch
  $('file-batch').value = '';
  $('meta-batch').textContent = '';
  $('meta-batch').classList.remove('show', 'error');
  $('preview-batch').innerHTML = '';
  $('card-batch').classList.remove('active');
  $('batch-execute-section').style.display = 'none';

  // Hide dynamic panels
  $('config-panel').style.display = 'none';
  $('reorder-panel').style.display = 'none';
  $('result-panel').style.display = 'none';
  $('batch-result-panel').style.display = 'none';

  toast('已重置，请重新上传', 'success');
}

// ==================== Utility ====================
function $(id) { return document.getElementById(id); }

function readExcel(file, onSuccess, onError) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1 });

      if (json.length === 0) { onError('文件为空'); return; }

      const headers = json[0].map(h => String(h || ''));
      const rows = json.slice(1).map(r => {
        const padded = [...r];
        while (padded.length < headers.length) padded.push('');
        return padded.map(c => String(c || ''));
      });

      onSuccess(rows, headers);
    } catch (err) {
      onError('文件格式错误: ' + err.message);
    }
  };
  reader.onerror = () => onError('读取文件失败');
  reader.readAsArrayBuffer(file);
}

function esc(text) {
  const d = document.createElement('div');
  d.textContent = String(text || '');
  return d.innerHTML;
}

function toast(msg, type = 'success') {
  const t = $('toast');
  t.textContent = msg;
  t.className = 'toast show ' + type;
  setTimeout(() => t.classList.remove('show'), 3000);
}
