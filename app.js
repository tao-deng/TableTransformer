// 全局状态
const state = {
    sourceData: null,
    filterData: null,
    sourceHeaders: [],
    filterHeaders: [],
    extractedData: null,
    extractedHeaders: null,
    extractedColumnOrder: [], // 存储列顺序索引
    batchConfigs: [], // 批量配置
    sourceFileName: '',
    filterFileName: '',
    batchFileName: ''
};

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    // 文件上传事件
    document.getElementById('sourceFile').addEventListener('change', handleSourceFile);
    document.getElementById('filterFile').addEventListener('change', handleFilterFile);
    document.getElementById('batchFile').addEventListener('change', handleBatchFile);

    // 按钮事件
    document.getElementById('extractBtn').addEventListener('click', extractData);
    document.getElementById('downloadBtn').addEventListener('click', downloadResult);
    document.getElementById('singleResetBtn').addEventListener('click', resetApp);
    document.getElementById('batchResetBtn').addEventListener('click', resetApp);
    document.getElementById('batchExtractBtn').addEventListener('click', batchExtract);

    // 模板下载事件
    document.getElementById('downloadSourceTemplate').addEventListener('click', downloadSourceTemplate);
    document.getElementById('downloadFilterTemplate').addEventListener('click', downloadFilterTemplate);
    document.getElementById('downloadBatchTemplate').addEventListener('click', downloadBatchTemplate);

    // 配置变化事件
    document.getElementById('filterColumn').addEventListener('change', updateFilterPreview);
    document.getElementById('batchColumn').addEventListener('change', updateBatchPreview);
});

// ==================== 模板下载 ====================
function downloadSourceTemplate() {
    const template = [
        ['姓名', '部门', '职位', '入职日期', '薪资', '城市'],
        ['张三', '研发部', '高级工程师', '2020-03-15', '35000', '北京'],
        ['李四', '销售部', '销售经理', '2019-06-20', '28000', '上海'],
        ['王五', '研发部', '工程师', '2021-01-10', '22000', '深圳'],
    ];

    downloadExcel(template, '总表模板.xlsx');
    showToast('总表模板已下载', 'success');
}

function downloadFilterTemplate() {
    const template = [
        ['需要提取的列名'],
        ['姓名'],
        ['部门'],
        ['薪资'],
    ];

    downloadExcel(template, '列名表模板.xlsx');
    showToast('列名表模板已下载', 'success');
}

function downloadBatchTemplate() {
    const template = [
        ['文件名', '列名'],
        ['基本信息', '姓名;部门;城市'],
        ['薪资信息', '姓名;部门;薪资'],
        ['职位信息', '姓名;职位;入职日期'],
    ];

    downloadExcel(template, '批量配置表模板.xlsx');
    showToast('批量配置表模板已下载', 'success');
}

function downloadExcel(data, filename) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // 使用 setTimeout 延迟下载，避免连续调用导致覆盖
    return new Promise((resolve) => {
        setTimeout(() => {
            XLSX.writeFile(wb, filename);
            resolve();
        }, 100);
    });
}

// ==================== 文件处理 ====================
function handleSourceFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    state.sourceFileName = file.name;
    document.getElementById('sourceFileName').textContent = file.name;

    readExcelFile(file, (data, headers) => {
        state.sourceData = data;
        state.sourceHeaders = headers;
        updateSourceUI();
        checkFilesLoaded();
        showToast(`总表已加载: ${data.length} 行, ${headers.length} 列`, 'success');
    }, (error) => {
        showFileError('sourceInfo', error);
        showToast(error, 'error');
    });
}

function handleFilterFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    state.filterFileName = file.name;
    document.getElementById('filterFileName').textContent = file.name;

    readExcelFile(file, (data, headers) => {
        state.filterData = data;
        state.filterHeaders = headers;
        updateFilterUI();
        checkFilesLoaded();
        showToast(`列名表已加载: ${data.length} 行`, 'success');
    }, (error) => {
        showFileError('filterInfo', error);
        showToast(error, 'error');
    });
}

function handleBatchFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    state.batchFileName = file.name;
    document.getElementById('batchFileName').textContent = file.name;

    readExcelFile(file, (data, headers) => {
        state.batchConfigs = data.map(row => ({
            filename: row[0] || '',
            columns: (row[1] || '').split(';').map(c => c.trim()).filter(c => c)
        }));

        updateBatchUI();
        checkFilesLoaded();
        showToast(`批量配置已加载: ${state.batchConfigs.length} 个任务`, 'success');
    }, (error) => {
        showFileError('batchInfo', error);
        showToast(error, 'error');
    });
}

function updateSourceUI() {
    const infoEl = document.getElementById('sourceInfo');
    const previewEl = document.getElementById('sourcePreview');

    infoEl.innerHTML = `✓ 已加载 <span class="count">${state.sourceData.length}</span> 行, <span class="count">${state.sourceHeaders.length}</span> 列`;
    infoEl.classList.remove('error');

    previewEl.innerHTML = `
        <div class="headers">可用列名: ${state.sourceHeaders.join(', ')}</div>
        <div>前2行数据预览:</div>
        ${state.sourceData.slice(0, 2).map(row =>
            `<div style="margin-left: 16px; margin-top: 4px;">• ${row.join(' | ')}</div>`
        ).join('')}
    `;
}

function updateFilterUI() {
    const infoEl = document.getElementById('filterInfo');
    const previewEl = document.getElementById('filterPreview');

    infoEl.innerHTML = `✓ 已加载 <span class="count">${state.filterData.length}</span> 行`;
    infoEl.classList.remove('error');

    previewEl.innerHTML = `
        <div class="headers">列名: ${state.filterHeaders.join(', ')}</div>
        <div>所有内容:</div>
        ${state.filterData.map(row =>
            `<div style="margin-left: 16px; margin-top: 4px;">• ${row.join(' | ')}</div>`
        ).join('')}
    `;
}

function updateBatchUI() {
    const infoEl = document.getElementById('batchInfo');
    const previewEl = document.getElementById('batchPreview');

    infoEl.innerHTML = `✓ 已加载 <span class="count">${state.batchConfigs.length}</span> 个提取任务`;
    infoEl.classList.remove('error');

    previewEl.innerHTML = state.batchConfigs.slice(0, 5).map((config, i) => `
        <div style="margin: 8px 0; padding: 10px; background: #f3f4f6; border-radius: 8px;">
            <div style="font-weight: 600; color: #6366f1;">📄 ${escapeHtml(config.filename)}</div>
            <div style="font-size: 0.9rem; color: #6b7280; margin-top: 4px;">列: ${config.columns.map(c => `<span style="background: white; padding: 2px 8px; border-radius: 4px; margin: 0 4px 0 0;">${escapeHtml(c)}</span>`).join('')}</div>
        </div>
    `).join('');

    if (state.batchConfigs.length > 5) {
        previewEl.innerHTML += `<div style="text-align: center; padding: 10px; color: #6b7280;">...还有 ${state.batchConfigs.length - 5} 个任务</div>`;
    }
}

function showFileError(elementId, error) {
    const el = document.getElementById(elementId);
    el.innerHTML = `x ${error}`;
    el.classList.add('error');
}

// ==================== 配置管理 ====================
function checkFilesLoaded() {
    const hasSource = state.sourceData !== null;
    const hasFilter = state.filterData !== null;
    const hasBatch = state.batchConfigs.length > 0;

    // 单提取模式
    if (hasSource && hasFilter) {
        populateColumnSelect();
    }

    // 批量提取模式
    if (hasSource && hasBatch) {
        populateBatchColumnSelect();
    }
}

function populateColumnSelect() {
    const select = document.getElementById('filterColumn');
    select.innerHTML = '';

    state.filterHeaders.forEach((header, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = header || `列 ${index + 1}`;
        select.appendChild(option);
    });

    select.value = 0;
    document.getElementById('singleConfigSection').style.display = 'block';
    document.getElementById('singleActionSection').style.display = 'block';
    updateFilterPreview();
}

function populateBatchColumnSelect() {
    const select = document.getElementById('batchColumn');
    select.innerHTML = '';

    // 假设配置表有"文件名"和"列名"两列
    const headers = ['文件名', '列名'];
    headers.forEach((header, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = header;
        select.appendChild(option);
    });

    select.value = 0;
    document.getElementById('batchConfigSection').style.display = 'block';
    document.getElementById('batchActionSection').style.display = 'block';
    updateBatchPreview();
}

function updateFilterPreview() {
    const filterColumnIndex = parseInt(document.getElementById('filterColumn').value);
    const previewEl = document.getElementById('filterValuesPreview');

    const columnNames = state.filterData
        .map(row => row[filterColumnIndex])
        .filter(v => v && v.trim() !== '')
        .map(v => v.trim());

    const uniqueColumnNames = [...new Set(columnNames)];

    previewEl.innerHTML = `
        <div style="margin-bottom: 10px; color: #6b7280;">
            共找到 <strong>${uniqueColumnNames.length}</strong> 个列名
        </div>
        ${uniqueColumnNames.map(name => {
            const exists = state.sourceHeaders.includes(name);
            return `
                <span class="filter-tag ${exists ? 'found' : 'not-found'}" title="${exists ? '总表中存在' : '总表中不存在'}">
                    ${exists ? '✓' : '✗'} ${escapeHtml(name)}
                </span>
            `;
        }).join('')}
    `;
}

function updateBatchPreview() {
    const previewEl = document.getElementById('batchConfigPreview');

    // 验证每个配置
    const results = state.batchConfigs.map((config, i) => {
        const missing = config.columns.filter(c => !state.sourceHeaders.includes(c));
        const valid = missing.length === 0;

        return {
            index: i,
            filename: config.filename,
            columns: config.columns,
            valid: valid,
            missing: missing
        };
    });

    const invalidCount = results.filter(r => !r.valid).length;

    previewEl.innerHTML = `
        <div style="margin-bottom: 15px;">
            <span style="color: ${invalidCount === 0 ? '#10b981' : '#ef4444'}; font-weight: 600;">
                ${invalidCount === 0 ? '✓ 所有配置有效' : `⚠️ ${invalidCount} 个配置存在问题`}
            </span>
        </div>
        ${results.map(r => `
            <div style="margin: 8px 0; padding: 10px; background: ${r.valid ? '#ecfdf5' : '#fef2f2'}; border-radius: 8px; border-left: 3px solid ${r.valid ? '#10b981' : '#ef4444'};">
                <div style="font-weight: 600; color: #374151;">${r.index + 1}. ${escapeHtml(r.filename) || '(未命名)'}</div>
                <div style="font-size: 0.9rem; margin-top: 6px;">
                    ${r.columns.map(c => `<span style="background: white; padding: 2px 8px; border-radius: 4px; margin: 0 4px 0 0; color: ${state.sourceHeaders.includes(c) ? '#374151' : '#ef4444'};">${escapeHtml(c)}</span>`).join('')}
                </div>
                ${!r.valid ? `<div style="color: #ef4444; font-size: 0.85rem; margin-top: 6px;">缺失列: ${r.missing.join(', ')}</div>` : ''}
            </div>
        `).join('')}
    `;
}

// ==================== 单次数据提取 ====================
function extractData() {
    const filterColumnIndex = parseInt(document.getElementById('filterColumn').value);

    const columnNames = state.filterData
        .map(row => row[filterColumnIndex])
        .filter(v => v && v.trim() !== '')
        .map(v => v.trim());

    const uniqueColumnNames = [...new Set(columnNames)];

    const columnIndices = [];
    const extractedHeaders = [];
    const missingColumns = [];

    uniqueColumnNames.forEach(name => {
        const index = state.sourceHeaders.indexOf(name);
        if (index !== -1) {
            columnIndices.push(index);
            extractedHeaders.push(name);
        } else {
            missingColumns.push(name);
        }
    });

    const extractedRows = state.sourceData.map(row => {
        return columnIndices.map(index => row[index]);
    });

    state.extractedHeaders = extractedHeaders;
    state.extractedData = extractedRows;
    state.extractedColumnOrder = extractedHeaders.map((_, i) => i);

    showResults(extractedRows, extractedHeaders, missingColumns, true);

    if (missingColumns.length > 0) {
        showToast(`提取完成，但有 ${missingColumns.length} 个列未找到`, 'error');
    } else {
        showToast(`成功提取 ${extractedHeaders.length} 列数据！`, 'success');
    }
}

// ==================== 批量提取 ====================
async function batchExtract() {
    if (!state.sourceData || state.batchConfigs.length === 0) {
        showToast('请先上传总表和批量配置表', 'error');
        return;
    }

    const results = [];
    let totalFiles = 0;
    let failedFiles = 0;

    for (let index = 0; index < state.batchConfigs.length; index++) {
        const config = state.batchConfigs[index];

        // 验证列
        const missing = config.columns.filter(c => !state.sourceHeaders.includes(c));
        if (missing.length > 0 || config.columns.length === 0) {
            failedFiles++;
            results.push({
                index: index,
                filename: config.filename,
                success: false,
                reason: missing.length > 0 ? `缺失列: ${missing.join(', ')}` : '没有指定列'
            });
            continue;
        }

        // 提取数据
        const columnIndices = config.columns.map(c => state.sourceHeaders.indexOf(c));
        const extractedRows = state.sourceData.map(row => columnIndices.map(idx => row[idx]));

        // 下载文件（使用 await 确保按顺序下载）
        try {
            const safeFilename = (config.filename || `提取结果_${index + 1}`).replace(/[^a-zA-Z0-9\u4e00-\u9fa5_-]/g, '_');
            const filename = `${safeFilename}.xlsx`;
            const data = [config.columns, ...extractedRows];

            await downloadExcel(data, filename);
            totalFiles++;
            results.push({
                index: index,
                filename: filename,
                success: true,
                columns: config.columns.length,
                rows: extractedRows.length
            });
        } catch (error) {
            failedFiles++;
            results.push({
                index: index,
                filename: config.filename,
                success: false,
                reason: error.message
            });
        }
    }

    // 显示结果
    showBatchResults(results, totalFiles, failedFiles);

    if (failedFiles === 0) {
        showToast(`批量提取完成！已下载 ${totalFiles} 个文件`, 'success');
    } else {
        showToast(`批量提取完成：成功 ${totalFiles} 个，失败 ${failedFiles} 个`, 'error');
    }
}

function showBatchResults(results, successCount, failCount) {
    const resultSection = document.getElementById('batchResultSection');
    const resultPreview = document.getElementById('batchResultPreview');

    resultSection.style.display = 'block';

    resultPreview.innerHTML = `
        <div style="text-align: center; padding: 30px;">
            <div style="font-size: 3rem; margin-bottom: 20px;">
                ${failCount === 0 ? '🎉' : '📊'}
            </div>
            <div style="font-size: 1.2rem; font-weight: 600; margin-bottom: 10px;">
                批量提取完成
            </div>
            <div style="color: #610b981; font-size: 1rem;">
                成功: <span style="color: #10b981; font-weight: 600;">${successCount}</span> |
                失败: <span style="${failCount > 0 ? 'color: #ef4444;' : ''} font-weight: 600;">${failCount}</span>
            </div>
        </div>
        ${results.map(r => `
            <div style="margin: 10px 0; padding: 12px; background: ${r.success ? '#ecfdf5' : '#fef2f2'}; border-radius: 8px; display: flex; align-items: center; gap: 12px;">
                <span style="font-size: 1.5rem;">${r.success ? '✅' : '❌'}</span>
                <div>
                    <div style="font-weight: 600; color: #374151;">${escapeHtml(r.filename)}</div>
                    <div style="font-size: 0.9rem; color: #6b7280;">
                        ${r.success ? `${r.columns} 列 × ${r.rows} 行` : r.reason}
                    </div>
                </div>
            </div>
        `).join('')}
    `;

    resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ==================== 结果显示 ====================
function showResults(resultRows, headers, missingColumns, showReorder = false) {
    const resultSection = document.getElementById('resultSection');
    const resultStats = document.getElementById('resultStats');
    const resultPreview = document.getElementById('resultPreview');
    const reorderSection = document.getElementById('columnReorderSection');

    resultSection.style.display = 'block';

    let stats = `
        提取了 <strong>${headers.length}</strong> 列, <strong>${resultRows.length}</strong> 行数据
    `;

    if (missingColumns.length > 0) {
        stats += `<br><span style="color: #ef4444; font-size: 0.9rem;">⚠️ ${missingColumns.length} 个列未找到: ${missingColumns.join(', ')}</span>`;
    }

    resultStats.innerHTML = stats;

    // 显示列重排界面
    if (showReorder && headers.length > 1) {
        reorderSection.style.display = 'block';
        renderReorderColumns(headers);
    } else {
        reorderSection.style.display = 'none';
    }

    if (resultRows.length === 0 || headers.length === 0) {
        resultPreview.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">😔</div>
                <p>没有提取到数据</p>
            </div>
        `;
    } else {
        // 按当前顺序显示
        const orderedHeaders = headers.map((_, i) => headers[state.extractedColumnOrder[i]]);
        const orderedRows = resultRows.map(row =>
            state.extractedColumnOrder.map(i => row[i])
        );

        const previewData = orderedRows.slice(0, 100);
        const previewHtml = createTableHtml(orderedHeaders, previewData);
        resultPreview.innerHTML = previewHtml;

        if (orderedRows.length > 100) {
            resultPreview.innerHTML += `
                <div style="text-align: center; padding: 20px; color: #6b7280;">
                    预览显示前 100 行，共 ${orderedRows.length} 条数据
                </div>
            `;
        }
    }

    resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ==================== 列重排功能 ====================
function renderReorderColumns(headers) {
    const container = document.getElementById('reorderColumnsContainer');
    container.innerHTML = '';

    state.extractedColumnOrder.forEach((originalIndex, currentIndex) => {
        const column = document.createElement('div');
        column.className = 'reorder-column';
        column.dataset.index = currentIndex;
        column.dataset.originalIndex = originalIndex;
        column.textContent = headers[originalIndex];
        column.draggable = true;

        // 拖拽事件
        column.addEventListener('dragstart', handleDragStart);
        column.addEventListener('dragover', handleDragOver);
        column.addEventListener('drop', handleDrop);
        column.addEventListener('dragend', handleDragEnd);

        // 移动按钮
        const moveLeft = document.createElement('button');
        moveLeft.className = 'move-btn';
        moveLeft.innerHTML = '←';
        moveLeft.title = '向左移动';
        moveLeft.addEventListener('click', () => moveColumn(currentIndex, -1));

        const moveRight = document.createElement('button');
        moveRight.className = 'move-btn';
        moveRight.innerHTML = '→';
        moveRight.title = '向右移动';
        moveRight.addEventListener('click', () => moveColumn(currentIndex, 1));

        const controls = document.createElement('div');
        controls.className = 'column-controls';
        controls.appendChild(moveLeft);
        controls.appendChild(moveRight);

        column.appendChild(controls);
        container.appendChild(column);
    });
}

let draggedIndex = null;

function handleDragStart(e) {
    draggedIndex = parseInt(this.dataset.index);
    this.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
}

function handleDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    this.classList.add('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    const targetIndex = parseInt(this.dataset.index);

    if (draggedIndex !== null && draggedIndex !== targetIndex) {
        // 交换位置
        const temp = state.extractedColumnOrder[draggedIndex];
        state.extractedColumnOrder.splice(draggedIndex, 1);
        state.extractedColumnOrder.splice(targetIndex, 0, temp);

        // 重新显示结果
        showResults(state.extractedData, state.extractedHeaders, [], true);
    }
}

function handleDragEnd() {
    this.classList.remove('dragging');
    document.querySelectorAll('.reorder-column').forEach(col => {
        col.classList.remove('drag-over');
    });
    draggedIndex = null;
}

function moveColumn(index, direction) {
    const newIndex = index + direction;
    if (newIndex < 0 || newIndex >= state.extractedColumnOrder.length) return;

    // 交换位置
    const temp = state.extractedColumnOrder[index];
    state.extractedColumnOrder[index] = state.extractedColumnOrder[newIndex];
    state.extractedColumnOrder[newIndex] = temp;

    showResults(state.extractedData, state.extractedHeaders, [], true);
}

// ==================== 下载 ====================
function downloadResult() {
    if (!state.extractedData || !state.extractedHeaders) {
        showToast('没有数据可下载', 'error');
        return;
    }

    const format = document.getElementById('outputFormat').value;
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    const filename = `提取结果_${timestamp}`;

    // 按当前顺序组织数据
    const orderedHeaders = state.extractedHeaders.map((_, i) =>
        state.extractedHeaders[state.extractedColumnOrder[i]]
    );
    const orderedRows = state.extractedData.map(row =>
        state.extractedColumnOrder.map(i => row[i])
    );

    const sheetData = [orderedHeaders, ...orderedRows];

    try {
        if (format === 'xlsx') {
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '提取结果');
            XLSX.writeFile(wb, `${filename}.xlsx`);
        } else if (format === 'csv') {
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            const csv = XLSX.utils.sheet_to_csv(ws);
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `${filename}.csv`;
            link.click();
            URL.revokeObjectURL(link.href);
        }
        showToast('下载已开始', 'success');
    } catch (error) {
        showToast('下载失败: ' + error.message, 'error');
    }
}

// ==================== 重置 ====================
function resetApp() {
    state.sourceData = null;
    state.filterData = null;
    state.sourceHeaders = [];
    state.filterHeaders = [];
    state.extractedData = null;
    state.extractedHeaders = null;
    state.extractedColumnOrder = [];
    state.batchConfigs = [];
    state.sourceFileName = '';
    state.filterFileName = '';
    state.batchFileName = '';

    // 重置UI
    document.getElementById('sourceFileName').textContent = '选择文件';
    document.getElementById('sourceInfo').innerHTML = '';
    document.getElementById('sourcePreview').innerHTML = '';
    document.getElementById('sourceFile').value = '';

    document.getElementById('filterFileName').textContent = '选择文件';
    document.getElementById('filterInfo').innerHTML = '';
    document.getElementById('filterPreview').innerHTML = '';
    document.getElementById('filterFile').value = '';

    document.getElementById('batchFileName').textContent = '选择文件';
    document.getElementById('batchInfo').innerHTML = '';
    document.getElementById('batchPreview').innerHTML = '';
    document.getElementById('batchFile').value = '';

    document.getElementById('singleExtractSection').style.display = 'none';
    document.getElementById('batchExtractSection').style.display = 'none';
    document.getElementById('resultSection').style.display = 'none';
    document.getElementById('batchResultSection').style.display = 'none';

    showToast('已重置，请重新上传文件', 'success');
}

// ==================== 工具函数 ====================
function readExcelFile(file, onSuccess, onError) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (jsonData.length === 0) {
                onError('文件为空');
                return;
            }

            const headers = jsonData[0].map(h => String(h || ''));
            const rows = jsonData.slice(1).map(row => {
                const paddedRow = [...row];
                while (paddedRow.length < headers.length) {
                    paddedRow.push('');
                }
                return paddedRow.map(cell => String(cell || ''));
            });

            onSuccess(rows, headers);
        } catch (error) {
            onError('文件格式错误: ' + error.message);
        }
    };
    reader.onerror = () => onError('读取文件失败');
    reader.readAsArrayBuffer(file);
}

function createTableHtml(headers, rows) {
    let html = '<table><thead><tr>';
    headers.forEach(h => {
        html += `<th>${escapeHtml(h)}</th>`;
    });
    html += '</tr></thead><tbody>';

    rows.forEach((row, rowIndex) => {
        html += '<tr>';
        row.forEach(cell => {
            html += `<td>${escapeHtml(cell)}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = String(text || '');
    return div.innerHTML;
}

function showToast(message, type = 'success') {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.className = 'toast show ' + type;

    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}
