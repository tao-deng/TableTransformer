// 存储数据
let sourceData = null;
let filterData = null;
let sourceHeaders = [];
let filterHeaders = [];
let extractedData = null;

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('sourceFile').addEventListener('change', handleSourceFile);
    document.getElementById('filterFile').addEventListener('change', handleFilterFile);
    document.getElementById('extractBtn').addEventListener('click', extractData);
    document.getElementById('downloadBtn').addEventListener('click', downloadResult);
});

// 处理源文件上传
function handleSourceFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    document.getElementById('sourceFileName').textContent = file.name;
    readExcelFile(file, (data, headers) => {
        sourceData = data;
        sourceHeaders = headers;
        document.getElementById('sourceInfo').innerHTML = `✓ 已加载 ${data.length} 行，${headers.length} 列`;
        document.getElementById('sourceInfo').classList.remove('error');
        checkBothFilesLoaded();
    }, (error) => {
        document.getElementById('sourceInfo').innerHTML = `✗ ${error}`;
        document.getElementById('sourceInfo').classList.add('error');
    });
}

// 处理筛选文件上传
function handleFilterFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    document.getElementById('filterFileName').textContent = file.name;
    readExcelFile(file, (data, headers) => {
        filterData = data;
        filterHeaders = headers;
        document.getElementById('filterInfo').innerHTML = `✓ 已加载 ${data.length} 行，${headers.length} 列`;
        document.getElementById('filterInfo').classList.remove('error');
        checkBothFilesLoaded();
    }, (error) => {
        document.getElementById('filterInfo').innerHTML = `✗ ${error}`;
        document.getElementById('filterInfo').classList.add('error');
    });
}

// 检查两个文件都已加载
function checkBothFilesLoaded() {
    if (sourceData && filterData) {
        populateColumnSelects();
        document.getElementById('optionsSection').style.display = 'block';
        updateFilterPreview();
        document.getElementById('previewSection').style.display = 'block';
        document.getElementById('actionSection').style.display = 'block';
    }
}

// 填充列选择下拉框
function populateColumnSelects() {
    const sourceSelect = document.getElementById('sourceColumn');
    const filterSelect = document.getElementById('filterColumn');

    sourceSelect.innerHTML = '';
    filterSelect.innerHTML = '';

    sourceHeaders.forEach((header, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = header || `列 ${index + 1}`;
        sourceSelect.appendChild(option);
    });

    filterHeaders.forEach((header, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = header || `列 ${index + 1}`;
        filterSelect.appendChild(option);
    });

    // 默认选择第一列
    sourceSelect.value = 0;
    filterSelect.value = 0;
}

// 更新筛选条件预览
function updateFilterPreview() {
    const filterColumnIndex = parseInt(document.getElementById('filterColumn').value);
    const preview = document.getElementById('filterPreview');
    const container = document.createElement('div');
    container.className = 'filter-tags';

    const filterValues = filterData.map(row => row[filterColumnIndex]).filter(v => v);
    const uniqueValues = [...new Set(filterValues)].slice(0, 50); // 最多显示50个

    uniqueValues.forEach(value => {
        const tag = document.createElement('span');
        tag.className = 'filter-tag';
        tag.textContent = value;
        container.appendChild(tag);
    });

    if (filterValues.length > 50) {
        const more = document.createElement('span');
        more.className = 'filter-tag';
        more.textContent = `...还有 ${filterValues.length - 50} 个`;
        container.appendChild(more);
    }

    preview.innerHTML = '';
    preview.appendChild(container);
}

// 监听筛选列变化
document.getElementById('filterColumn').addEventListener('change', updateFilterPreview);

// 读取 Excel 文件
function readExcelFile(file, onSuccess, onError) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // 读取第一个工作表
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // 转换为 JSON
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (jsonData.length === 0) {
                onError('文件为空');
                return;
            }

            // 第一行为表头
            const headers = jsonData[0].map(h => String(h || ''));
            const rows = jsonData.slice(1).map(row => {
                // 确保每行有和表头一样多的列
                const paddedRow = [...row];
                while (paddedRow.length < headers.length) {
                    paddedRow.push('');
                }
                return paddedRow.map(cell => String(cell || ''));
            });

            onSuccess(rows, headers);
        } catch (error) {
            onError('文件格式错误，请检查文件');
        }
    };
    reader.onerror = () => onError('读取文件失败');
    reader.readAsArrayBuffer(file);
}

// 提取数据
function extractData() {
    const sourceColumnIndex = parseInt(document.getElementById('sourceColumn').value);
    const filterColumnIndex = parseInt(document.getElementById('filterColumn').value);
    const matchMode = document.getElementById('matchMode').value;

    // 获取筛选值（去重）
    const filterValues = new Set(
        filterData
            .map(row => row[filterColumnIndex])
            .filter(v => v && v.trim() !== '')
            .map(v => v.trim())
    );

    // 提取匹配的行
    const resultRows = sourceData.filter(row => {
        const sourceValue = (row[sourceColumnIndex] || '').trim();

        switch (matchMode) {
            case 'exact':
                return filterValues.has(sourceValue);
            case 'contain':
                return Array.from(filterValues).some(fv => sourceValue.includes(fv));
            case 'startsWith':
                return Array.from(filterValues).some(fv => sourceValue.startsWith(fv));
            default:
                return filterValues.has(sourceValue);
        }
    });

    extractedData = resultRows;

    // 显示结果
    document.getElementById('resultSection').style.display = 'block';
    document.getElementById('resultStats').textContent = `找到 ${resultRows.length} 条匹配数据`;

    // 预览前 100 行
    const previewData = resultRows.slice(0, 100);
    const previewHtml = createTableHtml(sourceHeaders, previewData);
    document.getElementById('resultPreview').innerHTML = previewHtml;

    if (resultRows.length > 100) {
        document.getElementById('resultStats').textContent += ` (预览显示前 100 行)`;
    }
}

// 创建表格 HTML
function createTableHtml(headers, rows) {
    let html = '<table><thead><tr>';
    headers.forEach(h => {
        html += `<th>${escapeHtml(h)}</th>`;
    });
    html += '</tr></thead><tbody>';

    rows.forEach(row => {
        html += '<tr>';
        row.forEach(cell => {
            html += `<td>${escapeHtml(cell)}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
}

// HTML 转义义
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// 下载结果
function downloadResult() {
    if (!extractedData || extractedData.length === 0) {
        alert('没有数据可下载');
        return;
    }

    const format = document.getElementById('outputFormat').value;
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    let filename = `提取结果_${timestamp}`;
    let content;

    // 创建包含表头的工作表数据
    const sheetData = [sourceHeaders, ...extractedData];

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
}
