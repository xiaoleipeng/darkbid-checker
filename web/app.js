// app.js - Pyodide前端控制逻辑
let pyodide = null;
let resultBlob = null;
let kwMap = {};

async function initPyodide() {
    const bar = document.getElementById('loadingBar');
    const text = document.getElementById('loadingText');
    const sub = document.getElementById('loadingSub');

    text.textContent = '正在加载 Pyodide 运行环境...';
    sub.textContent = '步骤 1/3：下载 WebAssembly 核心';
    bar.style.width = '10%';

    pyodide = await loadPyodide();
    bar.style.width = '40%';

    text.textContent = '正在安装 Python 依赖包...';
    sub.textContent = '步骤 2/3：安装 python-docx';
    await pyodide.loadPackage(['micropip']);
    bar.style.width = '60%';

    await pyodide.runPythonAsync(`
import micropip
await micropip.install('python-docx')
    `);
    bar.style.width = '85%';

    text.textContent = '正在加载检查脚本...';
    sub.textContent = '步骤 3/3：初始化完成';
    const resp = await fetch('checker.py?v=' + Date.now());
    const code = await resp.text();
    await pyodide.runPythonAsync(code);
    bar.style.width = '100%';

    setTimeout(() => {
        document.getElementById('loading-overlay').classList.add('hidden');
    }, 300);
    document.getElementById('runBtn').disabled = false;
}

// 文件选择 - 点击和拖拽
const fileDrop = document.getElementById('fileDrop');
const fileInput = document.getElementById('fileInput');

fileDrop.addEventListener('click', () => fileInput.click());
fileDrop.addEventListener('dragover', (e) => { e.preventDefault(); fileDrop.classList.add('dragover'); });
fileDrop.addEventListener('dragleave', () => fileDrop.classList.remove('dragover'));
fileDrop.addEventListener('drop', (e) => {
    e.preventDefault();
    fileDrop.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.docx') || file.name.endsWith('.pdf'))) {
        fileInput.files = e.dataTransfer.files;
        showFileName(file);
    } else {
        alert('请选择 .docx 或 .pdf 格式的文件');
    }
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) showFileName(e.target.files[0]);
});

function showFileName(file) {
    const size = file.size < 1024*1024
        ? (file.size/1024).toFixed(0) + 'KB'
        : (file.size/1024/1024).toFixed(1) + 'MB';
    document.getElementById('fileName').textContent = `✓ ${file.name} (${size})`;
}

// 关键词编辑弹窗
function openKwModal() {
    const modal = document.createElement('div');
    modal.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);display:flex;align-items:center;justify-content:center;z-index:1000;';
    let rows = Object.entries(kwMap).map(([k,v]) => `<tr><td class="kw-cell" style="padding:8px 16px;border-bottom:1px solid #eee;cursor:text">${escHtml(k)}</td><td class="kw-cell" style="padding:8px 16px;border-bottom:1px solid #eee;cursor:text">${escHtml(v)}</td></tr>`).join('');
    modal.innerHTML = `<div style="background:#fff;border-radius:12px;padding:24px;max-width:520px;width:90%;max-height:75vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,0.3)">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
            <h3 style="font-size:16px;color:#333">关键词映射</h3>
            <span style="cursor:pointer;font-size:20px;color:#999;padding:4px 8px" id="closeModal">✕</span>
        </div>
        <div style="font-size:11px;color:#999;margin-bottom:12px">💡 双击单元格编辑，回车确认</div>
        <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead><tr><th style="text-align:left;padding:8px 16px;background:#f5f5f5;border-bottom:2px solid #e5e7eb">关键词</th><th style="text-align:left;padding:8px 16px;background:#f5f5f5;border-bottom:2px solid #e5e7eb">替换为</th></tr></thead>
            <tbody id="kwTableBody">${rows}</tbody>
        </table>
        <div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap">
            <label style="padding:6px 14px;background:#fff;color:#666;border:1px solid #ddd;border-radius:6px;font-size:12px;cursor:pointer">📄 导入txt<input type="file" accept=".txt" style="display:none" id="kwFileInModal"></label>
            <button id="kwAddBtn" style="padding:6px 14px;background:#fff;color:#2563eb;border:1px solid #2563eb;border-radius:6px;font-size:12px;cursor:pointer">+ 添加行</button>
            <button id="kwSaveBtn" style="padding:6px 14px;background:#2563eb;color:#fff;border:none;border-radius:6px;font-size:12px;cursor:pointer;margin-left:auto">保存</button>
        </div>
    </div>`;
    document.body.appendChild(modal);
    modal.querySelector('#closeModal').onclick = () => modal.remove();
    modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });

    function enableEdit(cell) {
        cell.addEventListener('dblclick', () => {
            if (cell.querySelector('input')) return;
            const orig = cell.textContent;
            const input = document.createElement('input');
            input.value = orig;
            input.style.cssText = 'width:100%;padding:4px 6px;border:1px solid #2563eb;border-radius:4px;font-size:13px;';
            cell.textContent = '';
            cell.appendChild(input);
            input.focus();
            input.select();
            const finish = () => { cell.textContent = input.value || orig; };
            input.addEventListener('keydown', (e) => { if (e.key === 'Enter') finish(); });
            input.addEventListener('blur', finish);
        });
    }
    modal.querySelectorAll('.kw-cell').forEach(enableEdit);

    // 导入txt文件
    modal.querySelector('#kwFileInModal').addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const text = await file.text();
        const tbody = modal.querySelector('#kwTableBody');
        text.split('\n').forEach(line => {
            line = line.trim();
            if (!line || line.startsWith('#')) return;
            let k, v;
            if (line.includes(',')) { [k, v] = line.split(',', 2); k = k.trim(); v = v.trim(); }
            else { k = line; v = '***'; }
            const tr = document.createElement('tr');
            tr.innerHTML = `<td class="kw-cell" style="padding:8px 16px;border-bottom:1px solid #eee;cursor:text">${escHtml(k)}</td><td class="kw-cell" style="padding:8px 16px;border-bottom:1px solid #eee;cursor:text">${escHtml(v)}</td>`;
            tbody.appendChild(tr);
            tr.querySelectorAll('.kw-cell').forEach(enableEdit);
        });
    });

    // 添加行
    modal.querySelector('#kwAddBtn').onclick = () => {
        const tbody = modal.querySelector('#kwTableBody');
        const tr = document.createElement('tr');
        tr.innerHTML = `<td class="kw-cell" style="padding:8px 16px;border-bottom:1px solid #eee;cursor:text">新关键词</td><td class="kw-cell" style="padding:8px 16px;border-bottom:1px solid #eee;cursor:text">***</td>`;
        tbody.appendChild(tr);
        tr.querySelectorAll('.kw-cell').forEach(enableEdit);
        tr.querySelector('.kw-cell').dispatchEvent(new Event('dblclick'));
    };

    // 保存
    modal.querySelector('#kwSaveBtn').onclick = () => {
        const newMap = {};
        modal.querySelectorAll('#kwTableBody tr').forEach(tr => {
            const cells = tr.querySelectorAll('td');
            const k = cells[0].textContent.trim();
            const v = cells[1].textContent.trim();
            if (k && k !== '新关键词') newMap[k] = v || '***';
        });
        kwMap = newMap;
        const info = document.getElementById('kwFileInfo');
        const count = Object.keys(kwMap).length;
        info.textContent = count > 0 ? `✓ 已配置 ${count} 条映射` : '未配置';
        info.style.cursor = count > 0 ? 'pointer' : 'default';
        modal.remove();
    };
}

document.getElementById('kwEditBtn').addEventListener('click', openKwModal);
document.getElementById('kwFileInfo').addEventListener('click', () => {
    if (Object.keys(kwMap).length > 0) openKwModal();
});

// 工具函数
function setStatus(type, msg) {
    const el = document.getElementById('status');
    el.className = 'status ' + type;
    el.innerHTML = msg;
}

function showProgress(show) {
    document.getElementById('progress').style.display = show ? 'block' : 'none';
}

function setProgress(pct) {
    document.getElementById('progressBar').style.width = pct + '%';
    // 同步更新 loading 状态中的百分比
    const el = document.getElementById('status');
    if (el.classList.contains('loading')) {
        const base = el.getAttribute('data-base-msg') || '';
        el.innerHTML = `${base} <strong>${pct}%</strong>`;
    }
}

function setLoading(msg) {
    const el = document.getElementById('status');
    el.className = 'status loading';
    el.setAttribute('data-base-msg', msg);
    el.innerHTML = `${msg} <strong>0%</strong>`;
}

function escHtml(s) {
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// 执行检查
async function runCheck() {
    const fileInput = document.getElementById('fileInput');
    if (!fileInput.files[0]) {
        alert('请先选择文件');
        return;
    }

    const file = fileInput.files[0];
    if (file.name.endsWith('.pdf')) {
        await runPdfCheck(file);
    } else {
        await runWordCheck(file);
    }
}

// PDF 检查逻辑
async function runPdfCheck(file) {
    const btn = document.getElementById('runBtn');
    btn.disabled = true;
    setLoading('⏳ 正在检查PDF文件...');
    showProgress(true);
    setProgress(10);
    document.getElementById('downloadBtn').style.display = 'none';
    document.getElementById('results').innerHTML = '';

    try {
        const arrayBuf = await file.arrayBuffer();
        // 缓存一份副本供后续修复使用（pdf.js 会 transfer 原 buffer）
        window._pdfFixData = {fileBuffer: arrayBuf.slice(0), fileName: file.name, issues: []};
        setProgress(20);

        const pdfjsLib = await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs');
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';

        const pdf = await pdfjsLib.getDocument({data: arrayBuf}).promise;
        setProgress(40);

        const categories = [];
        document.querySelectorAll('.categories input:checked').forEach(cb => categories.push(cb.value));
        const catSet = new Set(categories);

        const keywords = Object.keys(kwMap);
        const issues = [];
        const startTime = Date.now();

        const A4_W = 595.276, A4_H = 841.89, TOL = 3.0;
        const MARGIN_TOP = 70.87, MARGIN_BOTTOM = 56.69, MARGIN_LEFT = 56.69, MARGIN_RIGHT = 56.69;
        const EXPECTED_SIZE = 14.0, SIZE_TOL = 0.5;
        const FONT_ALIASES = ['宋体', 'SimSun', 'simsun', 'STSong', 'Songti', 'FangSong'];
        const EN_PUNCTS = new Set(',.;:!?()[]');

        for (let i = 0; i < pdf.numPages; i++) {
            const page = await pdf.getPage(i + 1);
            const pn = i + 1;
            const vp = page.getViewport({scale: 1});

            // page: 纸张尺寸
            if (catSet.has('page')) {
                if (Math.abs(vp.width - A4_W) > TOL || Math.abs(vp.height - A4_H) > TOL) {
                    issues.push(['page', pn, `纸张尺寸 ${(vp.width/72*25.4).toFixed(1)}×${(vp.height/72*25.4).toFixed(1)}mm，不是A4`]);
                }
            }

            const textContent = await page.getTextContent();
            const items = textContent.items;
            if (!items.length) continue;

            // page: 页边距估算
            if (catSet.has('page') && items.length > 0) {
                let minX = Infinity, minY = Infinity, maxX = 0, maxY = 0;
                items.forEach(item => {
                    if (!item.str || !item.str.trim()) return;
                    const tx = item.transform;
                    const x = tx[4], y = tx[5];
                    if (x < minX) minX = x;
                    if (x + item.width > maxX) maxX = x + item.width;
                    const yTop = vp.height - y;
                    if (yTop < minY) minY = yTop;
                    if (yTop + item.height > maxY) maxY = yTop + item.height;
                });
                if (minX < MARGIN_LEFT - TOL)
                    issues.push(['page', pn, `左边距约${(minX/72*25.4).toFixed(1)}mm，小于要求(20mm)`]);
                if (minY < MARGIN_TOP - TOL)
                    issues.push(['page', pn, `上边距约${(minY/72*25.4).toFixed(1)}mm，小于要求(25mm)`]);
                if (vp.width - maxX < MARGIN_RIGHT - TOL)
                    issues.push(['page', pn, `右边距约${((vp.width - maxX)/72*25.4).toFixed(1)}mm，小于要求(20mm)`]);
                if (vp.height - maxY < MARGIN_BOTTOM - TOL)
                    issues.push(['page', pn, `下边距约${((vp.height - maxY)/72*25.4).toFixed(1)}mm，小于要求(20mm)`]);
            }

            // 逐项检查 font/color/punct/identity
            const fontIssuesOnPage = new Set();
            const sizeIssuesOnPage = new Set();

            for (const item of items) {
                const text = item.str;
                if (!text || !text.trim()) continue;
                const fontSize = item.transform ? Math.abs(item.transform[3]) || Math.abs(item.transform[0]) : 0;
                const fontName = item.fontName || '';

                // font
                if (catSet.has('font')) {
                    const isExpected = FONT_ALIASES.some(a => fontName.toLowerCase().includes(a.toLowerCase()));
                    if (!isExpected && fontName && !fontIssuesOnPage.has(fontName)) {
                        fontIssuesOnPage.add(fontName);
                        issues.push(['font', pn, `字体「${fontName}」非宋体，内容: '${text.slice(0,15)}'`]);
                    }
                    if (fontSize > 0 && Math.abs(fontSize - EXPECTED_SIZE) > SIZE_TOL) {
                        const sizeKey = fontSize.toFixed(1);
                        if (!sizeIssuesOnPage.has(sizeKey)) {
                            sizeIssuesOnPage.add(sizeKey);
                            issues.push(['font', pn, `字号${fontSize.toFixed(1)}pt，要求14pt(四号)，内容: '${text.slice(0,15)}'`]);
                        }
                    }
                }

                // color
                if (catSet.has('color') && item.color) {
                    const [r, g, b] = item.color;
                    if (r !== undefined && (r !== 0 || g !== 0 || b !== 0)) {
                        const hex = `#${Math.round(r*255).toString(16).padStart(2,'0')}${Math.round(g*255).toString(16).padStart(2,'0')}${Math.round(b*255).toString(16).padStart(2,'0')}`.toUpperCase();
                        issues.push(['color', pn, `颜色${hex}非黑色，内容: '${text.slice(0,15)}'`]);
                    }
                }

                // punct
                if (catSet.has('punct')) {
                    const cleaned = text.replace(/\d[.,]\d/g, '');
                    const found = [...cleaned].filter(c => EN_PUNCTS.has(c));
                    if (found.length > 0) {
                        const unique = [...new Set(found)].sort();
                        issues.push(['punct', pn, `英文标点 ${unique.map(c => `'${c}'`).join(' ')}，内容: '${text.slice(0,20)}'`]);
                    }
                }

                // identity
                if (catSet.has('identity') && keywords.length > 0) {
                    for (const kw of keywords) {
                        if (text.includes(kw)) {
                            issues.push(['identity', pn, `发现身份信息「${kw}」，内容: '${text.slice(0,30)}'`]);
                        }
                    }
                }
            }

            setProgress(40 + Math.round(50 * (i + 1) / pdf.numPages));
        }

        // 去重
        const seen = new Set();
        const deduped = issues.filter(([cat, pg, msg]) => {
            const key = `${cat}|${pg}|${msg.slice(0,50)}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });

        setProgress(95);
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);

        if (deduped.length === 0) {
            setStatus('success', `✓ PDF检查通过，文档符合暗标编制要求！(${elapsed}s)`);
        } else {
            setStatus('warning', `⚠️ PDF共发现 ${deduped.length} 个问题 (${elapsed}s)。PDF无法自动修复，请在Word源文件中修改后重新导出。`);
            renderPdfResults(deduped, pdf, file);
        }
        setProgress(100);
        setTimeout(() => showProgress(false), 500);

    } catch (err) {
        setStatus('error', '❌ PDF处理失败: ' + err.message);
        showProgress(false);
        console.error(err);
    }

    btn.disabled = false;
}

function renderPdfResults(issues, pdfDoc, file) {
    const catNames = {page:'页面设置', spacing:'行间距', align:'对齐缩进', font:'字体字号', color:'颜色', punct:'标点符号', identity:'身份信息'};
    const fixableCats = new Set(['color', 'punct', 'identity']);
    const fixableCount = issues.filter(([cat]) => fixableCats.has(cat)).length;
    const unfixableCount = issues.length - fixableCount;

    let html = `<div class="results-header"><span>共 ${issues.length} 个问题</span></div>`;

    // 修复按钮区域
    if (fixableCount > 0) {
        html += `<div style="margin:12px 0;padding:14px;background:#fffbeb;border:1px solid #fbbf24;border-radius:8px;">
            <div style="font-size:13px;color:#92400e;margin-bottom:8px">
                <strong>⚠️ PDF修复风险提示：</strong>
                <ul style="margin:6px 0 0 16px;line-height:1.8">
                    <li>修复可能导致部分文字位置偏移或排版变化</li>
                    <li>修复后字体可能回退为默认字体</li>
                    <li>仅能修复：<strong>身份信息替换为***</strong>（${fixableCount}项可检测）</li>
                    <li>颜色和标点修复需使用命令行版（pymupdf），网页版暂不支持</li>
                    ${unfixableCount > 0 ? `<li>无法修复：字体/字号/行间距/页边距/对齐（${unfixableCount}项），请在Word中处理</li>` : ''}
                    <li>原始文件不会被修改，修复结果另存为新文件</li>
                </ul>
            </div>
            <label style="display:flex;align-items:center;gap:6px;font-size:13px;color:#92400e;cursor:pointer;margin-bottom:10px">
                <input type="checkbox" id="pdfFixConfirm"> 我已了解风险，确认尝试修复
            </label>
            <button id="pdfFixBtn" class="btn" disabled style="padding:8px 16px;background:#d97706;color:#fff;border:none;border-radius:6px;font-size:13px;cursor:pointer;opacity:0.5"
                onclick="executePdfFix()">🔧 尝试修复 PDF（${fixableCount}项可修复）</button>
        </div>`;
    }

    html += '<table><thead><tr><th style="width:15%">类别</th><th style="width:10%">页码</th><th>问题描述</th><th style="width:12%">状态</th></tr></thead><tbody>';
    issues.slice(0, 100).forEach(([cat, pg, msg]) => {
        const fixable = fixableCats.has(cat);
        const statusHtml = fixable
            ? '<span style="color:#d97706;font-size:11px">可修复</span>'
            : '<span style="color:#999;font-size:11px">不可修复</span>';
        html += `<tr><td><span class="tag tag-page">${catNames[cat] || cat}</span></td><td>第${pg}页</td><td>${escHtml(msg)}</td><td>${statusHtml}</td></tr>`;
    });
    if (issues.length > 100) {
        html += `<tr><td colspan="4" style="text-align:center;color:#999">... 还有 ${issues.length - 100} 项问题未显示</td></tr>`;
    }
    html += '</tbody></table>';
    document.getElementById('results').innerHTML = html;

    // 更新修复数据中的 issues
    if (window._pdfFixData) window._pdfFixData.issues = issues;

    // 勾选确认后启用按钮
    const checkbox = document.getElementById('pdfFixConfirm');
    const fixBtn = document.getElementById('pdfFixBtn');
    if (checkbox && fixBtn) {
        checkbox.addEventListener('change', () => {
            fixBtn.disabled = !checkbox.checked;
            fixBtn.style.opacity = checkbox.checked ? '1' : '0.5';
        });
    }
}

async function executePdfFix() {
    const {fileBuffer, fileName} = window._pdfFixData || {};
    if (!fileBuffer) return;

    const btn = document.getElementById('pdfFixBtn');
    btn.disabled = true;
    btn.textContent = '⏳ 正在修复...';

    try {
        // 加载 pdf-lib 用于修复
        const {PDFDocument, rgb} = await import('https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm');

        const pdfDoc = await PDFDocument.load(fileBuffer.slice(0));
        const pages = pdfDoc.getPages();
        let fixCount = 0;

        // 用 pdf.js 找到需要修复的文本位置
        const pdfjsLib = await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs');
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';
        const pdfjs = await pdfjsLib.getDocument({data: fileBuffer.slice(0)}).promise;

        const keywords = Object.keys(kwMap);
        console.log('[PDF修复] keywords:', keywords, 'pages:', pdfjs.numPages, 'bufferSize:', fileBuffer.byteLength);

        // PDF-lib 主要能做的修复有限，这里做文本覆盖式修复
        for (let i = 0; i < pdfjs.numPages; i++) {
            const pdfPage = await pdfjs.getPage(i + 1);
            const textContent = await pdfPage.getTextContent();
            const page = pages[i];
            const {height} = page.getSize();

            for (const item of textContent.items) {
                const text = item.str;
                if (!text || !text.trim()) continue;

                // 身份信息替换（用***替代，纯ASCII可正常写入）
                if (keywords.length > 0) {
                    for (const kw of keywords) {
                        if (text.includes(kw)) {
                            console.log(`[PDF修复] 找到匹配: "${kw}" in "${text.slice(0,30)}" page ${i+1}`);
                            const tx = item.transform;
                            const x = tx[4], y = tx[5];
                            const fontSize = Math.abs(tx[3]) || Math.abs(tx[0]) || 12;
                            const replacement = (kwMap[kw] || '***').replace(/[^\x00-\x7F]/g, '***');
                            page.drawRectangle({x, y: y - 2, width: item.width + 2, height: fontSize + 4, color: rgb(1,1,1), borderWidth: 0});
                            page.drawText(replacement, {x, y, size: fontSize, color: rgb(0,0,0)});
                            fixCount++;
                            break;
                        }
                    }
                }
            }
        }

        // 注意：标点修复和颜色修复因pdf-lib不支持中文字体，在网页版中无法实现
        // 请使用命令行版（pymupdf）进行完整修复

        if (fixCount > 0) {
            const pdfBytes = await pdfDoc.save();
            const blob = new Blob([pdfBytes], {type: 'application/pdf'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = fileName.replace('.pdf', '_已修复.pdf');
            a.click();
            URL.revokeObjectURL(a.href);
            btn.textContent = `✓ 已修复 ${fixCount} 处，文件已下载`;
            btn.style.background = '#16a34a';
        } else {
            btn.textContent = '未找到可修复内容';
            btn.style.background = '#999';
        }

    } catch (err) {
        btn.textContent = '❌ 修复失败: ' + err.message;
        btn.style.background = '#dc2626';
        console.error(err);
    }
}

const PUNCTUATION_MAP_JS = {',':'，', '.':'。', ':':'：', ';':'；', '!':'！', '?':'？', '(':'（', ')':'）', '[':'【', ']':'】'};

// Word 检查逻辑（原有逻辑）
async function runWordCheck(file) {

    const btn = document.getElementById('runBtn');
    btn.disabled = true;
    setLoading('⏳ 正在处理Word文件...');
    showProgress(true);
    setProgress(10);
    document.getElementById('downloadBtn').style.display = 'none';
    document.getElementById('results').innerHTML = '';

    try {
        // 清理上次残留文件
        try { pyodide.FS.unlink('/input.docx'); } catch(e) {}
        try { pyodide.FS.unlink('/output.docx'); } catch(e) {}

        const arrayBuf = await file.arrayBuffer();
        const uint8 = new Uint8Array(new Uint8Array(arrayBuf));
        if (uint8.length === 0) {
            throw new Error('文件读取为空，请重新选择文件');
        }
        try {
            pyodide.FS.writeFile('/input.docx', uint8);
        } catch (e) {
            throw new Error('文件写入失败，请检查文件是否为有效的 .docx 文件');
        }
        // 验证文件写入成功
        const stat = pyodide.FS.stat('/input.docx');
        if (stat.size === 0) {
            throw new Error('文件写入为空，请重新选择文件');
        }
        setProgress(30);

        const categories = [];
        document.querySelectorAll('.categories input:checked').forEach(cb => {
            categories.push(cb.value);
        });

        const addComments = document.getElementById('addComments').checked;

        setProgress(40);
        await new Promise(r => setTimeout(r, 0));
        const startTime = Date.now();

        // 核心处理会阻塞主线程，切换为不确定进度动画（CSS动画不受阻塞影响）
        document.getElementById('progressBar').classList.add('indeterminate');

        await pyodide.runPythonAsync(`
import json

keyword_map = json.loads('${JSON.stringify(kwMap).replace(/\\/g, '\\\\').replace(/'/g, "\\'")}')
categories = set(json.loads('${JSON.stringify(categories)}'))
add_comments = ${addComments ? 'True' : 'False'}
`);
        await new Promise(r => setTimeout(r, 0));

        await pyodide.runPythonAsync(`
details, fix_count, warn_count = check_fix_annotate(
    '/input.docx', '/output.docx',
    keyword_map=keyword_map if keyword_map else None,
    categories=categories,
    add_comments=add_comments
)
`);

        // 核心处理完成，恢复确定进度
        document.getElementById('progressBar').classList.remove('indeterminate');
        setProgress(85);
        await new Promise(r => setTimeout(r, 0));

        await pyodide.runPythonAsync(`
result_json = json.dumps({
    'fix_count': fix_count,
    'warn_count': warn_count,
    'details': details
}, ensure_ascii=False)
        `);

        setProgress(80);

        const resultJson = pyodide.globals.get('result_json');
        const result = JSON.parse(resultJson);
        const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);

        setProgress(90);

        if (result.fix_count === 0 && result.warn_count === 0) {
            setStatus('success', `✓ 所有检查通过，文档符合暗标编制要求！(${elapsed}s)`);
            showProgress(false);
        } else {
            const outputData = pyodide.FS.readFile('/output.docx');
            resultBlob = new Blob([outputData], {type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});

            setStatus('success', `✓ 已修复 ${result.fix_count} 项问题` + (result.warn_count ? `，${result.warn_count} 项警告` : '') + ` (${elapsed}s)`);
            document.getElementById('downloadBtn').style.display = 'inline-flex';
            renderResults(result.details);
        }
        setProgress(100);
        setTimeout(() => showProgress(false), 500);

    } catch (err) {
        setStatus('error', '❌ 处理失败: ' + err.message);
        showProgress(false);
        console.error(err);
    }

    btn.disabled = false;
}

function renderResults(details, page = 1, pageSize = 50) {
    if (!details.length) return;
    window._resultDetails = details;
    const totalPages = Math.ceil(details.length / pageSize);
    const start = (page - 1) * pageSize;
    const end = Math.min(start + pageSize, details.length);
    const pageItems = details.slice(start, end);

    let html = `<div class="results-header"><span>共 ${details.length} 项修改明细（第${page}/${totalPages}页，${start+1}-${end}条）</span></div>`;
    html += '<table><thead><tr><th style="width:35%">位置</th><th>修复内容</th></tr></thead><tbody>';
    pageItems.forEach(([loc, fix]) => {
        let locHtml;
        if (loc.includes('[表格内]')) {
            locHtml = `<span class="tag tag-table">表格</span> ${escHtml(loc.replace('[表格内] ', ''))}`;
        } else if (loc === '页面设置') {
            locHtml = `<span class="tag tag-page">页面</span> 全局设置`;
        } else {
            locHtml = escHtml(loc);
        }
        html += `<tr><td>${locHtml}</td><td>${escHtml(fix)}</td></tr>`;
    });
    html += '</tbody></table>';

    if (totalPages > 1) {
        html += '<div style="display:flex;justify-content:center;align-items:center;gap:8px;margin-top:12px;flex-wrap:wrap">';
        html += `<button class="page-btn" onclick="renderResults(window._resultDetails,1)" ${page===1?'disabled':''}>&laquo;</button>`;
        html += `<button class="page-btn" onclick="renderResults(window._resultDetails,${page-1})" ${page===1?'disabled':''}>‹</button>`;
        // 显示页码
        let startP = Math.max(1, page - 2), endP = Math.min(totalPages, page + 2);
        for (let p = startP; p <= endP; p++) {
            html += `<button class="page-btn${p===page?' active':''}" onclick="renderResults(window._resultDetails,${p})">${p}</button>`;
        }
        html += `<button class="page-btn" onclick="renderResults(window._resultDetails,${page+1})" ${page===totalPages?'disabled':''}>›</button>`;
        html += `<button class="page-btn" onclick="renderResults(window._resultDetails,${totalPages})" ${page===totalPages?'disabled':''}>&raquo;</button>`;
        html += '</div>';
    }

    document.getElementById('results').innerHTML = html;
}

function downloadResult() {
    if (!resultBlob) return;
    const fileInput = document.getElementById('fileInput');
    const origName = fileInput.files[0].name.replace('.docx', '');
    const categories = [];
    document.querySelectorAll('.categories input:checked').forEach(cb => {
        categories.push(cb.value);
    });
    const allCats = ['page','spacing','align','font','color','punct','identity'];
    const suffix = categories.length === allCats.length ? '_已修正' : '_' + categories.sort().join('+');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(resultBlob);
    a.download = origName + suffix + '.docx';
    a.click();
    URL.revokeObjectURL(a.href);
}

// 启动
initPyodide();

// 规则弹窗
const ruleTip = document.getElementById('ruleTip');
const rulePopup = document.getElementById('rulePopup');
ruleTip.addEventListener('click', (e) => {
    e.stopPropagation();
    rulePopup.classList.toggle('show');
});
document.addEventListener('click', (e) => {
    if (!rulePopup.contains(e.target) && e.target !== ruleTip) {
        rulePopup.classList.remove('show');
    }
});
