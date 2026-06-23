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
    const resp = await fetch('checker.py');
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
    if (file && file.name.endsWith('.docx')) {
        fileInput.files = e.dataTransfer.files;
        showFileName(file);
    } else {
        alert('请选择 .docx 格式的文件');
    }
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) showFileName(e.target.files[0]);
});

function showFileName(file) {
    document.getElementById('fileName').textContent = `✓ ${file.name} (${(file.size/1024/1024).toFixed(1)}MB)`;
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
    el.textContent = msg;
}

function showProgress(show) {
    document.getElementById('progress').style.display = show ? 'block' : 'none';
}

function setProgress(pct) {
    document.getElementById('progressBar').style.width = pct + '%';
}

function escHtml(s) {
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// 执行检查
async function runCheck() {
    const fileInput = document.getElementById('fileInput');
    if (!fileInput.files[0]) {
        alert('请先选择Word文件');
        return;
    }

    const btn = document.getElementById('runBtn');
    btn.disabled = true;
    setStatus('loading', '⏳ 正在处理...');
    showProgress(true);
    setProgress(10);
    document.getElementById('downloadBtn').style.display = 'none';
    document.getElementById('results').innerHTML = '';

    try {
        const file = fileInput.files[0];
        const arrayBuf = await file.arrayBuffer();
        const uint8 = new Uint8Array(arrayBuf);
        pyodide.FS.writeFile('/input.docx', uint8);
        setProgress(30);

        const categories = [];
        document.querySelectorAll('.categories input:checked').forEach(cb => {
            categories.push(cb.value);
        });

        const addComments = document.getElementById('addComments').checked;

        setProgress(50);
        await pyodide.runPythonAsync(`
import json
keyword_map = json.loads('${JSON.stringify(kwMap).replace(/'/g, "\\'")}')
categories = set(json.loads('${JSON.stringify(categories)}'))
add_comments = ${addComments ? 'True' : 'False'}

details, fix_count, warn_count = check_fix_annotate(
    '/input.docx', '/output.docx',
    keyword_map=keyword_map if keyword_map else None,
    categories=categories,
    add_comments=add_comments
)

result_json = json.dumps({
    'fix_count': fix_count,
    'warn_count': warn_count,
    'details': details
}, ensure_ascii=False)
        `);

        setProgress(80);

        const resultJson = pyodide.globals.get('result_json');
        const result = JSON.parse(resultJson);

        setProgress(90);

        if (result.fix_count === 0 && result.warn_count === 0) {
            setStatus('success', '✓ 所有检查通过，文档符合暗标编制要求！');
            showProgress(false);
        } else {
            const outputData = pyodide.FS.readFile('/output.docx');
            resultBlob = new Blob([outputData], {type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});

            setStatus('success', `✓ 已修复 ${result.fix_count} 项问题` + (result.warn_count ? `，${result.warn_count} 项警告` : ''));
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

function renderResults(details) {
    if (!details.length) return;
    let html = `<div class="results-header"><span>共 ${details.length} 项修改明细</span></div>`;
    html += '<table><thead><tr><th style="width:35%">位置</th><th>修复内容</th></tr></thead><tbody>';
    details.forEach(([loc, fix]) => {
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
