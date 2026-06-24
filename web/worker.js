// worker.js - Web Worker 中运行 Pyodide 处理 Word 文件
importScripts('https://cdn.jsdelivr.net/pyodide/v0.25.1/full/pyodide.js');

let pyodide = null;

async function initPyodide() {
    postMessage({type: 'loading', step: 1, msg: '下载 WebAssembly 核心'});
    pyodide = await loadPyodide();
    postMessage({type: 'loading', step: 2, msg: '安装 python-docx'});
    await pyodide.loadPackage(['micropip']);
    await pyodide.runPythonAsync(`
import micropip
await micropip.install('python-docx')
    `);
    postMessage({type: 'loading', step: 3, msg: '加载检查脚本'});
    const resp = await fetch('checker.py?v=' + Date.now());
    const code = await resp.text();
    await pyodide.runPythonAsync(code);
    postMessage({type: 'ready'});
}

async function processFile(data) {
    const {fileBuffer, kwMap, categories, addComments} = data;

    try {
        // 写入文件
        const uint8 = new Uint8Array(fileBuffer);
        try { pyodide.FS.unlink('/input.docx'); } catch(e) {}
        try { pyodide.FS.unlink('/output.docx'); } catch(e) {}
        pyodide.FS.writeFile('/input.docx', uint8);

        postMessage({type: 'progress', pct: 10, msg: '文件已加载'});

        // 注册进度回调（Python 调用此函数时 Worker 线程可自由 postMessage）
        pyodide.globals.set('_worker_progress', (pct, msg) => {
            postMessage({type: 'progress', pct: Math.round(pct), msg: msg || ''});
        });

        // 执行处理
        await pyodide.runPythonAsync(`
import json

keyword_map = json.loads('''${JSON.stringify(kwMap).replace(/\\/g, '\\\\').replace(/'/g, "\\'")}''')
categories = set(json.loads('${JSON.stringify(categories)}'))
add_comments = ${addComments ? 'True' : 'False'}

details, fix_count, warn_count = check_fix_annotate(
    '/input.docx', '/output.docx',
    keyword_map=keyword_map if keyword_map else None,
    categories=categories,
    add_comments=add_comments,
    progress_callback=_worker_progress
)

result_json = json.dumps({
    'fix_count': fix_count,
    'warn_count': warn_count,
    'details': details
}, ensure_ascii=False)
        `);

        const resultJson = pyodide.globals.get('result_json');
        const result = JSON.parse(resultJson);

        let outputBuffer = null;
        if (result.fix_count > 0 || result.warn_count > 0) {
            const outputData = pyodide.FS.readFile('/output.docx');
            outputBuffer = outputData.buffer;
        }

        postMessage({type: 'done', result, outputBuffer}, outputBuffer ? [outputBuffer] : []);

    } catch (err) {
        postMessage({type: 'error', msg: err.message});
    }
}

// 监听主线程消息
self.onmessage = async (e) => {
    const {type} = e.data;
    if (type === 'init') {
        await initPyodide();
    } else if (type === 'process') {
        await processFile(e.data);
    }
};
