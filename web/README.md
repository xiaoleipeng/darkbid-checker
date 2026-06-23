# Word暗标编制要求检查工具 - 网页版

## 在线使用

部署到 GitHub Pages 后，直接打开网页即可使用。

🔒 **隐私安全**：所有文件仅在您的浏览器本地处理（基于Pyodide/WebAssembly），不会上传到任何服务器。

## 部署到 GitHub Pages

1. 将 `web/` 目录下的文件推送到 GitHub 仓库
2. 在仓库 Settings → Pages 中，选择部署来源为 `main` 分支的 `/web` 目录（或 root）
3. 等待部署完成后访问 `https://<用户名>.github.io/<仓库名>/`

## 本地开发测试

```bash
cd web/
python3 serve.py
# 浏览器自动打开 http://localhost:8000
```

## 功能

- ✅ 页面设置检查修复（A4纸张、页边距、页眉页脚页码）
- ✅ 行间距固定值30磅、段前段后间距归零
- ✅ 左对齐、首行缩进2字符、删除开头空格
- ✅ 字体宋体四号、去除加粗/倾斜/下划线（表格内仅检查样式不改字体字号）
- ✅ 文字颜色统一黑色
- ✅ 英文标点→中文标点
- ✅ 身份信息关键词替换（支持导入映射文件）
- ✅ 可选添加批注说明

## 技术实现

- 前端：HTML + JavaScript
- Python运行时：[Pyodide](https://pyodide.org/) (WebAssembly)
- docx处理：python-docx（在浏览器中通过micropip安装）
- 部署：GitHub Pages（纯静态托管）
