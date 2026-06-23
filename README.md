# darkbid-checker

Word投标文件技术标（暗标）格式自动检查与修复工具。

## 功能

自动检测并修复暗标编制中不符合规范的格式问题，支持一键修复+批注说明。

| 检查类别 | 内容 |
|---------|------|
| page | 纸张A4、页边距（上2.5cm/其余2cm）、清除页眉页脚页码、检测目录 |
| spacing | 行间距固定值30磅、段前段后间距归零 |
| align | 左对齐、首行缩进2字符、删除开头空格 |
| font | 宋体四号常规、去除加粗/倾斜/下划线（表格内仅检查样式不改字体字号） |
| color | 所有文字颜色→黑色（含表格内） |
| punct | 英文标点→中文标点 |
| identity | 投标人名称等身份信息替换 |

## 使用方式

### 方式一：网页版（推荐）

在线访问：`https://<your-username>.github.io/darkbid-checker/`

- 打开网页即可使用，无需安装任何软件
- 所有文件仅在浏览器本地处理，**不会上传任何服务器**
- 基于 Pyodide（浏览器端 Python + python-docx），处理精度与本地版一致

### 方式二：本地GUI

```bash
pip install -r requirements.txt
python3 check_word_gui.py
```

### 方式三：命令行

```bash
# 全部修复
python3 check_word_comment.py 文件.docx

# 指定类别修复
python3 check_word_comment.py 文件.docx --only page,spacing,font

# 带身份信息检测
python3 check_word_comment.py 文件.docx --keywords "XX公司,张三"
```

## 安装

```bash
git clone https://github.com/<your-username>/darkbid-checker.git
cd darkbid-checker
pip install -r requirements.txt
```

依赖：`python-docx >= 1.2.0`，Python 3.8+

## 部署网页版到 GitHub Pages

1. 将代码推送到 GitHub
2. Settings → Pages → Source 选 `main` 分支，目录选 `/web`
3. 等待部署完成，访问 `https://<用户名>.github.io/darkbid-checker/`

## 本地开发测试网页版

```bash
cd web/
python3 serve.py
# 浏览器自动打开 http://localhost:8080
```

## 项目结构

```
darkbid-checker/
├── web/                      # 网页版（可部署到GitHub Pages）
│   ├── index.html            # 页面
│   ├── app.js                # 前端交互逻辑
│   ├── checker.py            # Python检查核心（Pyodide中运行）
│   └── serve.py              # 本地测试服务器
├── check_word_gui.py         # 桌面GUI版（tkinter）
├── check_word_comment.py     # 命令行版（自动修复+批注）
├── check_word.py             # 命令行版（仅检查报告）
├── keywords_example.txt      # 关键词映射文件示例
├── requirements.txt          # Python依赖
├── DESIGN.md                 # 设计文档
└── README.md
```

## 关键词映射文件格式

```txt
# 注释行
XX科技有限公司,投标人
张三,项目经理
李四,技术负责人
公司logo,***
```

每行格式：`关键词,替换文本`。没有逗号的行默认替换为 `***`。

## 隐私安全

- **网页版**：基于 Pyodide (WebAssembly)，Python 完全在浏览器中运行，文件不离开本地
- **本地版**：文件仅在本机处理，不产生网络请求
- 原始文件不会被修改，修复结果另存为新文件

## 已知限制

- 页数统计为估算值（python-docx 无法获取精确渲染页数）
- 封面页无法程序化判断，需用户自行确认
- 图片/Logo 中的身份信息无法自动检测
- 网页版首次加载需下载约10MB运行环境（有浏览器缓存）

## License

MIT
