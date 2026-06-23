"""
Word暗标编制要求检查工具 - Pyodide版核心逻辑
在浏览器中通过Pyodide运行，提供 check_fix_annotate() 函数。
"""
import io
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

# 常量
FONT_NAME = '宋体'
FONT_SIZE = Pt(14)
PAGE_WIDTH = Cm(21.0)
PAGE_HEIGHT = Cm(29.7)
MARGIN_TOP = Cm(2.5)
MARGIN_BOTTOM = Cm(2.0)
MARGIN_LEFT = Cm(2.0)
MARGIN_RIGHT = Cm(2.0)
LINE_SPACING_VAL = Pt(30)
BLACK_COLOR = RGBColor(0, 0, 0)
AUTHOR = "暗标检查工具"
INITIALS = "AH"

PUNCTUATION_MAP = {
    ',': '，', '.': '。', ':': '：', ';': '；',
    '!': '！', '?': '？', '(': '（', ')': '）',
    '[': '【', ']': '】',
}
PUNCT_SET = set(PUNCTUATION_MAP.keys())
DIGIT_PUNCT_PATTERN = re.compile(r'\d[.,]\d')

# 预计算 qn 值避免重复调用
QN_WPR = qn('w:rPr')
QN_WFONTS = qn('w:rFonts')
QN_EASTASIA = qn('w:eastAsia')
QN_INSTR = qn('w:instrText')
QN_TC = qn('w:tc')
QN_TBL = qn('w:tbl')

# 批注合并阈值：超过此数量的修复不逐条加批注
COMMENT_LIMIT = 200
# details 最大条目数
DETAILS_LIMIT = 500


def _set_font_name(run):
    run.font.name = FONT_NAME
    rpr = run._element.find(QN_WPR)
    if rpr is None:
        rpr = run._element.makeelement(QN_WPR, {})
        run._element.insert(0, rpr)
    rFonts = rpr.find(QN_WFONTS)
    if rFonts is None:
        rFonts = rpr.makeelement(QN_WFONTS, {})
        rpr.insert(0, rFonts)
    rFonts.set(QN_EASTASIA, FONT_NAME)


def _fix_punctuation_in_run(run):
    text = run.text
    if not any(ch in PUNCT_SET for ch in text):
        return
    result = []
    for i, ch in enumerate(text):
        if ch in PUNCTUATION_MAP:
            if ch in '.,':
                before_digit = i > 0 and text[i-1].isdigit()
                after_digit = i < len(text)-1 and text[i+1].isdigit()
                if before_digit and after_digit:
                    result.append(ch)
                    continue
            result.append(PUNCTUATION_MAP[ch])
        else:
            result.append(ch)
    run.text = ''.join(result)


def _process_runs_font_color(runs, in_table, categories, fix_count, details, details_full):
    """批量处理 runs 的字体和颜色，返回更新后的 fix_count"""
    check_font = 'font' in categories
    check_color = 'color' in categories

    for run in runs:
        if not run.text.strip():
            continue
        font = run.font
        fixed = False

        if check_font:
            if not in_table:
                # 字体名
                if font.name and font.name != FONT_NAME:
                    _set_font_name(run)
                    fix_count += 1
                    fixed = True
                else:
                    rpr = run._element.find(QN_WPR)
                    if rpr is not None:
                        rFonts = rpr.find(QN_WFONTS)
                        if rFonts is not None:
                            ea = rFonts.get(QN_EASTASIA)
                            if ea and ea != FONT_NAME:
                                _set_font_name(run)
                                fix_count += 1
                                fixed = True
                # 字号
                if font.size and font.size != FONT_SIZE:
                    font.size = FONT_SIZE
                    fix_count += 1
                    fixed = True
            # 加粗/倾斜/下划线（表格内外均检查）
            if font.bold:
                font.bold = False
                fix_count += 1
                fixed = True
            if font.italic:
                font.italic = False
                fix_count += 1
                fixed = True
            if font.underline:
                font.underline = False
                fix_count += 1
                fixed = True

        if check_color:
            if font.color and font.color.rgb and font.color.rgb != BLACK_COLOR:
                font.color.rgb = BLACK_COLOR
                fix_count += 1
                fixed = True

        if fixed and details_full:
            run_loc = ("[表格内] " if in_table else "") + (run.text[:15] + "..." if len(run.text) > 15 else run.text)
            details.append((run_loc, "字体/颜色修复"))

    return fix_count


def check_fix_annotate(doc_path, output_path, keyword_map=None, categories=None, add_comments=True):
    """核心修复逻辑，返回 (details, fix_count, warn_count)"""
    ALL_CATEGORIES = {'page', 'spacing', 'align', 'font', 'color', 'punct', 'identity'}
    if categories is None:
        categories = ALL_CATEGORIES

    # 支持路径或BytesIO
    if isinstance(doc_path, str):
        with open(doc_path, 'rb') as f:
            doc = Document(io.BytesIO(f.read()))
    else:
        doc = Document(doc_path)

    fix_count = 0
    warn_count = 0
    details = []
    details_full = True  # 当 details 太多时停止记录明细

    # 收集表格内所有 paragraph 元素的 id，用于快速判断
    table_para_ids = set()
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    table_para_ids.add(id(para))

    first_para = doc.paragraphs[0] if doc.paragraphs else None

    # TOC目录检测 + 页面设置（合并到一次遍历前处理）
    if 'page' in categories:
        # TOC 检测只扫前50段（目录通常在文档开头）
        for para in doc.paragraphs[:50]:
            for instr in para._element.findall('.//' + QN_INSTR):
                if instr.text and 'TOC' in instr.text.upper():
                    details.append(("目录", "⚠️ 存在自动目录，需手动删除"))
                    warn_count += 1
                    break

        # 页面设置
        page_fixes = []
        tolerance = Cm(0.1)
        margin_tolerance = Cm(0.05)
        for section in doc.sections:
            if section.page_width and abs(section.page_width - PAGE_WIDTH) > tolerance:
                page_fixes.append(f"纸张宽度 → 21.0cm")
                section.page_width = PAGE_WIDTH
            if section.page_height and abs(section.page_height - PAGE_HEIGHT) > tolerance:
                page_fixes.append(f"纸张高度 → 29.7cm")
                section.page_height = PAGE_HEIGHT
            if section.top_margin is not None and abs(section.top_margin - MARGIN_TOP) > margin_tolerance:
                page_fixes.append(f"上边距 → 2.5cm")
                section.top_margin = MARGIN_TOP
            if section.bottom_margin is not None and abs(section.bottom_margin - MARGIN_BOTTOM) > margin_tolerance:
                page_fixes.append(f"下边距 → 2.0cm")
                section.bottom_margin = MARGIN_BOTTOM
            if section.left_margin is not None and abs(section.left_margin - MARGIN_LEFT) > margin_tolerance:
                page_fixes.append(f"左边距 → 2.0cm")
                section.left_margin = MARGIN_LEFT
            if section.right_margin is not None and abs(section.right_margin - MARGIN_RIGHT) > margin_tolerance:
                page_fixes.append(f"右边距 → 2.0cm")
                section.right_margin = MARGIN_RIGHT
            if section.header and section.header.is_linked_to_previous is False:
                if ''.join(p.text for p in section.header.paragraphs).strip():
                    page_fixes.append("已清除页眉")
                    for p in section.header.paragraphs:
                        p.clear()
            if section.footer and section.footer.is_linked_to_previous is False:
                if ''.join(p.text for p in section.footer.paragraphs).strip():
                    page_fixes.append("已清除页脚/页码")
                    for p in section.footer.paragraphs:
                        p.clear()
        if page_fixes:
            fix_count += len(page_fixes)
            for f in page_fixes:
                details.append(("页面设置", f))

    # 页数估算
    total_paras = len(doc.paragraphs)
    estimated_pages = total_paras // 25
    if estimated_pages > 300:
        details.append(("全文", f"⚠️ 估算页数约{estimated_pages}页，可能超过300页限制"))
        warn_count += 1

    # 预编译关键词检查
    kw_items = list(keyword_map.items()) if keyword_map and 'identity' in categories else []

    check_spacing = 'spacing' in categories
    check_align = 'align' in categories
    check_punct = 'punct' in categories
    check_font_or_color = 'font' in categories or 'color' in categories

    # 主循环：逐段处理（跳过表格内段落，后面单独处理）
    for para in doc.paragraphs:
        text = para.text
        if not text.strip() or not para.runs:
            continue

        in_table = id(para) in table_para_ids
        if in_table:
            continue  # 表格内段落在下面统一处理

        para_fixed = False

        if check_spacing:
            pf = para.paragraph_format
            need_fix = False
            if pf.line_spacing_rule != WD_LINE_SPACING.EXACTLY:
                need_fix = True
            elif pf.line_spacing is None or abs(pf.line_spacing - LINE_SPACING_VAL) > Pt(0.5):
                need_fix = True
            if need_fix:
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                pf.line_spacing = LINE_SPACING_VAL
                fix_count += 1
                para_fixed = True
            if pf.space_before and pf.space_before > Pt(0):
                pf.space_before = Pt(0)
                fix_count += 1
                para_fixed = True
            if pf.space_after and pf.space_after > Pt(0):
                pf.space_after = Pt(0)
                fix_count += 1
                para_fixed = True

        if check_align:
            pf = para.paragraph_format
            if pf.alignment is not None and pf.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                fix_count += 1
                para_fixed = True
            if pf.first_line_indent is None or abs(pf.first_line_indent - Pt(28)) > Pt(1):
                pf.first_line_indent = Pt(28)
                fix_count += 1
                para_fixed = True
            if text[0] in (' ', '\u3000'):
                para.runs[0].text = para.runs[0].text.lstrip(' \u3000')
                fix_count += 1
                para_fixed = True

        if check_punct:
            # 快速检查是否有英文标点
            if PUNCT_SET & set(text):
                cleaned = DIGIT_PUNCT_PATTERN.sub('', text)
                if any(ch in PUNCT_SET for ch in cleaned):
                    for run in para.runs:
                        _fix_punctuation_in_run(run)
                    fix_count += 1
                    para_fixed = True

        if kw_items:
            for kw, repl in kw_items:
                if kw in text:
                    for run in para.runs:
                        if kw in run.text:
                            run.text = run.text.replace(kw, repl)
                    fix_count += 1
                    warn_count += 1
                    para_fixed = True

        if check_font_or_color:
            fix_count = _process_runs_font_color(para.runs, False, categories, fix_count, details, details_full)

        if para_fixed and details_full:
            para_loc = text[:25] + "..." if len(text) > 25 else text
            details.append((para_loc, "格式修复"))
            if len(details) >= DETAILS_LIMIT:
                details_full = False

    # 表格内段落处理
    if check_font_or_color or check_align:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if not para.text.strip() or not para.runs:
                            continue

                        if check_align:
                            pf = para.paragraph_format
                            if pf.alignment is not None and pf.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                                pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                fix_count += 1

                        if check_font_or_color:
                            fix_count = _process_runs_font_color(para.runs, True, categories, fix_count, details, details_full)

    # 批注：大文档只在首段加一条汇总批注
    if add_comments and fix_count > 0 and first_para and first_para.runs:
        if fix_count > COMMENT_LIMIT:
            doc.add_comment(
                runs=first_para.runs,
                text=f"【暗标检查工具】共修复 {fix_count} 项格式问题，{warn_count} 项警告。详见修改明细。",
                author=AUTHOR, initials=INITIALS
            )
        else:
            doc.add_comment(
                runs=first_para.runs,
                text=f"【暗标检查工具】共修复 {fix_count} 项格式问题。",
                author=AUTHOR, initials=INITIALS
            )

    if fix_count == 0 and warn_count == 0:
        return details, 0, 0

    # 截断 details
    if not details_full:
        details.append(("...", f"共 {fix_count} 项修复，仅显示前 {DETAILS_LIMIT} 条"))

    if isinstance(output_path, str):
        buf = io.BytesIO()
        doc.save(buf)
        with open(output_path, 'wb') as f:
            f.write(buf.getvalue())
    else:
        doc.save(output_path)
    return details, fix_count, warn_count
