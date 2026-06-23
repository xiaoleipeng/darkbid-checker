"""
Word暗标编制要求检查工具 - Pyodide版核心逻辑
在浏览器中通过Pyodide运行，提供 check_fix_annotate() 函数。
"""
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
DIGIT_PUNCT_PATTERN = re.compile(r'\d[.,]\d')


def _is_in_table(paragraph):
    parent = paragraph._element.getparent()
    while parent is not None:
        if parent.tag.endswith('}tc') or parent.tag.endswith('}tbl'):
            return True
        parent = parent.getparent()
    return False


def _comment(doc, runs, text):
    doc.add_comment(runs=runs, text=text, author=AUTHOR, initials=INITIALS)


def _set_font_name(run):
    run.font.name = FONT_NAME
    rpr = run._element.find(qn('w:rPr'))
    if rpr is None:
        rpr = run._element.makeelement(qn('w:rPr'), {})
        run._element.insert(0, rpr)
    rFonts = rpr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), FONT_NAME)


def _fix_punctuation_in_run(run):
    text = run.text
    result = []
    for i, ch in enumerate(text):
        if ch in PUNCTUATION_MAP:
            before_digit = i > 0 and text[i-1].isdigit()
            after_digit = i < len(text)-1 and text[i+1].isdigit()
            if before_digit and after_digit and ch in '.,':
                result.append(ch)
            else:
                result.append(PUNCTUATION_MAP[ch])
        else:
            result.append(ch)
    run.text = ''.join(result)


def check_fix_annotate(doc_path, output_path, keyword_map=None, categories=None, add_comments=True):
    """核心修复逻辑，返回 (details, fix_count, warn_count)"""
    ALL_CATEGORIES = {'page', 'spacing', 'align', 'font', 'color', 'punct', 'identity'}
    if categories is None:
        categories = ALL_CATEGORIES

    doc = Document(doc_path)
    fix_count = 0
    warn_count = 0
    details = []

    first_para = doc.paragraphs[0] if doc.paragraphs else None

    # TOC目录检测
    if 'page' in categories:
        for para in doc.paragraphs:
            for instr in para._element.findall('.//' + qn('w:instrText')):
                if instr.text and 'TOC' in instr.text.upper():
                    details.append(("目录", "⚠️ 存在自动目录，需手动删除"))
                    warn_count += 1
                    if add_comments and para.runs:
                        _comment(doc, para.runs, "【⚠️ 需手动处理】存在自动目录，请手动删除。")
                    break

    # 页数估算
    total_paras = len(doc.paragraphs)
    estimated_pages = total_paras // 25
    if estimated_pages > 300:
        details.append(("全文", f"⚠️ 估算页数约{estimated_pages}页，可能超过300页限制"))
        warn_count += 1

    # 页面设置
    if 'page' in categories:
        page_fixes = []
        tolerance = Cm(0.1)
        margin_tolerance = Cm(0.05)
        for section in doc.sections:
            if section.page_width and abs(section.page_width - PAGE_WIDTH) > tolerance:
                page_fixes.append(f"纸张宽度: {section.page_width.cm:.1f}cm → 21.0cm")
                section.page_width = PAGE_WIDTH
            if section.page_height and abs(section.page_height - PAGE_HEIGHT) > tolerance:
                page_fixes.append(f"纸张高度: {section.page_height.cm:.1f}cm → 29.7cm")
                section.page_height = PAGE_HEIGHT
            if section.top_margin is not None and abs(section.top_margin - MARGIN_TOP) > margin_tolerance:
                page_fixes.append(f"上边距: {section.top_margin.cm:.2f}cm → 2.5cm")
                section.top_margin = MARGIN_TOP
            if section.bottom_margin is not None and abs(section.bottom_margin - MARGIN_BOTTOM) > margin_tolerance:
                page_fixes.append(f"下边距: {section.bottom_margin.cm:.2f}cm → 2.0cm")
                section.bottom_margin = MARGIN_BOTTOM
            if section.left_margin is not None and abs(section.left_margin - MARGIN_LEFT) > margin_tolerance:
                page_fixes.append(f"左边距: {section.left_margin.cm:.2f}cm → 2.0cm")
                section.left_margin = MARGIN_LEFT
            if section.right_margin is not None and abs(section.right_margin - MARGIN_RIGHT) > margin_tolerance:
                page_fixes.append(f"右边距: {section.right_margin.cm:.2f}cm → 2.0cm")
                section.right_margin = MARGIN_RIGHT
            if section.header and section.header.is_linked_to_previous is False:
                header_text = ''.join(p.text for p in section.header.paragraphs).strip()
                if header_text:
                    page_fixes.append("已清除页眉")
                    for p in section.header.paragraphs:
                        p.clear()
            if section.footer and section.footer.is_linked_to_previous is False:
                footer_text = ''.join(p.text for p in section.footer.paragraphs).strip()
                if footer_text:
                    page_fixes.append("已清除页脚/页码")
                    for p in section.footer.paragraphs:
                        p.clear()
        if page_fixes and first_para and first_para.runs:
            if add_comments:
                _comment(doc, first_para.runs, "【已修复 - 页面设置】\n" + "\n".join(f"• {f}" for f in page_fixes))
            fix_count += len(page_fixes)
            for f in page_fixes:
                details.append(("页面设置", f))

    # 逐段检查
    for para in doc.paragraphs:
        if not para.text.strip() or not para.runs:
            continue
        in_table = _is_in_table(para)
        para_loc = para.text[:25] + "..." if len(para.text) > 25 else para.text
        para_fixes = []

        if not in_table:
            if 'spacing' in categories:
                pf = para.paragraph_format
                if pf.line_spacing_rule != WD_LINE_SPACING.EXACTLY or pf.line_spacing != LINE_SPACING_VAL:
                    para_fixes.append("行间距 → 固定值30磅")
                    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                    pf.line_spacing = LINE_SPACING_VAL
                if pf.space_before and pf.space_before > Pt(0):
                    para_fixes.append("段前间距 → 0")
                    pf.space_before = Pt(0)
                if pf.space_after and pf.space_after > Pt(0):
                    para_fixes.append("段后间距 → 0")
                    pf.space_after = Pt(0)

            if 'align' in categories:
                pf = para.paragraph_format
                if pf.alignment is not None and pf.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                    para_fixes.append("对齐 → 左对齐")
                    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if pf.first_line_indent is None or abs(pf.first_line_indent - Pt(28)) > Pt(1):
                    para_fixes.append("首行缩进 → 2字符")
                    pf.first_line_indent = Pt(28)
                if para.text.startswith(' ') or para.text.startswith('\u3000'):
                    para_fixes.append("删除开头空格")
                    para.runs[0].text = para.runs[0].text.lstrip(' \u3000')

            if 'punct' in categories:
                cleaned = DIGIT_PUNCT_PATTERN.sub('', para.text)
                en_puncts = set(ch for ch in cleaned if ch in PUNCTUATION_MAP)
                if en_puncts:
                    para_fixes.append(f"标点: {' '.join(repr(c) for c in sorted(en_puncts))} → 中文")
                    for run in para.runs:
                        _fix_punctuation_in_run(run)

        if 'identity' in categories and keyword_map:
            for kw, repl in keyword_map.items():
                if kw in para.text:
                    para_fixes.append(f"身份信息 '{kw}' → '{repl}'")
                    for run in para.runs:
                        if kw in run.text:
                            run.text = run.text.replace(kw, repl)
                    warn_count += 1

        if para_fixes:
            if add_comments:
                _comment(doc, para.runs, "【已修复】\n" + "\n".join(f"• {f}" for f in para_fixes))
            fix_count += len(para_fixes)
            for f in para_fixes:
                details.append((para_loc, f))

        # 字体检查
        if 'font' in categories or 'color' in categories:
            for run in para.runs:
                if not run.text.strip():
                    continue
                font = run.font
                run_fixes = []
                run_loc = run.text[:15] + "..." if len(run.text) > 15 else run.text

                if 'font' in categories:
                    if not in_table:
                        if font.name and font.name != FONT_NAME:
                            run_fixes.append(f"字体: {font.name} → 宋体")
                            _set_font_name(run)
                        else:
                            rpr = run._element.find(qn('w:rPr'))
                            if rpr is not None:
                                rFonts = rpr.find(qn('w:rFonts'))
                                if rFonts is not None:
                                    ea = rFonts.get(qn('w:eastAsia'))
                                    if ea and ea != FONT_NAME:
                                        run_fixes.append(f"中文字体: {ea} → 宋体")
                                        _set_font_name(run)
                        if font.size and font.size != FONT_SIZE:
                            run_fixes.append(f"字号: {font.size.pt:.1f}pt → 14pt")
                            font.size = FONT_SIZE
                    if font.bold:
                        run_fixes.append("去除加粗")
                        font.bold = False
                    if font.italic:
                        run_fixes.append("去除倾斜")
                        font.italic = False
                    if font.underline:
                        run_fixes.append("去除下划线")
                        font.underline = False

                if 'color' in categories:
                    if font.color and font.color.rgb and font.color.rgb != BLACK_COLOR:
                        run_fixes.append(f"颜色: #{font.color.rgb} → 黑色")
                        font.color.rgb = BLACK_COLOR

                if run_fixes:
                    if add_comments:
                        prefix = "【已修复 - 表格内字体】" if in_table else "【已修复 - 字体】"
                        _comment(doc, [run], prefix + "\n" + "\n".join(f"• {f}" for f in run_fixes))
                    fix_count += len(run_fixes)
                    loc_prefix = "[表格内] " if in_table else ""
                    for f in run_fixes:
                        details.append((loc_prefix + run_loc, f))

    # 表格内段落
    if 'font' in categories or 'color' in categories or 'align' in categories:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if not para.text.strip() or not para.runs:
                            continue
                        para_loc = "[表格内] " + (para.text[:15] + "..." if len(para.text) > 15 else para.text)

                        if 'align' in categories:
                            pf = para.paragraph_format
                            if pf.alignment is not None and pf.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                                align_map = {WD_ALIGN_PARAGRAPH.CENTER: "居中", WD_ALIGN_PARAGRAPH.RIGHT: "右对齐", WD_ALIGN_PARAGRAPH.JUSTIFY: "两端对齐"}
                                fix_desc = f"对齐: {align_map.get(pf.alignment, '?')} → 左对齐"
                                pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                if add_comments:
                                    _comment(doc, para.runs, "【已修复 - 表格内】\n• " + fix_desc)
                                fix_count += 1
                                details.append((para_loc, fix_desc))

                        if 'font' in categories or 'color' in categories:
                            for run in para.runs:
                                if not run.text.strip():
                                    continue
                                font = run.font
                                run_fixes = []
                                run_loc = "[表格内] " + (run.text[:15] + "..." if len(run.text) > 15 else run.text)

                                if 'font' in categories:
                                    if font.bold:
                                        run_fixes.append("去除加粗")
                                        font.bold = False
                                    if font.italic:
                                        run_fixes.append("去除倾斜")
                                        font.italic = False
                                    if font.underline:
                                        run_fixes.append("去除下划线")
                                        font.underline = False

                                if 'color' in categories:
                                    if font.color and font.color.rgb and font.color.rgb != BLACK_COLOR:
                                        run_fixes.append(f"颜色: #{font.color.rgb} → 黑色")
                                        font.color.rgb = BLACK_COLOR

                                if run_fixes:
                                    if add_comments:
                                        _comment(doc, [run], "【已修复 - 表格内字体】\n" + "\n".join(f"• {f}" for f in run_fixes))
                                    fix_count += len(run_fixes)
                                    for f in run_fixes:
                                        details.append((run_loc, f))

    if fix_count == 0 and warn_count == 0:
        return details, 0, 0

    doc.save(output_path)
    return details, fix_count, warn_count
