#!/usr/bin/env python3
"""
PDF文档暗标编制要求检查工具
检测PDF是否符合暗标格式规范，输出检查报告。
可选 --fix 模式：尝试修复部分问题（颜色、标点、身份信息）。

⚠️ PDF修复风险提示：
  - PDF不是结构化文档，修复可能导致排版错乱、文字重叠或丢失
  - 仅能修复：颜色→黑色、英文标点→中文、身份关键词替换
  - 无法修复：字体、字号、行间距、页边距、对齐方式
  - 强烈建议在Word源文件中修改后重新导出PDF

依赖：PyMuPDF (fitz)
"""

import sys
import re
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("错误: 需要安装 PyMuPDF，请执行: pip install pymupdf")
    sys.exit(1)

# 常量
A4_WIDTH_PT = 595.276
A4_HEIGHT_PT = 841.89
MARGIN_TOP_PT = 70.87
MARGIN_BOTTOM_PT = 56.69
MARGIN_LEFT_PT = 56.69
MARGIN_RIGHT_PT = 56.69
TOLERANCE_PT = 3.0

EXPECTED_FONT_ALIASES = {'宋体', 'SimSun', 'SimSun-ExtB', 'STSong', 'Songti SC', 'FangSong'}
EXPECTED_SIZE = 14.0
SIZE_TOLERANCE = 0.5
LINE_SPACING_PT = 30.0
LINE_SPACING_TOLERANCE = 2.0

PUNCTUATION_MAP = {
    ',': '，', '.': '。', ':': '：', ';': '；',
    '!': '！', '?': '？', '(': '（', ')': '）',
    '[': '【', ']': '】',
}
PUNCTUATION_EN = set(PUNCTUATION_MAP.keys())
DIGIT_PUNCT_PATTERN = re.compile(r'\d[.,]\d')

ALL_CATEGORIES = {'page', 'spacing', 'align', 'font', 'color', 'punct', 'identity'}
# --fix 模式能修复的类别
FIXABLE_CATEGORIES = {'color', 'punct', 'identity'}


class Issue:
    def __init__(self, category, page, message):
        self.category = category
        self.page = page
        self.message = message

    def __str__(self):
        return f"[{self.category}] 第{self.page}页: {self.message}"


def check_pdf(pdf_path, keywords=None, categories=None):
    """检查 PDF，返回问题列表"""
    if categories is None:
        categories = ALL_CATEGORIES

    doc = fitz.open(pdf_path)
    issues = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        pn = page_num + 1

        if 'page' in categories:
            rect = page.rect
            w, h = rect.width, rect.height
            if abs(w - A4_WIDTH_PT) > TOLERANCE_PT or abs(h - A4_HEIGHT_PT) > TOLERANCE_PT:
                issues.append(Issue('page', pn,
                    f"纸张尺寸 {w/72*25.4:.1f}×{h/72*25.4:.1f}mm，不是A4(210×297mm)"))

            blocks = page.get_text("blocks")
            if blocks:
                content_left = min(b[0] for b in blocks)
                content_top = min(b[1] for b in blocks)
                content_right = max(b[2] for b in blocks)
                content_bottom = max(b[3] for b in blocks)
                if content_left < MARGIN_LEFT_PT - TOLERANCE_PT:
                    issues.append(Issue('page', pn, f"左边距约{content_left/72*25.4:.1f}mm，小于要求(20mm)"))
                if content_top < MARGIN_TOP_PT - TOLERANCE_PT:
                    issues.append(Issue('page', pn, f"上边距约{content_top/72*25.4:.1f}mm，小于要求(25mm)"))
                if rect.width - content_right < MARGIN_RIGHT_PT - TOLERANCE_PT:
                    issues.append(Issue('page', pn, f"右边距约{(rect.width-content_right)/72*25.4:.1f}mm，小于要求(20mm)"))
                if rect.height - content_bottom < MARGIN_BOTTOM_PT - TOLERANCE_PT:
                    issues.append(Issue('page', pn, f"下边距约{(rect.height-content_bottom)/72*25.4:.1f}mm，小于要求(20mm)"))

        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue
            lines = block.get("lines", [])

            if 'spacing' in categories and len(lines) >= 2:
                for i in range(1, len(lines)):
                    y0 = lines[i-1]["bbox"][1]
                    y1 = lines[i]["bbox"][1]
                    gap = y1 - y0
                    if gap > 0 and (gap > LINE_SPACING_PT + 5 or gap < LINE_SPACING_PT - 5):
                        issues.append(Issue('spacing', pn, f"行间距约{gap:.1f}pt，要求固定值30磅"))
                        break

            for line in lines:
                for span in line.get("spans", []):
                    text = span.get("text", "").strip()
                    if not text:
                        continue
                    font_name = span.get("font", "")
                    font_size = span.get("size", 0)
                    color_int = span.get("color", 0)
                    flags = span.get("flags", 0)

                    if 'font' in categories:
                        is_expected = any(alias in font_name for alias in EXPECTED_FONT_ALIASES)
                        if not is_expected and font_name:
                            issues.append(Issue('font', pn, f"字体「{font_name}」非宋体，内容: '{text[:15]}'"))
                        if font_size > 0 and abs(font_size - EXPECTED_SIZE) > SIZE_TOLERANCE:
                            issues.append(Issue('font', pn, f"字号{font_size:.1f}pt，要求14pt(四号)，内容: '{text[:15]}'"))
                        if flags & 16:
                            issues.append(Issue('font', pn, f"检测到加粗，内容: '{text[:15]}'"))
                        if flags & 2:
                            issues.append(Issue('font', pn, f"检测到倾斜，内容: '{text[:15]}'"))

                    if 'color' in categories:
                        if color_int != 0 and color_int != 0x000000:
                            r = (color_int >> 16) & 0xFF
                            g = (color_int >> 8) & 0xFF
                            b = color_int & 0xFF
                            issues.append(Issue('color', pn, f"颜色#{r:02X}{g:02X}{b:02X}非黑色，内容: '{text[:15]}'"))

                    if 'punct' in categories:
                        cleaned = DIGIT_PUNCT_PATTERN.sub('', text)
                        en_puncts = set(ch for ch in cleaned if ch in PUNCTUATION_EN)
                        if en_puncts:
                            issues.append(Issue('punct', pn, f"英文标点 {' '.join(repr(c) for c in sorted(en_puncts))}，内容: '{text[:20]}'"))

                    if 'identity' in categories and keywords:
                        for kw in keywords:
                            if kw in text:
                                issues.append(Issue('identity', pn, f"发现身份信息「{kw}」，内容: '{text[:30]}'"))

        if 'align' in categories:
            blocks = page.get_text("blocks")
            for block in blocks:
                block_text = block[4] if len(block) > 4 else ""
                if isinstance(block_text, str) and block_text.strip():
                    if not (block_text.startswith('  ') or block_text.startswith('\u3000\u3000')):
                        if block_text.startswith(' ') or block_text.startswith('\u3000'):
                            issues.append(Issue('align', pn, f"段落开头有空格，内容: '{block_text.strip()[:20]}'"))

    doc.close()

    seen = set()
    deduped = []
    for issue in issues:
        key = (issue.category, issue.page, issue.message[:50])
        if key not in seen:
            seen.add(key)
            deduped.append(issue)
    return deduped


def fix_pdf(pdf_path, output_path, keywords=None, categories=None):
    """
    尝试修复 PDF 中可修复的问题。
    可修复项：color（颜色→黑色）、punct（英文标点→中文）、identity（关键词替换）
    返回：(fix_count, unfixable_issues)
    """
    if categories is None:
        categories = FIXABLE_CATEGORIES
    else:
        categories = categories & FIXABLE_CATEGORIES

    doc = fitz.open(pdf_path)
    fix_count = 0

    for page_num in range(len(doc)):
        page = doc[page_num]

        # 颜色修复：将所有非黑色文字改为黑色
        if 'color' in categories:
            text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
            for block in text_dict.get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        color_int = span.get("color", 0)
                        if color_int != 0 and color_int != 0x000000:
                            # 用黑色重新绘制该文字
                            text = span.get("text", "")
                            if not text.strip():
                                continue
                            bbox = fitz.Rect(span["bbox"])
                            font_size = span.get("size", 12)
                            # 先用白色矩形覆盖原文字
                            page.draw_rect(bbox, color=None, fill=(1, 1, 1))
                            # 重新插入黑色文字
                            page.insert_text(
                                fitz.Point(bbox.x0, bbox.y1 - 2),
                                text, fontsize=font_size, color=(0, 0, 0)
                            )
                            fix_count += 1

        # 标点修复
        if 'punct' in categories:
            for en, cn in PUNCTUATION_MAP.items():
                instances = page.search_for(en)
                for inst in instances:
                    # 获取上下文判断是否在数字间
                    surrounding = page.get_text("text", clip=fitz.Rect(
                        inst.x0 - 10, inst.y0, inst.x1 + 10, inst.y1))
                    if DIGIT_PUNCT_PATTERN.search(surrounding):
                        continue
                    page.add_redact_annot(inst, cn, fontsize=0)
                    fix_count += 1
            page.apply_redactions()

        # 身份信息替换
        if 'identity' in categories and keywords:
            for kw in keywords:
                instances = page.search_for(kw)
                for inst in instances:
                    page.add_redact_annot(inst, "***", fontsize=0)
                    fix_count += 1
            page.apply_redactions()

    if fix_count > 0:
        doc.save(output_path)
    doc.close()
    return fix_count


def main():
    if len(sys.argv) < 2:
        print("用法: python3 check_pdf.py <PDF文件路径> [选项]")
        print()
        print("选项:")
        print("  --only <类别>    仅检查指定类别（逗号分隔）")
        print("  --keywords <词>  身份信息关键词（逗号分隔）")
        print("  --fix            尝试修复可修复的问题（颜色/标点/身份信息）")
        print()
        print("支持的类别:")
        print("  page      纸张A4、页边距")
        print("  spacing   行间距（固定值30磅）")
        print("  align     首行缩进、开头空格")
        print("  font      字体（宋体四号）、加粗/倾斜")
        print("  color     字体颜色（黑色）")
        print("  punct     英文标点")
        print("  identity  身份信息关键词")
        print()
        print("示例:")
        print("  python3 check_pdf.py 文件.pdf")
        print("  python3 check_pdf.py 文件.pdf --fix")
        print("  python3 check_pdf.py 文件.pdf --fix --keywords \"XX公司,张三\"")
        print("  python3 check_pdf.py 文件.pdf --only page,font")
        sys.exit(0)

    pdf_path = sys.argv[1]
    keywords = None
    categories = None
    do_fix = '--fix' in sys.argv

    if '--keywords' in sys.argv:
        idx = sys.argv.index('--keywords')
        if idx + 1 < len(sys.argv):
            keywords = [k.strip() for k in sys.argv[idx+1].split(',') if k.strip()]

    if '--only' in sys.argv:
        idx = sys.argv.index('--only')
        if idx + 1 < len(sys.argv):
            categories = set(k.strip() for k in sys.argv[idx+1].split(',') if k.strip())
            invalid = categories - ALL_CATEGORIES
            if invalid:
                print(f"错误: 无效类别 {invalid}，支持: {', '.join(sorted(ALL_CATEGORIES))}")
                sys.exit(1)

    if not Path(pdf_path).exists():
        print(f"错误: 文件不存在 - {pdf_path}")
        sys.exit(1)

    print(f"正在检查: {pdf_path}")
    if categories:
        print(f"检查类别: {', '.join(sorted(categories))}")
    if keywords:
        print(f"身份关键词: {', '.join(keywords)}")
    print()

    issues = check_pdf(pdf_path, keywords, categories)

    if not issues:
        print("✓ 所有检查通过，PDF文档符合暗标编制要求！")
        return

    # 按类别分组输出
    by_cat = {}
    for issue in issues:
        by_cat.setdefault(issue.category, []).append(issue)

    cat_names = {'page': '页面设置', 'spacing': '行间距', 'align': '对齐缩进',
                 'font': '字体字号', 'color': '颜色', 'punct': '标点符号', 'identity': '身份信息'}

    print(f"共发现 {len(issues)} 个问题：")
    print("=" * 60)
    for cat in ['page', 'spacing', 'align', 'font', 'color', 'punct', 'identity']:
        if cat in by_cat:
            cat_issues = by_cat[cat]
            fixable = "✓可修复" if cat in FIXABLE_CATEGORIES else "✗不可修复"
            print(f"\n【{cat_names[cat]}】({len(cat_issues)}项) [{fixable}]")
            for issue in cat_issues[:20]:
                print(f"  ⚠️  {issue}")
            if len(cat_issues) > 20:
                print(f"  ... 还有 {len(cat_issues)-20} 项同类问题")

    print("\n" + "=" * 60)

    if not do_fix:
        print("⚠️  注意: PDF文件修复有风险，建议在Word源文件中修改后重新导出PDF。")
        print("    如需尝试修复，请添加 --fix 参数（仅能修复颜色/标点/身份信息）。")
        return

    # 执行修复
    fixable_issues = [i for i in issues if i.category in FIXABLE_CATEGORIES]
    unfixable_issues = [i for i in issues if i.category not in FIXABLE_CATEGORIES]

    if not fixable_issues:
        print("\n所有问题均不可通过PDF修复（字体/字号/行间距/页边距/对齐），请在Word中修改。")
        return

    print(f"\n{'='*60}")
    print("⚠️  PDF修复风险提示：")
    print("  • 修复可能导致部分文字位置偏移或排版变化")
    print("  • 修复后的字体可能与原文不同（回退为默认字体）")
    print("  • 原始文件不会被修改，修复结果另存为新文件")
    print(f"  • 可修复 {len(fixable_issues)} 项（颜色/标点/身份信息）")
    if unfixable_issues:
        print(f"  • 无法修复 {len(unfixable_issues)} 项（字体/字号/行间距/页边距/对齐）")
    print(f"{'='*60}")

    confirm = input("\n确认修复？(y/N): ").strip().lower()
    if confirm != 'y':
        print("已取消修复。")
        return

    p = Path(pdf_path)
    output_path = p.with_stem(p.stem + '_已修复').as_posix()

    fix_cats = (categories & FIXABLE_CATEGORIES) if categories else None
    fix_count = fix_pdf(pdf_path, output_path, keywords, fix_cats)

    if fix_count > 0:
        print(f"\n✓ 已修复 {fix_count} 处，输出文件: {output_path}")
        if unfixable_issues:
            print(f"⚠️  仍有 {len(unfixable_issues)} 项无法修复（字体/字号/行间距等），请在Word中处理。")
    else:
        print("\n未找到可修复的内容。")


if __name__ == "__main__":
    main()
