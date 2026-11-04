#!/usr/bin/env python3
"""交付前检查:validate.py、markitdown、LibreOffice 渲染都看不到的几类问题。

用法:
    python3 scripts/check_delivery.py out.pptx --template templates/company-template.pptx
    python3 scripts/check_delivery.py out.pptx --scaffold templates/campus-hire-defence.pptx
    python3 scripts/check_delivery.py out.pptx --template ... --fix   # 自动修 docProps 并重新检查

检查项:
  1. XML 声明:每个 .xml/.rels 部件开头必须带 standalone="yes",缺了真实 PowerPoint 会清空该页
  2. docProps:作者字段不能是模板遗留值或邮箱;app.xml 的页数和 TitlesOfParts 不能是旧值
  3. 示例文案残留:--template 指定的源文件里所有幻灯片都是示例页,成稿里出现同一段原文即失败
  4. 场景模板占位:--scaffold 指定的场景模板里,同一页重复出现的文字是槽位占位,成稿里仍在即失败;
     其他与场景模板相同的文字列为"需确认"(多数是有意保留的章节标题)
  5. 通用占位符与规划标签:lorem / TODO / xxx / 待补充 / [外部资料] / [阻塞待确认] 等

只依赖 Python 标准库。退出码:0 = 全部通过,1 = 有失败项。
"""
import argparse
import re
import sys
import zipfile
from collections import Counter

MIN_LEN = 6  # 短于这个长度的段落(页码、"01"、"本科")不参与原文比对
PLACEHOLDER_RE = re.compile(
    r"lorem|ipsum|\bTODO\b|\bx{3,}\b|\bX{3,}\b|\[insert|待补充|待填写|此处输入|请输入|单击此处",
    re.IGNORECASE,
)
PLANNING_TAG_RE = re.compile(r"\[(外部资料|阻塞待确认)\]")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")
DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def slide_order(z):
    """按 <p:sldIdLst> 的顺序返回幻灯片部件路径,只含成稿里真正出现的页。"""
    pres = z.read("ppt/presentation.xml").decode("utf-8")
    rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    targets = {}
    for rel in re.findall(r"<Relationship\b[^>]*>", rels):
        rid = re.search(r'\bId="([^"]+)"', rel)
        target = re.search(r'\bTarget="([^"]+)"', rel)
        if rid and target:
            targets[rid.group(1)] = target.group(1)
    lst = re.search(r"<p:sldIdLst>(.*?)</p:sldIdLst>", pres, re.S)
    ids = re.findall(r'r:id="([^"]+)"', lst.group(1)) if lst else []
    return ["ppt/" + targets[i].lstrip("/").removeprefix("ppt/") for i in ids if i in targets]


def paragraphs(xml):
    out = []
    for body in re.findall(r"<a:p\b[^>]*>(.*?)</a:p>", xml, re.S):
        text = "".join(re.findall(r"<a:t\b[^>]*>([^<]*)</a:t>", body)).strip()
        if text:
            out.append(text)
    return out


def deck_paragraphs(path):
    """返回 [(页号, 段落文字)],页号从 1 开始,按成稿顺序。"""
    with zipfile.ZipFile(path) as z:
        return [(i, t) for i, part in enumerate(slide_order(z), 1)
                for t in paragraphs(z.read(part).decode("utf-8"))]


def worth_comparing(text):
    return len(text) >= MIN_LEN and not text.replace(" ", "").isdigit()


def check_declarations(z):
    bad = [n for n in z.namelist()
           if n.endswith((".xml", ".rels")) and b'standalone="yes"' not in z.read(n)[:200]]
    return [f"{n} 的 XML 声明缺少 standalone=\"yes\"——真实 PowerPoint 会把该部件当损坏文件清空。"
            f"不要用 minidom 的 toxml(encoding=...) 写回,改为手动拼接声明(见 SKILL.md「填充模板时」)"
            for n in bad]


def author_fields(xml):
    return dict(re.findall(r"<(dc:creator|cp:lastModifiedBy)>([^<]*)<", xml))


def slide_titles(app):
    """TitlesOfParts 依次是字体、主题、幻灯片标题等分组,组名和条数在 HeadingPairs 里;只取幻灯片标题组。"""
    pairs = re.search(r"<HeadingPairs>(.*?)</HeadingPairs>", app, re.S)
    parts = re.search(r"<TitlesOfParts>(.*?)</TitlesOfParts>", app, re.S)
    if not pairs or not parts:
        return []
    groups = re.findall(r"<vt:lpstr>([^<]*)</vt:lpstr>\s*</vt:variant>\s*<vt:variant>\s*<vt:i4>(\d+)</vt:i4>",
                        pairs.group(1))
    names = re.findall(r"<vt:lpstr>([^<]*)</vt:lpstr>", parts.group(1))
    start = 0
    for name, count in groups:
        count = int(count)
        if "幻灯片" in name or "slide" in name.lower():
            return names[start:start + count]
        start += count
    return []


def check_docprops(z, template_path, slide_count, deck_texts):
    fails = []
    names = z.namelist()
    inherited = {}
    if template_path:
        with zipfile.ZipFile(template_path) as t:
            if "docProps/core.xml" in t.namelist():
                inherited = author_fields(t.read("docProps/core.xml").decode("utf-8"))
    if "docProps/core.xml" in names:
        for field, value in author_fields(z.read("docProps/core.xml").decode("utf-8")).items():
            if not value:
                continue
            if EMAIL_RE.search(value):
                fails.append(f"docProps/core.xml 的 <{field}> 是邮箱「{value}」,分发出去会带走内部账号")
            elif inherited.get(field) == value:
                fails.append(f"docProps/core.xml 的 <{field}>「{value}」是从模板继承来的原作者")
    if "docProps/app.xml" in names:
        app = z.read("docProps/app.xml").decode("utf-8")
        m = re.search(r"<Slides>(\d+)</Slides>", app)
        if m and int(m.group(1)) != slide_count:
            fails.append(f"docProps/app.xml 的 <Slides> 是 {m.group(1)},成稿实际 {slide_count} 页")
        titles = slide_titles(app)
        if titles:
            joined = "\n".join(deck_texts)
            stale = [s for s in titles if worth_comparing(s) and s not in joined]
            if stale:
                fails.append("docProps/app.xml 的 <TitlesOfParts> 留着成稿里没有的旧标题:" + "、".join(stale[:5])
                             + (" 等" if len(stale) > 5 else ""))
    return fails


def check_template_residue(deck, template_path):
    example = {t for _, t in deck_paragraphs(template_path) if worth_comparing(t)}
    return [f"第 {page} 页残留模板示例文案:「{text}」——改成用户内容或删除该形状"
            for page, text in deck if text in example]


def check_scaffold(deck, scaffold_path):
    src = deck_paragraphs(scaffold_path)
    per_slide = Counter(src)
    slot = {t for (_, t), n in per_slide.items() if n >= 2 and worth_comparing(t)}
    seen = {t for _, t in src if worth_comparing(t)}
    fails = [f"第 {page} 页仍是场景模板的槽位占位文字:「{text}」" for page, text in deck if text in slot]
    review = sorted({f"第 {page} 页「{text}」" for page, text in deck if text in seen and text not in slot})
    return fails, review


def check_generic(deck):
    fails = []
    for page, text in deck:
        if PLACEHOLDER_RE.search(text):
            fails.append(f"第 {page} 页有占位符文字:「{text}」")
        m = PLANNING_TAG_RE.search(text)
        if m:
            fails.append(f"第 {page} 页出现规划标签 {m.group(0)}:页面上改用「来源:…」,或先把阻塞项向用户确认")
    return fails


def fix_docprops(path, template_path):
    """清掉继承来的作者、同步页数、删掉 TitlesOfParts/HeadingPairs。纯字符串替换,保留原 XML 声明。"""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
        slide_count = len(slide_order(z))
    inherited = {}
    if template_path:
        with zipfile.ZipFile(template_path) as t:
            if "docProps/core.xml" in t.namelist():
                inherited = author_fields(t.read("docProps/core.xml").decode("utf-8"))
    out = []
    for info, data in items:
        if info.filename == "docProps/core.xml":
            s = data.decode("utf-8")
            for field, value in author_fields(s).items():
                if value and (EMAIL_RE.search(value) or inherited.get(field) == value):
                    s = s.replace(f"<{field}>{value}<", f"<{field}><")
            data = s.encode("utf-8")
        elif info.filename == "docProps/app.xml":
            s = data.decode("utf-8")
            s = re.sub(r"<Slides>\d+</Slides>", f"<Slides>{slide_count}</Slides>", s)
            s = re.sub(r"<HeadingPairs>.*?</HeadingPairs>", "", s, flags=re.S)
            s = re.sub(r"<TitlesOfParts>.*?</TitlesOfParts>", "", s, flags=re.S)
            data = s.encode("utf-8")
        out.append((info, data))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)


def run(args):
    deck = deck_paragraphs(args.deck)
    with zipfile.ZipFile(args.deck) as z:
        slide_count = len(slide_order(z))
        fails = check_declarations(z)
        fails += check_docprops(z, args.template, slide_count, [t for _, t in deck])
    review = []
    if args.template:
        fails += check_template_residue(deck, args.template)
    if args.scaffold:
        f, review = check_scaffold(deck, args.scaffold)
        fails += f
    fails += check_generic(deck)
    return fails, review


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--template", help="成稿派生自的公司模板;其中所有幻灯片都视为示例页")
    ap.add_argument("--scaffold", help="成稿派生自的场景模板(templates/ 下的现成模板)")
    ap.add_argument("--fix", action="store_true", help="自动修复 docProps 问题后重新检查(原地改写 deck)")
    args = ap.parse_args()

    if args.fix:
        fix_docprops(args.deck, args.template or args.scaffold)
    fails, review = run(args)

    for line in review:
        print(f"需确认  {line}——与场景模板原文相同,确认是有意保留的标题而不是没改的提示语")
    for line in fails:
        print(f"失败    {line}")
    if fails:
        print(f"\n{len(fails)} 项失败。docProps 类问题可加 --fix 自动修;其余回到对应页修改,重新打包后再跑本脚本。")
        sys.exit(1)
    print("交付前检查通过" + (f"({len(review)} 项需人工确认)" if review else ""))


if __name__ == "__main__":
    main()
