#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 终端AI指南.md 转成:
  1) 终端AI指南-公众号.html  —— 内联样式,浏览器打开后全选复制,粘贴进公众号编辑器即可(图片会自动上传)
  2) 终端AI指南-公众号.docx  —— 配色一致的 Word 版,图片已内嵌
配色:爱马仕橙 #F37021 + 蒂芙尼蓝 #0ABAB5
"""
import base64, html, mimetypes, re, subprocess, sys, tempfile, zipfile
from urllib.parse import unquote
from pathlib import Path
from lxml import etree

ORANGE, ORANGE_D, ORANGE_L = "#F37021", "#D95F12", "#FFF6F0"
TIFFANY, TIFFANY_D, TIFFANY_L = "#0ABAB5", "#079C98", "#EAF9F8"
INK, BODY, MUTED = "#2B2B2B", "#3F3F3F", "#8A8A8A"
CODE_BG, CODE_HDR, CODE_FG = "#22252B", "#1B1E23", "#E6E8EB"
MONO = "Menlo,Consolas,'Courier New',monospace"
SANS = "-apple-system,BlinkMacSystemFont,'PingFang SC','Hiragino Sans GB','Microsoft YaHei',sans-serif"

# 用法: python3 build_wechat.py [文章.md]   默认处理 终端AI指南.md
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("终端AI指南.md")
OUT_HTML = SRC.with_name(SRC.stem + "-公众号.html")
OUT_HTML_SINGLE = SRC.with_name(SRC.stem + "-公众号-单文件.html")
OUT_DOCX = SRC.with_name(SRC.stem + "-公众号.docx")

# 针对个别文章的笔误修正(换文章时留空即可,匹配不到就是空操作)
PER_ARTICLE_FIXES = (
    ("claude_cod\n", "claude_code\n"),
    ("因为我已经再环境变量中", "因为我已经在环境变量中"),
)

# ---------------------------------------------------------------- 解析 markdown
def parse(md_text):
    lines = md_text.split("\n")
    blocks, i = [], 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("```"):
            lang = line[3:].strip()
            i += 1
            buf = []
            while i < len(lines) and not lines[i].startswith("```"):
                buf.append(lines[i]); i += 1
            i += 1
            blocks.append(("code", lang, "\n".join(buf)))
            continue
        m = re.match(r"^(#{1,6})\s+(.*?)\s*$", line)
        if m:
            text = re.sub(r"\*\*(.+?)\*\*", r"\1", m.group(2))
            if text.strip() != "[TOC]":
                blocks.append(("head", len(m.group(1)), text))
            i += 1
            continue
        m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)\s*$", line.strip())
        if m:
            blocks.append(("img", m.group(1), m.group(2))); i += 1; continue
        if not line.strip():
            i += 1; continue
        buf = [line.strip()]; i += 1
        while i < len(lines) and lines[i].strip() and not re.match(r"^(#{1,6}\s|!\[|```)", lines[i]):
            buf.append(lines[i].strip()); i += 1
        blocks.append(("para", " ".join(buf)))
    return blocks

# ---------------------------------------------------------------- 行内渲染
def normalize_punct(s):
    """中文排版:把中文句子里的半角 , . : 换成全角(URL 和文件名先挖出来保护)。"""
    if not re.search(r"[\u4e00-\u9fff]", s):
        return s
    holes = []
    def stash(m):
        holes.append(m.group(0))
        return "\x00%d\x00" % (len(holes) - 1)
    s = re.sub(r"https?://[^\s]+", stash, s)
    s = re.sub(r"\b[\w./-]+\.(png|jpe?g|gif|sh|md|json|jsonc|toml|js|py|txt)\b", stash, s)
    s = re.sub(r"([\u4e00-\u9fff]),", r"\1，", s)
    s = re.sub(r",(?=[A-Za-z0-9])", "，", s)
    s = re.sub(r",(?=[\u4e00-\u9fff])", "，", s)
    s = re.sub(r"([\u4e00-\u9fff])\.(?=\s|[\u4e00-\u9fff]|$)", r"\1。", s)
    s = re.sub(r"([\u4e00-\u9fff]):(?=\s*\S)", r"\1：", s)
    return re.sub(r"\x00(\d+)\x00", lambda m: holes[int(m.group(1))], s)

def inline(text):
    t = html.escape(text)
    t = re.sub(r"\*\*(.+?)\*\*", r'<strong style="color:%s;">\1</strong>' % INK, t)
    t = re.sub(r"(https?://[^\s<)&]+)",
               r'<span style="color:%s;word-break:break-all;">\1</span>' % TIFFANY_D, t)
    return t

def highlight_code(code):
    out = []
    for line in code.split("\n"):
        s = html.escape(line)
        if s.strip().startswith("#"):
            out.append('<span style="color:#7F8896;">%s</span>' % s)
            continue
        s = re.sub(r"(https?://[^\s<)&]+)", r'<span style="color:#4FD3CE;">\1</span>', s)
        s = re.sub(r"(&lt;[^&]*?&gt;)", r'<span style="color:#FFC46B;">\1</span>', s)
        s = re.sub(r"^(\s*)(npm|bash|export|brew|curl|sudo)(?=\s)",
                   r'\1<span style="color:#FF9A5C;font-weight:600;">\2</span>', s)
        out.append(s)
    return "\n".join(out)

# ---------------------------------------------------------------- 块渲染
def r_head(lvl, text):
    text = normalize_punct(text)
    m = re.match(r"^(\d+)\s*[.,、)]\s*(.+)$", text)
    # H5(或更深)且以数字开头 —— 步骤
    if m and lvl >= 5:
        return ('<p style="margin:24px 0 12px;font-size:15px;font-weight:700;color:%s;line-height:1.8;">'
                '<span style="background-color:%s;color:#FFFFFF;border-radius:9px;padding:2px 9px;'
                'font-size:13px;margin-right:9px;">%s</span>%s</p>'
                % (INK, ORANGE, m.group(1), inline(m.group(2))))
    # H2 —— 章节大标题:爱马仕橙,最大最重,带浅橙底
    if lvl == 2:
        return ('<p style="margin:44px 0 20px;padding:11px 14px;border-left:5px solid %s;'
                'border-radius:0 8px 8px 0;background-color:%s;'
                'background-image:linear-gradient(90deg,#FFEDE0 0%%,#FFFFFF 96%%);'
                'font-size:20px;font-weight:800;color:%s;letter-spacing:1.2px;line-height:1.5;">%s</p>'
                % (ORANGE, ORANGE_L, INK, inline(text)))
    # H3 —— 小节标题:蒂芙尼蓝,中等,带浅蓝底
    if lvl == 3:
        return ('<p style="margin:30px 0 14px;padding:9px 13px;border-left:4px solid %s;'
                'border-radius:0 8px 8px 0;background-color:%s;'
                'background-image:linear-gradient(90deg,#E1F6F5 0%%,#FFFFFF 96%%);'
                'font-size:17px;font-weight:700;color:%s;letter-spacing:.9px;line-height:1.55;">%s</p>'
                % (TIFFANY, TIFFANY_L, TIFFANY_D, inline(text)))
    # H4 —— 分组标题:墨色小字 + 蓝色菱形前缀,无底色
    if lvl == 4:
        return ('<p style="margin:26px 0 12px;font-size:15.5px;font-weight:700;color:%s;'
                'letter-spacing:.5px;line-height:1.7;">'
                '<span style="color:%s;font-size:12px;margin-right:8px;">&#9670;</span>%s</p>'
                % (INK, TIFFANY, inline(text)))
    return ('<p style="margin:22px 0 10px;font-size:15px;font-weight:700;color:%s;'
            'letter-spacing:.5px;line-height:1.7;">%s</p>' % (INK, inline(text)))

def r_para(text):
    text = normalize_punct(text)
    if "不需要翻墙" in text:
        return ('<p style="margin:0 0 18px;text-align:center;">'
                '<span style="display:inline-block;background-color:%s;color:#FFFFFF;font-size:13px;'
                'letter-spacing:1px;border-radius:20px;padding:6px 18px;">%s</span></p>'
                % (ORANGE, inline(text)))
    if "platform.deepseek.com/api_keys" in text:
        return ('<section style="margin:0 0 18px;background-color:%s;border-left:4px solid %s;'
                'border-radius:0 8px 8px 0;padding:13px 15px;">'
                '<p style="margin:0;font-size:14px;line-height:1.85;color:%s;">%s</p></section>'
                % (TIFFANY_L, TIFFANY, INK, inline(text)))
    return ('<p style="margin:0 0 16px;font-size:15px;line-height:1.9;color:%s;'
            'letter-spacing:.4px;text-align:justify;">%s</p>' % (BODY, inline(text)))

def r_code(lang, code):
    label = lang or infer_lang(code)
    return ('<section style="margin:0 0 20px;border-radius:10px;overflow:hidden;background-color:%s;">'
            '<p style="margin:0;padding:7px 15px;background-color:%s;font-size:11px;letter-spacing:1.5px;'
            'color:#8B93A1;font-family:%s;">%s</p>'
            '<p style="margin:0;padding:14px 16px;font-family:%s;font-size:12.5px;line-height:1.85;'
            'color:%s;white-space:pre-wrap;word-break:break-all;">%s</p></section>'
            % (CODE_BG, CODE_HDR, MONO, html.escape(label), MONO, CODE_FG,
               highlight_code(code)))

def infer_lang(code):
    """根据内容推断代码块标签。"""
    lines = [l.strip() for l in code.split("\n") if l.strip() and not l.strip().startswith("#")]
    if lines and all(l.startswith("export ") for l in lines):
        return "环境变量"
    return "SHELL"

def r_img(src):
    return ('<section style="margin:0 0 20px;">'
            '<img src="%s" alt="" style="width:100%%;display:block;border-radius:10px;" /></section>' % src)

# ---------------------------------------------------------------- 组装 HTML
def build_html(blocks):
    title, toc, body = "终端使用 AI 的几种方法", [], []
    for b in blocks:
        if b[0] == "head":
            if b[1] == 1:
                title = b[2].replace("终端使用AI 的", "终端使用 AI 的"); continue
            if b[1] == 2:
                toc.append(b[2])
            body.append(r_head(b[1], b[2]))
        elif b[0] == "para":   body.append(r_para(b[1]))
        elif b[0] == "code":   body.append(r_code(b[1], b[2]))
        elif b[0] == "img":    body.append(r_img(b[2]))

    toc_html = ""
    if toc:
        items = "".join(
            '<p style="margin:0 0 8px;font-size:14px;line-height:1.9;color:%s;">'
            '<span style="color:%s;font-weight:700;margin-right:8px;">%02d</span>%s</p>'
            % (INK, ORANGE, n, inline(t)) for n, t in enumerate(toc, 1))
        toc_html = ('<section style="margin:22px 0 0;background-color:#FAFAFB;border-radius:12px;'
                    'padding:18px 20px;">'
                    '<p style="margin:0 0 12px;font-size:12px;letter-spacing:3px;color:%s;'
                    'font-weight:700;">C O N T E N T S</p>%s</section>' % (TIFFANY, items))

    cover = (
        '<section style="background-color:%s;background-image:linear-gradient(135deg,%s 0%%,%s 100%%);'
        'border-radius:14px;padding:30px 22px 26px;">'
        '<p style="margin:0 0 12px;font-size:11px;letter-spacing:4px;color:#FFD9C0;font-weight:700;">'
        'T E R M I N A L &nbsp;&#183;&nbsp; A I</p>'
        '<p style="margin:0;font-size:26px;line-height:1.42;font-weight:800;color:#FFFFFF;'
        'letter-spacing:.5px;">%s</p>'
        '<p style="margin:14px 0 0;font-size:13px;line-height:1.9;color:#FFE9DC;">'
        'Claude &#183; Codex &#183; opencode 三件套<br>统一走 DeepSeek 配置，全程不需要翻墙</p>'
        '<section style="height:3px;width:56px;background-color:%s;border-radius:2px;'
        'margin-top:18px;"></section></section>' % (ORANGE, ORANGE, ORANGE_D, html.escape(title), TIFFANY))

    footer = ('<section style="margin:34px 0 0;height:2px;border-radius:1px;background-image:'
              'linear-gradient(90deg,%s 0%%,%s 100%%);opacity:.45;"></section>'
              '<p style="margin:18px 0 0;text-align:center;font-size:12px;letter-spacing:3px;'
              'color:%s;">&#8212;&#8212; 完 &#8212;&#8212;</p>' % (ORANGE, TIFFANY, MUTED))

    return ("<!DOCTYPE html>\n<html lang=\"zh-CN\">\n<head>\n<meta charset=\"utf-8\">\n"
            "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
            "<title>%s</title>\n</head>\n<body style=\"margin:0;padding:0;background-color:#F5F5F7;\">\n"
            "<section style=\"max-width:677px;margin:0 auto;padding:22px 18px 40px;background-color:#FFFFFF;"
            "font-family:%s;color:%s;-webkit-font-smoothing:antialiased;\">\n%s\n%s\n%s\n%s\n</section>\n"
            "</body>\n</html>\n" % (html.escape(title), SANS, BODY, cover, toc_html, "\n".join(body), footer))

# ---------------------------------------------------------------- 生成 docx
def inline_images(html_text):
    """把 ./assets/xxx.png 换成 base64 data URI,产出不依赖外部文件夹的单文件 HTML。"""
    def repl(m):
        p = Path(unquote(m.group(1)))
        if not p.exists():
            return m.group(0)
        mime = mimetypes.guess_type(p.name)[0] or "image/png"
        b64 = base64.b64encode(p.read_bytes()).decode("ascii")
        return 'src="data:%s;base64,%s"' % (mime, b64)
    return re.sub(r'src="(\./assets/[^"]+)"', repl, html_text)


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def wq(t): return "{%s}%s" % (W, t)

def style_rpr(styles_root, style_id, font, half_pt, color, bold):
    for st in styles_root.iter(wq("style")):
        if st.get(wq("styleId")) != style_id:
            continue
        for tag in ("pPr", "rPr"):
            el = st.find(wq(tag))
            if el is not None:
                st.remove(el)
        pPr = etree.SubElement(st, wq("pPr"))
        rPr = etree.SubElement(st, wq("rPr"))
        rf = etree.SubElement(rPr, wq("rFonts"))
        for a in ("ascii", "hAnsi", "eastAsia", "cs"):
            rf.set(wq(a), font)
        if bold:
            etree.SubElement(rPr, wq("b"))
        c = etree.SubElement(rPr, wq("color")); c.set(wq("val"), color.lstrip("#"))
        sz = etree.SubElement(rPr, wq("sz")); sz.set(wq("val"), str(half_pt))
        etree.SubElement(rPr, wq("szCs")).set(wq("val"), str(half_pt))
        return

def build_docx(clean_md: Path, out: Path):
    tmp = Path(tempfile.mkdtemp())
    ref = tmp / "ref.docx"
    with open(ref, "wb") as f:
        subprocess.run(["pandoc", "--print-default-data-file", "reference.docx"],
                       stdout=f, check=True)
    d = tmp / "ref"; d.mkdir()
    with zipfile.ZipFile(ref) as z:
        z.extractall(d)
    sp = d / "word" / "styles.xml"
    tree = etree.parse(str(sp)); root = tree.getroot()

    for rfp in root.iter(wq("rFonts")):
        for a in ("ascii", "hAnsi", "eastAsia", "cs"):
            rfp.set(wq(a), "PingFang SC")
    for sid, size, color, bold in (("Heading1", 34, ORANGE, True),
                                   ("Heading2", 30, ORANGE_D, True),
                                   ("Heading3", 27, TIFFANY_D, True),
                                   ("Heading4", 25, INK, True),
                                   ("Heading5", 24, TIFFANY_D, True),
                                   ("BodyText", 21, BODY, False),
                                   ("FirstParagraph", 21, BODY, False)):
        style_rpr(root, sid, "PingFang SC", size, color, bold)
    tree.write(str(sp), xml_declaration=True, encoding="UTF-8", standalone=True)

    new_ref = tmp / "ref2.docx"
    with zipfile.ZipFile(new_ref, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(d.rglob("*")):
            if p.is_file():
                z.write(p, p.relative_to(d).as_posix())

    subprocess.run(["pandoc", str(clean_md.resolve()), "-o", str(out.resolve()),
                    "--reference-doc=%s" % new_ref,
                    "--resource-path=%s" % Path.cwd().resolve()], check=True)

# ---------------------------------------------------------------- main
def main():
    if not SRC.exists():
        sys.exit("找不到输入文件: %s" % SRC)
    raw = SRC.read_text(encoding="utf-8")

    # ---- 源 md 规范化(幂等):清空无意义的图片 alt;修正 opencode 那组步骤标题的层级 ----
    fixed = re.sub(r"!\[[^\]]*\]\(", "![](", raw)
    fixed = re.sub(r"^####\s+1\.\s+启动 opencode", "##### 1. 启动 opencode", fixed, flags=re.M)
    if fixed != raw:
        SRC.write_text(fixed, encoding="utf-8")
        print("[OK] 已规范化源 md(图片 alt 清空 / 步骤标题层级修正)")
    clean = fixed

    # ---- 仅用于渲染的内容修正(不改动源 md)----
    for old, new in PER_ARTICLE_FIXES:
        clean = clean.replace(old, new)
    clean = clean.replace("[TOC]\n", "")   # 目录改用样式化卡片呈现

    blocks = parse(clean)
    html_text = build_html(blocks)
    OUT_HTML.write_text(html_text, encoding="utf-8")
    print("[OK] HTML ->", OUT_HTML, "(%.1f KB)" % (OUT_HTML.stat().st_size / 1024))

    OUT_HTML_SINGLE.write_text(inline_images(html_text), encoding="utf-8")
    print("[OK] 单文件 HTML ->", OUT_HTML_SINGLE, "(%.2f MB)" % (OUT_HTML_SINGLE.stat().st_size / 1024 / 1024))

    tmp_md = Path(tempfile.mkdtemp()) / "clean.md"
    # pandoc 会把图片 alt 文本渲染成可见图注,这里清空 alt,只保留图片本身
    docx_md = re.sub(r"!\[[^\]]*\]\(", "![](", clean)
    tmp_md.write_text(docx_md, encoding="utf-8")
    try:
        build_docx(tmp_md, OUT_DOCX)
        print("[OK] DOCX ->", OUT_DOCX, "(%.1f KB)" % (OUT_DOCX.stat().st_size / 1024))
    except Exception as e:
        print("[WARN] docx 生成失败:", e, file=sys.stderr)

if __name__ == "__main__":
    main()
