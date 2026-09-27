# -*- coding: utf-8 -*-
"""
Build index.html from 简历.docx.

简历.docx is the SINGLE SOURCE OF TRUTH. Nothing is hardcoded: the title, the
table geometry, cell shading, cell alignment, bold/size/colour of every run and
the paragraph structure are all read out of the Word file. Edit Word, re-run
this script, and index.html follows.

Usage:
    python build_html.py            # build index.html
    python build_html.py --check    # build + verify every docx text is present

Exit code 0 = OK, 1 = build/consistency problem.
"""
import html as H
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from docx import Document
from docx.oxml.ns import qn

BASE = os.path.dirname(os.path.abspath(__file__))
DOCX = os.path.join(BASE, "简历.docx")
HTML_OUT = os.path.join(BASE, "index.html")

# ---- design constants (used only as fallbacks / chrome around Word content) --
BODY_FONT = '"SimSun", "宋体", "Songti SC", "Noto Serif CJK SC", serif'
BORDER = "#B8C4D8"
DEFAULT_PT = 10.5
TITLE_PT = 20
TITLE_COLOR = "#1F4E9C"
PAGE_W = "18cm"

# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------
_MULTISPACE = re.compile(r"[ \t\u00a0]{2,}")


def collapse(s):
    """Collapse runs of meaningless whitespace, keep single separators."""
    s = s.replace("\u00a0", " ")
    s = _MULTISPACE.sub(" ", s)
    return s


def esc(s):
    return H.escape(s, quote=False)


def half_points_to_pt(sz):
    try:
        return int(sz) / 2.0
    except (TypeError, ValueError):
        return None


def fmt_pt(pt):
    return ("%g" % round(pt, 2)) + "pt"


# ---------------------------------------------------------------------------
# run-level extraction
# ---------------------------------------------------------------------------
def run_props(r):
    rPr = r.find(qn("w:rPr"))
    p = {"b": False, "i": False, "sz": None, "color": None, "font": None}
    if rPr is None:
        return p
    if rPr.find(qn("w:b")) is not None:
        b = rPr.find(qn("w:b"))
        p["b"] = b.get(qn("w:val")) not in ("0", "false", "off")
    if rPr.find(qn("w:i")) is not None:
        p["i"] = True
    s = rPr.find(qn("w:sz"))
    if s is not None:
        p["sz"] = s.get(qn("w:val"))
    c = rPr.find(qn("w:color"))
    if c is not None:
        col = c.get(qn("w:val"))
        if col and col.lower() not in ("auto", "000000"):
            p["color"] = col
    f = rPr.find(qn("w:rFonts"))
    if f is not None:
        p["font"] = f.get(qn("w:eastAsia")) or f.get(qn("w:ascii"))
    return p


def run_css(p):
    css = []
    if p["b"]:
        css.append("font-weight:700")
    if p["i"]:
        css.append("font-style:italic")
    if p["sz"]:
        pt = half_points_to_pt(p["sz"])
        if pt:
            css.append("font-size:" + fmt_pt(pt))
    if p["color"]:
        css.append("color:#" + p["color"])
    if p["font"]:
        css.append('font-family:"%s",%s' % (p["font"], BODY_FONT))
    return ";".join(css)


def para_runs(p):
    """Return [(text, props), ...] for one w:p."""
    out = []
    for r in p.findall(qn("w:r")):
        t = "".join(x.text or "" for x in r.iter(qn("w:t")))
        if t:
            out.append((t, run_props(r)))
    return out


def group_runs(runs):
    """Merge consecutive runs that share styling -> [(text, props), ...]."""
    groups = []
    for text, props in runs:
        if groups and groups[-1][1] == props:
            groups[-1][0] += text
        else:
            groups.append([text, props])
    return [(t, p) for t, p in groups]


def render_para(runs):
    """Render runs of ONE paragraph -> (hoistable_css, inline_html).

    Whitespace is normalised across the whole paragraph so that spaces split
    across runs behave like one string, while each run keeps its own styling.

    Literal full-width spaces (U+3000) typed at the start of a paragraph are
    treated as a manual first-line indent and converted to text-indent, so
    they survive HTML whitespace collapsing. Real first-line indent set in
    Word (paragraph format) is read separately in para_ind_css().

    When the whole paragraph shares one style (the common case) that styling is
    returned as `hoistable_css` so the caller can put it straight on the
    wrapper element instead of emitting a redundant <span>.
    """
    groups = group_runs(runs)
    if not groups:
        return "", ""
    raw = "".join(t for t, _ in groups)
    lead = raw[:len(raw) - len(raw.lstrip("\u3000"))]
    raw = raw[len(lead):]
    whole = collapse(raw).strip()
    if not whole:
        return "", ""
    css = run_css(groups[0][1]) if len(groups) == 1 else ""
    if lead:
        css = (css + ";" if css else "") + "text-indent:%gem" % len(lead)
    if len(groups) == 1:
        return css, esc(whole)
    parts = []
    for i, (t, p) in enumerate(groups):
        t = collapse(t)
        if i == 0:
            t = t.lstrip()
        if i == len(groups) - 1:
            t = t.rstrip()
        if not t:
            continue
        rcss = run_css(p)
        if rcss:
            parts.append('<span style="%s">%s</span>' % (rcss, esc(t)))
        else:
            parts.append(esc(t))
    return css, "".join(parts)


# ---------------------------------------------------------------------------
# paragraph / cell extraction
# ---------------------------------------------------------------------------
def para_jc(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    j = pPr.find(qn("w:jc"))
    return j.get(qn("w:val")) if j is not None else None


def para_ind_css(p):
    """Word paragraph indent (w:pPr/w:ind) -> CSS.

    Chinese resumes indent the first line via paragraph FORMAT ("首行缩进2字符"),
    not via space characters, so this attribute must be read explicitly:
      firstLineChars=200  -> text-indent:2em   (2 characters, font-relative)
      firstLine=420       -> text-indent:21pt  (420 twips, fallback)
      left=120            -> margin-left:6pt   (whole-paragraph left indent)
      hanging=*           -> out-dented first line (padding + negative indent)
    """
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return []
    e = pPr.find(qn("w:ind"))
    if e is None:
        return []

    def num(attr):
        v = e.get(qn("w:" + attr))
        try:
            return int(v) if v is not None else None
        except (TypeError, ValueError):
            return None

    css = []
    flc = num("firstLineChars")            # 1/100 char units -> em
    fl = num("firstLine")                  # twips -> pt
    hgc = num("hangingChars")
    hg = num("hanging")
    lcc = num("leftChars") or num("startChars")
    lw = num("left") or num("start")

    if hgc:
        hang = "%gem" % (hgc / 100.0)
    elif hg:
        hang = "%gpt" % (hg / 20.0)
    else:
        hang = None
    if hang:
        # hanging indent: every line except the first is pushed in
        css.append("padding-left:%s" % hang)
        css.append("text-indent:-%s" % hang)
    elif flc:
        css.append("text-indent:%gem" % (flc / 100.0))
    elif fl:
        css.append("text-indent:%gpt" % (fl / 20.0))
    if lcc:
        css.append("margin-left:%gem" % (lcc / 100.0))
    elif lw:
        css.append("margin-left:%gpt" % (lw / 20.0))
    return css


def para_style(p):
    """(inline css, text-html) for a single w:p, or None when empty."""
    run_css_s, body = render_para(para_runs(p))
    if not body:
        return None
    css = []
    if run_css_s:
        css.append(run_css_s)
    css.extend(para_ind_css(p))
    jc = para_jc(p)
    if jc == "center":
        css.append("text-align:center")
    elif jc == "right":
        css.append("text-align:right")
    elif jc == "both":
        css.append("text-align:justify")
    return ";".join(css), body


def cell_info(tc):
    tcPr = tc.find(qn("w:tcPr"))
    info = {"span": 1, "fill": None, "valign": None, "vmerge": None, "paras": []}
    if tcPr is not None:
        gs = tcPr.find(qn("w:gridSpan"))
        if gs is not None:
            try:
                info["span"] = int(gs.get(qn("w:val")))
            except (TypeError, ValueError):
                pass
        shd = tcPr.find(qn("w:shd"))
        if shd is not None:
            fill = shd.get(qn("w:fill"))
            if fill and fill.lower() not in ("auto", "ffffff"):
                info["fill"] = fill
        va = tcPr.find(qn("w:vAlign"))
        if va is not None:
            info["valign"] = va.get(qn("w:val"))
        vm = tcPr.find(qn("w:vMerge"))
        if vm is not None:
            info["vmerge"] = vm.get(qn("w:val")) or "continue"
    for p in tc.findall(qn("w:p")):
        st = para_style(p)
        if st:
            info["paras"].append(st)
    return info


def table_grid_pct(tbl):
    """Column widths as percentages from w:tblGrid."""
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        return None
    widths = []
    for gc in grid.findall(qn("w:gridCol")):
        try:
            widths.append(int(gc.get(qn("w:w"))))
        except (TypeError, ValueError):
            widths.append(0)
    total = sum(widths)
    if not total:
        return None
    return ["%.2f%%" % (w * 100.0 / total) for w in widths]


# ---------------------------------------------------------------------------
# table rendering
# ---------------------------------------------------------------------------
def render_table(tbl):
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        rows.append([cell_info(tc) for tc in tr.findall(qn("w:tc"))])

    # resolve vertical merges (same cell index scan; no vMerge in this doc but
    # the Word format allows it and we must not silently drop content)
    rowspan = {}
    skip = set()
    for ri, cells in enumerate(rows):
        for ci, c in enumerate(cells):
            if c["vmerge"] == "restart":
                n = 1
                for rj in range(ri + 1, len(rows)):
                    nxt = rows[rj][ci] if ci < len(rows[rj]) else None
                    if nxt is not None and nxt["vmerge"] == "continue":
                        skip.add((rj, ci))
                        n += 1
                    else:
                        break
                rowspan[(ri, ci)] = n

    grid = table_grid_pct(tbl)
    out = ['    <table class="resume">']
    if grid:
        out.append("      <colgroup>")
        for w in grid:
            out.append('        <col style="width:%s">' % w)
        out.append("      </colgroup>")

    for ri, cells in enumerate(rows):
        out.append("      <tr>")
        for ci, c in enumerate(cells):
            if (ri, ci) in skip:
                continue
            attrs = []
            if c["span"] > 1:
                attrs.append('colspan="%d"' % c["span"])
            if (ri, ci) in rowspan and rowspan[(ri, ci)] > 1:
                attrs.append('rowspan="%d"' % rowspan[(ri, ci)])
            css = []
            if c["fill"]:
                css.append("background:#" + c["fill"])
            if c["valign"] == "center":
                css.append("vertical-align:middle")
            elif c["valign"] == "bottom":
                css.append("vertical-align:bottom")
            if css:
                attrs.append('style="%s"' % ";".join(css))
            attr_s = (" " + " ".join(attrs)) if attrs else ""
            if not c["paras"]:
                out.append("        <td%s></td>" % attr_s)
                continue
            out.append("        <td%s>" % attr_s)
            for pj, body in c["paras"]:
                if pj:
                    out.append('          <div class="para" style="%s">%s</div>' % (pj, body))
                else:
                    out.append('          <div class="para">%s</div>' % body)
            out.append("        </td>")
        out.append("      </tr>")
    out.append("    </table>")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# document rendering (body order preserved: paragraphs AND tables)
# ---------------------------------------------------------------------------
def body_blocks(doc):
    """Yield ('p', paragraph_element) / ('tbl', table_element) in body order."""
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield "p", child
        elif tag == "tbl":
            yield "tbl", child


def render_title(p_el):
    """<h1> from the first non-empty paragraph, honouring its own formatting."""
    runs = para_runs(p_el)
    run_css_s, _body = render_para(runs)
    text = collapse("".join(t for t, _ in runs)).strip()
    if not text:
        return None
    # title styling: prefer explicit run props, else the design default
    css = ["text-align:center", "letter-spacing:4px",
           "font-size:" + fmt_pt(TITLE_PT), "font-weight:700",
           "color:" + TITLE_COLOR]
    explicit = []
    for _, props in runs:
        if props.get("sz"):
            pt = half_points_to_pt(props["sz"])
            if pt:
                explicit = [c for c in explicit if not c.startswith("font-size")]
                explicit.append("font-size:" + fmt_pt(pt))
        if props.get("b"):
            explicit.append("font-weight:700")
        if props.get("color"):
            explicit = [c for c in explicit if not c.startswith("color")]
            explicit.append("color:#" + props["color"])
    if not explicit and run_css_s:
        explicit = [c for c in run_css_s.split(";") if c]
    if explicit:
        css = ["text-align:center", "letter-spacing:4px"] + explicit
    return '  <h1 class="title" style="%s">%s</h1>' % (";".join(css), esc(text))


def build():
    doc = Document(DOCX)

    blocks = []
    title_done = False
    for kind, el in body_blocks(doc):
        if kind == "p":
            if not title_done:
                h1 = render_title(el)
                if h1:
                    blocks.append(h1)
                    title_done = True
                continue
            st = para_style(el)
            if st:
                pj, body = st
                blocks.append('  <p class="line"%s>%s</p>'
                              % ((' style="%s"' % pj) if pj else "", body))
        else:
            blocks.append(render_table(el))

    html_doc = HTML_TEMPLATE.replace("__BODY__", "\n".join(blocks))
    with open(HTML_OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write(html_doc)
    return html_doc


CSS = """  * { margin: 0; padding: 0; box-sizing: border-box; }
  html, body {
    font-family: %(font)s;
    color: #1A1A1A;
    background: #EEF1F6;
    -webkit-text-size-adjust: 100%%;
  }
  .page {
    width: 100%%;
    max-width: %(pagew)s;
    margin: 0 auto;
    padding: 0.6cm 0 1.4cm;
  }
  h1.title {
    font-family: %(font)s;
    margin: 0 0 14px;
  }
  table.resume {
    width: 100%%;
    border-collapse: collapse;
    table-layout: fixed;
    background: #FFFFFF;
    font-size: %(dsize)s;
    line-height: 1.5;
    border: 1px solid %(border)s;
  }
  table.resume td {
    border: 1px solid %(border)s;
    padding: 5px 8px;
    vertical-align: middle;
    word-break: break-word;
    overflow-wrap: anywhere;
  }
  .para + .para { margin-top: 2px; }
  p.line { font-size: %(dsize)s; line-height: 1.5; margin: 0 0 6px; }

  /* ---- floating "save as PDF" button (screen only) ---- */
  /* an <a download> pointing at the pre-rendered resume.pdf: clicking it
     starts a browser download directly - colors, margins and page content
     are exactly what weasyprint rendered, with NO browser print headers
     (URL / date / title) and no "background graphics" surprises. */
  #save-pdf {
    position: fixed;
    right: 20px;
    bottom: 20px;
    z-index: 9999;
    padding: 11px 20px;
    font-family: %(font)s;
    font-size: 14px;
    color: #fff;
    background: %(title_color)s;
    border: none;
    border-radius: 999px;
    box-shadow: 0 3px 10px rgba(0, 0, 0, 0.25);
    cursor: pointer;
    text-decoration: none;
  }
  #save-pdf:hover { background: #163A75; }

  /* ---- print / Ctrl+P fallback ---- */
  /* keep background colors when the page is printed anyway */
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  @page { size: A4; margin: 1.2cm 1.6cm 1.1cm 1.6cm; }
  @media print {
    html, body { background: #fff; }
    .page { max-width: none; padding: 0; }
    #save-pdf { display: none !important; }
    table.resume { border: 1px solid %(border)s; }
    tr { page-break-inside: avoid; }
    thead { display: table-header-group; }
  }
""" % {
    "font": BODY_FONT,
    "pagew": PAGE_W,
    "dsize": fmt_pt(DEFAULT_PT),
    "border": BORDER,
    "title_color": TITLE_COLOR,
}

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>个人简历</title>
<style>
__CSS__
</style>
</head>
<body>
  <div class="page">
__BODY__
  </div>
  <a id="save-pdf" href="resume.pdf" download="个人简历-赵浩.pdf">保存为 PDF</a>
</body>
</html>
""".replace("__CSS__", CSS)


# ---------------------------------------------------------------------------
# consistency check: every docx text must appear in the HTML, in order
# ---------------------------------------------------------------------------
def check(html_doc, doc):
    def clean(s):
        return re.sub(r"\s+", "", s)

    texts = []
    for p in doc.paragraphs:
        if p.text.strip():
            texts.append(p.text)
    for tbl in doc.tables:
        for tr in tbl._tbl.findall(qn("w:tr")):
            for tc in tr.findall(qn("w:tc")):
                for p in tc.findall(qn("w:p")):
                    s = "".join(t.text or "" for t in p.iter(qn("w:t")))
                    if s.strip():
                        texts.append(s)

    body = html_doc.split("<body", 1)[1].split("</body>", 1)[0]
    body = re.sub(r"<script.*?</script>", "", body, flags=re.S)
    visible = H.unescape(re.sub(r"<[^>]+>", "", body))
    visible = re.sub(r"保存为 PDF", "", visible)
    htext = clean(visible)

    missing = [t for t in texts if clean(t) not in htext]
    # order check
    order_ok = True
    pos = -1
    for t in texts:
        i = htext.find(clean(t))
        if i < pos:
            order_ok = False
            break
        pos = i
    return len(texts), missing, order_ok


PDF_OUT = "resume.pdf"


def build_pdf():
    """Regenerate resume.pdf from index.html (best effort).

    weasyprint renders exactly what the CSS defines: background colors kept,
    the #save-pdf button hidden via @media print, A4 margins from @page, and
    NO browser print headers (URL / date / title). If weasyprint is not
    installed the existing resume.pdf (if any) is left untouched.
    """
    try:
        import weasyprint
    except Exception as e:
        print("PDF_SKIPPED weasyprint unavailable (%s)" % e.__class__.__name__)
        return False
    weasyprint.HTML(filename=HTML_OUT).write_pdf(PDF_OUT)
    print("PDF_OK %s bytes=%d" % (PDF_OUT, os.path.getsize(PDF_OUT)))
    return True


def main():
    if not os.path.exists(DOCX):
        print("ERROR: 简历.docx not found at %s" % DOCX)
        return 1
    html_doc = build()
    n, missing, order_ok = check(html_doc, Document(DOCX))
    size = os.path.getsize(HTML_OUT)
    print("BUILD_OK index.html bytes=%d texts=%d missing=%d order_ok=%s"
          % (size, n, len(missing), order_ok))
    for m in missing[:10]:
        print("  MISSING: %r" % m)
    if missing or not order_ok:
        return 1
    build_pdf()
    return 0


if __name__ == "__main__":
    sys.exit(main())