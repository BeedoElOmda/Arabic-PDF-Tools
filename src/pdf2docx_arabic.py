"""PDF -> editable Word (.docx) with correct Arabic.

1. Repairs the PDF's Arabic letter maps (ligatures / presentation forms ->
   normal letters, see fix_arabic_pdf.py), so no boxes appear.
2. Rebuilds the layout with pdf2docx (paragraphs, real Word tables, images).
3. Makes every Arabic paragraph proper right-to-left Word text: RTL paragraph
   direction, RTL runs, Arabic ("complex script") font, size and bold.

Output: "<name> (Word).docx" next to the PDF.
Usage: pythonw pdf2docx_arabic.py file1.pdf [file2.pdf ...]
"""
import copy
import os
import re
import sys
import tempfile
import traceback

FONT = "Arial"   # Arabic-capable font present on every Windows PC
ARABIC = re.compile(r"[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(tag):
    return "{%s}%s" % (W, tag)


# Schema order of the child elements we touch (Word is strict about order).
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
             "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
             "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
             "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
             "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
             "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
             "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
             "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
             "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
             "specVanish", "oMath"]


def sort_children(el, order):
    rank = {q(t): i for i, t in enumerate(order)}
    kids = list(el)
    kids.sort(key=lambda k: rank.get(k.tag, len(order)))
    for k in kids:
        el.remove(k)
    for k in kids:
        el.append(k)


def get_or_add(parent, tag):
    el = parent.find(q(tag))
    if el is None:
        el = parent.makeelement(q(tag), {})
        parent.append(el)
    return el


def para_text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def fix_run(r, rtl):
    rpr = r.find(q("rPr"))
    if rpr is None:
        rpr = r.makeelement(q("rPr"), {})
        r.insert(0, rpr)
    f = get_or_add(rpr, "rFonts")
    for a in ("ascii", "hAnsi", "cs"):
        f.set(q(a), FONT)
    if rpr.find(q("b")) is not None and rpr.find(q("b")).get(q("val")) not in ("0", "false"):
        get_or_add(rpr, "bCs")
    if rpr.find(q("i")) is not None and rpr.find(q("i")).get(q("val")) not in ("0", "false"):
        get_or_add(rpr, "iCs")
    sz = rpr.find(q("sz"))
    if sz is not None:
        get_or_add(rpr, "szCs").set(q("val"), sz.get(q("val")))
    text = "".join(t.text or "" for t in r.iter(q("t")))
    if rtl and ARABIC.search(text):
        get_or_add(rpr, "rtl")
        get_or_add(rpr, "lang").set(q("bidi"), "ar-EG")
    sort_children(rpr, RPR_ORDER)


def merge_runs(p):
    """Join neighbouring runs with identical formatting (keeps Arabic letters joined)."""
    from lxml import etree
    runs = [r for r in p if r.tag == q("r")]
    prev = None
    for r in runs:
        ts = r.findall(q("t"))
        simple = len(ts) == 1 and all(c.tag in (q("rPr"), q("t")) for c in r)
        if prev is not None and simple and r.getprevious() is prev:
            a = prev.find(q("rPr"))
            b = r.find(q("rPr"))
            if etree.tostring(a if a is not None else etree.Element("x")) == \
               etree.tostring(b if b is not None else etree.Element("x")):
                pt = prev.find(q("t"))
                pt.text = (pt.text or "") + (ts[0].text or "")
                pt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                p.remove(r)
                continue
        prev = r if simple else None


CONTROL = re.compile(r"[\x00-\x08\x0b-\x1f\x7f]")


def fix_paragraph(p):
    for t in p.iter(q("t")):          # drop invisible control characters some PDFs carry
        if t.text and CONTROL.search(t.text):
            t.text = CONTROL.sub("", t.text)
    text = para_text(p)
    rtl = bool(ARABIC.search(text))
    merge_runs(p)
    for r in p.iter(q("r")):
        fix_run(r, rtl)
    # Rows laid out with tab stops are measured left-to-right: keep them LTR
    # (their Arabic runs are already marked RTL, so the words read correctly).
    if not rtl or p.find(".//" + q("tab")) is not None and p.find(q("pPr") + "/" + q("tabs")) is not None:
        return
    ppr = p.find(q("pPr"))
    if ppr is None:
        ppr = p.makeelement(q("pPr"), {})
        p.insert(0, ppr)
    get_or_add(ppr, "bidi")
    jc = ppr.find(q("jc"))
    val = jc.get(q("val")) if jc is not None else None
    # In an RTL paragraph Word's "left/right" mean start/end: map to what the PDF showed.
    if val in ("both", "distribute", "center"):
        pass
    else:
        get_or_add(ppr, "jc").set(q("val"), "left")   # = right-aligned in RTL paragraph
    ind = ppr.find(q("ind"))
    if ind is not None:   # pdf2docx writes physical indents; swap them for RTL
        l, rgt = ind.get(q("left")), ind.get(q("right"))
        if l is not None or rgt is not None:
            ind.set(q("left"), rgt or "0")
            ind.set(q("right"), l or "0")
    sort_children(ppr, PPR_ORDER)


def fix_styles(doc):
    """Default document font: Arial for Latin and Arabic."""
    st = doc.styles.element
    rpr_default = st.find(q("docDefaults"))
    if rpr_default is None:
        return
    rpr = rpr_default.find(q("rPrDefault") + "/" + q("rPr"))
    if rpr is None:
        return
    f = get_or_add(rpr, "rFonts")
    for a in ("ascii", "hAnsi", "cs"):
        f.set(q(a), FONT)
    sort_children(rpr, RPR_ORDER)


_WIN_FONTS = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts")
FONT_FILES = {   # used only to measure text width
    True: [os.path.join(_WIN_FONTS, "arialbd.ttf"), "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"],
    False: [os.path.join(_WIN_FONTS, "arial.ttf"), "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"],
}
_fonts = {}


def measure(text, size, bold):
    """Width in points of `text` in Arial (bold or regular), or None."""
    from PIL import ImageFont
    if bold not in _fonts:
        _fonts[bold] = None
        for ff in FONT_FILES[bold]:
            if os.path.exists(ff):
                try:
                    _fonts[bold] = ImageFont.truetype(ff, 100)
                    break
                except Exception:
                    pass
    f = _fonts[bold]
    if f is None:
        return None
    try:
        w = f.getlength(text, direction="rtl")
    except Exception:
        w = f.getlength(text)
    return w * size / 100


def font_scale(fixed_pdf):
    """How much to shrink text so Arial takes the same width as the PDF's font.

    Measures every Arabic text span in Arial (bold or regular, as in the PDF)
    and compares it with its width in the PDF; returns the median ratio.
    """
    import pymupdf
    ratios = []
    doc = pymupdf.open(fixed_pdf)
    for page in doc:
        for b in page.get_text("dict", flags=0)["blocks"]:
            for l in b.get("lines", []):
                for sp in l["spans"]:
                    t = sp["text"].strip()
                    if len(t) < 8 or not ARABIC.search(t):
                        continue
                    bold = bool(sp["flags"] & 16) or any(k in sp["font"].lower() for k in ("bold", "heading", "black"))
                    m = measure(t, sp["size"], bold)
                    w = sp["bbox"][2] - sp["bbox"][0]
                    if m and w > 0:
                        ratios.append(w / m)
    doc.close()
    if not ratios:
        return 1.0
    ratios.sort()
    return max(0.6, min(1.0, ratios[len(ratios) // 2] * 0.95))


TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
               "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
               "tblLook", "tblCaption", "tblDescription"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
              "textDirection", "tcFitText", "vAlign", "hideMark"]
BORDER_ORDER = ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl", "between", "bar"]


def schema_cleanup(d):
    """Fix pdf2docx output so Word opens it without complaints:
    colours without '#', border elements in schema order, zoom percent."""
    body = d.element.body
    for el in body.iter():
        for attr in (q("color"), q("fill")):
            v = el.get(attr)
            if v and v.startswith("#"):
                el.set(attr, v[1:].upper())
    for tag in ("tcBorders", "pBdr", "tblBorders"):
        for el in body.iter(q(tag)):
            sort_children(el, BORDER_ORDER)
    for el in body.iter(q("pPr")):
        sort_children(el, PPR_ORDER)
    for el in body.iter(q("rPr")):
        if el.getparent() is not None and el.getparent().tag in (q("r"), q("pPr")):
            sort_children(el, RPR_ORDER)
    for el in body.iter(q("tblPr")):
        sort_children(el, TBLPR_ORDER)
    for el in body.iter(q("tcPr")):
        seen_tags = set()
        for child in list(el):              # duplicates are not allowed
            if child.tag in seen_tags:
                el.remove(child)
            seen_tags.add(child.tag)
        sort_children(el, TCPR_ORDER)
    for el in body.iter(q("shd")):
        if el.get(q("val")) is None:
            el.set(q("val"), "clear")
    zoom = d.settings.element.find(q("zoom"))
    if zoom is not None and zoom.get(q("percent")) is None:
        zoom.set(q("percent"), "100")


def scale_sizes(root, scale):
    if scale >= 0.995:
        return
    for tag in ("sz", "szCs"):
        for el in root.iter(q(tag)):
            v = el.get(q("val"))
            if v and v.isdigit():
                el.set(q("val"), str(max(10, int(round(int(v) * scale)))))


def is_slide_deck(pdf_path):
    """Mostly landscape pages = slides -> keep the exact layout instead of flowing text."""
    import pymupdf
    doc = pymupdf.open(pdf_path)
    n = doc.page_count
    wide = sum(1 for p in doc if p.rect.width > p.rect.height * 1.15)
    doc.close()
    return n > 0 and wide >= 0.6 * n


def convert(pdf_path, mode="auto"):
    """mode: 'auto' (documents -> flowing text, slides -> layout), 'flow' or 'layout'."""
    if mode == "auto":
        mode = "layout" if is_slide_deck(pdf_path) else "flow"
    if mode == "flow":
        return convert_flow(pdf_path)
    return convert_layout(pdf_path)


def convert_flow(pdf_path):
    """Document PDFs: rebuild as flowing, editable right-to-left Word text."""
    import io
    import pikepdf
    import docx_flow
    from fix_arabic_pdf import fix_font, expand_ligatures, fill_missing_codes, UNKNOWN_GUESSES, DOT_PLACEHOLDERS

    pdf = pikepdf.open(pdf_path)
    seen, ligs, coverage = set(), {}, {}
    UNKNOWN_GUESSES.clear()
    DOT_PLACEHOLDERS.clear()
    fill_missing_codes(pdf, ligs)
    for obj in pdf.objects:
        if isinstance(obj, pikepdf.Dictionary) and obj.get("/Type") == "/Font":
            fix_font(obj, seen, ligs, coverage)
    buf = io.BytesIO()
    pdf.save(buf)
    pdf.close()
    base, _ = os.path.splitext(pdf_path)
    out = base + " (Word).docx"
    from fix_arabic_pdf import UNKNOWN_GUESSES
    n = docx_flow.convert(buf.getvalue(), out, lambda c: expand_ligatures(c, ligs), coverage,
                          {v: UNKNOWN_GUESSES.get(v, []) for k, v in ligs.items() if k.startswith("\x00")})
    return out, n


def convert_layout(pdf_path):
    """Slide-like PDFs: keep positions (pdf2docx) and fix Arabic direction/fonts."""
    import pikepdf
    import docx
    from pdf2docx import Converter
    from fix_arabic_pdf import fix_font, fill_missing_codes

    tmpdir = tempfile.mkdtemp(prefix="pdf2docx_ar_")
    fixed = os.path.join(tmpdir, "fixed.pdf")
    raw = os.path.join(tmpdir, "raw.docx")

    pdf = pikepdf.open(pdf_path)
    fill_missing_codes(pdf)
    seen = set()
    for obj in pdf.objects:
        if isinstance(obj, pikepdf.Dictionary) and obj.get("/Type") == "/Font":
            fix_font(obj, seen)
    pdf.save(fixed)
    pdf.close()

    cv = Converter(fixed)
    cv.convert(raw)
    n = len(cv.fitz_doc)
    cv.close()

    scale = font_scale(fixed)
    d = docx.Document(raw)
    fix_styles(d)
    parts = [d.element.body]
    for s in d.sections:
        for hf in (s.header, s.footer):
            if hf is not None and not hf.is_linked_to_previous:
                parts.append(hf._element)
    for part in parts:
        for p in list(part.iter(q("p"))):
            fix_paragraph(p)
        scale_sizes(part, scale)

    base, _ = os.path.splitext(pdf_path)
    out = base + " (Word).docx"
    schema_cleanup(d)
    d.core_properties.author = "Beedo El-Omda"
    d.core_properties.last_modified_by = "Beedo El-Omda"
    d.save(out)
    try:
        os.remove(fixed)
        os.remove(raw)
        os.rmdir(tmpdir)
    except OSError:
        pass
    return out, n


# ---------------------------------------------------------------- shell ---
def notify(title, msg):
    if os.name == "nt" and os.path.basename(sys.executable).lower() == "pythonw.exe":
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, title, 0x40)
    else:
        print(f"{title}: {msg}")


def install_sendto(pythonw):
    """Create the Send To shortcut (called by Install.cmd)."""
    import subprocess
    sendto = os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "SendTo")
    lnk = os.path.join(sendto, "PDF to Word (Arabic, editable).lnk")
    script = os.path.abspath(__file__)

    def ps_q(s):
        return "'" + s.replace("'", "''") + "'"

    icon = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "System32", "imageres.dll") + ",2"
    args = '"' + script + '"'
    ps = (
        f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut({ps_q(lnk)});"
        f"$s.TargetPath={ps_q(pythonw)};$s.Arguments={ps_q(args)};"
        f"$s.WorkingDirectory={ps_q(os.path.dirname(script))};$s.IconLocation={ps_q(icon)};"
        "$s.Description='Convert PDF to editable Word with correct Arabic';$s.Save()"
    )
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps], check=True)


def main():
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    args = sys.argv[1:]
    if len(args) == 2 and args[0] == "--install-sendto":
        install_sendto(args[1])
        return
    files = [f for f in args if f.lower().endswith(".pdf")]
    if not files:
        notify("PDF to Word", "Select PDF files, then right-click > Send to > PDF to Word (Arabic, editable).")
        return
    done, failed = [], []
    for f in files:
        try:
            out, n = convert(f)
            done.append(f"{os.path.basename(out)}  ({n} pages)")
        except Exception:
            failed.append(f"{os.path.basename(f)}:\n{traceback.format_exc(limit=2)}")
    msg = ("Created:\n" + "\n".join(done) + "\n\nEnjoy , Beedo El-Omda") if done else ""
    if failed:
        msg += ("\n\n" if msg else "") + "Failed:\n" + "\n".join(failed)
    notify("PDF to Word (Arabic, editable)", msg)


if __name__ == "__main__":
    main()
