"""PDF -> editable PPTX with correct Arabic.

1. Repairs the PDF's Arabic letter maps (ligatures / presentation forms ->
   normal letters), so no boxes appear.
2. For each page: renders the page WITHOUT its text as the slide background
   (pictures, shapes, colours stay exactly as in the PDF), then places every
   text back on top as ONE editable text box per shape / paragraph block,
   lines joined into reflowing right-to-left paragraphs (original size,
   colour, bold and line spacing).

Output: "<name> (editable).pptx" next to the PDF.
Usage: pythonw pdf2pptx.py file1.pdf [file2.pdf ...]
       add --image to get picture-only slides instead.
"""
import io
import os
import re
import sys
import traceback

DPI = 200
FONT = "Arial"            # Arabic-capable font present on every Windows PC
FONT_FILES = [  # used only to measure text width so lines keep their original length
    os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", "arialbd.ttf"),
    os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", "arial.ttf"),
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]
_measure = {}


def text_width(text, size):
    """Width in points of text set in FONT at `size`, or None if unknown."""
    from PIL import ImageFont
    key = round(size * 4)
    if key not in _measure:
        _measure[key] = None
        for ff in FONT_FILES:
            if os.path.exists(ff):
                try:
                    _measure[key] = ImageFont.truetype(ff, key)  # 4x for precision
                    break
                except Exception:
                    pass
    font = _measure[key]
    if font is None:
        return None
    try:
        return font.getlength(text, direction="rtl" if ARABIC.search(text) else None) / 4
    except Exception:
        return font.getlength(text) / 4


ARABIC = re.compile(r"[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]")


# ---------------------------------------------------------------- repair ---
def repaired_for_text(path):
    """(PDF bytes, expand function, font coverage) using the shared Arabic text engine."""
    import pikepdf
    from fix_arabic_pdf import fix_font, expand_ligatures, fill_missing_codes, UNKNOWN_GUESSES, DOT_PLACEHOLDERS

    pdf = pikepdf.open(path)
    seen, ligs, coverage = set(), {}, {}
    UNKNOWN_GUESSES.clear()
    DOT_PLACEHOLDERS.clear()
    fill_missing_codes(pdf, ligs)
    for obj in pdf.objects:
        if isinstance(obj, pikepdf.Dictionary) and obj.get("/Type") == "/Font":
            fix_font(obj, seen, ligs, coverage)
    buf = io.BytesIO()
    pdf.save(buf)
    guesses = {v: UNKNOWN_GUESSES.get(v, []) for k, v in ligs.items() if k.startswith("\x00")}
    return buf.getvalue(), (lambda c: expand_ligatures(c, ligs)), coverage, guesses


def engine_lines(page, rects, expand):
    import docx_flow
    return docx_flow.build_lines(docx_flow.page_glyphs(page, expand), rects, split_gap=1.2)


def engine_rows(lines):
    """Rows for the slide builder, from the shared text engine (correct order,
    repaired letters, no kashida, spaces fixed). Returns (rows, unreadable bboxes)."""
    rows, keep = [], []
    for ln in lines:
        bbox = (ln["x0"], ln["top"], ln["x1"], ln["bottom"])
        if ln.get("bad"):
            keep.append(bbox)          # unreliable font: leave it in the picture
            continue
        spans = []
        for text, (fam, size, bold, italic, color) in ln["runs"]:
            spans.append({"text": text, "size": size, "color": color, "bbox": bbox,
                          "flags": (16 if bold else 0) | (2 if italic else 0),
                          "font": fam + ("-Bold" if bold else ""), "origin": (ln["x0"], ln["y"])})
        if not spans:
            continue
        rows.append({"bbox": bbox, "rtl": ln["rtl"], "spans": spans, "size": ln["size"], "base": ln["y"]})
    return rows, keep


def repaired_pdf_bytes(path):
    """Return PDF bytes with Arabic letter maps repaired (see fix_arabic_pdf.py)."""
    import pikepdf
    from fix_arabic_pdf import fix_font, fill_missing_codes

    pdf = pikepdf.open(path)
    fill_missing_codes(pdf)
    seen = set()
    for obj in pdf.objects:
        if isinstance(obj, pikepdf.Dictionary) and obj.get("/Type") == "/Font":
            fix_font(obj, seen)
    buf = io.BytesIO()
    pdf.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------- text rows ---
def _spans(page):
    out = []
    for block in page.get_text("dict", flags=0)["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            if abs(line["dir"][1]) > 0.01:      # skip rotated text
                continue
            for s in line["spans"]:
                if s["text"].strip():
                    out.append(s)
    return out


def _finish_row(spans):
    rtl = any(ARABIC.search(s["text"]) for s in spans)
    spans.sort(key=lambda s: -s["bbox"][2] if rtl else s["bbox"][0])
    return {
        "bbox": (min(s["bbox"][0] for s in spans), min(s["bbox"][1] for s in spans),
                 max(s["bbox"][2] for s in spans), max(s["bbox"][3] for s in spans)),
        "rtl": rtl,
        "spans": spans,
        "size": max(s["size"] for s in spans),
        "base": max(s["origin"][1] for s in spans),
    }


def page_rows(page, rects=()):
    """Group spans into visual rows. Spans on the same line but far apart
    (separate columns / table cells) stay in separate rows."""
    spans = sorted(_spans(page), key=lambda s: -s["bbox"][2])
    rows = []
    for s in spans:
        x0, y0, x1, y1 = s["bbox"]
        cy = (y0 + y1) / 2
        box = container_of({"bbox": s["bbox"]}, rects)
        best = None
        for r in rows:
            if abs(r["cy"] - cy) > s["size"] * 0.4:
                continue
            if r["box"] != box:          # different shape / column / table cell
                continue
            rx0, rx1 = r["x0"], r["x1"]
            gap = max(rx0 - x1, x0 - rx1, 0)
            if gap < s["size"] * 1.2:
                best = r
                break
        if best is None:
            rows.append({"cy": cy, "x0": x0, "x1": x1, "spans": [s], "box": box})
        else:
            best["spans"].append(s)
            best["x0"], best["x1"] = min(best["x0"], x0), max(best["x1"], x1)
    return [_finish_row(r["spans"]) for r in rows]


# --------------------------------------------------------- smart grouping ---
def shape_rects(page):
    """Filled/outlined rectangles (incl. rounded) drawn on the page = text containers."""
    pr = page.rect
    rects = []
    try:
        drawings = page.get_drawings()
    except Exception:
        return rects
    for d in drawings:
        r = d.get("rect")
        if r is None or (d.get("fill") is None and d.get("color") is None):
            continue
        if r.width < 25 or r.height < 12:
            continue
        if r.width * r.height > 0.9 * pr.width * pr.height:
            continue
        rects.append((r.x0, r.y0, r.x1, r.y1))
    return rects


def image_rects(src_doc, pno):
    """Boxes that are part of a picture (slides exported with framed text boxes
    drawn as images): flat, nearly rectangular areas of the text-free page."""
    try:
        import numpy as np
        import cv2
        import pymupdf
    except Exception:
        return []
    tmp = pymupdf.open()
    tmp.insert_pdf(src_doc, from_page=pno, to_page=pno)
    p = tmp[0]
    for b in p.get_text("dict", flags=0)["blocks"]:
        if b.get("type") == 0:
            for l in b["lines"]:
                for s in l["spans"]:
                    p.add_redact_annot(pymupdf.Rect(s["bbox"]))
    p.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                       graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                       text=pymupdf.PDF_REDACT_TEXT_REMOVE)
    zoom = 1.5
    pix = p.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
    a = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3].astype(np.int16)
    T = 22
    dx = np.abs(np.diff(a, axis=1)).max(axis=2) > T
    dy = np.abs(np.diff(a, axis=0)).max(axis=2) > T
    edge = np.zeros(a.shape[:2], bool)
    edge[:, 1:] |= dx
    edge[:, :-1] |= dx
    edge[1:, :] |= dy
    edge[:-1, :] |= dy
    n, _, stats, _ = cv2.connectedComponentsWithStats((~edge).astype(np.uint8), connectivity=4)
    H, W = edge.shape
    out = []
    for i in range(1, n):
        x, y, w, h, area = (int(v) for v in stats[i])
        if w < 30 * zoom or h < 14 * zoom or w * h > 0.8 * W * H:
            continue
        if area < 0.72 * w * h:          # not a box: photo, gradient, irregular area
            continue
        out.append((x / zoom, y / zoom, (x + w) / zoom, (y + h) / zoom))
    return out


def container_of(row, rects):
    x0, y0, x1, y1 = row["bbox"]
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    best = None
    for r in rects:
        if r[0] - 3 <= x0 and x1 <= r[2] + 3 and r[1] - 3 <= cy <= r[3] + 3 and r[0] <= cx <= r[2]:
            area = (r[2] - r[0]) * (r[3] - r[1])
            if best is None or area < best[0]:
                best = (area, r)
    return best[1] if best else None


def group_boxes(rows, rects=()):
    """Merge consecutive rows that belong to the same paragraph/column into
    one text box."""
    boxes = []
    for r in sorted(rows, key=lambda r: (r["bbox"][1], -r["bbox"][2])):
        x0, y0, x1, y1 = r["bbox"]
        r["box"] = container_of(r, rects)
        target = None
        for b in boxes:
            last = b["rows"][-1]
            lx0, ly0, lx1, ly1 = last["bbox"]
            if b["rtl"] != r["rtl"] or b["box"] != r["box"]:
                continue
            if not (0.75 <= r["size"] / last["size"] <= 1.33):
                continue
            # next line = baseline 0.8x..1.9x the font size below the last line
            dbase = r["base"] - last["base"]
            if not (last["size"] * 0.8 <= dbase <= last["size"] * 1.9):
                continue
            bx0, bx1 = b["x0"], b["x1"]
            overlap = min(x1, bx1) - max(x0, bx0)
            if overlap < 0.5 * min(x1 - x0, bx1 - bx0):
                continue
            tol = r["size"] * 1.5
            aligned = (abs(x1 - bx1) < tol) or (abs(x0 - bx0) < tol) or \
                      (abs((x0 + x1) / 2 - (bx0 + bx1) / 2) < tol)
            if not aligned:
                continue
            target = b
            break
        if target is None:
            boxes.append({"rows": [r], "rtl": r["rtl"], "x0": x0, "x1": x1, "box": r["box"]})
        else:
            target["rows"].append(r)
            target["x0"], target["x1"] = min(target["x0"], x0), max(target["x1"], x1)
    for b in boxes:
        b["y0"] = min(r["bbox"][1] for r in b["rows"])
        b["y1"] = max(r["bbox"][3] for r in b["rows"])
    return boxes


def background_png(src_doc, pno, keep=()):
    """Render page with all text removed (images and shapes kept).
    Text inside `keep` boxes (unreadable fonts) stays in the picture."""
    import pymupdf
    tmp = pymupdf.open()
    tmp.insert_pdf(src_doc, from_page=pno, to_page=pno)
    p = tmp[0]
    for b in p.get_text("dict", flags=0)["blocks"]:
        if b.get("type") != 0:
            continue
        for l in b["lines"]:
            for s in l["spans"]:
                x0, y0, x1, y1 = s["bbox"]
                cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                if any(k[0] - 2 <= cx <= k[2] + 2 and k[1] - 2 <= cy <= k[3] + 2 for k in keep):
                    continue
                p.add_redact_annot(pymupdf.Rect(s["bbox"]))
    p.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                       graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                       text=pymupdf.PDF_REDACT_TEXT_REMOVE)
    return p.get_pixmap(dpi=DPI).tobytes("png")


# ------------------------------------------------------------ pptx build ---
def _row_text(row):
    parts, prev = [], None
    for s in row["spans"]:
        text = s["text"]
        if prev is not None:
            gap = (prev["bbox"][0] - s["bbox"][2]) if row["rtl"] else (s["bbox"][0] - prev["bbox"][2])
            if gap > s["size"] * 0.15 and not text.startswith(" ") and not prev["text"].endswith(" "):
                text = " " + text
        parts.append((s, text))
        prev = s
    return parts


BULLET = re.compile(r"^\s*([\u2022\u25AA\u25CF\u2013\-\*\u00B7]|[0-9\u0660-\u0669]+[\.\)\-:]|[\u0623-\u064A][\.\)\-])\s*")


def _style_run(r, s, scale, k, rtl):
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    f = r.font
    f.size = Pt(max(6, round(s["size"] * scale * k / 12700 * 2) / 2))
    f.bold = bool(s["flags"] & 16) or "bold" in s["font"].lower() or "heading" in s["font"].lower()
    f.italic = bool(s["flags"] & 2)
    c = s["color"]
    f.color.rgb = RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)
    f.name = FONT
    rPr = r._r.get_or_add_rPr()
    el = rPr.find(qn("a:cs"))
    if el is None:
        el = rPr.makeelement(qn("a:cs"), {})
        rPr.append(el)
    el.set("typeface", FONT)
    if rtl:
        rPr.set("lang", "ar-EG")


def add_box(slide, box, k):
    """One editable text box per paragraph group. k = EMU per PDF point."""
    from pptx.util import Emu, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.oxml.ns import qn

    rows, rtl = box["rows"], box["rtl"]
    x0, x1, y0, y1 = box["x0"], box["x1"], box["y0"], box["y1"]
    W = x1 - x0
    tol = max(r["size"] for r in rows) * 1.2

    # ---- alignment of the original text
    centred = len(rows) > 1 and all(abs((r["bbox"][0] + r["bbox"][2]) / 2 - (x0 + x1) / 2) < tol for r in rows) \
        and any(r["bbox"][2] - r["bbox"][0] < W - 2 * tol for r in rows)
    full = [r for r in rows[:-1] if r["bbox"][2] - r["bbox"][0] > W - tol]
    justified = len(rows) > 2 and len(full) >= 0.6 * (len(rows) - 1)
    if centred:
        align = PP_ALIGN.CENTER
    else:
        # justified text is set flush right (left for English): every original
        # line keeps its own break, and PowerPoint would stretch the spaces of
        # a justified line ending in a break into wide gaps between words
        align = PP_ALIGN.RIGHT if rtl else PP_ALIGN.LEFT

    # ---- split rows into paragraphs: a row that ends early ends its paragraph
    paras, cur = [], []
    for i, r in enumerate(rows):
        text = "".join(t for _, t in _row_text(r))
        if cur and BULLET.match(text):
            paras.append(cur)
            cur = []
        cur.append(r)
        rx0, rx1 = r["bbox"][0], r["bbox"][2]
        short = (rx0 > x0 + tol * 2) if rtl else (rx1 < x1 - tol * 2)
        if (centred or short) and i < len(rows) - 1:   # centred titles keep their line breaks
            paras.append(cur)
            cur = []
    if cur:
        paras.append(cur)

    # ---- geometry: the box covers the original lines (plus a little room),
    #      never spilling out of the frame the text sits in
    if len(rows) == 1:
        slack = W * 0.12 + 6
        if centred or align == PP_ALIGN.CENTER:
            left, right = x0 - slack / 2, x1 + slack / 2
        elif rtl:
            left, right = x0 - slack, x1
        else:
            left, right = x0, x1 + slack
    else:
        pad = W * 0.015 + 2
        left, right = x0 - pad, x1 + pad
    frame = box.get("box")
    if frame:
        left, right = max(left, frame[0] + 1), min(right, frame[2] - 1)
        if right - left < W:                       # keep at least the text's own width
            left, right = x0, x1
    left = max(0, left)
    width = right - left

    # ---- every original line keeps its own line, so the font is scaled until
    #      the widest line fits the box (no extra wrapping, nothing overflows)
    scale = 1.0
    room = width * 0.97
    for r in rows:
        m = text_width("".join(t for _, t in _row_text(r)), r["size"])
        if m and m > room:
            scale = min(scale, room / m)
    scale = max(0.6, scale)

    tb = slide.shapes.add_textbox(Emu(int(left * k)), Emu(int(y0 * k)),
                                  Emu(int(width * k)), Emu(int((y1 - y0) * k)))
    tf = tb.text_frame
    tf.word_wrap = len(rows) > 1      # a single line never wraps (e.g. "1 / 44")
    tf.auto_size = None
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.TOP if len(rows) > 1 else MSO_ANCHOR.MIDDLE

    # line pitch of the original (baseline to baseline)
    pitch = None
    if len(rows) > 1:
        d = sorted(b["base"] - a["base"] for a, b in zip(rows, rows[1:]))
        pitch = d[len(d) // 2]

    first = True
    for para in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if rtl:
            p._p.get_or_add_pPr().set("rtl", "1")
        if pitch and pitch > 0:
            p.line_spacing = Pt(round(pitch * k / 12700 * 2) / 2)
        for j, r in enumerate(para):
            parts = _row_text(r)
            if j > 0:
                # keep the original line break, so each line stays where it was
                p.add_line_break()
                br = p._p.findall(qn("a:br"))[-1]
                rpr = br.makeelement(qn("a:rPr"), {"sz": str(int(Pt(max(6, round(r["size"] * scale * k / 12700 * 2) / 2)).pt * 100))})
                if rtl:
                    rpr.set("lang", "ar-EG")
                br.append(rpr)
            if parts:
                s0, t0 = parts[0]
                parts[0] = (s0, t0.lstrip(" "))
                s1, t1 = parts[-1]
                parts[-1] = (s1, t1.rstrip(" "))
            clean, prev_space = [], False
            for s, t in parts:
                t = re.sub(r"\s{2,}", " ", t)
                if prev_space:
                    t = t.lstrip(" ")
                if t:
                    clean.append((s, t))
                    prev_space = t.endswith(" ")
            parts = clean
            for s, t in parts:
                if not t:
                    continue
                run = p.add_run()
                run.text = t
                _style_run(run, s, scale, k, rtl)


def convert(pdf_path, image_only=False):
    import pymupdf
    from pptx import Presentation
    from pptx.util import Emu

    import docx_flow
    data, expand, coverage, guesses = repaired_for_text(pdf_path)
    doc = pymupdf.open(stream=data, filetype="pdf")
    docx_flow.learn_font_names(doc)
    docx_flow._UNREADABLE.clear()
    docx_flow._UNREADABLE.update(n for n, c in coverage.items() if c < 0.5)
    first = doc[0].rect
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(int(12192000 * first.height / first.width))
    k = 12192000 / first.width
    blank = prs.slide_layouts[6]

    # first pass: text lines of every page, so letters the PDF lost and doubtful
    # spaces can be settled with the whole document's vocabulary
    page_rects, page_lines = [], []
    if not image_only:
        for pno, page in enumerate(doc):
            rects = shape_rects(page)
            try:
                rects += image_rects(doc, pno)       # framed boxes that are part of a picture
            except Exception:
                pass
            page_rects.append(rects)
            page_lines.append(engine_lines(page, rects, expand))
        docx_flow.resolve_unknowns(page_lines, guesses)
        docx_flow.resolve_doubtful_spaces(page_lines)

    for pno, page in enumerate(doc):
        slide = prs.slides.add_slide(blank)
        if image_only:
            img = page.get_pixmap(dpi=DPI).tobytes("png")
            slide.shapes.add_picture(io.BytesIO(img), 0, 0, prs.slide_width, prs.slide_height)
            continue
        rects = page_rects[pno]
        rows, keep = engine_rows(page_lines[pno])
        boxes = group_boxes(rows, rects)
        bg = background_png(doc, pno, keep)
        slide.shapes.add_picture(io.BytesIO(bg), 0, 0, prs.slide_width, prs.slide_height)
        for box in boxes:
            add_box(slide, box, k)

    base, _ = os.path.splitext(pdf_path)
    out = base + (" (images).pptx" if image_only else " (editable).pptx")
    prs.core_properties.author = "Beedo El-Omda"
    prs.core_properties.last_modified_by = "Beedo El-Omda"
    prs.save(out)
    return out, len(doc)


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
    lnk = os.path.join(sendto, "PDF to PPTX (Arabic, editable).lnk")
    script = os.path.abspath(__file__)

    def q(s):
        return "'" + s.replace("'", "''") + "'"

    icon = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "System32", "imageres.dll") + ",68"
    args = '"' + script + '"'
    ps = (
        f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut({q(lnk)});"
        f"$s.TargetPath={q(pythonw)};$s.Arguments={q(args)};"
        f"$s.WorkingDirectory={q(os.path.dirname(script))};$s.IconLocation={q(icon)};"
        "$s.Description='Convert PDF to editable PowerPoint with correct Arabic';$s.Save()"
    )
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps], check=True)


def main():
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    args = sys.argv[1:]
    if len(args) == 2 and args[0] == "--install-sendto":
        install_sendto(args[1])
        return
    image_only = "--image" in args
    files = [f for f in args if f.lower().endswith(".pdf")]
    if not files:
        notify("PDF to PPTX", "Select PDF files, then right-click > Send to > PDF to PPTX (Arabic, editable).")
        return
    done, failed = [], []
    for f in files:
        try:
            out, n = convert(f, image_only)
            done.append(f"{os.path.basename(out)}  ({n} slides)")
        except Exception:
            failed.append(f"{os.path.basename(f)}:\n{traceback.format_exc(limit=2)}")
    msg = ("Created:\n" + "\n".join(done) + "\n\nEnjoy , Beedo El-Omda") if done else ""
    if failed:
        msg += ("\n\n" if msg else "") + "Failed:\n" + "\n".join(failed)
    notify("PDF to PPTX (Arabic, editable)", msg)


if __name__ == "__main__":
    main()
