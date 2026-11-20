"""Rebuild an Arabic (or mixed) PDF as a flowing, editable Word document.

Built for document-style PDFs (books, reports, theses): real paragraphs with
right-to-left direction, the original fonts, sizes, bold and colours, footnote
numbers as superscript, images in place, tables as Word tables, one PDF page
per Word page.

Text order is rebuilt from glyph positions (not from the PDF's text order),
which fixes punctuation, brackets, numbers and English words inside Arabic.
"""
import io
import os
import re
import unicodedata

ARABIC = re.compile(r"[؀-ۿݐ-ݿﭐ-﷿\uFE70-\uFEFF\uEC00-\uEFFF]")
LTR_STRONG = re.compile(r"[A-Za-z0-9À-ɏ٠-٩۰-۹]")
MIRROR = {"(": ")", ")": "(", "[": "]", "]": "[", "{": "}", "}": "{", "<": ">", ">": "<",
          "«": "»", "»": "«"}
TATWEEL = "ـ"
DOUBT = "\u2063"          # marker for a space that may be an artefact inside a word
BULLETS = "-•●▪■◆❖➢✓✔–"


# ------------------------------------------------------------- fonts ---
FONT_ALIASES = {
    "timesnewroman": "Times New Roman", "arial": "Arial", "calibri": "Calibri",
    "simplifiedarabic": "Simplified Arabic", "traditionalarabic": "Traditional Arabic",
    "arabictypesetting": "Arabic Typesetting", "sakkalmajalla": "Sakkal Majalla",
    "andalus": "Andalus", "tahoma": "Tahoma", "segoeui": "Segoe UI",
}


_EMBEDDED_FAMILY = {}     # PDF font name -> family name read from the embedded font
_UNREADABLE = set()       # fonts whose text can't be trusted (no letter table) -> keep as picture


def learn_font_names(doc):
    """Read real family names ('SKR HEAD1', 'Simplified Arabic') from embedded fonts."""
    import logging
    logging.getLogger("fontTools").setLevel(logging.ERROR)
    try:
        from fontTools.ttLib import TTFont
    except Exception:
        return
    seen = set()
    for pno in range(doc.page_count):
        for xref, ext, _t, basefont, _n, _e in doc.get_page_fonts(pno):
            if xref in seen or basefont in _EMBEDDED_FAMILY:
                continue
            seen.add(xref)
            try:
                _name, ext, _st, buf = doc.extract_font(xref)
                if not buf or ext not in ("ttf", "otf", "cff", "ttc"):
                    continue
                t = TTFont(io.BytesIO(buf), lazy=True)
                fam = t["name"].getDebugName(16) or t["name"].getDebugName(1)
                if fam:
                    _EMBEDDED_FAMILY[basefont.split("+")[-1]] = fam.strip()
            except Exception:
                continue


def font_family(pdf_font):
    """'BCDEEE+SimplifiedArabic-Bold' -> ('Simplified Arabic', bold, italic)."""
    name = pdf_font.split("+")[-1]
    if name in _EMBEDDED_FAMILY:
        low = name.lower()
        return (_EMBEDDED_FAMILY[name], any(k in low for k in ("bold", "black", "heavy")),
                "italic" in low or "oblique" in low)
    low = name.lower()
    bold = any(k in low for k in ("bold", "black", "heavy", "semibold", "demi"))
    italic = "italic" in low or "oblique" in low
    base = re.split(r"[-,]", name)[0]
    base = re.sub(r"(PSMT|PS|MT)$", "", base)
    key = base.replace(" ", "").lower()
    if key in FONT_ALIASES:
        return FONT_ALIASES[key], bold, italic
    nice = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", base).strip() or "Arial"
    return nice, bold, italic


# ------------------------------------------------------ glyph clusters ---
def page_glyphs(page, expand):
    """All glyphs on the page as clusters (base char + its combining marks)."""
    import pymupdf
    raw = None
    for flags in (pymupdf.TEXT_PRESERVE_LIGATURES | pymupdf.TEXT_PRESERVE_WHITESPACE,
                  pymupdf.TEXTFLAGS_RAWDICT, 0):
        try:
            raw = page.get_text("rawdict", flags=flags)
            break
        except Exception:
            continue
    if raw is None:
        return []
    try:
        from fix_arabic_pdf import DOT_PLACEHOLDERS as dots
    except Exception:
        dots = set()
    glyphs = []
    seq = 0
    for b in raw["blocks"]:
        if b.get("type") != 0:
            continue
        for l in b["lines"]:
            if abs(l["dir"][1]) > 0.05:          # skip rotated text
                continue
            for s in l["spans"]:
                seq += 1
                fam, bold, italic = font_family(s["font"])
                bold = bold or bool(s["flags"] & 16)
                italic = italic or bool(s["flags"] & 2)
                style = (fam, round(s["size"] * 2) / 2, bold, italic, s["color"])
                bad = s["font"].split("+")[-1] in _UNREADABLE
                for c in s["chars"]:
                    ch = expand(c["c"]).replace("\u200b", "")
                    if not ch:
                        continue
                    x0, y0, x1, y1 = c["bbox"]
                    g = {"t": ch, "x": c["origin"][0], "y": c["origin"][1], "x0": x0, "x1": x1,
                         "y0": y0, "y1": y1, "size": s["size"], "style": style, "bad": bad, "seq": seq}
                    g["mark"] = bool(ch.strip()) and all(unicodedata.combining(k) for k in ch)
                    g["dot"] = bool(dots) and all(k in dots for k in ch)
                    glyphs.append(g)
    if dots:
        glyphs = _attach_dots(glyphs)
    return glyphs


def _attach_dots(glyphs):
    """Dots some PDFs draw as separate zero-width pieces join the letter they sit
    on (found by position), so the pair can be read as one letter later."""
    bases = [g for g in glyphs if not g.get("dot")]
    for d in glyphs:
        if not d.get("dot"):
            continue
        best, score = None, None
        for b in bases:
            if not b["t"].strip() or abs(b["y"] - d["y"]) > 1.2 * b["size"]:
                continue
            dist = 0 if b["x0"] - 0.5 <= d["x"] <= b["x1"] + 0.5 else min(abs(d["x"] - b["x0"]), abs(d["x"] - b["x1"]))
            if dist > 0.3 * b["size"]:
                continue
            sc = (dist, abs(b["y"] - d["y"]))
            if score is None or sc < score:
                best, score = b, sc
        if best is not None:
            best["t"] += d["t"]
    return bases


# ------------------------------------------------------------- lines ---
def _container(g, rects):
    cx, cy = (g["x0"] + g["x1"]) / 2, g["y"] - 0.3 * g["size"]
    best = None
    for r in rects:
        if r[0] <= cx <= r[2] and r[1] <= cy <= r[3]:
            a = (r[2] - r[0]) * (r[3] - r[1])
            if best is None or a < best[0]:
                best = (a, r)
    return best[1] if best else None


def _rtl_key(g):
    """Right-to-left visual order. A space whose origin is on its right edge was
    drawn from the end of the previous letter: it comes after that letter."""
    x1 = g["x1"]
    if not g["t"].strip() and g["x"] > (g["x0"] + x1) / 2:
        x1 -= 0.3
    return (-x1, -g["x"])


def _place_spaces(gs):
    """Some PDFs store a space glyph stretched back over the letter before it, so
    sorting by right edge puts it one letter too early. A space that clearly lies
    over the gap between two letters is moved into that gap."""
    solid = [g for g in gs if g["t"].strip()]
    if len(solid) < 2 or len(solid) == len(gs):
        return gs
    after = {}
    for g in gs:
        if g["t"].strip():
            continue
        w = g["x1"] - g["x0"]
        best, score = None, max(0.2 * w, 0.2 * g["size"])
        for i in range(len(solid) - 1):
            a, b = solid[i], solid[i + 1]          # a is right of b
            ov = min(g["x1"], a["x0"]) - max(g["x0"], b["x1"])
            if ov > score:
                best, score = i, ov
        if best is not None:
            after[id(g)] = id(solid[best])
    if not after:
        return gs
    moved = set(after)
    out = []
    for g in gs:
        if id(g) in moved:
            continue
        out.append(g)
        out.extend(s for s in gs if after.get(id(s)) == id(g))
    return out


def _drop_hidden_copy(gs):
    """Some slides hold the same title twice at slightly different sizes, one
    copy covered by the other. Keep only the copy drawn last (the visible one)."""
    groups = {}
    for g in gs:
        groups.setdefault(round(g["size"], 1), []).append(g)
    if len(groups) < 2:
        return gs
    import difflib
    big = sorted((v for v in groups.values() if len(v) >= 4), key=lambda v: -max(g.get("seq", 0) for g in v))
    drop = set()

    def text(v):        # Arabic letters only, right to left
        return re.sub("[^\u0621-\u064A]", "", "".join(g["t"] for g in sorted(v, key=lambda g: -g["x1"])))
    for i, a in enumerate(big):
        if id(a[0]) in drop:
            continue
        ta = text(a)
        ax0, ax1 = min(g["x0"] for g in a), max(g["x1"] for g in a)
        for b in big[i + 1:]:
            tb = text(b)
            bx0, bx1 = min(g["x0"] for g in b), max(g["x1"] for g in b)
            ov = min(ax1, bx1) - max(ax0, bx0)
            short = min(len(ta), len(tb))
            if short < 4 or ov < 0.6 * min(ax1 - ax0, bx1 - bx0):
                continue
            same = sum(m.size for m in difflib.SequenceMatcher(None, ta, tb, autojunk=False).get_matching_blocks())
            if same >= 0.8 * short and short >= 0.6 * max(len(ta), len(tb)):
                drop.update(id(g) for g in b)
    return [g for g in gs if id(g) not in drop] if drop else gs


def build_lines(glyphs, rects=(), split_gap=4.0):
    """Group glyphs into visual lines by baseline, split at big horizontal gaps
    and where a line crosses from one drawn box / column / cell into another."""
    glyphs = sorted(glyphs, key=lambda g: g["y"])
    rows = []
    for g in glyphs:
        for r in reversed(rows[-4:]):
            if abs(r["y"] - g["y"]) <= max(1.5, 0.35 * min(r["size"], g["size"])) or \
               (g["size"] < 0.75 * r["size"] and r["y"] - 0.6 * r["size"] <= g["y"] <= r["y"] + 1):
                r["g"].append(g)
                if g["size"] >= r["size"]:
                    r["size"] = g["size"]
                    r["y"] = g["y"]
                break
        else:
            rows.append({"y": g["y"], "size": g["size"], "g": [g]})
    # small raised text (footnote numbers) belongs to the line it sits on
    rows.sort(key=lambda r: -r["size"])
    merged = []
    for r in rows:
        host = None
        if r["size"] < 0.8 * max((m["size"] for m in merged), default=0):
            rx0 = min(g["x0"] for g in r["g"]); rx1 = max(g["x1"] for g in r["g"])
            for m in merged:
                if m["size"] * 0.8 <= r["size"]:
                    continue
                mx0 = min(g["x0"] for g in m["g"]); mx1 = max(g["x1"] for g in m["g"])
                if (m["y"] - 0.75 * m["size"] <= r["y"] <= m["y"] + 0.15 * m["size"]
                        and rx1 >= mx0 - m["size"] and rx0 <= mx1 + m["size"] and len(r["g"]) <= 8):
                    host = m
                    break
        if host:
            host["g"].extend(r["g"])
        else:
            merged.append(r)
    rows = merged
    for r in rows:
        r["g"] = _drop_hidden_copy(r["g"])
    lines = []
    for r in rows:
        gs = _place_spaces(sorted(r["g"], key=_rtl_key))
        if rects:
            for g in gs:
                g["box"] = _container(g, rects)
        seg = [gs[0]]
        last_solid = gs[0] if gs[0]["t"].strip() else None
        for g in gs[1:]:
            gap = seg[-1]["x0"] - g["x1"]
            other_box = (bool(rects) and g["t"].strip() and last_solid is not None
                         and g.get("box") != last_solid.get("box"))
            if g["t"].strip():
                last_solid_next = g
            else:
                last_solid_next = last_solid
            size_jump = abs(g["size"] - seg[-1]["size"]) > 0.2 * max(g["size"], seg[-1]["size"]) and gap > 0.5 * r["size"]
            if gap > split_gap * r["size"] or other_box or size_jump:   # separate column / cell / shape
                lines.append(make_line(seg, r))
                seg = []
            seg.append(g)
            last_solid = last_solid_next
        lines.append(make_line(seg, r))
    lines = [l for l in lines if l["text"].strip()]
    lines.sort(key=lambda l: (round(l["y"], 0), -l["x1"]))
    return lines


def drop_inner_spaces(gs):
    """Remove 'space' glyphs that sit on top of letters (Word's justification artefacts)."""
    solid = [(g["x0"], g["x1"]) for g in gs if g["t"].strip()]
    out = []
    for g in gs:
        if not g["t"].strip():
            # the stray ones lie across several letters; a normal space in slanted
            # (italic) text only overlaps its two neighbours
            covered = sum(1 for a, b in solid
                          if min(g["x1"], b) - max(g["x0"], a) > 0.3 * max(0.01, b - a))
            if covered >= 3:
                continue
            w = max(0.01, g["x1"] - g["x0"])
            # how much of the space is under letters (union of overlaps)
            spans = sorted((max(g["x0"], a), min(g["x1"], b)) for a, b in solid if min(g["x1"], b) > max(g["x0"], a))
            under, end = 0.0, g["x0"]
            for a, b in spans:
                if b > end:
                    under += b - max(a, end)
                    end = b
            if covered >= 1 and under > 0.9 * w:   # possibly a stray space inside a word: the vocabulary decides
                g = dict(g, t=DOUBT)
        out.append(g)
    return out


def attach_marks(gs):
    """Put each haraka (fatha, damma, tanween, shadda...) on the letter it sits over.

    PDFs store marks as separate glyphs, often out of order; matching them by
    position keeps them on the right letter and stops stray spaces before them.
    """
    bases = [g for g in gs if not g.get("mark")]
    marks = [g for g in gs if g.get("mark")]
    if not marks or not bases:
        return gs
    bases = [dict(g) for g in bases]
    letters = [g for g in bases if g["t"].strip()] or bases
    for m in marks:
        cx = (m["x0"] + m["x1"]) / 2
        best, score = None, None
        for b in letters:
            ov = min(m["x1"], b["x1"]) - max(m["x0"], b["x0"])
            dist = 0 if b["x0"] <= cx <= b["x1"] else min(abs(cx - b["x0"]), abs(cx - b["x1"]))
            sc = (dist, -ov)
            if score is None or sc < score:
                best, score = b, sc
        best["t"] += m["t"]
    return bases


def make_line(gs, row):
    """Visual (right->left sorted) glyphs -> logical text runs."""
    gs = attach_marks(gs)
    gs = drop_inner_spaces(gs) or gs
    arabic = sum(1 for g in gs if ARABIC.search(g["t"]))
    latin = sum(1 for g in gs if re.search(r"[A-Za-z]", g["t"]))
    rtl = arabic > 0 and arabic >= 0.25 * latin   # Arabic with some English still reads right-to-left
    if not rtl:
        gs = sorted(gs, key=lambda g: g["x"])
    seq = list(gs)
    # Some PDFs store brackets as drawn (mirrored), others in reading order.
    # Reading right-to-left, the first bracket of a line is normally an opening one.
    mirror = False
    if rtl:
        for g in seq:
            if g["t"] in "()[]{}":
                mirror = g["t"] in ")]}"
                break
    if rtl:
        # numbers / English inside Arabic are drawn left-to-right: flip those runs back
        out, i = [], 0
        while i < len(seq):
            if LTR_STRONG.search(seq[i]["t"]):
                j = i
                while j + 1 < len(seq) and (LTR_STRONG.search(seq[j + 1]["t"]) or
                                            (seq[j + 1]["t"] in " .,:/%-_'&+" and j + 2 < len(seq)
                                             and LTR_STRONG.search(seq[j + 2]["t"]))):
                    j += 1
                out.extend(reversed(seq[i:j + 1]))
                i = j + 1
            else:
                g = seq[i]
                if mirror and g["t"] in MIRROR:
                    g = dict(g, t=MIRROR[g["t"]])
                out.append(g)
                i += 1
        seq = out
    # spaces from gaps the PDF didn't store as characters
    for k, g in enumerate(seq):           # e.g. tanween+alef shapes stored mark-first ("ًا" -> "اً")
        t = g["t"]
        j = 0
        while j < len(t) and unicodedata.combining(t[j]):
            j += 1
        if 0 < j < len(t):
            seq[k] = dict(g, t=t[j:] + t[:j])
    pieces = []
    prev = None
    for g in seq:
        if prev is not None:
            gap = (prev["x0"] - g["x1"]) if rtl else (g["x0"] - prev["x1"])
            both_ar = ARABIC.search(g["t"]) and ARABIC.search(prev["t"])
            need = (0.2 if both_ar else 0.22) * max(g["size"], prev["size"])
            inside_brackets = prev["t"][-1:] in "([{" or g["t"][:1] in ")]}"
            starts_mark = bool(g["t"]) and unicodedata.combining(g["t"][0]) > 0
            if gap > need and not inside_brackets and not starts_mark and not g["t"].startswith((" ", DOUBT)) and not prev["t"].endswith((" ", DOUBT)):
                if not (rtl and LTR_STRONG.search(g["t"]) and LTR_STRONG.search(prev["t"]) and gap < 0.4 * g["size"]):
                    pieces.append((" ", prev["style"]))
        pieces.append((g["t"].replace(TATWEEL, "").replace("\u00a0", " "), g["style"]))
        prev = g
    pieces = [(re.sub(r" {2,}", " ", t), st) for t, st in pieces]
    runs = []
    for t, st in pieces:
        if not t:
            continue
        if runs and runs[-1][1] == st:
            runs[-1][0] += t
        else:
            runs.append([t, st])
    for i in range(1, len(runs)):
        if runs[i - 1][0].endswith(" ") and runs[i][0].startswith(" "):
            runs[i][0] = runs[i][0].lstrip(" ")
    runs = [r for r in runs if r[0]]
    text = "".join(r[0] for r in runs)
    main = max((g["size"] for g in gs), default=row["size"])
    bad = sum(1 for g in gs if g.get("bad") and g["t"].strip()) > 0.5 * max(1, sum(1 for g in gs if g["t"].strip()))
    ink = [g for g in gs if g["t"].strip()] or gs     # trailing spaces don't widen the line
    return {"runs": runs, "text": re.sub(r"\s+", " ", text).strip(), "rtl": rtl, "bad": bad,
            "x0": min(g["x0"] for g in ink), "x1": max(g["x1"] for g in ink),
            "y": row["y"], "top": min(g["y0"] for g in gs), "bottom": max(g["y1"] for g in gs),
            "size": main}


# -------------------------------------------------------- paragraphs ---
def build_paragraphs(lines, left, right):
    """Join lines into paragraphs. left/right = text area of the page."""
    width = right - left
    paras = []
    for ln in lines:
        p = paras[-1] if paras else None
        new = True
        if p is not None:
            last = p["lines"][-1]
            pitch = ln["y"] - last["y"]
            same_size = abs(ln["size"] - last["size"]) <= 1.0
            starts_bullet = ln["text"][:1] in BULLETS or re.match(r"^(\d+|[٠-٩]+)\s*[-.)]", ln["text"])
            last_full = (last["x0"] - left) < 0.08 * width if last["rtl"] else (right - last["x1"]) < 0.08 * width
            if (same_size and ln["rtl"] == last["rtl"] and 0 < pitch <= 1.9 * last["size"]
                    and last_full and not starts_bullet and not p.get("centered")):
                new = False
        if new:
            paras.append({"lines": [ln]})
        else:
            p["lines"].append(ln)
    for p in paras:
        ls = p["lines"]
        p["rtl"] = ls[0]["rtl"]
        p["x0"] = min(l["x0"] for l in ls)
        p["x1"] = max(l["x1"] for l in ls)
        p["size"] = max(l["size"] for l in ls)
        mid = (left + right) / 2
        c = (p["x0"] + p["x1"]) / 2
        w = p["x1"] - p["x0"]
        if len(ls) == 1 and abs(c - mid) < 0.06 * width and w < 0.8 * width:
            p["align"] = "center"
        elif len(ls) > 1:
            p["align"] = "both"
        else:
            p["align"] = "right" if p["rtl"] else "left"
        # indent of the first line / whole paragraph relative to the text area
        p["ind_start"] = max(0.0, (right - p["x1"]) if p["rtl"] else (p["x0"] - left))
        if p["align"] == "center":
            p["ind_start"] = 0.0
        p["pitch"] = (ls[-1]["y"] - ls[0]["y"]) / (len(ls) - 1) if len(ls) > 1 else None
    return paras


# -------------------------------------------------------------- docx ---
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(t):
    return "{%s}%s" % (W, t)


def add_run(par, text, style, line_size, superscript=False):
    from docx.shared import Pt, RGBColor
    fam, size, bold, italic, color = style
    r = par.add_run(text)
    f = r.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = fam
    f.color.rgb = RGBColor((color >> 16) & 255, (color >> 8) & 255, color & 255)
    if superscript:
        f.superscript = True
        f.size = Pt(max(size * 1.4, line_size * 0.75) if size < line_size else size)
    rpr = r._r.get_or_add_rPr()
    fonts = rpr.find(_q("rFonts"))
    for a in ("ascii", "hAnsi", "cs"):
        fonts.set(_q(a), fam)
    for tag, on in (("bCs", bold), ("iCs", italic)):
        if on:
            rpr.append(rpr.makeelement(_q(tag), {}))
    szcs = rpr.makeelement(_q("szCs"), {_q("val"): str(int(round(f.size.pt * 2)))})
    rpr.append(szcs)
    if ARABIC.search(text):
        rpr.append(rpr.makeelement(_q("rtl"), {}))
        rpr.append(rpr.makeelement(_q("lang"), {_q("bidi"): "ar-EG"}))
    from pdf2docx_arabic import sort_children, RPR_ORDER
    sort_children(rpr, RPR_ORDER)
    return r


def write_paragraph(doc_or_cell, p, expand_sup=True):
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    par = doc_or_cell.add_paragraph()
    pf = par.paragraph_format
    from docx.enum.text import WD_LINE_SPACING
    pf.space_before = Pt(p.get("space_before", 0))
    pf.space_after = Pt(p.get("space_after", 0))
    pf.widow_control = False
    if p.get("pitch"):
        pf.line_spacing = Pt(round(p["pitch"] * 0.97, 1))
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    align = {"center": WD_ALIGN_PARAGRAPH.CENTER, "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
             "right": WD_ALIGN_PARAGRAPH.RIGHT, "left": WD_ALIGN_PARAGRAPH.LEFT}[p["align"]]
    if p["rtl"]:
        ppr = par._p.get_or_add_pPr()
        ppr.append(ppr.makeelement(_q("bidi"), {}))
        # in an RTL paragraph Word's left/right are swapped (start/end)
        if align == WD_ALIGN_PARAGRAPH.RIGHT:
            align = WD_ALIGN_PARAGRAPH.LEFT
        elif align == WD_ALIGN_PARAGRAPH.LEFT:
            align = WD_ALIGN_PARAGRAPH.RIGHT
    pf.alignment = align
    if p.get("ind_start", 0) > 3:
        # w:start = the side the text starts on (right for Arabic), read the same by Word and LibreOffice
        ppr = par._p.get_or_add_pPr()
        ind = ppr.makeelement(_q("ind"), {_q("start"): str(int(p["ind_start"] * 20)), _q("end"): "0"})
        ppr.append(ind)
    line_size = p["size"]
    for li, ln in enumerate(p["lines"]):
        if li:
            par.add_run(" ")
        runs = [[t, st] for t, st in ln["runs"]]
        # one space at most between words; no padding at line ends (justification spaces)
        for r in runs:
            r[0] = re.sub(r"\s{2,}", " ", r[0])
        while runs and not runs[0][0].strip():
            runs.pop(0)
        while runs and not runs[-1][0].strip():
            runs.pop()
        if runs:
            runs[0][0] = runs[0][0].lstrip()
            runs[-1][0] = runs[-1][0].rstrip()
        for a, b in zip(runs, runs[1:]):
            if a[0].endswith(" ") and b[0].startswith(" "):
                b[0] = b[0].lstrip()
        for text, st in runs:
            if not text:
                continue
            sup = st[1] < 0.72 * line_size and re.fullmatch(r"\s*[\(\)\[\]\d٠-٩*]+\s*", text) is not None
            add_run(par, text, st, line_size, superscript=sup)
    from pdf2docx_arabic import sort_children, PPR_ORDER
    sort_children(par._p.get_or_add_pPr(), PPR_ORDER)
    return par


HARAKAT = re.compile(r"[\u064B-\u065F\u0670]")
LETTERS = "ءآأؤإئابةتثجحخدذرزسشصضطظعغفقكلمنهوىي"


def resolve_unknowns(page_lines, guesses):
    """Fill letters the PDF lost, using the document's own vocabulary.

    guesses: {placeholder: [(letter, shape score), ...]}. An unknown token is a
    run of placeholders inside a word (some PDFs draw one letter as several
    pieces, or several letters as one). Every word holding unknown tokens is
    matched against the words that are readable elsewhere in the document; each
    match votes for what its tokens stand for. The best supported reading of
    each token is taken, put back into the other words, and the rest voted on
    again. Tokens that can't be decided are dropped (or set from their shape).
    """
    phs = "".join(guesses)
    if not phs:
        return
    tok_re = re.compile("[%s]+" % re.escape(phs))
    letters = "\u0621-\u064A"
    word_re = re.compile("[%s\u0671-\u06D3%s\u064B-\u065F\u0670]+" % (letters, re.escape(phs)))
    vocab = {}
    words = {}
    for ls in page_lines:
        for l in ls:
            for w in word_re.findall(l["text"]):
                w = HARAKAT.sub("", w)
                if tok_re.search(w):
                    words[w] = words.get(w, 0) + 1
                elif len(w) > 1:
                    vocab[w] = vocab.get(w, 0) + 1
    if not words:
        return
    try:
        from arabic_common_words import WORDS
        for i, v in enumerate(WORDS):
            # the list is roughly in order of frequency: early words count more
            vocab[v] = max(vocab.get(v, 0), 5000.0 / (i + 10))
    except Exception:
        pass
    by_len = {}
    for v, n in vocab.items():
        by_len.setdefault(len(v), []).append((v, n))
    choice = {}

    def vote_round():
        votes = {}
        for w, wn in words.items():
            parts = tok_re.split(w)
            toks = tok_re.findall(w)
            unknown = [t for t in toks if t not in choice]
            if not unknown:
                continue
            for pre in ("", "و", "ف", "ب", "ل", "ك"):
                if pre and not parts[0].startswith(pre):
                    continue
                ps = list(parts)
                ps[0] = ps[0][len(pre):]
                rx, fixed = "", sum(len(p) for p in ps)
                for i, t in enumerate(toks):
                    rx += re.escape(ps[i])
                    if t in choice:
                        rx += re.escape(choice[t])
                        fixed += len(choice[t])
                    else:
                        rx += "([%s]{1,3})" % letters
                rx += re.escape(ps[-1])
                cre = re.compile(rx + "$")
                nunk = sum(1 for t in toks if t not in choice)
                hits = []
                for L in range(fixed + nunk, fixed + 3 * nunk + 1):
                    for v, n in by_len.get(L, ()):
                        m = cre.match(v)
                        if m:
                            hits.append((m.groups(), n))
                if not hits:
                    continue
                # every readable word this one could be gets an equal share of its vote
                # a reading scores one point for every different word it makes
                # readable (plus a small share, so frequent words break ties)
                readings = {}
                for groups, n in hits:
                    key = tuple(g.replace("\u0649", "\u064A") for g in groups)
                    readings[key] = readings.get(key, 0) + n
                total = sum(readings.values())
                seen_vals = set()
                for groups, n in readings.items():
                    share = 0.5 * n / total           # frequent readings win ties
                    gi = iter(groups)
                    for t in toks:
                        if t not in choice:
                            val = next(gi)
                            d = votes.setdefault(t, {})
                            d[val] = d.get(val, 0.0) + share
                            if (t, val) not in seen_vals:
                                seen_vals.add((t, val))
                                d[val] += 1.0
                break
        return votes

    for _ in range(10):
        votes = vote_round()
        progress = False
        # settle the clearest tokens first; the rest are voted on again with them in
        ranked = []
        for t, vs in votes.items():
            best = sorted(vs.items(), key=lambda x: (-round(x[1], 6), len(x[0])))
            top = best[0][1]
            second = best[1][1] if len(best) > 1 else 0.0
            if top >= 2.0 and top >= second + 1.0:
                ranked.append((top - second, t, best[0][0]))
        ranked.sort(reverse=True)
        for _, t, val in ranked[:max(1, len(ranked) // 2)]:
            choice[t] = val
            progress = True
        if not progress:
            break

    last = vote_round()
    for t, vs in last.items():
        if t not in choice:
            ranked = sorted(vs.items(), key=lambda x: (-round(x[1], 6), len(x[0])))
            top = ranked[0][1]
            second = ranked[1][1] if len(ranked) > 1 else 0.0
            # not clear-cut, but better than a missing letter -- except for a whole
            # word that could be any common word (then a wrong word is worse)
            alone = all(re.fullmatch("[وفبلك]?" + re.escape(t), w) for w in words if t in w)
            if top >= 1.0 and (top - second >= 0.002 or not alone):
                choice[t] = ranked[0][0]
    if os.environ.get("ARPDF_DEBUG"):
        for t, vs in last.items():
            print("VOTES", repr(t), sorted(vs.items(), key=lambda x: -x[1])[:4])

    def single(t):
        prior = dict(guesses.get(t, [])) if len(t) == 1 else {}
        return max(prior, key=prior.get) if prior and max(prior.values()) >= 0.5 else ""

    if os.environ.get("ARPDF_DEBUG"):
        alltok = {}
        for w, n in words.items():
            for t in tok_re.findall(w):
                alltok[t] = alltok.get(t, 0) + n
        for t, n in sorted(alltok.items(), key=lambda x: -x[1])[:40]:
            print(repr(t), n, repr(choice.get(t)), [tok_re.sub("?", w) for w in words if t in w][:5])

    def fix_word(m):
        w = m.group()
        if not tok_re.search(w):
            return w
        out = tok_re.sub(lambda t: choice[t.group()] if t.group() in choice else single(t.group()), w)
        if out.endswith("\u064A") and vocab.get(out[:-1] + "\u0649", 0) > vocab.get(out, 0):
            out = out[:-1] + "\u0649"
        return out

    def fix(text):
        return word_re.sub(fix_word, text)
    for ls in page_lines:
        for l in ls:
            if not tok_re.search(l["text"]):
                continue
            for r in l["runs"]:
                r[0] = fix(r[0])
            l["text"] = fix(l["text"])
    return choice


def resolve_doubtful_spaces(page_lines):
    """Join word pieces split by a doubtful space when the joined word occurs
    elsewhere in the document (e.g. 'إدا رة' -> 'إدارة'); otherwise keep a space."""
    letters = "\u0621-\u064A\u064B-\u0652\u0670"
    vocab = {}
    for ls in page_lines:
        for l in ls:
            for w in re.findall("[%s%s]+" % (letters, DOUBT), l["text"]):
                if DOUBT not in w:
                    w = HARAKAT.sub("", w)
                    vocab[w] = vocab.get(w, 0) + 1
    for ls in page_lines:
        for l in ls:
            if DOUBT not in l["text"]:
                continue
            text = l["text"]
            decisions = []
            for i, ch in enumerate(text):
                if ch != DOUBT:
                    continue
                a = HARAKAT.sub("", re.search("([%s]*)$" % letters, text[:i]).group(1))
                b = HARAKAT.sub("", re.match("([%s]*)" % letters, text[i + 1:]).group(1))
                if i + 1 < len(text) and unicodedata.combining(text[i + 1]):
                    decisions.append("")          # never a space before a haraka
                    continue
                whole = vocab.get(a + b, 0)
                # join when the whole word is at least half as common as its rarer piece
                decisions.append("" if a and b and whole >= 1 and
                                 whole >= 0.5 * min(vocab.get(a, 0), vocab.get(b, 0)) else " ")
            k = iter(decisions)
            l["text"] = re.sub(r"\s+", " ", "".join(next(k, " ") if ch == DOUBT else ch for ch in text)).strip()
            k = iter(decisions)
            for r in l["runs"]:
                if DOUBT in r[0]:
                    r[0] = "".join(next(k, " ") if ch == DOUBT else ch for ch in r[0])


def is_page_number(ln, page_h):
    return (ln["y"] > page_h * 0.9 or ln["y"] < page_h * 0.06) and \
        re.fullmatch(r"[\s\-–—|]*[\d٠-٩]{1,4}[\s\-–—|]*", ln["text"]) is not None


def convert(pdf_bytes, out_path, expand, font_coverage=None, unknown_guesses=None):
    """Build the flowing .docx. Returns number of pages.

    font_coverage: {font name: share of its used glyphs traceable to letters}
    from fix_font(); fonts below 50% are kept as pictures (text unreliable).
    """
    _UNREADABLE.clear()
    for name, cov in (font_coverage or {}).items():
        if cov < 0.5:
            _UNREADABLE.add(name)
    import pymupdf
    import docx
    from docx.shared import Pt
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_BREAK

    src = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    learn_font_names(src)
    d = docx.Document()
    # empty template paragraph
    body = d.element.body
    for p in list(body.iter(_q("p"))):
        body.remove(p)
    st = d.styles["Normal"]
    st.font.size = Pt(12)
    rpr = st.element.get_or_add_rPr()
    fonts = rpr.find(_q("rFonts"))
    if fonts is None:
        fonts = rpr.makeelement(_q("rFonts"), {})
        rpr.insert(0, fonts)

    first = src[0].rect
    sec = d.sections[0]
    sec.page_width, sec.page_height = Pt(first.width), Pt(first.height)
    if first.width > first.height:
        sec.orientation = WD_ORIENT.LANDSCAPE

    # text area: from all pages' text extents (robust against odd pages)
    lefts, rights, tops, bottoms = [], [], [], []
    page_lines, page_tables = [], []
    for page in src:
        tables = []
        try:
            found = page.find_tables().tables
            vlines = vertical_strokes(page) if found else []
            for t in found:
                if t.row_count >= 2 and t.col_count >= 2 and is_real_table(t, vlines):
                    tables.append(t)
        except Exception:
            pass
        page_tables.append(tables)
        cell_rects = [tuple(c) for t in tables for c in t.cells if c]
        lines = [l for l in build_lines(page_glyphs(page, expand), cell_rects)
                 if not is_page_number(l, page.rect.height)]
        page_lines.append(lines)
        for l in lines:
            if len(l["text"]) > 40:
                lefts.append(l["x0"]); rights.append(l["x1"])
            tops.append(l["top"]); bottoms.append(l["bottom"])

    resolve_unknowns(page_lines, unknown_guesses or {})
    resolve_doubtful_spaces(page_lines)

    def pct(v, q, dflt):
        v = sorted(v)
        return v[int(q * (len(v) - 1))] if v else dflt

    left = pct(lefts, 0.1, 56.0)
    right = pct(rights, 0.9, first.width - 56.0)
    sec.left_margin = Pt(max(18, left - 4))
    sec.right_margin = Pt(max(18, first.width - right - 4))
    sec.top_margin = Pt(max(18, min(pct(tops, 0.02, 56), 90) - 6))
    # text area ends low (where page numbers were) so dense pages have room to spare
    sec.bottom_margin = Pt(max(30, min(first.height - pct(bottoms, 0.98, first.height - 56) - 6, 42)))
    sec.footer_distance = Pt(14)
    add_page_number(sec)
    if src.page_count and any(ARABIC.search(l["text"]) for ls in page_lines for l in ls):
        sp = sec._sectPr
        bidi = sp.makeelement(_q("bidi"), {})
        grid = sp.find(_q("docGrid"))
        if grid is not None:
            grid.addprevious(bidi)        # schema order: ... bidi, rtlGutter, docGrid
        else:
            sp.append(bidi)

    usable = (first.width - left - (first.width - right))
    top_y = sec.top_margin.pt
    ratios = []
    for ls in page_lines:
        for p in build_paragraphs(ls, left, right):
            if p.get("pitch") and p["size"]:
                ratios.append(p["pitch"] / p["size"])
    ratio = sorted(ratios)[len(ratios) // 2] if ratios else 1.5
    for pno, page in enumerate(src):
        items = []           # (y, kind, payload)
        lines = page_lines[pno]

        # images: placed behind the text at their exact page position
        images = []
        try:
            for info in page.get_image_info(xrefs=True):
                bb = pymupdf.Rect(info["bbox"]) & page.rect
                if bb.width < 12 or bb.height < 12:
                    continue
                images.append((info, bb))
        except Exception:
            pass

        # tables -> Word tables; their text is taken out of the flow
        tables = page_tables[pno]
        flow = []
        for ln in [l for l in lines if l.get("bad")]:
            pad = 0.25 * ln["size"]
            bb = pymupdf.Rect(ln["x0"] - pad, ln["top"] - pad, ln["x1"] + pad, ln["bottom"] + pad) & page.rect
            images.append(({"xref": 0}, bb))
        lines = [l for l in lines if not l.get("bad")]
        for ln in lines:
            cy = (ln["top"] + ln["bottom"]) / 2
            cx = (ln["x0"] + ln["x1"]) / 2
            if any(t.bbox[0] - 2 <= cx <= t.bbox[2] + 2 and t.bbox[1] - 2 <= cy <= t.bbox[3] + 2 for t in tables):
                continue
            flow.append(ln)
        for t in tables:
            items.append((t.bbox[1], "table", t))

        for p in build_paragraphs(flow, left, right):
            items.append((p["lines"][0]["top"], "para", p))
        items.sort(key=lambda it: it[0])

        last_par = None
        first_par = None
        prev_bottom = top_y          # where the previous item ended on this page
        for _, kind, obj in items:
            if kind == "para":
                pitch = obj.get("pitch") or obj["size"] * ratio
                obj["pitch"] = pitch
                box_top = obj["lines"][0]["y"] - 0.78 * pitch
                box_bottom = obj["lines"][-1]["y"] + 0.22 * pitch
                gap = max(0.0, (box_top - prev_bottom) * 0.95)
                if first_par is None and gap > 6:       # space-before is dropped after a page break
                    sp = tiny_paragraph(d)
                    sp.paragraph_format.line_spacing = Pt(gap)
                    first_par = sp
                    gap = 0.0
                obj["space_before"] = gap
                prev_bottom = max(prev_bottom, box_bottom)
                last_par = write_paragraph(d, obj)
                first_par = first_par or last_par
            elif kind == "table":
                add_table(d, page, obj, lines, usable, ratio)
                prev_bottom = max(prev_bottom, obj.bbox[3])
                last_par = None
        if images:
            if first_par is None:
                first_par = last_par = tiny_paragraph(d)
            for info, bb in images:
                add_floating_image(first_par, src, page, info, bb)
        if pno < src.page_count - 1:
            if last_par is None:
                last_par = tiny_paragraph(d)
            last_par.add_run().add_break(WD_BREAK.PAGE)
    zoom = d.settings.element.find(_q("zoom"))
    if zoom is not None and zoom.get(_q("percent")) is None:
        zoom.set(_q("percent"), "100")
    d.core_properties.author = "Beedo El-Omda"
    d.core_properties.last_modified_by = "Beedo El-Omda"
    d.save(out_path)
    return src.page_count


def add_page_number(sec):
    """Centered page number field in the footer."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    par = sec.footer.paragraphs[0]
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run()
    run.font.size = Pt(11)
    for kind, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            el = run._r.makeelement(_q("fldChar"), {_q("fldCharType"): kind})
        else:
            el = run._r.makeelement(_q("instrText"), {})
            el.text = text
        run._r.append(el)


def vertical_strokes(page):
    """x positions and spans of drawn vertical lines / thin rectangles on the page."""
    out = []
    try:
        for dr in page.get_drawings():
            for it in dr.get("items", []):
                if it[0] == "l":
                    a, b = it[1], it[2]
                    if abs(a.x - b.x) < 1.5 and abs(a.y - b.y) > 6:
                        out.append((a.x, min(a.y, b.y), max(a.y, b.y)))
                elif it[0] == "re":
                    r = it[1]
                    if r.width < 2.5 and r.height > 6:
                        out.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
                    elif r.width > 6 and r.height > 6 and dr.get("color") is not None:   # outlined box
                        out += [(r.x0, r.y0, r.y1), (r.x1, r.y0, r.y1)]
    except Exception:
        pass
    return out


def is_real_table(t, vlines):
    """Underlines and highlights also form 'grids'; a real table has drawn column borders
    and text in several columns."""
    x0, y0, x1, y1 = t.bbox
    h = y1 - y0
    # Word draws borders as many short pieces (one per row): add up the pieces
    # at each column position (within ~3 pt) before judging the coverage.
    cover = {}
    for x, a, b in vlines:
        if x0 - 3 <= x <= x1 + 3:
            seg = min(b, y1) - max(a, y0)
            if seg > 0:
                key = next((k for k in cover if abs(k - x) <= 3), x)
                cover[key] = cover.get(key, 0.0) + seg
    xs = {k for k, v in cover.items() if v > 0.5 * h}
    if len(xs) < 3:                       # need outer borders + at least one inner column line
        return False
    try:
        rows = t.extract()
        multi = sum(1 for r in rows if sum(1 for c in r if c and str(c).strip()) >= 2)
        return multi >= 2
    except Exception:
        return True


def tiny_paragraph(d):
    from docx.shared import Pt
    from docx.enum.text import WD_LINE_SPACING
    par = d.add_paragraph()
    pf = par.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing = Pt(1)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    return par


def image_png(doc, page, info, bb):
    """The picture itself (with transparency), without any text drawn over it."""
    import pymupdf
    xref = info.get("xref") or 0
    if xref > 0:
        try:
            pix = pymupdf.Pixmap(doc, xref)
            smask = doc.xref_get_key(xref, "SMask")
            if smask[0] == "xref":
                mask = pymupdf.Pixmap(doc, int(smask[1].split()[0]))
                pix = pymupdf.Pixmap(pix, mask)
            if pix.n - pix.alpha >= 4:          # CMYK -> RGB
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            if pix.width * pix.height > 12_000_000:
                pix.shrink(1)
            return pix.tobytes("png")
        except Exception:
            pass
    try:
        return page.get_pixmap(clip=bb, dpi=150).tobytes("png")
    except Exception:
        return None


_PIC_ID = [1000]


def add_floating_image(par, doc, page, info, bb):
    """Picture anchored to the page at the PDF position, behind the text."""
    from docx.shared import Pt
    from lxml import etree
    data = image_png(doc, page, info, bb)
    if not data:
        return
    run = par.add_run()
    run.add_picture(io.BytesIO(data), width=Pt(bb.width), height=Pt(bb.height))
    inline = run._r.find(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline")
    if inline is None:
        return
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    emu = lambda v: str(int(round(v * 12700)))
    _PIC_ID[0] += 1
    anchor = etree.SubElement(inline.getparent(), "{%s}anchor" % WP, {
        "distT": "0", "distB": "0", "distL": "0", "distR": "0", "simplePos": "0",
        "relativeHeight": str(_PIC_ID[0]), "behindDoc": "1", "locked": "0",
        "layoutInCell": "1", "allowOverlap": "1"})
    etree.SubElement(anchor, "{%s}simplePos" % WP, {"x": "0", "y": "0"})
    ph = etree.SubElement(anchor, "{%s}positionH" % WP, {"relativeFrom": "page"})
    etree.SubElement(ph, "{%s}posOffset" % WP).text = emu(bb.x0)
    pv = etree.SubElement(anchor, "{%s}positionV" % WP, {"relativeFrom": "page"})
    etree.SubElement(pv, "{%s}posOffset" % WP).text = emu(bb.y0)
    ext = inline.find("{%s}extent" % WP)
    anchor.append(ext)
    etree.SubElement(anchor, "{%s}effectExtent" % WP, {"l": "0", "t": "0", "r": "0", "b": "0"})
    etree.SubElement(anchor, "{%s}wrapNone" % WP)
    for tag in ("docPr", "cNvGraphicFramePr"):
        el = inline.find("{%s}%s" % (WP, tag))
        if el is not None:
            if tag == "docPr":
                el.set("id", str(_PIC_ID[0]))
            anchor.append(el)
    graphic = inline.find("{http://schemas.openxmlformats.org/drawingml/2006/main}graphic")
    anchor.append(graphic)
    inline.getparent().remove(inline)


def add_image(d, page, info, bb, usable, space_before=0.0):
    import pymupdf
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    try:
        pix = page.get_pixmap(clip=bb, dpi=150)
        data = io.BytesIO(pix.tobytes("png"))
    except Exception:
        return None
    par = d.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(space_before)
    par.paragraph_format.space_after = Pt(0)
    w = min(bb.width, usable)
    h = bb.height * (w / bb.width)
    par.add_run().add_picture(data, width=Pt(w), height=Pt(h))
    return par


def add_table(d, page, t, lines, usable, ratio=1.5):
    """Word table from a detected PDF table, with the PDF's column widths and
    row heights; cell text from our ordered lines."""
    from docx.shared import Pt
    from docx.enum.table import WD_ROW_HEIGHT_RULE
    rows = t.rows
    ncols = t.col_count
    rtl = any(l["rtl"] for l in lines)
    table = d.add_table(rows=len(rows), cols=ncols)
    table.style = "Table Grid"
    table.autofit = False
    tblPr = table._tbl.tblPr
    if rtl:
        bv = tblPr.makeelement(_q("bidiVisual"), {})
        anchor = None
        for tag in ("tblOverlap", "tblpPr", "tblStyle"):      # schema order: tblStyle, tblpPr, tblOverlap, bidiVisual
            anchor = tblPr.find(_q(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addnext(bv)
        else:
            tblPr.insert(0, bv)
    # fixed layout + small cell margins, like the PDF
    mar = tblPr.makeelement(_q("tblCellMar"), {})
    for side in ("top", "start", "bottom", "end"):
        mar.append(mar.makeelement(_q(side), {_q("w"): "0" if side in ("top", "bottom") else "60", _q("type"): "dxa"}))
    tblPr.append(mar)
    from pdf2docx_arabic import sort_children
    sort_children(tblPr, ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
                          "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
                          "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"])

    # column edges from the PDF cells
    edges = []
    for c in range(ncols):
        xs0 = [r.cells[c][0] for r in rows if c < len(r.cells) and r.cells[c]]
        xs1 = [r.cells[c][2] for r in rows if c < len(r.cells) and r.cells[c]]
        edges.append((min(xs0) if xs0 else None, max(xs1) if xs1 else None))
    total = (t.bbox[2] - t.bbox[0]) or 1
    widths = []
    for a, b in edges:
        widths.append((b - a) if a is not None and b is not None else total / ncols)
    scale = min(1.0, usable / sum(widths)) if sum(widths) else 1.0
    widths = [w * scale for w in widths]
    order_cols = list(range(ncols))[::-1] if rtl else list(range(ncols))   # bidiVisual: first Word column = rightmost

    for ri, row in enumerate(rows):
        cells = list(row.cells)
        try:
            h = row.bbox[3] - row.bbox[1]
            tr = table.rows[ri]
            tr.height = Pt(max(8, h * 0.95))
            tr.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        except Exception:
            pass
        for ci, src_i in enumerate(order_cols):
            cell = table.cell(ri, ci)
            cell.width = Pt(widths[src_i])
            cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)
            bbox = cells[src_i] if src_i < len(cells) else None
            if bbox is None:
                tiny_paragraph(cell)
                continue
            inside = [l for l in lines if bbox[0] - 2 <= (l["x0"] + l["x1"]) / 2 <= bbox[2] + 2
                      and bbox[1] - 2 <= (l["top"] + l["bottom"]) / 2 <= bbox[3] + 2]
            if not inside:
                tiny_paragraph(cell)
                continue
            for p in build_paragraphs(inside, bbox[0], bbox[2]):
                p["ind_start"] = 0
                p["space_after"] = 0
                p["space_before"] = 0
                p["pitch"] = p.get("pitch") or p["size"] * ratio * 0.9
                if p["align"] == "both":          # cells: no justification (it spreads digits apart)
                    p["align"] = "right" if p["rtl"] else "left"
                write_paragraph(cell, p)
    for ci, w in enumerate(order_cols):
        table.columns[ci].width = Pt(widths[w])
    tiny_paragraph(d)
