"""Repair Arabic text in PDFs so converters (PDFGear, Word, Acrobat) produce
editable, correct Arabic.

Mac-made PDFs often map Arabic glyphs to Unicode "presentation forms"
(U+FB50-U+FDFF, U+FE70-U+FEFF), including ligatures such as U+FC70 (تر) or
U+FCCC (لم). Windows fonts lack most of these, so converted text shows boxes.
This rewrites every font's ToUnicode map to normal Arabic letters (NFKC),
leaving the page appearance untouched.

Output: "<name> (fixed).pdf" next to the original.
Usage: python fix_arabic_pdf.py file1.pdf [file2.pdf ...]
"""
import os
import re
import sys
import unicodedata

import logging

import pikepdf

logging.getLogger("fontTools").setLevel(logging.ERROR)

# common bullets of Zapf Dingbats / Monotype Sorts / Wingdings (private-use codes)
DINGBATS = {"l": "\u25CF", "n": "\u25A0", "u": "\u25C6", "v": "\u2756", "q": "\u2751",
            "o": "\u2752", "p": "\u2752", "s": "\u25B2", "t": "\u25BC", "3": "\u2713",
            "4": "\u2714", "\u00a7": "\u25AA", "\u00d8": "\u27A2", "\u00fc": "\u2714"}

HEX = re.compile(r"<([0-9A-Fa-f]+)>")


def is_presentation(ch):
    o = ord(ch)
    return 0xFB50 <= o <= 0xFDFF or 0xFE70 <= o <= 0xFEFF


# Persian/Urdu codes that Arabic documents often carry by mistake (same look in Arabic text)
PERSIAN_TO_ARABIC = {"\u06CC": "\u064A", "\u06BE": "\u0647", "\u06A9": "\u0643", "\u06C0": "\u0629",
                     "\u06C1": "\u0647", "\u06D2": "\u064A"}


def norm(s):
    out = []
    for ch in s:
        if is_presentation(ch):
            d = unicodedata.normalize("NFKC", ch)
            out.append(d.strip() or ch)
        else:
            out.append(ch)
    return "".join(PERSIAN_TO_ARABIC.get(c, c) for c in "".join(out))


def h2s(h):
    return bytes.fromhex(h).decode("utf-16-be", errors="replace")


def s2h(s):
    return s.encode("utf-16-be").hex().upper()


def parse_cmap(text):
    """Return (codespace_lines, {code_hex: unicode_str}, code_width)."""
    mapping = {}
    cs = re.search(r"begincodespacerange(.*?)endcodespacerange", text, re.S)
    codespace = cs.group(1).strip() if cs else "<00> <FF>"
    for block in re.findall(r"beginbfchar(.*?)endbfchar", text, re.S):
        for src, dst in re.findall(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>", block):
            mapping[src.upper()] = h2s(dst)
    for block in re.findall(r"beginbfrange(.*?)endbfrange", text, re.S):
        for line in re.finditer(r"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(\[[^\]]*\]|<[0-9A-Fa-f]*>)", block):
            lo, hi, dst = line.group(1), line.group(2), line.group(3)
            width = len(lo)
            a, b = int(lo, 16), int(hi, 16)
            if dst.startswith("["):
                dsts = HEX.findall(dst)
                for i, c in enumerate(range(a, b + 1)):
                    if i < len(dsts):
                        mapping[format(c, f"0{width}X")] = h2s(dsts[i])
            else:
                base = bytearray(bytes.fromhex(dst[1:-1]))
                for i, c in enumerate(range(a, b + 1)):
                    cur = bytearray(base)
                    if cur:
                        cur[-1] = (cur[-1] + i) & 0xFF
                    mapping[format(c, f"0{width}X")] = bytes(cur).decode("utf-16-be", errors="replace")
    return codespace, mapping


def build_cmap(codespace, mapping):
    items = sorted(mapping.items())
    parts = [
        "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n"
        "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n"
        "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n"
        f"1 begincodespacerange\n{codespace}\nendcodespacerange\n"
    ]
    for i in range(0, len(items), 100):
        chunk = items[i:i + 100]
        parts.append(f"{len(chunk)} beginbfchar\n")
        parts += [f"<{k}> <{s2h(v)}>\n" for k, v in chunk]
        parts.append("endbfchar\n")
    parts.append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n")
    return "".join(parts).encode("latin-1")


GLYPH_UNI = re.compile(r"uni((?:[0-9A-Fa-f]{4})+)|u([0-9A-Fa-f]{4,6})")


def glyph_text(name):
    """'uni06440645062C.Init' -> 'لمج', 'uni0644_uni0627' -> 'لا', else None."""
    out = ""
    for part in name.split(".")[0].split("_"):
        m = GLYPH_UNI.fullmatch(part)
        if not m:
            return None
        if m.group(1):
            h = m.group(1)
            out += "".join(chr(int(h[i:i + 4], 16)) for i in range(0, len(h), 4))
        else:
            out += chr(int(m.group(2), 16))
    return out or None


def font_glyph_names(font):
    """{code_hex: glyph name} for a simple TrueType font, or {}."""
    try:
        import io
        from fontTools.ttLib import TTFont
        ff = font.FontDescriptor.get("/FontFile2")
        if ff is None:
            return {}
        t = TTFont(io.BytesIO(ff.read_bytes()), lazy=True)
        cm = t["cmap"].getcmap(1, 0) or t["cmap"].getcmap(3, 0) or t["cmap"].getcmap(3, 1)
        out = {}
        for code, name in (cm.cmap.items() if cm else []):
            out[format(code & 0xFF, "02X")] = name
        return out
    except Exception:
        return {}


ARABIC_CH = re.compile(r"[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]")


def glyph_texts(tt):
    """{glyph id: text} from an embedded TrueType font's own data.

    Uses the font's cmap, 'uniXXXX' glyph names, and the GSUB table to trace
    shaped / ligature / kashida variants (e.g. 'glyph00523') back to letters.
    """
    order = tt.getGlyphOrder()
    gid_of = {n: i for i, n in enumerate(order)}
    text = {}
    try:
        for cp, name in (tt.getBestCmap() or {}).items():
            g = gid_of.get(name)
            if g is not None and (g not in text or cp < 0xFB50):   # prefer base letters
                text[g] = chr(cp)
    except Exception:
        pass
    for g, name in enumerate(order):
        if g not in text:
            t = glyph_text(name)
            if t:
                text[g] = t
    if "GSUB" not in tt:
        return text
    rules = []                      # (target glyph, [source glyphs])
    try:
        for lookup in tt["GSUB"].table.LookupList.Lookup:
            for st in lookup.SubTable:
                if getattr(st, "LookupType", None) == 7 or hasattr(st, "ExtSubTable"):
                    st = st.ExtSubTable
                lt = getattr(st, "LookupType", None)
                if lt == 1:
                    rules += [(dst, [src]) for src, dst in st.mapping.items()]
                elif lt == 3:
                    for src, alts in st.alternates.items():
                        rules += [(a, [src]) for a in alts]
                elif lt == 4:
                    for first, ligs in st.ligatures.items():
                        for lig in ligs:
                            rules.append((lig.LigGlyph, [first] + list(lig.Component)))
    except Exception:
        return text
    for _ in range(4):              # propagate through chains of substitutions
        changed = False
        for dst, srcs in rules:
            g = gid_of.get(dst)
            if g is None or g in text:
                continue
            parts = [text.get(gid_of.get(s)) for s in srcs]
            if all(parts):
                text[g] = "".join(parts)
                changed = True
        if not changed:
            break
    return text


# ------------------------------------------------ installed-font fallback ---
_SYS_FONTS = None


def system_font(ps_name):
    """Find the installed font file (Windows) whose PostScript name matches, or None."""
    global _SYS_FONTS
    if _SYS_FONTS is None:
        _SYS_FONTS = {}
        dirs = [os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "Fonts"),
                os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")]
        try:
            from fontTools.ttLib import TTFont
        except Exception:
            return None
        for d in dirs:
            if not os.path.isdir(d):
                continue
            for fn in os.listdir(d):
                if not fn.lower().endswith((".ttf", ".otf")):
                    continue
                path = os.path.join(d, fn)
                try:
                    t = TTFont(path, lazy=True, fontNumber=0)
                    n = t["name"].getDebugName(6)
                    if n:
                        _SYS_FONTS.setdefault(n.replace(" ", "").lower(), path)
                    t.close()
                except Exception:
                    pass
    return _SYS_FONTS.get(ps_name.replace(" ", "").lower())


def outline_key(glyphset, name):
    """Hashable outline of a glyph (same design -> same key, whatever its glyph number)."""
    try:
        from fontTools.pens.recordingPen import DecomposingRecordingPen
        pen = DecomposingRecordingPen(glyphset)
        glyphset[name].draw(pen)
        if not pen.value:
            return None
        return hash(tuple((op, tuple(tuple(round(c) for c in pt) for pt in args if pt is not None))
                          for op, args in pen.value))
    except Exception:
        return None


_SYS_TEXTS = {}


def system_texts(path):
    """(texts by glyph id, texts by outline) for an installed font, cached."""
    if path not in _SYS_TEXTS:
        from fontTools.ttLib import TTFont
        full = TTFont(path, lazy=False)
        by_gid = glyph_texts(full)
        gs, order = full.getGlyphSet(), full.getGlyphOrder()
        by_outline = {}
        for g, t in by_gid.items():
            k = outline_key(gs, order[g])
            if k is not None:
                by_outline.setdefault(k, t)
        _SYS_TEXTS[path] = (len(order), by_gid, by_outline)
    return _SYS_TEXTS[path]


# ------------------------------------------------------- shape matching ---
def _raster(glyphset, name, n=40, stretch=True):
    try:
        import numpy as np
        from PIL import Image, ImageDraw
        from fontTools.pens.recordingPen import DecomposingRecordingPen
        pen = DecomposingRecordingPen(glyphset)
        glyphset[name].draw(pen)
    except Exception:
        return None
    polys, cur = [], []
    for op, args in pen.value:
        if op == "moveTo":
            cur = [args[0]]
        elif op in ("lineTo", "qCurveTo", "curveTo"):
            cur.extend(a for a in args if a is not None)
        elif op in ("closePath", "endPath"):
            if cur:
                polys.append(cur)
            cur = []
    pts = [p for poly in polys for p in poly]
    if not pts:
        return None
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    x0, y1 = min(xs), max(ys)
    sx, sy = (max(xs) - x0) or 1, (y1 - min(ys)) or 1
    if not stretch:
        sx = sy = max(sx, sy)
    # stretch=True: fill the box in both directions, so kashida-widened variants
    # of a letter still look like the letter
    im = Image.new("L", (n, n), 0)
    d = ImageDraw.Draw(im)
    for poly in polys:
        if len(poly) > 2:
            d.polygon([((p[0] - x0) / sx * (n - 1), (y1 - p[1]) / sy * (n - 1)) for p in poly], fill=255, outline=255)
    return np.array(im) > 0


def match_by_shape(tt, gids, gtext):
    """Guess letters for unnamed glyphs by comparing outlines with known letters.

    Returns (sure, guesses): sure = {gid: letter} for clear matches;
    guesses = {gid: [(letter, score), ...]} best first, for the rest.
    """
    import numpy as np
    gs, order = tt.getGlyphSet(), tt.getGlyphOrder()
    known = []
    for g, t in gtext.items():
        b = unicodedata.normalize("NFKC", t)
        if len(b) == 1 and 0x0621 <= ord(b) <= 0x064A:
            r = _raster(gs, order[g])
            if r is not None and r.sum():
                known.append((r, b))
    sure, guesses = {}, {}
    for g in gids:
        r = _raster(gs, order[g])
        if r is None or not r.sum():
            continue
        best = {}
        for k, t in known:
            sc = np.logical_and(r, k).sum() / np.logical_or(r, k).sum()
            if sc > best.get(t, -1):
                best[t] = float(sc)
        ranked = sorted(best.items(), key=lambda x: -x[1])
        if not ranked:
            continue
        if ranked[0][1] >= 0.6 and (len(ranked) < 2 or ranked[0][1] - ranked[1][1] >= 0.1):
            sure[g] = ranked[0][0]
        else:
            guesses[g] = ranked[:6]
    return sure, guesses


def is_kashida_shape(tt, name):
    """A thin horizontal stroke near the baseline = kashida (tatweel), not a letter."""
    try:
        upm = tt["head"].unitsPerEm
        if "glyf" in tt:
            g = tt["glyf"][name]
            g.recalcBounds(tt["glyf"])
            if not hasattr(g, "yMax") or g.numberOfContours == 0:
                return False
            w, h, ymin = g.xMax - g.xMin, g.yMax - g.yMin, g.yMin
        else:
            from fontTools.pens.boundsPen import BoundsPen
            pen = BoundsPen(tt.getGlyphSet())
            tt.getGlyphSet()[name].draw(pen)
            if not pen.bounds:
                return False
            x0, y0, x1, y1 = pen.bounds
            w, h, ymin = x1 - x0, y1 - y0, y0
        return h < 0.12 * upm and w > 1.5 * h and -0.15 * upm < ymin < 0.25 * upm
    except Exception:
        return False


def cid_font_texts(font, suspicious, all_codes=()):
    """{code_hex(4): text} for a Type0/Identity font with embedded TrueType, or {}.

    suspicious: codes whose PDF mapping looks wrong (blank, alef, ...); only for
    those are the slower fallbacks (installed font, shape matching) used.
    """
    try:
        import io
        from fontTools.ttLib import TTFont
        if font.get("/Subtype") != "/Type0" or str(font.get("/Encoding")) not in ("/Identity-H", "/Identity-V"):
            return {}, {}, None, {}
        df = font.DescendantFonts[0]
        ff = df.FontDescriptor.get("/FontFile2")
        if ff is None:
            return {}, {}, None, {}
        tt = TTFont(io.BytesIO(ff.read_bytes()), lazy=False)
        gtext = glyph_texts(tt)
        c2g = df.get("/CIDToGIDMap")
        if c2g is None or str(c2g) == "/Identity":
            cid2gid = None
        else:
            data = c2g.read_bytes()
            cid2gid = {format(c, "04X"): data[2 * c] << 8 | data[2 * c + 1] for c in range(len(data) // 2)}

        def gid(code):
            return int(code, 16) if cid2gid is None else cid2gid.get(code)

        order = tt.getGlyphOrder()
        # unnamed glyphs that are just a stroke: kashida, whatever the PDF claims
        for c in all_codes:
            g = gid(c)
            if g is not None and g not in gtext and g < len(order) and is_kashida_shape(tt, order[g]):
                gtext[g] = "\u0640"
                suspicious = suspicious | {c}
        missing = {gid(c) for c in suspicious if gid(c) is not None and gid(c) not in gtext
                   and gid(c) < len(order) and order[gid(c)] not in ("space", "uni00A0", ".notdef", "nbspace")}
        if missing:      # same font installed on this PC? (subsets keep the original glyph numbers)
            ps = str(font.get("/BaseFont", "")).lstrip("/").split("+")[-1]
            path = system_font(ps)
            if path:
                try:
                    n_full, by_gid, by_outline = system_texts(path)
                    gs = tt.getGlyphSet()
                    for g in list(missing):
                        k = outline_key(gs, order[g])
                        t = by_outline.get(k) if k is not None else None
                        if t is None and n_full == len(order):
                            t = by_gid.get(g)          # same version: glyph numbers match
                        if t:
                            gtext[g] = t
                            missing.discard(g)
                except Exception:
                    pass
        guesses = {}
        if missing:
            sure, guesses = match_by_shape(tt, missing, gtext)
            gtext.update(sure)
        out = {}
        for code in suspicious | set(format(g, "04X") for g in gtext if cid2gid is None):
            g = gid(code)
            if g is not None and g in gtext:
                out[code] = gtext[g]
        out_guess = {}
        for code in suspicious:
            g = gid(code)
            if g in guesses:
                out_guess[code] = guesses[g]
        return out, gtext, gid, out_guess
    except Exception:
        return {}, {}, None, {}


UNKNOWN_GUESSES = {}   # placeholder char -> [(letter, shape score), ...]
DOT_PLACEHOLDERS = set()   # placeholders of zero-width pieces (dots drawn separately)


def _next_pua(codes):
    used = set(codes.values())
    c = 0xEC00 + len(codes)
    while chr(c) in used:
        c += 1
    return chr(c)


def fix_font(font, seen, ligature_codes=None, report=None):
    """Rewrite one font's ToUnicode map in place. Returns number of codes changed.

    ligature_codes: optional dict. When given, every multi-letter glyph (e.g. لا,
    لمج) is mapped to ONE private-use character and recorded here, so text
    extraction can't reorder its letters; expand_ligatures() restores them.
    """
    tu = font.get("/ToUnicode")
    if tu is None or tu.objgen in seen:
        return 0
    seen.add(tu.objgen)
    try:
        codespace, mapping = parse_cmap(tu.read_bytes().decode("latin-1"))
    except Exception:
        return 0
    # Arabic fonts: trust the font's own glyph data where the PDF's map disagrees
    # (Word's PDF export maps kashida to "ا", ain/diacritics/ط/ك to spaces, etc.)
    ftext, fguess = {}, {}
    if any(ARABIC_CH.search(v) for v in mapping.values()):
        susp = {k.upper().zfill(4) for k, v in mapping.items()
                if not v.strip() or v in ("\u0627", "\ufffd") or len(v) > 1 or is_presentation(v[:1] or " ")}
        ftext, gtext, gid, fguess = cid_font_texts(font, susp, {k.upper().zfill(4) for k in mapping})
        if report is not None and gid is not None:
            # how many of the glyphs this PDF uses can be traced to letters?
            visible = [k for k, v in mapping.items() if v.strip()]
            known = sum(1 for k in visible if gid(k.upper().zfill(4)) in gtext)
            report[str(font.get("/BaseFont", "")).lstrip("/").split("+")[-1]] = known / max(1, len(visible))
    names = None
    new = {}
    for k, v in mapping.items():
        ft = ftext.get(k.upper().zfill(4))
        if ft and norm(ft) != norm(v) and (ARABIC_CH.search(ft) or (ft.isspace() and ARABIC_CH.search(v))):
            v = ft
        nv = norm(v)
        if any(0xE000 <= ord(c) <= 0xF8FF or c == "\ufffd" for c in nv):
            if names is None:
                names = font_glyph_names(font)
            g = glyph_text(names.get(k, "")) if names.get(k) else None
            if g:
                nv = norm(g)
            elif len(nv) == 1 and 0xF020 <= ord(nv) <= 0xF0FF:
                nv = DINGBATS.get(chr(ord(nv) - 0xF000), "\u25AA")
        kk = k.upper().zfill(4)
        if ligature_codes is not None and kk in fguess and not nv.strip():
            # unknown letter: keep a placeholder; the document's own words decide later
            key = "\x00?" + str(font.get("/BaseFont", "")) + ":" + kk
            if key not in ligature_codes:
                ligature_codes[key] = _next_pua(ligature_codes)
                UNKNOWN_GUESSES[ligature_codes[key]] = fguess[kk]
            new[k] = ligature_codes[key]
            continue
        if ligature_codes is not None and len(nv) > 1 and ARABIC_CH.search(nv) and nv.strip() == nv:
            if nv not in ligature_codes:
                ligature_codes[nv] = _next_pua(ligature_codes)
            nv = ligature_codes[nv]
        new[k] = nv
    n = sum(1 for k in mapping if new[k] != mapping[k])
    if n:
        tu.write(build_cmap(codespace, new))
    return n


def expand_ligatures(text, ligature_codes):
    """Undo the private-use placeholders made by fix_font(..., ligature_codes)."""
    if not ligature_codes:
        return text
    back = {v: k for k, v in ligature_codes.items() if not k.startswith("\x00")}
    return "".join(back.get(c, c) for c in text)


# ------------------------------------------- codes missing from ToUnicode ---
def _font_program(font):
    """(TTFont, code -> glyph name function, is_simple) for fonts with an embedded
    TrueType program, else None."""
    import io
    from fontTools.ttLib import TTFont
    st = font.get("/Subtype")
    if st == "/Type0":
        if str(font.get("/Encoding")) not in ("/Identity-H", "/Identity-V"):
            return None
        df = font.DescendantFonts[0]
        ff = df.FontDescriptor.get("/FontFile2")
        if ff is None:
            return None
        tt = TTFont(io.BytesIO(ff.read_bytes()), lazy=False)
        order = tt.getGlyphOrder()
        c2g = df.get("/CIDToGIDMap")
        cid2gid = None
        if c2g is not None and str(c2g) != "/Identity":
            data = c2g.read_bytes()
            cid2gid = [data[2 * c] << 8 | data[2 * c + 1] for c in range(len(data) // 2)]

        def name(code):
            g = code if cid2gid is None else (cid2gid[code] if code < len(cid2gid) else None)
            return order[g] if g is not None and 0 < g < len(order) else None
        return tt, name, False
    if st == "/TrueType":
        fd = font.get("/FontDescriptor")
        ff = fd.get("/FontFile2") if fd is not None else None
        if ff is None:
            return None
        tt = TTFont(io.BytesIO(ff.read_bytes()), lazy=False)
        tables = {(t.platformID, t.platEncID): t.cmap for t in tt["cmap"].tables} if "cmap" in tt else {}

        def name(code):
            for key, off in (((3, 0), 0xF000), ((1, 0), 0), ((3, 1), 0), ((3, 0), 0)):
                m = tables.get(key)
                if m and code + off in m:
                    return m[code + off]
            return None
        return tt, name, True
    return None


def _good_text(v):
    return bool(v.strip()) and not any(0xE000 <= ord(c) <= 0xF8FF or c == "\ufffd" for c in v)


_REF = {}


def _ref_raster(ch, n=40):
    """Raster of a Latin character in a standard font (Helvetica), for checking
    whether a glyph really is the character its encoding claims."""
    if ch not in _REF:
        r = None
        try:
            import numpy as np
            import pymupdf
            from PIL import Image
            d = pymupdf.open()
            p = d.new_page(width=80, height=80)
            p.insert_text((15, 60), ch, fontsize=48, fontname="helv")
            pix = p.get_pixmap(alpha=False)
            a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, 0] < 128
            ys, xs = np.nonzero(a)
            if len(xs):
                c = a[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
                r = np.array(Image.fromarray(c.astype(np.uint8) * 255).resize((n, n))) > 127
        except Exception:
            r = None
        _REF[ch] = r
    return _REF[ch]


def _looks_like(gs, gname, ch):
    r, R = _raster(gs, gname), _ref_raster(ch)
    if r is None or R is None or not r.sum():
        return False
    return (r & R).sum() / (r | R).sum() >= 0.5


def _map_bare_font(pdf, obj, lib, arabic_names, ligature_codes):
    """A simple TrueType font without any ToUnicode map, from a family that
    carries Arabic elsewhere in the document (macOS exports split an Arabic
    font into a mapped part and an unmapped MacRoman part). Build its map."""
    if str(obj.get("/BaseFont", "")).split("+")[-1] not in arabic_names:
        return 0
    prog = _font_program(obj)
    if prog is None:
        return 0
    tt, name, _ = prog
    gs = tt.getGlyphSet()
    order = tt.getGlyphOrder()
    first = int(obj.get("/FirstChar", 0))
    used = [first + i for i in range(len(obj.get("/Widths", [])))]
    enc = "mac_roman" if str(obj.get("/Encoding", "")) == "/MacRomanEncoding" else "cp1252"
    mapping, unresolved = {}, []
    for c in used:
        n = name(c)
        if not n or n == ".notdef":
            continue
        if not re.fullmatch(r"glyph\d+|g\d+|gid\d+", n):
            # a real glyph name (zero, comma, uni0628...) says what it is
            t = glyph_text(n)
            if not t:
                try:
                    t = bytes([c]).decode(enc)
                except Exception:
                    t = ""
            if t:
                mapping[format(c, "02X")] = t
                continue
        key = outline_key(gs, n)
        cand = lib.get(key) if key is not None else None
        if cand:
            mapping[format(c, "02X")] = max(cand.items(), key=lambda x: x[1])[0]
            continue
        try:
            ch = bytes([c]).decode(enc)
        except Exception:
            ch = ""
        if not ch.strip() or n in ("space", ".notdef") or _looks_like(gs, n, ch):
            mapping[format(c, "02X")] = ch or " "
        else:
            unresolved.append((c, n))
    if not any(ARABIC_CH.search(v) for v in mapping.values()) and not unresolved:
        return 0
    if unresolved:
        gidx = {n: i for i, n in enumerate(order)}
        gtext = {gidx[name(int(k, 16))]: v for k, v in mapping.items()
                 if name(int(k, 16)) in gidx and ARABIC_CH.search(v)}
        sure, guesses = {}, {}
        try:
            sure, guesses = match_by_shape(tt, [gidx[n] for _, n in unresolved if n in gidx], gtext)
        except Exception:
            pass
        zero = {first + i for i, w in enumerate(obj.get("/Widths", [])) if float(w) == 0}
        for c, n in unresolved:
            code = format(c, "02X")
            g = gidx.get(n)
            if c in zero and ligature_codes is None:
                mapping[code] = "\u200b"
            elif g in sure and c not in zero:
                mapping[code] = sure[g]
            elif ligature_codes is not None:
                ok = outline_key(gs, n)    # same glyph in another subset -> same placeholder
                key = "\x00?" + (("o%d" % ok) if ok is not None else str(obj.get("/BaseFont", "")) + ":" + code)
                if key not in ligature_codes:
                    ligature_codes[key] = _next_pua(ligature_codes)
                    UNKNOWN_GUESSES[ligature_codes[key]] = guesses.get(g, [])
                if c in zero:
                    DOT_PLACEHOLDERS.add(ligature_codes[key])
                mapping[code] = ligature_codes[key]
            else:
                mapping[code] = "\u200b"          # unknown: better nothing than junk
    obj.ToUnicode = pdf.make_stream(build_cmap("<00> <FF>", mapping))
    return len(mapping)


def fill_missing_codes(pdf, ligature_codes=None):
    """Some PDFs (e.g. macOS / Keynote exports) leave glyphs out of a font's
    ToUnicode map, so text extraction shows Latin junk (%, N, Z, W ...) instead
    of Arabic letters. The same glyph outline is usually mapped in another
    subset of the font elsewhere in the document; failing that, its shape is
    compared with the letters this font does map. Missing codes are added."""
    fonts, bare = [], []
    lib = {}
    done = set()
    for obj in pdf.objects:
        if not isinstance(obj, pikepdf.Dictionary) or obj.get("/Type") != "/Font":
            continue
        tu = obj.get("/ToUnicode")
        if tu is None:
            if obj.get("/Subtype") == "/TrueType" and obj.objgen not in done:
                done.add(obj.objgen)
                bare.append(obj)
            continue
        if tu.objgen in done:
            continue
        done.add(tu.objgen)
        try:
            prog = _font_program(obj)
            if prog is None:
                continue
            codespace, mapping = parse_cmap(tu.read_bytes().decode("latin-1"))
        except Exception:
            continue
        tt, name, simple = prog
        gs = tt.getGlyphSet()
        for k, v in mapping.items():
            n = name(int(k, 16))
            if n and _good_text(v):
                key = outline_key(gs, n)
                if key is not None:
                    lib.setdefault(key, {}).setdefault(norm(v), 0)
                    lib[key][norm(v)] += 1
        fonts.append((obj, tu, prog, codespace, mapping))
    added = 0
    arabic_names = {str(o.get("/BaseFont", "")).split("+")[-1] for o, _, _, _, m in fonts
                    if any(ARABIC_CH.search(v) for v in m.values())}
    for obj in bare:
        try:
            added += _map_bare_font(pdf, obj, lib, arabic_names, ligature_codes)
        except Exception:
            pass
    for obj, tu, (tt, name, simple), codespace, mapping in fonts:
        if not any(ARABIC_CH.search(v) for v in mapping.values()):
            continue
        if simple:
            first = int(obj.get("/FirstChar", 0))
            widths = list(obj.get("/Widths", []))
            used = [first + i for i in range(len(widths))]     # zero-width parts too
            width = 2
        else:
            used = []
            width = 4
        have = {int(k, 16) for k in mapping}
        missing = [c for c in used if c not in have and name(c) not in (None, ".notdef")]
        if not missing:
            continue
        gs = tt.getGlyphSet()
        order = tt.getGlyphOrder()
        new = dict(mapping)
        unresolved = []
        for c in missing:
            n = name(c)
            key = outline_key(gs, n)
            cand = lib.get(key) if key is not None else None
            if cand:
                new[format(c, "0%dX" % width)] = max(cand.items(), key=lambda x: x[1])[0]
            elif n not in (".notdef", "space"):
                unresolved.append((c, n))
        if unresolved:
            sure, guesses, gidx = {}, {}, {}
            try:
                gidx = {n: i for i, n in enumerate(order)}
                gtext = {}
                for k, v in mapping.items():
                    n = name(int(k, 16))
                    if n in gidx and _good_text(v):
                        gtext[gidx[n]] = v
                sure, guesses = match_by_shape(tt, [gidx[n] for _, n in unresolved if n in gidx], gtext)
            except Exception:
                pass
            zero = {first + i for i, w in enumerate(widths) if float(w) == 0} if simple else set()
            for c, n in unresolved:
                code = format(c, "0%dX" % width)
                g = gidx.get(n)
                if c in zero and ligature_codes is None:
                    new[code] = "\u200b"           # dots/marks drawn as separate zero-width pieces
                elif g in sure and c not in zero:
                    new[code] = sure[g]
                elif ligature_codes is not None:
                    # unknown glyph (letter or ligature): a placeholder that the
                    # document's own words resolve later
                    ok = outline_key(gs, n)
                    key = "\x00?" + (("o%d" % ok) if ok is not None else str(obj.get("/BaseFont", "")) + ":" + code)
                    if key not in ligature_codes:
                        ligature_codes[key] = _next_pua(ligature_codes)
                        UNKNOWN_GUESSES[ligature_codes[key]] = guesses.get(g, [])
                    if c in zero:
                        DOT_PLACEHOLDERS.add(ligature_codes[key])
                    new[code] = ligature_codes[key]
        if len(new) > len(mapping):
            added += len(new) - len(mapping)
            tu.write(build_cmap(codespace, new))
    return added


def fix_pdf(path):
    pdf = pikepdf.open(path)
    fill_missing_codes(pdf)
    fixed_fonts = changed_codes = 0
    seen = set()
    for obj in pdf.objects:
        if not isinstance(obj, pikepdf.Dictionary) or obj.get("/Type") != "/Font":
            continue
        n = fix_font(obj, seen)
        if n:
            fixed_fonts += 1
            changed_codes += n
    base, _ = os.path.splitext(path)
    out = base + " (fixed).pdf"
    pdf.save(out)
    return out, fixed_fonts, changed_codes


if __name__ == "__main__":
    for p in sys.argv[1:]:
        out, f, c = fix_pdf(p)
        print(f"{os.path.basename(out)}: fixed {c} codes in {f} fonts")
