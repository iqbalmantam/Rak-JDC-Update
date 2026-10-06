"""Baca kotak teks nama customer di atas layout rak (drawing di dalam file .xlsx)."""
import colorsys
import re
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
_ln = lambda e: e.tag.split("}")[-1]
EMU_PX = 9525
ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]


def _theme(z):
    try:
        root = ET.fromstring(z.read("xl/theme/theme1.xml"))
    except KeyError:
        return {}
    out = {}
    cs = root.find(".//a:clrScheme", NS)
    for ch in cs:
        name = _ln(ch)
        c = ch[0]
        out[name] = c.get("val") if _ln(c) == "srgbClr" else c.get("lastClr")
    return out


def _color(fill, theme):
    """Warna kotak seperti tampil di Excel (termasuk lumMod/lumOff dan transparansi di atas putih)."""
    c = fill[0]
    kind = _ln(c)
    hexv = c.get("val") if kind == "srgbClr" else theme.get(c.get("val"), "808080")
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    alpha = 1.0
    for m in c:
        t = _ln(m)
        v = int(m.get("val")) / 100000
        if t == "lumMod":
            l *= v
        elif t == "lumOff":
            l += v
        elif t == "alpha":
            alpha = v
    l = min(1, max(0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    r, g, b = (alpha * x + (1 - alpha) for x in (r, g, b))
    return "#%02x%02x%02x" % tuple(round(x * 255) for x in (r, g, b))


def _hex_hls(h):
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (1, 3, 5))
    return colorsys.rgb_to_hls(r, g, b)


def _hls_hex(h, l, s):
    r, g, b = colorsys.hls_to_rgb(h, min(1, max(0, l)), min(1, max(0, s)))
    return "#%02x%02x%02x" % tuple(round(x * 255) for x in (r, g, b))


PUCAT = ["#e0457b", "#c9a227", "#e59ad8", "#2aa89a"]  # pengganti warna kotak yang terlalu pucat


def _sheet_drawing(z, sheet_index):
    """Path drawing untuk sheet ke-N (1-based), jatuh ke drawing pertama bila tidak ketemu."""
    rel = f"xl/worksheets/_rels/sheet{sheet_index}.xml.rels"
    try:
        root = ET.fromstring(z.read(rel))
        for r in root:
            if r.get("Type", "").endswith("/drawing"):
                t = r.get("Target")
                return "xl/drawings/" + t.split("/")[-1]
    except KeyError:
        pass
    names = [n for n in z.namelist() if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n)]
    return sorted(names)[0] if names else None


def baca_customer(file, ws, sheet_index=1):
    pucat = list(PUCAT)
    """Return (customers, boxes): customers=[{name,color}], boxes dengan nama & area dalam EMU."""
    if hasattr(file, "seek"):
        file.seek(0)
    z = zipfile.ZipFile(file)
    dpath = _sheet_drawing(z, sheet_index)
    if not dpath:
        return [], []
    theme = _theme(z)
    root = ET.fromstring(z.read(dpath))

    wcol, hrow = {}, {}
    for d in ws.column_dimensions.values():
        if d.width:
            for c in range(d.min or 1, (d.max or d.min or 1) + 1):
                wcol[c] = d.width
    dw = ws.sheet_format.defaultColWidth or 8.43
    dh = ws.sheet_format.defaultRowHeight or 15
    colw = lambda c: round((wcol.get(c, dw) * 7 + 5)) * EMU_PX
    rowh = lambda r: (ws.row_dimensions[r].height or dh) * 12700
    cx, cy = [0], [0]
    for c in range(1, 200):
        cx.append(cx[-1] + colw(c))
    for r in range(1, 400):
        cy.append(cy[-1] + rowh(r))

    def pos(e):
        col = int(e.find("xdr:col", NS).text)
        row = int(e.find("xdr:row", NS).text)
        return (cx[col] + int(e.find("xdr:colOff", NS).text),
                cy[row] + int(e.find("xdr:rowOff", NS).text))

    boxes = []
    for anc in root:
        sp = anc.find("xdr:sp", NS)
        fr, to = anc.find("xdr:from", NS), anc.find("xdr:to", NS)
        if sp is None or fr is None or to is None:
            continue
        txt = " ".join("".join(t.text or "" for t in p.iter(f"{{{NS['a']}}}t")) for p in sp.iter(f"{{{NS['a']}}}p")).strip()
        fill = sp.find("xdr:spPr/a:solidFill", NS)
        if fill is None:
            continue
        (x0, y0), (x1, y1) = pos(fr), pos(to)
        boxes.append(dict(name=" ".join(txt.split()), color=_color(fill, theme), x0=x0, y0=y0, x1=x1, y1=y1))

    # kotak tanpa nama memakai nama kotak bernama yang warnanya sama
    named = {}
    for b in boxes:
        if b["name"]:
            named.setdefault(b["color"], b["name"])
    for b in boxes:
        b["inferred"] = False
        if not b["name"]:
            b["name"] = named.get(b["color"], "")
            b["inferred"] = bool(b["name"])
    boxes = [b for b in boxes if b["name"]]
    names = []
    for b in boxes:
        if b["name"] not in [n["name"] for n in names]:
            names.append(dict(name=b["name"], color=b["color"]))
    seen = {}
    for n in names:  # dua customer berwarna sama: yang kedua dibuat sedikit lebih terang
        if _hex_hls(n["color"])[1] > 0.80 and pucat:
            n["color"] = pucat.pop(0)
        if n["color"] in seen:
            h, l, s_ = _hex_hls(n["color"])
            n["color"] = _hls_hex(h, l + 0.16, s_)
        seen[n["color"]] = 1
    for b in boxes:
        b["idx"] = [n["name"] for n in names].index(b["name"])
    return names, boxes, (cx, cy, colw, rowh)


def cust_of(boxes, geo, row, col):
    """Index customer untuk sel (row, col) 1-based; -1 bila tidak tertutup kotak mana pun."""
    cx, cy, colw, rowh = geo
    x = cx[col - 1] + colw(col) / 2
    y = cy[row - 1] + rowh(row) / 2
    hit = [b for b in boxes if b["x0"] <= x <= b["x1"] and b["y0"] <= y <= b["y1"]]
    if not hit:
        return -1
    hit.sort(key=lambda b: (b["x1"] - b["x0"]) * (b["y1"] - b["y0"]))
    return hit[0]["idx"]


def tandai_customer(data, file, ws, sheet_index):
    """Tambahkan indeks customer sebagai elemen ke-5 tiap sel, plus data['cust'] dan data['cnote']."""
    try:
        names, boxes, geo = baca_customer(file, ws, sheet_index)
    except Exception:
        return data
    if not names:
        return data
    for r in data["racks"]:
        for c in r["cells"]:
            c.append(cust_of(boxes, geo, c[0], c[1]))
    data["cust"] = names
    inf = {}
    for b in boxes:
        if b["inferred"]:
            inf[b["name"]] = inf.get(b["name"], 0) + 1
    notes = [f"{n} kotak tanpa nama dibaca sebagai {nm} (warna kotak sama)." for nm, n in inf.items()]
    notes.append("Batas customer mengikuti posisi kotak teks di Excel, jadi sel di tepi kotak bisa bergeser satu kolom.")
    data["cnote"] = notes
    return data
