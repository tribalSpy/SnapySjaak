"""Maak Stickers - dubbelklik-script. Verwacht Sticker test.xlsx en Halindeling.xlsx in dezelfde map."""
import os, sys, subprocess
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
STICKER_XLSX = SCRIPT_DIR / "Sticker test.xlsx"
HAL_XLSX = SCRIPT_DIR / "Halindeling.xlsx"


def ensure_packages():
    needed = {"pandas": "pandas", "openpyxl": "openpyxl", "reportlab": "reportlab"}
    miss = [pkg for mod, pkg in needed.items() if not __try_import(mod)]
    if miss:
        print(f"Installeren: {', '.join(miss)} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", *miss])


def __try_import(m):
    try:
        __import__(m); return True
    except ImportError:
        return False


ensure_packages()

import struct, zlib, zipfile, re
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF

PAGE_W, PAGE_H = A4
DRAW_W, DRAW_H = PAGE_H, PAGE_W


def _fix_sst(unc):
    if unc.rstrip().endswith(b"</sst>"):
        return unc
    last = unc.rfind(b"</si>")
    if last == -1:
        h = unc.find(b"<sst")
        if h == -1: return unc
        return unc[:unc.find(b">", h) + 1] + b"</sst>"
    body = unc[:last + 5]
    m = re.search(rb'uniqueCount="(\d+)"', body)
    target = int(m.group(1)) if m else 0
    cur = body.count(b"<si>") + body.count(b"<si ")
    return body + b"<si><t></t></si>" * max(0, target - cur) + b"</sst>"


def repair_xlsx(src, dst):
    data = Path(src).read_bytes()
    entries, i = [], 0
    while i < len(data) - 30:
        if data[i:i + 4] == b"PK\x03\x04":
            nlen = struct.unpack("<H", data[i + 26:i + 28])[0]
            elen = struct.unpack("<H", data[i + 28:i + 30])[0]
            comp = struct.unpack("<I", data[i + 18:i + 22])[0]
            name = data[i + 30:i + 30 + nlen].decode("utf-8", "replace")
            ds = i + 30 + nlen + elen
            entries.append((name, ds, comp))
            i = ds + max(comp, 1)
        else:
            i += 1
    files = {}
    for name, ds, comp in entries:
        d = zlib.decompressobj(-15)
        avail = data[ds:ds + comp] if comp <= len(data) - ds else data[ds:]
        try: unc = d.decompress(avail)
        except zlib.error: unc = b""
        if name == "xl/sharedStrings.xml":
            unc = _fix_sst(unc)
        files[name] = unc
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zf:
        for n, c in files.items():
            zf.writestr(n, c)


def load_xlsx_safe(path, **kw):
    try:
        zipfile.ZipFile(path).namelist()
        return pd.read_excel(path, **kw)
    except zipfile.BadZipFile:
        rep = SCRIPT_DIR / ("_repaired_" + Path(path).name)
        repair_xlsx(path, rep)
        return pd.read_excel(rep, **kw)


def build_location_lookup():
    df = load_xlsx_safe(HAL_XLSX, sheet_name="Blad1", header=None)
    df["loc"] = df[0].ffill()
    df = df.dropna(subset=[1])
    df = df[df[1].astype(str).str.match(r"^[A-Za-z0-9#]+$")]
    out = {}
    for _, r in df.iterrows():
        if r[1] not in out:
            out[r[1]] = r["loc"]
    return out


def location_sort_key(loc):
    if not loc: return []
    return [int(p) if p.isdigit() else p for p in re.findall(r"[A-Za-z]+|\d+", str(loc))]


def draw_qr(c, value, x, y, size):
    qc = qr.QrCodeWidget(str(value))
    b = qc.getBounds()
    w, h = b[2] - b[0], b[3] - b[1]
    d = Drawing(size, size, transform=[size / w, 0, 0, size / h, 0, 0])
    d.add(qc)
    renderPDF.draw(d, c, x, y)


def draw_sticker(c, location, truck_num, counter, klantcode, klantnaam, c1, c2):
    c.saveState(); c.translate(PAGE_W, 0); c.rotate(90)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(40, 545, str(location) if location else "")
    if truck_num is not None:
        c.setFont("Helvetica-Bold", 24)
        c.drawString(40, 510, f"Truck {truck_num}")
    c.setFont("Helvetica-Bold", 110)
    cw = c.stringWidth(str(klantcode), "Helvetica-Bold", 110)
    c.drawString((DRAW_W - cw) / 2, 485, str(klantcode))
    c.setFont("Helvetica-Bold", 56)
    tw = c.stringWidth(counter, "Helvetica-Bold", 56)
    c.drawString(DRAW_W - 40 - tw, 510, counter)
    naam = str(klantnaam).upper() if klantnaam else ""
    c.setFont("Helvetica-Bold", 32)
    nw = c.stringWidth(naam, "Helvetica-Bold", 32)
    c.drawString((DRAW_W - nw) / 2, 425, naam)
    qr_size = 200
    qr_x = DRAW_W - qr_size - 40
    draw_qr(c, klantcode, qr_x, 110, qr_size)
    car_center = (40 + qr_x - 25) / 2
    if c1:
        c.setFont("Helvetica-Bold", 80)
        bw = c.stringWidth(str(c1), "Helvetica-Bold", 80)
        c.drawString(car_center - bw / 2, 220, str(c1))
    if c2:
        c.setFont("Helvetica-Bold", 80)
        vw = c.stringWidth(str(c2), "Helvetica-Bold", 80)
        c.drawString(car_center - vw / 2, 80, str(c2))
    c.restoreState(); c.showPage()


def make_pdf(rows, truck_label, loc_lookup, out_path):
    def key(r):
        c1 = "" if pd.isna(r.get("Carrier 1")) else str(r.get("Carrier 1")).strip()
        loc = loc_lookup.get(r["klant"], "")
        return (c1, location_sort_key(loc), str(loc))
    rows = sorted(rows, key=key)
    c = canvas.Canvas(str(out_path), pagesize=A4)
    n, miss = 0, []
    for r in rows:
        klant = r["klant"]; naam = r["naam"]
        c1v = r.get("Carrier 1"); c2v = r.get("Carrier 2")
        c1 = "" if pd.isna(c1v) else str(c1v).strip()
        c2 = "" if pd.isna(c2v) else str(c2v).strip()
        try: cnt = int(r["Split CC"])
        except Exception: cnt = 1
        if cnt < 1: cnt = 1
        loc = loc_lookup.get(klant, "")
        if not loc: miss.append(klant)
        for i in range(1, cnt + 1):
            draw_sticker(c, loc, truck_label, f"{i}-{cnt}", klant, naam, c1, c2)
            n += 1
    c.save()
    return n, miss


def main():
    print("Stickers genereren...")
    if not STICKER_XLSX.exists():
        print(f"FOUT: {STICKER_XLSX.name} ontbreekt"); return
    if not HAL_XLSX.exists():
        print(f"FOUT: {HAL_XLSX.name} ontbreekt"); return
    loc = build_location_lookup()
    print(f"Halindeling: {len(loc)} klanten")
    df = load_xlsx_safe(STICKER_XLSX)
    print(f"Sticker rijen: {len(df)}")
    trucks = sorted(df.dropna(subset=["Split"])["Split"].unique())
    all_miss = []
    for t in trucks:
        sub = df[df["Split"] == t].to_dict("records")
        out = SCRIPT_DIR / f"Stickers_truck_{int(t)}.pdf"
        n, m = make_pdf(sub, int(t), loc, out)
        all_miss.extend(m)
        print(f"  Truck {int(t)}: {len(sub)} klanten, {n} stickers -> {out.name}")
    nos = df[df["Split"].isna()].to_dict("records")
    if nos:
        out = SCRIPT_DIR / "Stickers_overig.pdf"
        n, m = make_pdf(nos, None, loc, out)
        all_miss.extend(m)
        print(f"  Overig: {len(nos)} klanten, {n} stickers -> {out.name}")
    if all_miss:
        print("Geen locatie voor:", ", ".join(sorted(set(all_miss))))
    print("Klaar.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback; traceback.print_exc()
    finally:
        if os.name == "nt":
            input("\nDruk Enter om af te sluiten...")
