"""Synthesize scanned survey pages (BMP) + definition CSV for OMR tests.

Geometry mimics what Excel VBA (WriteOmrDef) measures: marker-relative
points. Renders at 150dpi with rotation, noise, gray pencil marks.
"""
import math
import random
import struct
import sys
from pathlib import Path

OUT = Path(__file__).resolve().parent

# --- form geometry (pt, origin = TL marker center) ---
DD, KK = 3, 12
W_PT, H_PT = 500.0, 360.0          # marker spacing
GRID = dict(x0=120.0, y0=70.0, pitchX=100.0, pitchY=22.0, cellW=96.0, cellH=20.0)
ID_X, ID_W = 0.0, 18.0
ID_MODS = [(16.0 + 17.0 * i, 15.0) for i in range(8)]   # (y, h)

PAGE_W_PT, PAGE_H_PT = 595.0, 842.0
MARGIN_X_PT, MARGIN_Y_PT = 45.0, 40.0    # TL marker center on page

DPI = 150
S = DPI / 72.0


def make_def(path):
    lines = ["PTMOMR,2",
             f"size,{W_PT:.2f},{H_PT:.2f}",
             f"days,{DD},slots,{KK}",
             f"grid,{GRID['x0']:.2f},{GRID['y0']:.2f},{GRID['pitchX']:.2f},"
             f"{GRID['pitchY']:.2f},{GRID['cellW']:.2f},{GRID['cellH']:.2f}"]
    idline = f"idstrip,{ID_X:.2f},{ID_W:.2f}"
    for y, h in ID_MODS:
        idline += f",{y:.2f},{h:.2f}"
    lines.append(idline)
    labels = []
    for d in range(DD):
        for k in range(KK):
            labels.append(f"6/{20 + d} {13 + k // 4}:{(k % 4) * 15:02d}")
    lines.append("labels," + ",".join(labels))
    for n in range(1, 36):
        lines.append(f"student,{n},テスト 生徒{n}")
    path.write_text("﻿" + "\r\n".join(lines) + "\r\n", encoding="utf-8")
    return labels


class Canvas:
    def __init__(self, w, h, bg=250):
        self.w, self.h = w, h
        self.px = bytearray([bg]) * (w * h)

    def rect(self, x0, y0, x1, y1, val):
        x0, y0 = max(0, int(x0)), max(0, int(y0))
        x1, y1 = min(self.w, int(x1)), min(self.h, int(y1))
        for y in range(y0, y1):
            base = y * self.w
            for x in range(x0, x1):
                self.px[base + x] = val


def render_page(student_id, marked, rot_deg=1.2, noise=4, seed=1,
                dpi=150, pencil=115):
    rng = random.Random(seed)
    S = dpi / 72.0
    w, h = int(PAGE_W_PT * S), int(PAGE_H_PT * S)
    c = Canvas(w, h)

    def pt2px(fx, fy):
        return ((MARGIN_X_PT + fx) * S, (MARGIN_Y_PT + fy) * S)

    # markers (18x16pt)
    for fx, fy in [(0, 0), (W_PT, 0), (0, H_PT), (W_PT, H_PT)]:
        x, y = pt2px(fx, fy)
        c.rect(x - 9 * S, y - 8 * S, x + 9 * S, y + 8 * S, 10)

    # id strip
    bits = [(student_id >> (6 - i)) & 1 for i in range(7)]
    bits.append(1 - sum(bits) % 2)   # odd parity
    for (y0, hh), b in zip(ID_MODS, bits):
        if b:
            x, y = pt2px(ID_X, y0)
            c.rect(x - ID_W / 2 * S, y - hh / 2 * S, x + ID_W / 2 * S, y + hh / 2 * S, 10)

    # grid lines (thin, light)
    gx0 = GRID["x0"] - GRID["cellW"] / 2
    gy0 = GRID["y0"] - GRID["cellH"] / 2
    for d in range(DD + 1):
        x, _ = pt2px(gx0 + d * GRID["pitchX"], 0)
        c.rect(x, pt2px(0, gy0)[1], x + 1, pt2px(0, gy0 + KK * GRID["pitchY"])[1], 150)
    for k in range(KK + 1):
        _, y = pt2px(0, gy0 + k * GRID["pitchY"])
        c.rect(pt2px(gx0, 0)[0], y, pt2px(gx0 + DD * GRID["pitchX"], 0)[0], y + 1, 150)

    # pencil marks (gray fill, inner 70%)
    for (d, k) in marked:
        fx = GRID["x0"] + d * GRID["pitchX"]
        fy = GRID["y0"] + k * GRID["pitchY"]
        x, y = pt2px(fx, fy)
        rw, rh = GRID["cellW"] * 0.35 * S, GRID["cellH"] * 0.35 * S
        c.rect(x - rw, y - rh, x + rw, y + rh, pencil)

    # some header-ish text noise lines
    x, y = pt2px(100, 14)
    c.rect(x, y, x + 200 * S, y + 2, 80)

    # rotate + noise into output canvas
    out = Canvas(w, h, bg=252)
    th = math.radians(rot_deg)
    cx, cy = w / 2, h / 2
    cos, sin = math.cos(th), math.sin(th)
    for y in range(h):
        for x in range(w):
            sx = cos * (x - cx) + sin * (y - cy) + cx
            sy = -sin * (x - cx) + cos * (y - cy) + cy
            if 0 <= sx < w and 0 <= sy < h:
                v = c.px[int(sy) * w + int(sx)]
            else:
                v = 252
            v += rng.randint(-noise, noise)
            out.px[y * w + x] = max(0, min(255, v))
    return out


def save_bmp(canvas, path):
    w, h = canvas.w, canvas.h
    row = (w * 3 + 3) & ~3
    size = 54 + row * h
    hdr = struct.pack("<2sIHHI", b"BM", size, 0, 0, 54)
    hdr += struct.pack("<IiiHHIIiiII", 40, w, h, 1, 24, 0, row * h, 2835, 2835, 0, 0)
    body = bytearray()
    for y in range(h - 1, -1, -1):
        line = bytearray()
        base = y * w
        for x in range(w):
            v = canvas.px[base + x]
            line += bytes([v, v, v])
        line += b"\x00" * (row - len(line))
        body += line
    path.write_bytes(hdr + body)


def main():
    make_def(OUT / "test_def.csv")
    cases = [
        # (id, marks, rotation, dpi, pencil, filename)
        (7, [(0, 1), (0, 5), (1, 3), (2, 11)], 1.2, 150, 115, "page_07.bmp"),
        (35, [(2, 0), (1, 7)], -1.6, 150, 115, "page_35.bmp"),
        (12, [], 0.4, 150, 115, "page_12_empty.bmp"),
        (64, [(0, 0), (2, 5)], 3.0, 150, 115, "page_64_rot3.bmp"),
        (3, [(1, 11), (2, 3)], -0.8, 100, 115, "page_03_100dpi.bmp"),
        (50, [(0, 2), (1, 9), (2, 7)], 1.0, 200, 140, "page_50_light.bmp"),
    ]
    manifest = []
    for sid, marks, rot, dpi, pencil, name in cases:
        page = render_page(sid, marks, rot_deg=rot, seed=sid, dpi=dpi, pencil=pencil)
        save_bmp(page, OUT / name)
        manifest.append(f"{name};{sid};" + "|".join(f"{d}-{k}" for d, k in marks))
        print("rendered", name)
    (OUT / "manifest.txt").write_text("\n".join(manifest))


if __name__ == "__main__":
    main()
