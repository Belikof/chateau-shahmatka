"""Наложить топографию на генплан + легенда высот."""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
GENPLAN = ROOT / "data" / "genplan_commerce.png"
TOPO = ROOT / "data" / "topo_elevation.png"
OUT = ROOT / "data" / "genplan_topo.png"
OUT_HI = ROOT / "data" / "genplan_topo_hi.png"

# зона участков на генплане (пиксели)
MAP_BOX = (12, 248, 1200, 1458)
TOPO_OPACITY = 0.52


def _font(size: int, bold: bool = False):
    paths = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in paths:
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def _topo_rgba(topo: Image.Image, size: tuple[int, int]) -> Image.Image:
    """Масштабировать топографию и убрать белый фон."""
    t = topo.convert("RGBA")
    # обрезка полей: контент по не-белым пикселям
    px = t.load()
    w, h = t.size
    xs, ys = [], []
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if r < 245 or g < 245 or b < 245:
                xs.append(x)
                ys.append(y)
    if xs:
        pad = 8
        box = (
            max(0, min(xs) - pad),
            max(0, min(ys) - pad),
            min(w, max(xs) + pad),
            min(h, max(ys) + pad),
        )
        t = t.crop(box)
    t = t.resize(size, Image.LANCZOS)
    px = t.load()
    for y in range(t.height):
        for x in range(t.width):
            r, g, b, a = px[x, y]
            lum = (r + g + b) / 3
            if lum > 248:
                px[x, y] = (r, g, b, 0)
            else:
                alpha = int(min(255, (255 - lum) * 1.6) * TOPO_OPACITY)
                px[x, y] = (r, g, b, alpha)
    return t


def _legend(draw: ImageDraw.ImageDraw, x0: int, y0: int) -> None:
    f_b = _font(13, True)
    f = _font(11)
    items = [
        ("130 м", (210, 55, 35)),
        ("110 м", (230, 170, 60)),
        ("90 м", (170, 210, 120)),
        ("80 м", (120, 200, 220)),
    ]
    w, h = 168, 118
    draw.rounded_rectangle((x0, y0, x0 + w, y0 + h), radius=8, fill=(255, 255, 255, 215))
    draw.text((x0 + 10, y0 + 6), "ВЫСОТЫ", fill=(30, 30, 30), font=f_b)
    yy = y0 + 26
    for label, color in items:
        draw.rectangle((x0 + 10, yy, x0 + 28, yy + 14), fill=color)
        draw.text((x0 + 34, yy - 1), label, fill=(40, 40, 40), font=f)
        yy += 22
    draw.text((x0 + 10, y0 + h - 18), "Запад → Восток", fill=(80, 80, 80), font=f)


def build(base_path: Path = GENPLAN, topo_path: Path = TOPO, out_path: Path = OUT) -> Path:
    if not base_path.exists():
        raise SystemExit(f"Нет генплана: {base_path}")
    if not topo_path.exists():
        raise SystemExit(f"Нет топографии: {topo_path}")

    base = Image.open(base_path).convert("RGBA")
    topo = Image.open(topo_path)
    x0, y0, x1, y1 = MAP_BOX
    mw, mh = x1 - x0, y1 - y0
    layer = _topo_rgba(topo, (mw, mh))

    out = base.copy()
    out.alpha_composite(layer, (x0, y0))

    draw = ImageDraw.Draw(out)
    _legend(draw, x0 + 14, y1 - 130)

    # подпись пика
    f = _font(10)
    draw.text((x1 - 155, y0 + 18), "Поливадина 130 м →", fill=(180, 50, 30), font=f)

    rgb = out.convert("RGB")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    rgb.save(out_path, "PNG", optimize=True)

    # версия 2× для чёткости в таблице
    hi = rgb.resize((rgb.width * 2, rgb.height * 2), Image.LANCZOS)
    hi.save(OUT_HI, "PNG", optimize=True)
    print(f"Сохранено: {out_path} ({rgb.width}×{rgb.height})")
    print(f"Сохранено: {OUT_HI} ({hi.width}×{hi.height})")
    return out_path


def main() -> None:
    build()


if __name__ == "__main__":
    main()
