from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(exist_ok=True)


def make_icon() -> None:
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = []

    for size in sizes:
        image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        scale = size / 256
        radius = int(44 * scale)
        draw.rounded_rectangle(
            (0, 0, size - 1, size - 1),
            radius=radius,
            fill=(22, 119, 255, 255),
        )
        draw.rounded_rectangle(
            (
                int(58 * scale),
                int(46 * scale),
                int(198 * scale),
                int(210 * scale),
            ),
            radius=int(18 * scale),
            fill=(255, 255, 255, 255),
        )
        draw.line(
            (
                int(85 * scale),
                int(100 * scale),
                int(171 * scale),
                int(100 * scale),
            ),
            fill=(22, 119, 255, 255),
            width=max(2, int(12 * scale)),
        )
        draw.line(
            (
                int(85 * scale),
                int(135 * scale),
                int(155 * scale),
                int(135 * scale),
            ),
            fill=(22, 119, 255, 255),
            width=max(2, int(12 * scale)),
        )
        draw.line(
            (
                int(85 * scale),
                int(170 * scale),
                int(135 * scale),
                int(170 * scale),
            ),
            fill=(22, 119, 255, 255),
            width=max(2, int(12 * scale)),
        )
        images.append(image)

    images[-1].save(ASSETS / "app.ico", sizes=[(size, size) for size in sizes])


if __name__ == "__main__":
    make_icon()
