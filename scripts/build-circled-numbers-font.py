"""Build the circled-number fallback from the pinned Noto Sans Symbols source.

Run with FontTools installed and pass the original TTF linked in fonts/README.md.
"""
from pathlib import Path
import sys

from fontTools import subset
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont

FONTS = Path(__file__).resolve().parents[1] / "fonts"
FAMILY = "Document Kit Circled Numbers"
CIRCLED_NUMBERS = (0x24EA, *range(0x2460, 0x2474))


def build(source: Path, style: str, weight: int) -> None:
    font = instantiateVariableFont(TTFont(source, recalcTimestamp=False), {"wght": weight})
    options = subset.Options()
    options.name_IDs = [0, 1, 2, 3, 4, 5, 6, 13, 14]
    selected = subset.Subsetter(options=options)
    selected.populate(unicodes=CIRCLED_NUMBERS)
    selected.subset(font)

    for record in font["name"].names:
        if record.nameID == 1:
            record.string = FAMILY
        elif record.nameID == 2:
            record.string = style
        elif record.nameID in (3, 4):
            record.string = f"{FAMILY} {style}"
        elif record.nameID == 6:
            record.string = f"DocumentKitCircledNumbers-{style}"
    font.save(FONTS / f"DocumentKitCircledNumbers-{style}.ttf")


if __name__ == "__main__":
    source = Path(sys.argv[1])
    for style, weight in (("Regular", 400), ("Bold", 700)):
        build(source, style, weight)
