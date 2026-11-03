"""
# TASK
Make PPT from images

Put images one by one in the slides of a PowerPoint presentation (PPTX) ordered by the timestamped screenshot names.

# OUTPUT
Python script for the task.

# USAGE
    python make_ppt_from_images.py                      # build img.pptx from the images next to this script
    python make_ppt_from_images.py -o course.pptx
    python make_ppt_from_images.py --title "Project Management - The Planning Phase" --fit cover
    python make_ppt_from_images.py --dry-run            # only print the slide order

# REQUIREMENTS
    pip install python-pptx pillow
"""

from __future__ import annotations

import argparse
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Sequence

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches

IMAGE_EXTENSIONS = {
    ".png", ".jpg", ".jpeg", ".jpe", ".gif", ".bmp", ".dib", ".tif", ".tiff", ".webp",
}
# Formats python-pptx can embed as-is; anything else is converted to PNG first.
NATIVE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".jpe", ".gif", ".bmp", ".dib", ".tif", ".tiff"}
SLIDE_SIZES = {
    "16:9": (13.333, 7.5),
    "16:10": (10.0, 6.25),
    "4:3": (10.0, 7.5),
    "3:2": (13.333, 8.889),
}
BLANK_LAYOUT_INDEX = 6

# Matches "Screenshot 2026-10-28 154358.png" and similar date/time styles,
# e.g. "2026-10-28 15-43-58", "2026-10-28_154358", "2026-10-28 1543".
TIMESTAMP_RE = re.compile(
    r"(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})"
    r"[ _T]+(?P<hour>\d{2})[-:.]?(?P<minute>\d{2})(?:[-:.]?(?P<second>\d{2}))?"
)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a PPTX deck with one image per slide, ordered by the timestamp in the file name.",
    )
    parser.add_argument("-i", "--input", default=".", help="folder with the images (default: current folder)")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="output .pptx path (default: <input folder name>.pptx next to the input folder)",
    )
    parser.add_argument(
        "--sort",
        choices=("timestamp", "name", "mtime"),
        default="timestamp",
        help="slide order; 'timestamp' parses the date/time from the file name and "
        "falls back to the file modification time (default: timestamp)",
    )
    parser.add_argument(
        "--slide-size",
        default="16:9",
        help="16:9, 16:10, 4:3, 3:2 or WIDTHxHEIGHT in inches (default: 16:9)",
    )
    parser.add_argument(
        "--fit",
        choices=("contain", "cover", "stretch"),
        default="contain",
        help="how each image fills the slide: 'contain' keeps the whole image without distortion, "
        "'cover' crops it to fill the slide, 'stretch' distorts it to fill the slide (default: contain)",
    )
    parser.add_argument("--margin", type=float, default=0.25, help="margin around each image in inches (default: 0.25)")
    parser.add_argument("--title", default=None, help="optional title added as the first slide")
    parser.add_argument("--background", default=None, help="slide background color as RRGGBB hex, e.g. FFFFFF")
    parser.add_argument("--no-notes", action="store_true", help="do not put the source image name in the speaker notes")
    parser.add_argument("-r", "--recursive", action="store_true", help="also scan sub folders")
    parser.add_argument("--dry-run", action="store_true", help="only print the slide order, do not write the deck")
    return parser.parse_args(argv)


def parse_slide_size(value: str) -> tuple[float, float]:
    if value in SLIDE_SIZES:
        return SLIDE_SIZES[value]
    match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*[xX*]\s*(\d+(?:\.\d+)?)", value.strip())
    if not match:
        raise argparse.ArgumentTypeError(
            f"invalid --slide-size {value!r}; use 16:9, 16:10, 4:3, 3:2 or WIDTHxHEIGHT"
        )
    width, height = float(match.group(1)), float(match.group(2))
    if width <= 0 or height <= 0:
        raise argparse.ArgumentTypeError("--slide-size width and height must be positive")
    return width, height


def natural_key(text: str) -> list[tuple[int, object]]:
    """Sort key that orders 'slide 2' before 'slide 10'."""
    return [(0, int(part)) if part.isdigit() else (1, part.lower()) for part in re.split(r"(\d+)", text) if part]


def timestamp_of(path: Path) -> datetime | None:
    match = TIMESTAMP_RE.search(path.stem)
    if not match:
        return None
    parts = {name: int(value or 0) for name, value in match.groupdict().items()}
    try:
        return datetime(
            parts["year"], parts["month"], parts["day"],
            parts["hour"], parts["minute"], parts["second"],
        )
    except ValueError:
        return None


def collect_images(input_dir: Path, recursive: bool) -> list[Path]:
    walker = input_dir.rglob("*") if recursive else input_dir.glob("*")
    return [path for path in walker if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS]


def sort_images(images: list[Path], mode: str) -> list[Path]:
    if mode == "name":
        return sorted(images, key=lambda path: natural_key(path.name))
    if mode == "mtime":
        return sorted(images, key=lambda path: (path.stat().st_mtime, natural_key(path.name)))
    return sorted(
        images,
        key=lambda path: (timestamp_of(path) or datetime.fromtimestamp(path.stat().st_mtime), natural_key(path.name)),
    )


def fit_contain(
    src_w: float, src_h: float, box_w: float, box_h: float
) -> tuple[float, float, tuple[float, float, float, float]]:
    scale = min(box_w / src_w, box_h / src_h)
    return src_w * scale, src_h * scale, (0.0, 0.0, 0.0, 0.0)


def fit_cover(
    src_w: float, src_h: float, box_w: float, box_h: float
) -> tuple[float, float, tuple[float, float, float, float]]:
    src_ratio = src_w / src_h
    box_ratio = box_w / box_h
    if src_ratio > box_ratio:
        side = (1.0 - box_ratio / src_ratio) / 2.0
        return box_w, box_h, (side, side, 0.0, 0.0)
    side = (1.0 - src_ratio / box_ratio) / 2.0
    return box_w, box_h, (0.0, 0.0, side, side)


FITTERS = {
    "contain": fit_contain,
    "cover": fit_cover,
    "stretch": lambda w, h, bw, bh: (bw, bh, (0.0, 0.0, 0.0, 0.0)),
}


def prepare_image(path: Path, tmp_dir: Path) -> Path:
    """Return a file python-pptx can embed, converting unsupported formats to PNG."""
    if path.suffix.lower() in NATIVE_EXTENSIONS:
        return path
    target = tmp_dir / f"{path.stem}.png"
    with Image.open(path) as image:
        if image.mode not in ("RGB", "RGBA"):
            image = image.convert("RGBA")
        image.save(target, "PNG")
    return target


def build_deck(
    images: list[Path],
    output: Path,
    slide_size: tuple[float, float],
    *,
    margin: float,
    fit: str,
    title: str | None,
    background: str | None,
    add_notes: bool,
) -> tuple[int, list[tuple[Path, str]]]:
    slide_w, slide_h = slide_size
    box_w = slide_w - 2 * margin
    box_h = slide_h - 2 * margin
    if box_w <= 0 or box_h <= 0:
        raise ValueError("--margin is too large for the selected --slide-size")

    prs = Presentation()
    prs.slide_width = Inches(slide_w)
    prs.slide_height = Inches(slide_h)
    blank_layout = prs.slide_layouts[BLANK_LAYOUT_INDEX]

    if title:
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = title

    failures: list[tuple[Path, str]] = []
    added = 0
    with tempfile.TemporaryDirectory(prefix="pptx_images_") as tmp:
        tmp_dir = Path(tmp)
        for image_path in images:
            try:
                embed_path = prepare_image(image_path, tmp_dir)
                with Image.open(embed_path) as image:
                    src_w, src_h = image.size
                if src_w <= 0 or src_h <= 0:
                    raise ValueError("image has an invalid size")
                disp_w, disp_h, (crop_left, crop_right, crop_top, crop_bottom) = FITTERS[fit](
                    src_w, src_h, box_w, box_h
                )

                slide = prs.slides.add_slide(blank_layout)
                if background:
                    fill = slide.background.fill
                    fill.solid()
                    fill.fore_color.rgb = RGBColor.from_string(background)

                picture = slide.shapes.add_picture(
                    str(embed_path),
                    Inches(margin + (box_w - disp_w) / 2),
                    Inches(margin + (box_h - disp_h) / 2),
                    width=Inches(disp_w),
                    height=Inches(disp_h),
                )
                picture.crop_left = crop_left
                picture.crop_right = crop_right
                picture.crop_top = crop_top
                picture.crop_bottom = crop_bottom

                if add_notes:
                    slide.notes_slide.notes_text_frame.text = image_path.name
                added += 1
            except Exception as exc:  # keep going so one bad file does not abort the deck
                failures.append((image_path, str(exc)))

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))
    return added, failures


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)

    input_dir = Path(args.input).expanduser().resolve()
    if not input_dir.is_dir():
        print(f"error: input folder not found: {input_dir}", file=sys.stderr)
        return 2

    try:
        slide_size = parse_slide_size(args.slide_size)
    except argparse.ArgumentTypeError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    background = None
    if args.background:
        background = args.background.strip().lstrip("#").upper()
        if not re.fullmatch(r"[0-9A-F]{6}", background):
            print("error: --background must be a 6 digit hex color such as FFFFFF", file=sys.stderr)
            return 2

    images = sort_images(collect_images(input_dir, args.recursive), args.sort)
    if not images:
        print(f"error: no images found in {input_dir}", file=sys.stderr)
        return 1

    output = Path(args.output).expanduser() if args.output else input_dir.parent / f"{input_dir.name}.pptx"
    output = output.resolve()

    print(f"Found {len(images)} image(s) in {input_dir}")

    if args.dry_run:
        for index, image_path in enumerate(images, start=1):
            print(f"{index:4d}. {image_path.relative_to(input_dir)}")
        print(f"Output would be: {output}")
        return 0

    try:
        added, failures = build_deck(
            images,
            output,
            slide_size,
            margin=args.margin,
            fit=args.fit,
            title=args.title,
            background=background,
            add_notes=not args.no_notes,
        )
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    for image_path, reason in failures:
        print(f"warning: skipped {image_path.name}: {reason}", file=sys.stderr)

    title_note = " plus 1 title slide" if args.title else ""
    print(f"Created {output} with {added} image slide(s){title_note}")
    if failures:
        print(f"Skipped {len(failures)} image(s)")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
