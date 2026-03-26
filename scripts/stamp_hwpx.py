#!/usr/bin/env python3
"""Generate a red stamp (도장) image and overlay it on "(인)" in an HWPX file.

Creates a transparent-background red circular seal with the given name,
then finds "(인)" text in the HWPX document and overlays the stamp image
at the same position and size.

Usage:
    # Apply stamp to existing HWPX
    python stamp_hwpx.py --name 홍길동 --input document.hwpx --output stamped.hwpx

    # Generate stamp image only
    python stamp_hwpx.py --name 홍길동 --image-only --output stamp.png

    # Custom stamp size (HWPUNIT, default 3300 ≈ 11.6mm)
    python stamp_hwpx.py --name 홍길동 --input doc.hwpx --output out.hwpx --stamp-size 4000
"""

import argparse
import copy
import math
import shutil
import sys
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, BadZipFile, ZipFile

from lxml import etree
from PIL import Image, ImageDraw, ImageFont

# ──────────────────────────────────────────
# Constants
# ──────────────────────────────────────────

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "opf": "http://www.idpf.org/2007/opf/",
}

# Default stamp size in HWPUNIT (3300 ≈ 11.6mm, matches standard 인감 size)
DEFAULT_STAMP_SIZE = 3300

# Stamp image resolution in pixels
STAMP_PX = 400

# Red color for stamp
STAMP_RED = (220, 30, 30, 230)  # slightly transparent for realism
STAMP_BORDER_RED = (200, 20, 20, 245)

# Korean fonts to try (Windows paths)
FONT_CANDIDATES = [
    "C:/Windows/Fonts/malgunbd.ttf",    # 맑은 고딕 Bold
    "C:/Windows/Fonts/malgun.ttf",       # 맑은 고딕
    "C:/Windows/Fonts/HANBatangB.TTF",   # 한컴바탕 Bold
    "C:/Windows/Fonts/batang.ttc",       # 바탕
    "C:/Windows/Fonts/gulim.ttc",        # 굴림
]


def _find_font() -> str:
    """Find an available Korean font."""
    for path in FONT_CANDIDATES:
        if Path(path).is_file():
            return path
    raise SystemExit("No Korean font found. Install 맑은 고딕 or similar.")


def generate_stamp(name: str, output_path: Path, px: int = STAMP_PX) -> None:
    """Generate a red circular stamp PNG with transparent background."""
    img = Image.new("RGBA", (px, px), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    center = px // 2
    radius = int(px * 0.45)
    border_width = max(3, int(px * 0.04))

    # Draw outer circle border
    for i in range(border_width):
        draw.ellipse(
            [center - radius + i, center - radius + i,
             center + radius - i, center + radius - i],
            outline=STAMP_BORDER_RED,
        )

    # Draw inner rough edge effect (looks more like real stamp)
    import random
    random.seed(hash(name))  # deterministic per name
    for angle in range(0, 360, 2):
        rad = math.radians(angle)
        r_var = radius - border_width - random.randint(0, 2)
        x1 = center + int(r_var * math.cos(rad))
        y1 = center + int(r_var * math.sin(rad))
        # Small dots along inner edge for texture
        if random.random() < 0.3:
            draw.ellipse([x1 - 1, y1 - 1, x1 + 1, y1 + 1], fill=STAMP_RED)

    # Determine text layout based on name length
    font_path = _find_font()
    name_chars = list(name)
    n = len(name_chars)

    if n <= 2:
        # Vertical layout for 2 chars
        _draw_vertical(draw, name_chars, center, radius, border_width, font_path, px)
    elif n <= 4:
        # 2x2 grid for 3-4 chars
        _draw_grid(draw, name_chars, center, radius, border_width, font_path, px)
    else:
        # Horizontal for longer names
        _draw_horizontal(draw, name, center, radius, border_width, font_path, px)

    img.save(str(output_path), "PNG")


def _draw_vertical(draw, chars, center, radius, bw, font_path, px):
    """Draw 1-2 characters vertically centered."""
    n = len(chars)
    text_area = int(radius * 1.2)
    font_size = int(text_area / max(n, 1) * 0.95)
    font = ImageFont.truetype(font_path, font_size)

    total_h = n * font_size
    start_y = center - total_h // 2

    for i, ch in enumerate(chars):
        bbox = draw.textbbox((0, 0), ch, font=font)
        tw = bbox[2] - bbox[0]
        x = center - tw // 2
        y = start_y + i * font_size
        draw.text((x, y), ch, fill=STAMP_RED, font=font)


def _draw_grid(draw, chars, center, radius, bw, font_path, px):
    """Draw 3-4 characters in a 2x2 grid.

    Korean seal reading order: right column top→bottom, then left column top→bottom.
    For 홍길동: 홍(top-right) 길(bottom-right) 동(top-left)
    For 홍길동X: 홍(top-right) 길(bottom-right) 동(top-left) X(bottom-left)
    """
    while len(chars) < 4:
        chars.append("")

    text_area = int(radius * 1.05)
    font_size = int(text_area * 0.52)
    font = ImageFont.truetype(font_path, font_size)

    cell = font_size + 4
    half_cell = cell // 2
    gap = int(cell * 0.05)

    # Cell centers for 2x2 grid
    # [top-left]  [top-right]
    # [bot-left]  [bot-right]
    cell_centers = {
        "TL": (center - half_cell - gap, center - half_cell - gap),
        "TR": (center + half_cell + gap, center - half_cell - gap),
        "BL": (center - half_cell - gap, center + half_cell + gap),
        "BR": (center + half_cell + gap, center + half_cell + gap),
    }

    # Reading order: right-col top→bottom, then left-col top→bottom
    # chars[0]→TR, chars[1]→BR, chars[2]→TL, chars[3]→BL
    slot_order = ["TR", "BR", "TL", "BL"]

    for i, slot in enumerate(slot_order):
        ch = chars[i] if i < len(chars) else ""
        if not ch:
            continue
        cx, cy = cell_centers[slot]
        bbox = draw.textbbox((0, 0), ch, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        x = cx - tw // 2
        y = cy - th // 2 - int(font_size * 0.1)
        draw.text((x, y), ch, fill=STAMP_RED, font=font)


def _draw_horizontal(draw, name, center, radius, bw, font_path, px):
    """Draw name horizontally for long names."""
    text_area = int(radius * 1.4)
    font_size = int(text_area / len(name) * 1.2)
    font_size = min(font_size, int(radius * 0.7))
    font = ImageFont.truetype(font_path, font_size)

    bbox = draw.textbbox((0, 0), name, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    x = center - tw // 2
    y = center - th // 2
    draw.text((x, y), name, fill=STAMP_RED, font=font)


def _build_pic_element(stamp_size: int, stamp_ref: str, name: str,
                       pic_id: int, inst_id: int) -> etree._Element:
    """Build the <hp:pic> XML element for the stamp image."""
    hp = NS["hp"]
    hc = NS["hc"]
    half = stamp_size // 2

    # The stamp overlaps (인) — positioned relative to paragraph
    # vertOffset negative = slightly above text baseline
    # horzOffset positive = to the right to cover (인)
    vert_offset = -(stamp_size // 3)
    horz_offset = stamp_size * 2 + stamp_size // 2  # approx position over (인)

    pic = etree.Element(f"{{{hp}}}pic")
    pic.set("id", str(pic_id))
    pic.set("zOrder", "999")
    pic.set("numberingType", "PICTURE")
    pic.set("textWrap", "IN_FRONT_OF_TEXT")
    pic.set("textFlow", "BOTH_SIDES")
    pic.set("lock", "0")
    pic.set("dropcapstyle", "None")
    pic.set("href", "")
    pic.set("groupLevel", "0")
    pic.set("instid", str(inst_id))
    pic.set("reverse", "0")

    etree.SubElement(pic, f"{{{hp}}}offset", x="0", y="0")
    etree.SubElement(pic, f"{{{hp}}}orgSz", width=str(stamp_size), height=str(stamp_size))
    etree.SubElement(pic, f"{{{hp}}}curSz", width=str(stamp_size), height=str(stamp_size))
    etree.SubElement(pic, f"{{{hp}}}flip", horizontal="0", vertical="0")
    etree.SubElement(pic, f"{{{hp}}}rotationInfo",
                     angle="0", centerX=str(half), centerY=str(half), rotateimage="0")

    ri = etree.SubElement(pic, f"{{{hp}}}renderingInfo")
    etree.SubElement(ri, f"{{{hc}}}transMatrix", e1="1", e2="0", e3="0", e4="0", e5="1", e6="0")
    etree.SubElement(ri, f"{{{hc}}}scaMatrix", e1="1", e2="0", e3="0", e4="0", e5="1", e6="0")
    etree.SubElement(ri, f"{{{hc}}}rotMatrix", e1="1", e2="0", e3="0", e4="0", e5="1", e6="0")

    img_el = etree.SubElement(pic, f"{{{hc}}}img")
    img_el.set("binaryItemIDRef", stamp_ref)
    img_el.set("bright", "0")
    img_el.set("contrast", "0")
    img_el.set("effect", "REAL_PIC")
    img_el.set("alpha", "0")

    ir = etree.SubElement(pic, f"{{{hp}}}imgRect")
    etree.SubElement(ir, f"{{{hc}}}pt0", x="0", y="0")
    etree.SubElement(ir, f"{{{hc}}}pt1", x=str(stamp_size), y="0")
    etree.SubElement(ir, f"{{{hc}}}pt2", x=str(stamp_size), y=str(stamp_size))
    etree.SubElement(ir, f"{{{hc}}}pt3", x="0", y=str(stamp_size))

    etree.SubElement(pic, f"{{{hp}}}imgClip", left="0", right="0", top="0", bottom="0")
    etree.SubElement(pic, f"{{{hp}}}inMargin", left="0", right="0", top="0", bottom="0")
    etree.SubElement(pic, f"{{{hp}}}imgDim", dimwidth=str(stamp_size), dimheight=str(stamp_size))
    etree.SubElement(pic, f"{{{hp}}}effects")

    sz = etree.SubElement(pic, f"{{{hp}}}sz")
    sz.set("width", str(stamp_size))
    sz.set("widthRelTo", "ABSOLUTE")
    sz.set("height", str(stamp_size))
    sz.set("heightRelTo", "ABSOLUTE")
    sz.set("protect", "0")

    pos = etree.SubElement(pic, f"{{{hp}}}pos")
    pos.set("treatAsChar", "0")
    pos.set("affectLSpacing", "0")
    pos.set("flowWithText", "1")
    pos.set("allowOverlap", "1")
    pos.set("holdAnchorAndSO", "0")
    pos.set("vertRelTo", "PARA")
    pos.set("horzRelTo", "PARA")
    pos.set("vertAlign", "TOP")
    pos.set("horzAlign", "LEFT")
    pos.set("vertOffset", str(vert_offset))
    pos.set("horzOffset", str(horz_offset))

    etree.SubElement(pic, f"{{{hp}}}outMargin", left="0", right="0", top="0", bottom="0")
    comment = etree.SubElement(pic, f"{{{hp}}}shapeComment")
    comment.text = f"stamp_{name}"

    return pic


def _find_and_stamp(root, name: str, stamp_ref: str, stamp_size: int) -> bool:
    """Find (인) in section XML and insert stamp pic overlapping it.

    Returns True if (인) was found and stamp inserted.
    """
    pic_id = 180200000 + abs(hash(name)) % 100000
    inst_id = pic_id - 1

    # Search all <hp:t> elements (covers paragraphs and tables)
    for t_elem in root.iter(f"{{{NS['hp']}}}t"):
        if t_elem.text and "(인)" in t_elem.text:
            run = t_elem.getparent()
            if run is None or not run.tag.endswith("}run"):
                continue

            text = t_elem.text
            in_pos = text.index("(인)")
            # Estimate char width ≈ 700 HWPUNIT for standard 10pt text
            char_width = 700
            horz_offset = in_pos * char_width - stamp_size // 6

            pic = _build_pic_element(stamp_size, stamp_ref, name, pic_id, inst_id)
            pos_el = pic.find(f"{{{NS['hp']}}}pos")
            if pos_el is not None:
                pos_el.set("horzOffset", str(horz_offset))

            # Add new run with pic after the (인) run
            new_run = etree.Element(f"{{{NS['hp']}}}run")
            new_run.set("charPrIDRef", run.get("charPrIDRef", "0"))
            new_run.append(pic)
            etree.SubElement(new_run, f"{{{NS['hp']}}}t")

            parent = run.getparent()
            run_index = list(parent).index(run)
            parent.insert(run_index + 1, new_run)

            print(f"  Found '(인)', stamp overlaid at horzOffset={horz_offset}")
            return True

    return False


def _add_in_text(root, name: str, signer_name: str | None) -> bool:
    """Add '(인)' text to the last paragraph that contains the signer name.

    If signer_name is given, appends ' (인)' to the paragraph containing it.
    Otherwise, appends ' (인)' to the last non-empty paragraph.
    Returns True if (인) was added.
    """
    target_run = None

    if signer_name:
        # Find paragraph containing the signer name
        for t_elem in root.iter(f"{{{NS['hp']}}}t"):
            if t_elem.text and signer_name in t_elem.text:
                target_run = t_elem.getparent()
                # Append (인) to the text
                t_elem.text = t_elem.text.rstrip() + " (인)"
                print(f"  Added '(인)' after '{signer_name}'")
                return True

    # Fallback: find the last non-empty paragraph's last run
    last_t = None
    for t_elem in root.iter(f"{{{NS['hp']}}}t"):
        if t_elem.text and t_elem.text.strip():
            last_t = t_elem

    if last_t is not None:
        last_t.text = last_t.text.rstrip() + " (인)"
        print(f"  Added '(인)' to last text: '{last_t.text}'")
        return True

    return False


def stamp_hwpx(name: str, input_path: Path, output_path: Path,
               stamp_size: int = DEFAULT_STAMP_SIZE,
               signer_name: str | None = None) -> None:
    """Apply a red stamp over (인) in an HWPX document.

    If (인) is not found in the document, it is automatically added
    next to the signer name (or the last non-empty paragraph), then
    the stamp is overlaid on top of it.
    """

    if not input_path.is_file():
        raise SystemExit(f"Input file not found: {input_path}")

    stamp_png = Path(tempfile.mktemp(suffix=".png"))
    try:
        generate_stamp(name, stamp_png, px=STAMP_PX)

        with tempfile.TemporaryDirectory() as tmpdir:
            work = Path(tmpdir) / "build"

            with ZipFile(input_path, "r") as zf:
                zf.extractall(work)

            # Copy stamp image to BinData
            bindata_dir = work / "BinData"
            bindata_dir.mkdir(exist_ok=True)
            stamp_fname = f"stamp_{name}.png"
            counter = 1
            while (bindata_dir / stamp_fname).exists():
                stamp_fname = f"stamp_{name}_{counter}.png"
                counter += 1
            shutil.copy2(stamp_png, bindata_dir / stamp_fname)
            stamp_ref = Path(stamp_fname).stem

            # Parse section0.xml
            section_path = work / "Contents" / "section0.xml"
            tree = etree.parse(str(section_path))
            root = tree.getroot()

            # Step 1: Check if (인) already exists
            has_in = False
            for t_elem in root.iter(f"{{{NS['hp']}}}t"):
                if t_elem.text and "(인)" in t_elem.text:
                    has_in = True
                    break

            # Step 2: If (인) not found, add it automatically
            if not has_in:
                print("  '(in)' not found - adding automatically")
                _add_in_text(root, name, signer_name)

            # Step 3: Find (인) and overlay stamp
            found = _find_and_stamp(root, name, stamp_ref, stamp_size)

            if not found:
                print("WARNING: Could not position stamp. '(인)' not found.",
                      file=sys.stderr)

            # Write updated section
            etree.indent(root, space="  ")
            tree.write(str(section_path), pretty_print=True,
                       xml_declaration=True, encoding="UTF-8")

            # Register stamp in content.hpf
            hpf_path = work / "Contents" / "content.hpf"
            if hpf_path.is_file():
                hpf_tree = etree.parse(str(hpf_path))
                hpf_root = hpf_tree.getroot()
                manifest = hpf_root.find(f".//{{{NS['opf']}}}manifest")
                if manifest is not None:
                    existing = manifest.find(f"opf:item[@id='{stamp_ref}']", NS)
                    if existing is None:
                        el = etree.SubElement(manifest, f"{{{NS['opf']}}}item")
                        el.set("id", stamp_ref)
                        el.set("href", f"BinData/{stamp_fname}")
                        el.set("media-type", "image/png")
                        el.set("isEmbeded", "1")
                etree.indent(hpf_root, space="  ")
                hpf_tree.write(str(hpf_path), pretty_print=True,
                               xml_declaration=True, encoding="UTF-8")

            # Pack HWPX
            mimetype_file = work / "mimetype"
            all_files = sorted(
                p.relative_to(work).as_posix()
                for p in work.rglob("*")
                if p.is_file()
            )

            with ZipFile(output_path, "w", ZIP_DEFLATED) as zf:
                zf.write(mimetype_file, "mimetype", compress_type=ZIP_STORED)
                for rel_path in all_files:
                    if rel_path == "mimetype":
                        continue
                    zf.write(work / rel_path, rel_path, compress_type=ZIP_DEFLATED)

        print(f"STAMPED: {output_path}")
        print(f"  Name: {name}")
        print(f"  Stamp size: {stamp_size} HWPUNIT")
        print(f"  Image: {stamp_fname}")

    finally:
        if stamp_png.exists():
            stamp_png.unlink()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate red stamp (도장) and overlay on (인) in HWPX"
    )
    parser.add_argument(
        "--name", "-n",
        required=True,
        help="Name to engrave on the stamp (e.g., 홍길동)",
    )
    parser.add_argument(
        "--input", "-i",
        type=Path,
        help="Input HWPX file (must contain '(인)' text)",
    )
    parser.add_argument(
        "--output", "-o",
        type=Path,
        required=True,
        help="Output file path (.hwpx or .png with --image-only)",
    )
    parser.add_argument(
        "--stamp-size",
        type=int,
        default=DEFAULT_STAMP_SIZE,
        help=f"Stamp size in HWPUNIT (default: {DEFAULT_STAMP_SIZE} ≈ 11.6mm)",
    )
    parser.add_argument(
        "--signer",
        help="Signer name/title to find in document for (인) placement "
             "(e.g., '학교장'). If omitted, (인) is added to last text.",
    )
    parser.add_argument(
        "--image-only",
        action="store_true",
        help="Only generate the stamp PNG image, do not modify HWPX",
    )
    args = parser.parse_args()

    if args.image_only:
        generate_stamp(args.name, args.output, px=STAMP_PX)
        print(f"STAMP IMAGE: {args.output}")
        print(f"  Name: {args.name}")
    else:
        if not args.input:
            raise SystemExit("--input is required when not using --image-only")
        stamp_hwpx(args.name, args.input, args.output, args.stamp_size,
                    signer_name=args.signer)


if __name__ == "__main__":
    main()
