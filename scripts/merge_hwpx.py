#!/usr/bin/env python3
"""Merge multiple HWPX documents into a single file.

Combines section content from multiple HWPX files while remapping style IDs
to avoid conflicts. The first file's styles serve as the base; subsequent
files' styles are appended with shifted IDs.

Handles:
  - charPr, paraPr, borderFill style ID remapping
  - BinData image deduplication and renaming
  - Paragraph/table element ID uniqueness
  - Page breaks between merged documents
  - Font merging from all source files

Usage:
    python merge_hwpx.py file1.hwpx file2.hwpx -o merged.hwpx
    python merge_hwpx.py file1.hwpx file2.hwpx file3.hwpx -o merged.hwpx --page-break
"""

import argparse
import copy
import hashlib
import shutil
import sys
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, BadZipFile, ZipFile

from lxml import etree

NS = {
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "opf": "http://www.idpf.org/2007/opf/",
}

# Attributes in section XML that reference style IDs
CHARPR_ATTRS = ["charPrIDRef"]
PARAPR_ATTRS = ["paraPrIDRef"]
BORDERFILL_ATTRS = ["borderFillIDRef"]
STYLE_ATTRS = ["styleIDRef"]
ID_ATTRS = ["id", "instid"]
BINDATA_ATTRS = ["binaryItemIDRef"]


def _file_hash(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


def _parse_hwpx(hwpx_path: Path) -> dict:
    """Extract header XML, section XML, and BinData from an HWPX file."""
    try:
        zf = ZipFile(hwpx_path, "r")
    except BadZipFile:
        raise SystemExit(f"Not a valid ZIP/HWPX: {hwpx_path}")

    result = {"path": hwpx_path, "bindata": {}, "bindata_bytes": {}}

    with zf:
        names = zf.namelist()

        # Header
        if "Contents/header.xml" not in names:
            raise SystemExit(f"Missing Contents/header.xml in {hwpx_path}")
        result["header"] = etree.fromstring(zf.read("Contents/header.xml"))

        # Section(s)
        sections = sorted(n for n in names if n.startswith("Contents/section") and n.endswith(".xml"))
        if not sections:
            raise SystemExit(f"No section XML found in {hwpx_path}")
        result["sections"] = []
        for sec_name in sections:
            result["sections"].append(etree.fromstring(zf.read(sec_name)))

        # BinData
        for name in names:
            if name.startswith("BinData/"):
                fname = name.split("/", 1)[1]
                if not fname:  # skip directory entries
                    continue
                data = zf.read(name)
                result["bindata"][fname] = data
                result["bindata_bytes"][fname] = _file_hash(data)

        # content.hpf for image references
        if "Contents/content.hpf" in names:
            result["content_hpf"] = etree.fromstring(zf.read("Contents/content.hpf"))

    return result


def _get_max_id(header_root, tag_path: str, ns: dict) -> int:
    """Get the maximum 'id' attribute value from a style collection."""
    max_id = 0
    for el in header_root.findall(tag_path, ns):
        try:
            val = int(el.get("id", "0"))
            if val > max_id:
                max_id = val
        except ValueError:
            pass
    return max_id


def _get_max_element_id(section_root) -> int:
    """Get maximum numeric id/instid from all elements in a section."""
    max_id = 0
    for el in section_root.iter():
        for attr in ID_ATTRS:
            val = el.get(attr, "")
            try:
                num = int(val)
                if num > max_id:
                    max_id = num
            except ValueError:
                pass
    return max_id


def _remap_section_ids(section_root, charpr_offset: int, parapr_offset: int,
                       borderfill_offset: int, style_offset: int,
                       element_id_offset: int, bindata_map: dict) -> None:
    """Remap all style/element ID references in a section XML tree."""
    for el in section_root.iter():
        # charPrIDRef
        for attr in CHARPR_ATTRS:
            val = el.get(attr)
            if val is not None:
                try:
                    old_id = int(val)
                    el.set(attr, str(old_id + charpr_offset))
                except ValueError:
                    pass

        # paraPrIDRef
        for attr in PARAPR_ATTRS:
            val = el.get(attr)
            if val is not None:
                try:
                    old_id = int(val)
                    el.set(attr, str(old_id + parapr_offset))
                except ValueError:
                    pass

        # borderFillIDRef
        for attr in BORDERFILL_ATTRS:
            val = el.get(attr)
            if val is not None:
                try:
                    old_id = int(val)
                    el.set(attr, str(old_id + borderfill_offset))
                except ValueError:
                    pass

        # styleIDRef
        for attr in STYLE_ATTRS:
            val = el.get(attr)
            if val is not None:
                try:
                    old_id = int(val)
                    el.set(attr, str(old_id + style_offset))
                except ValueError:
                    pass

        # element IDs (id, instid)
        for attr in ID_ATTRS:
            val = el.get(attr)
            if val is not None:
                try:
                    old_id = int(val)
                    el.set(attr, str(old_id + element_id_offset))
                except ValueError:
                    pass

        # BinData references
        for attr in BINDATA_ATTRS:
            val = el.get(attr)
            if val is not None and val in bindata_map:
                el.set(attr, bindata_map[val])


def _append_styles(base_header, new_header, ns: dict) -> dict:
    """Append styles from new_header to base_header, returning offset map."""
    offsets = {}

    # borderFills
    base_bf = base_header.find(".//hh:borderFills", ns)
    new_bf = new_header.find(".//hh:borderFills", ns)
    if base_bf is not None and new_bf is not None:
        max_id = _get_max_id(base_header, ".//hh:borderFill", ns)
        bf_offset = max_id  # new IDs start after max
        offsets["borderfill"] = bf_offset
        for bf in new_bf.findall("hh:borderFill", ns):
            new_el = copy.deepcopy(bf)
            old_id = int(new_el.get("id", "0"))
            new_el.set("id", str(old_id + bf_offset))
            base_bf.append(new_el)
        base_bf.set("itemCnt", str(len(base_bf.findall("hh:borderFill", ns))))
    else:
        offsets["borderfill"] = 0

    # charProperties
    base_cp = base_header.find(".//hh:charProperties", ns)
    new_cp = new_header.find(".//hh:charProperties", ns)
    if base_cp is not None and new_cp is not None:
        max_id = _get_max_id(base_header, ".//hh:charPr", ns)
        cp_offset = max_id + 1
        offsets["charpr"] = cp_offset
        for cp in new_cp.findall("hh:charPr", ns):
            new_el = copy.deepcopy(cp)
            old_id = int(new_el.get("id", "0"))
            new_el.set("id", str(old_id + cp_offset))
            # Remap borderFillIDRef inside charPr
            old_bf = new_el.get("borderFillIDRef")
            if old_bf is not None:
                try:
                    new_el.set("borderFillIDRef", str(int(old_bf) + offsets["borderfill"]))
                except ValueError:
                    pass
            base_cp.append(new_el)
        base_cp.set("itemCnt", str(len(base_cp.findall("hh:charPr", ns))))
    else:
        offsets["charpr"] = 0

    # tabProperties
    base_tp = base_header.find(".//hh:tabProperties", ns)
    new_tp = new_header.find(".//hh:tabProperties", ns)
    if base_tp is not None and new_tp is not None:
        max_id = _get_max_id(base_header, ".//hh:tabPr", ns)
        tp_offset = max_id + 1
        offsets["tabpr"] = tp_offset
        for tp in new_tp.findall("hh:tabPr", ns):
            new_el = copy.deepcopy(tp)
            old_id = int(new_el.get("id", "0"))
            new_el.set("id", str(old_id + tp_offset))
            base_tp.append(new_el)
        base_tp.set("itemCnt", str(len(base_tp.findall("hh:tabPr", ns))))
    else:
        offsets["tabpr"] = 0

    # paraPrProperties (paragraph styles)
    base_pp = base_header.find(".//hh:paraPrProperties", ns)
    new_pp = new_header.find(".//hh:paraPrProperties", ns)
    if base_pp is not None and new_pp is not None:
        max_id = _get_max_id(base_header, ".//hh:paraPr", ns)
        pp_offset = max_id + 1
        offsets["parapr"] = pp_offset
        for pp in new_pp.findall("hh:paraPr", ns):
            new_el = copy.deepcopy(pp)
            old_id = int(new_el.get("id", "0"))
            new_el.set("id", str(old_id + pp_offset))
            # Remap tabPrIDRef inside paraPr
            old_tab = new_el.get("tabPrIDRef")
            if old_tab is not None:
                try:
                    new_el.set("tabPrIDRef", str(int(old_tab) + offsets.get("tabpr", 0)))
                except ValueError:
                    pass
            base_pp.append(new_el)
        base_pp.set("itemCnt", str(len(base_pp.findall("hh:paraPr", ns))))
    else:
        offsets["parapr"] = 0

    # styles
    base_st = base_header.find(".//hh:styles", ns)
    new_st = new_header.find(".//hh:styles", ns)
    if base_st is not None and new_st is not None:
        max_id = 0
        for s in base_st.findall("hh:style", ns):
            try:
                val = int(s.get("id", "0"))
                if val > max_id:
                    max_id = val
            except ValueError:
                pass
        st_offset = max_id + 1
        offsets["style"] = st_offset
        for st in new_st.findall("hh:style", ns):
            new_el = copy.deepcopy(st)
            old_id = int(new_el.get("id", "0"))
            new_el.set("id", str(old_id + st_offset))
            # Remap charPrIDRef/paraPrIDRef inside style
            for ref_attr, off_key in [("charPrIDRef", "charpr"), ("paraPrIDRef", "parapr")]:
                old_ref = new_el.get(ref_attr)
                if old_ref is not None:
                    try:
                        new_el.set(ref_attr, str(int(old_ref) + offsets[off_key]))
                    except (ValueError, KeyError):
                        pass
            # Remap nextStyleIDRef
            old_next = new_el.get("nextStyleIDRef")
            if old_next is not None:
                try:
                    new_el.set("nextStyleIDRef", str(int(old_next) + st_offset))
                except ValueError:
                    pass
            base_st.append(new_el)
        base_st.set("itemCnt", str(len(base_st.findall("hh:style", ns))))
    else:
        offsets["style"] = 0

    # Merge fonts (add missing fonts from new_header)
    _merge_fonts(base_header, new_header, ns)

    return offsets


def _merge_fonts(base_header, new_header, ns: dict) -> None:
    """Add any new font faces from new_header into base_header."""
    base_ff = base_header.find(".//hh:fontfaces", ns)
    new_ff = new_header.find(".//hh:fontfaces", ns)
    if base_ff is None or new_ff is None:
        return

    for lang_attr in ["HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"]:
        base_face = None
        new_face = None
        for ff in base_ff.findall("hh:fontface", ns):
            if ff.get("lang") == lang_attr:
                base_face = ff
                break
        for ff in new_ff.findall("hh:fontface", ns):
            if ff.get("lang") == lang_attr:
                new_face = ff
                break

        if base_face is not None and new_face is not None:
            existing_names = set()
            for f in base_face.findall("hh:font", ns):
                existing_names.add(f.get("face", ""))

            max_font_id = 0
            for f in base_face.findall("hh:font", ns):
                try:
                    fid = int(f.get("id", "0"))
                    if fid > max_font_id:
                        max_font_id = fid
                except ValueError:
                    pass

            for f in new_face.findall("hh:font", ns):
                face_name = f.get("face", "")
                if face_name not in existing_names:
                    new_font = copy.deepcopy(f)
                    max_font_id += 1
                    new_font.set("id", str(max_font_id))
                    base_face.append(new_font)
                    existing_names.add(face_name)

            base_face.set("fontCnt", str(len(base_face.findall("hh:font", ns))))


def _make_page_break_paragraph(ns: dict) -> etree._Element:
    """Create a paragraph element with a page break."""
    p = etree.Element(f"{{{ns['hp']}}}p")
    p.set("id", "0")
    p.set("paraPrIDRef", "0")
    p.set("styleIDRef", "0")
    p.set("pageBreak", "1")
    p.set("columnBreak", "0")
    p.set("merged", "0")
    run = etree.SubElement(p, f"{{{ns['hp']}}}run")
    run.set("charPrIDRef", "0")
    t = etree.SubElement(run, f"{{{ns['hp']}}}t")
    t.text = None
    return p


def merge(input_files: list[Path], output_path: Path, page_break: bool = True) -> None:
    """Main merge logic."""
    if len(input_files) < 2:
        raise SystemExit("At least 2 HWPX files are required for merge.")

    # Parse all input files
    parsed = []
    for f in input_files:
        if not f.is_file():
            raise SystemExit(f"File not found: {f}")
        parsed.append(_parse_hwpx(f))

    # Use first file as base
    base = parsed[0]
    merged_header = copy.deepcopy(base["header"])
    merged_bindata: dict[str, bytes] = dict(base["bindata"])
    bindata_hash_map: dict[str, str] = dict(base["bindata_bytes"])  # hash -> filename

    # Collect all section body content
    # Each section's <hp:p> children are the content
    all_section_contents = []
    for sec in base["sections"]:
        all_section_contents.append(("base", sec, {}))

    # Track element ID offset
    global_max_elem_id = 0
    for sec in base["sections"]:
        mid = _get_max_element_id(sec)
        if mid > global_max_elem_id:
            global_max_elem_id = mid

    # Process additional files
    for idx, extra in enumerate(parsed[1:], start=1):
        # Merge styles into base header, get offsets
        offsets = _append_styles(merged_header, extra["header"], NS)

        # Handle BinData — deduplicate by content hash
        bindata_map = {}  # old_ref -> new_ref
        for fname, data in extra["bindata"].items():
            file_hash = _file_hash(data)
            # Check if identical content already exists
            existing = None
            for existing_name, existing_hash in bindata_hash_map.items():
                if existing_hash == file_hash:
                    existing = existing_name
                    break

            if existing:
                # Same content, reuse existing reference
                old_ref = Path(fname).stem
                new_ref = Path(existing).stem
                bindata_map[old_ref] = new_ref
            else:
                # New image — ensure unique filename
                new_fname = fname
                counter = 1
                while new_fname in merged_bindata:
                    stem = Path(fname).stem
                    suffix = Path(fname).suffix
                    new_fname = f"{stem}_{counter}{suffix}"
                    counter += 1

                merged_bindata[new_fname] = data
                bindata_hash_map[new_fname] = file_hash
                old_ref = Path(fname).stem
                new_ref = Path(new_fname).stem
                bindata_map[old_ref] = new_ref

        # Remap and collect sections
        for sec in extra["sections"]:
            sec_copy = copy.deepcopy(sec)
            element_id_offset = global_max_elem_id + 1000000 * idx
            _remap_section_ids(
                sec_copy,
                charpr_offset=offsets["charpr"],
                parapr_offset=offsets["parapr"],
                borderfill_offset=offsets["borderfill"],
                style_offset=offsets["style"],
                element_id_offset=element_id_offset,
                bindata_map=bindata_map,
            )
            all_section_contents.append(("extra", sec_copy, offsets))

            mid = _get_max_element_id(sec_copy)
            if mid > global_max_elem_id:
                global_max_elem_id = mid

    # Build merged section0.xml
    # Take the first section as the base structure, append content from others
    base_section = copy.deepcopy(all_section_contents[0][1])

    # Find the body content container — the <hp:p> elements at top level
    for source_type, sec, offsets in all_section_contents[1:]:
        if page_break:
            pb = _make_page_break_paragraph(NS)
            base_section.append(pb)

        # Append all <hp:p> elements from additional sections
        for child in list(sec):
            # Skip first <hp:p> that contains <hp:secPr> (page setup)
            has_secpr = child.find(f".//{{{NS['hp']}}}secPr") is not None
            if has_secpr:
                continue
            base_section.append(copy.deepcopy(child))

    # Now build the output HWPX
    # Use first file as the structural base
    with tempfile.TemporaryDirectory() as tmpdir:
        work = Path(tmpdir) / "build"

        # Extract first file as base structure
        with ZipFile(input_files[0], "r") as zf:
            zf.extractall(work)

        # Write merged header
        etree.indent(merged_header, space="  ")
        header_tree = etree.ElementTree(merged_header)
        header_tree.write(
            str(work / "Contents" / "header.xml"),
            pretty_print=True,
            xml_declaration=True,
            encoding="UTF-8",
        )

        # Write merged section
        etree.indent(base_section, space="  ")
        section_tree = etree.ElementTree(base_section)
        section_tree.write(
            str(work / "Contents" / "section0.xml"),
            pretty_print=True,
            xml_declaration=True,
            encoding="UTF-8",
        )

        # Write merged BinData — clear existing first to avoid stale files
        bindata_dir = work / "BinData"
        if bindata_dir.is_dir():
            shutil.rmtree(bindata_dir)
        bindata_dir.mkdir(exist_ok=True)
        for fname, data in merged_bindata.items():
            (bindata_dir / fname).write_bytes(data)

        # Update content.hpf — register all BinData images
        hpf_path = work / "Contents" / "content.hpf"
        if hpf_path.is_file():
            hpf_tree = etree.parse(str(hpf_path))
            hpf_root = hpf_tree.getroot()
            manifest = hpf_root.find(".//opf:manifest", NS)
            if manifest is not None:
                mime_map = {
                    ".jpg": "image/jpg", ".jpeg": "image/jpg",
                    ".png": "image/png", ".gif": "image/gif", ".bmp": "image/bmp",
                }
                for fname in merged_bindata:
                    item_id = Path(fname).stem
                    suffix = Path(fname).suffix.lower()
                    if suffix not in mime_map:
                        continue
                    existing = manifest.find(f"opf:item[@id='{item_id}']", NS)
                    if existing is None:
                        el = etree.SubElement(manifest, f"{{{NS['opf']}}}item")
                        el.set("id", item_id)
                        el.set("href", f"BinData/{fname}")
                        el.set("media-type", mime_map[suffix])
                        el.set("isEmbeded", "1")

            etree.indent(hpf_root, space="  ")
            hpf_tree.write(
                str(hpf_path),
                pretty_print=True,
                xml_declaration=True,
                encoding="UTF-8",
            )

        # Remove extra section files (section1.xml, etc.) if any
        contents_dir = work / "Contents"
        for f in contents_dir.glob("section*.xml"):
            if f.name != "section0.xml":
                f.unlink()

        # Pack as HWPX
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

    # Validate
    errors = _validate(output_path)
    if errors:
        print(f"WARNING: {output_path} has issues:", file=sys.stderr)
        for e in errors:
            print(f"  - {e}", file=sys.stderr)
    else:
        print(f"MERGED: {output_path}")
        print(f"  Sources: {len(input_files)} files")
        for f in input_files:
            print(f"    - {f.name}")
        print(f"  Page breaks: {'yes' if page_break else 'no'}")


def _validate(hwpx_path: Path) -> list[str]:
    """Quick structural validation."""
    errors = []
    required = ["mimetype", "Contents/content.hpf", "Contents/header.xml", "Contents/section0.xml"]

    try:
        zf = ZipFile(hwpx_path, "r")
    except BadZipFile:
        return [f"Not a valid ZIP: {hwpx_path}"]

    with zf:
        names = zf.namelist()
        for r in required:
            if r not in names:
                errors.append(f"Missing: {r}")

        if "mimetype" in names:
            content = zf.read("mimetype").decode("utf-8").strip()
            if content != "application/hwp+zip":
                errors.append(f"Bad mimetype: {content}")
            if names[0] != "mimetype":
                errors.append("mimetype not first entry")
            info = zf.getinfo("mimetype")
            if info.compress_type != ZIP_STORED:
                errors.append("mimetype not ZIP_STORED")

        for name in names:
            if name.endswith(".xml") or name.endswith(".hpf"):
                try:
                    etree.fromstring(zf.read(name))
                except etree.XMLSyntaxError as e:
                    errors.append(f"Malformed XML: {name}: {e}")

    return errors


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge multiple HWPX documents into one"
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="Input HWPX files to merge (order matters: first = base)",
    )
    parser.add_argument(
        "--output", "-o",
        type=Path,
        required=True,
        help="Output merged .hwpx file path",
    )
    parser.add_argument(
        "--no-page-break",
        action="store_true",
        default=False,
        help="Do not insert page breaks between merged documents",
    )
    args = parser.parse_args()

    merge(
        input_files=args.inputs,
        output_path=args.output,
        page_break=not args.no_page_break,
    )


if __name__ == "__main__":
    main()
