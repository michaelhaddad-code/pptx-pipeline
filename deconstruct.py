"""
Step 1: DECONSTRUCT
Extracts all components from a .pptx file into a structured component library.
Usage: python deconstruct.py <input.pptx> [--library LIBRARY_DIR] [--force]
"""

import argparse
import zipfile
import os
import re
import shutil
import json
import sys
import logging
from datetime import datetime
from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)


def deconstruct(pptx_path: str, library_path: str = "component_library", force: bool = False):
    logger.info("Deconstructing: %s", pptx_path)
    logger.info("Output library: %s", library_path)

    # Clean and recreate library
    if os.path.exists(library_path):
        if not force:
            response = input(f"Library '{library_path}' already exists. Overwrite? [y/N]: ").strip().lower()
            if response != "y":
                print("Aborted.")
                return
        # Create timestamped backup before deletion
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{library_path}_backup_{timestamp}"
        shutil.copytree(library_path, backup_path)
        logger.info("  Backup created: %s", backup_path)
        shutil.rmtree(library_path)
    os.makedirs(library_path)

    with zipfile.ZipFile(pptx_path, "r") as z:
        all_files = z.namelist()

        # ── 1. Extract everything raw (preserves full fidelity) ──────────────
        raw_dir = os.path.join(library_path, "_raw")
        # Safe extraction — prevent zip slip (path traversal)
        for member in z.infolist():
            target_path = os.path.realpath(os.path.join(raw_dir, member.filename))
            if not target_path.startswith(os.path.realpath(raw_dir)):
                raise ValueError(f"Zip slip detected: {member.filename} escapes {raw_dir}")
            z.extract(member, raw_dir)
        logger.info("  ok Raw extraction: %d files -> %s", len(all_files), raw_dir)

        # ── 2. Theme ──────────────────────────────────────────────────────────
        theme_dir = os.path.join(library_path, "theme")
        os.makedirs(theme_dir)
        for f in all_files:
            if f.startswith("ppt/theme/"):
                dest = os.path.join(theme_dir, os.path.basename(f))
                with z.open(f) as src, open(dest, "wb") as dst:
                    dst.write(src.read())
        logger.info("  ok Theme files extracted")

        # ── 3. Slide master & layouts ─────────────────────────────────────────
        master_dir = os.path.join(library_path, "slide_master")
        os.makedirs(master_dir)
        for f in all_files:
            if f.startswith("ppt/slideMasters/") or f.startswith("ppt/slideLayouts/"):
                rel_path = f.replace("ppt/slideMasters/", "slideMasters/").replace("ppt/slideLayouts/", "slideLayouts/")
                dest = os.path.join(master_dir, rel_path)
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with z.open(f) as src, open(dest, "wb") as dst:
                    dst.write(src.read())
        logger.info("  ok Slide master & layouts extracted")

        # ── 4. Media assets ───────────────────────────────────────────────────
        media_dir = os.path.join(library_path, "media")
        os.makedirs(media_dir)
        media_files = [f for f in all_files if f.startswith("ppt/media/")]
        for f in media_files:
            dest = os.path.join(media_dir, os.path.basename(f))
            with z.open(f) as src, open(dest, "wb") as dst:
                dst.write(src.read())
        logger.info("  ok Media assets: %d files", len(media_files))

        # ── 5. Per-slide extraction ───────────────────────────────────────────
        slide_files = sorted([f for f in all_files if f.startswith("ppt/slides/slide") and f.endswith(".xml")])
        slides_meta = []

        for slide_file in slide_files:
            slide_num = int(re.search(r'slide(\d+)\.xml$', slide_file).group(1))
            slide_dir = os.path.join(library_path, "slides", f"slide_{slide_num}")
            os.makedirs(slide_dir, exist_ok=True)

            # Raw slide XML
            slide_xml_bytes = z.read(slide_file)
            slide_xml_path = os.path.join(slide_dir, "slide.xml")
            with open(slide_xml_path, "wb") as f:
                f.write(slide_xml_bytes)

            # Relationships file
            rels_file = f"ppt/slides/_rels/slide{slide_num}.xml.rels"
            rels_data = {}
            if rels_file in all_files:
                rels_xml_bytes = z.read(rels_file)
                with open(os.path.join(slide_dir, "slide.xml.rels"), "wb") as f:
                    f.write(rels_xml_bytes)
                # Parse rels to find which media this slide uses
                rels_root = ET.fromstring(rels_xml_bytes)
                for rel in rels_root:
                    rid = rel.get("Id")
                    target = rel.get("Target", "")
                    rtype = rel.get("Type", "").split("/")[-1]
                    rels_data[rid] = {"type": rtype, "target": target}

            # Parse slide XML for shape metadata
            ns = {
                "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            }
            root = ET.fromstring(slide_xml_bytes)
            shapes = []

            def _extract_geometry(elem):
                """Extract position and size from a:xfrm or p:xfrm."""
                # Shapes use a:xfrm inside p:spPr; graphicFrames use p:xfrm directly
                xfrm = elem.find(".//p:spPr/a:xfrm", ns)
                if xfrm is None:
                    xfrm = elem.find("p:xfrm", ns)
                if xfrm is None:
                    return None
                off = xfrm.find("a:off", ns)
                ext = xfrm.find("a:ext", ns)
                if off is None or ext is None:
                    return None
                return {
                    "x": int(off.get("x", 0)),
                    "y": int(off.get("y", 0)),
                    "cx": int(ext.get("cx", 0)),
                    "cy": int(ext.get("cy", 0)),
                }

            def _extract_table_grid(elem):
                """Extract column widths, row heights, and per-row font sizes from a table."""
                tbl = elem.find(".//a:tbl", ns)
                if tbl is None:
                    return None
                grid_cols = tbl.findall("a:tblGrid/a:gridCol", ns)
                col_widths = [int(gc.get("w", 0)) for gc in grid_cols]
                rows = tbl.findall("a:tr", ns)
                row_heights = [int(tr.get("h", 0)) for tr in rows]

                # Extract font sizes per row (to detect header/data/summary fonts)
                row_fonts = []
                for tr in rows:
                    sizes = set()
                    for rpr in tr.findall(".//a:rPr", ns):
                        sz = rpr.get("sz")
                        if sz:
                            sizes.add(int(sz))
                    for def_rpr in tr.findall(".//a:pPr/a:defRPr", ns):
                        sz = def_rpr.get("sz")
                        if sz:
                            sizes.add(int(sz))
                    row_fonts.append(sorted(sizes) if sizes else [])

                return {
                    "columns": len(col_widths),
                    "rows": len(row_heights),
                    "col_widths": col_widths,
                    "row_heights": row_heights,
                    "row_fonts": row_fonts,
                }

            def _extract_shape(elem, parent_group=None):
                """Extract shape info from an element, recursing into groups."""
                tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                shape_info = {"type": tag}

                # Shape name/id
                cnvpr = elem.find(".//p:cNvPr", ns)
                if cnvpr is None:
                    cnvpr = elem.find(".//a:cNvPr", ns)
                if cnvpr is not None:
                    shape_info["id"] = cnvpr.get("id")
                    shape_info["name"] = cnvpr.get("name")

                # Geometry (position + size)
                geo = _extract_geometry(elem)
                if geo:
                    shape_info["geometry"] = geo

                # Table grid info
                if tag == "graphicFrame":
                    table_grid = _extract_table_grid(elem)
                    if table_grid:
                        shape_info["table_grid"] = table_grid

                # Image embed reference
                if tag == "pic":
                    blip = elem.find(".//a:blip", ns)
                    if blip is not None:
                        embed = blip.get(f"{{{ns['r']}}}embed")
                        if embed:
                            shape_info["image_rid"] = embed

                # Extract font sizes from text runs
                font_sizes_found = set()
                for rpr in elem.findall(".//a:rPr", ns):
                    sz = rpr.get("sz")
                    if sz:
                        font_sizes_found.add(int(sz))
                # Also check default paragraph run properties
                for def_rpr in elem.findall(".//a:pPr/a:defRPr", ns):
                    sz = def_rpr.get("sz")
                    if sz:
                        font_sizes_found.add(int(sz))
                if font_sizes_found:
                    shape_info["font_sizes"] = sorted(font_sizes_found)

                # Extract text content if any
                texts = []
                for t in elem.findall(".//a:t", ns):
                    if t.text:
                        texts.append(t.text)
                if texts:
                    shape_info["text_preview"] = " ".join(texts)[:100]

                if parent_group is not None:
                    shape_info["parent_group"] = parent_group

                shapes.append(shape_info)

                # Recurse into group shapes to enumerate children
                if tag == "grpSp":
                    group_id = shape_info.get("id", "unknown")
                    for child in elem:
                        child_tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                        if child_tag in ("sp", "pic", "grpSp", "graphicFrame", "cxnSp"):
                            _extract_shape(child, parent_group=group_id)

            sp_tree = root.find(".//p:spTree", ns)
            if sp_tree is not None:
                for child in sp_tree:
                    _extract_shape(child)

            # Build slide metadata
            slide_meta = {
                "slide_number": slide_num,
                "slide_xml": f"slides/slide_{slide_num}/slide.xml",
                "rels_xml": f"slides/slide_{slide_num}/slide.xml.rels",
                "relationships": rels_data,
                "shape_count": len(shapes),
                "shapes": shapes,
            }
            slides_meta.append(slide_meta)

            # Save per-slide metadata
            with open(os.path.join(slide_dir, "metadata.json"), "w") as f:
                json.dump(slide_meta, f, indent=2)

            logger.info("  ok Slide %d: %d shapes, %d relationships", slide_num, len(shapes), len(rels_data))

        # ── 6. Global manifest ────────────────────────────────────────────────
        manifest = {
            "source": os.path.basename(pptx_path),
            "total_slides": len(slide_files),
            "total_media": len(media_files),
            "structure": {
                "raw": "_raw/",
                "theme": "theme/",
                "slide_master": "slide_master/",
                "media": "media/",
                "slides": "slides/slide_N/",
            },
            "slides": slides_meta,
        }

        with open(os.path.join(library_path, "manifest.json"), "w") as f:
            json.dump(manifest, f, indent=2)

        logger.info("\ndone Deconstruction complete.")
        logger.info("   Slides: %d", len(slide_files))
        logger.info("   Media:  %d", len(media_files))
        logger.info("   Library: %s/", library_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Deconstruct a .pptx file into a structured component library.")
    parser.add_argument("pptx", nargs="?", default="Slides_Examples.pptx", help="Path to the input .pptx file")
    parser.add_argument("--library", default="component_library", help="Output library directory (default: component_library)")
    parser.add_argument("--force", action="store_true", help="Overwrite existing library without prompting")
    args = parser.parse_args()
    deconstruct(args.pptx, args.library, args.force)
