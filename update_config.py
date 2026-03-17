"""
Step 3 (Row 2): UPDATE CONFIG
Loads new data from the data/ folder into the existing config.json,
resolving token mappings with actual values ready for injection.
Usage: python update_config.py [-c CONFIG] [-d DATA_DIR]
"""

import os
import json
import csv
import sys
import re
import glob
import argparse
import logging
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from layout import compute_table_layout, compute_image_fit, compute_text_font_scale, read_image_dimensions

logger = logging.getLogger(__name__)


def _load_xlsx(xlsx_path: str) -> list:
    """Load an .xlsx file using stdlib only (zipfile + XML parsing).

    Returns a list of row dicts (like csv.DictReader), using the first row
    as column headers. Only reads the first sheet.
    """
    _SS_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    with zipfile.ZipFile(xlsx_path, "r") as z:
        # Load shared strings table (cells reference strings by index)
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_xml = z.read("xl/sharedStrings.xml")
            ss_root = ET.fromstring(ss_xml)
            for si in ss_root.findall(f"{{{_SS_NS}}}si"):
                # <si> can contain <t> directly or <r><t> runs
                texts = []
                for t in si.iter(f"{{{_SS_NS}}}t"):
                    if t.text:
                        texts.append(t.text)
                shared_strings.append("".join(texts))

        # Find the first sheet's path from workbook.xml.rels
        sheet_path = None
        if "xl/workbook.xml" in z.namelist():
            # Read rels to find sheet1
            rels_path = "xl/_rels/workbook.xml.rels"
            if rels_path in z.namelist():
                rels_root = ET.fromstring(z.read(rels_path))
                for rel in rels_root:
                    target = rel.get("Target", "")
                    if "worksheets/sheet1" in target or "worksheets/sheet" in target:
                        sheet_path = "xl/" + target.lstrip("/")
                        break

        if not sheet_path:
            # Fallback: try common path
            for candidate in ["xl/worksheets/sheet1.xml", "xl/worksheets/sheet.xml"]:
                if candidate in z.namelist():
                    sheet_path = candidate
                    break

        if not sheet_path:
            return []

        sheet_xml = z.read(sheet_path)
        sheet_root = ET.fromstring(sheet_xml)

    # Parse rows
    rows_data = []
    for row in sheet_root.iter(f"{{{_SS_NS}}}row"):
        cells = []
        for cell in row.findall(f"{{{_SS_NS}}}c"):
            cell_type = cell.get("t", "")
            value_elem = cell.find(f"{{{_SS_NS}}}v")
            value = value_elem.text if value_elem is not None else ""

            if cell_type == "s" and value:
                # Shared string reference
                idx = int(value)
                value = shared_strings[idx] if idx < len(shared_strings) else ""
            elif cell_type == "inlineStr":
                # Inline string
                is_elem = cell.find(f"{{{_SS_NS}}}is")
                if is_elem is not None:
                    texts = [t.text or "" for t in is_elem.iter(f"{{{_SS_NS}}}t")]
                    value = "".join(texts)

            cells.append(value or "")
        rows_data.append(cells)

    if len(rows_data) < 2:
        return []

    headers = rows_data[0]
    result = []
    for row_cells in rows_data[1:]:
        row_dict = {}
        for i, header in enumerate(headers):
            if header:
                row_dict[header] = row_cells[i] if i < len(row_cells) else ""
        if any(v for v in row_dict.values()):
            result.append(row_dict)

    return result


def find_screenshots(search_dir: str, match_key: str) -> str:
    """Find a screenshot image matching a key in a directory with date-stamped subfolders.

    Ported from Lara's _find_screenshot(): scans subfolders newest-first,
    matches filenames ending with _<KEY>.png (case-insensitive, spaces -> underscores).

    Also supports flat directory (no subfolders) and .jpg/.jpeg files.

    Args:
        search_dir: directory to search (may contain date-stamped subfolders)
        match_key: the identifier to match (e.g. "US PRO")

    Returns:
        str path to the matching image, or None
    """
    search_path = Path(search_dir)
    if not search_path.exists():
        return None

    key_normalized = match_key.strip().replace(" ", "_").upper()
    image_exts = {".png", ".jpg", ".jpeg"}

    # Try date-stamped subfolders first (newest first)
    subdirs = sorted(
        [d for d in search_path.iterdir() if d.is_dir()],
        reverse=True,
    )

    for folder in subdirs:
        for img in folder.iterdir():
            if img.suffix.lower() in image_exts:
                if img.stem.upper().endswith(f"_{key_normalized}"):
                    return str(img)

    # Fallback: flat directory
    for img in search_path.iterdir():
        if img.is_file() and img.suffix.lower() in image_exts:
            if img.stem.upper().endswith(f"_{key_normalized}"):
                return str(img)

    return None


def load_data_sources(data_dir: str) -> dict:
    """
    Loads all data files from the data/ folder into a unified key-value store.
    Supports: .csv, .json, .txt, .md

    The returned dict contains a special key '_load_warnings' (list of str)
    with any non-fatal issues encountered during loading (e.g. encoding failures).
    Callers can check this list to detect problems programmatically.
    """
    data = {}
    load_warnings = []

    if not os.path.exists(data_dir):
        logger.error("ERROR Data directory not found: %s", data_dir)
        sys.exit(1)

    # ── CSV files ─────────────────────────────────────────────────
    for csv_path in glob.glob(os.path.join(data_dir, "*.csv")):
        # Excel on Windows exports CSV as cp1252 by default;
        # try utf-8-sig (handles BOM), utf-8, cp1252, latin-1 in order.
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                with open(csv_path, newline='', encoding=enc) as f:
                    reader = csv.DictReader(f)
                    rows = list(reader)
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        else:
            msg = f"Could not decode CSV with any known encoding: {os.path.basename(csv_path)}"
            logger.warning("WARNING %s", msg)
            load_warnings.append(msg)
            continue
        # Support two formats:
        # Format A: field,value (key-value pairs) — requires exact column names
        # Format B: any tabular CSV (loaded as list)
        cols = set(rows[0].keys()) if rows else set()
        is_kv = False
        if rows and cols == {"field", "value"} and len(cols) == 2:
            # Extra guard: only treat as key-value if "field" column has unique,
            # non-empty string keys (avoids misidentifying regular 2-column data)
            fields = [row["field"] for row in rows]
            if all(fields) and len(fields) == len(set(fields)):
                for row in rows:
                    data[row["field"]] = row["value"]
                logger.info("  ok Loaded key-value CSV: %s (%d fields)", os.path.basename(csv_path), len(rows))
                is_kv = True
        if not is_kv:
            key = Path(csv_path).stem
            data[key] = rows
            logger.info("  ok Loaded tabular CSV: %s (%d rows)", os.path.basename(csv_path), len(rows))

    # ── JSON files ────────────────────────────────────────────────
    for json_path in glob.glob(os.path.join(data_dir, "*.json")):
        with open(json_path, encoding='utf-8') as f:
            content = json.load(f)
        if isinstance(content, dict):
            data.update(content)
        else:
            key = Path(json_path).stem
            data[key] = content
        logger.info("  ok Loaded JSON: %s", os.path.basename(json_path))

    # ── Text / Markdown files ─────────────────────────────────────
    for txt_path in glob.glob(os.path.join(data_dir, "*.txt")) + glob.glob(os.path.join(data_dir, "*.md")):
        key = Path(txt_path).stem
        with open(txt_path, encoding='utf-8') as f:
            data[key] = f.read().strip()
        logger.info("  ok Loaded text: %s", os.path.basename(txt_path))

    # ── Excel files (.xlsx) ──────────────────────────────────────
    for xlsx_path in glob.glob(os.path.join(data_dir, "*.xlsx")):
        try:
            rows = _load_xlsx(xlsx_path)
            if not rows:
                continue
            # Same logic as CSV: detect key-value vs tabular
            cols = set(rows[0].keys()) if rows else set()
            is_kv = False
            if rows and cols == {"field", "value"} and len(cols) == 2:
                fields = [row["field"] for row in rows]
                if all(fields) and len(fields) == len(set(fields)):
                    for row in rows:
                        data[row["field"]] = row["value"]
                    logger.info("  ok Loaded key-value XLSX: %s (%d fields)", os.path.basename(xlsx_path), len(rows))
                    is_kv = True
            if not is_kv:
                key = Path(xlsx_path).stem
                data[key] = rows
                logger.info("  ok Loaded tabular XLSX: %s (%d rows)", os.path.basename(xlsx_path), len(rows))
        except Exception as e:
            msg = f"Failed to load XLSX: {os.path.basename(xlsx_path)}: {e}"
            logger.warning("WARNING %s", msg)
            load_warnings.append(msg)

    # ── Images (just register their paths) ───────────────────────
    visuals_dir = os.path.join(data_dir, "visuals")
    if os.path.exists(visuals_dir):
        for img_path in glob.glob(os.path.join(visuals_dir, "*")):
            key = Path(img_path).stem
            data[f"visual:{key}"] = img_path
        logger.info("  ok Registered visuals: %d files", len(glob.glob(os.path.join(visuals_dir, '*'))))

    # ── Screenshot directories (scan for matching images) ────────
    screenshots_dir = os.path.join(data_dir, "screenshots")
    if os.path.exists(screenshots_dir):
        # Each subdirectory represents a source category (e.g. "outlook", "fiscal")
        for source_dir in Path(screenshots_dir).iterdir():
            if source_dir.is_dir():
                # Register as a screenshot source for find_screenshots()
                data[f"screenshot_dir:{source_dir.name}"] = str(source_dir)
                logger.info("  ok Registered screenshot source: %s", source_dir.name)

    if load_warnings:
        data["_load_warnings"] = load_warnings
    return data


def _resolve_nested(field: str, data: dict):
    """
    Walk *data* according to a dotted/indexed path and return the value.

    Supported syntax examples:
        customer.name        -> data["customer"]["name"]
        items[0]             -> data["items"][0]
        customers[2].name    -> data["customers"][2]["name"]

    Raises KeyError / IndexError / TypeError on any lookup failure.
    """
    # Split on '.' but keep array indices attached to their segment.
    # e.g. "customers[2].name" -> ["customers[2]", "name"]
    segments = field.split(".")
    # Regex to pull an optional index off a segment: "key[0]" -> ("key", "0")
    _idx_re = re.compile(r'^([^\[]+)\[(\d+)\]$')

    current = data
    for seg in segments:
        m = _idx_re.match(seg)
        if m:
            key, idx = m.group(1), int(m.group(2))
            current = current[key]     # dict / list lookup for the key part
            current = current[idx]     # integer index
        else:
            current = current[seg]     # plain dict key
    return current


def resolve_field(field: str, data: dict) -> dict:
    """
    Resolve a data field name to its value from the data store.
    Returns a dict with 'value' and '_resolved' flag.

    Supports nested access via dot notation and array indexing:
        customer.name        -> data["customer"]["name"]
        items[0]             -> data["items"][0]
        customers[2].name    -> data["customers"][2]["name"]

    Simple flat keys are tried first for backward compatibility.
    """
    if not field:
        return {"value": None, "_resolved": False}

    # 1. Flat lookup
    if field in data:
        val = data[field]
        # For lists/dicts (e.g. table data), return as JSON string
        if isinstance(val, (list, dict)):
            return {"value": json.dumps(val), "_resolved": True}
        return {"value": str(val), "_resolved": True}

    # 2. Nested / indexed lookup
    try:
        value = _resolve_nested(field, data)
        if isinstance(value, (list, dict)):
            return {"value": json.dumps(value), "_resolved": True}
        return {"value": str(value), "_resolved": True}
    except (KeyError, IndexError, TypeError):
        logger.warning("WARNING Field not found in data: '%s'", field)
        return {"value": None, "_resolved": False}


def resolve_token(token_ref: str, data: dict) -> dict:
    """
    Legacy: Resolves a token reference like {{field_name}} to its actual value.
    Prefer resolve_field() for new code.
    """
    if token_ref.startswith("{{") and token_ref.endswith("}}"):
        field = token_ref[2:-2].strip()
        result = resolve_field(field, data)
        if not result["_resolved"]:
            # Preserve original token string for backward compatibility
            result["value"] = token_ref
        return result

    return {"value": token_ref, "_resolved": True}


def update_config(config_path: str, data_dir: str):
    logger.info("Updating config: %s", config_path)
    logger.info("Data directory: %s", data_dir)

    # Load config
    if not os.path.exists(config_path):
        logger.error("ERROR Config not found: %s. Run generate_config.py first.", config_path)
        sys.exit(1)

    with open(config_path, encoding='utf-8') as f:
        config = json.load(f)

    # Load all data sources
    logger.info("\nLoading data sources...")
    data = load_data_sources(data_dir)
    logger.info("  Total fields loaded: %d", len(data))

    # Process each slide and shape
    logger.info("\nResolving tokens...")
    updates = 0
    warnings = 0

    for slide_key, slide in config["slides"].items():
        for shape in slide["shapes"]:
            if not shape.get("is_dynamic"):
                continue

            data_field = shape.get("data_field", "")
            if not data_field:
                # Legacy fallback: check for old tokens dict
                tokens = shape.get("tokens", {})
                if tokens:
                    resolved_tokens = {}
                    for token, field_ref in tokens.items():
                        result = resolve_token(field_ref, data)
                        resolved_tokens[token] = result
                        if result["_resolved"]:
                            updates += 1
                        else:
                            warnings += 1
                    shape["resolved_tokens"] = resolved_tokens
                continue

            result = resolve_field(data_field, data)
            if result["_resolved"]:
                shape["resolved_value"] = result["value"]
                updates += 1
                logger.info("  ok %s -> %s: '%s' -> '%s'", slide_key, shape['shape_name'], data_field, result['value'][:80] if result['value'] else "")
            else:
                warnings += 1

        # Handle dynamic images
        for image in slide.get("images", []):
            if not image.get("is_dynamic"):
                continue

            source = image.get("source", "")
            resolved = None

            if source.startswith("visual:"):
                # Direct visual reference: "visual:chart_name"
                if source in data:
                    resolved = data[source]
                else:
                    # Try without prefix
                    visual_key = source
                    stem = Path(source.replace("visual:", "")).stem
                    visual_key = f"visual:{stem}"
                    if visual_key in data:
                        resolved = data[visual_key]

            elif source.startswith("screenshot:"):
                # Screenshot directory scan: "screenshot:outlook/US PRO"
                # Format: screenshot:<dir_name>/<match_key>
                parts = source.replace("screenshot:", "").split("/", 1)
                if len(parts) == 2:
                    dir_name, match_key = parts
                    dir_key = f"screenshot_dir:{dir_name}"
                    if dir_key in data:
                        found = find_screenshots(data[dir_key], match_key)
                        if found:
                            resolved = found
                        else:
                            logger.warning("  WARNING No screenshot found for '%s' in %s", match_key, dir_name)

            elif source:
                # Legacy: bare filename -> look up as visual
                image_key = f"visual:{Path(source).stem}"
                if image_key in data:
                    resolved = data[image_key]

            if resolved:
                image["resolved_source"] = resolved
                updates += 1
                logger.info("  ok %s -> image %s: resolved to %s", slide_key, image.get('rid', '?'), resolved)

    # ── Post-resolution: compute layout for dynamic tables ──
    for slide_key, slide in config["slides"].items():
        for shape in slide["shapes"]:
            layout = shape.get("layout")
            if not layout or layout.get("type") != "dynamic_table":
                continue
            if not shape.get("is_dynamic"):
                continue

            resolved_value = shape.get("resolved_value")
            if not resolved_value:
                continue

            try:
                rows_data = json.loads(resolved_value)
            except (json.JSONDecodeError, TypeError):
                continue

            if not isinstance(rows_data, list):
                continue

            geometry = shape.get("geometry", {})
            if not geometry:
                continue

            n_data_rows = len(rows_data)
            table_grid = shape.get("table_grid", {})
            # header_rows defaults to 1, summary row is counted by compute_table_layout
            computed = compute_table_layout(geometry, layout, n_data_rows)
            layout["_computed"] = computed
            logger.info("  ok %s -> table %s: %d data rows, row_h=%d, fonts=%s%s",
                        slide_key, shape["shape_name"], n_data_rows,
                        computed["row_height"], computed["font_sizes"],
                        " [single-row]" if computed["single_row"] else "")

    # ── Post-resolution: compute layout for dynamic text shapes ──
    for slide_key, slide in config["slides"].items():
        for shape in slide["shapes"]:
            layout = shape.get("layout")
            if not layout or layout.get("type") != "auto_fit_text":
                continue
            if not shape.get("is_dynamic"):
                continue

            resolved_value = shape.get("resolved_value")
            if not resolved_value:
                continue

            geometry = shape.get("geometry", {})
            if not geometry.get("cx") or not geometry.get("cy"):
                continue

            min_font = layout.get("min_font_size", 600)
            max_font = layout.get("max_font_size", 2400)

            # Read the original font size from the shape XML if available,
            # otherwise use max_font as the starting point
            original_font = max_font

            computed_sz = compute_text_font_scale(
                resolved_value,
                geometry["cx"], geometry["cy"],
                original_font, min_font, max_font,
            )
            layout["_computed"] = {"font_size": computed_sz}

            if computed_sz < original_font:
                logger.info("  ok %s -> %s: text auto-fit %d -> %d (%.0f%% of max)",
                            slide_key, shape["shape_name"],
                            original_font, computed_sz,
                            computed_sz / original_font * 100)

    # ── Post-resolution: compute layout for dynamic images ──
    for slide_key, slide in config["slides"].items():
        for image in slide.get("images", []):
            if not image.get("is_dynamic") or not image.get("resolved_source"):
                continue

            source_path = image["resolved_source"]
            if not os.path.exists(source_path):
                continue

            # Find the matching shape's layout config
            target_shape_id = image.get("target_shape_id")
            layout = image.get("layout")

            # If no layout on the image entry, try to pull from the matching shape
            if not layout and target_shape_id:
                for shape in slide["shapes"]:
                    if shape.get("shape_id") == target_shape_id and shape.get("layout", {}).get("type") == "dynamic_image":
                        layout = shape["layout"]
                        break

            if not layout or layout.get("type") != "dynamic_image":
                continue

            max_cx = layout.get("max_cx", 0)
            max_cy = layout.get("max_cy", 0)
            if max_cx <= 0 or max_cy <= 0:
                continue

            dims = read_image_dimensions(source_path)
            if dims is None:
                logger.warning("  WARNING Could not read dimensions from: %s", source_path)
                continue

            img_w, img_h = dims
            fit_mode = layout.get("fit", "stretch")
            anchor = layout.get("anchor", "center")

            computed = compute_image_fit(img_w, img_h, max_cx, max_cy, fit=fit_mode, anchor=anchor)
            image["_computed"] = computed
            logger.info("  ok %s -> image %s: %dx%d px -> cx=%d cy=%d (offset %d,%d) [%s]",
                        slide_key, image.get("rid", "?"), img_w, img_h,
                        computed["cx"], computed["cy"],
                        computed["offset_x"], computed["offset_y"], fit_mode)

    # Save updated config
    with open(config_path, "w", encoding='utf-8') as f:
        json.dump(config, f, indent=2)

    logger.info("\ndone Config updated.")
    logger.info("   Tokens resolved: %d", updates)
    if warnings:
        logger.warning("   WARNING Unresolved fields: %d (check field names in data file)", warnings)
    logger.info("   Saved: %s", config_path)

    return config


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Update config.json with resolved data from the data/ folder."
    )
    parser.add_argument(
        "-c", "--config",
        default="configs/slides_examples.json",
        help="Path to the config JSON file (default: configs/slides_examples.json)",
    )
    parser.add_argument(
        "-d", "--data-dir",
        default="data",
        help="Path to the data directory (default: data)",
    )
    args = parser.parse_args()
    update_config(args.config, args.data_dir)
