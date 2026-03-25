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

from layout import compute_table_layout, compute_image_fit, compute_text_font_scale, read_image_dimensions, read_image_dpi

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
                        clean = target.lstrip("/")
                        if clean.startswith("xl/"):
                            sheet_path = clean
                        else:
                            sheet_path = "xl/" + clean
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


def apply_mappings(config: dict, mappings_path: str) -> int:
    """Apply mappings from mappings.json into the config.

    First resets all mapping fields (is_dynamic, data_field, source) so that
    mappings.json is the sole source of truth. Then applies each mapping entry.

    Args:
        config: the loaded config dict (modified in place)
        mappings_path: path to mappings.json

    Returns:
        number of mappings applied
    """
    if not os.path.exists(mappings_path):
        logger.info("No mappings file found at %s, skipping.", mappings_path)
        return 0

    with open(mappings_path, encoding="utf-8") as f:
        mappings_data = json.load(f)

    mapping_list = mappings_data.get("mappings", [])
    if not mapping_list:
        logger.info("Mappings file is empty, skipping.")
        return 0

    # ── Reset all mapping fields so mappings.json is sole source of truth ──
    for slide_key, slide in config["slides"].items():
        for shape in slide["shapes"]:
            shape["is_dynamic"] = False
            shape.pop("data_field", None)
            shape.pop("resolved_value", None)
            shape.pop("resolved_tokens", None)
        for image in slide.get("images", []):
            image["is_dynamic"] = False
            image.pop("source", None)
            image.pop("resolved_source", None)

    # ── Apply each mapping ──
    applied = 0
    for entry in mapping_list:
        slide_key = entry.get("slide", "")
        shape_id = str(entry.get("shape_id", ""))
        mapping_type = entry.get("type", "")

        slide = config["slides"].get(slide_key)
        if slide is None:
            logger.warning("  WARNING Mapping references unknown slide: %s", slide_key)
            continue

        # Find the shape
        target_shape = None
        for shape in slide["shapes"]:
            if str(shape.get("shape_id", "")) == shape_id:
                target_shape = shape
                break

        if target_shape is None:
            logger.warning("  WARNING Mapping references unknown shape_id=%s on %s", shape_id, slide_key)
            continue

        # Mark shape as dynamic
        target_shape["is_dynamic"] = True

        if mapping_type in ("text", "table"):
            target_shape["data_field"] = entry.get("data_field", "")
            logger.info("  ok %s shape_id=%s -> data_field='%s' (%s)",
                        slide_key, shape_id, target_shape["data_field"], mapping_type)

        elif mapping_type == "image":
            source = entry.get("source", "")
            target_shape["data_field"] = source

            # Also mark the corresponding images[] entry
            rid = target_shape.get("image_rid", "")
            if rid:
                for image in slide.get("images", []):
                    if image.get("rid") == rid:
                        image["is_dynamic"] = True
                        image["source"] = source
                        break
            else:
                # Try matching by target_shape_id
                for image in slide.get("images", []):
                    if str(image.get("target_shape_id", "")) == shape_id:
                        image["is_dynamic"] = True
                        image["source"] = source
                        break

            logger.info("  ok %s shape_id=%s -> source='%s' (image)",
                        slide_key, shape_id, source)

        applied += 1

    logger.info("Applied %d mappings from %s", applied, mappings_path)
    return applied


def update_config(config_path: str, data_dir: str, mappings_path: str = None):
    logger.info("Updating config: %s", config_path)
    logger.info("Data directory: %s", data_dir)

    # Load config
    if not os.path.exists(config_path):
        logger.error("ERROR Config not found: %s. Run generate_config.py first.", config_path)
        sys.exit(1)

    with open(config_path, encoding='utf-8') as f:
        config = json.load(f)

    # ── Idempotency: reset geometries and computed fields to original values ──
    # This ensures re-running update_config doesn't compound changes.
    library_path = os.path.join(os.path.dirname(config_path), "..", "component_library")
    manifest_path = os.path.join(library_path, "manifest.json")
    if os.path.exists(manifest_path):
        with open(manifest_path, encoding='utf-8') as f:
            manifest = json.load(f)
        # Build lookup: (slide_number, shape_id) -> original geometry from manifest
        # Shape IDs are only unique within a slide, not across slides
        original_geometries = {}
        for slide_meta in manifest["slides"]:
            slide_num = slide_meta.get("slide_number")
            for shape in slide_meta["shapes"]:
                sid = shape.get("id")
                if sid and shape.get("geometry") and slide_num is not None:
                    original_geometries[(slide_num, str(sid))] = dict(shape["geometry"])

        # Reset all shape geometries and clear computed fields
        for slide_key, slide in config["slides"].items():
            slide_num = slide.get("slide_number")
            for shape in slide["shapes"]:
                sid = str(shape.get("shape_id", ""))
                key = (slide_num, sid)
                if key in original_geometries:
                    shape["geometry"] = dict(original_geometries[key])
                # Clear previously computed values so they're recalculated fresh
                # Only clear resolved values when the shape has a data_field mapping;
                # manually set values (no data_field) are preserved.
                data_field = shape.get("data_field", "")
                if data_field:
                    shape.pop("resolved_value", None)
                    shape.pop("resolved_tokens", None)
                layout = shape.get("layout")
                if layout:
                    layout.pop("_computed", None)
            # Clear image computed fields
            for image in slide.get("images", []):
                image.pop("_computed", None)
                image.pop("resolved_source", None)
        logger.info("Reset geometries and computed fields from manifest (idempotent).")

    # ── Apply mappings from mappings.json ──
    if mappings_path is None:
        # Default: same directory as config, named <deck>_mappings.json
        config_dir = os.path.dirname(config_path)
        deck_name = config.get("deck", "")
        if deck_name:
            mappings_path = os.path.join(config_dir, f"{deck_name}_mappings.json")
        else:
            mappings_path = os.path.join(config_dir, "mappings.json")

    applied = apply_mappings(config, mappings_path)
    if applied > 0:
        logger.info("")

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

            # Literal values: "literal:some text" resolves directly
            if data_field.startswith("literal:"):
                shape["resolved_value"] = data_field[len("literal:"):]
                updates += 1
                logger.info("  ok %s -> %s: literal -> '%s'", slide_key, shape['shape_name'], shape['resolved_value'][:80])
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

            # Use the median of captured font_sizes (>= 7pt) as baseline;
            # this represents the body font, not the largest decorative font.
            # Fall back to max_font only if font_sizes is absent or empty.
            raw_sizes = shape.get("font_sizes", [])
            body_sizes = [s for s in raw_sizes if s >= 700]
            if body_sizes:
                sorted_sizes = sorted(body_sizes)
                mid = len(sorted_sizes) // 2
                if len(sorted_sizes) % 2 == 0:
                    original_font = (sorted_sizes[mid - 1] + sorted_sizes[mid]) // 2
                else:
                    original_font = sorted_sizes[mid]
            else:
                original_font = max_font

            if ":" in resolved_value:
                logger.debug("Resolved value for shape '%s' contains colons — verify _split_by_run_structure output", shape.get("shape_name", "?"))

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
            # Default to fit_width so images scale proportionally
            fit_mode = layout.get("fit", "fit_width")
            anchor = layout.get("anchor", "center")

            dpi_x, _dpi_y = read_image_dpi(source_path)
            computed = compute_image_fit(img_w, img_h, max_cx, max_cy, fit=fit_mode, anchor=anchor, dpi=dpi_x)
            computed["img_width_px"] = img_w
            computed["img_height_px"] = img_h
            image["_computed"] = computed
            logger.info("  ok %s -> image %s: %dx%d px -> cx=%d cy=%d (offset %d,%d) [%s]",
                        slide_key, image.get("rid", "?"), img_w, img_h,
                        computed["cx"], computed["cy"],
                        computed["offset_x"], computed["offset_y"], fit_mode)

    # ── Post-resolution: column-aware image stacking per slide ──
    # Runs on every slide.  Only images whose computed cy differs from the
    # original cy (i.e. they actually expanded) trigger shifting.
    SLIDE_WIDTH_MID = 6096000   # half of standard 12192000 EMU width
    SLIDE_H = 6858000           # standard slide height in EMU

    for slide_key, slide in config["slides"].items():
        all_images = slide.get("images", [])

        # ── 1. Identify expanded images (computed cy != original cy) ──
        expanded = []   # list of dicts with image info
        for image in all_images:
            if not image.get("is_dynamic") or not image.get("_computed"):
                continue
            target_shape_id = image.get("target_shape_id")
            if not target_shape_id:
                continue
            # Find the corresponding shape to get original geometry
            for shape in slide["shapes"]:
                if shape.get("shape_id") == target_shape_id and shape.get("geometry"):
                    orig_geo = shape["geometry"]
                    orig_cy = orig_geo.get("cy", 0)
                    new_cy = image["_computed"].get("cy", orig_cy)
                    if new_cy != orig_cy:
                        center_x = orig_geo.get("x", 0) + orig_geo.get("cx", 0) // 2
                        expanded.append({
                            "image": image,
                            "shape": shape,
                            "target_shape_id": target_shape_id,
                            "orig_x": orig_geo.get("x", 0),
                            "orig_cx": orig_geo.get("cx", 0),
                            "orig_y": orig_geo.get("y", 0),
                            "orig_cy": orig_cy,
                            "new_cy": new_cy,
                            "center_x": center_x,
                            "column": "left" if center_x < SLIDE_WIDTH_MID else "right",
                        })
                    break

        if not expanded:
            continue

        # ── 2. Group by column ──
        columns = {}
        for entry in expanded:
            col = entry["column"]
            columns.setdefault(col, []).append(entry)

        for col_name, col_entries in columns.items():
            # Sort by y position (top to bottom)
            col_entries.sort(key=lambda e: e["orig_y"])

            # ── Item 12: Uniform sizing — normalize cy to max among expanded in column ──
            if len(col_entries) > 1:
                max_cy = max(e["new_cy"] for e in col_entries)
                for entry in col_entries:
                    entry["new_cy"] = max_cy
                    entry["image"]["_computed"]["cy"] = max_cy

            # ── 3-5. Compute deltas and shift shapes below ──
            accumulated_delta = 0
            shifted_shape_ids = set()  # track already-shifted shapes

            for entry in col_entries:
                orig_y = entry["orig_y"]
                orig_cy = entry["orig_cy"]
                new_cy = entry["new_cy"]
                delta = new_cy - orig_cy
                target_sid = entry["target_shape_id"]

                # Apply accumulated delta to this image's shape position
                entry["shape"]["geometry"]["cy"] = new_cy
                if accumulated_delta > 0 and target_sid not in shifted_shape_ids:
                    entry["shape"]["geometry"]["y"] = orig_y + accumulated_delta
                    entry["image"]["_computed"]["new_y"] = orig_y + accumulated_delta
                    shifted_shape_ids.add(target_sid)

                logger.info("  ok %s -> image %s (%s col): cy %d -> %d, delta=%d, accum=%d",
                            slide_key, target_sid, col_name, orig_cy, new_cy, delta, accumulated_delta)

                # ── Item 5/9: Detect and resize overlays for this image ──
                img_x = entry["orig_x"]
                img_cx = entry["orig_cx"]
                img_x_end = img_x + img_cx
                img_y_end = orig_y + orig_cy
                cy_ratio = new_cy / orig_cy if orig_cy > 0 else 1.0

                for shape in slide["shapes"]:
                    if shape.get("shape_id") == target_sid:
                        continue
                    geo = shape.get("geometry")
                    if not geo:
                        continue
                    sx = geo.get("x", 0)
                    scx = geo.get("cx", 0)
                    sy = geo.get("y", 0)
                    scy = geo.get("cy", 0)
                    sx_end = sx + scx

                    # Check x-range overlap
                    x_overlaps = sx < img_x_end and sx_end > img_x
                    if not x_overlaps:
                        continue

                    # Check y within image's original y..y+cy range
                    y_within = orig_y <= sy < img_y_end

                    if not y_within:
                        continue

                    # Check cy condition: close to image cy (within 20%) OR fully contained
                    cy_close = abs(scy - orig_cy) <= orig_cy * 0.20 if orig_cy > 0 else False
                    fully_contained = (sy >= orig_y) and (sy + scy <= img_y_end)

                    if cy_close or fully_contained:
                        geo["cy"] = int(scy * cy_ratio)
                        # Also shift overlay by accumulated delta
                        if accumulated_delta > 0 and shape["shape_id"] not in shifted_shape_ids:
                            geo["y"] = sy + accumulated_delta
                            shifted_shape_ids.add(shape["shape_id"])
                        logger.info("  ok %s -> overlay %s: resized cy=%d, ratio=%.2f",
                                    slide_key, shape["shape_id"], geo["cy"], cy_ratio)

                # Accumulate delta for shapes below this image
                accumulated_delta += delta

            # ── 4-5. Shift all shapes below expanded images in this column ──
            if accumulated_delta > 0:
                # Determine column x-range from expanded images
                col_x_min = min(e["orig_x"] for e in col_entries)
                col_x_max = max(e["orig_x"] + e["orig_cx"] for e in col_entries)
                # The bottom of the lowest original image marks the threshold
                bottom_of_last_orig = max(e["orig_y"] + e["orig_cy"] for e in col_entries)

                for shape in slide["shapes"]:
                    sid = shape.get("shape_id")
                    if sid in shifted_shape_ids:
                        continue
                    geo = shape.get("geometry")
                    if not geo:
                        continue
                    sy = geo.get("y", 0)
                    sx = geo.get("x", 0)
                    scx = geo.get("cx", 0)
                    sx_end = sx + scx

                    # Shape must be in the same column (x overlap) and below the expanded images
                    x_overlaps = sx < col_x_max and sx_end > col_x_min
                    if not x_overlaps:
                        continue
                    if sy < bottom_of_last_orig:
                        continue

                    geo["y"] = sy + accumulated_delta
                    shifted_shape_ids.add(sid)
                    logger.info("  ok %s -> shifted %s down by %d to y=%d",
                                slide_key, sid, accumulated_delta, geo["y"])

        # ── Item 11: Bounds check — warn if any shape exceeds slide height ──
        for shape in slide["shapes"]:
            geo = shape.get("geometry")
            if not geo:
                continue
            bottom = geo.get("y", 0) + geo.get("cy", 0)
            if bottom > SLIDE_H:
                logger.warning("  WARNING %s -> shape %s exceeds slide height: y+cy=%d > %d",
                               slide_key, shape.get("shape_id", "?"), bottom, SLIDE_H)

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
    parser.add_argument(
        "-m", "--mappings",
        default=None,
        help="Path to mappings.json (default: configs/<deck>_mappings.json)",
    )
    args = parser.parse_args()
    update_config(args.config, args.data_dir, mappings_path=args.mappings)
