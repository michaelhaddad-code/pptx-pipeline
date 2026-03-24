"""
Step 2: GENERATE CONFIG
Reads the component library and produces a config.json describing every slide.
All shapes are listed neutrally — classification as dynamic/static happens
interactively in Step 3 (Map) when the user brings their data.
Runs once per deck. If config already exists, skips generation (unless --force).
Usage: python generate_config.py [--library_dir DIR] [--configs_dir DIR] [--force]
"""

import argparse
import os
import json
import sys
import logging
import statistics

logger = logging.getLogger(__name__)


# Shape types that are purely structural (group wrappers, connectors) — excluded from config
STRUCTURAL_TYPES = {"nvGrpSpPr", "grpSpPr", "cxnSp"}


def get_shape_category(shape: dict) -> str:
    """Categorize shape type for the config."""
    stype = shape.get("type", "")
    if stype == "pic":
        return "image"
    elif stype == "sp":
        return "text"
    elif stype == "grpSp":
        return "group"
    elif stype == "graphicFrame":
        return "table"
    elif stype in STRUCTURAL_TYPES:
        return "structural"
    else:
        return "unknown"


def _make_layout_stub(category: str, shape: dict) -> dict:
    """Create a layout rule stub derived from the template's actual dimensions.

    All baselines come from the template itself — no hardcoded magic numbers.
    """
    if category == "text":
        # Derive font bounds from the template's actual font sizes
        font_sizes = shape.get("font_sizes", [])
        if font_sizes:
            # Fix 23: use actual font data from shape runs
            max_font = max(font_sizes)
            min_font = max(400, min(font_sizes) // 2)

            # Fix 41: per-run font metadata for multi-size shapes
            header_font = max(font_sizes)
            body_font = statistics.median(font_sizes)
            # Use body_font as max_font for auto-fit so decorative headers
            # don't inflate the baseline and over-shrink body text
            max_font = int(body_font)
        else:
            # Fallback: read from geometry — estimate reasonable font from shape height
            header_font = None
            body_font = None
            geo = shape.get("geometry", {})
            cy = geo.get("cy", 0)
            if cy > 0:
                # Rough heuristic: shape height in EMU / 12700 gives points,
                # shape can fit ~1-3 lines, so max font ~ cy / 12700 / 2 * 100
                max_font = max(800, int(cy / 12700 / 2 * 100))
                min_font = max(400, max_font // 4)
            else:
                max_font = 2400
                min_font = 600

        stub = {
            "type": "auto_fit_text",
            "min_font_size": min_font,
            "max_font_size": max_font,
        }
        if header_font is not None:
            stub["header_font"] = header_font
            stub["body_font"] = int(body_font)
        return stub

    elif category == "table":
        table_grid = shape.get("table_grid", {})
        row_heights = table_grid.get("row_heights", [])
        row_fonts = table_grid.get("row_fonts", [])

        # Derive row height baseline from the template's actual row heights
        if row_heights:
            # Use the most common row height (data rows) as baseline
            # Skip first row (header) if there are multiple rows
            data_heights = row_heights[1:] if len(row_heights) > 1 else row_heights
            row_h_baseline = max(data_heights) if data_heights else row_heights[0]
            row_h_min = min(row_heights) // 2  # allow shrinking to half
            row_h_min = max(row_h_min, 100000)  # but not below ~0.1 inch
        else:
            geo = shape.get("geometry", {})
            n_rows = table_grid.get("rows", 1) or 1
            cy = geo.get("cy", 0)
            row_h_baseline = cy // n_rows if cy > 0 else 348711
            row_h_min = row_h_baseline // 2

        # font_scale_baseline = the row height at which fonts are 100%
        # This IS the template's row height — that's the "normal" size
        font_scale_baseline = row_h_baseline

        # Derive font sizes from the template's actual per-row fonts
        header_font = 1000
        summary_font = 1700
        data_font = 1200

        if row_fonts:
            # First row = header
            if row_fonts[0]:
                header_font = max(row_fonts[0])
            # Last row = summary (if more than 2 rows)
            if len(row_fonts) > 2 and row_fonts[-1]:
                summary_font = max(row_fonts[-1])
            # Middle rows = data
            data_rows_fonts = row_fonts[1:-1] if len(row_fonts) > 2 else row_fonts[1:]
            if data_rows_fonts:
                all_data_fonts = [f for rf in data_rows_fonts for f in rf]
                if all_data_fonts:
                    data_font = max(all_data_fonts)

        return {
            "type": "dynamic_table",
            "fixed_columns": table_grid.get("columns", 0),
            "header_rows": 1,
            "row_height_baseline": row_h_baseline,
            "row_height_min": row_h_min,
            "font_scale_baseline": font_scale_baseline,
            "font_sizes": {
                "header": header_font,
                "summary_row": summary_font,
                "data_row": data_font,
            },
        }

    elif category == "image":
        geo = shape.get("geometry", {})
        return {
            "type": "dynamic_image",
            "fit": "stretch",
            "anchor": "center",
            "max_cx": geo.get("cx", 0),
            "max_cy": geo.get("cy", 0),
        }
    return None


def generate_config(
    library_path: str = "component_library",
    configs_dir: str = "configs",
    force: bool = False,
):
    # Load manifest
    manifest_path = os.path.join(library_path, "manifest.json")
    if not os.path.exists(manifest_path):
        logger.info("No manifest found at %s. Run deconstruct.py first.", manifest_path)
        sys.exit(1)

    with open(manifest_path) as f:
        manifest = json.load(f)

    deck_name = os.path.splitext(manifest["source"])[0].replace(" ", "_").lower()
    os.makedirs(configs_dir, exist_ok=True)
    config_path = os.path.join(configs_dir, f"{deck_name}.json")

    # Check if config already exists
    if os.path.exists(config_path) and not force:
        logger.info("Config already exists for '%s' -- skipping generation.", deck_name)
        logger.info("  Config: %s", config_path)
        logger.info("  Use --force to overwrite.")
        return config_path

    # Fix 49: when --force, load existing config to preserve mapping work
    existing_shapes_by_id = {}
    if os.path.exists(config_path) and force:
        logger.info("Config exists for '%s' -- overwriting (--force).", deck_name)
        with open(config_path) as f:
            existing_config = json.load(f)
        for slide_data in existing_config.get("slides", {}).values():
            for s in slide_data.get("shapes", []):
                sid = s.get("shape_id")
                if sid:
                    existing_shapes_by_id[sid] = s

    logger.info("Generating config for: %s", deck_name)

    config = {
        "deck": deck_name,
        "source": manifest["source"],
        "total_slides": manifest["total_slides"],
        "slides": {}
    }

    for slide_meta in manifest["slides"]:
        slide_num = slide_meta["slide_number"]
        slide_key = f"slide_{slide_num}"
        shapes_config = []

        for shape in slide_meta["shapes"]:
            # Skip purely structural elements
            if shape.get("type") in STRUCTURAL_TYPES:
                continue

            category = get_shape_category(shape)

            shape_entry = {
                "shape_id": shape.get("id", ""),
                "shape_name": shape.get("name", ""),
                "category": category,
                "is_dynamic": False,  # user defines this interactively in Step 3 (Map)
                "text_preview": shape.get("text_preview", ""),
                "data_field": "",  # set during mapping
            }

            # Geometry (position + size in EMU)
            if shape.get("geometry"):
                shape_entry["geometry"] = shape["geometry"]

            # Table grid info (column/row counts and sizes)
            if shape.get("table_grid"):
                shape_entry["table_grid"] = shape["table_grid"]

            # Image embed reference
            if shape.get("image_rid"):
                shape_entry["image_rid"] = shape["image_rid"]

            # Layout rule stub for dynamic sizing
            layout = _make_layout_stub(category, shape)
            if layout:
                shape_entry["layout"] = layout

            # Track parent group for nested shapes
            if shape.get("parent_group"):
                shape_entry["parent_group"] = shape["parent_group"]

            # Fix 49: preserve mapping data from previous config if shape still exists
            prev = existing_shapes_by_id.get(shape_entry["shape_id"])
            if prev:
                for key in ("is_dynamic", "data_field", "resolved_tokens", "resolved_value"):
                    if key in prev:
                        shape_entry[key] = prev[key]

            shapes_config.append(shape_entry)

        # Also include relationships (images linked to this slide)
        # Cross-reference with pic shapes to link rid -> shape_id
        rid_to_shape = {}
        for sc in shapes_config:
            if sc.get("image_rid"):
                rid_to_shape[sc["image_rid"]] = sc["shape_id"]

        relationships = []
        for rid, rel in slide_meta.get("relationships", {}).items():
            if rel["type"] == "image":
                img_rel = {
                    "rid": rid,
                    "target": rel["target"],
                    "is_dynamic": False,  # user defines this in Step 3 (Map)
                    "source": "",  # set during mapping
                }
                if rid in rid_to_shape:
                    img_rel["target_shape_id"] = rid_to_shape[rid]
                relationships.append(img_rel)

        config["slides"][slide_key] = {
            "slide_number": slide_num,
            "shape_count": len(shapes_config),
            "shapes": shapes_config,
            "images": relationships
        }

        logger.info("  Slide %d: %d shapes", slide_num, len(shapes_config))

    # Save config
    with open(config_path, "w") as f:
        json.dump(config, f, indent=2)

    logger.info("\nConfig generated: %s", config_path)
    logger.info("   Run Step 3 (Map) to interactively classify shapes and assign data fields.")

    return config_path


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate a config.json describing every slide in the component library."
    )
    parser.add_argument(
        "--library_dir",
        default="component_library",
        help="Path to the component library directory (default: component_library)",
    )
    parser.add_argument(
        "--configs_dir",
        default="configs",
        help="Path to the configs output directory (default: configs)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing config file if it already exists",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    generate_config(
        library_path=args.library_dir,
        configs_dir=args.configs_dir,
        force=args.force,
    )
