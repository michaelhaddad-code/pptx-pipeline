"""
Step 2: GENERATE CONFIG
Reads the component library and produces a config.json describing every slide.
Runs once per deck. If config already exists, skips generation (unless --force).
Usage: python generate_config.py [--library_dir DIR] [--configs_dir DIR] [--force] [--hints FILE]
"""

import argparse
import os
import json
import sys
import logging

logger = logging.getLogger(__name__)


# Shape types that are purely structural — never dynamic
STRUCTURAL_TYPES = {"nvGrpSpPr", "grpSpPr", "cxnSp"}

# Shape name patterns that are likely static design elements
STATIC_NAME_PATTERNS = ["Oval", "Connector", "Straight Connector"]

# Default text previews that look like dynamic placeholders
DEFAULT_DYNAMIC_HINTS = [
    "xxx", "$XX", "$YY", "$AA", "$ZZ", "NN ", "A/B", "as of",
    "##", "TBD", "N/A", "XX%", "YY%", "$0", "0.0",
    "YYYY", "MM/DD", "Q1", "Q2", "Q3", "Q4",
    "{{", "}}", "[placeholder]", "[TBD]", "[date]",
    "lorem", "ipsum", "sample", "example",
]


def load_dynamic_hints(hints_file: str = None) -> list:
    """Load dynamic hints from an optional JSON file, falling back to defaults."""
    hints = list(DEFAULT_DYNAMIC_HINTS)
    if hints_file and os.path.exists(hints_file):
        with open(hints_file) as f:
            custom_hints = json.load(f)
        if isinstance(custom_hints, list):
            hints.extend(custom_hints)
        else:
            logger.warning("Warning: hints file %s should contain a JSON array of strings. Ignoring.", hints_file)
    return hints


def looks_dynamic(shape: dict, dynamic_hints: list) -> bool:
    """Heuristic — flag shapes that look like they contain weekly data."""
    text = shape.get("text_preview", "")
    return any(hint.lower() in text.lower() for hint in dynamic_hints)


def looks_static(shape: dict) -> bool:
    """Heuristic — flag shapes that are structural or design-only."""
    if shape.get("type") in STRUCTURAL_TYPES:
        return True
    name = shape.get("name", "")
    return any(pattern in name for pattern in STATIC_NAME_PATTERNS)


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
        template_fonts = shape.get("font_sizes", [])
        if template_fonts:
            max_font = max(template_fonts)
            # min is 25% of the template font (floor at 400 = 4pt for readability)
            min_font = max(400, max_font // 4)
        else:
            # Fallback: read from geometry — estimate reasonable font from shape height
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
        return {
            "type": "auto_fit_text",
            "min_font_size": min_font,
            "max_font_size": max_font,
        }

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
    hints_file: str = None,
):
    dynamic_hints = load_dynamic_hints(hints_file)

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

    if os.path.exists(config_path) and force:
        logger.info("Config exists for '%s' -- overwriting (--force).", deck_name)

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
            dynamic = looks_dynamic(shape, dynamic_hints)
            static = looks_static(shape)

            shape_entry = {
                "shape_id": shape.get("id", ""),
                "shape_name": shape.get("name", ""),
                "category": category,
                "is_dynamic": dynamic and not static,
                "text_preview": shape.get("text_preview", ""),
                "data_field": "",  # field name in data files (e.g. "revenue", "metrics.total")
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
                    "is_dynamic": False,
                    "source": "",
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

        dynamic_count = sum(1 for s in shapes_config if s["is_dynamic"])
        logger.info("  Slide %d: %d shapes (%d flagged as dynamic)", slide_num, len(shapes_config), dynamic_count)

    # Save config
    with open(config_path, "w") as f:
        json.dump(config, f, indent=2)

    logger.info("\nConfig generated: %s", config_path)
    logger.info("   Review and fill in 'data_field' for dynamic shapes.")

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
    parser.add_argument(
        "--hints",
        default=None,
        help="Path to an optional JSON file containing additional dynamic hint strings",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    generate_config(
        library_path=args.library_dir,
        configs_dir=args.configs_dir,
        force=args.force,
        hints_file=args.hints,
    )
