"""
Layout computation module — pure math for dynamic sizing.
No I/O, no XML, no hardcoded template values.

All baselines come from the template via generate_config.
This module just does the arithmetic.

All dimensions are in EMU (English Metric Units):
  1 inch = 914400 EMU
  1 point = 12700 EMU
  1 pixel (96 DPI) = 9525 EMU
"""

import struct
import math


# ── Unit Constants (physical, not template-specific) ────────────────────────

EMU_PER_INCH = 914400
EMU_PER_PT = 12700
EMU_PER_PX_96DPI = 9525


# ── Table Layout ─────────────────────────────────────────────────────────────

def compute_table_layout(geometry, layout_config, n_data_rows):
    """Compute row heights and font sizes for a variable-row table.

    All baselines (row_height_baseline, font_scale_baseline, font_sizes)
    are expected to come from the template via layout_config.

    Args:
        geometry: dict with cx, cy (shape dimensions in EMU)
        layout_config: dict with sizing rules (from config "layout" field)
        n_data_rows: int, number of actual data rows (excludes header/summary)

    Returns:
        dict with:
            row_height: int (EMU)
            total_rows: int
            total_height: int (EMU)
            font_sizes: dict with header, summary_row, data_row (hundredths of pt)
            single_row: bool — True if only one data row (special formatting)
    """
    header_rows = layout_config.get("header_rows", 1)
    row_h_baseline = layout_config.get("row_height_baseline", 0)
    row_h_min = layout_config.get("row_height_min", 0)
    font_baseline = layout_config.get("font_scale_baseline", 0)
    font_sizes = layout_config.get("font_sizes", {})

    # Total rows: header + summary + data
    summary_rows = 1
    n_total = header_rows + summary_rows + max(n_data_rows, 0)

    # Available height = full shape height
    # (frame padding is template-specific, already baked into geometry)
    available_h = geometry.get("cy", 0)

    # Compute row height
    if n_total > 0 and available_h > 0:
        row_h = available_h // n_total
    else:
        row_h = row_h_baseline

    # Clamp between min and baseline
    if row_h_baseline > 0:
        row_h = min(row_h, row_h_baseline)
    if row_h_min > 0:
        row_h = max(row_h_min, row_h)

    # Font scaling: ratio of new row height to template's original row height
    if font_baseline > 0:
        scale = min(1.0, row_h / font_baseline)
    else:
        scale = 1.0

    hdr_font = font_sizes.get("header", 1000)
    sum_font = font_sizes.get("summary_row", 1700)
    data_font = font_sizes.get("data_row", 1200)

    # Scale fonts, with floors at 50% of original (never unreadable)
    computed_fonts = {
        "header": max(hdr_font // 2, int(hdr_font * scale)),
        "summary_row": max(sum_font // 2, int(sum_font * scale)),
        "data_row": max(data_font // 2, int(data_font * scale)),
    }

    return {
        "row_height": row_h,
        "total_rows": n_total,
        "total_height": row_h * n_total,
        "font_sizes": computed_fonts,
        "single_row": n_data_rows <= 1,
    }


# ── Image Layout ─────────────────────────────────────────────────────────────

def read_image_dimensions(file_path):
    """Read width and height from a PNG or JPEG file header (stdlib only).

    Args:
        file_path: str, path to image file

    Returns:
        (width_px, height_px) or None if format is unrecognized
    """
    with open(file_path, "rb") as f:
        header = f.read(32)

    # PNG: 8-byte signature, then IHDR chunk with width (4B) + height (4B)
    if header[:8] == b"\x89PNG\r\n\x1a\n":
        width = struct.unpack(">I", header[16:20])[0]
        height = struct.unpack(">I", header[20:24])[0]
        return (width, height)

    # JPEG: scan for SOF0/SOF2 marker
    if header[:2] == b"\xff\xd8":
        return _read_jpeg_dimensions(file_path)

    return None


def _read_jpeg_dimensions(file_path):
    """Parse JPEG to find SOF marker and extract dimensions."""
    with open(file_path, "rb") as f:
        f.read(2)  # skip SOI
        while True:
            marker = f.read(2)
            if len(marker) < 2:
                return None
            if marker[0] != 0xFF:
                return None
            code = marker[1]
            # SOF0 (0xC0) or SOF2 (0xC2) — baseline or progressive
            if code in (0xC0, 0xC2):
                f.read(3)  # length (2B) + precision (1B)
                height = struct.unpack(">H", f.read(2))[0]
                width = struct.unpack(">H", f.read(2))[0]
                return (width, height)
            # Skip other markers
            if code == 0xD9:  # EOI
                return None
            if code == 0x00 or (0xD0 <= code <= 0xD7):
                continue  # padding or RST
            length_bytes = f.read(2)
            if len(length_bytes) < 2:
                return None
            length = struct.unpack(">H", length_bytes)[0]
            f.read(length - 2)
    return None


def compute_image_fit(img_width_px, img_height_px, max_cx, max_cy,
                      fit="contain", anchor="center"):
    """Compute fitted image dimensions and offset within a bounding box.

    Args:
        img_width_px, img_height_px: source image size in pixels
        max_cx, max_cy: bounding box in EMU
        fit: "contain" (fit within), "cover" (fill, crop), "stretch",
             "fit_width" (match width, compute height proportionally)
        anchor: "center", "top-left", "top-center"

    Returns:
        dict with cx, cy (fitted EMU), offset_x, offset_y (EMU)
    """
    if img_width_px <= 0 or img_height_px <= 0:
        return {"cx": max_cx, "cy": max_cy, "offset_x": 0, "offset_y": 0}

    img_w_emu = img_width_px * EMU_PER_PX_96DPI
    img_h_emu = img_height_px * EMU_PER_PX_96DPI

    if fit == "stretch":
        return {"cx": max_cx, "cy": max_cy, "offset_x": 0, "offset_y": 0}

    if fit == "fit_width":
        # Scale to match shape width, compute height proportionally
        # No height cap — the auto-stacker handles overflow by scaling all sections
        scale = max_cx / img_w_emu
        fit_cx = max_cx
        fit_cy = int(img_h_emu * scale)
        return {"cx": fit_cx, "cy": fit_cy, "offset_x": 0, "offset_y": 0}

    if fit == "cover":
        scale = max(max_cx / img_w_emu, max_cy / img_h_emu)
    else:  # contain
        scale = min(max_cx / img_w_emu, max_cy / img_h_emu)

    fit_cx = int(img_w_emu * scale)
    fit_cy = int(img_h_emu * scale)

    # Clamp to bounds (cover can exceed)
    fit_cx = min(fit_cx, max_cx)
    fit_cy = min(fit_cy, max_cy)

    # Anchor offset
    if anchor == "top-left":
        offset_x, offset_y = 0, 0
    elif anchor == "top-center":
        offset_x = (max_cx - fit_cx) // 2
        offset_y = 0
    else:  # center
        offset_x = (max_cx - fit_cx) // 2
        offset_y = (max_cy - fit_cy) // 2

    return {
        "cx": fit_cx,
        "cy": fit_cy,
        "offset_x": offset_x,
        "offset_y": offset_y,
    }


# ── Slide-level Image Stacking ──────────────────────────────────────────────

SLIDE_HEIGHT_EMU = 6858000  # Standard 7.5" slide height
SLIDE_TOP_MARGIN = 150000   # Top margin padding
SLIDE_BOTTOM_MARGIN = 150000  # Bottom margin padding


def compute_slide_image_stack(image_entries, static_shapes=None):
    """Restack all images on a slide so nothing overlaps.

    All sections (dynamic and static) are scaled by a uniform factor if the
    total height exceeds available slide space. This keeps row sizes consistent
    across all tables. Static image content isn't replaced, but their shape
    dimensions are scaled so they don't hog space.

    Args:
        image_entries: list of dicts, each with:
            - target_shape_id: str
            - _computed: dict with cx, cy
            - _original_geometry: dict with x, y, cx, cy
            - _label_shape: optional dict with shape_id and geometry
        static_shapes: list of dicts, each with:
            - shape_id: str
            - geometry: dict with x, y, cx, cy
            - _label_shape: optional dict with shape_id and geometry

    Returns:
        list of dicts with: shape_id, cx, cy, new_x, new_y, scale_factor,
        and optionally label_shape_id, label_new_y for repositioned labels.
    """
    if not image_entries:
        return []

    gap = 50000  # ~0.05" gap between sections
    label_height = 310000  # ~0.25" for section label

    # Build a unified list of sections sorted by original Y position
    sections = []

    for e in image_entries:
        y = e["_original_geometry"].get("y", 0)
        if e.get("_label_shape"):
            y = min(y, e["_label_shape"].get("geometry", {}).get("y", y))
        sections.append({
            "type": "dynamic",
            "sort_y": y,
            "entry": e,
        })

    for s in (static_shapes or []):
        y = s["geometry"].get("y", 0)
        if s.get("_label_shape"):
            y = min(y, s["_label_shape"].get("geometry", {}).get("y", y))
        sections.append({
            "type": "static",
            "sort_y": y,
            "entry": s,
        })

    sections.sort(key=lambda s: s["sort_y"])

    n_sections = len(sections)

    # Determine the top starting Y from the first section
    first_y = sections[0]["sort_y"]
    available_height = SLIDE_HEIGHT_EMU - first_y - SLIDE_BOTTOM_MARGIN

    # Compute fixed overhead: labels + gaps + static image heights
    total_gaps = gap * max(n_sections - 1, 0)
    total_labels = sum(
        label_height for sec in sections if sec["entry"].get("_label_shape")
    )
    total_static = sum(
        sec["entry"]["geometry"].get("cy", 0)
        for sec in sections if sec["type"] == "static"
    )

    # Remaining space for dynamic images
    n_dynamic = sum(1 for sec in sections if sec["type"] == "dynamic")
    available_for_dynamic = available_height - total_gaps - total_labels - total_static

    if n_dynamic >= 2 and available_for_dynamic > 0:
        # Multiple dynamic images: distribute space proportionally based on
        # each image's natural (fit_width) height so aspect ratios are preserved.
        proportional_heights = []
        for sec in sections:
            if sec["type"] == "dynamic":
                proportional_heights.append(sec["entry"]["_computed"]["cy"])
        total_proportional = sum(proportional_heights) or 1
        for sec in sections:
            if sec["type"] == "dynamic":
                ratio = sec["entry"]["_computed"]["cy"] / total_proportional
                sec["entry"]["_computed"]["cy"] = int(available_for_dynamic * ratio)
    elif n_dynamic == 1 and available_for_dynamic > 0:
        # Single dynamic image: use its proportional height, capped to available space
        for sec in sections:
            if sec["type"] == "dynamic":
                proportional_cy = sec["entry"]["_computed"]["cy"]
                sec["entry"]["_computed"]["cy"] = min(proportional_cy, available_for_dynamic)

    # Compute total content height for overflow check
    total_content_height = 0
    for sec in sections:
        if sec["type"] == "static":
            total_content_height += sec["entry"]["geometry"].get("cy", 0)
        else:
            total_content_height += sec["entry"]["_computed"]["cy"]
        if sec["entry"].get("_label_shape"):
            total_content_height += label_height

    total_needed = total_content_height + total_gaps
    if total_needed > available_height and total_needed > 0:
        scale_factor = available_height / total_needed
    else:
        scale_factor = 1.0

    # Lay out sections top-to-bottom
    results = []
    current_y = first_y

    for sec in sections:
        if sec["type"] == "static":
            s = sec["entry"]
            geo = s["geometry"]
            scaled_cy = int(geo.get("cy", 0) * scale_factor)

            result = {
                "shape_id": s["shape_id"],
                "cx": geo.get("cx", 0),
                "cy": scaled_cy,
                "new_x": geo.get("x", 0),
                "new_y": current_y,
                "scale_factor": scale_factor,
                "is_static": True,
            }

            if s.get("_label_shape"):
                result["label_shape_id"] = s["_label_shape"]["shape_id"]
                result["label_new_y"] = current_y
                result["label_cy"] = int(label_height * scale_factor)
                current_y += int(label_height * scale_factor)

            result["new_y"] = current_y
            current_y += scaled_cy + int(gap * scale_factor)

            results.append(result)

        else:
            e = sec["entry"]
            shape_id = e["target_shape_id"]
            orig_geo = e["_original_geometry"]
            computed = e["_computed"]

            scaled_cy = int(computed["cy"] * scale_factor)

            result = {
                "shape_id": shape_id,
                "cx": computed["cx"],
                "cy": scaled_cy,
                "new_x": orig_geo.get("x", 0),
                "new_y": current_y,
                "scale_factor": scale_factor,
            }

            if e.get("_label_shape"):
                result["label_shape_id"] = e["_label_shape"]["shape_id"]
                result["label_new_y"] = current_y
                result["label_cy"] = int(label_height * scale_factor)
                current_y += int(label_height * scale_factor)

            result["new_y"] = current_y
            current_y += scaled_cy + int(gap * scale_factor)

            results.append(result)

    return results


# ── Text Layout ──────────────────────────────────────────────────────────────

def compute_text_font_scale(text, shape_cx, shape_cy,
                            original_font_size, min_font, max_font):
    """Estimate a font size that fits text within a shape.

    Heuristic only — no text shaping engine. Uses average character width
    as ~60% of font em-square.

    Args:
        text: str, the text to fit
        shape_cx, shape_cy: shape dimensions in EMU
        original_font_size: int, in hundredths of a point (e.g. 1400 = 14pt)
        min_font, max_font: int, font size bounds (hundredths of a point)

    Returns:
        int — computed font size in hundredths of a point
    """
    if not text or shape_cx <= 0 or shape_cy <= 0 or original_font_size <= 0:
        return original_font_size

    font_size = min(original_font_size, max_font)

    # Convert font size to EMU (hundredths of point -> EMU)
    font_emu = font_size * EMU_PER_PT / 100

    # Estimate characters per line (~60% of em width per char)
    char_width_emu = font_emu * 0.6
    if char_width_emu <= 0:
        return font_size

    chars_per_line = shape_cx / char_width_emu
    if chars_per_line <= 0:
        return font_size

    # Estimate lines needed
    lines_needed = math.ceil(len(text) / chars_per_line)
    # Single line: font em is sufficient. Multi-line: 120% spacing between lines.
    if lines_needed <= 1:
        height_needed = font_emu
    else:
        height_needed = font_emu + (lines_needed - 1) * font_emu * 1.2

    if height_needed > shape_cy and height_needed > 0:
        scale = shape_cy / height_needed
        font_size = int(font_size * scale)

    return max(min_font, min(font_size, max_font))
