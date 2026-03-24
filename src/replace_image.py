"""
Replace an image in a deconstructed PPTX component library.

Handles three concerns:
1. Swap the media file on disk
2. Resize the shape to fit the new image proportionally (full width, height scales)
3. Find overlay shapes (same position/size as the image) and resize them to match
4. Push all shapes below the image down by the height delta

Usage:
    python replace_image.py <library_dir> <config_path> <slide_number> <shape_id> <new_image_path>
"""

import os
import re
import json
import sys
import shutil
import logging
from PIL import Image

logger = logging.getLogger(__name__)

# EMU constants
EMU_PER_INCH = 914400
EMU_PER_PT = 12700
SLIDE_WIDTH_EMU = 12192000  # standard 16:9 slide width


def _find_shape_span(xml_str: str, shape_id: str) -> tuple | None:
    """Find start/end offsets of the shape element containing the given cNvPr id."""
    # Find the cNvPr with matching id
    cnvpr_pattern = re.compile(
        r'<[^>]*?:cNvPr\b[^>]*?\bid\s*=\s*"' + re.escape(shape_id) + r'"'
    )
    m = cnvpr_pattern.search(xml_str)
    if not m:
        return None

    # Walk backwards to find the enclosing shape tag
    shape_tags = ("p:sp", "p:pic", "p:grpSp", "p:graphicFrame")
    pos = m.start()

    # Find the nearest opening shape tag before this position
    best_start = -1
    best_tag = None
    for tag in shape_tags:
        # Look for <p:pic>, <p:sp>, etc.
        idx = xml_str.rfind(f"<{tag}", 0, pos)
        if idx > best_start:
            best_start = idx
            best_tag = tag

    if best_start < 0 or best_tag is None:
        return None

    # Find the matching closing tag
    close_tag = f"</{best_tag}>"
    end_idx = xml_str.find(close_tag, pos)
    if end_idx < 0:
        return None

    return (best_start, end_idx + len(close_tag))


def _get_shape_offset_and_extent(shape_xml: str) -> dict | None:
    """Extract x, y, cx, cy from the first <a:off> and <a:ext> in a shape."""
    off_m = re.search(
        r'<[^>]*?:off\b[^>]*?\bx\s*=\s*"(\d+)"[^>]*?\by\s*=\s*"(\d+)"', shape_xml
    )
    ext_m = re.search(
        r'<[^>]*?:ext\b[^>]*?\bcx\s*=\s*"(\d+)"[^>]*?\bcy\s*=\s*"(\d+)"', shape_xml
    )
    if not off_m or not ext_m:
        return None
    return {
        "x": int(off_m.group(1)),
        "y": int(off_m.group(2)),
        "cx": int(ext_m.group(1)),
        "cy": int(ext_m.group(2)),
    }


def _set_shape_cy(shape_xml: str, new_cy: int) -> str:
    """Replace the first <a:ext cy="..."> value in shape XML."""
    return re.sub(
        r'(<[^>]*?:ext\b[^>]*?\bcy\s*=\s*")(\d+)(")',
        lambda m: m.group(1) + str(new_cy) + m.group(3),
        shape_xml,
        count=1,
    )


def _set_shape_y(shape_xml: str, new_y: int) -> str:
    """Replace the first <a:off y="..."> value in shape XML."""
    return re.sub(
        r'(<[^>]*?:off\b[^>]*?\by\s*=\s*")(\d+)(")',
        lambda m: m.group(1) + str(new_y) + m.group(3),
        shape_xml,
        count=1,
    )


def _find_overlays(xml_str: str, target_y: int, target_cy: int,
                   target_x: int, target_cx: int,
                   image_shape_id: str, all_shapes: list) -> list:
    """Find shapes that overlap the image and have matching height (overlays/highlight boxes).

    An overlay is a shape whose:
    - y position is within 5000 EMU of the image y
    - cy (height) is within 5000 EMU of the image cy
    - x position overlaps horizontally with the image
    - is not the image itself
    """
    TOLERANCE = 5000  # EMU tolerance for position matching
    overlays = []

    for s in all_shapes:
        sid = str(s.get("shape_id", ""))
        if sid == str(image_shape_id):
            continue

        geo = s.get("geometry", {})
        sy = geo.get("y", 0)
        scy = geo.get("cy", 0)
        sx = geo.get("x", 0)
        scx = geo.get("cx", 0)

        # Check vertical alignment: y and cy should be close to the image's
        y_close = abs(sy - target_y) < TOLERANCE
        cy_close = abs(scy - target_cy) < TOLERANCE

        # Check horizontal overlap
        img_left = target_x
        img_right = target_x + target_cx
        shape_left = sx
        shape_right = sx + scx
        h_overlap = shape_left < img_right and shape_right > img_left

        if y_close and cy_close and h_overlap:
            overlays.append(s)

    return overlays


def replace_image(library_dir: str, config_path: str,
                  slide_number: int, shape_id: int,
                  new_image_path: str) -> str:
    """Replace an image in the component library and adjust layout.

    Args:
        library_dir: path to component_library
        config_path: path to config JSON
        slide_number: 1-based slide number
        shape_id: numeric shape id (from config)
        new_image_path: path to the replacement image file

    Returns:
        path to the modified slide XML
    """
    # ── Load config ──
    with open(config_path) as f:
        cfg = json.load(f)

    slide_key = f"slide_{slide_number}"
    slide_cfg = cfg["slides"][slide_key]
    all_shapes = slide_cfg["shapes"]
    images = slide_cfg.get("images", [])

    # ── Find the image entry in config ──
    image_entry = None
    for img in images:
        if str(img.get("target_shape_id", "")) == str(shape_id):
            image_entry = img
            break

    if image_entry is None:
        raise ValueError(f"No image entry found for shape_id={shape_id} on slide {slide_number}")

    media_target = image_entry["target"]  # e.g. "../media/image45.jpeg"
    media_filename = os.path.basename(media_target)
    media_path = os.path.join(library_dir, "_raw", "ppt", "media", media_filename)

    # ── Get new image dimensions ──
    new_img = Image.open(new_image_path)
    new_width, new_height = new_img.size
    new_aspect = new_width / new_height
    logger.info("New image: %dx%d (aspect=%.4f)", new_width, new_height, new_aspect)

    # ── Read slide XML ──
    slide_xml_path = os.path.join(library_dir, "_raw", "ppt", "slides", f"slide{slide_number}.xml")
    with open(slide_xml_path, "r", encoding="utf-8") as f:
        xml_str = f.read()

    # ── Find the image shape and get its current geometry ──
    sid_str = str(shape_id)
    span = _find_shape_span(xml_str, sid_str)
    if span is None:
        raise ValueError(f"Shape id={shape_id} not found in slide XML")

    shape_xml = xml_str[span[0]:span[1]]
    geo = _get_shape_offset_and_extent(shape_xml)
    if geo is None:
        raise ValueError(f"Could not parse geometry for shape id={shape_id}")

    original_x = geo["x"]
    original_y = geo["y"]
    original_cx = geo["cx"]
    original_cy = geo["cy"]
    logger.info("Original shape: x=%d y=%d cx=%d cy=%d", original_x, original_y, original_cx, original_cy)

    # ── Compute new height: keep full width, scale height proportionally ──
    new_cy = int(original_cx / new_aspect)
    delta = new_cy - original_cy
    logger.info("New cy=%d (delta=%d, %.2f inches)", new_cy, delta, delta / EMU_PER_INCH)

    # ── Step 1: Replace the media file ──
    ext = os.path.splitext(new_image_path)[1].lower()
    media_ext = os.path.splitext(media_path)[1].lower()
    shutil.copy2(new_image_path, media_path)
    logger.info("Replaced media file: %s", media_path)

    # ── Step 2: Resize the image shape ──
    new_shape_xml = _set_shape_cy(shape_xml, new_cy)
    xml_str = xml_str[:span[0]] + new_shape_xml + xml_str[span[1]:]
    logger.info("Resized shape id=%s: cy %d -> %d", sid_str, original_cy, new_cy)

    # ── Step 3: Find and resize overlay shapes ──
    overlays = _find_overlays(xml_str, original_y, original_cy,
                              original_x, original_cx, sid_str, all_shapes)
    overlay_ids = set()
    for ov in overlays:
        ov_id = str(ov["shape_id"])
        overlay_ids.add(ov_id)
        ov_span = _find_shape_span(xml_str, ov_id)
        if ov_span is None:
            logger.warning("Overlay shape id=%s not found in XML", ov_id)
            continue
        ov_xml = xml_str[ov_span[0]:ov_span[1]]
        new_ov_xml = _set_shape_cy(ov_xml, new_cy)
        xml_str = xml_str[:ov_span[0]] + new_ov_xml + xml_str[ov_span[1]:]
        logger.info("Resized overlay id=%s ('%s'): cy %d -> %d",
                     ov_id, ov.get("shape_name", ""), original_cy, new_cy)

    # ── Step 4: Push down shapes below the image (same column only) ──
    if delta > 0:
        image_bottom = original_y + original_cy
        skip_ids = {sid_str} | overlay_ids
        shifted = 0

        # Determine which column the image is in (left vs right of slide center)
        half_slide = SLIDE_WIDTH_EMU // 2
        image_center_x = original_x + original_cx // 2
        image_is_left = image_center_x < half_slide

        for s in all_shapes:
            s_id = str(s.get("shape_id", ""))
            if s_id in skip_ids:
                continue

            geo_s = s.get("geometry", {})
            sy = geo_s.get("y", 0)
            sx = geo_s.get("x", 0)
            scx = geo_s.get("cx", 0)

            # Only push shapes in the same column
            shape_center_x = sx + scx // 2
            shape_is_left = shape_center_x < half_slide
            if shape_is_left != image_is_left:
                continue

            # Only push shapes whose top is at or below the image's original bottom
            if sy < image_bottom:
                continue

            s_span = _find_shape_span(xml_str, s_id)
            if s_span is None:
                continue

            s_xml = xml_str[s_span[0]:s_span[1]]
            new_y = sy + delta
            new_s_xml = _set_shape_y(s_xml, new_y)
            xml_str = xml_str[:s_span[0]] + new_s_xml + xml_str[s_span[1]:]
            shifted += 1
            logger.info("Pushed down id=%s ('%s'): y %d -> %d",
                         s_id, s.get("shape_name", ""), sy, new_y)

        logger.info("Pushed %d shapes down by %d EMU (%.2f in)", shifted, delta, delta / EMU_PER_INCH)

    # ── Write modified XML ──
    with open(slide_xml_path, "w", encoding="utf-8") as f:
        f.write(xml_str)
    logger.info("Wrote: %s", slide_xml_path)

    return slide_xml_path


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    if len(sys.argv) < 6:
        print("Usage: python replace_image.py <library_dir> <config_path> <slide_number> <shape_id> <new_image_path>")
        sys.exit(1)

    library_dir = sys.argv[1]
    config_path = sys.argv[2]
    slide_number = int(sys.argv[3])
    shape_id = int(sys.argv[4])
    new_image_path = sys.argv[5]

    replace_image(library_dir, config_path, slide_number, shape_id, new_image_path)
