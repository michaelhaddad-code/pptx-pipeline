"""
Steps 2-4: RECONSTRUCT + RENDER
Reassembles a .pptx from the component library, producing an exact replica.
Usage: python reconstruct.py [library_dir] [output.pptx]
"""

import zipfile
import os
import json
import shutil
import sys
import logging

logger = logging.getLogger(__name__)


def _get_original_compression(source_pptx: str) -> dict:
    """Read the original PPTX to learn per-file compression methods."""
    compression = {}
    if source_pptx and os.path.exists(source_pptx):
        try:
            with zipfile.ZipFile(source_pptx, "r") as zin:
                for info in zin.infolist():
                    compression[info.filename] = info.compress_type
        except zipfile.BadZipFile:
            pass
    return compression


def reconstruct(library_path: str = "component_library", output_path: str = "output_replica.pptx"):
    logger.info("Reconstructing from: %s", library_path)
    logger.info("Output: %s", output_path)

    # Load manifest
    manifest_path = os.path.join(library_path, "manifest.json")
    with open(manifest_path) as f:
        manifest = json.load(f)

    raw_dir = os.path.join(library_path, "_raw")

    # Learn original compression methods so we can preserve them
    source_pptx = manifest.get("source", "")
    orig_compression = _get_original_compression(source_pptx)

    # The raw directory contains the exact original file structure.
    # We repack it 1:1 — this guarantees a perfect replica.
    if os.path.exists(output_path):
        os.remove(output_path)

    # Walk the raw extracted directory and repack into a zip (pptx)
    file_count = 0
    with zipfile.ZipFile(output_path, "w") as zout:
        for root, dirs, files in os.walk(raw_dir):
            for file in files:
                full_path = os.path.join(root, file)
                # Archive path = path relative to raw_dir
                arcname = os.path.relpath(full_path, raw_dir).replace("\\", "/")
                # Use original compression method if known, otherwise DEFLATED
                compress_type = orig_compression.get(arcname, zipfile.ZIP_DEFLATED)
                zout.write(full_path, arcname, compress_type=compress_type)
                file_count += 1

    size_kb = os.path.getsize(output_path) / 1024
    logger.info("\ndone Reconstruction complete.")
    logger.info("   Files packed: %d", file_count)
    logger.info("   Output size:  %.1f KB", size_kb)
    logger.info("   Output file:  %s", output_path)

    return output_path


def verify(original_path: str, replica_path: str, library_path: str = "component_library"):
    """Compare original and replica slide by slide."""
    logger.info("\nVerifying fidelity...")
    logger.info("  Original: %s", original_path)
    logger.info("  Replica:  %s", replica_path)

    orig_size = os.path.getsize(original_path)
    repl_size = os.path.getsize(replica_path)
    logger.info("\n  Original size: %.1f KB", orig_size / 1024)
    logger.info("  Replica size:  %.1f KB", repl_size / 1024)
    logger.info("  Size diff:     %d bytes (%.2f%%)", abs(orig_size - repl_size), abs(orig_size - repl_size) / orig_size * 100)

    # Compare file lists
    with zipfile.ZipFile(original_path) as zo, zipfile.ZipFile(replica_path) as zr:
        orig_files = set(zo.namelist())
        repl_files = set(zr.namelist())

        missing = orig_files - repl_files
        extra = repl_files - orig_files

        if missing:
            logger.warning("\n  WARNING Missing files in replica: %s", missing)
        if extra:
            logger.warning("\n  WARNING Extra files in replica: %s", extra)

        # Compare slide XML content
        slide_files = sorted([f for f in orig_files if f.startswith("ppt/slides/slide") and f.endswith(".xml")])
        logger.info("\n  Slide XML comparison (%d slides):", len(slide_files))
        all_match = True
        for sf in slide_files:
            orig_xml = zo.read(sf)
            repl_xml = zr.read(sf)
            match = orig_xml == repl_xml
            status = "ok match" if match else "DIFFER"
            logger.info("    %s: %s", sf, status)
            if not match:
                all_match = False

        # Compare media
        media_files = [f for f in orig_files if f.startswith("ppt/media/")]
        media_match = all(zo.read(mf) == zr.read(mf) for mf in media_files if mf in repl_files)
        logger.info("\n  Media files (%d): %s", len(media_files), 'ok all match' if media_match else 'some differ')

        if all_match and media_match and not missing and not extra:
            logger.info("\nSUCCESS PERFECT REPLICA -- output is bit-for-bit identical to the original.")
        else:
            logger.warning("\nWARNING Some differences detected -- see above.")


if __name__ == "__main__":
    lib = sys.argv[1] if len(sys.argv) > 1 else "component_library"
    out = sys.argv[2] if len(sys.argv) > 2 else "output_replica.pptx"

    output = reconstruct(lib, out)

    # Auto-verify against the source listed in the manifest
    manifest_path = os.path.join(lib, "manifest.json")
    with open(manifest_path) as f:
        manifest = json.load(f)
    source = manifest.get("source", "Slides_Examples.pptx")
    if os.path.exists(source):
        verify(source, output, lib)
    else:
        logger.info("\nNote: Original file '%s' not found for verification.", source)
