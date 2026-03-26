"""
One-command entry point for updating a PowerPoint with new data.

Usage:
    python update.py "My Presentation.pptx" data/
    python update.py "My Presentation.pptx" data/ --output result.pptx
    python update.py "My Presentation.pptx" data/ --mappings configs/my_mappings.json

The user never needs to know about config files, shape IDs, or XML.
"""

import argparse
import json
import logging
import os
import re
import sys
import glob as _glob
from pathlib import Path

# Ensure emoji/unicode can print on Windows terminals
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ── Make src/ importable ──────────────────────────────────────────────
_src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src")
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)

from deconstruct import deconstruct
from generate_config import generate_config
from update_config import update_config, load_data_sources
from inject import inject
from reconstruct import reconstruct


# ═══════════════════════════════════════════════════════════════════════
# Friendly logger — translates internal messages to plain English
# ═══════════════════════════════════════════════════════════════════════

class FriendlyFilter(logging.Filter):
    """Intercepts pipeline log records and rewrites them for a non-technical user."""

    # Patterns to suppress entirely (too internal to show)
    SUPPRESS = [
        re.compile(r"^Reset geometries"),
        re.compile(r"^  ok "),
        re.compile(r"^\s*$"),
        re.compile(r"^={5,}"),
        re.compile(r"^  STEP:"),
        re.compile(r"^  VERIFY"),
        re.compile(r"^Applied \d+ mappings"),
        re.compile(r"^Mappings file is empty"),
        re.compile(r"^No mappings file found"),
        re.compile(r"^Updating config:"),
        re.compile(r"^Data directory:"),
        re.compile(r"^Loading data sources"),
        re.compile(r"^  Total fields loaded"),
        re.compile(r"^Resolving tokens"),
        re.compile(r"^Config generated:"),
        re.compile(r"^Generating config for:"),
        re.compile(r"^Config already exists"),
        re.compile(r"^Config exists for"),
        re.compile(r"^  Config:"),
        re.compile(r"^  Use --force"),
        re.compile(r"^  Slide \d+: \d+ shapes"),
        re.compile(r"^Run Step 3"),
        re.compile(r"^Deconstructing:"),
        re.compile(r"^Output library:"),
        re.compile(r"^  Skipping backup"),
        re.compile(r"^Backup created:"),
        re.compile(r"^ok "),
        re.compile(r"^Reconstructing from:"),
        re.compile(r"^Output:"),
        re.compile(r"^   Files packed:"),
        re.compile(r"^   Output size:"),
        re.compile(r"^   Output file:"),
        re.compile(r"^\[ok\]"),
        re.compile(r"^\[skip\]"),
        re.compile(r"^\[layout\]"),
        re.compile(r"^\[path\]"),
        re.compile(r"^\[DONE\]"),
        re.compile(r"^   Slides modified:"),
        re.compile(r"^   Total replacements:"),
        re.compile(r"^Verifying fidelity"),
        re.compile(r"^  Original:"),
        re.compile(r"^  Replica:"),
        re.compile(r"^  Original size:"),
        re.compile(r"^  Replica size:"),
        re.compile(r"^  Size diff:"),
        re.compile(r"^  Slide XML comparison"),
        re.compile(r"^    ppt/slides"),
        re.compile(r"^  Media files"),
        re.compile(r"^SUCCESS PERFECT"),
        re.compile(r"^done "),
        re.compile(r"^PIPELINE COMPLETE"),
        re.compile(r"^  PIPELINE COMPLETE"),
        re.compile(r"^Removed backup:"),
        re.compile(r"^Cleaned \d+ backup"),
        re.compile(r"^Removed clean backup:"),
        re.compile(r"^  Skipping backup"),
        re.compile(r"^   Slides:"),
        re.compile(r"^   Media:"),
        re.compile(r"^   Library:"),
        re.compile(r"^  Slide \d+:"),
        re.compile(r"^  Total fields loaded"),
        re.compile(r"^Total fields loaded"),
        re.compile(r"^   Tokens resolved"),
        re.compile(r"^Tokens resolved"),
        re.compile(r"^   Saved:"),
        re.compile(r"^Saved:"),
        re.compile(r"^Injecting from config"),
        re.compile(r"^Library:"),
        re.compile(r"^  Clean backup"),
        re.compile(r"^Clean backup"),
        re.compile(r"^  \[ Slide \d+"),
        re.compile(r"^\[ Slide \d+"),
        re.compile(r"^      ->"),
        re.compile(r"^    ->"),
        re.compile(r"^  ->"),
        re.compile(r"^Config updated"),
        re.compile(r"^done Config updated"),
        re.compile(r"^   Slides modified"),
        re.compile(r"^   Total replacements"),
        re.compile(r"^   Files packed"),
        re.compile(r"^   Output size"),
        re.compile(r"^   Output file"),
        re.compile(r"^   Slides:"),
        re.compile(r"^   Media:"),
        re.compile(r"^   Library:"),
        re.compile(r"Skipping backup"),
        re.compile(r"^  Slide \d+: \d+"),
        re.compile(r"^Slide \d+: \d+"),
    ]

    # Patterns to rewrite to friendly messages
    REWRITE = [
        # Warnings the user should see in plain English
        (re.compile(r"WARNING Mapping references unknown slide: (.+)"),
         lambda m: f"  ⚠️  Could not find slide {m.group(1)} in your PowerPoint — skipping that mapping."),
        (re.compile(r"WARNING Mapping references unknown shape_id=\S+ on (.+)"),
         lambda m: f"  ⚠️  A placeholder on {m.group(1)} could not be found — it may have been deleted. Skipping."),
        (re.compile(r"WARNING Field not found in data: '(.+)'"),
         lambda m: f"  ⚠️  Could not find \"{m.group(1)}\" in your data files — that field will be left as-is."),
        (re.compile(r"WARNING No screenshot found for '(.+)'"),
         lambda m: f"  ⚠️  No image found for \"{m.group(1)}\" — that placeholder will be left empty."),
        (re.compile(r"WARNING Unresolved fields: (\d+)"),
         lambda m: f"  ⚠️  {m.group(1)} field(s) could not be filled — check that your data file has all the columns you need."),
        (re.compile(r"WARNING Missing files in replica"),
         lambda m: "  ⚠️  Some parts of the original file could not be included — the output may look slightly different."),
        (re.compile(r"WARNING .+exceeds slide height"),
         lambda m: "  ⚠️  Some content may extend past the bottom of a slide — consider using a smaller image or less data."),
        (re.compile(r"\[ERROR\] (.+)"),
         lambda m: f"  ❌ Something went wrong: {m.group(1)}"),
        (re.compile(r"\[FAILED\] (.+)"),
         lambda m: f"  ❌ {m.group(1)}"),
        (re.compile(r"ERROR (.+)"),
         lambda m: f"  ❌ {m.group(1)}"),
        (re.compile(r"\[!\] Shape.*?id=(\d+)"),
         lambda m: f"  ⚠️  A placeholder could not be updated (internal ref {m.group(1)}) — skipping."),
        (re.compile(r"\[!\] table .+ no <a:tbl>"),
         lambda _: "  ⚠️  A table placeholder could not be updated — skipping."),
        (re.compile(r"\[!\] table .+ failed to parse"),
         lambda _: "  ⚠️  Table data could not be read — check your data file format."),
        (re.compile(r"\[!\] table .+ need at least 2 rows"),
         lambda _: "  ⚠️  A table needs at least a header row and one data row — skipping."),
    ]

    def filter(self, record):
        msg = record.getMessage().strip()
        if not msg:
            return False

        # Check rewrite list first — these are warnings we want the user to see
        for pattern, rewriter in self.REWRITE:
            m = pattern.search(msg)
            if m:
                record.msg = rewriter(m)
                record.args = ()
                return True

        # Suppress everything else from pipeline modules.
        # Only our own print() statements (which bypass logging) reach the user.
        return False


class FriendlyFormatter(logging.Formatter):
    """Strips log level prefixes — the emoji/text in the message is enough."""
    def format(self, record):
        return record.getMessage()


def setup_friendly_logging(verbose=False):
    """Replace the default logging with human-friendly output."""
    if verbose:
        # In verbose mode, show everything (for debugging)
        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%H:%M:%S",
        )
        return

    # Friendly mode: filter + clean formatting
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(logging.DEBUG)
    handler.addFilter(FriendlyFilter())
    handler.setFormatter(FriendlyFormatter())

    root.handlers.clear()
    root.addHandler(handler)


# ═══════════════════════════════════════════════════════════════════════
# Interactive mapping
# ═══════════════════════════════════════════════════════════════════════

def _describe_category(cat):
    """Return a user-friendly label for a shape category."""
    return {
        "text": "Text",
        "image": "Image",
        "table": "Table",
        "group": "Group",
    }.get(cat, cat.title())


def _describe_shape(shape):
    """Return a one-line description of a shape for the user."""
    cat = _describe_category(shape.get("category", ""))
    preview = shape.get("text_preview", "").strip()
    name = shape.get("shape_name", "")

    if preview:
        # Truncate long previews
        if len(preview) > 60:
            preview = preview[:57] + "..."
        return f'{cat}: "{preview}"'
    elif cat == "Image":
        return f'Image placeholder ({name})'
    elif cat == "Table":
        grid = shape.get("table_grid", {})
        cols = grid.get("col_count", "?")
        rows = grid.get("row_count", "?")
        return f'Table ({cols} columns, {rows} rows)'
    else:
        return f'{cat}: {name}'


def _collect_data_choices(data_dir):
    """Scan the data directory and return available data fields and image files."""
    fields = []   # (display_name, data_field_value)
    images = []   # (display_name, filename)

    # CSV / JSON / XLSX data columns
    data = load_data_sources(data_dir)
    for key, val in data.items():
        if key.startswith("_"):
            continue
        if key.startswith("visual:") or key.startswith("screenshot_dir:"):
            continue
        if isinstance(val, list):
            # Tabular data — show the key
            fields.append((f"{key} (table, {len(val)} rows)", key))
        else:
            fields.append((f"{key}", key))

    # Image files in data/visuals/
    visuals_dir = os.path.join(data_dir, "visuals")
    if os.path.exists(visuals_dir):
        for img_path in sorted(_glob.glob(os.path.join(visuals_dir, "*"))):
            if os.path.isfile(img_path):
                fname = os.path.basename(img_path)
                stem = Path(img_path).stem
                images.append((fname, stem))

    return fields, images


def _get_runs_from_xml(library_path, slide_num, shape_id):
    """Read the XML to find text runs for a shape — used to show which run to target."""
    slide_xml_path = os.path.join(library_path, f"slides/slide_{slide_num}/slide.xml")
    if not os.path.exists(slide_xml_path):
        return []
    with open(slide_xml_path, encoding="utf-8") as f:
        xml = f.read()

    # Find the shape by cNvPr id
    pattern = rf'<p:sp\b[^>]*>.*?<p:cNvPr[^>]*\bid="{shape_id}"[^>]*/?>.*?</p:sp>'
    match = re.search(pattern, xml, re.DOTALL)
    if not match:
        return []

    runs = re.findall(r'<a:r>(.*?)</a:r>', match.group(), re.DOTALL)
    result = []
    for r in runs:
        t = re.search(r'<a:t>(.*?)</a:t>', r, re.DOTALL)
        if t:
            text = t.group(1).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
            result.append(text)
    return result


def interactive_mapping(config_path, data_dir, library_path="component_library"):
    """Walk the user through mapping shapes to data, slide by slide."""
    with open(config_path, encoding="utf-8") as f:
        config = json.load(f)

    # Suppress pipeline logging during data scan
    prev_level = logging.getLogger().level
    logging.getLogger().setLevel(logging.CRITICAL)
    fields, images = _collect_data_choices(data_dir)
    logging.getLogger().setLevel(prev_level)

    mappings = []

    print("\n" + "─" * 50)
    print("  Step 3: Tell us what goes where")
    print("─" * 50)
    print()
    print("I'll show you what's on each slide.")
    print("Pick which data should replace each item,")
    print("or press Enter to skip.\n")

    for slide_key in sorted(config["slides"].keys(), key=lambda k: config["slides"][k]["slide_number"]):
        slide = config["slides"][slide_key]
        slide_num = slide["slide_number"]
        shapes = slide["shapes"]

        # Summarize what's on this slide
        text_shapes = [s for s in shapes if s["category"] == "text" and s.get("text_preview", "").strip()]
        image_shapes = [s for s in shapes if s["category"] == "image"]
        table_shapes = [s for s in shapes if s["category"] == "table"]

        mappable = text_shapes + table_shapes + image_shapes
        if not mappable:
            continue

        counts = []
        if text_shapes:
            counts.append(f"{len(text_shapes)} text field{'s' if len(text_shapes) != 1 else ''}")
        if table_shapes:
            counts.append(f"{len(table_shapes)} table{'s' if len(table_shapes) != 1 else ''}")
        if image_shapes:
            counts.append(f"{len(image_shapes)} image{'s' if len(image_shapes) != 1 else ''}")

        print(f"━━━ Slide {slide_num} — found {', '.join(counts)} ━━━\n")

        # Ask user if they want to map this slide
        proceed = input(f"  Map this slide? (y/n, default: y): ").strip().lower()
        if proceed == "n":
            print()
            continue

        for shape in mappable:
            cat = shape.get("category", "")
            desc = _describe_shape(shape)

            print(f"\n  [{slide_key.replace('_', ' ').title()} — {_describe_category(cat)}]")
            print(f"  Currently: {desc}")

            # For text shapes with multiple runs, show them
            if cat == "text":
                runs = _get_runs_from_xml(library_path, slide_num, shape["shape_id"])
                if len(runs) > 1:
                    print(f"  Text parts:")
                    for i, run_text in enumerate(runs):
                        print(f"    Part {i}: \"{run_text}\"")

            # Build choices
            choices = []
            if cat == "image":
                if not images:
                    print("  (No image files found in your data/visuals/ folder)")
                    print()
                    continue
                for i, (display, stem) in enumerate(images, 1):
                    choices.append((str(i), display, stem))
            elif cat == "table":
                table_fields = [(d, f) for d, f in fields if "table" in d.lower() or "row" in d.lower()]
                if not table_fields:
                    table_fields = fields
                for i, (display, field) in enumerate(table_fields, 1):
                    choices.append((str(i), display, field))
            else:
                for i, (display, field) in enumerate(fields, 1):
                    choices.append((str(i), display, field))

            # Also offer literal text entry for text shapes
            if cat == "text":
                choices.append(("t", "Type a value", None))

            choices.append(("", "Skip (keep as-is)", None))

            print()
            for num, display, _ in choices:
                if num:
                    print(f"    {num}. {display}")
                else:
                    print(f"    Enter. {display}")

            choice = input("\n  Your choice: ").strip()

            if not choice:
                continue

            # Handle literal text entry
            if choice.lower() == "t" and cat == "text":
                value = input("  Type the value: ").strip()
                if not value:
                    continue

                target_run = None
                runs = _get_runs_from_xml(library_path, slide_num, shape["shape_id"])
                if len(runs) > 1:
                    run_choice = input(f"  Which part to replace? (0-{len(runs)-1}, or Enter for all): ").strip()
                    if run_choice.isdigit() and 0 <= int(run_choice) < len(runs):
                        target_run = int(run_choice)

                mapping = {
                    "slide": slide_key,
                    "shape_id": shape["shape_id"],
                    "shape_name": shape["shape_name"],
                    "type": "text",
                    "data_field": f"literal:{value}",
                }
                if target_run is not None:
                    mapping["target_run"] = target_run
                mappings.append(mapping)
                print(f"  ✅ Mapped → \"{value}\"")
                continue

            # Match numbered choice
            matched = None
            for num, display, field_val in choices:
                if num and choice == num:
                    matched = (display, field_val)
                    break

            if not matched:
                print("  (Skipped)")
                continue

            display, field_val = matched

            if cat == "image":
                # Find the full filename
                source_file = None
                for fname, stem in images:
                    if stem == field_val:
                        source_file = fname
                        break
                mappings.append({
                    "slide": slide_key,
                    "shape_id": shape["shape_id"],
                    "shape_name": shape["shape_name"],
                    "type": "image",
                    "data_field": f"visual:{field_val}",
                    "source": source_file or field_val,
                })
                print(f"  ✅ Mapped → {source_file}")
            elif cat == "table":
                mappings.append({
                    "slide": slide_key,
                    "shape_id": shape["shape_id"],
                    "shape_name": shape["shape_name"],
                    "type": "table",
                    "data_field": field_val,
                })
                print(f"  ✅ Mapped → {display}")
            else:
                # Text field
                target_run = None
                runs = _get_runs_from_xml(library_path, slide_num, shape["shape_id"])
                if len(runs) > 1:
                    run_choice = input(f"  Which part to replace? (0-{len(runs)-1}, or Enter for all): ").strip()
                    if run_choice.isdigit() and 0 <= int(run_choice) < len(runs):
                        target_run = int(run_choice)

                mapping = {
                    "slide": slide_key,
                    "shape_id": shape["shape_id"],
                    "shape_name": shape["shape_name"],
                    "type": "text",
                    "data_field": field_val,
                }
                if target_run is not None:
                    mapping["target_run"] = target_run
                mappings.append(mapping)
                print(f"  ✅ Mapped → {display}")

        print()

    return mappings


# ═══════════════════════════════════════════════════════════════════════
# Main pipeline
# ═══════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Update a PowerPoint file with new data.",
        epilog="Example: python update.py \"My Slides.pptx\" data/",
    )
    parser.add_argument("pptx", help="Path to your PowerPoint file")
    parser.add_argument("data_dir", nargs="?", default="data",
                        help="Folder containing your data files (default: data/)")
    parser.add_argument("--output", "-o", default=None,
                        help="Output file name (default: <name>_updated.pptx)")
    parser.add_argument("--mappings", "-m", default=None,
                        help="Use a saved mappings file instead of interactive mapping")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Show detailed technical output (for troubleshooting)")
    args = parser.parse_args()

    setup_friendly_logging(verbose=args.verbose)

    # Validate inputs
    if not os.path.exists(args.pptx):
        print(f"  ❌ File not found: {args.pptx}")
        sys.exit(1)

    if not os.path.exists(args.data_dir):
        print(f"  ❌ Data folder not found: {args.data_dir}")
        sys.exit(1)

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        base = os.path.splitext(os.path.basename(args.pptx))[0]
        output_dir = os.path.dirname(args.pptx) or "."
        output_path = os.path.join(output_dir, f"{base}_updated.pptx")

    # Ensure output directory exists
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # ── Step 1: Read your PowerPoint ──────────────────────────────
    print("\n  📖 Reading your PowerPoint...")
    deconstruct(args.pptx, "component_library", force=True, no_backup=True)
    print("  ✅ Reading your PowerPoint... done")

    # ── Step 2: Scan your slides ──────────────────────────────────
    print("\n  🔍 Scanning your slides...")
    config_path = generate_config(
        library_path="component_library",
        configs_dir="configs",
        force=True,
    )
    print("  ✅ Scanning your slides... done")

    # ── Step 3: Tell us what goes where ───────────────────────────
    deck_name = os.path.splitext(os.path.basename(args.pptx))[0].replace(" ", "_").lower()
    config_path = os.path.join("configs", f"{deck_name}.json")
    mappings_path = args.mappings or os.path.join("configs", f"{deck_name}_mappings.json")

    if args.mappings and os.path.exists(args.mappings):
        print(f"\n  📋 Using saved mappings: {args.mappings}")
    else:
        mappings = interactive_mapping(config_path, args.data_dir)

        if not mappings:
            print("\n  ℹ️  No mappings created. Your PowerPoint will be unchanged.")
            # Still write an empty mappings file so the pipeline doesn't error
            mappings_data = {"deck": deck_name, "mappings": []}
        else:
            mappings_data = {"deck": deck_name, "mappings": mappings}
            print(f"\n  ✅ {len(mappings)} mapping(s) saved.")

        # Write mappings file
        with open(mappings_path, "w", encoding="utf-8") as f:
            json.dump(mappings_data, f, indent=2)

    # ── Step 4: Load your data & update slides ────────────────────
    print("\n  📊 Loading your data...")
    update_config(config_path, args.data_dir, mappings_path=mappings_path)
    print("  ✅ Loading your data... done")

    print("\n  ✏️  Updating your slides...")
    inject(config_path, "component_library")
    print("  ✅ Updating your slides... done")

    # ── Step 5: Build your new PowerPoint ─────────────────────────
    print("\n  📦 Building your new PowerPoint...")
    reconstruct("component_library", output_path)
    print("  ✅ Building your new PowerPoint... done")

    # ── Done ──────────────────────────────────────────────────────
    size_kb = os.path.getsize(output_path) / 1024
    print("\n" + "─" * 50)
    print(f"  ✅ Your updated PowerPoint is ready!")
    print(f"     {output_path} ({size_kb:.0f} KB)")
    print("─" * 50 + "\n")


if __name__ == "__main__":
    main()
