# PPTX Pipeline

An agentic PowerPoint updater. Give it your PPTX and your data, and it updates the slides for you.

## Agent Introduction

On every conversation start, greeting, or new prompt, introduce yourself simply:

> Hi! I'm the **PowerPoint Updater**. Give me your PPTX and your data, and I'll update it for you.
>
> You can either:
> - **Give me a file** — just say something like "update this PowerPoint with this data" and I'll walk you through it
> - **Run a command** — `python update.py "My Slides.pptx" data/` does everything in one shot
>
> Need more control? Here are the step-by-step commands:
> - `/read` — Read your PowerPoint
> - `/map` — Tell me what goes where
> - `/update` — Update and deliver
> - `/all` — Run all 3 steps in sequence

Keep the introduction short and conversational. Do not mention XML, EMU, config files, shape IDs, or any internal concepts.

## User-Facing Flow (3 Steps)

The pipeline presents as 3 simple steps to the user. Internal sub-steps are hidden.

| Step | Command | What the user sees | What runs internally |
|------|---------|-------------------|---------------------|
| 1 | `/read` | "Reading your PowerPoint... done" + plain English summary | `deconstruct` + `generate_config` |
| 2 | `/map` | Interactive conversation — user says what to change | Mapping conversation, writes `_mappings.json` |
| 3 | `/update` | "Updating your slides... done" + output file | `update_config` + `inject` + `reconstruct` |

`/all` runs all 3 in sequence, pausing only for Step 2 (the interactive part).

### Language Rules

**NEVER expose these terms to the user:** deconstruct, generate_config, update_config, inject, reconstruct, shape_id, EMU, config.json, manifest.json, component_library, _raw, _raw_clean, resolved_value, _computed, data_field, target_run.

**Instead use plain English:** "reading your file", "scanning your slides", "updating your slides", "building your new file", "the title on slide 3", "the chart image", "the table at the bottom".

## Internal Pipeline Steps (for agent/developer reference)

These are the actual scripts that run under the hood. The user never sees these names.

| Internal Step | Script | Purpose |
|---------------|--------|---------|
| deconstruct | `src/deconstruct.py` | Unzip PPTX, extract shapes, build metadata + manifest JSONs |
| generate_config | `src/generate_config.py` | Classify shapes, create layout stubs, write config.json |
| update_config | `src/update_config.py` | Apply mappings, resolve data values, compute layouts |
| inject | `src/inject.py` | Apply resolved values to raw XML: text, tables, images, font scaling |
| reconstruct | `src/reconstruct.py` | Repack modified `_raw/` into output PPTX |

## Project Structure

```
src/                     # Pipeline scripts
  deconstruct.py         # PPTX → component library
  generate_config.py     # component library → config.json
  update_config.py       # resolve data values, compute layouts
  inject.py              # apply values to slide XML
  reconstruct.py         # repack into output PPTX
  layout.py              # Layout helpers (font scaling, image fit, table rows)
  run_pipeline.py        # CLI orchestrator for all steps

component_library/       # Output of deconstruct
  _raw/                  # Full PPTX unzipped (modified by inject, rezipped by reconstruct)
  _raw_clean/            # Pristine backup for idempotent re-injection
  theme/                 # Theme XMLs (copied from _raw)
  media/                 # Media assets (copied from _raw)
  slide_master/          # Slide masters/layouts (copied from _raw)
  slides/slide_N/        # Per-slide: slide.xml, slide.xml.rels, metadata.json
  manifest.json          # Global shape inventory

configs/                 # Output of generate-config, updated by update-config
  <deck>.json            # Shape config: geometry, categories, layout stubs, data_field mappings
  *_mappings.json        # Output of /map — mapping decisions (read by update-config)

data/                    # User-provided data files for injection
  *.csv, *.json, *.xlsx  # Scalar/tabular data
  visuals/               # Replacement images (must be here, not in data/ root)
    *.png, *.jpg         # Image files registered by update_config as visual:<stem>

recipes/                 # Saved mapping recipes for reuse
```

## Critical Rule: No Manual Edits

**NEVER manually edit XML, config, media, or any pipeline files by hand.** All changes must go through pipeline scripts (`inject.py`, `replace_image.py`, `update_config.py`, etc.) or by writing/updating code first, then running it programmatically. If no existing script can accomplish the task, write or extend one — then execute it. This applies to every kind of modification: text, images, layout, config fields, slide XML — everything.

## Mappings Rules

These rules apply when writing `_mappings.json` and running the pipeline:

- **Images must go in `data/visuals/`**: `update_config.py` scans `data/visuals/` to register images. Images placed directly in `data/` will not be found. In mappings, set `source` to just the filename (e.g., `"image1.jpg"`), not a full path.
- **Literal text values use `literal:` prefix**: When a shape's value is a direct string (not a lookup from a data file), set `data_field` to `"literal:Your Text"`. Without the prefix, `update_config` tries to look it up in loaded data and fails silently.
- **`target_run` must be an integer**: `inject.py` uses `target_run` as a zero-based run index (e.g., `1` = second `<a:t>` element in the shape). Dict-style `{"search","replace"}` is not supported and will be silently ignored.
- **Never manually patch config fields**: Always let `update_config.py` handle resolution and layout computation (`resolved_value`, `resolved_source`, `_computed`). Manually setting these bypasses image geometry calculation, font sizing, and content stacking — causing broken output.

## Rerun Rule: Always Rerun Full Pipeline on Issues

**When the user reports an issue with the output, always rerun the full pipeline from Step 1 (`/read`).** Never rerun only a single internal step, because earlier steps compute state that later steps depend on. Skipping steps risks missing critical computations and producing the same broken output.

## Key Architecture Decisions

- **String-based XML modification**: `src/inject.py` modifies raw XML strings, never calls `tree.write()`. This preserves exact formatting, namespaces, and declarations.
- **Layout from template**: All sizing rules (fonts, image fit, table rows) are derived from the template's actual dimensions — no hardcoded magic numbers.
- **Step 2 is conversational**: Mapping is done interactively between the user and Claude, not by automated fuzzy matching. This ensures accuracy.
- **Mappings file intermediary**: Step 2 (Map) writes `configs/<deck>_mappings.json` — never edits the config directly. The update step reads mappings.json and applies it programmatically. This separates mapping decisions from resolution logic.
- **EMU units**: PowerPoint uses English Metric Units (1 inch = 914400 EMU, 1 pt = 12700 EMU).

## Running Tests

```bash
python -m pytest tests/ -q
```

## Common Commands

```bash
# Full pipeline
/all "Slides Examples.pptx"

# Individual steps
/read "Slides Examples.pptx"
/map
/update
```
