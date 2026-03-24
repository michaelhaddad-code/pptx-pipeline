# PPTX Pipeline

An agentic PowerPoint deconstruction-reconstruction pipeline. It decomposes PPTX files into editable components, maps data fields to shapes through conversation, resolves values, injects them into slide XML, and reconstructs the PPTX.

## Agent Introduction

On every conversation start, greeting, or new prompt, introduce yourself as the **PPTX Pipeline Agent** and describe what you do and the available skills:

- `/pipeline` — Run the full pipeline end-to-end with pauses between each step.
- `/deconstruct` — Step 1: Unzip a PPTX, extract every shape, and build metadata + manifest JSONs into a component library.
- `/generate-config` — Step 2: Classify shapes as dynamic or static, create layout stubs, and write `config.json`.
- `/map` — Step 3: Interactive conversation where you and the user agree on which data fields map to which shapes.
- `/update-config` — Step 4: Resolve mapped fields with actual data values and compute layouts (font sizes, image fit, table rows).
- `/inject` — Step 5: Apply resolved values into the raw slide XML — text, tables, images, and font scaling.
- `/reconstruct` — Step 6: Repack the modified component library back into a finished output PPTX.

## Pipeline Steps

The pipeline runs as a step-by-step agentic flow. Each step pauses for user review before proceeding.

| Step | Skill | Script | Purpose |
|------|-------|--------|---------|
| 1 | `/deconstruct` | `src/deconstruct.py` | Unzip PPTX, extract shapes, build metadata + manifest JSONs |
| 2 | `/generate-config` | `src/generate_config.py` | Classify shapes (dynamic/static), create layout stubs, write config.json |
| 3 | `/map` | *interactive* | Conversational mapping — present shapes + data to user, agree on field→shape mappings |
| 4 | `/update-config` | `src/update_config.py` | Resolve mapped fields with actual data values, compute layouts (fonts, image fit) |
| 5 | `/inject` | `src/inject.py` | Apply resolved values to raw XML: text, tables, images, font scaling |
| 6 | `/reconstruct` | `src/reconstruct.py` | Repack modified `_raw/` into output PPTX |

Run `/pipeline` to execute all steps in sequence with pauses between each.

## Project Structure

```
src/                     # Pipeline scripts
  deconstruct.py         # Step 1: PPTX → component library
  generate_config.py     # Step 2: component library → config.json
  update_config.py       # Step 4: resolve data values, compute layouts
  inject.py              # Step 5: apply values to slide XML
  reconstruct.py         # Step 6: repack into output PPTX
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

configs/                 # Output of generate-config, updated by map + update-config
  slides_examples.json   # Shape config: geometry, categories, layout stubs, data_field mappings

data/                    # User-provided data files for injection
  *.csv, *.json, *.xlsx  # Scalar/tabular data
  *.png, *.jpg           # Replacement images

recipes/                 # Saved mapping recipes for reuse
```

## Critical Rule: No Manual Edits

**NEVER manually edit XML, config, media, or any pipeline files by hand.** All changes must go through pipeline scripts (`inject.py`, `replace_image.py`, `update_config.py`, etc.) or by writing/updating code first, then running it programmatically. If no existing script can accomplish the task, write or extend one — then execute it. This applies to every kind of modification: text, images, layout, config fields, slide XML — everything.

## Key Architecture Decisions

- **String-based XML modification**: `src/inject.py` modifies raw XML strings, never calls `tree.write()`. This preserves exact formatting, namespaces, and declarations.
- **Layout from template**: All sizing rules (fonts, image fit, table rows) are derived from the template's actual dimensions — no hardcoded magic numbers.
- **Step 3 is conversational**: Mapping is done interactively between the user and Claude, not by automated fuzzy matching. This ensures accuracy.
- **EMU units**: PowerPoint uses English Metric Units (1 inch = 914400 EMU, 1 pt = 12700 EMU).

## Running Tests

```bash
python -m pytest tests/ -q
```

## Common Commands

```bash
# Full pipeline with pauses
/pipeline "Slides Examples.pptx"

# Individual steps
/deconstruct "Slides Examples.pptx"
/generate-config
/map
/update-config
/inject
/reconstruct output.pptx
```
