---
name: read
description: "Step 1: Read a PowerPoint file — deconstructs and scans it automatically. Use when the user provides a PPTX file."
allowed-tools: Bash, Read, Glob, Grep
argument-hint: "<pptx-file>"
---

# Step 1 — Read Your PowerPoint

Automatically read and scan the user's PowerPoint file. No user input needed.

**Input file**: $ARGUMENTS (default: "Slides Examples.pptx")

## Instructions

Run both internal steps silently (deconstruct + generate_config) and present ONE combined result.

1. Show a brief status message: "Reading your PowerPoint..."

2. Run deconstruct:

```python
import sys, logging
sys.path.insert(0, "src")
logging.basicConfig(level=logging.WARNING)
from deconstruct import deconstruct
deconstruct("$ARGUMENTS", "component_library", force=True)
```

3. Immediately run generate_config (no pause, no user prompt):

```python
from generate_config import generate_config
generate_config("component_library", "configs", force=True)
```

4. Load the config to build a summary:

```python
import json, glob
config_files = [f for f in glob.glob("configs/*.json") if "_mappings" not in f]
config_path = sorted(config_files)[-1]
with open(config_path) as f:
    cfg = json.load(f)
```

5. Present a **plain English summary** to the user. NO shape IDs, NO technical details, NO XML/EMU/config jargon. Just:
   - How many slides
   - Per slide: a one-line description of what's on it (e.g., "Title slide", "Dashboard with 3 charts and some metrics", "Table with account data")
   - Count of images, text areas, and tables across the deck
   - What data files are available in data/ (if any)

Example output style:
```
Done! Here's what's in your deck:

- **Slide 1** — Title: "MACC Health & Performance"
- **Slide 2** — Dashboard with 4 charts, metrics, and cohort breakdowns
- **Slide 3** — Quality scorecard with 3 charts and performance quadrants
...

Found 57 images, 120 text areas, and 2 tables across 8 slides.

You have data ready: milestone_pipeline.xlsx, plus 11 images in your visuals folder.
```

6. Then ask: **"Which slides do you want to update, and what changes do you need?"**

## Important

- Do NOT pause between deconstruct and generate_config — run them back to back
- Do NOT show internal step names (deconstruct, generate_config)
- Do NOT show shape IDs, XML details, or config file paths
- Do NOT tell the user to run another command — just ask what they want to change
- Keep logging at WARNING level so internal details stay hidden
- If either step fails, show a friendly error message, not a stack trace
