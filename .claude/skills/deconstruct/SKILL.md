---
name: deconstruct
description: "Step 1: Deconstruct a PPTX into component library. Unzips, extracts shapes, builds metadata and manifest JSONs. Use at the start of the pipeline."
allowed-tools: Bash, Read, Glob, Grep
argument-hint: "<pptx-file>"
---

# Step 1: Deconstruct

Deconstruct the PowerPoint file into the component library.

**Input file**: $ARGUMENTS (default: "Slides Examples.pptx")

## Instructions

1. Run deconstruct with logging enabled:

```python
import sys, logging
sys.path.insert(0, "src")
logging.basicConfig(level=logging.INFO)
from deconstruct import deconstruct
deconstruct("$ARGUMENTS", "component_library", force=True)
```

2. After it completes, read the manifest to summarize what was extracted:

```python
import json
with open("component_library/manifest.json") as f:
    manifest = json.load(f)
```

3. Present a summary to the user:
   - Number of slides
   - Number of media files
   - Per slide: shape count and types (text, image, table, group)
   - Any notable shapes (tables, charts, grouped elements)

4. Tell the user to run `/generate-config` as the next step.
