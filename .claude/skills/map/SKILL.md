---
name: map
description: "Step 3: Interactive mapping — conversationally map data fields to shapes. Replaces auto_map. Use after /generate-config."
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
---

# Step 3: Interactive Mapping

Map data fields to presentation shapes through conversation with the user.

## Instructions

This step is **interactive**. Do NOT auto-complete — work with the user step by step.

### Phase 1: Inventory

1. Find the config file (look in configs/ for the most recent one):

```python
import json, glob
config_files = glob.glob("configs/*.json")
# Exclude mappings files
config_files = [f for f in config_files if "_mappings" not in f and "mappings.json" not in f]
```

2. Load the config to get all shapes (read-only — do NOT modify the config):

```python
with open(config_path) as f:
    cfg = json.load(f)
```

3. Load the data schema from the data/ directory:

```python
import os
data_dir = "data"
data_files = os.listdir(data_dir) if os.path.exists(data_dir) else []
```

For each data file, show its contents/structure (field names, sample values, image files).

4. Present TWO lists side by side to the user:

**Available Data:**
- Scalar fields (from CSVs/JSONs): field name, sample value
- Tabular sources: file name, columns, row count
- Images: file names

**Shapes Needing Mapping (per slide):**
- Shape id, name, category, text preview
- For images: current image target, nearby text labels

### Phase 2: Mapping Conversation

5. Go slide by slide. For each slide with shapes:
   - Show the shapes and ask the user which data field each one maps to
   - For images: ask which replacement image goes where
   - For tables: ask which tabular data source feeds the table
   - Confirm each mapping before moving on

6. If the user provides a screenshot of the slide, use it to understand the visual layout and suggest mappings.

7. The user may say things like:
   - "shape 12 maps to revenue_field"
   - "the top-right chart gets replaced with tom_jerry.png"
   - "skip this slide"
   - "that table gets the execution_data.csv"
   - "replace image 31 with my_image.jpg"

### Phase 3: Write Mappings File

8. Once all mappings are agreed, write them to `configs/<deck>_mappings.json`:

```python
import json
from datetime import datetime

mappings = {
    "deck": "<deck_name>",
    "created": datetime.now().isoformat(),
    "mappings": [
        # Text mapping example:
        {
            "slide": "slide_2",
            "shape_id": "5",
            "type": "text",
            "data_field": "report_title"
        },
        # Image mapping example:
        {
            "slide": "slide_3",
            "shape_id": "32",
            "type": "image",
            "source": "visual:my_image"
        },
        # Table mapping example:
        {
            "slide": "slide_7",
            "shape_id": "7",
            "type": "table",
            "data_field": "execution_data"
        }
    ]
}

with open("configs/<deck>_mappings.json", "w") as f:
    json.dump(mappings, f, indent=2)
```

**CRITICAL: Do NOT modify the config file. Only write mappings.json.**

The `type` field must be one of: `"text"`, `"image"`, `"table"`.
- Text and table mappings use `data_field` (data key, dot-path, or `literal:...`).
- Image mappings use `source` (`visual:name`, `screenshot:dir/key`, or bare filename path).
- Presence in mappings.json implies the shape is dynamic — no `is_dynamic` field needed.

9. Optionally save the mapping as a recipe for reuse:

```python
import shutil
shutil.copy("configs/<deck>_mappings.json", "recipes/<deck>_recipe.json")
```

10. Present a final summary of all mappings and tell the user to run `/update-config` next.

## Important

- **NEVER edit the config file** — only write mappings.json
- ALWAYS pause and ask the user before writing mappings
- Show what you're about to write and get confirmation
- If the user is unsure about a mapping, skip it and come back later
- Keep track of unmapped shapes and remind the user at the end
- The config is read-only during this step — used only to display shape info
