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

1. Load the config to get all dynamic shapes:

```python
import json
with open("configs/slides_examples.json") as f:
    cfg = json.load(f)
```

2. Load the data schema from the data/ directory:

```python
import os
data_dir = "data"
data_files = os.listdir(data_dir) if os.path.exists(data_dir) else []
```

For each data file, show its contents/structure (field names, sample values, image files).

3. Present TWO lists side by side to the user:

**Available Data:**
- Scalar fields (from CSVs/JSONs): field name, sample value
- Tabular sources: file name, columns, row count
- Images: file names

**Shapes Needing Mapping (per slide):**
- Shape id, name, category, text preview
- For images: current image target, nearby text labels

### Phase 2: Mapping Conversation

4. Go slide by slide. For each slide with dynamic shapes:
   - Show the shapes and ask the user which data field each one maps to
   - For images: ask which replacement image goes where
   - For tables: ask which tabular data source feeds the table
   - Confirm each mapping before moving on

5. If the user provides a screenshot of the slide, use it to understand the visual layout and suggest mappings.

6. The user may say things like:
   - "shape 12 maps to revenue_field"
   - "the top-right chart gets replaced with tom_jerry.png"
   - "skip this slide"
   - "that table gets the execution_data.csv"

### Phase 3: Write Mappings

7. Once all mappings are agreed, write them into the config:

```python
import json
with open("configs/slides_examples.json") as f:
    cfg = json.load(f)

# For each mapping, set data_field on the shape:
# shape["data_field"] = "field_name"
# For images: image["source"] = "visual:filename" or path

with open("configs/slides_examples.json", "w") as f:
    json.dump(cfg, f, indent=2)
```

8. Optionally save the mapping as a recipe for reuse:

```python
# Save to recipes/slides_examples_recipe.json
```

9. Present a final summary of all mappings and tell the user to run `/update-config` next.

## Important

- ALWAYS pause and ask the user before writing mappings
- Show what you're about to write and get confirmation
- If the user is unsure about a mapping, skip it and come back later
- Keep track of unmapped shapes and remind the user at the end
