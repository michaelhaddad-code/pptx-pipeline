---
name: map
description: "Step 2: Interactive mapping — tell the agent what to change in your slides. The conversational step between reading and updating."
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
---

# Step 2 — Tell Me What Goes Where

Interactive mapping conversation. The user describes what they want changed, and the agent figures out which shapes and data to connect.

## Instructions

This step is **interactive**. Do NOT auto-complete — work with the user step by step.

### Phase 1: Show What's Available

1. Find the config file:

```python
import json, glob
config_files = [f for f in glob.glob("configs/*.json") if "_mappings" not in f]
config_path = sorted(config_files)[-1]
with open(config_path) as f:
    cfg = json.load(f)
```

2. Load the data from data/:

```python
import os
data_dir = "data"
data_files = os.listdir(data_dir) if os.path.exists(data_dir) else []
```

3. Present what's available in **plain English**:
   - Per slide: describe what's on it (text content, images, tables) — NO shape IDs in the initial presentation
   - Available data: file names, field names, sample values, image files
   - Keep it conversational: "Slide 3 has a title, some quality metrics, quadrant labels, and 3 chart images"

### Phase 2: Mapping Conversation

4. Let the user describe what they want in natural language. They might say:
   - "Replace the title date with 'March 15'"
   - "Swap the bar chart on slide 3 with my new one"
   - "Put the pipeline data in that table on slide 8"
   - "Change the quadrant numbers to 2, 3, 4, 5"
   - "Skip slide 4"

5. When the user references a shape, use your knowledge of the config to find the right shape internally. If ambiguous, describe the options in plain English and ask the user to clarify.

6. If the user provides a screenshot, use it to understand the layout and help identify shapes.

7. For images the user wants to replace:
   - Copy the image to `data/visuals/` if it's not already there
   - Use the filename (without extension) as the visual identifier

8. For text that needs XML run targeting (e.g., replacing just "xxx" in a longer string), read the slide XML to find the correct target_run index. Do this silently — the user doesn't need to know about runs.

### Phase 3: Confirm and Write Mappings

9. Before writing anything, show a **simple confirmation table** to the user:

```
Here's what I'll update:

| Slide | What | New Value |
|-------|------|-----------|
| 3 | Title date | "tpid" |
| 3 | Quadrant a | 2 |
| 3 | Quadrant b | 3 |
| 3 | Bar chart | your_image.jpg |
```

10. Once confirmed, write mappings to `configs/<deck>_mappings.json`:

```python
import json
from datetime import datetime

mappings = {
    "deck": "<deck_name>",
    "created": datetime.now().isoformat(),
    "mappings": [
        # Text: {"slide": "slide_N", "shape_id": "ID", "type": "text", "data_field": "literal:value"}
        # Text with target_run: add "target_run": N
        # Image: {"slide": "slide_N", "shape_id": "ID", "type": "image", "source": "visual:filename_stem"}
        # Table: {"slide": "slide_N", "shape_id": "ID", "type": "table", "data_field": "source_name"}
    ]
}

with open(f"configs/{deck}_mappings.json", "w") as f:
    json.dump(mappings, f, indent=2)
```

**CRITICAL: Do NOT modify the config file. Only write mappings.json.**

The `type` field must be one of: `"text"`, `"image"`, `"table"`.
- Text and table mappings use `data_field` (data key, dot-path, or `literal:...`).
- Image mappings use `source` (`visual:name`, `screenshot:dir/key`, or bare filename path).
- Presence in mappings.json implies the shape is dynamic — no `is_dynamic` field needed.

11. After writing, say: **"Got it! Ready to update your slides?"**

## Important

- **NEVER edit the config file** — only write mappings.json
- **NEVER show shape IDs** to the user unless they specifically ask for technical details
- Describe shapes by their content: "the title", "the bar chart on the left", "the table at the bottom"
- ALWAYS pause and confirm the mapping table before writing
- If the user is unsure, skip it and come back later
- Keep the conversation natural — the user shouldn't need to know about shape IDs, XML, or config internals
