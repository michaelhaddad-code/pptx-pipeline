---
name: generate-config
description: "Step 2: Generate config.json from the component library. Classifies shapes as dynamic/static, creates layout stubs. Use after /deconstruct."
allowed-tools: Bash, Read, Glob, Grep
---

# Step 2: Generate Config

Generate the configuration file from the component library.

## Instructions

1. Run generate_config with logging:

```python
import logging
logging.basicConfig(level=logging.INFO)
from generate_config import generate_config
generate_config("component_library", "configs", force=True)
```

2. After it completes, load and analyze the config:

```python
import json
with open("configs/slides_examples.json") as f:
    cfg = json.load(f)
```

3. Present a summary to the user showing per slide:
   - Dynamic shapes: name, category (text/image/table), text preview (first 60 chars)
   - Static shapes: count only
   - Images: rid, target file
   - Layout stubs created (auto_fit, dynamic_table, dynamic_image)

4. Highlight which shapes were detected as dynamic and why (placeholder patterns found).

5. Tell the user to run `/map` as the next step to map data fields to these shapes.
