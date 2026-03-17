---
name: update-config
description: "Step 4: Resolve mapped fields with actual data values and compute layouts (font sizes, image fit, table rows). Use after /map."
allowed-tools: Bash, Read, Glob, Grep
---

# Step 4: Update Config

Resolve all mapped data fields with actual values and compute layouts.

## Instructions

1. Run update_config with logging:

```python
import logging
logging.basicConfig(level=logging.INFO)
from update_config import update_config
result = update_config("configs/slides_examples.json", "data")
```

2. After it completes, summarize the resolution results:
   - How many fields were resolved successfully
   - How many fields failed to resolve (and which ones)
   - Layout computations performed:
     - Text shapes: computed font sizes
     - Tables: computed row heights and font sizes
     - Images: computed fit dimensions and offsets

3. If any fields failed to resolve, warn the user and suggest:
   - Check the data file for the missing field
   - Re-run `/map` to fix the mapping
   - Or manually add the data to the data/ directory

4. Tell the user to run `/inject` as the next step.
