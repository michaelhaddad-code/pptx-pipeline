---
name: update-config
description: "Step 4: Resolve mapped fields with actual data values and compute layouts (font sizes, image fit, table rows). Use after /map."
allowed-tools: Bash, Read, Glob, Grep
---

# Step 4: Update Config

Apply mappings from mappings.json, resolve all data fields, and compute layouts.

## Instructions

1. Find the config and mappings files:

```python
import glob
config_files = [f for f in glob.glob("configs/*.json") if "_mappings" not in f]
# Mappings file follows naming convention: configs/<deck>_mappings.json
```

2. Run update_config with logging. The script automatically reads mappings.json and applies them before resolving values:

```python
import sys, logging
sys.path.insert(0, "src")
logging.basicConfig(level=logging.INFO)
from update_config import update_config
result = update_config("configs/<deck>.json", "data")
```

The script will:
- Reset all mapping fields (idempotent)
- Apply mappings from `configs/<deck>_mappings.json`
- Load data sources from `data/`
- Resolve all mapped fields to actual values
- Compute layouts (font sizes, image fit, table rows)

3. After it completes, summarize the resolution results:
   - How many mappings were applied from mappings.json
   - How many fields were resolved successfully
   - How many fields failed to resolve (and which ones)
   - Layout computations performed:
     - Text shapes: computed font sizes
     - Tables: computed row heights and font sizes
     - Images: computed fit dimensions and offsets

4. If any fields failed to resolve, warn the user and suggest:
   - Check the data file for the missing field
   - Re-run `/map` to fix the mapping
   - Or manually add the data to the data/ directory

5. Tell the user to run `/inject` as the next step.
