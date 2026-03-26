---
name: update
description: "Step 3: Update slides and deliver the output PPTX. Runs update_config + inject + reconstruct automatically. Use after /map."
allowed-tools: Bash, Read, Glob, Grep
---

# Step 3 — Update and Deliver

Automatically apply all mappings, inject values, and build the output file. No user input needed.

## Instructions

Run all three internal steps silently (update_config + inject + reconstruct) and present ONE combined result.

1. Show a brief status message: "Updating your slides..."

2. Find the config and mappings files:

```python
import glob
config_files = [f for f in glob.glob("configs/*.json") if "_mappings" not in f]
config_path = sorted(config_files)[-1]
# Derive deck name from config filename
import os
deck = os.path.splitext(os.path.basename(config_path))[0]
mappings_path = f"configs/{deck}_mappings.json"
```

3. Run update_config:

```python
import sys, logging
sys.path.insert(0, "src")
logging.basicConfig(level=logging.WARNING)
from update_config import update_config
update_config(config_path, "data", mappings_path)
```

4. Immediately run inject (no pause):

```python
from inject import inject
inject(config_path, "component_library")
```

5. Immediately run reconstruct (no pause):

```python
from reconstruct import reconstruct
output_path = f"output/{deck}_updated.pptx"
reconstruct("component_library", output_path)
```

6. Present a **plain English result** to the user:
   - Confirm what was changed (e.g., "Updated 5 text values and swapped 1 image on slide 3")
   - Give them the output file path
   - Tell them to open it and check

Example output style:
```
Done! Your updated PowerPoint is ready.

Changes made:
- Slide 3: Updated title text, 4 quadrant values, and replaced 1 chart image

Your file: output/test9_updated.pptx
```

## Important

- Do NOT pause between internal steps — run them all back to back
- Do NOT show internal step names (update_config, inject, reconstruct)
- Do NOT show shape IDs, XML details, config paths, or EMU values
- Do NOT tell the user to run another command
- Keep logging at WARNING level so internal details stay hidden
- If any step fails, show a friendly error message and suggest re-running from the start
- The output filename should be based on the deck name: output/<deck>_updated.pptx
