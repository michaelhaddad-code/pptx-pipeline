---
name: inject
description: "Step 5: Inject resolved values into slide XML — text, tables, images, font scaling. Modifies _raw/ in component library. Use after /update-config."
allowed-tools: Bash, Read, Glob, Grep
---

# Step 5: Inject

Apply all resolved values into the raw slide XML files.

## Instructions

1. Run inject with logging:

```python
import logging
logging.basicConfig(level=logging.INFO)
from inject import inject
inject("configs/slides_examples.json", "component_library", dry_run=False)
```

2. After it completes, summarize what was injected:
   - Per slide: number of text replacements, token replacements, table injections
   - Images swapped (which files replaced which)
   - Font adjustments applied (shrink or autofit)
   - Table row adjustments (rows added/removed)

3. If there were any errors or warnings, report them clearly.

4. Tell the user to run `/reconstruct` as the next step to produce the output PPTX.

## Dry Run Option

If the user asks for a preview first, run with `dry_run=True` to show what would change without modifying files. Then ask if they want to proceed with the real injection.
