---
name: reconstruct
description: "Step 6: Repack the modified component library into an output PPTX file. Final step of the pipeline. Use after /inject."
allowed-tools: Bash, Read, Glob
argument-hint: "<output-file.pptx>"
---

# Step 6: Reconstruct

Repack the modified component library into an output PPTX.

**Output file**: $ARGUMENTS (default: "output.pptx")

## Instructions

1. Run reconstruct with logging:

```python
import logging
logging.basicConfig(level=logging.INFO)
from reconstruct import reconstruct
output_path = reconstruct("component_library", "$ARGUMENTS")
```

2. After it completes, report:
   - Files packed count
   - Output file size
   - Output file path

3. Optionally run verify against the original:

```python
from reconstruct import verify
verify("Slides Examples.pptx", "$ARGUMENTS", "component_library")
```

4. Report verification results:
   - File list comparison (missing/extra)
   - Modified slides (expected — these are the ones we injected into)
   - Media file changes (expected — these are the images we swapped)

5. Tell the user the output is ready to open and review.
