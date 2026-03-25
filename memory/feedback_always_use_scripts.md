---
name: Always use pipeline scripts
description: Never manually edit config JSON — always use the pipeline scripts (update_config.py, inject.py, etc.) to compute values
type: feedback
---

Never manually set config fields (like `images` array, `_computed`, `resolved_value`) by hand. Always use the pipeline scripts to do it.

**Why:** Manually writing config fields skips layout computation (image fit, font sizing, geometry). This caused distorted oversized images and text overflow because `_computed` was missing.

**How to apply:** During mapping, only set `is_dynamic`, `data_field`, and `source` fields. Let `update_config.py` handle resolution and layout computation. Let `inject.py` handle XML modification.
