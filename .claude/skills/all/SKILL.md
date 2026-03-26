---
name: all
description: "Run the full 3-step pipeline: read the PowerPoint, map interactively, then update and deliver. Use when the user provides a PPTX and wants the full flow."
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
argument-hint: "<pptx-file>"
---

# Full Pipeline — 3 Steps

Run the complete pipeline: Read → Map → Update.

**Input file**: $ARGUMENTS (default: "Slides Examples.pptx")

## Instructions

### Step 1 — Read Your PowerPoint
Run `/read $ARGUMENTS` logic:
- Run deconstruct + generate_config back to back (no pause between them)
- Show a plain English summary of what's in the deck
- Then ask: "Which slides do you want to update, and what changes do you need?"

### Step 2 — Tell Me What Goes Where
Run `/map` logic:
- This is the interactive mapping conversation
- User describes what they want changed in plain English
- Agent figures out which shapes map to which data
- Show a simple confirmation table of all mappings before proceeding
- Ask: "Look good? I'll update your slides now."

### Step 3 — Update and Deliver
Run `/update` logic:
- Run update_config + inject + reconstruct back to back (no pause between them)
- Show a plain English summary of what was changed
- Deliver the output file path

## Important

- Only pause between the 3 user-facing steps, NOT between internal sub-steps
- NEVER mention internal step names (deconstruct, generate_config, update_config, inject, reconstruct)
- NEVER show shape IDs, XML, EMU, config paths, or other technical details
- Keep the conversation natural and non-technical
- If any step fails, show a friendly error and help the user fix it
