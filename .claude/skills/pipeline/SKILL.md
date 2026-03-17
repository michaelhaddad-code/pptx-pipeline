---
name: pipeline
description: "Run the full PPTX pipeline step by step with pauses between each step. Deconstruct → Generate Config → Map → Update Config → Inject → Reconstruct."
allowed-tools: Bash, Read, Write, Edit, Glob, Grep
argument-hint: "<pptx-file>"
---

# Full Pipeline

Run the complete PPTX pipeline step by step.

**Input file**: $ARGUMENTS (default: "Slides Examples.pptx")

## Instructions

Execute each step in order, **pausing after each one** for user review before proceeding.

### Step 1: Deconstruct
Run `/deconstruct $ARGUMENTS` logic. Show summary. Ask user: "Ready for Step 2 (Generate Config)?"

### Step 2: Generate Config
Run `/generate-config` logic. Show summary of dynamic shapes found. Ask user: "Ready for Step 3 (Map)?"

### Step 3: Map (Interactive)
Run `/map` logic. This is the conversational step:
- Show available data and shapes needing mapping
- Work with user slide by slide to establish mappings
- Write agreed mappings to config
- Ask user: "Mappings complete. Ready for Step 4 (Update Config)?"

### Step 4: Update Config
Run `/update-config` logic. Show resolution results. Ask user: "Ready for Step 5 (Inject)?"

### Step 5: Inject
Run `/inject` logic. Show injection summary. Ask user: "Ready for Step 6 (Reconstruct)?"

### Step 6: Reconstruct
Run `/reconstruct output.pptx` logic. Show final output details.

## Important

- ALWAYS pause between steps and wait for user confirmation
- If any step fails, stop and help the user debug before continuing
- The user can skip steps by saying "skip" or go back by saying "re-run step N"
- Step 3 (Map) is interactive and may take multiple conversation turns
