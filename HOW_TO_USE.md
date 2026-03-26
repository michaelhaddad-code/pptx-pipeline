# How to Update Your PowerPoint in 4 Steps

## What You Need

1. **Your PowerPoint file** — the template you want to update
2. **A data folder** — a folder containing the new values you want to put in

Your data folder should look something like this:

```
data/
  report_data.csv          ← spreadsheets with your numbers
  quarterly_metrics.xlsx   ← Excel files work too
  notes.json               ← or JSON files
  visuals/                 ← a subfolder for replacement images
    chart_sales.png
    team_photo.jpg
```

**Data files** can be CSV, Excel (.xlsx), JSON, or plain text (.txt).
Put any **replacement images** inside a `visuals/` subfolder.

---

## How to Run It

Open a terminal in the project folder and run:

```
python update.py "My Presentation.pptx" data/
```

That's it. One command.

The updated file will be saved as **My Presentation_updated.pptx** in the same folder as your original.

To save it somewhere else:

```
python update.py "My Presentation.pptx" data/ --output "C:\Users\Me\Desktop\result.pptx"
```

---

## How the Mapping Step Works

After reading your PowerPoint and scanning your data, the tool walks you through each slide and asks what should go where.

You'll see something like this:

```
━━━ Slide 3 — found 5 text fields, 1 image ━━━

  Map this slide? (y/n, default: y): y

  [Slide 3 — Text]
  Currently: "MACC Quality & Performance | TPID"
  Text parts:
    Part 0: "MACC Quality & Performance | "
    Part 1: "TPID"
    Part 2: "as of Mar 05"

    1. report_title
    2. period_label
    3. account_name
    t. Type a value
    Enter. Skip (keep as-is)

  Your choice: t
  Type the value: Contoso Corp
  Which part to replace? (0-2, or Enter for all): 1
  ✅ Mapped → "Contoso Corp"
```

For each item on each slide, you can:
- **Pick a number** to use data from your files
- **Press `t`** to type a value directly
- **Press Enter** to skip it (leave it unchanged)

If a text has multiple parts (like a title with different colors), you'll be asked which part to replace.

---

## Reusing Your Mappings

The first time you run the tool, your choices are saved automatically. To reuse them next time:

```
python update.py "My Presentation.pptx" data/ --mappings configs/my_mappings.json
```

This skips the interactive step and applies your saved mappings directly — useful when you update the same deck regularly with new data.

---

## When Something Is Skipped

If you see a warning like:

> ⚠️  Could not find "sales_total" in your data files — that field will be left as-is.

Here's what to check:

- [ ] **Is the column name spelled correctly?** Open your CSV/Excel file and check the header row matches exactly.
- [ ] **Is the file in the right folder?** Data files go in the `data/` folder, images go in `data/visuals/`.
- [ ] **Is the file format supported?** Use CSV, Excel (.xlsx), JSON, or plain text.
- [ ] **Did the PowerPoint change?** If someone edited the template since your last run, some placeholders may have moved or been deleted. Try running again with the latest version of the file.

---

## Quick Reference

| What you want to do | Command |
|---|---|
| Update a PowerPoint | `python update.py "file.pptx" data/` |
| Save to a specific location | `python update.py "file.pptx" data/ --output "result.pptx"` |
| Reuse previous mappings | `python update.py "file.pptx" data/ --mappings configs/saved.json` |
| See detailed output | `python update.py "file.pptx" data/ --verbose` |
