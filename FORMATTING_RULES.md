# Formatting & Layout Rules

How the pipeline handles different types of content when updating your PowerPoint.

---

## Text Rules

### Auto-shrink to fit
When new text is longer than what was originally in the shape, the font size shrinks to fit.
- The font never shrinks below 75% of the original size
- Absolute minimum font size: 6pt (anything smaller is unreadable)
- If the text fits on one line at the maximum font size, it stays at maximum — no shrinking

### Multi-part text (runs)
PowerPoint text often has multiple styled parts — for example, a bold label followed by regular body text. The pipeline preserves this:
- Each part keeps its own font, color, bold/italic settings
- You can target a specific part to replace (e.g., just the date in "Report as of Mar 05")
- Parts that aren't targeted stay exactly as they are

### Label : Value pairs
Text like `"ACR Growth: 10% target"` is treated as a label ("ACR Growth") and a value ("10% target"). When replacing:
- The label stays bold (or whatever style it had)
- The value gets the body style
- This only applies when there's a colon within the first 60 characters

### Multi-paragraph text
When text has multiple paragraphs (bullet points, line breaks), each paragraph is matched one-to-one:
- Paragraph 1 in the new text replaces Paragraph 1 in the original
- Formatting (bullets, indentation, spacing) is preserved per paragraph
- This only works when the new text has the same number of paragraphs as the original

### Font size estimation
The pipeline estimates how much text fits in a box using these assumptions:
- Each character is roughly 55% as wide as the font height
- Line spacing is 120% of font height
- PowerPoint text boxes have ~0.1 inch padding on each side
- Multi-paragraph text gets a 20% safety margin; single paragraph gets 10%

---

## Image Rules

### Replacement sizing
When you replace an image, the new image is sized to match the original placeholder:
- **Width stays the same** — always matches the original shape's width
- **Height adjusts proportionally** based on the new image's aspect ratio
- If the new image's aspect ratio is within 5% of the original, the height doesn't change (avoids tiny shifts)

### When images grow taller
If a replacement image is taller than the original (different aspect ratio):
- Everything below that image on the slide gets pushed down to make room
- Any shapes sitting on top of the image (like a transparent overlay) get resized proportionally
- Overlay detection: a shape counts as "on top of" an image if it's within ~1.6pt of the image's position and fits within its width

### Multiple images on the same slide
When a slide has several images stacked vertically:
- The pipeline detects which images are in the same column (based on horizontal overlap)
- Available vertical space is divided proportionally — taller images get more space
- All images are scaled by the same factor if they'd overflow the slide
- A small gap (~0.05 inch) is maintained between images

### Image DPI
The pipeline reads DPI metadata from PNG and JPEG files to calculate correct physical size. If no DPI metadata is found, it assumes 96 DPI (standard screen resolution).

---

## Table Rules

### Row heights
When table data has more or fewer rows than the template:
- Rows are added or removed to match the data
- Row height = total table height / number of rows
- Row height never exceeds the template's original row height
- Row height never goes below the template's minimum row height

### Table font scaling
When rows get shorter (more rows squeezed in), fonts scale down proportionally:
- Scale factor = new row height / original row height
- Fonts never scale UP, only down
- Fonts never shrink below 50% of their original size

### Default table font sizes
| Part | Default size | Minimum |
|------|-------------|---------|
| Header row | 10pt | 5pt |
| Total/summary row | 17pt | 8.5pt |
| Data rows | 12pt | 6pt |

### Cell content
- Column matching is by header name (must match exactly)
- Each cell's first text element gets the new value; extras are cleared
- Rows are processed bottom-to-top internally (doesn't affect output)

---

## Slide-Level Layout Rules

### Content stacking
When images change size, the pipeline reorganizes the slide:
1. Detect columns (left side vs right side) based on horizontal position
2. Within each column, sort content top-to-bottom
3. Allocate space: labels + gaps + static content first, remaining space to dynamic images
4. If total content exceeds slide height, scale everything down uniformly

### Gaps and spacing
| Between what | Gap size |
|---|---|
| Section label and its image | ~0.005 inch |
| Between image sections | ~0.05 inch |
| Bottom of slide margin | ~0.05 inch |
| Label reserved height | ~0.25 inch |

### Overflow handling
If content would extend past the bottom of the slide after all layout:
- All sections (dynamic and static) are scaled down by a uniform factor
- This keeps proportions consistent — no single image gets shrunk more than others
- A warning is shown if anything still overflows after scaling

---

## What Stays the Same

These things are never changed by the pipeline:
- Slide backgrounds, themes, and master layouts
- Shape positions (x, y) unless pushed down by an expanding image above
- Shape widths (cx) — only heights may change
- Colors, gradients, borders, shadows on shapes
- Animations, transitions, notes
- Non-dynamic shapes (anything you don't map stays untouched)
