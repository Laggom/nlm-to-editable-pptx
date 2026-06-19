# Verifying the result

The conversion is visual, so the best check is a side-by-side comparison of each original
slide against the rebuilt slide. Two ways:

## 1. Open in PowerPoint (definitive)

Open the output `.pptx`. Click on any text — it should be a selectable, editable text box.
Confirm the background no longer shows the original (now-erased) text underneath.

## 2. Generate a comparison HTML (headless-friendly)

When PowerPoint isn't available (servers, Databricks, CI), render each rebuilt slide to a PNG
and place it next to the original. Use `scripts/compare_html.py`:

```bash
python scripts/compare_html.py --workdir <workdir-used-for-convert> --out comparison.html
```

It reads the extracted slide PNGs (`slides/`) and cleaned backgrounds (`clean/`) plus the OCR
blocks, draws the text onto the cleaned backgrounds (approximating PowerPoint's rendering with
PIL), and writes a single self-contained HTML with original-vs-rebuilt columns. Images are
base64-embedded so the file opens anywhere.

Note: the PIL preview is an approximation (it may not render every unicode subscript the same
way PowerPoint does). It's for layout/position sanity-checking, not pixel-perfect proofing —
PowerPoint itself is the source of truth.

## What to look for

- **Text position** matches the original (titles, labels, body land in the right place).
- **Font sizes** look consistent within a slide (titles bigger than body, no random giant label).
- **No ghost text** in the background (erase worked).
- **Illustrations/tables preserved** (gpt-image-2 should keep them faithful).
- **Korean/formulas correct** (spot-check a slide with math).

If one slide is off (e.g. a font size looks wrong, or erase left a smudge), re-running the
conversion for that single slide usually fixes it — the models are slightly nondeterministic.
