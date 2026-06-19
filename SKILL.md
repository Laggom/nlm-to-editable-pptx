---
name: nlm-to-editable-pptx
description: >-
  Use when the user has an image-based / non-editable PPTX or PDF (especially exported
  from NotebookLM, or any "scanned" deck where the slides are full-page pictures with no
  selectable text) and wants it converted into an editable PowerPoint with real, editable
  text boxes. Trigger this whenever the user mentions NotebookLM slides, "make this PPT
  editable", "the text isn't selectable", "convert this PDF deck to editable slides",
  Korean slide OCR, or rebuilding a deck so the text can be edited — even if they don't
  say the word "skill". Works on Windows/macOS/Linux and is portable to Databricks.
---

# NotebookLM → Editable PPTX

## What this does

NotebookLM (and many "export to slides" tools) produce decks where **each slide is a single
full-page image** — the text looks real but cannot be selected or edited. This skill rebuilds
such a deck into a **genuinely editable `.pptx`**:

1. The original text is **erased from the background** by an image model (so no ghost text remains).
2. The text is **re-extracted with precise position, size, color** by a vision model.
3. A new `.pptx` is assembled: cleaned background image + native, editable text boxes on top.

The result keeps the original look (illustrations, diagrams, colors) but every piece of text
is now a real PowerPoint text box you can click and edit. Korean and math/formulas are handled.

## When to use vs. not

- **Use** when text must become editable AND the original visuals should be preserved.
- **Use** for NotebookLM PPTX/PDF, scanned decks, screenshot decks, image-only slides.
- **Don't** use to author a brand-new presentation from scratch (no source deck).
- **Don't** use if the deck already has editable text (check by selecting text in PowerPoint).

## Requirements

- **`OPENAI_API_KEY`** environment variable (the skill reads it automatically).
  - Windows: `setx OPENAI_API_KEY "sk-..."` (new shell) or `$env:OPENAI_API_KEY="sk-..."` (current PowerShell).
  - macOS/Linux: `export OPENAI_API_KEY=sk-...`
- **Python 3.11+** with `python-pptx` and `pillow`. For PDF input also `pymupdf`.
  - `pip install python-pptx pillow pymupdf`
- Models used (overridable via env vars): image edit **`gpt-image-2`**, OCR **`gpt-5.5`**.

## How to run

The whole pipeline is one self-contained module: `scripts/nlm2pptx.py`. Prefer running it
directly — it needs no other files from this skill.

```bash
python scripts/nlm2pptx.py <input.pptx|input.pdf> <output.pptx>
```

Common options:
- `--no-erase` — skip text removal; keep the original background and just overlay editable
  text. Much faster and cheaper, but the original (non-editable) text stays visible underneath.
  Use when the user only needs selectable/editable text and doesn't mind the original showing.
- `--font "맑은 고딕"` — font name written into the pptx (default `맑은 고딕`, Windows Korean
  default). The actual rendering uses whatever font the viewer's PowerPoint has.
- `--workdir DIR` — keep intermediate slide PNGs, cleaned backgrounds, and `convert.log` for
  inspection. **Always pass this when debugging** — the per-slide timing and any HTTP/timeout
  errors are written there.
- `--workers N` — slides processed in parallel (default 6, `1` = sequential). Parallelism is
  the difference between ~2.5 min and ~10+ min for a 12-slide deck, since each slide's
  image/OCR call is I/O-bound. Lower it only if you hit rate limits.

Each slide calls the OCR model (~30–90s) and, unless `--no-erase`, the image model (~30–60s).
With the default `--workers 6`: ~2.5 min for `--no-erase`, ~7 min full. Tell the user this up
front. The pipeline logs per-slide progress and timing (`[1/4]…[4/4]`, then `slideN: 완료`);
a single slide that times out is retried automatically and, if it still fails, is filled with
an empty-text fallback so the rest of the deck still completes.

## Workflow when invoked

1. Confirm the input file path and that it's image-based (if unsure, note that editable decks
   don't need this skill).
2. Check `OPENAI_API_KEY` is set; if not, ask the user to set it (see Requirements).
3. Run `scripts/nlm2pptx.py input output`. Stream progress to the user (it prints `[1/4]…[4/4]`).
4. When done, report the output path. Offer to generate a side-by-side comparison HTML
   (original vs. result) for verification — see `references/verify.md`.

## Programmatic use (notebooks / web app / Databricks)

`nlm2pptx.py` is a plain module with no global state or hardcoded paths — import and call it.
This is the path for the eventual Databricks web app.

```python
from nlm2pptx import convert
convert("input.pptx", "output.pptx")           # full pipeline (erase + ocr + build)
convert("input.pdf", "out.pptx", erase=False)  # fast: overlay text only
```

Individual stages (`extract_slides`, `erase_text`, `ocr_slide`, `build_pptx`) are exported too,
so a web app can parallelize per-slide erase/OCR across workers and stream progress. See
`references/architecture.md` for the stage contracts and how to port to a Spark/job context.

## Notes & limits

- Formulas are converted to readable unicode (e.g. `m₀/√(1−v²/c²)`), not rendered math — they
  stay editable as text. This matches how lightweight slide tools handle math.
- Font sizes come from the vision model's per-block estimate plus a width-fit safety clamp, so
  text stays inside its box. If a specific slide looks off, re-running that slide usually helps.
- The cleaned background is a regenerated image; `gpt-image-2` preserves the original layout far
  better than `gpt-image-1-mini` (which distorts illustrations) — keep the default.
