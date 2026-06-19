# Architecture & Porting Guide

How `nlm2pptx.py` is structured, and how to embed it in a notebook, a job, or a Databricks web app.

## Design principles

- **No environment coupling.** Stdlib + `python-pptx` + `pillow` only; OpenAI is called via
  `urllib` (no SDK). All I/O is via function arguments — no hardcoded paths, no `open` command,
  no system-font lookups. Runs identically on Windows, macOS, Linux, and Databricks.
- **API key from `OPENAI_API_KEY`** (or passed explicitly as `api_key=`). Override base URL with
  `OPENAI_BASE_URL` (e.g. Azure/OpenAI-compatible gateways).
- **Models via env vars**: `NLM2PPTX_IMAGE_MODEL` (default `gpt-image-2`),
  `NLM2PPTX_OCR_MODEL` (default `gpt-5.5`), `NLM2PPTX_FONT` (default `맑은 고딕`).

## Stage contracts

| Function | Input | Output | Calls API? |
|---|---|---|---|
| `extract_slides(input_path, out_dir)` | pptx/pdf path | list of slide PNG paths | no |
| `erase_text(png, out_path)` | one slide PNG | text-erased PNG path | yes (image) |
| `ocr_slide(png)` | one slide PNG | `list[block]` (text, box_2d, font_size_pt, bold, align, color) | yes (vision) |
| `latex2plain(s)` | string | unicode-plain string | no |
| `build_pptx(bg_paths, blocks_per_slide, out)` | backgrounds + blocks | pptx path | no |
| `convert(input, output, erase=, font=, api_key=)` | full orchestration | pptx path | yes |

`block` shape (OCR output): `{"text": str, "box_2d": [ymin,xmin,ymax,xmax] (0-1000),
"font_size_pt": float, "bold": bool, "align": "left|center|right", "color": "RRGGBB"}`.

Coordinates are a 0–1000 grid over a 16:9 slide; `build_pptx` maps them to EMU.

## Porting to parallel / Databricks

The per-slide stages (`erase_text`, `ocr_slide`) are independent and I/O-bound — ideal for
parallelism. `convert()` runs them sequentially for simplicity; for scale, fan out yourself:

```python
from concurrent.futures import ThreadPoolExecutor
from nlm2pptx import extract_slides, erase_text, ocr_slide, build_pptx

pngs = extract_slides(inp, workdir)
with ThreadPoolExecutor(max_workers=8) as ex:
    cleans = list(ex.map(lambda p: erase_text(p, p+".clean.png"), pngs))
    blocks = list(ex.map(ocr_slide, pngs))
build_pptx(cleans, blocks, out)
```

**Databricks web app notes:**
- Read `OPENAI_API_KEY` from a Databricks secret scope into the env, or pass `api_key=` directly.
- Use a job-local temp dir (`tempfile.mkdtemp()`) for `workdir`; don't assume `/tmp` semantics.
- For a request/response web app, run `extract → (parallel erase+ocr) → build` inside the
  handler and stream the `[1/4]…[4/4]` progress, or push slide-level progress events.
- All paths are passed in, so DBFS/volume paths (`/Volumes/...`, `/dbfs/...`) work unchanged.
- No browser/`open` calls in the core module — safe for headless execution.

## Why these choices

- **gpt-image-2 for erase**: preserves original illustrations/table structure; `gpt-image-1-mini`
  is cheaper but visibly distorts diagrams (tested side-by-side).
- **gpt-5.5 for OCR**: gives per-block font size + bold that older vision models flatten,
  which is what makes the rebuilt text sizes look consistent instead of all-one-size.
- **Font name only, not embedded**: pptx stores the font *name*; the viewer's PowerPoint renders
  it. `맑은 고딕` is the Windows Korean default, so it looks right on the most common target.
