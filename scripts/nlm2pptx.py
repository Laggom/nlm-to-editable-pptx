"""
nlm2pptx — NotebookLM(이미지 기반) PPTX/PDF를 '편집 가능한' PPTX로 변환하는 핵심 로직.

설계 원칙 (이식성):
- 표준 라이브러리 + python-pptx + pillow 만 사용. OpenAI 호출은 urllib(표준)로 직접.
- 경로/환경 의존 없음: 모든 입출력은 함수 인자. OS 무관(Windows/Linux/macOS/Databricks).
- API 키는 환경변수 OPENAI_API_KEY 에서 읽음 (인자로 직접 주입도 가능).
- macOS 전용 도구(editppt, open, 시스템 폰트 경로) 일절 사용 안 함.

파이프라인:
  1) extract_slides(): pptx/pdf → 페이지별 PNG
  2) erase_text():     각 PNG → 글자 지운 배경 PNG  (gpt-image-2, images/edits)
  3) ocr_slide():      각 PNG → 텍스트 블록 JSON     (gpt-5.5, vision + JSON)
  4) build_pptx():     배경 + OCR블록 → 편집 가능 PPTX (python-pptx)

CLI:
  python nlm2pptx.py <input.pptx|input.pdf> <output.pptx> [--workdir DIR] [--no-erase] [--font "맑은 고딕"]
"""
from __future__ import annotations
import base64, json, os, re, sys, time, uuid, zipfile, io, glob, argparse, logging, threading
import urllib.request, urllib.error
from concurrent.futures import ThreadPoolExecutor, as_completed

# ──────────────────────────────────────────────────────────────────────────
# 로깅: 타임스탬프 + 단계/슬라이드/소요시간. 콘솔 + (workdir 지정 시) 파일.
# 환경변수 NLM2PPTX_LOG_LEVEL 로 조절(기본 INFO, 디버그는 DEBUG).
# ──────────────────────────────────────────────────────────────────────────
log = logging.getLogger("nlm2pptx")
if not log.handlers:
    _h = logging.StreamHandler(sys.stdout)
    _h.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
    log.addHandler(_h)
log.setLevel(getattr(logging, os.environ.get("NLM2PPTX_LOG_LEVEL", "INFO").upper(), logging.INFO))

def _add_file_log(workdir: str):
    """workdir/convert.log 에 전체 로그를 함께 기록."""
    fh = logging.FileHandler(os.path.join(workdir, "convert.log"), encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    fh.setLevel(logging.DEBUG)
    log.addHandler(fh)
    return fh

# ──────────────────────────────────────────────────────────────────────────
# 설정 (기본 모델)
# ──────────────────────────────────────────────────────────────────────────
IMAGE_MODEL = os.environ.get("NLM2PPTX_IMAGE_MODEL", "gpt-image-2")   # 글자 제거
OCR_MODEL   = os.environ.get("NLM2PPTX_OCR_MODEL", "gpt-5.5")          # OCR
DEFAULT_FONT = os.environ.get("NLM2PPTX_FONT", "맑은 고딕")            # 한글 폰트(이름만 기입)
OPENAI_BASE = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
IMAGE_SIZE  = "1536x1024"   # 16:9 가로 (gpt-image 허용 사이즈)


def _api_key(explicit: str | None = None) -> str:
    key = explicit or os.environ.get("OPENAI_API_KEY")
    if not key:
        raise RuntimeError(
            "OpenAI API 키가 없습니다. 환경변수 OPENAI_API_KEY 를 설정하거나 api_key 인자로 전달하세요."
        )
    return key


# ──────────────────────────────────────────────────────────────────────────
# 1) 슬라이드 이미지 추출
# ──────────────────────────────────────────────────────────────────────────
def extract_slides(input_path: str, out_dir: str) -> list[str]:
    """pptx 또는 pdf에서 페이지별 PNG를 out_dir에 추출. 정렬된 경로 리스트 반환."""
    os.makedirs(out_dir, exist_ok=True)
    ext = os.path.splitext(input_path)[1].lower()
    paths: list[str] = []

    if ext in (".pptx", ".ppt"):
        # pptx는 zip. ppt/media/imageN.* 가 슬라이드 풀페이지 이미지.
        with zipfile.ZipFile(input_path) as z:
            media = sorted(
                [n for n in z.namelist() if n.startswith("ppt/media/image")],
                key=lambda n: int(re.search(r"image(\d+)", n).group(1)),
            )
            for i, name in enumerate(media, 1):
                data = z.read(name)
                # PNG로 정규화 (pillow)
                from PIL import Image
                im = Image.open(io.BytesIO(data)).convert("RGB")
                p = os.path.join(out_dir, f"slide{i}.png")
                im.save(p)
                paths.append(p)

    elif ext == ".pdf":
        # PDF는 PyMuPDF(fitz)로 페이지 렌더. (없으면 안내 후 예외)
        try:
            import fitz  # PyMuPDF
        except ImportError:
            raise RuntimeError("PDF 입력은 PyMuPDF가 필요합니다: pip install pymupdf")
        doc = fitz.open(input_path)
        for i, page in enumerate(doc, 1):
            pix = page.get_pixmap(dpi=200)
            p = os.path.join(out_dir, f"slide{i}.png")
            pix.save(p)
            paths.append(p)
    else:
        raise ValueError(f"지원하지 않는 입력 형식: {ext} (pptx 또는 pdf)")

    if not paths:
        raise RuntimeError(f"{input_path} 에서 슬라이드 이미지를 찾지 못했습니다.")
    return paths


# ──────────────────────────────────────────────────────────────────────────
# 2) 글자 제거 (gpt-image-2, images/edits 멀티파트)
# ──────────────────────────────────────────────────────────────────────────
ERASE_PROMPT = (
    "Remove all text and letters from this image while preserving the background, "
    "illustrations, diagrams, graphs, and all non-text visual elements exactly. "
    "Do not add any new text. Keep the same composition, colors, and lighting."
)

def erase_text(png_path: str, out_path: str, api_key: str | None = None,
               model: str | None = None, retries: int = 3) -> str:
    """슬라이드 PNG에서 글자만 제거한 배경 이미지를 생성해 out_path에 저장."""
    key = _api_key(api_key)
    model = model or IMAGE_MODEL
    img = open(png_path, "rb").read()
    boundary = "----" + uuid.uuid4().hex
    def part_field(name, val):
        return (f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n{val}\r\n').encode()
    def part_file(name, fn, data):
        return (f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"; '
                f'filename="{fn}"\r\nContent-Type: image/png\r\n\r\n').encode() + data + b"\r\n"
    body = (part_field("model", model) + part_field("prompt", ERASE_PROMPT)
            + part_field("size", IMAGE_SIZE)
            + part_file("image", os.path.basename(png_path), img)
            + (f"--{boundary}--\r\n").encode())
    name = os.path.basename(png_path)
    last = None
    for a in range(retries):
        t0 = time.time()
        try:
            log.debug(f"erase {name}: 시도 {a+1}/{retries} ({model})")
            req = urllib.request.Request(
                f"{OPENAI_BASE}/images/edits", data=body,
                headers={"Authorization": f"Bearer {key}",
                         "Content-Type": f"multipart/form-data; boundary={boundary}"})
            d = json.loads(urllib.request.urlopen(req, timeout=200).read())
            b64 = d["data"][0]["b64_json"]
            with open(out_path, "wb") as f:
                f.write(base64.b64decode(b64))
            log.info(f"erase {name}: OK ({time.time()-t0:.1f}s)")
            return out_path
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}: {e.read()[:200].decode('utf-8','ignore')}"
            log.warning(f"erase {name}: 시도 {a+1} 실패 ({time.time()-t0:.1f}s) — {last}")
            time.sleep(5)
        except Exception as e:
            last = f"{type(e).__name__}: {e}"
            log.warning(f"erase {name}: 시도 {a+1} 실패 ({time.time()-t0:.1f}s) — {last}")
            time.sleep(5)
    raise RuntimeError(f"글자 제거 실패({name}) {retries}회: {last}")


# ──────────────────────────────────────────────────────────────────────────
# 3) OCR (gpt-5.5, vision + 구조화 JSON)
# ──────────────────────────────────────────────────────────────────────────
OCR_PROMPT = """You are an expert slide OCR system. Analyze this 16:9 slide and return ONLY JSON: {"blocks":[...]}.
Each block = one visually-coherent text line or label. Fields:
- text: exact text. For math, convert to readable plain unicode (use ², ₂, √, Σ, ±, ÷, →, Greek letters; NOT LaTeX backslashes).
- box_2d: [ymin,xmin,ymax,xmax] in 0-1000 coords, TIGHT around the glyphs.
- font_size_pt: REAL point size for a 7.5-inch-tall slide (measure glyph height precisely; titles ~36-44, body ~14-20, captions ~10-13).
- bold: true/false
- align: left|center|right
- color: hex like FFFFFF
Group multi-line paragraphs into one block with \\n. Be precise about font_size_pt — it must match visual glyph height."""

def ocr_slide(png_path: str, api_key: str | None = None,
              model: str | None = None, retries: int = 3) -> list[dict]:
    """슬라이드 PNG에서 텍스트 블록 리스트(JSON) 추출."""
    key = _api_key(api_key)
    model = model or OCR_MODEL
    img_b64 = base64.b64encode(open(png_path, "rb").read()).decode()
    body = {
        "model": model,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": OCR_PROMPT},
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_b64}"}},
        ]}],
        "response_format": {"type": "json_object"},
    }
    name = os.path.basename(png_path)
    last = None
    for a in range(retries):
        t0 = time.time()
        try:
            log.debug(f"ocr {name}: 시도 {a+1}/{retries} ({model})")
            req = urllib.request.Request(
                f"{OPENAI_BASE}/chat/completions",
                data=json.dumps(body).encode(),
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"})
            d = json.loads(urllib.request.urlopen(req, timeout=200).read())
            content = d["choices"][0]["message"]["content"]
            obj = json.loads(content)
            blocks = obj.get("blocks", obj if isinstance(obj, list) else [])
            log.info(f"ocr {name}: OK {len(blocks)}블록 ({time.time()-t0:.1f}s)")
            return blocks
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}: {e.read()[:200].decode('utf-8','ignore')}"
            log.warning(f"ocr {name}: 시도 {a+1} 실패 ({time.time()-t0:.1f}s) — {last}")
            time.sleep(5)
        except Exception as e:
            last = f"{type(e).__name__}: {e}"
            log.warning(f"ocr {name}: 시도 {a+1} 실패 ({time.time()-t0:.1f}s) — {last}")
            time.sleep(5)
    raise RuntimeError(f"OCR 실패({name}) {retries}회: {last}")


# ──────────────────────────────────────────────────────────────────────────
# LaTeX → 유니코드 평문 (gpt-5.5가 대부분 평문으로 주지만, 혹시 모를 폴백)
# ──────────────────────────────────────────────────────────────────────────
_GREEK = {r"\\Psi":"Ψ",r"\\psi":"ψ",r"\\alpha":"α",r"\\beta":"β",r"\\gamma":"γ",r"\\delta":"δ",
          r"\\Delta":"Δ",r"\\theta":"θ",r"\\omega":"ω",r"\\Omega":"Ω",r"\\lambda":"λ",r"\\mu":"μ",
          r"\\nu":"ν",r"\\pi":"π",r"\\sigma":"σ",r"\\phi":"φ",r"\\rho":"ρ",r"\\tau":"τ",r"\\epsilon":"ε"}
_SYM = {r"\\sum":"Σ",r"\\int":"∫",r"\\partial":"∂",r"\\cdot":"·",r"\\times":"×",r"\\pm":"±",
        r"\\mp":"∓",r"\\infty":"∞",r"\\approx":"≈",r"\\neq":"≠",r"\\leq":"≤",r"\\geq":"≥",
        r"\\rightarrow":"→",r"\\to":"→",r"\\hbar":"ℏ",r"\\,":" ",r"\\!":"",r"\\quad":"  ",r"\\ ":" "}
_SUP = str.maketrans("0123456789+-=()n", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿ")
_SUB = str.maketrans("0123456789+-=()", "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎")

def latex2plain(t: str) -> str:
    if not t:
        return t
    s = t
    for _ in range(8):
        before = s
        s = re.sub(r"\\sqrt\s*\{([^{}]*)\}", r"√(\1)", s)
        s = re.sub(r"\\frac\s*\{([^{}]*)\}\s*\{([^{}]*)\}", r"(\1)/(\2)", s)
        if s == before:
            break
    s = re.sub(r"\\sqrt\s*\{([^{}]*)\}", r"√(\1)", s)
    s = re.sub(r"\\vec\s*\{([^{}]*)\}", r"\1⃗", s)
    for k, v in {**_GREEK, **_SYM}.items():
        s = re.sub(k + r"(?![A-Za-z])", v, s)
    sup = lambda m: m.group(1).translate(_SUP) if all(c in "0123456789+-=()n" for c in m.group(1)) else "^"+m.group(1)
    sub = lambda m: m.group(1).translate(_SUB) if all(c in "0123456789+-=()" for c in m.group(1)) else "_"+m.group(1)
    s = re.sub(r"\^\{([^{}]*)\}", sup, s)
    s = re.sub(r"\^(-?[0-9]+)", sup, s)
    s = re.sub(r"\^([A-Za-z0-9])", sup, s)
    s = re.sub(r"_\{([^{}]*)\}", sub, s)
    s = s.replace("{", "").replace("}", "")
    s = re.sub(r"\\([A-Za-z]+)", r"\1", s)
    return re.sub(r"\s+", " ", s).strip()


# ──────────────────────────────────────────────────────────────────────────
# 4) PPTX 조립
# ──────────────────────────────────────────────────────────────────────────
def build_pptx(bg_paths: list[str], ocr_blocks_per_slide: list[list[dict]],
               out_path: str, font: str | None = None) -> str:
    """배경 이미지 + 슬라이드별 OCR 블록 → 편집 가능 PPTX 저장."""
    from pptx import Presentation
    from pptx.util import Emu, Pt, Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.dml.color import RGBColor

    font = font or DEFAULT_FONT
    EMU_W, EMU_H = Inches(13.333), Inches(7.5)
    ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    prs = Presentation()
    prs.slide_width, prs.slide_height = EMU_W, EMU_H
    blank = prs.slide_layouts[6]

    for bg, blocks in zip(bg_paths, ocr_blocks_per_slide):
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(bg, 0, 0, width=EMU_W, height=EMU_H)
        for b in blocks:
            txt = latex2plain((b.get("text") or "").strip())
            box = b.get("box_2d")
            if not txt or not box or len(box) != 4:
                continue
            ymin, xmin, ymax, xmax = box
            # 텍스트 줄 높이 기준으로 폰트 산정: 한 줄 높이 ≈ font_size.
            # OCR이 멀티라인 블록의 box를 전체 높이로 주므로 줄 수로 나눠 한 줄 높이 추정.
            nlines = txt.count("\n") + 1
            line_h_pt = (ymax - ymin) / 1000 * EMU_H / 12700.0 / max(nlines, 1)
            fp = float(b.get("font_size_pt") or 16)
            # 모델 폰트와 박스에서 역산한 줄 높이 중 작은 쪽 근처로 — 박스를 넘지 않게.
            fp = min(fp, line_h_pt * 1.05)
            fp = max(8.0, min(fp, 48.0))
            # 박스: 좌상단은 OCR 좌표 그대로. 줄바꿈을 끄므로 너비는 텍스트가 담길 만큼.
            x_emu = int(xmin/1000 * EMU_W); y_emu = int(ymin/1000 * EMU_H)
            avail_pt = (EMU_W - x_emu) / 12700.0   # 시작점부터 슬라이드 우측 끝까지(pt)
            # 한 줄(가장 긴 줄)의 추정 폭: 한글≈0.95·fp, 그 외≈0.52·fp 폭 가정.
            def _line_w_pt(line, size):
                wide = sum(1 for c in line if ord(c) > 0x1100)   # CJK 등 전각
                return (wide * 0.95 + (len(line) - wide) * 0.52) * size
            longest_line = max(txt.split("\n"), key=len) if txt else ""
            # 폭이 가용 폭을 넘으면 폰트를 비례 축소(긴 한 줄 캡션이 오른쪽으로 잘리는 것 방지).
            est = _line_w_pt(longest_line, fp)
            if est > avail_pt and est > 0:
                fp = max(8.0, fp * avail_pt / est)
            need_pt = _line_w_pt(longest_line, fp) * 1.08 + 4   # 약간 여유
            w = Emu(int(min(max(need_pt * 12700.0, 0.04 * EMU_W), EMU_W - x_emu)))
            h = Emu(int(max((ymax-ymin)/1000, 0.02) * EMU_H))
            tb = slide.shapes.add_textbox(Emu(x_emu), Emu(y_emu), w, h)
            tf = tb.text_frame
            tf.word_wrap = False  # 줄바꿈 금지 — OCR이 분리한 줄 단위를 그대로 유지
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]; p.alignment = ALIGN.get(b.get("align", "left"), PP_ALIGN.LEFT)
            run = p.add_run(); run.text = txt
            run.font.size = Pt(round(fp, 1)); run.font.name = font
            if b.get("bold"):
                run.font.bold = True
            col = (b.get("color") or "FFFFFF").replace("#", "")
            try:
                run.font.color.rgb = RGBColor.from_string(col if len(col) == 6 else "FFFFFF")
            except Exception:
                run.font.color.rgb = RGBColor.from_string("FFFFFF")
    prs.save(out_path)
    return out_path


# ──────────────────────────────────────────────────────────────────────────
# 오케스트레이션 (전체 파이프라인)
# ──────────────────────────────────────────────────────────────────────────
def convert(input_path: str, output_path: str, workdir: str | None = None,
            erase: bool = True, font: str | None = None,
            api_key: str | None = None, max_workers: int = 6) -> str:
    """전체 변환. erase=False면 원본 배경 위에 텍스트만 올림(빠름, 글자 잔존).

    슬라이드별 글자제거/OCR은 max_workers 개 스레드로 병렬 실행(I/O 바운드라 효과 큼).
    max_workers=1 이면 순차. workdir 지정 시 convert.log 에 전체 로그 기록.
    """
    import tempfile
    t_start = time.time()
    workdir = workdir or tempfile.mkdtemp(prefix="nlm2pptx_")
    slides_dir = os.path.join(workdir, "slides")
    clean_dir = os.path.join(workdir, "clean")
    os.makedirs(clean_dir, exist_ok=True)
    fh = _add_file_log(workdir)
    log.info(f"=== convert 시작: {input_path} → {output_path} "
             f"(erase={erase}, workers={max_workers}, workdir={workdir}) ===")
    try:
        # 1) 추출
        t0 = time.time()
        slide_pngs = extract_slides(input_path, slides_dir)
        n = len(slide_pngs)
        log.info(f"[1/4] 슬라이드 추출: {n}장 ({time.time()-t0:.1f}s)")

        # 2+3) 슬라이드별 (글자제거 + OCR) 병렬
        bg_paths = [None] * n
        ocr_results = [None] * n
        errors = []

        def process(idx_png):
            i, png = idx_png
            tag = f"slide{i+1}"
            try:
                if erase:
                    clean = os.path.join(clean_dir, f"clean{i+1}.png")
                    erase_text(png, clean, api_key=api_key)
                    bg_paths[i] = clean
                else:
                    bg_paths[i] = png
                ocr_results[i] = ocr_slide(png, api_key=api_key)
                log.info(f"{tag}: 완료 (bg={'clean' if erase else 'orig'}, "
                         f"{len(ocr_results[i])}블록)")
            except Exception as e:
                log.error(f"{tag}: 처리 실패 — {type(e).__name__}: {e}")
                errors.append((i+1, str(e)))
                # 실패해도 진행: 배경은 원본, 텍스트는 빈 것으로 폴백
                bg_paths[i] = bg_paths[i] or png
                ocr_results[i] = ocr_results[i] or []

        stage = "글자제거+OCR" if erase else "OCR"
        log.info(f"[2/4][3/4] {stage} 병렬 시작 ({n}장, workers={max_workers})")
        t0 = time.time()
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            list(ex.map(process, list(enumerate(slide_pngs))))
        log.info(f"[2/4][3/4] {stage} 완료 ({time.time()-t0:.1f}s, 실패 {len(errors)}장)")

        # 4) 조립
        t0 = time.time()
        build_pptx(bg_paths, ocr_results, output_path, font=font)
        log.info(f"[4/4] PPTX 조립 → {output_path} ({time.time()-t0:.1f}s)")

        if errors:
            log.warning(f"일부 슬라이드 실패(빈 텍스트로 대체): {[e[0] for e in errors]}")
        log.info(f"=== 완료: {output_path} (총 {time.time()-t_start:.1f}s) ===")
        return output_path
    finally:
        log.removeHandler(fh)
        fh.close()


def _cli():
    ap = argparse.ArgumentParser(description="NotebookLM 이미지 PPTX/PDF → 편집 가능 PPTX")
    ap.add_argument("input", help="입력 .pptx 또는 .pdf")
    ap.add_argument("output", help="출력 .pptx")
    ap.add_argument("--workdir", default=None, help="중간 파일 디렉토리(기본: 임시)")
    ap.add_argument("--no-erase", action="store_true", help="글자 제거 생략(원본 배경 유지, 빠름)")
    ap.add_argument("--font", default=None, help=f"한글 폰트명(기본: {DEFAULT_FONT})")
    ap.add_argument("--workers", type=int, default=6, help="슬라이드 병렬 처리 수(기본 6, 1=순차)")
    args = ap.parse_args()
    convert(args.input, args.output, workdir=args.workdir,
            erase=not args.no_erase, font=args.font, max_workers=args.workers)


if __name__ == "__main__":
    _cli()
