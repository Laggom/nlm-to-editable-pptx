"""
원본 슬라이드 vs 재구성 결과를 비교하는 단일 HTML 생성 (검증용, 헤드리스 친화).

convert()를 --workdir 와 함께 돌렸다면 그 디렉토리에 slides/ 와 clean/ 이 남아있다.
이 스크립트는 거기에 OCR을 한 번 더 돌려(또는 캐시가 있으면 사용) 텍스트를 PIL로
clean 배경 위에 근사 렌더하고, 원본과 나란히 비교하는 HTML을 만든다.

사용:
  python compare_html.py --workdir <workdir> --out comparison.html
폰트 자동탐색(맑은 고딕/Noto/AppleSDGothicNeo 등)이 실패하면 PIL 기본 폰트로 폴백한다.
"""
import argparse, base64, glob, json, os, re, sys
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from nlm2pptx import ocr_slide, latex2plain

# 폰트 자동 탐색 (OS 무관)
FONT_CANDIDATES = [
    "C:/Windows/Fonts/malgun.ttf",                       # Windows 맑은 고딕
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",        # macOS
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",  # Linux Noto
    "/usr/share/fonts/opentype/noto/NotoSansCJKkr-Regular.otf",
]
def _font(px):
    for fp in FONT_CANDIDATES:
        if os.path.exists(fp):
            try: return ImageFont.truetype(fp, px)
            except Exception: pass
    return ImageFont.load_default()

def _b64(p):
    return "data:image/png;base64," + base64.b64encode(open(p, "rb").read()).decode()

def _slide_num(path):
    m = re.search(r"(\d+)", os.path.basename(path))
    return int(m.group(1)) if m else 0

def render_rebuilt(clean_png, blocks, out_png):
    bg = Image.open(clean_png).convert("RGB")
    W, H = bg.size
    d = ImageDraw.Draw(bg)
    for b in blocks:
        t = latex2plain((b.get("text") or "").strip())
        box = b.get("box_2d")
        if not t or not box or len(box) != 4:
            continue
        ymin, xmin, ymax, xmax = box
        fp = float(b.get("font_size_pt") or 16)
        longest = max((len(ln) for ln in t.split("\n")), default=1)
        box_w_pt = (xmax - xmin) / 1000 * 13.333 * 72
        if longest > 0:
            fp = min(fp, box_w_pt / (longest * 0.55) * 1.25)
        fp = max(8.0, min(fp, 48.0))
        px = max(int(fp * H / 540), 9)
        col = (b.get("color") or "FFFFFF").replace("#", "")
        try: rgb = tuple(int(col[i:i+2], 16) for i in (0, 2, 4))
        except Exception: rgb = (255, 255, 255)
        d.multiline_text((int(xmin/1000*W), int(ymin/1000*H)), t, fill=rgb, font=_font(px))
    bg.save(out_png)
    return out_png

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workdir", required=True, help="convert()에서 쓴 작업 디렉토리")
    ap.add_argument("--out", default="comparison.html")
    args = ap.parse_args()

    slides = sorted(glob.glob(os.path.join(args.workdir, "slides", "slide*.png")), key=_slide_num)
    if not slides:
        sys.exit(f"slides/ 에서 슬라이드 PNG를 찾지 못함: {args.workdir}")
    clean_dir = os.path.join(args.workdir, "clean")
    prev_dir = os.path.join(args.workdir, "preview")
    os.makedirs(prev_dir, exist_ok=True)

    rows = ""
    for sp in slides:
        n = _slide_num(sp)
        clean = os.path.join(clean_dir, f"clean{n}.png")
        bg = clean if os.path.exists(clean) else sp
        blocks = ocr_slide(sp)
        rebuilt = render_rebuilt(bg, blocks, os.path.join(prev_dir, f"rebuilt{n}.png"))
        rows += f"""
        <section><h2>슬라이드 {n}</h2><div class="pair">
          <figure><figcaption>원본 (편집 불가)</figcaption><img src="{_b64(sp)}"></figure>
          <figure><figcaption>재구성 (편집 가능)</figcaption><img src="{_b64(rebuilt)}"></figure>
        </div></section>"""

    html = f"""<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>변환 비교</title>
<style>body{{font-family:sans-serif;margin:0;background:#0f1620;color:#e8eef5}}
header{{padding:20px 28px;background:#16202e;border-bottom:2px solid #2a3a52}}
section{{padding:20px 28px;border-bottom:1px solid #243246}}
h2{{font-size:15px;color:#7fd0ff;margin:0 0 12px}}
.pair{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
figure{{margin:0}}figcaption{{font-size:12px;color:#9fb3c8;margin-bottom:6px}}
img{{width:100%;border:1px solid #2a3a52;border-radius:6px;display:block}}
@media(max-width:900px){{.pair{{grid-template-columns:1fr}}}}</style></head><body>
<header><h1>원본 vs 재구성 비교</h1></header>{rows}</body></html>"""
    open(args.out, "w", encoding="utf-8").write(html)
    print("SAVED:", args.out)

if __name__ == "__main__":
    main()
