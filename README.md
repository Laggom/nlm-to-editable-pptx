# nlm-to-editable-pptx

NotebookLM(또는 이미지 기반/스캔) PPTX·PDF 덱을 **편집 가능한 PowerPoint**로 변환하는 Claude Code 스킬.

NotebookLM이 내보낸 슬라이드는 각 페이지가 통이미지라 글자를 선택·편집할 수 없습니다.
이 스킬은 (1) 배경에서 글자를 지우고 (2) 글자를 위치·크기·색까지 다시 인식해서
(3) 깨끗한 배경 위에 진짜 편집 가능한 텍스트박스를 얹은 새 `.pptx`를 만듭니다.
원본의 그림·도표·색감은 유지되고, 한국어와 수식도 처리됩니다.

## 파이프라인

```
입력(pptx/pdf) → 슬라이드 PNG 추출 → [글자 제거: gpt-image-2] + [OCR: gpt-5.5] → pptx 조립
```

### 글자 제거 시 종횡비 보존 (v1.1.0)

`gpt-image-2`는 출력 사이즈가 고정(예: `1536×1024`, 3:2)이라, 16:9 슬라이드를 **그대로** 보내면
모델이 비율을 맞추려 그림·그래프를 재배치(reflow)해 레이아웃이 틀어집니다. 그러면 원본 좌표로
얹은 편집 텍스트가 배경과 어긋나 보입니다. v1.1.0은 이를 다음으로 해결합니다:

1. 원본을 출력 사이즈 캔버스에 **비율 유지 레터박스**로 패딩해 전송 → 모델이 비율을 바꿀
   필요가 없어 재구성하지 않고 글자만 지움
2. 결과에서 패딩 여백을 **크롭**해 원본 비율을 복원

여기에 "비텍스트 픽셀은 이동·크기변경·재구성하지 말라"는 **composition-lock 프롬프트**를 더해
글자 잔존·엉뚱한 내용 생성을 줄였습니다. 결과적으로 편집 텍스트가 배경과 정확히 정합됩니다
(특히 그래프·도표 슬라이드에서 효과가 큽니다). 자세한 내력은 [`CHANGELOG.md`](CHANGELOG.md).

## 결과 비교 (원본 → 변경 후)

각 쌍의 **왼쪽이 NotebookLM 원본**(통이미지, 편집 불가), **오른쪽이 변환 결과를
실제 PowerPoint/LibreOffice로 렌더한 모습**(원본 글자를 제거한 배경 위에 편집 가능한 텍스트).
한국어·수식·표가 모두 편집 가능한 텍스트로 재구성되며, 텍스트 위치·크기·줄바꿈이 원본에
맞도록 조정됩니다.

**표지**
![표지 비교](examples/compare_slide1_cover.png)

**수식 슬라이드** (LaTeX가 아닌 읽기 쉬운 유니코드로: `m₀/√(1−v²/c²)`)
![수식 비교](examples/compare_slide2_formula.png)

**그래프 슬라이드** (종횡비 보존 erase로 배경·축·범례가 텍스트와 정합 — v1.1.0)
![그래프 비교](examples/compare_slide5_graph.png)

**표 슬라이드**
![표 비교](examples/compare_slide8_table.png)

- 12슬라이드 전체 비교: [`examples/comparison.html`](examples/comparison.html) (내려받아 브라우저로 열기)
- **입력 샘플** (NotebookLM 원본, 편집 불가 통이미지):
  [`examples/sample_input_notebooklm.pptx`](examples/sample_input_notebooklm.pptx) (12슬라이드),
  [`examples/sample_input_notebooklm_2.pptx`](examples/sample_input_notebooklm_2.pptx) (15슬라이드)
- **출력 샘플** (변환 결과, 편집 가능): [`examples/sample_output.pptx`](examples/sample_output.pptx)

입력→출력 쌍이 함께 있으니, 직접 `python scripts/nlm2pptx.py examples/sample_input_notebooklm.pptx my_output.pptx`
로 재현해 결과를 비교해 볼 수 있습니다.

## 설치

```bash
pip install -r requirements.txt          # python-pptx, pillow, (pdf면) pymupdf
export OPENAI_API_KEY=sk-...              # Windows: setx OPENAI_API_KEY "sk-..."
```

Claude Code 스킬로 쓰려면 이 폴더를 `~/.claude/skills/nlm-to-editable-pptx/` 에 두면 됩니다
(Windows: `%USERPROFILE%\.claude\skills\nlm-to-editable-pptx\`).

## 사용

### CLI

```bash
python scripts/nlm2pptx.py input.pptx output.pptx
python scripts/nlm2pptx.py deck.pdf out.pptx --no-erase      # 빠른 모드(글자제거 생략)
python scripts/nlm2pptx.py input.pptx out.pptx --workers 6   # 병렬(기본 6, 1=순차)
python scripts/nlm2pptx.py input.pptx out.pptx --workdir ./wd  # 중간파일+convert.log 보존
python scripts/nlm2pptx.py input.pptx out.pptx --tables      # 표를 네이티브 표로 분리(아래 참고)
```

### 표 분리 (`--tables`)

기본 변환은 "배경 이미지 + 편집 가능한 텍스트"만 만듭니다. `--tables` 를 주면 **표를
편집 가능한 네이티브 PowerPoint 표로** 분리합니다:

- **표** → **네이티브 PowerPoint 표** (셀 텍스트·열너비·행높이·셀색을 원본에서 추정)
- **그림/차트/아이콘** → 인식하지 않고 **배경에 그대로 둠** (그림 인식 정확도가 낮아 제외)
- **배경** → 글자 + 표 영역을 제거한 깨끗한 이미지, 슬라이드 비율은 원본에 맞춤
- 표 셀에 들어가는 글자만 표로 흡수, 나머지 글자는 모두 편집 가능한 텍스트로 유지

표 검출은 hybrid(img2table + 비전 표 구조)이며 슬라이드당 비전 1회 호출이 추가됩니다.
`img2table`, `opencv-python` 미설치 시 표는 비전 검출 결과로 폴백합니다(requirements 의 optional 항목).

슬라이드별 글자제거/OCR은 기본 6스레드 병렬(12장 기준 전체 ~7분, `--no-erase` ~2.5분).
`--workdir` 를 주면 `convert.log` 에 슬라이드별 소요시간·재시도·에러가 기록됩니다.

### Python (노트북 / 웹앱 / Databricks)

```python
from nlm2pptx import convert
convert("input.pptx", "output.pptx")            # 전체
convert("input.pdf",  "out.pptx", erase=False)  # 빠른 모드
```

개별 단계(`extract_slides`, `erase_text`, `ocr_slide`, `build_pptx`)도 import 가능해
슬라이드 단위 병렬화가 쉽습니다. 자세한 내용은 `references/architecture.md`.

### 검증

```bash
python scripts/nlm2pptx.py input.pptx out.pptx --workdir ./wd
python scripts/compare_html.py --workdir ./wd --out comparison.html   # 원본 vs 결과 비교
```

## 환경 변수

| 변수 | 기본값 | 설명 |
|---|---|---|
| `OPENAI_API_KEY` | (필수) | OpenAI API 키 |
| `OPENAI_BASE_URL` | `https://api.openai.com/v1` | 게이트웨이/Azure 등 |
| `NLM2PPTX_IMAGE_MODEL` | `gpt-image-2` | 글자 제거 모델 |
| `NLM2PPTX_OCR_MODEL` | `gpt-5.5` | OCR 모델 |
| `NLM2PPTX_FONT` | `맑은 고딕` | pptx에 기입할 폰트명 |

## 참고

- 글자 제거는 슬라이드당 이미지 모델을 1회 호출(~30–60초). 12장이면 수 분 소요.
  속도/비용이 중요하면 `--no-erase`(글자 제거 없이 텍스트만 오버레이).
- 수식은 편집 가능한 유니코드 평문으로 변환(`m₀/√(1−v²/c²)`)됩니다.
- `gpt-image-2`는 `gpt-image-1-mini`보다 원본 그림/표를 충실히 보존합니다(기본값 유지 권장).
- 글자 제거는 원본 비율을 레터박스로 보존한 뒤 크롭합니다(v1.1.0). 출력 사이즈는 `NLM2PPTX_IMAGE_MODEL`
  모델이 허용하는 값이어야 하며 `IMAGE_SIZE` 상수로 정합니다(기본 `1536x1024`).
- 단, 모델이 글자 외 요소까지 크게 재구성하는 일부 슬라이드는 레터박스만으로 완전히 막지 못할 수
  있습니다. 그런 경우는 `--no-erase`(원본 배경 유지)가 더 안전합니다.

## 라이선스

MIT
