# 객체 분리(Object Separation) 설계 — A/B/C 비교

날짜: 2026-06-22
대상 저장소: `nlm-to-editable-pptx` (Python 포팅본)

## 1. 목표 & 범위

NotebookLM 통이미지 슬라이드를 변환할 때, 지금은 "글자만" 편집 가능하게 하고 표·그림·차트·아이콘은 모두 배경 이미지 한 장에 남긴다. 이 기능은 **비텍스트 객체를 인식해 개별 객체로 분리**한다:

- **표** → 네이티브 PowerPoint 표(`python-pptx` `add_table`, 셀 텍스트 = 기존 OCR 매핑)
- **사진/일러스트/아이콘/차트** → 원본에서 잘라 **개별 이미지**(`add_picture`, 같은 위치)
- 나머지 텍스트 → 기존 `build_pptx` 텍스트박스

핵심 결론(사전 조사): 인식(레이아웃 검출·표 구조 인식)은 **성숙한 오픈소스를 재사용**하고, pptx 조립 글루만 직접 구현한다. "이미지→객체분리 편집가능 pptx" 완성품은 존재하지 않음.

## 2. 아키텍처

```
원본 슬라이드 PNG ─┐
                  ├─► Detector.detect(image, ocr_blocks) ─► objects:[{type, box_2d, table?:{rows,cols,cells}}]
wd/ocr 캐시 ───────┘                                          │
                                                              ▼
                                       Assembler (공통, 검출기 무관)
                                         ├─ table  → add_table(셀=OCR 텍스트 좌표 매핑)
                                         ├─ image/icon/chart → 원본 crop → add_picture(동일 위치)
                                         └─ 그 외 텍스트 → 기존 build_pptx 텍스트박스
                                                              ▼
                                       SK_objsep_{A|B|C}.pptx
```

- `Detector`: 단일 인터페이스 `detect(image_path, ocr_blocks) -> list[DetectedObject]`. A/B/C는 그 구현체.
- `DetectedObject`: `{type: "table"|"image"|"icon"|"chart", box_2d:[ymin,xmin,ymax,xmax](0–1000), confidence, table: {n_rows, n_cols, cells:[{r,c,box_2d}]}? }`
- `Assembler`: 검출 결과 + OCR 블록 + 배경을 받아 pptx 생성. **검출 방식과 독립** → A/B/C가 공유.

## 3. 세 검출기 (라이선스 안전: MIT/Apache/BSD만)

| | 라이브러리 | 라이선스 | 설치/비용 | 비고 |
|---|---|---|---|---|
| **A 경량** | `img2table` + OpenCV | MIT | pip, 소형, 무료 | 괘선 표 강함, 무괘선·음영표 약함 |
| **B 모델** | MS Table Transformer(검출+구조) + OpenCV(그림영역) | MIT (+torch BSD) | torch/transformers ~2GB | 무괘선 표도 강함 |
| **C 비전** | 기존 gpt-5.5에 객체+표구조 JSON 추가 | (API) | 설치 0, API 비용 | 레이아웃/무괘선 무관, 한글 강함 |

- 표 셀 텍스트는 세 방식 모두 **기존 OCR 좌표를 셀 격자에 매핑**(중복 OCR 없음).
- 라이선스 회피: DocLayout-YOLO(AGPL), surya/marker(커스텀/상업 임계), Detectron2 기반 LayoutParser(Windows 설치난).

## 4. 공통 Assembler

- 표: `slide.shapes.add_table(rows, cols, x, y, w, h)`; 셀 박스에 포함되는 OCR 블록 텍스트를 해당 셀에 기입. 표 영역의 텍스트박스는 표로 흡수(중복 방지).
- 이미지/아이콘/차트: 원본 PNG에서 `box_2d` 크롭 → `add_picture` 동일 위치. 배경엔 남겨두되 오버레이가 덮음(이동 시에만 아래 중복; v1 허용). 표는 불투명 셀이 배경을 덮으므로 inpaint 불필요.
- 폴백: 검출 실패/저신뢰 → 그 영역은 기존 방식(배경에 남김). 표 셀 매핑 실패 → 그 표는 crop 이미지로.

## 5. 비교 하니스 (핵심 산출물)

- A·B·C **병렬 실행**으로 각각 `SK_objsep_A/B/C.pptx` 생성.
- LibreOffice 렌더 → **원본/A/B/C 4단 비교 HTML**.
- 방식별 **지표 표**: 소요시간, 추가 API 호출 수(비용 프록시), 설치 용량, 표 인식 성공/실패 슬라이드 수.
- 이 지표·시각 비교로 채택안 결정.

## 6. 테스트 입력 (확보 완료)

- **그림/아이콘 분리**: NBLM2PPTX(github laihenyi) 데모의 **실제 NotebookLM 슬라이드 이미지** 사용 — `demo-before.png`(만화 일러스트+텍스트+인용박스), `demo-v1.1-original.jpg`(주방 일러스트+제목+불릿+빨간 금지 아이콘). 인터넷의 진짜 NotebookLM 산출물. (라이선스 MIT)
  - 위치: `D:\Work\coding\lm2ppt\_nblm_samples\assets\`
  - 단일 슬라이드들 → 합쳐 소규모 멀티슬라이드 테스트 덱 구성.
- **표 분리**: 이 NotebookLM 샘플엔 표가 없음 → **SK 슬라이드 5·6을 로컬 전용**으로 표 케이스 테스트(기밀, 비공개/비배포).
- 참고: NotebookLM은 2026.2부터 .pptx 직접 내보내기(편집가능)도 생겼으나, PDF/구버전 내보내기는 여전히 통이미지 → 본 기능 대상.

## 7. 에러 처리 / 폴백

- 모든 검출 단계는 실패 시 "기존 동작(배경 이미지 + 텍스트박스)"로 안전 폴백 → 결과물이 절대 깨지지 않음.
- 표 구조 신뢰도 임계 미만 → crop 이미지로 강등.

## 8. 범위 밖 (YAGNI)

- 차트 → 편집 가능 차트 객체 재구성 ❌ (crop 이미지까지만)
- 병합 셀/회전/중첩 표 정밀 복원 ❌ (단순 행·열만)

## 9. 환경 리스크 / 로지스틱스

- **사내 SSL(self-signed in chain)**: A·B 신규 패키지(특히 B torch ~2GB) pip 다운로드가 막힐 수 있음 → `--trusted-host` 또는 사내 CA 필요. C는 설치가 없어 영향 없음.
- **git push 대상 미정**: origin이 `Laggom/...`(권한 없음), gh CLI 없음. 현재는 로컬 커밋만 복원 지점. push하려면 사용자 소유 원격 필요.

## 10. 성공 기준

- 실제 NotebookLM 입력에서: 표 슬라이드 → 네이티브 표(행·열·셀 텍스트), 그림/아이콘 → 개별 이미지로 분리.
- 4단 비교 HTML + 지표 표로 A/B/C의 품질·시간·비용을 한눈에 비교해 채택 결정 가능.
