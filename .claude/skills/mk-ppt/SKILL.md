---
name: mk-ppt
description: agent_09 프로젝트에서 python-pptx로 PowerPoint 발표자료(.pptx)를 만든다. research/의 자료조사 .md나 report/의 보고서 내용을 슬라이드로 옮기거나, 주제만 주고 발표자료를 만들어 달라는 요청에 쓴다. "PPT 만들어줘", "발표자료", "슬라이드로 정리", "프레젠테이션", "파워포인트", ".pptx", "발표용으로" 같은 표현이 나오면 형식을 명시하지 않았더라도 이 스킬을 사용한다. 기존 .pptx에 표·차트·노트를 추가하거나 고치는 작업에도 쓴다.
---

# mk-ppt: 발표자료(.pptx) 만들기

이 프로젝트(agent_09)의 규칙에 맞춰 python-pptx로 발표자료를 만든다.
슬라이드는 직접 코드로 그리지 않는다. **슬라이드 내용을 JSON 사양(spec)으로 쓰고 `scripts/build_pptx.py`에 넘기는 방식**을 기본으로 한다.
이렇게 하면 디자인(맑은 고딕, 색, 제목 띠, 쪽 번호)이 매번 같게 나오고, 한글 글꼴·글머리표처럼 python-pptx에서 자주 틀리는 부분을 스크립트가 처리한다.

## 실행 환경

- Python은 PATH에 없다. 항상 전체 경로로 실행한다.
  `C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe`
- python-pptx 1.0.2, Pillow, lxml이 설치되어 있다.
- 아래에서 `$PY`는 위 python.exe 경로, `$SKILL`은 `.claude/skills/mk-ppt`를 뜻한다.

## 작업 순서

### 1. 프로젝트 규칙 따르기
CLAUDE.md의 작업 규칙이 이 스킬보다 우선한다.
- 실행 전에 todo list를 보고하고 승인을 받는다.
- 이 스킬로 만드는 파일도 todo list에 적는다. 슬라이드 개요(제목 목록)를 함께 보여 주면 승인하는 사람이 결과를 미리 판단할 수 있다.

### 2. 원본 자료 확인
- 사용자가 주제만 말하면, 먼저 `research/`에서 같은 주제의 최신 버전(`주제_자료조사_vN.md` 중 N이 가장 큰 것)을 찾는다.
- 있으면 그 내용을 쓰고, 없으면 자료조사가 먼저 필요한지 사용자에게 묻는다.
- 발표자료에 들어가는 사실은 원본 자료에 있는 것만 쓴다. 자료에 없는 내용을 지어내면 보고서와 발표자료가 서로 어긋나게 된다.

### 3. 슬라이드 개요 짜기
- 보고서와 마찬가지로 **서론 → 본론 → 결론** 흐름으로 짠다.
- 기본 구성(자료 분량에 따라 10~18장)
  1. 표지(`title`)
  2. 서론: 발표 목차 또는 배경(`bullets`)
  3. 본론: 주제별 `section` 간지 + 내용 슬라이드
  4. 결론: 핵심 정리·의의(`bullets`)
  5. 참고 출처(`bullets`, 작은 글씨)
  6. 마무리(`closing`)
- 자료의 형태에 맞는 슬라이드를 고른다.
  - 원본 자료의 **표는 `table`**로 옮긴다.
  - **연도·수량 같은 숫자 비교는 `chart`**로 바꾸면 한눈에 들어온다.
  - **대비되는 두 가지(사실 vs 각색, 장점 vs 한계)는 `two_column`**으로 보여 준다.
- 슬라이드 하나에 글머리표는 **6개 이하, 한 줄 40자 안팎**으로 한다. python-pptx는 글자가 상자를 넘쳐도 알려 주지 않기 때문에 처음부터 짧게 쓰는 것이 유일한 방법이다. 내용이 많으면 슬라이드를 나눈다.
- 문장보다 핵심어 위주로 쓰고, 자세한 설명은 `notes`(발표자 노트)에 넣는다.

### 4. spec 작성
- spec JSON은 프로젝트 안이 아닌 임시 폴더에 저장한다. 세션 scratchpad가 있으면 그곳을 쓰고, 없으면 `$env:TEMP`를 쓴다. UTF-8로 저장한다.
- 형식은 아래 "spec 형식"을 따른다.

### 5. 생성
```powershell
& $PY "$SKILL\scripts\build_pptx.py" <spec.json> "ppt\주제_발표자료_vN.pptx"
```
- 저장 위치는 `ppt/`, 파일 이름은 `주제_발표자료_vN.pptx`다.
- 같은 이름이 이미 있으면 스크립트가 덮어쓰지 않고 멈춘다. `ppt/`를 확인해 다음 버전 번호를 쓴다. 기존 파일을 고쳐 달라는 요청도 새 버전 파일로 만든다.
- 먼저 임시 폴더에 생성해 본다. 출력에 `경고:`가 있으면(글머리표 과다, 긴 문장, 표 행 과다) spec을 고쳐 경고가 없어질 때까지 다시 생성한다. 그다음 최종본만 `ppt/`에 생성한다. 이렇게 해야 `ppt/`에 중간 버전이 쌓이지 않는다.

### 6. 검증
```powershell
& $PY "$SKILL\scripts\build_pptx.py" --check "ppt\주제_발표자료_vN.pptx"
```
다음을 확인한다.
- 슬라이드 수와 각 슬라이드 제목이 개요와 같은가
- 사용 글꼴이 맑은 고딕뿐인가
- "슬라이드 밖으로 나감" 경고가 없는가

결과 보고에는 저장 경로, 슬라이드 목록, 확인하지 못한 점을 적는다. 글자 넘침은 PowerPoint로 열어 봐야만 알 수 있으므로 그 점을 함께 알린다.

### 7. 마무리
- 결과 파일은 `ppt/`에만 둔다. 루트에 두지 않는다.
- 새 폴더가 생겼으면 CLAUDE.md의 폴더 구조에 추가한다.

## spec 형식

```json
{
  "title": "허준과 동의보감",
  "author": "작성자",
  "footer": "허준과 동의보감",
  "size": "16:9",
  "font": "맑은 고딕",
  "theme": {"primary": "1F3864", "accent": "C55A11", "text": "222222", "muted": "6B7280", "light": "EEF2F7"},
  "slides": [
    {"type": "title", "title": "허준과 동의보감", "subtitle": "조선 의학을 집대성한 의학자\n2026-10-13", "notes": "발표자 노트"},
    {"type": "section", "label": "본론 1", "title": "생애와 관직 활동"},
    {"type": "bullets", "title": "내의원 입문", "bullets": ["서자 출신 — 잡과인 의학의 길", {"text": "1569년 유희춘의 추천", "level": 1}]},
    {"type": "two_column", "title": "사실과 각색",
     "left": {"heading": "기록된 사실", "bullets": ["..."]},
     "right": {"heading": "드라마 속 창작", "bullets": ["..."]}},
    {"type": "table", "title": "동의보감의 구성", "columns": ["편", "권수", "내용"],
     "rows": [["내경편", "4권", "몸 안의 장부와 생리"]], "col_widths": [1, 1, 4]},
    {"type": "chart", "title": "편별 권수", "chart_type": "column",
     "categories": ["내경편", "외형편"], "series": [{"name": "권수", "values": [4, 4]}]},
    {"type": "image", "title": "신형장부도", "path": "C:/경로/그림.png", "caption": "설명"},
    {"type": "closing", "title": "감사합니다", "subtitle": "질의응답"}
  ]
}
```

| type | 필수 키 | 선택 키 |
|---|---|---|
| `title`, `closing` | `title` | `subtitle`(줄바꿈 `\n`) |
| `section` | `title` | `label`(예: "본론 1") |
| `bullets` | `title`, `bullets` | `font_size`(기본 20) |
| `two_column` | `title`, `left`, `right` (`heading`, `bullets`) | `font_size`(기본 18) |
| `table` | `title`, `columns`, `rows` | `col_widths`(비율), `font_size`(기본 16, 7행 이상 13) |
| `chart` | `title`, `categories`, `series` | `chart_type`: column, column_stacked, bar, bar_stacked, line, pie, doughnut, area / `data_labels`(기본 true) |
| `image` | `title`, `path` | `caption` |

- 모든 슬라이드에 `notes`(발표자 노트)를 넣을 수 있다.
- `bullets` 항목은 문자열이거나 `{"text": ..., "level": 0~2}` 형식이다.
- 표의 행은 본문 영역에 **8행 정도까지** 들어간다. 그보다 많으면 슬라이드를 나눈다.
- 원형(`pie`, `doughnut`) 차트는 백분율 레이블을 자동으로 단다. 계열이 1개인 막대·꺾은선 차트는 `primary` 색을 쓴다.

## 스크립트로 안 되는 작업

- 다음은 python-pptx로 코드를 직접 짠다.
  - 기존 .pptx 수정
  - 하이퍼링크, 도형 다이어그램, 연결선
  - 분산형·거품형 차트
  - 슬라이드 삭제·순서 변경
- 이때 `references/python-pptx_기능.md`를 읽는다. 기능별 코드 조각, 지원하지 않는 기능의 우회법, 자주 틀리는 점(한글 `ea` 글꼴, 글머리 기호)이 있다.
- 코드를 직접 짤 때도 `build_pptx.py`의 `apply_font()`, `set_bullet()`을 가져다 쓰면 글꼴과 글머리표가 기존 결과물과 같게 나온다.
- 기존 파일을 수정할 때도 원본을 열어 고친 뒤 **새 버전 이름으로 저장**한다.

## 지원하지 않는 요청

다음 기능은 python-pptx로 할 수 없다. 그럴 듯하게 흉내 내지 말고 사용자에게 바로 알린다.
- 애니메이션, 화면 전환 효과
- SmartArt
- 오디오
- PDF나 이미지로 변환
- .ppt 파일 열기
