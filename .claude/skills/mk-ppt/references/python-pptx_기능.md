# python-pptx 1.0.2 기능 참고

`scripts/build_pptx.py`가 지원하지 않는 것을 직접 코드로 만들 때 읽는다.
아래 내용은 이 PC에 설치된 1.0.2 소스 코드로 확인했다.

## 목차
1. 프레젠테이션 파일
2. 슬라이드
3. 도형 추가
4. 도형 공통 속성
5. 텍스트
6. 표
7. 차트
8. 색·채우기·선
9. 동작과 링크
10. 단위
11. 지원하지 않는 기능과 우회법
12. 자주 틀리는 점

---

## 1. 프레젠테이션 파일
```python
from pptx import Presentation
from pptx.util import Inches
prs = Presentation()                       # 새 파일 (기본 4:3)
prs = Presentation("기존.pptx")            # 열기 (.pptx만 가능, .ppt 불가)
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)   # 16:9
prs.core_properties.title = "제목"          # author, keywords, subject, created 등
prs.save("결과.pptx")
```
- 마스터·레이아웃: `prs.slide_master`, `prs.slide_masters`, `prs.slide_layouts`, `prs.notes_master`

## 2. 슬라이드
```python
slide = prs.slides.add_slide(prs.slide_layouts[6])   # 0 제목, 1 제목+내용, 5 제목만, 6 빈 화면
prs.slides.index(slide); prs.slides.get(slide_id)
slide.background.fill.solid(); slide.background.fill.fore_color.rgb = RGBColor(0x1F, 0x38, 0x64)
slide.follow_master_background                        # 마스터 배경 사용 여부
slide.notes_slide.notes_text_frame.text = "발표자 노트"
prs.slide_layouts.get_by_name("Title Only"); prs.slide_layouts.remove(layout)  # 사용 중이 아닌 레이아웃만
layout.used_by_slides
```

## 3. 도형 추가 (`slide.shapes`)
| 기능 | 코드 |
|---|---|
| 텍스트 상자 | `add_textbox(left, top, w, h)` |
| 자동 도형 | `add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)` — `pptx.enum.shapes.MSO_SHAPE`에 180여 종 |
| 그림 | `add_picture("a.png", l, t, width=None, height=None)` — 한쪽만 주면 비율 유지 |
| 표 | `add_table(rows, cols, l, t, w, h).table` |
| 차트 | `add_chart(XL_CHART_TYPE.X, l, t, w, h, chart_data).chart` |
| 연결선 | `add_connector(MSO_CONNECTOR.ELBOW, x1, y1, x2, y2)` → `begin_connect(shape, 0)`, `end_connect(shape, 2)` |
| 그룹 | `grp = add_group_shape(); grp.shapes.add_shape(...)` |
| 동영상 | `add_movie("a.mp4", l, t, w, h, poster_frame_image="p.png", mime_type="video/mp4")` |
| OLE 개체 | `add_ole_object("a.xlsx", PROG_ID.XLSX, l, t)` |
| 자유형 | `fb = build_freeform(x, y); fb.add_line_segments([(x1,y1),(x2,y2)], close=True); fb.convert_to_shape()` |
| 개체 틀 채우기 | `placeholder.insert_picture(path)`, `insert_table(r, c)`, `insert_chart(type, data)` |

## 4. 도형 공통 속성
- `left`, `top`, `width`, `height`, `rotation`, `name`, `shape_id`, `shape_type`
- `shape.fill`, `shape.line`, `shape.shadow.inherit = False`(그림자 끄기)
- 자동 도형 조절점: `shape.adjustments[0] = 0.2`
- 그림 자르기: `pic.crop_left = 0.1`(비율), 원본: `pic.image.blob`, `pic.image.ext`
- 자동 도형 안에도 글자를 넣을 수 있음: `shape.text_frame`

## 5. 텍스트
```python
tf = shape.text_frame
tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.margin_left = Inches(0.1)
tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE   # 표시 옵션일 뿐, 실제 글자 크기는 PowerPoint가 열 때 조정
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; p.level = 1
p.line_spacing = 1.2; p.space_before = Pt(6); p.space_after = Pt(6)
run = p.add_run(); run.text = "글자"
run.font.size = Pt(18); run.font.bold = True; run.font.color.rgb = RGBColor.from_string("C55A11")
run.hyperlink.address = "https://..."
p.add_line_break()
```
- `tf.fit_text(font_family, max_size)`는 글꼴 파일(.ttf)이 필요해 맑은 고딕(.ttc)에서는 실패할 수 있다.

## 6. 표
```python
table = slide.shapes.add_table(4, 3, l, t, w, h).table
table.columns[0].width = Inches(2); table.rows[0].height = Inches(0.5)
cell = table.cell(0, 0); cell.text = "머리글"
cell.merge(table.cell(0, 2))      # 병합, cell.split()으로 해제
cell.fill.solid(); cell.fill.fore_color.rgb = ...
cell.vertical_anchor = MSO_ANCHOR.MIDDLE; cell.margin_left = Inches(0.1)
table.first_row = True; table.horz_banding = True   # last_row, first_col, last_col, vert_banding
```

## 7. 차트
- 종류(`XL_CHART_TYPE`): `COLUMN_CLUSTERED/STACKED`, `BAR_CLUSTERED/STACKED`, `LINE`, `LINE_MARKERS`, `PIE`, `DOUGHNUT`, `AREA`, `XY_SCATTER`, `BUBBLE`, `RADAR` 등 3D 포함
- 데이터: `CategoryChartData`(범주형), `XyChartData`(분산형), `BubbleChartData`(거품형)
```python
data = CategoryChartData(); data.categories = ["1분기", "2분기"]
data.add_series("매출", (120, 150), number_format="#,##0")
chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, l, t, w, h, data).chart
chart.has_title = True; chart.chart_title.text_frame.text = "제목"
chart.has_legend = True; chart.legend.position = XL_LEGEND_POSITION.BOTTOM; chart.legend.include_in_layout = False
chart.value_axis.maximum_scale = 200; chart.value_axis.major_unit = 50
chart.value_axis.has_major_gridlines = False; chart.value_axis.tick_labels.number_format = '0"%"'
chart.category_axis.reverse_order = True
plot = chart.plots[0]; plot.gap_width = 80; plot.overlap = 0; plot.vary_by_categories = False
plot.has_data_labels = True; plot.data_labels.show_value = True; plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
plot.series[0].format.fill.solid(); plot.series[0].points[1].format.fill.fore_color.rgb = ...   # 막대 하나만 강조
plot.series[0].smooth = True; plot.series[0].marker.style = XL_MARKER_STYLE.CIRCLE   # 꺾은선
chart.replace_data(new_data)   # 기존 차트 데이터 교체
```

## 8. 색·채우기·선
- 채우기: `fill.solid()`, `fill.gradient()` + `gradient_angle`, `gradient_stops`, `fill.patterned()` + `pattern`, `fill.background()`(채우기 없음)
- 색: `color.rgb = RGBColor(...)`, `color.theme_color = MSO_THEME_COLOR.ACCENT_1`, `color.brightness = 0.4`(-1~1)
- 선: `line.color.rgb`, `line.width = Pt(1.5)`, `line.dash_style = MSO_LINE.DASH`, 선 없애기 `line.fill.background()`

## 9. 동작과 링크
```python
shape.click_action.hyperlink.address = "https://..."
shape.click_action.target_slide = prs.slides[3]     # 다른 슬라이드로 이동
run.hyperlink.address = "https://..."
```

## 10. 단위
`Inches`, `Cm`, `Mm`, `Pt`, `Emu`, `Centipoints` — 모두 EMU 정수(1 in = 914400 EMU). `.inches`, `.pt`, `.cm`로 되돌려 읽는다.

---

## 11. 지원하지 않는 기능과 우회법
| 기능 | 상태 | 우회법 |
|---|---|---|
| 슬라이드 삭제 | API 없음 | 아래 XML 코드 |
| 슬라이드 순서 변경 | API 없음 | 아래 XML 코드 |
| 슬라이드 복제 | API 없음 | 새 슬라이드를 만들고 도형 XML을 `copy.deepcopy`로 복사(그림·차트는 관계(rel)도 복사해야 해서 까다로움 — 가능하면 처음부터 다시 생성) |
| 도형 삭제 | API 없음 | `sp = shape._element; sp.getparent().remove(sp)` |
| 애니메이션·전환 효과 | 없음 | 지원하지 않는다고 사용자에게 알린다 |
| SmartArt | 없음 | 자동 도형과 연결선으로 비슷하게 그린다 |
| 그림자 세부 설정, 반사, 네온, 3D | 없음(`shadow.inherit`만) | XML 직접 편집. 보통은 생략 |
| 오디오 | 없음 | 지원하지 않는다고 알린다 |
| PDF·이미지 변환 | 없음 | PowerPoint에서 직접 저장하도록 안내 |
| .ppt(구버전), 암호 파일 | 열 수 없음 | PowerPoint에서 .pptx로 다시 저장하도록 안내 |

```python
# 슬라이드 삭제 (index 번째)
xml_slides = prs.slides._sldIdLst
sld = xml_slides[index]
prs.part.drop_rel(sld.rId)
xml_slides.remove(sld)

# 슬라이드 순서 변경 (old → new 위치)
xml_slides = prs.slides._sldIdLst
el = xml_slides[old]
xml_slides.remove(el)
xml_slides.insert(new, el)
```

## 12. 자주 틀리는 점
- **한글 글꼴:** `font.name`은 `a:latin`만 바꾼다. 한글은 `a:ea` 글꼴을 쓰므로 그것도 지정해야 맑은 고딕으로 나온다. `build_pptx.py`의 `apply_font()`를 참고한다.
- **글머리 기호:** 빈 레이아웃의 텍스트 상자에는 기호가 없다. `a:buChar`를 XML로 넣어야 한다(`set_bullet()` 참고).
- **넘침 확인 불가:** python-pptx는 글자를 렌더링하지 않는다. 그래서 글자가 상자를 넘치는지 알 수 없다. 글자 수와 글머리표 개수를 보수적으로 잡는다.
- **기본 크기는 4:3:** `Presentation()`의 기본 크기는 10 x 7.5 in이다. 16:9가 필요하면 직접 바꾼다.
