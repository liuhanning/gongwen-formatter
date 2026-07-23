# rev.docx format extraction vs current standard

- Canonical sample: `贵州省低空基础设施和信息服务行业研究报告（20260413）rev.docx`
- Current standard file: `咨询报告统一格式要求.md`
- Extracted: 524 non-empty paragraphs, 6 tables, 3 sections

## Representative paragraph formats

| Item | Paragraph | Style | Align | Font | Size | Bold | Indent / spacing |
| --- | --- | --- | --- | --- | --- | --- | --- |
| cover title 1 | #2 贵州省“十五五”时期 | - | center | 方正小标宋简体 | 22pt | False | left=-, first=-, line=18pt, before=-, after=0pt |
| cover title 2 | #3 低空基础设施和信息服务行业研究报告 | - | center | 方正小标宋简体 | 22pt | False | left=-, first=-, line=18pt, before=-, after=0pt |
| cover org | #10 中咨海外咨询有限公司 | - | center | 楷体_GB2312 | 16pt | False | left=-, first=0pt, line=28pt, before=-, after=0pt |
| cover date | #11 2026 年 4 月 | - | center | 仿宋_GB2312 | 16pt | False | left=-, first=0pt, line=28pt, before=-, after=0pt |
| toc title | #12 目   录 | heading 1 | center | 黑体 | 20pt | True | left=-, first=-, line=30pt, before=6pt, after=6pt |
| toc 1 | #14 第一章 贵州省低空基础设施和信息服务行业研究背景- 1 - | toc 1 | - | 仿宋 | 16pt | True | left=-, first=-, line=-, before=-, after=- |
| toc 2 | #15 第一节 国家和贵州省低空经济发展情况- 1 - | toc 2 | - | 仿宋 | 16pt | False | left=22pt, first=-, line=-, before=-, after=- |
| toc 3 | #16 一、国家低空经济发展情况- 1 - | toc 3 | - | 仿宋 | 16pt | False | left=44pt, first=-, line=-, before=-, after=- |
| body paragraph | #150 国家低空经济发展的经济基础已经显现。我国低空经济已进入由离散的科研验证、小范围商业试点向规模化、业务化应用拓展的阶段。2024 年 3 月， | - | both | 仿宋_GB2312 | 16pt | False | left=-, first=32pt, line=28pt, before=-, after=0pt |
| table title | #181 表 2-1 贵州省市（州）飞行基础设施现状表（截至 2025 年底） | - | center | 黑体 | 12pt | False | left=-, first=-, line=12pt, before=6pt, after=6pt |
| figure title | #168 图 1-1 贵州省运输机场分布表 | - | both | 仿宋_GB2312 | 16pt | False | left=-, first=32pt, line=28pt, before=-, after=0pt |

## Sections and page numbering

| Section | Page | Margins | Page number | Grid | Footer refs |
| --- | --- | --- | --- | --- | --- |
| 1 | 21.00cm x 29.70cm | top 3.70cm, bottom 3.50cm, left 2.80cm, right 2.60cm, header 1.80cm, footer 1.80cm | fmt=-, start=- | type=-, linePitch=360 | - |
| 2 | 21.00cm x 29.70cm | top 3.70cm, bottom 3.50cm, left 2.80cm, right 2.60cm, header 1.80cm, footer 1.80cm | fmt=upperRoman, start=1 | type=-, linePitch=360 | footer1.xml:22 / 2 / 2 |
| 3 | 21.00cm x 29.70cm | top 3.70cm, bottom 3.50cm, left 2.80cm, right 2.60cm, header 1.80cm, footer 1.80cm | fmt=numberInDash, start=1 | type=-, linePitch=360 | footer2.xml:77 / 7 / 7 |

## Tables

| Table | Align | Borders | Cell margins | First cell sample |
| --- | --- | --- | --- | --- |
| 1 | center | top:single/4/000000, left:single/4/000000, bottom:single/4/000000, right:single/4/000000, insideH:single/4/000000, insideV:single/4/000000 | top 0pt, left 5.4pt, bottom 0pt, right 5.4pt | 黑体 12pt; 机场类型 |
| 2 | center | top:single/4/000000, left:single/4/000000, bottom:single/4/000000, right:single/4/000000, insideH:single/4/000000, insideV:single/4/000000 | top 0pt, left 5.4pt, bottom 0pt, right 5.4pt | 黑体 12pt; 市（州）名称 |
| 3 | center | top:single/4/000000, left:single/4/000000, bottom:single/4/000000, right:single/4/000000, insideH:single/4/000000, insideV:single/4/000000 | top 0pt, left 5.4pt, bottom 0pt, right 5.4pt | 黑体 12pt; 市州 |
| 4 | center | top:single/4/000000, left:single/4/000000, bottom:single/4/000000, right:single/4/000000, insideH:single/4/000000, insideV:single/4/000000 | top 0pt, left 5.4pt, bottom 0pt, right 5.4pt | 黑体 12pt; 设施类型 |
| 5 | center | top:single/4/000000, left:single/4/000000, bottom:single/4/000000, right:single/4/000000, insideH:single/4/000000, insideV:single/4/000000 | top 0pt, left 5.4pt, bottom 0pt, right 5.4pt | 黑体 12pt; 设施类型 |
| 6 | center | top:single/4/000000, left:single/4/000000, bottom:single/4/000000, right:single/4/000000, insideH:single/4/000000, insideV:single/4/000000 | top 0pt, left 5.4pt, bottom 0pt, right 5.4pt | 黑体 12pt; 指标 |

## Comparison summary

| Module | Current standard | rev.docx extraction | Assessment |
| --- | --- | --- | --- |
| Page setup | A4; margins top 3.7, bottom 3.5, left 2.8, right 2.6 cm; header/footer 1.8 cm; grid aligned | 21.00cm x 29.70cm; top 3.70cm, bottom 3.50cm, left 2.80cm, right 2.60cm, header 1.80cm, footer 1.80cm; type=-, linePitch=360 | Matches size/margins/header/footer; docGrid linePitch exists, type is implicit/empty. |
| Cover title | FZ Xiaobiao Song 22pt, centered, first-line indent 0 | 方正小标宋简体 22pt and 方正小标宋简体 22pt; align=center; bold=False | Matches. |
| Cover org | Org full name, centered, no prompt labels | 中咨海外咨询有限公司; 楷体_GB2312 16pt; align=center; first=0pt | Matches; standard should explicitly name this font. |
| Cover date | Current standard only says cover paragraphs centered/no first-line indent | 2026 年 4 月; 仿宋_GB2312 16pt; align=center; first=0pt | Gap: add explicit date font rule. |
| TOC title | Hei 20pt centered | 黑体 20pt; align=center; bold=True | Mostly matches; decide whether bold is mandatory. |
| TOC entries | Fangsong 16pt; preserve tab leader/page number; level indentation | TOC1 仿宋 16pt bold=True left=0; TOC2 left=22pt; TOC3 left=44pt | Matches; add TOC1 bold and exact indent guidance if canonical. |
| Page numbering | Cover no page no.; TOC upperRoman; body starts at 1; body page format - 1 - | sec1 fmt=- start=-; sec2 fmt=upperRoman start=1; sec3 fmt=numberInDash start=1 | Matches; body uses Word numberInDash. |
| Body heading hierarchy | Chapter 22pt centered; section 16pt centered; lower levels by numbering | Representative style hierarchy is heading 1/2/3/4 plus numbered body headings; see paragraph table. | Mostly matches. |
| Body text | Fangsong_GB2312 16pt; exact 28pt line; justified; first-line indent 2 chars | 仿宋_GB2312 16pt; align=both; first=32pt; line=28pt | Matches; firstLine is 32pt in XML. |
| Tables | Black single borders; cell margins L/R 5.4pt, T/B 0; header Hei 12pt; body Fangsong 12pt | 6 tables; top:single/4/000000, left:single/4/000000, bottom:single/4/000000, right:single/4/000000, insideH:single/4/000000, insideV:single/4/000000; top 0pt, left 5.4pt, bottom 0pt, right 5.4pt; first row 黑体 12pt; 机场类型 | Matches table structure/header. |

## Recommended standard updates

- Treat `贵州省低空基础设施和信息服务行业研究报告（20260413）rev.docx` as the canonical sample for cover, TOC, sections/page numbering, heading hierarchy, body, and tables.
- Add explicit cover date rule: `仿宋_GB2312 16pt, centered, first-line indent 0`.
- Add explicit cover org rule: `楷体_GB2312 16pt, centered, first-line indent 0`.
- Decide whether TOC title bold and TOC1 bold should be hard requirements; rev.docx has both.
- Preserve TOC style indentation/tab stops; do not flatten TOC entries when formatting.
- Body page number should use Word `numberInDash` behavior for `- 1 -` style.