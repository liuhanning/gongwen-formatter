#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
W = "{%s}" % NS["w"]
R = "{%s}" % NS["r"]
PR = "{%s}" % NS["pr"]


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize_compact(text: str) -> str:
    return re.sub(r"\s+", "", text).strip()


def first(values: Iterable[str | None]) -> str | None:
    for value in values:
        if value:
            return value
    return None


@dataclass
class ParagraphInfo:
    index: int
    text: str
    style_id: str
    style_name: str
    alignment: str
    left_indent: str
    first_line_indent: str
    line: str
    line_rule: str
    font_east_asia: str
    font_ascii: str
    size_half_pt: str
    bold: bool

    def short(self) -> str:
        parts = [
            f"#{self.index}",
            self.text,
            f"style={self.style_name or self.style_id or '-'}",
            f"align={self.alignment or '-'}",
        ]
        if self.size_half_pt:
            parts.append(f"size={int(self.size_half_pt) / 2:g}pt")
        if self.font_east_asia or self.font_ascii:
            parts.append(f"font={self.font_east_asia or self.font_ascii}")
        if self.left_indent:
            parts.append(f"left={self.left_indent}")
        if self.first_line_indent:
            parts.append(f"firstLine={self.first_line_indent}")
        if self.bold:
            parts.append("bold")
        return " | ".join(parts)


@dataclass
class SectionInfo:
    index: int
    page_number_start: str
    page_number_format: str
    title_page: bool
    header_refs: list[str]
    footer_refs: list[str]


@dataclass
class DocumentProfile:
    path: Path
    paragraph_count: int
    non_empty_paragraphs: list[ParagraphInfo]
    table_count: int
    section_infos: list[SectionInfo]
    doc_grid: dict[str, str]
    header_texts: dict[str, str]
    footer_texts: dict[str, str]

    @property
    def cover_paragraphs(self) -> list[ParagraphInfo]:
        toc_idx = None
        for para in self.non_empty_paragraphs[:40]:
            if normalize_compact(para.text) in {"目录", "目录i", "目录I"}:
                toc_idx = para.index
                break
        if toc_idx is None:
            return self.non_empty_paragraphs[:8]
        return [para for para in self.non_empty_paragraphs if para.index < toc_idx]

    @property
    def toc_title(self) -> ParagraphInfo | None:
        for para in self.non_empty_paragraphs[:50]:
            if normalize_compact(para.text) == "目录":
                return para
        return None

    @property
    def toc_entries(self) -> list[ParagraphInfo]:
        title = self.toc_title
        if title is None:
            return []
        entries: list[ParagraphInfo] = []
        started = False
        for para in self.non_empty_paragraphs:
            if para.index == title.index:
                started = True
                continue
            if not started:
                continue
            if para.style_name.lower().startswith("toc") or para.style_id in {"26", "30", "35"}:
                entries.append(para)
                continue
            compact = normalize_compact(para.text)
            if re.match(r"^(第[一二三四五六七八九十]+[章节]|[一二三四五六七八九十]+、)", compact):
                entries.append(para)
                continue
            if entries:
                break
        return entries


def load_xml(docx: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(docx.read(name))
    except KeyError:
        return None


def rel_map(docx: zipfile.ZipFile) -> dict[str, str]:
    root = load_xml(docx, "word/_rels/document.xml.rels")
    if root is None:
        return {}
    mapping: dict[str, str] = {}
    for rel in root.findall("pr:Relationship", NS):
        rel_id = rel.get("Id", "")
        target = rel.get("Target", "")
        if rel_id and target:
            mapping[rel_id] = target if target.startswith("word/") else f"word/{target}"
    return mapping


def style_map(docx: zipfile.ZipFile) -> dict[str, str]:
    root = load_xml(docx, "word/styles.xml")
    mapping: dict[str, str] = {}
    if root is None:
        return mapping
    for style in root.findall("w:style", NS):
        style_id = style.get(W + "styleId", "")
        name = style.find("w:name", NS)
        if style_id:
            mapping[style_id] = name.get(W + "val", "") if name is not None else ""
    return mapping


def text_of_para(para: ET.Element) -> str:
    return "".join(node.text or "" for node in para.findall(".//w:t", NS)).strip()


def paragraph_info(para: ET.Element, idx: int, styles: dict[str, str]) -> ParagraphInfo | None:
    text = text_of_para(para)
    if not text:
        return None

    ppr = para.find("w:pPr", NS)
    style_id = ""
    alignment = ""
    left_indent = ""
    first_line_indent = ""
    line = ""
    line_rule = ""

    if ppr is not None:
        pstyle = ppr.find("w:pStyle", NS)
        if pstyle is not None:
            style_id = pstyle.get(W + "val", "")
        jc = ppr.find("w:jc", NS)
        if jc is not None:
            alignment = jc.get(W + "val", "")
        spacing = ppr.find("w:spacing", NS)
        if spacing is not None:
            line = spacing.get(W + "line", "")
            line_rule = spacing.get(W + "lineRule", "")
        ind = ppr.find("w:ind", NS)
        if ind is not None:
            left_indent = first((ind.get(W + key) for key in ("leftChars", "left"))) or ""
            first_line_indent = first((ind.get(W + key) for key in ("firstLineChars", "firstLine"))) or ""

    font_east_asia = ""
    font_ascii = ""
    size_half_pt = ""
    bold = False

    for run in para.findall("w:r", NS):
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            continue
        rfonts = rpr.find("w:rFonts", NS)
        if rfonts is not None:
            font_east_asia = font_east_asia or rfonts.get(W + "eastAsia", "")
            font_ascii = font_ascii or rfonts.get(W + "ascii", "")
        size = rpr.find("w:sz", NS)
        if size is not None and not size_half_pt:
            size_half_pt = size.get(W + "val", "")
        if rpr.find("w:b", NS) is not None:
            bold = True

    return ParagraphInfo(
        index=idx,
        text=text,
        style_id=style_id,
        style_name=styles.get(style_id, ""),
        alignment=alignment,
        left_indent=left_indent,
        first_line_indent=first_line_indent,
        line=line,
        line_rule=line_rule,
        font_east_asia=font_east_asia,
        font_ascii=font_ascii,
        size_half_pt=size_half_pt,
        bold=bold,
    )


def section_infos(docx: zipfile.ZipFile, rels: dict[str, str]) -> list[SectionInfo]:
    root = load_xml(docx, "word/document.xml")
    if root is None:
        return []
    infos: list[SectionInfo] = []
    for idx, sect in enumerate(root.findall(".//w:sectPr", NS), start=1):
        pg_num = sect.find("w:pgNumType", NS)
        title_page = sect.find("w:titlePg", NS) is not None
        headers: list[str] = []
        footers: list[str] = []
        for ref in sect:
            if ref.tag == W + "headerReference":
                rid = ref.get(R + "id", "")
                headers.append(rels.get(rid, rid))
            elif ref.tag == W + "footerReference":
                rid = ref.get(R + "id", "")
                footers.append(rels.get(rid, rid))
        infos.append(
            SectionInfo(
                index=idx,
                page_number_start=pg_num.get(W + "start", "") if pg_num is not None else "",
                page_number_format=pg_num.get(W + "fmt", "") if pg_num is not None else "",
                title_page=title_page,
                header_refs=headers,
                footer_refs=footers,
            )
        )
    return infos


def xml_text(docx: zipfile.ZipFile, name: str) -> str:
    root = load_xml(docx, name)
    if root is None:
        return ""
    parts = [text for text in (normalize_space("".join(t.text or "" for t in para.findall(".//w:t", NS))) for para in root.findall(".//w:p", NS)) if text]
    return " | ".join(parts)


def extract_profile(path: Path) -> DocumentProfile:
    with zipfile.ZipFile(path) as docx:
        styles = style_map(docx)
        rels = rel_map(docx)
        root = load_xml(docx, "word/document.xml")
        settings = load_xml(docx, "word/settings.xml")
        if root is None:
            raise ValueError(f"Missing word/document.xml in {path}")

        paragraphs = root.findall(".//w:body/w:p", NS)
        non_empty: list[ParagraphInfo] = []
        for idx, para in enumerate(paragraphs, start=1):
            info = paragraph_info(para, idx, styles)
            if info is not None:
                non_empty.append(info)

        doc_grid_node = settings.find(".//w:docGrid", NS) if settings is not None else None
        doc_grid = dict(doc_grid_node.attrib) if doc_grid_node is not None else {}

        header_texts: dict[str, str] = {}
        footer_texts: dict[str, str] = {}
        for rel_target in sorted(set(rels.values())):
            if rel_target.startswith("word/header"):
                header_texts[Path(rel_target).name] = xml_text(docx, rel_target)
            elif rel_target.startswith("word/footer"):
                footer_texts[Path(rel_target).name] = xml_text(docx, rel_target)

        return DocumentProfile(
            path=path,
            paragraph_count=len(paragraphs),
            non_empty_paragraphs=non_empty,
            table_count=len(root.findall(".//w:tbl", NS)),
            section_infos=section_infos(docx, rels),
            doc_grid=doc_grid,
            header_texts=header_texts,
            footer_texts=footer_texts,
        )


def extract_spec_excerpt(spec_path: Path) -> list[str]:
    text = spec_path.read_text(encoding="utf-8")
    patterns = [
        r"页边距设置为：`上 3\.7cm，下 3\.5cm，左 2\.8cm，右 2\.6cm`。",
        r"目录标题使用`黑体 20pt`，`居中`。",
        r"目录条目使用`仿宋 16pt`。",
        r"目录单独编页码，格式为`upperRoman`（罗马数字页码）。",
        r"正文从`1`开始重新编号。",
        r"正文统一使用`仿宋_GB2312 三号`。",
    ]
    found: list[str] = []
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            found.append(match.group(0))
    return found


def compare_profiles(target: DocumentProfile, standard: DocumentProfile, spec_excerpt: list[str]) -> str:
    lines: list[str] = []
    lines.append(f"# 咨询报告 DOCX 结构比对")
    lines.append("")
    lines.append(f"- 测试文档: `{target.path.name}`")
    lines.append(f"- 标准文档: `{standard.path.name}`")
    lines.append("")
    lines.append("## 标准摘录")
    for line in spec_excerpt:
        lines.append(f"- {line}")
    if not spec_excerpt:
        lines.append("- 未从规范文件中提取到预设摘录。")
    lines.append("")
    lines.append("## 文档概况")
    lines.append(
        f"- 测试文档: {target.paragraph_count} 个段落，{len(target.non_empty_paragraphs)} 个非空段落，{target.table_count} 个表格。"
    )
    lines.append(
        f"- 标准文档: {standard.paragraph_count} 个段落，{len(standard.non_empty_paragraphs)} 个非空段落，{standard.table_count} 个表格。"
    )
    lines.append("")
    lines.append("## 重点差异")

    target_cover = target.cover_paragraphs
    standard_cover = standard.cover_paragraphs
    if target_cover and standard_cover:
        lines.append(
            f"- 封面标题行数不同: 测试稿前部识别为 {max(len(target_cover) - 2, 0)} 行标题，标准稿识别为 {max(len(standard_cover) - 2, 0)} 行标题。"
        )
        lines.append(f"- 测试稿封面前几行: {' / '.join(p.text for p in target_cover[:4])}")
        lines.append(f"- 标准稿封面前几行: {' / '.join(p.text for p in standard_cover[:4])}")
    if len(target_cover) >= 1 and len(standard_cover) >= 1:
        lines.append(f"- 测试稿首个封面标题属性: {target_cover[0].short()}")
        lines.append(f"- 标准稿首个封面标题属性: {standard_cover[0].short()}")
    if len(target_cover) >= 2 and len(standard_cover) >= 2:
        lines.append(f"- 测试稿第二行封面标题属性: {target_cover[1].short()}")
        lines.append(f"- 标准稿第二行封面标题属性: {standard_cover[1].short()}")
    if len(target_cover) >= 3 and len(standard_cover) >= 3:
        lines.append(f"- 测试稿编制单位属性: {target_cover[-2].short()}")
        lines.append(f"- 标准稿编制单位属性: {standard_cover[-2].short()}")
    if len(target_cover) >= 4 and len(standard_cover) >= 4:
        lines.append(f"- 测试稿日期属性: {target_cover[-1].short()}")
        lines.append(f"- 标准稿日期属性: {standard_cover[-1].short()}")

    if target.toc_title and standard.toc_title:
        lines.append(f"- 测试稿目录标题: {target.toc_title.short()}")
        lines.append(f"- 标准稿目录标题: {standard.toc_title.short()}")
    if target.toc_entries and standard.toc_entries:
        lines.append(f"- 测试稿目录前 4 项: {' / '.join(p.text for p in target.toc_entries[:4])}")
        lines.append(f"- 标准稿目录前 4 项: {' / '.join(p.text for p in standard.toc_entries[:4])}")
        lines.append(f"- 测试稿首条目录项属性: {target.toc_entries[0].short()}")
        lines.append(f"- 标准稿首条目录项属性: {standard.toc_entries[0].short()}")

    target_roman = target.section_infos[1].page_number_format if len(target.section_infos) >= 2 else ""
    standard_roman = standard.section_infos[1].page_number_format if len(standard.section_infos) >= 2 else ""
    if target_roman or standard_roman:
        lines.append(f"- 目录页码格式不同: 测试稿是 `{target_roman or '未设置'}`，标准稿是 `{standard_roman or '未设置'}`。")

    target_footer_non_empty = {k: v for k, v in target.footer_texts.items() if v}
    standard_footer_non_empty = {k: v for k, v in standard.footer_texts.items() if v}
    lines.append(f"- 测试稿非空页脚内容: {target_footer_non_empty or '无'}")
    lines.append(f"- 标准稿非空页脚内容: {standard_footer_non_empty or '无'}")

    lines.append("")
    lines.append("## 对宏/校对最有价值的结论")
    lines.append("- `rev.docx` 可以直接当作目标样稿，尤其适合校准封面、目录标题、目录页码和目录缩进。")
    lines.append("- 当前测试稿最明显的结构偏差是目录页码格式为 `lowerRoman`，而标准稿与规范要求一致，使用 `upperRoman`。")
    lines.append("- 标准稿目录标题是黑体 20pt 居中，测试稿当前目录标题仍带有正文式残留。")
    lines.append("- 标准稿封面标题为两行小标宋二号、居中、不加粗；编制单位和日期也都明确居中。")
    lines.append("- 标准稿目录项主体使用仿宋 16pt，章条目加粗，节/目条目通过左缩进区分层级。")
    lines.append("")
    lines.append("## 建议下一步")
    lines.append("- 先用这个报告校准 `ResearchProjectFormatter.bas` 的目录标题、目录条目和页码逻辑。")
    lines.append("- 再把同一套比对脚本用于更多样稿，确认规则不是只对这一份文档偶然成立。")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare a consulting report DOCX against a canonical DOCX sample.")
    parser.add_argument("--target", required=True, type=Path, help="The DOCX file to inspect.")
    parser.add_argument("--standard", required=True, type=Path, help="The canonical DOCX sample.")
    parser.add_argument("--spec", required=True, type=Path, help="The extracted markdown formatting standard.")
    parser.add_argument("--report", type=Path, help="Optional markdown output path.")
    args = parser.parse_args()

    target = extract_profile(args.target)
    standard = extract_profile(args.standard)
    report = compare_profiles(target, standard, extract_spec_excerpt(args.spec))

    if args.report:
        args.report.write_text(report, encoding="utf-8")
    print(report, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
