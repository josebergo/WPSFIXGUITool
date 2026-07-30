from __future__ import annotations

import copy
import hashlib
import io
import json
import os
import posixpath
import re
import shutil
import subprocess
import zipfile
from dataclasses import asdict, dataclass, field
from pathlib import Path, PurePosixPath
from typing import Callable

from lxml import etree
from PIL import Image


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W, "r": R, "a": A, "pr": PR, "ct": CT}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
IMAGE_REL_SUFFIX = "/image"
PAGE_LABELS = ("页次", "页码", "page")
ProgressCallback = Callable[[int, str], None]


def _noop_progress(_: int, __: str) -> None:
    pass


def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


@dataclass
class ConversionReport:
    source: str
    output: str | None = None
    report_file: str | None = None
    source_sha256: str = ""
    source_size: int = 0
    page_fields_found: int = 0
    numpages_fields_found: int = 0
    malformed_page_fields_found: int = 0
    page_fields_rebuilt: int = 0
    include_picture_fields_found: int = 0
    include_picture_fields_unwrapped: int = 0
    linked_images_found: int = 0
    tiff_images_converted: int = 0
    bmp_images_converted: int = 0
    remaining_include_picture_fields: int = 0
    remaining_external_image_refs: int = 0
    embedded_media: int = 0
    zip_crc_ok: bool = False
    xml_parse_ok: bool = False
    image_relationships_ok: bool = False
    fields_balanced: bool = False
    word_fields_updated: bool = False
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    def save(self, path: Path) -> None:
        path.write_text(json.dumps(self.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        self.report_file = str(path)


class DocxPackage:
    def __init__(self, path: Path):
        self.path = path
        if not zipfile.is_zipfile(path):
            raise ValueError("所选文件不是有效的 DOCX/ZIP 文件。")
        with zipfile.ZipFile(path, "r") as source:
            bad = source.testzip()
            if bad:
                raise ValueError(f"DOCX 压缩包损坏：{bad}")
            self.files = {item.filename: source.read(item.filename) for item in source.infolist()}
            self.infos = {item.filename: copy.copy(item) for item in source.infolist()}
        required = {"[Content_Types].xml", "word/document.xml"}
        missing = required.difference(self.files)
        if missing:
            raise ValueError(f"DOCX 缺少必要内容：{', '.join(sorted(missing))}")

    def write(self, output: Path) -> None:
        temp = output.with_name(output.name + ".tmp")
        if temp.exists():
            temp.unlink()
        output.parent.mkdir(parents=True, exist_ok=True)
        try:
            with zipfile.ZipFile(temp, "w") as target:
                for name, data in self.files.items():
                    info = self.infos.get(name)
                    if info is None:
                        info = zipfile.ZipInfo(name)
                        info.compress_type = zipfile.ZIP_DEFLATED
                    target.writestr(info, data)
            with zipfile.ZipFile(temp, "r") as check:
                bad = check.testzip()
                if bad:
                    raise ValueError(f"输出 DOCX 压缩包损坏：{bad}")
            os.replace(temp, output)
        finally:
            if temp.exists():
                temp.unlink()


def _parse(data: bytes) -> etree._Element:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)
    return etree.fromstring(data, parser=parser)


def _serialize(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def unique_output_path(source: Path) -> Path:
    source = source.resolve()
    base = source.with_name(source.stem + "-WPS兼容版" + source.suffix)
    if not base.exists():
        return base
    counter = 2
    while True:
        candidate = source.with_name(source.stem + f"-WPS兼容版-{counter}" + source.suffix)
        if not candidate.exists():
            return candidate
        counter += 1


def _xml_parts(package: DocxPackage):
    for name, data in package.files.items():
        if name.endswith(".xml") or name.endswith(".rels"):
            yield name, data


def _inventory(package: DocxPackage, report: ConversionReport) -> None:
    instructions: list[str] = []
    begins = ends = 0
    linked = 0
    for name, data in _xml_parts(package):
        if not name.startswith("word/") or name.endswith(".rels"):
            continue
        root = _parse(data)
        instructions.extend((node.text or "").strip() for node in root.xpath(".//w:instrText", namespaces=NS))
        instructions.extend((node.get(qn(W, "instr")) or "").strip() for node in root.xpath(".//w:fldSimple", namespaces=NS))
        begins += len(root.xpath(".//w:fldChar[@w:fldCharType='begin']", namespaces=NS))
        ends += len(root.xpath(".//w:fldChar[@w:fldCharType='end']", namespaces=NS))
        linked += len(root.xpath(".//a:blip[@r:link]", namespaces=NS))
    normalized = [re.sub(r"\s+", " ", value.upper()) for value in instructions]
    report.page_fields_found = sum(bool(re.search(r"(^|\s)PAGE($|\s)", value)) for value in normalized)
    report.numpages_fields_found = sum("NUMPAGES" in value for value in normalized)
    report.malformed_page_fields_found = sum(
        bool(re.search(r"(^|\s)PAGE($|\s|[}/])", value)) and "NUMPAGES" in value
        for value in normalized
    )
    report.include_picture_fields_found = sum("INCLUDEPICTURE" in value for value in normalized)
    report.linked_images_found = linked
    report.fields_balanced = begins == ends


def _make_run(*, text: str | None = None, field_char: str | None = None,
              instruction: str | None = None, dirty: bool = False) -> etree._Element:
    run = etree.Element(qn(W, "r"))
    properties = etree.SubElement(run, qn(W, "rPr"))
    fonts = etree.SubElement(properties, qn(W, "rFonts"))
    fonts.set(qn(W, "ascii"), "Times New Roman")
    fonts.set(qn(W, "hAnsi"), "Times New Roman")
    fonts.set(qn(W, "eastAsia"), "宋体")
    etree.SubElement(properties, qn(W, "sz")).set(qn(W, "val"), "24")
    etree.SubElement(properties, qn(W, "szCs")).set(qn(W, "val"), "24")
    if field_char:
        node = etree.SubElement(run, qn(W, "fldChar"))
        node.set(qn(W, "fldCharType"), field_char)
        if dirty:
            node.set(qn(W, "dirty"), "true")
    elif instruction is not None:
        node = etree.SubElement(run, qn(W, "instrText"))
        node.set(XML_SPACE, "preserve")
        node.text = instruction
    else:
        node = etree.SubElement(run, qn(W, "t"))
        if text and (text.startswith(" ") or text.endswith(" ")):
            node.set(XML_SPACE, "preserve")
        node.text = text or ""
    return run


def _add_field(paragraph: etree._Element, instruction: str) -> None:
    paragraph.append(_make_run(field_char="begin", dirty=True))
    paragraph.append(_make_run(instruction=f" {instruction} "))
    paragraph.append(_make_run(field_char="separate"))
    paragraph.append(_make_run(text="1"))
    paragraph.append(_make_run(field_char="end"))


def _replace_page_cell(cell: etree._Element) -> None:
    cell_properties = cell.find(qn(W, "tcPr"))
    for child in list(cell):
        if child is not cell_properties:
            cell.remove(child)
    paragraph = etree.SubElement(cell, qn(W, "p"))
    properties = etree.SubElement(paragraph, qn(W, "pPr"))
    spacing = etree.SubElement(properties, qn(W, "spacing"))
    spacing.set(qn(W, "before"), "0")
    spacing.set(qn(W, "after"), "0")
    spacing.set(qn(W, "line"), "340")
    spacing.set(qn(W, "lineRule"), "exact")
    etree.SubElement(properties, qn(W, "jc")).set(qn(W, "val"), "center")
    _add_field(paragraph, "PAGE")
    paragraph.append(_make_run(text=" / "))
    _add_field(paragraph, "NUMPAGES")


def _cell_text(cell: etree._Element) -> str:
    return "".join(cell.xpath(".//w:t/text()", namespaces=NS)).strip()


def _has_page_pair(cell: etree._Element) -> bool:
    instructions = [re.sub(r"\s+", " ", value.upper()).strip()
                    for value in cell.xpath(".//w:instrText/text()", namespaces=NS)]
    has_page = any(bool(re.search(r"(^|\s)PAGE($|\s)", value)) and "NUMPAGES" not in value
                   for value in instructions)
    has_numpages = any("NUMPAGES" in value and "PAGE }" not in value for value in instructions)
    return has_page and has_numpages


def _repair_page_fields(package: DocxPackage, report: ConversionReport) -> None:
    parts = [name for name in package.files if re.fullmatch(r"word/(header|footer)\d+\.xml", name)]
    for name in parts:
        root = _parse(package.files[name])
        candidates: dict[int, etree._Element] = {}

        for instruction in root.xpath(".//w:instrText", namespaces=NS):
            value = (instruction.text or "").upper()
            if "PAGE" in value and "NUMPAGES" in value:
                cells = instruction.xpath("ancestor::w:tc[1]", namespaces=NS)
                if cells:
                    candidates[id(cells[0])] = cells[0]

        for label_cell in root.xpath(".//w:tc", namespaces=NS):
            label = re.sub(r"\s+", "", _cell_text(label_cell)).lower()
            if not any(token in label for token in PAGE_LABELS):
                continue
            row = label_cell.getparent()
            cells = row.findall(qn(W, "tc")) if row is not None else []
            try:
                index = cells.index(label_cell)
            except ValueError:
                continue
            if index + 1 >= len(cells):
                report.warnings.append(f"{name}：页码标签右侧没有数值单元格。")
                continue
            target = cells[index + 1]
            if _has_page_pair(target):
                continue
            target_text = re.sub(r"\s+", "", _cell_text(target))
            has_page_code = any("PAGE" in (node.text or "").upper()
                                for node in target.xpath(".//w:instrText", namespaces=NS))
            has_shape = bool(target.xpath(".//w:drawing|.//w:pict", namespaces=NS))
            safe_literal = not target_text or bool(re.fullmatch(r"\d{1,4}([/／-]\d{1,4})?", target_text))
            if has_page_code or has_shape or safe_literal:
                candidates[id(target)] = target
            else:
                report.warnings.append(f"{name}：疑似页码单元格含业务文字，已跳过。")

        changed = False
        for cell in candidates.values():
            if not _has_page_pair(cell):
                _replace_page_cell(cell)
                report.page_fields_rebuilt += 1
                changed = True
        if changed:
            package.files[name] = _serialize(root)


def _unwrap_include_pictures(package: DocxPackage, report: ConversionReport) -> None:
    for name in list(package.files):
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        root = _parse(package.files[name])
        changed = False
        for paragraph in root.xpath(".//w:p", namespaces=NS):
            search_from = 0
            while True:
                children = list(paragraph)
                instruction_index = next(
                    (i for i in range(search_from, len(children))
                     if any("INCLUDEPICTURE" in (node.text or "").upper()
                            for node in children[i].xpath(".//w:instrText", namespaces=NS))),
                    None,
                )
                if instruction_index is None:
                    break
                begin_index = next((i for i in range(instruction_index - 1, -1, -1)
                                    if children[i].xpath(".//w:fldChar[@w:fldCharType='begin']", namespaces=NS)), None)
                separate_index = next((i for i in range(instruction_index + 1, len(children))
                                       if children[i].xpath(".//w:fldChar[@w:fldCharType='separate']", namespaces=NS)), None)
                start = separate_index + 1 if separate_index is not None else instruction_index + 1
                end_index = next((i for i in range(start, len(children))
                                  if children[i].xpath(".//w:fldChar[@w:fldCharType='end']", namespaces=NS)), None)
                if begin_index is None or separate_index is None or end_index is None:
                    report.warnings.append(f"{name}：发现结构不完整的 INCLUDEPICTURE 字段。")
                    search_from = instruction_index + 1
                    continue
                result_nodes = children[separate_index + 1:end_index]
                has_cached_image = any(node.xpath(".//w:drawing|.//w:pict", namespaces=NS) for node in result_nodes)
                if not has_cached_image:
                    report.warnings.append(f"{name}：远程图片没有可恢复的缓存图，已保留原字段。")
                    search_from = end_index + 1
                    continue
                for index in list(range(begin_index, separate_index + 1)) + [end_index]:
                    paragraph.remove(children[index])
                report.include_picture_fields_unwrapped += 1
                changed = True
                search_from = begin_index + len(result_nodes)
        if changed:
            package.files[name] = _serialize(root)


def _relationship_source(rels_name: str) -> str:
    path = PurePosixPath(rels_name)
    if rels_name == "_rels/.rels":
        return ""
    return str(path.parent.parent / path.name[:-5])


def _resolve_target(source_part: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _unique_media_name(package: DocxPackage, stem: str) -> str:
    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", stem).strip("._") or "image"
    candidate = f"word/media/{safe_stem}.png"
    counter = 2
    while candidate in package.files:
        candidate = f"word/media/{safe_stem}_{counter}.png"
        counter += 1
    return candidate


def _to_png(data: bytes) -> bytes:
    with Image.open(io.BytesIO(data)) as image:
        output = io.BytesIO()
        mode = "RGBA" if "A" in image.getbands() else "RGB"
        image.convert(mode).save(output, "PNG")
        return output.getvalue()


def _ensure_png_content_type(package: DocxPackage) -> None:
    name = "[Content_Types].xml"
    root = _parse(package.files[name])
    if not root.xpath("./ct:Default[translate(@Extension, 'PNG', 'png')='png']", namespaces=NS):
        node = etree.SubElement(root, qn(CT, "Default"))
        node.set("Extension", "png")
        node.set("ContentType", "image/png")
        package.files[name] = _serialize(root)


def _normalize_images(package: DocxPackage, report: ConversionReport) -> None:
    converted: dict[str, str] = {}
    for rels_name in [name for name in package.files if name.endswith(".rels")]:
        root = _parse(package.files[rels_name])
        source_part = _relationship_source(rels_name)
        changed = False
        for relationship in root.xpath("./pr:Relationship", namespaces=NS):
            if not (relationship.get("Type") or "").endswith(IMAGE_REL_SUFFIX):
                continue
            if relationship.get("TargetMode") == "External":
                continue
            target = relationship.get("Target") or ""
            actual = _resolve_target(source_part, target)
            extension = PurePosixPath(actual).suffix.lower()
            if extension not in {".tif", ".tiff", ".bmp"} or actual not in package.files:
                continue
            if actual not in converted:
                new_actual = _unique_media_name(package, PurePosixPath(actual).stem + "_wps")
                package.files[new_actual] = _to_png(package.files[actual])
                converted[actual] = new_actual
                if extension in {".tif", ".tiff"}:
                    report.tiff_images_converted += 1
                else:
                    report.bmp_images_converted += 1
            relationship.set("Target", posixpath.relpath(converted[actual], posixpath.dirname(source_part)))
            changed = True
        if changed:
            package.files[rels_name] = _serialize(root)
    if converted:
        _ensure_png_content_type(package)


def _enable_field_updates(package: DocxPackage) -> None:
    name = "word/settings.xml"
    if name not in package.files:
        return
    root = _parse(package.files[name])
    node = root.find(qn(W, "updateFields"))
    if node is None:
        node = etree.Element(qn(W, "updateFields"))
        root.insert(0, node)
    node.set(qn(W, "val"), "true")
    package.files[name] = _serialize(root)


def _audit(package: DocxPackage, report: ConversionReport) -> None:
    report.zip_crc_ok = True
    report.xml_parse_ok = True
    missing_images: list[str] = []
    external = linked = include = begins = ends = 0
    for name, data in _xml_parts(package):
        try:
            root = _parse(data)
        except etree.XMLSyntaxError as exc:
            report.xml_parse_ok = False
            report.warnings.append(f"{name} XML 无法解析：{exc}")
            continue
        if name.endswith(".rels"):
            source_part = _relationship_source(name)
            for relationship in root.xpath("./pr:Relationship", namespaces=NS):
                if not (relationship.get("Type") or "").endswith(IMAGE_REL_SUFFIX):
                    continue
                if relationship.get("TargetMode") == "External":
                    external += 1
                else:
                    actual = _resolve_target(source_part, relationship.get("Target") or "")
                    if actual not in package.files:
                        missing_images.append(actual)
            continue
        if not name.startswith("word/"):
            continue
        linked += len(root.xpath(".//a:blip[@r:link]", namespaces=NS))
        include += sum("INCLUDEPICTURE" in (node.text or "").upper()
                       for node in root.xpath(".//w:instrText", namespaces=NS))
        begins += len(root.xpath(".//w:fldChar[@w:fldCharType='begin']", namespaces=NS))
        ends += len(root.xpath(".//w:fldChar[@w:fldCharType='end']", namespaces=NS))
    report.remaining_include_picture_fields = include
    report.remaining_external_image_refs = external + linked
    report.image_relationships_ok = not missing_images
    report.fields_balanced = begins == ends
    report.embedded_media = len([name for name in package.files if name.startswith("word/media/") and not name.endswith("/")])
    for image in sorted(set(missing_images)):
        report.warnings.append(f"缺少内嵌图片：{image}")


def _word_update_fields(path: Path) -> bool:
    if os.name != "nt":
        return False
    shell = shutil.which("pwsh.exe") or shutil.which("powershell.exe")
    if not shell:
        return False
    script = r"""
$ErrorActionPreference = 'Stop'
$path = $env:WPSFIX_DOCX_PATH
$word = New-Object -ComObject Word.Application
try {
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($path, $false, $false, $false)
  $doc.Repaginate()
  foreach ($storyType in 1..17) {
    try {
      $range = $doc.StoryRanges.Item($storyType)
      while ($null -ne $range) {
        if ($range.Fields.Count -gt 0) { [void]$range.Fields.Update() }
        $range = $range.NextStoryRange
      }
    } catch {}
  }
  $doc.Repaginate()
  $doc.Save()
  $doc.Close(0)
} finally {
  $word.Quit()
}
"""
    try:
        child_env = os.environ.copy()
        child_env["WPSFIX_DOCX_PATH"] = str(path)
        result = subprocess.run(
            [shell, "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, text=True, timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            env=child_env,
        )
    except (OSError, subprocess.TimeoutExpired):
        return False
    return result.returncode == 0


def inspect_docx(source: Path) -> ConversionReport:
    source = source.resolve()
    report = ConversionReport(
        source=str(source),
        source_sha256=_sha256(source),
        source_size=source.stat().st_size,
    )
    package = DocxPackage(source)
    _inventory(package, report)
    _audit(package, report)
    return report


def convert_docx(source: Path, output: Path | None = None, *,
                 progress: ProgressCallback | None = None,
                 update_with_word: bool = True) -> ConversionReport:
    progress = progress or _noop_progress
    source = source.resolve()
    if source.suffix.lower() != ".docx":
        raise ValueError("请选择 .docx 文件。")
    if not source.is_file():
        raise FileNotFoundError(source)
    output = (output or unique_output_path(source)).resolve()
    if output == source:
        raise ValueError("输出文件不能覆盖原文件。")
    if output.exists():
        raise FileExistsError(output)

    progress(5, "正在验证并读取 DOCX……")
    package = DocxPackage(source)
    report = ConversionReport(
        source=str(source), output=str(output),
        source_sha256=_sha256(source), source_size=source.stat().st_size,
    )

    progress(15, "正在检查页码域、图片和媒体格式……")
    _inventory(package, report)

    progress(30, "正在修复 PAGE / NUMPAGES 页码域……")
    _repair_page_fields(package, report)

    progress(50, "正在将远程图片转换为内嵌图片……")
    _unwrap_include_pictures(package, report)

    progress(65, "正在转换 WPS 高风险图片格式……")
    _normalize_images(package, report)
    _enable_field_updates(package)

    progress(75, "正在重打包并校验 DOCX……")
    _audit(package, report)
    package.write(output)

    progress(85, "正在更新页码字段……")
    if update_with_word:
        report.word_fields_updated = _word_update_fields(output)
        if not report.word_fields_updated:
            report.warnings.append("未调用 Word 更新字段；WPS 打开文档时会自动更新页码。")

    progress(95, "正在执行最终检查……")
    final_package = DocxPackage(output)
    _audit(final_package, report)
    if not all((report.zip_crc_ok, report.xml_parse_ok, report.image_relationships_ok, report.fields_balanced)):
        raise ValueError("输出文件未通过结构校验。")

    report_path = output.with_suffix(".report.json")
    report.report_file = str(report_path)
    report.save(report_path)
    progress(100, "转换完成。")
    return report
