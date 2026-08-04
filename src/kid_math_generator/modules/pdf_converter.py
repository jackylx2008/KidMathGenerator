"""根据当前操作系统调用 Microsoft Word 将指定 DOCX 转换为 PDF。"""

from __future__ import annotations

import io
import platform
import tempfile
from collections.abc import Iterable
from pathlib import Path
from typing import Any

from pypdf import PdfReader, PdfWriter

from logging_config import get_logger

LOGGER = get_logger(__name__)
WORD_PDF_FORMAT = 17
PDF_PAINT_OPERATORS = {
    b"S",
    b"s",
    b"f",
    b"F",
    b"f*",
    b"B",
    b"B*",
    b"b",
    b"b*",
    b"Do",
    b"sh",
    b"INLINE IMAGE",
}
PDF_TEXT_OPERATORS = {b"Tj", b"TJ", b"'", b'"'}


def convert_docx_files(
    docx_files: Iterable[str | Path],
    *,
    output_dir: str | Path,
    delete_source: bool = False,
) -> list[Path]:
    """在 Windows 或 macOS 上转换显式传入的 DOCX 文件。"""
    target_dir = Path(output_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    sources = [Path(path).resolve() for path in docx_files]
    for source in sources:
        if not source.is_file():
            raise FileNotFoundError(source)

    system = platform.system()
    if system == "Windows":
        return _convert_on_windows(sources, target_dir, delete_source)
    if system == "Darwin":
        return _convert_on_macos(sources, target_dir, delete_source)
    raise RuntimeError(f"Microsoft Word PDF 转换不支持当前系统: {system}")


def _convert_on_windows(
    sources: list[Path],
    target_dir: Path,
    delete_source: bool,
) -> list[Path]:
    """通过 comtypes 复用一个 Word 实例完成 Windows 转换。"""
    import comtypes.client

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    generated: list[Path] = []

    try:
        for source in sources:
            target = target_dir / f"{source.stem}.pdf"
            document = None
            try:
                LOGGER.info("使用 Word for Windows 转换 PDF: %s -> %s", source, target)
                document = word.Documents.Open(str(source))
                document.SaveAs(str(target), FileFormat=WORD_PDF_FORMAT)
            finally:
                if document is not None:
                    document.Close(False)

            _record_generated_file(source, target, generated, delete_source)
    finally:
        word.Quit()

    return generated


def _convert_on_macos(
    sources: list[Path],
    target_dir: Path,
    delete_source: bool,
) -> list[Path]:
    """通过 docx2pdf 的 JXA 自动化接口调用 Word for Mac。"""
    from docx2pdf import convert

    generated: list[Path] = []
    for source in sources:
        target = target_dir / f"{source.stem}.pdf"
        LOGGER.info("使用 Word for Mac 转换 PDF: %s -> %s", source, target)
        try:
            convert(str(source), str(target))
        except SystemExit as error:
            raise RuntimeError(f"Word for Mac 转换失败: {source}") from error
        _record_generated_file(source, target, generated, delete_source)
    return generated


def _record_generated_file(
    source: Path,
    target: Path,
    generated: list[Path],
    delete_source: bool,
) -> None:
    """验证转换产物，并按配置处理源文件。"""
    if not target.is_file() or target.stat().st_size == 0:
        raise RuntimeError(f"PDF 未正常生成: {target}")
    remove_blank_pdf_pages(target)
    generated.append(target)
    if delete_source:
        source.unlink()


def remove_blank_pdf_pages(pdf_path: str | Path) -> int:
    """删除没有可见文字、图片、注释或矢量绘制内容的 PDF 页面。"""
    path = Path(pdf_path).resolve()
    reader = PdfReader(io.BytesIO(path.read_bytes()))
    retained_pages = [page for page in reader.pages if not _is_blank_page(page)]
    removed_count = len(reader.pages) - len(retained_pages)
    if removed_count == 0:
        return 0
    if not retained_pages:
        LOGGER.warning("PDF 的所有页面均被识别为空白，已保留原文件: %s", path)
        return 0

    writer = PdfWriter()
    for page in retained_pages:
        writer.add_page(page)
    if reader.metadata:
        writer.add_metadata(
            {
                str(key): str(value)
                for key, value in reader.metadata.items()
                if key is not None and value is not None
            }
        )

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix=f".{path.stem}-",
            suffix=".pdf.tmp",
            dir=path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
            writer.write(temporary_file)
        temporary_path.replace(path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()

    LOGGER.info("已从 PDF 删除 %s 个空白页: %s", removed_count, path)
    return removed_count


def _is_blank_page(page) -> bool:
    if (page.extract_text() or "").strip():
        return False
    if page.get("/Annots"):
        return False

    contents = page.get_contents()
    if contents is None:
        return True
    for operands, operator in contents.operations:
        if operator in PDF_PAINT_OPERATORS:
            return False
        if operator in PDF_TEXT_OPERATORS and _contains_visible_text(operands):
            return False
    return True


def _contains_visible_text(value: Any) -> bool:
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, bytes):
        return bool(value.strip())
    if isinstance(value, (list, tuple)):
        return any(_contains_visible_text(item) for item in value)
    return False
