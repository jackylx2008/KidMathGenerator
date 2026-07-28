"""使用 Microsoft Word 将指定 DOCX 转换为 PDF。"""

from __future__ import annotations

import platform
from collections.abc import Iterable
from pathlib import Path

from logging_config import get_logger

LOGGER = get_logger(__name__)
WORD_PDF_FORMAT = 17


def convert_docx_files(
    docx_files: Iterable[str | Path],
    *,
    output_dir: str | Path,
    delete_source: bool = False,
) -> list[Path]:
    """仅转换显式传入的 DOCX 文件，避免误处理当前目录其他文档。"""
    if platform.system() != "Windows":
        raise RuntimeError("Microsoft Word PDF 转换仅支持 Windows")

    import comtypes.client

    target_dir = Path(output_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    sources = [Path(path).resolve() for path in docx_files]
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    generated: list[Path] = []

    try:
        for source in sources:
            if not source.is_file():
                raise FileNotFoundError(source)
            target = target_dir / f"{source.stem}.pdf"
            document = None
            try:
                LOGGER.info("转换 PDF: %s -> %s", source, target)
                document = word.Documents.Open(str(source))
                document.SaveAs(str(target), FileFormat=WORD_PDF_FORMAT)
            finally:
                if document is not None:
                    document.Close(False)

            if not target.is_file() or target.stat().st_size == 0:
                raise RuntimeError(f"PDF 未正常生成: {target}")
            generated.append(target)
            if delete_source:
                source.unlink()
    finally:
        word.Quit()

    return generated
