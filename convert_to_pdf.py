"""指定 DOCX 文件转换工具

用途：
  使用本机 Microsoft Word 把一个或多个明确指定的 DOCX 转换为 PDF，
  不再扫描并误处理当前目录中的所有文档。

必填参数：
  files   一个或多个 DOCX 文件路径。

可选参数：
  --output-dir   PDF 输出目录，默认 output。
  --delete-source   转换成功后删除源 DOCX。

示例：
  python convert_to_pdf.py output/sample.docx

输出：
  PDF 写入 --output-dir 指定目录，日志写入 log。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from kid_math_generator.modules.pdf_converter import convert_docx_files
from logging_config import get_logger, setup_logger


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("files", nargs="+", type=Path, help="待转换的 DOCX")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=PROJECT_ROOT / "output",
    )
    parser.add_argument("--delete-source", action="store_true")
    args = parser.parse_args(argv)

    setup_logger()
    logger = get_logger(__name__)
    outputs = convert_docx_files(
        args.files,
        output_dir=args.output_dir,
        delete_source=args.delete_source,
    )
    logger.info("已生成 PDF：%s", ", ".join(map(str, outputs)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
