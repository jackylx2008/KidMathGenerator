"""验证烟雾测试 PDF 的页数、题量、范围和答案。"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from pypdf import PdfReader

EXPECTED_PDF_COUNT = 4
EXPECTED_PAGES = 2
EXPECTED_PROBLEMS = 40


def read_pdf(path: Path) -> tuple[int, str]:
    reader = PdfReader(path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    return len(reader.pages), text


def validate(output_dir: Path) -> None:
    pdfs = sorted(output_dir.glob("*.pdf"))
    if len(pdfs) != EXPECTED_PDF_COUNT:
        raise AssertionError(f"应生成 {EXPECTED_PDF_COUNT} 个 PDF，实际为 {len(pdfs)}")

    documents: dict[str, str] = {}
    for path in pdfs:
        if path.stat().st_size == 0:
            raise AssertionError(f"PDF 为空: {path}")
        pages, text = read_pdf(path)
        if pages != EXPECTED_PAGES:
            raise AssertionError(f"{path.name} 应为 {EXPECTED_PAGES} 页，实际为 {pages}")
        documents[path.name] = text

    addition_answer = next(
        text
        for name, text in documents.items()
        if name.startswith("addition") and "answer" in name
    )
    addition_equations = re.findall(r"(\d+)([+-])(\d+)=(\d+)", addition_answer)
    if len(addition_equations) != EXPECTED_PROBLEMS:
        raise AssertionError(f"加减法答案数量异常: {len(addition_equations)}")
    for left, operator, right, answer in addition_equations:
        expected = int(left) + int(right) if operator == "+" else int(left) - int(right)
        if expected != int(answer):
            raise AssertionError(f"加减法答案错误: {left}{operator}{right}={answer}")

    multiplication_answer = next(
        text
        for name, text in documents.items()
        if name.startswith("multiplication") and "answer" in name
    )
    multiplication_equations = re.findall(
        r"(\d+)\s*[^\d=\s]\s*(\d+)=(\d+)",
        multiplication_answer,
    )
    if len(multiplication_equations) != EXPECTED_PROBLEMS:
        raise AssertionError(f"乘法答案数量异常: {len(multiplication_equations)}")
    for left, right, answer in multiplication_equations:
        factors = int(left), int(right)
        if not all(1 <= factor <= 6 for factor in factors):
            raise AssertionError(f"乘法因数超出 1–6: {left} × {right}")
        if int(left) * int(right) != int(answer):
            raise AssertionError(f"乘法答案错误: {left} × {right} = {answer}")

    for name, text in documents.items():
        if "answer" not in name and text.count("=") != EXPECTED_PROBLEMS:
            raise AssertionError(f"{name} 题量异常: {text.count('=')}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("tmp/smoke_output/pdf"),
    )
    args = parser.parse_args()
    validate(args.output_dir)
    print(f"PDF 校验通过: {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
