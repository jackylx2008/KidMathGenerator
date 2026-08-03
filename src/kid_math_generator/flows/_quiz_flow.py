"""各类练习工作流共享的文档与 PDF 编排步骤。"""

from __future__ import annotations

import random
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from kid_math_generator.context import AppContext
from kid_math_generator.modules.document_builder import DocumentPair, QuizDocumentBuilder
from kid_math_generator.modules.pdf_converter import convert_docx_files


@dataclass(frozen=True, slots=True)
class QuizFlowResult:
    docx_files: tuple[Path, ...]
    pdf_files: tuple[Path, ...]


def run_quiz_flow(
    context: AppContext,
    problem_factory: Callable[[], tuple[str, str]],
    *,
    rng: random.Random,
) -> QuizFlowResult:
    """按生成 DOCX、转换 PDF、汇总结果三个阶段执行。"""
    flow_config = context.flow_config
    builder = QuizDocumentBuilder(
        flow_config,
        asset_dir=context.project_root / "src",
        rng=rng,
    )
    pair = builder.build(
        problem_factory,
        output_dir=context.output_dir,
    )

    return finish_quiz_flow(context, pair)


def finish_quiz_flow(
    context: AppContext,
    pair: DocumentPair,
) -> QuizFlowResult:
    """复用 DOCX 转 PDF 和结构化结果汇总步骤。"""

    pdf_files: tuple[Path, ...] = ()
    if bool(context.app_config.get("convert_to_pdf", True)):
        converted = convert_docx_files(
            (pair.question, pair.answer),
            output_dir=context.output_dir,
            delete_source=bool(
                context.app_config.get("delete_docx_after_pdf", True)
            ),
        )
        pdf_files = tuple(converted)

    return QuizFlowResult(
        docx_files=tuple(
            path for path in (pair.question, pair.answer) if path.exists()
        ),
        pdf_files=pdf_files,
    )
