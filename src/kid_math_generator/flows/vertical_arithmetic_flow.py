"""加减法竖式计算工作流。"""

from __future__ import annotations

import random

from kid_math_generator.context import AppContext
from kid_math_generator.flows._quiz_flow import QuizFlowResult, finish_quiz_flow
from kid_math_generator.modules.vertical_arithmetic import (
    VerticalArithmeticProblemGenerator,
)
from kid_math_generator.modules.vertical_document_builder import (
    VerticalArithmeticDocumentBuilder,
)


def run(context: AppContext) -> QuizFlowResult:
    """按配置生成加减法竖式，并复用统一 PDF 收尾流程。"""
    seed = context.app_config.get("random_seed")
    rng = random.Random(seed)
    generator = VerticalArithmeticProblemGenerator(
        context.flow_config.get("settings", []),
        rng=rng,
    )
    builder = VerticalArithmeticDocumentBuilder(
        context.flow_config,
        asset_dir=context.project_root / "src",
        rng=rng,
    )
    pair = builder.build(generator.generate, output_dir=context.output_dir)
    return finish_quiz_flow(context, pair)
