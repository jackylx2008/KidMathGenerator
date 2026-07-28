"""加减法口算题工作流。"""

from __future__ import annotations

import random

from kid_math_generator.context import AppContext
from kid_math_generator.flows._quiz_flow import QuizFlowResult, run_quiz_flow
from kid_math_generator.modules.addition_subtraction import (
    AdditionSubtractionProblemGenerator,
)


def run(context: AppContext) -> QuizFlowResult:
    """创建加减法题目源，并复用统一文档与 PDF 流程。"""
    seed = context.app_config.get("random_seed")
    rng = random.Random(seed)
    generator = AdditionSubtractionProblemGenerator(
        context.flow_config.get("settings", []),
        rng=rng,
    )
    return run_quiz_flow(context, generator.generate, rng=rng)
