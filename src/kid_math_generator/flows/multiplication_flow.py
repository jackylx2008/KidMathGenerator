"""九九乘法口算题工作流。"""

from __future__ import annotations

import random

from kid_math_generator.context import AppContext
from kid_math_generator.flows._quiz_flow import QuizFlowResult, run_quiz_flow
from kid_math_generator.modules.multiplication import (
    MultiplicationTableProblemGenerator,
)


def run(context: AppContext) -> QuizFlowResult:
    """按 flows.multiplication 的因数范围生成乘法题。"""
    config = context.flow_config
    seed = context.app_config.get("random_seed")
    rng = random.Random(seed)
    generator = MultiplicationTableProblemGenerator(
        int(config.get("factor_min", 1)),
        int(config.get("factor_max", 6)),
        rng=rng,
    )
    return run_quiz_flow(context, generator.generate, rng=rng)
