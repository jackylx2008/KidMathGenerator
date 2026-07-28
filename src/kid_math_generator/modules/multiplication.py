"""九九乘法口算题生成能力。"""

from __future__ import annotations

import random


class MultiplicationTableProblemGenerator:
    """在指定因数范围内生成九九乘法题。"""

    def __init__(
        self,
        factor_min: int = 1,
        factor_max: int = 6,
        *,
        rng: random.Random | None = None,
    ) -> None:
        if factor_min < 1 or factor_max > 9:
            raise ValueError("九九乘法因数范围必须位于 1–9")
        if factor_min > factor_max:
            raise ValueError("factor_min 不能大于 factor_max")
        self.factor_min = factor_min
        self.factor_max = factor_max
        self.rng = rng or random.Random()

    @property
    def unique_problem_count(self) -> int:
        width = self.factor_max - self.factor_min + 1
        return width * width

    def generate(self) -> tuple[str, str]:
        """生成一道两个因数均位于配置范围内的乘法题。"""
        left = self.rng.randint(self.factor_min, self.factor_max)
        right = self.rng.randint(self.factor_min, self.factor_max)
        return f"{left} × {right} =", f"{left}×{right}={left * right}"
