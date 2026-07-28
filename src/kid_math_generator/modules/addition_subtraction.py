"""加减法口算题生成能力。"""

from __future__ import annotations

import random
from collections.abc import Mapping, Sequence
from typing import Any


class AdditionSubtractionProblemGenerator:
    """根据数值、步数、运算比例和结果范围生成加减法题目。"""

    _ALIASES = {
        "+": "+",
        "加": "+",
        "add": "+",
        "addition": "+",
        "-": "-",
        "减": "-",
        "sub": "-",
        "subtract": "-",
        "subtraction": "-",
    }

    def __init__(
        self,
        settings: Sequence[Mapping[str, Any]],
        *,
        rng: random.Random | None = None,
    ) -> None:
        if not settings:
            raise ValueError("加减法 settings 不能为空")
        self.settings = [dict(setting) for setting in settings]
        self.rng = rng or random.Random()

    def generate(self) -> tuple[str, str]:
        """随机生成一道满足当前配置约束的加减法题目。"""
        setting = self.rng.choice(self.settings)
        steps = max(1, int(setting.get("steps", 1)))
        operand_ranges = [
            (
                int(setting.get("term2_min", 1)),
                int(setting.get("term2_max", 100)),
            ),
            (
                int(setting.get("term3_min", 1)),
                int(setting.get("term3_max", 100)),
            ),
        ]
        first_min = int(setting.get("term1_min", 1))
        first_max = int(setting.get("term1_max", 100))
        result_min = int(setting.get("result_min", 0))
        result_max = int(setting.get("result_max", 1000))
        mid_min = int(setting.get("mid_result_min", 0))
        first_result_min = int(setting.get("first_result_min", mid_min))
        first_result_max = int(setting.get("first_result_max", result_max))

        for _ in range(500):
            first = self.rng.randint(first_min, first_max)
            current = first
            terms = [str(first)]
            first_step_text = ""
            valid = True

            for step_index in range(steps):
                operator = self._choose_operator(setting, step_index + 1)
                operand_min, operand_max = operand_ranges[min(step_index, 1)]

                if step_index == 0 and steps >= 2:
                    step_min = max(first_result_min, mid_min)
                    step_max = first_result_max
                elif step_index < steps - 1:
                    step_min = mid_min
                    step_max = result_max
                else:
                    step_min = result_min
                    step_max = result_max

                operand_result = self._choose_operand(
                    current,
                    operator,
                    operand_min,
                    operand_max,
                    step_min,
                    step_max,
                )
                if operand_result is None:
                    valid = False
                    break

                previous = current
                operand, current = operand_result
                terms.extend((operator, str(operand)))
                if step_index == 0 and steps >= 2:
                    first_step_text = f"({previous}{operator}{operand}={current})"

            if valid and result_min <= current <= result_max:
                compact = "".join(terms)
                problem = " ".join(terms) + " ="
                answer = f"{compact}={current}"
                if first_step_text:
                    answer = f"{answer} {first_step_text}"
                return problem, answer

        raise ValueError("当前加减法配置无法生成满足范围的题目")

    def _choose_operator(
        self,
        setting: Mapping[str, Any],
        step_number: int,
    ) -> str:
        ratios = setting.get(
            f"operator_ratios{step_number}",
            setting.get("operator_ratios"),
        )
        weighted = self._parse_ratios(ratios)
        if weighted:
            operators, weights = weighted
            return self.rng.choices(operators, weights=weights, k=1)[0]

        fallback = setting.get(
            f"operators{step_number}",
            setting.get("operators", ["+", "-"]),
        )
        operators = [
            normalized
            for item in fallback
            if (normalized := self._normalize_operator(item)) is not None
        ]
        if not operators:
            raise ValueError("加减法运算符配置为空")
        return self.rng.choice(operators)

    def _parse_ratios(
        self,
        ratios: Any,
    ) -> tuple[list[str], list[float]] | None:
        if not isinstance(ratios, Mapping):
            return None

        weights: dict[str, float] = {}
        for operator, raw_weight in ratios.items():
            normalized = self._normalize_operator(operator)
            if normalized is None:
                continue
            try:
                weight = float(str(raw_weight).strip().rstrip("%"))
            except (TypeError, ValueError):
                continue
            if weight > 0:
                weights[normalized] = weights.get(normalized, 0.0) + weight

        if not weights:
            return None
        return list(weights), list(weights.values())

    @classmethod
    def _normalize_operator(cls, operator: Any) -> str | None:
        return cls._ALIASES.get(str(operator).strip().lower())

    def _choose_operand(
        self,
        current: int,
        operator: str,
        operand_min: int,
        operand_max: int,
        result_min: int,
        result_max: int,
    ) -> tuple[int, int] | None:
        candidates: list[tuple[int, int]] = []
        for operand in range(operand_min, operand_max + 1):
            result = current + operand if operator == "+" else current - operand
            if result_min <= result <= result_max:
                candidates.append((operand, result))
        return self.rng.choice(candidates) if candidates else None
