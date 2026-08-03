"""竖式计算题目的结构化模型与生成能力。"""

from __future__ import annotations

import random
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Any, Literal, TypeAlias

Operation: TypeAlias = Literal[
    "addition",
    "subtraction",
    "multiplication",
    "division",
]


@dataclass(frozen=True, slots=True)
class AdditionWorking:
    """按最高位到最低位记录需要写在对应数位上方的进位。"""

    carries: tuple[int | None, ...]


@dataclass(frozen=True, slots=True)
class SubtractionWorking:
    """按最高位到最低位记录当前数位是否向左侧借位。"""

    borrows: tuple[bool, ...]


@dataclass(frozen=True, slots=True)
class MultiplicationWorking:
    """为后续乘法竖式预留的部分积结构。"""

    partial_products: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class DivisionStep:
    """为后续除法竖式预留的单步计算结构。"""

    partial_dividend: int
    quotient_digit: int
    product: int
    remainder: int


@dataclass(frozen=True, slots=True)
class DivisionWorking:
    """为后续除法竖式预留的商、余数和计算步骤。"""

    quotient: int
    remainder: int
    steps: tuple[DivisionStep, ...]


VerticalWorking: TypeAlias = (
    AdditionWorking
    | SubtractionWorking
    | MultiplicationWorking
    | DivisionWorking
)


@dataclass(frozen=True, slots=True)
class VerticalArithmeticProblem:
    """一道与显示格式解耦的竖式计算题。"""

    operation: Operation
    left: int
    right: int
    result: int
    working: VerticalWorking
    remainder: int | None = None

    @property
    def unique_key(self) -> tuple[Operation, int, int]:
        return self.operation, self.left, self.right

    @property
    def symbol(self) -> str:
        return {
            "addition": "+",
            "subtraction": "−",
            "multiplication": "×",
            "division": "÷",
        }[self.operation]


class VerticalArithmeticProblemGenerator:
    """按带权配置生成加法或减法竖式题。"""

    _OPERATION_ALIASES = {
        "addition": "addition",
        "add": "addition",
        "+": "addition",
        "subtraction": "subtraction",
        "subtract": "subtraction",
        "sub": "subtraction",
        "-": "subtraction",
        "−": "subtraction",
        "multiplication": "multiplication",
        "multiply": "multiplication",
        "mul": "multiplication",
        "*": "multiplication",
        "×": "multiplication",
        "division": "division",
        "divide": "division",
        "div": "division",
        "/": "division",
        "÷": "division",
    }
    _IMPLEMENTED_OPERATIONS = {"addition", "subtraction"}
    _MODES = {"none", "required", "any"}

    def __init__(
        self,
        settings: Sequence[Mapping[str, Any]],
        *,
        rng: random.Random | None = None,
    ) -> None:
        if not settings:
            raise ValueError("竖式计算 settings 不能为空")
        self.settings = [self._validate_setting(dict(setting)) for setting in settings]
        self.weights = [float(setting.get("weight", 1)) for setting in self.settings]
        if not any(weight > 0 for weight in self.weights):
            raise ValueError("竖式计算配置至少需要一个正数 weight")
        self.rng = rng or random.Random()

    def generate(self) -> VerticalArithmeticProblem:
        """从带权配置中选择一种规则并生成满足约束的题目。"""
        setting = self.rng.choices(self.settings, weights=self.weights, k=1)[0]
        operation = setting["operation"]
        if operation == "addition":
            return self._generate_addition(setting)
        if operation == "subtraction":
            return self._generate_subtraction(setting)
        raise ValueError(f"暂未实现竖式运算: {operation}")

    def _generate_addition(
        self,
        setting: Mapping[str, Any],
    ) -> VerticalArithmeticProblem:
        left_range = self._operand_range(setting, "left")
        right_range = self._operand_range(setting, "right")
        for _ in range(2000):
            left = self.rng.randint(*left_range)
            right = self.rng.randint(*right_range)
            result = left + right
            carries = self._addition_carries(left, right, result)
            if not self._result_in_range(result, setting):
                continue
            if not self._working_count_matches(
                setting,
                "carry",
                sum(carry is not None for carry in carries),
            ):
                continue
            return VerticalArithmeticProblem(
                operation="addition",
                left=left,
                right=right,
                result=result,
                working=AdditionWorking(carries),
            )
        raise ValueError("当前加法竖式配置无法生成满足进位和结果范围的题目")

    def _generate_subtraction(
        self,
        setting: Mapping[str, Any],
    ) -> VerticalArithmeticProblem:
        left_range = self._operand_range(setting, "left")
        right_range = self._operand_range(setting, "right")
        for _ in range(2000):
            left = self.rng.randint(*left_range)
            right = self.rng.randint(*right_range)
            if left < right:
                continue
            result = left - right
            borrows = self._subtraction_borrows(left, right)
            if not self._result_in_range(result, setting):
                continue
            if not self._working_count_matches(
                setting,
                "borrow",
                sum(borrows),
            ):
                continue
            return VerticalArithmeticProblem(
                operation="subtraction",
                left=left,
                right=right,
                result=result,
                working=SubtractionWorking(borrows),
            )
        raise ValueError("当前减法竖式配置无法生成满足借位和结果范围的题目")

    def _validate_setting(self, setting: dict[str, Any]) -> dict[str, Any]:
        raw_operation = str(setting.get("operation", "")).strip().lower()
        operation = self._OPERATION_ALIASES.get(raw_operation)
        if operation is None:
            raise ValueError(f"未知竖式运算类型: {raw_operation or '<empty>'}")
        if operation not in self._IMPLEMENTED_OPERATIONS:
            raise ValueError(f"暂未实现竖式运算: {operation}")
        setting["operation"] = operation

        weight = float(setting.get("weight", 1))
        if weight < 0:
            raise ValueError("竖式计算 weight 不能为负数")

        self._operand_range(setting, "left")
        self._operand_range(setting, "right")
        mode_name = "carry" if operation == "addition" else "borrow"
        mode = str(setting.get(mode_name, "any")).strip().lower()
        if mode not in self._MODES:
            raise ValueError(
                f"{mode_name} 必须是 none、required 或 any，实际为: {mode}"
            )
        setting[mode_name] = mode

        if operation == "subtraction" and bool(setting.get("allow_negative", False)):
            raise ValueError("第一阶段减法竖式暂不支持负数结果")
        return setting

    @staticmethod
    def _operand_range(
        setting: Mapping[str, Any],
        name: str,
    ) -> tuple[int, int]:
        minimum = int(setting.get(f"{name}_min", 0))
        maximum = int(setting.get(f"{name}_max", 99))
        if minimum < 0 or maximum < 0:
            raise ValueError("第一阶段竖式计算暂不支持负数操作数")
        if minimum > maximum:
            raise ValueError(f"{name}_min 不能大于 {name}_max")
        return minimum, maximum

    @staticmethod
    def _result_in_range(result: int, setting: Mapping[str, Any]) -> bool:
        minimum = int(setting.get("result_min", 0))
        maximum = int(setting.get("result_max", 10**12))
        return minimum <= result <= maximum

    @classmethod
    def _working_count_matches(
        cls,
        setting: Mapping[str, Any],
        name: str,
        count: int,
    ) -> bool:
        mode = str(setting.get(name, "any"))
        if mode == "none" and count != 0:
            return False
        if mode == "required" and count == 0:
            return False

        minimum = int(setting.get(f"{name}_count_min", 0))
        maximum = int(setting.get(f"{name}_count_max", 10**6))
        return minimum <= count <= maximum

    @staticmethod
    def _addition_carries(
        left: int,
        right: int,
        result: int,
    ) -> tuple[int | None, ...]:
        width = max(len(str(left)), len(str(right)), len(str(result)))
        left_digits = [int(digit) for digit in str(left).zfill(width)]
        right_digits = [int(digit) for digit in str(right).zfill(width)]
        carries: list[int | None] = [None] * width
        carry = 0
        for index in range(width - 1, -1, -1):
            total = left_digits[index] + right_digits[index] + carry
            carry = total // 10
            if carry and index > 0:
                carries[index - 1] = carry
        return tuple(carries)

    @staticmethod
    def _subtraction_borrows(left: int, right: int) -> tuple[bool, ...]:
        width = max(len(str(left)), len(str(right)))
        left_digits = [int(digit) for digit in str(left).zfill(width)]
        right_digits = [int(digit) for digit in str(right).zfill(width)]
        borrows = [False] * width
        borrowed_from_right = 0
        for index in range(width - 1, -1, -1):
            current = left_digits[index] - borrowed_from_right
            if current < right_digits[index]:
                borrows[index] = True
                borrowed_from_right = 1
            else:
                borrowed_from_right = 0
        return tuple(borrows)
