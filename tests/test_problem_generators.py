"""题目算法的边界与正确性测试。"""

from __future__ import annotations

import random
import re
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from kid_math_generator.modules.addition_subtraction import (
    AdditionSubtractionProblemGenerator,
)
from kid_math_generator.modules.multiplication import (
    MultiplicationTableProblemGenerator,
)


class AdditionSubtractionTests(unittest.TestCase):
    def test_generated_answers_obey_result_range(self) -> None:
        generator = AdditionSubtractionProblemGenerator(
            [
                {
                    "steps": 1,
                    "term1_min": 1,
                    "term1_max": 20,
                    "term2_min": 1,
                    "term2_max": 20,
                    "operators1": ["+", "-"],
                    "result_min": 0,
                    "result_max": 30,
                }
            ],
            rng=random.Random(7),
        )

        for _ in range(100):
            problem, answer = generator.generate()
            self.assertRegex(problem, r"^\d+ [+-] \d+ =$")
            result = int(answer.rsplit("=", 1)[1])
            self.assertGreaterEqual(result, 0)
            self.assertLessEqual(result, 30)

    def test_impossible_configuration_fails_explicitly(self) -> None:
        generator = AdditionSubtractionProblemGenerator(
            [
                {
                    "steps": 1,
                    "term1_min": 1,
                    "term1_max": 1,
                    "term2_min": 1,
                    "term2_max": 1,
                    "operators1": ["+"],
                    "result_min": 99,
                    "result_max": 99,
                }
            ],
            rng=random.Random(1),
        )
        with self.assertRaisesRegex(ValueError, "无法生成"):
            generator.generate()


class MultiplicationTests(unittest.TestCase):
    def test_default_one_to_six_range_and_answers(self) -> None:
        generator = MultiplicationTableProblemGenerator(rng=random.Random(11))
        self.assertEqual(generator.unique_problem_count, 36)

        for _ in range(100):
            problem, answer = generator.generate()
            match = re.fullmatch(r"(\d+) × (\d+) =", problem)
            self.assertIsNotNone(match)
            left, right = map(int, match.groups())
            self.assertIn(left, range(1, 7))
            self.assertIn(right, range(1, 7))
            self.assertEqual(answer, f"{left}×{right}={left * right}")

    def test_range_is_limited_to_nine(self) -> None:
        with self.assertRaisesRegex(ValueError, "1–9"):
            MultiplicationTableProblemGenerator(1, 10)


if __name__ == "__main__":
    unittest.main()
