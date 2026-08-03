"""配置加载和文档构建测试。"""

from __future__ import annotations

import random
import sys
import tempfile
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from docx import Document

from kid_math_generator.config_loader import load_config
from kid_math_generator.modules.document_builder import QuizDocumentBuilder
from kid_math_generator.modules.multiplication import (
    MultiplicationTableProblemGenerator,
)


class ConfigTests(unittest.TestCase):
    def test_project_config_contains_all_flows(self) -> None:
        config = load_config(PROJECT_ROOT / "config.yaml")
        self.assertIn("addition_subtraction", config["flows"])
        self.assertIn("vertical_arithmetic", config["flows"])
        multiplication = config["flows"]["multiplication"]
        self.assertEqual(multiplication["factor_min"], 1)
        self.assertEqual(multiplication["factor_max"], 6)


class DocumentBuilderTests(unittest.TestCase):
    def test_builds_question_and_answer_documents(self) -> None:
        config = {
            "pages": 1,
            "count": 6,
            "columns": 2,
            "title": "乘法测试",
            "output_file": "questions.docx",
            "output_file_answer": "answers.docx",
            "orientation": "landscape",
            "font_name": "黑体",
            "font_size": 18,
            "hard_label": False,
        }
        rng = random.Random(3)
        problem_generator = MultiplicationTableProblemGenerator(rng=rng)
        builder = QuizDocumentBuilder(
            config,
            asset_dir=PROJECT_ROOT / "src",
            rng=rng,
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            result = builder.build(
                problem_generator.generate,
                output_dir=temp_dir,
            )
            self.assertTrue(result.question.is_file())
            self.assertTrue(result.answer.is_file())

            question_doc = Document(result.question)
            answer_doc = Document(result.answer)
            question_text = "\n".join(
                cell.text
                for table in question_doc.tables
                for row in table.rows
                for cell in row.cells
            )
            answer_text = "\n".join(
                cell.text
                for table in answer_doc.tables
                for row in table.rows
                for cell in row.cells
            )
            self.assertEqual(question_text.count("="), 6)
            self.assertEqual(answer_text.count("="), 6)
            self.assertIn("×", question_text)


if __name__ == "__main__":
    unittest.main()
