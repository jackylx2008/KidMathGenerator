"""不同操作系统下的 DOCX 转 PDF 分派测试。"""

from __future__ import annotations

import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from kid_math_generator.modules.pdf_converter import convert_docx_files


class PdfConverterTests(unittest.TestCase):
    def test_macos_uses_docx2pdf_and_can_delete_source(self) -> None:
        calls: list[tuple[str, str]] = []
        fake_docx2pdf = types.ModuleType("docx2pdf")

        def fake_convert(source: str, target: str) -> None:
            calls.append((source, target))
            Path(target).write_bytes(b"%PDF-macOS")

        fake_docx2pdf.convert = fake_convert

        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source = directory / "sample.docx"
            source.write_bytes(b"sample docx")
            output_dir = directory / "output"

            with (
                patch(
                    "kid_math_generator.modules.pdf_converter.platform.system",
                    return_value="Darwin",
                ),
                patch.dict(sys.modules, {"docx2pdf": fake_docx2pdf}),
            ):
                generated = convert_docx_files(
                    [source],
                    output_dir=output_dir,
                    delete_source=True,
                )

            target = output_dir.resolve() / "sample.pdf"
            self.assertEqual(generated, [target])
            self.assertEqual(calls, [(str(source.resolve()), str(target))])
            self.assertTrue(target.is_file())
            self.assertFalse(source.exists())

    def test_windows_uses_comtypes_word_automation(self) -> None:
        class FakeDocument:
            def __init__(self) -> None:
                self.closed = False

            def SaveAs(self, target: str, *, FileFormat: int) -> None:
                self.file_format = FileFormat
                Path(target).write_bytes(b"%PDF-Windows")

            def Close(self, save_changes: bool) -> None:
                self.closed = not save_changes

        class FakeDocuments:
            def __init__(self) -> None:
                self.opened: list[str] = []
                self.document = FakeDocument()

            def Open(self, source: str) -> FakeDocument:
                self.opened.append(source)
                return self.document

        class FakeWord:
            def __init__(self) -> None:
                self.Documents = FakeDocuments()
                self.quit_called = False

            def Quit(self) -> None:
                self.quit_called = True

        word = FakeWord()
        fake_client = types.ModuleType("comtypes.client")
        fake_client.CreateObject = lambda _name: word
        fake_comtypes = types.ModuleType("comtypes")
        fake_comtypes.client = fake_client

        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source = directory / "sample.docx"
            source.write_bytes(b"sample docx")
            output_dir = directory / "output"

            with (
                patch(
                    "kid_math_generator.modules.pdf_converter.platform.system",
                    return_value="Windows",
                ),
                patch.dict(
                    sys.modules,
                    {"comtypes": fake_comtypes, "comtypes.client": fake_client},
                ),
            ):
                generated = convert_docx_files([source], output_dir=output_dir)

            target = output_dir.resolve() / "sample.pdf"
            self.assertEqual(generated, [target])
            self.assertEqual(word.Documents.opened, [str(source.resolve())])
            self.assertEqual(word.Documents.document.file_format, 17)
            self.assertTrue(word.Documents.document.closed)
            self.assertTrue(word.quit_called)

    def test_macos_reports_docx2pdf_failure_without_deleting_source(self) -> None:
        fake_docx2pdf = types.ModuleType("docx2pdf")

        def fake_convert(_source: str, _target: str) -> None:
            raise SystemExit(1)

        fake_docx2pdf.convert = fake_convert

        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "sample.docx"
            source.write_bytes(b"sample docx")
            with (
                patch(
                    "kid_math_generator.modules.pdf_converter.platform.system",
                    return_value="Darwin",
                ),
                patch.dict(sys.modules, {"docx2pdf": fake_docx2pdf}),
            ):
                with self.assertRaisesRegex(RuntimeError, "Word for Mac 转换失败"):
                    convert_docx_files(
                        [source],
                        output_dir=Path(temp_dir) / "output",
                        delete_source=True,
                    )

            self.assertTrue(source.is_file())

    def test_unsupported_system_fails_explicitly(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "sample.docx"
            source.write_bytes(b"sample docx")
            with patch(
                "kid_math_generator.modules.pdf_converter.platform.system",
                return_value="Linux",
            ):
                with self.assertRaisesRegex(RuntimeError, "Linux"):
                    convert_docx_files([source], output_dir=Path(temp_dir) / "output")


if __name__ == "__main__":
    unittest.main()
