"""统一日志初始化测试。"""

from __future__ import annotations

import logging
import sys
import tempfile
import unittest
from logging.handlers import RotatingFileHandler
from pathlib import Path
from unittest.mock import patch

import logging_config


class LoggingConfigTests(unittest.TestCase):
    def test_uses_entry_name_and_rotating_file_in_logs_directory(self) -> None:
        root_logger = logging.getLogger()
        previous_handlers = list(root_logger.handlers)
        previous_level = root_logger.level

        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                with (
                    patch.object(logging_config, "PROJECT_ROOT", Path(temp_dir)),
                    patch.object(sys, "argv", ["sample_entry.py"]),
                ):
                    configured = logging_config.setup_logger("DEBUG")
                    configured.debug("sample message")

                rotating_handlers = [
                    handler
                    for handler in configured.handlers
                    if isinstance(handler, RotatingFileHandler)
                ]
                self.assertEqual(len(rotating_handlers), 1)
                self.assertEqual(rotating_handlers[0].maxBytes, 10 * 1024 * 1024)
                self.assertEqual(rotating_handlers[0].backupCount, 5)

                log_path = Path(temp_dir) / "logs" / "sample_entry.log"
                self.assertTrue(log_path.is_file())
                self.assertIn("sample message", log_path.read_text(encoding="utf-8"))
            finally:
                for handler in list(root_logger.handlers):
                    root_logger.removeHandler(handler)
                    handler.close()
                root_logger.handlers.extend(previous_handlers)
                root_logger.setLevel(previous_level)


if __name__ == "__main__":
    unittest.main()
