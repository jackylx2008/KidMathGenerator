"""配置文件、环境变量与跨平台路径解析测试。"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from kid_math_generator.config_loader import (
    get_cloudstation_root,
    load_config,
)


class ConfigLoaderTests(unittest.TestCase):
    def test_loads_common_env_without_overwriting_process_environment(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            env_path = directory / "common.env"
            env_path.write_text(
                "EXISTING_VALUE=from-file\n"
                "CLOUDSTATION_ROOT_MACOS=~/SynologyDrive/\n",
                encoding="utf-8",
            )
            config_path = directory / "config.yaml"
            config_path.write_text(
                "existing: '${EXISTING_VALUE}'\n"
                "fallback: '${MISSING_VALUE:-sample}'\n"
                "cloud_path: '${CLOUDSTATION_ROOT}/demo/input.xlsx'\n",
                encoding="utf-8",
            )

            with (
                patch.dict(os.environ, {"EXISTING_VALUE": "from-process"}, clear=True),
                patch(
                    "kid_math_generator.config_loader.platform.system",
                    return_value="Darwin",
                ),
            ):
                config = load_config(config_path, env_path=env_path)

            self.assertEqual(config["existing"], "from-process")
            self.assertEqual(config["fallback"], "sample")
            self.assertEqual(
                config["cloud_path"],
                str(Path.home() / "SynologyDrive" / "demo" / "input.xlsx"),
            )

    def test_explicit_cloudstation_root_has_highest_priority(self) -> None:
        environment = {
            "CLOUDSTATION_ROOT": "~/explicit-root",
            "CLOUDSTATION_ROOT_LINUX": "~/platform-root",
        }
        with (
            patch.dict(os.environ, environment, clear=True),
            patch(
                "kid_math_generator.config_loader.platform.system",
                return_value="Linux",
            ),
        ):
            root = get_cloudstation_root()

        self.assertEqual(root, str(Path.home() / "explicit-root"))


if __name__ == "__main__":
    unittest.main()
