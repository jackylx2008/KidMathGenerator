"""九九乘法口算题生成工具

用途：
  生成指定因数范围内的九九乘法题目卷、答案卷，并可转换为 PDF。
  默认范围为 1–6，即两个因数都从 1 到 6。

配置文件：
  默认读取根目录 config.yaml；common.env 可提供本机环境变量。
  出题范围位于 flows.multiplication.factor_min 和 factor_max。

可选参数：
  --config-file   指定 YAML 配置文件，默认使用根目录 config.yaml。

示例：
  python multiplication_quiz.py
  python multiplication_quiz.py --config-file config.yaml

输出：
  DOCX 写入 output/docx，PDF 写入 output/pdf，日志写入 log。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from kid_math_generator.config_loader import load_config
from kid_math_generator.context import AppContext
from kid_math_generator.flows.multiplication_flow import run
from logging_config import get_logger, setup_logger


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--config-file",
        type=Path,
        default=PROJECT_ROOT / "config.yaml",
        help="YAML 配置文件路径",
    )
    args = parser.parse_args(argv)

    config = load_config(args.config_file)
    setup_logger(config.get("app", {}).get("log_level", "INFO"))
    logger = get_logger(__name__)
    context = AppContext(PROJECT_ROOT, config, "multiplication", logger)
    result = run(context)
    logger.info("九九乘法口算题生成完成：%s", result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
