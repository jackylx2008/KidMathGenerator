"""加减法口算题生成工具

用途：
  按配置生成加减法题目卷、答案卷，并可通过 Microsoft Word 转换为 PDF。
  核心算法位于 src/kid_math_generator，当前文件只负责入口接线。

配置文件：
  默认读取根目录 config.yaml；common.env 可提供本机环境变量。
  加减法范围和运算比例位于 flows.addition_subtraction。

可选参数：
  --config-file   指定 YAML 配置文件，默认使用根目录 config.yaml。

示例：
  python addition_subtraction_quiz.py
  python addition_subtraction_quiz.py --config-file config.yaml

输出：
  DOCX 和 PDF 统一写入 output，日志写入 logs。
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
from kid_math_generator.flows.addition_subtraction_flow import run
from logging_config import get_logger, setup_logger


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
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
    context = AppContext(PROJECT_ROOT, config, "addition_subtraction", logger)
    result = run(context)
    logger.info("加减法口算题生成完成：%s", result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
