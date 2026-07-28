"""旧版加减法入口兼容工具

用途：
  保留原有 python PrimarySchoolMathematics.py 命令，实际转交给新的
  addition_subtraction_quiz.py 入口执行。

配置文件：
  默认读取根目录 config.yaml，也支持 --config-file 指定其他配置。

示例：
  python PrimarySchoolMathematics.py

输出：
  与 addition_subtraction_quiz.py 相同。
"""

from __future__ import annotations

from addition_subtraction_quiz import main


if __name__ == "__main__":
    raise SystemExit(main())
