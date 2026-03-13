---
name: math-generator-workflow
description: "**WORKFLOW SKILL** — 从配置调整到 Word 文档生成的完整口算题生成流程。用于：修改 `config.yaml` 题库配置、运行 `PrimarySchoolMathematics.py` 生成逻辑、验证生成的 `docx` 结果。适用于：调整数学算式步数、范围及排版布局。"
---

# 数学口算题生成工作流 (Math Generator Workflow)

这个技能涵盖了从配置修改到最终文档生成的完整步骤，旨在确保题目符合特定的教育难度要求，并且排版正确。

## 1. 题目配置分析 (Config Adjustment)
在生成题目之前，必须先检查并根据需要更新 [config.yaml](config.yaml)。

- **步数控制 (`steps`)**: 
  - 1: 基础加减法 (e.g., 5 + 3)
  - 2: 连加连减/混合运算 (e.g., 5 + 3 - 2)
- **数值范围 (`term1_min/max`, `term2_min/max`)**: 确保数字大小适合孩子。
- **运算符池 (`operators`)**: `["+", "-"]` 或 `["+", "-", "*", "/"]`。
- **结果限制 (`result_min/max`)**: 防止出现负数或超出孩子计算范围的大数。

## 2. 代码逻辑验证 (Logic Check)
核心逻辑位于 [PrimarySchoolMathematics.py](PrimarySchoolMathematics.py)。

- **生成算法**: `generate_problem` 方法使用循环重试机制直到满足结果范围。
- **去重机制**: `create_docx` 中实现了页内和全局去重策略。
- **字体与格式**: 文档导出使用 `python-docx`，需确保系统安装了配置中指定的字体（如“黑体”）。

## 3. 文档生成流程 (Execution)
按照以下步骤执行生成：

1. **环境准备**: 确保已安装依赖：`pip install -r requirements.txt`。
2. **执行脚本**: 运行 `python PrimarySchoolMathematics.py`。
3. **日志监控**: 观察控制台输出或 [logs/](logs/) 目录下的日志，确保没有“严重警告”或死循环。

## 4. 质量核对表 (Success Criteria)
- [ ] **唯一性**: 每一页的题目是否通过了去重逻辑。
- [ ] **格式**: 导出的 `.docx` 文件页边距、字体、姓名日期栏是否符合预期。
- [ ] **难度**: 抽查几道生成的题目，确认其数值范围和步数符合配置。

## 常见问题处理
- **生成缓慢/死循环**: 检查 `result_min/max` 是否设置过窄，或者 `steps` 导致的结果概率太低。
- **中文字体不显示**: 检查 `rPr.rFonts.set(qn("w:eastAsia"), font_name)` 逻辑是否正确注入到文档片段中。
