---
name: math-generator-workflow
description: "**WORKFLOW SKILL** — 从配置调整、题目生成到 Word/PDF 逐页校验的口算题完整流程。适用于加减法或 1-9 范围内的九九乘法。"
---

# 数学口算题生成工作流

## 1. 选择入口和配置

- 加减法：`addition_subtraction_quiz.py`
- 九九乘法：`multiplication_quiz.py`
- 旧版兼容入口：`PrimarySchoolMathematics.py`

统一检查根目录 `config.yaml`：

- 公共项位于 `app`。
- 加减法配置位于 `flows.addition_subtraction`。
- 九九乘法配置位于 `flows.multiplication`。
- 乘法的 `factor_min` 和 `factor_max` 同时约束两个因数，合法范围为 1-9。

## 2. 代码职责

- `src/kid_math_generator/modules/addition_subtraction.py`：加减法算法。
- `src/kid_math_generator/modules/multiplication.py`：九九乘法算法。
- `src/kid_math_generator/modules/document_builder.py`：DOCX 排版和图章。
- `src/kid_math_generator/modules/pdf_converter.py`：指定文件的 Word PDF 转换。
- `src/kid_math_generator/flows/`：场景编排。

入口脚本中不得重新实现题目算法、文档排版或 PDF 转换。

## 3. 自动化检查

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
flake8 .
```

必须验证：

- 加减法结果满足配置范围。
- 乘法两个因数都满足 `factor_min` / `factor_max`。
- 题目卷与答案卷题量一致。
- 答案重新计算正确。

## 4. 完整烟雾测试

```powershell
python addition_subtraction_quiz.py --config-file tests/fixtures/smoke_config.yaml
python multiplication_quiz.py --config-file tests/fixtures/smoke_config.yaml
python tests/validate_smoke_pdfs.py
```

测试输出位于 `tmp/smoke_output`，不进入 Git。

## 5. 视觉验收

最终 PDF 必须逐页渲染为 PNG 并检查：

- 页面数量与配置一致，没有额外空白页。
- 标题、信息行、题目和答案没有裁切或重叠。
- 图章只出现在题目卷，且不遮挡题目。
- 数学符号与中文字体显示正常。
- 每页题量与列数符合配置。

未完成逐页视觉检查时，不得声明 PDF 验收通过。
