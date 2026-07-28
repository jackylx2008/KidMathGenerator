# Kid Math Generator

用于生成小学生加减法和九九乘法口算题。程序会分别创建题目卷、答案卷 DOCX，并在 Windows 上调用 Microsoft Word 转换为 PDF。

## 运行环境

- Python 3.12+
- Windows 与 Microsoft Word（仅 PDF 转换需要）

安装依赖：

```powershell
python -m pip install -r requirements.txt
```

## 运行命令

生成加减法题：

```powershell
python addition_subtraction_quiz.py
```

生成九九乘法题：

```powershell
python multiplication_quiz.py
```

旧命令仍兼容，会转交给新的加减法入口：

```powershell
python PrimarySchoolMathematics.py
```

可以指定其他配置文件：

```powershell
python multiplication_quiz.py --config-file tests/fixtures/smoke_config.yaml
```

## 配置

统一配置位于根目录 `config.yaml`：

- `app`
  - `log_level`：日志级别。
  - `output_dir`：输出根目录，默认 `output`。
  - `convert_to_pdf`：是否调用 Word 转换 PDF。
  - `delete_docx_after_pdf`：PDF 成功后是否删除 DOCX，默认保留。
  - `random_seed`：可选随机种子，设置整数后可以复现同一批题。
- `flows.addition_subtraction`
  - 页数、每页题量、排版、图章和加减法数值范围。
  - `settings` 支持一步或多步运算、运算符比例及结果范围。
- `flows.multiplication`
  - `factor_min` / `factor_max` 控制两个因数的范围。
  - 默认值为 `1` 和 `6`，因此只生成两个因数均在 1-6 内的乘法题。

本机差异可写入不入库的 `common.env`，格式参考 `common.env.example`。

## 输出

- `output/docx/`：题目卷和答案卷 DOCX。
- `output/pdf/`：题目卷和答案卷 PDF。
- `log/`：按入口脚本命名的滚动日志。

开启加减法 `hard_label` 后，程序读取 `src` 根目录中的第一张图片作为图章，只盖在题目卷上；输出文件名会追加 `_难题`。

## 项目结构

```text
src/kid_math_generator/
├── config_loader.py
├── context.py
├── modules/
│   ├── addition_subtraction.py
│   ├── multiplication.py
│   ├── document_builder.py
│   └── pdf_converter.py
└── flows/
    ├── addition_subtraction_flow.py
    ├── multiplication_flow.py
    └── _quiz_flow.py
```

- `modules/` 提供单一、可复用的题目生成、文档排版和 PDF 转换能力。
- `flows/` 组织具体场景的执行步骤。
- 根目录入口脚本只负责配置、日志、上下文和工作流调用。
- 根目录 `logging_config.py` 统一配置控制台与滚动文件日志。

## 测试和 PDF 校验

运行单元测试和代码规范检查：

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
flake8 .
```

运行可复现的完整烟雾测试，会各生成 2 页、每页 20 题的加减法和乘法题目卷/答案卷：

```powershell
python addition_subtraction_quiz.py --config-file tests/fixtures/smoke_config.yaml
python multiplication_quiz.py --config-file tests/fixtures/smoke_config.yaml
python tests/validate_smoke_pdfs.py
```

PDF 校验会检查文件数量和页数、题量、加减法答案、乘法答案，以及乘法因数是否都位于 1-6。发布前还应使用 Poppler 将所有 PDF 页面渲染成 PNG，逐页检查裁切、重叠、字体和分页。

## Git 同步

`COMMON_PROJECT_SKILLS.md`、`common.env`、日志、输出文件和测试临时产物不会提交。代码、README、`config.yaml`、`requirements.txt` 和 `common.env.example` 应正常同步。

```powershell
git pull --rebase origin main
git push origin main
```
