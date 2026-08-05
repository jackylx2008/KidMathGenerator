# Kid Math Generator

用于生成小学生加减法口算题、九九乘法口算题和竖式计算练习。程序会分别创建题目卷、答案卷 DOCX，并在 Windows 或 macOS 上调用 Microsoft Word 转换为 PDF。

## 运行环境

- Python 3.12+
- Windows 或 macOS
- Microsoft Word（仅自动转换 PDF 时需要）

安装依赖：

```powershell
python -m pip install -r requirements.txt
```

依赖会根据当前系统自动选择 PDF 转换组件：Windows 安装 `comtypes`，
macOS 安装 `docx2pdf`。macOS 首次转换时需要允许终端或 Python 控制
Microsoft Word。Linux 暂不支持 Microsoft Word 自动转换，可将
`app.convert_to_pdf` 设为 `false`，仅生成 DOCX。

macOS 路径已使用 Microsoft Word for Mac 完成 DOCX 到 PDF 的实际转换验证。

## 运行命令

生成加减法题：

```powershell
python addition_subtraction_quiz.py
```

生成九九乘法题：

```powershell
python multiplication_quiz.py
```

生成加减法竖式计算练习：

```powershell
python vertical_arithmetic_quiz.py
```

可以指定其他配置文件：

```powershell
python multiplication_quiz.py --config-file tests/fixtures/smoke_config.yaml
```

生成一页竖式 DOCX 烟雾测试文件（不调用 Word 转 PDF）：

```powershell
python vertical_arithmetic_quiz.py --config-file tests/fixtures/vertical_smoke_config.yaml
```

转换一个或多个指定的 DOCX 文件：

```powershell
python convert_to_pdf.py output/sample.docx --output-dir output
```

## 配置

统一配置位于根目录 `config.yaml`：

- `app`
  - `log_level`：日志级别。
  - `output_dir`：输出根目录，默认 `output`。
  - `convert_to_pdf`：是否调用 Word 转换 PDF。
  - `delete_docx_after_pdf`：PDF 成功后是否删除 DOCX，默认删除。
  - `random_seed`：可选随机种子，设置整数后可以复现同一批题。
- `flows.addition_subtraction`
  - 页数、每页题量、排版、图章和加减法数值范围。
  - `settings` 支持一步或多步运算、运算符比例及结果范围。
- `flows.multiplication`
  - `factor_min` / `factor_max` 控制两个因数的范围。
  - 默认值为 `1` 和 `9`，因此生成完整的九九乘法题。
  - `label_enabled` 启用题目卷盖章；每页从 `src` 下的图片随机选择一张，
    经过透明化、压缩和小角度旋转后放置在页面左上角，答案卷不盖章。
- `flows.vertical_arithmetic`
  - `pages`、`count`、`columns` 控制页数、每页题量和列数。
  - `settings` 中通过 `operation` 选择 `addition` 或 `subtraction`。
  - `weight` 控制不同规则在混合练习中的比例。
  - `carry` / `borrow` 支持 `none`、`required`、`any`。
  - `carry_count_min`、`carry_count_max`、`borrow_count_min`、
    `borrow_count_max` 可进一步控制进位或借位次数。
  - `show_working_in_answer` 控制答案卷是否显示进位和借位标记。
  - `operator_font_name` / `operator_font_size` 统一控制加减运算符的字体和字号。
  - 运算符始终紧贴两个操作数中位数较多者的最左侧，不受结果位数影响。
  - `label_enabled` 启用题目卷盖章；每页从 `src` 下的图片随机选择一张，
    经过透明化和小角度旋转后放置在页面左上角，答案卷不盖章。
  - `hard_label_max_width_px` 限制嵌入图章的像素宽度，避免多页 PDF 体积过大。

所有题目卷和答案卷默认使用 A4 横向页面，竖式计算也遵循这一规则。新增生成流程时应复用通用文档生成器的页面设置；只有明确配置 `orientation: portrait` 时才改为纵向。

竖式第一阶段支持加法进位和减法借位。数据模型已经为乘法部分积、除法商、
余数和分步计算预留结构，后续会沿用同一入口扩展。

本机差异可写入不入库的 `common.env`，格式参考 `common.env.example`。配置值支持
`${ENV_VAR:-default}` 形式的环境变量覆盖；进程中已有的环境变量优先于
`common.env`。

跨平台同步路径统一使用 `${CLOUDSTATION_ROOT}`。加载配置时会优先读取显式的
`CLOUDSTATION_ROOT`，否则根据系统选择
`CLOUDSTATION_ROOT_WINDOWS`、`CLOUDSTATION_ROOT_MACOS` 或
`CLOUDSTATION_ROOT_LINUX`，并自动展开 `~`。

## 输出

- `output/`：统一存放口算和竖式练习的题目卷、答案卷 DOCX/PDF。
- `logs/`：按入口脚本命名的滚动日志；单文件上限 10 MB，保留 5 份备份。

生成过程中会先创建 DOCX；确认对应 PDF 已成功生成且文件非空后，默认删除该 DOCX。PDF 转换失败时会保留 DOCX，便于恢复和排查。
PDF 转换成功后还会校核每一页，并自动删除不含文字、图片、注释或矢量绘制内容的空白页；若整份 PDF 都被判定为空白，则保留原文件以避免误删。

开启加减法 `hard_label` 后，程序读取 `src` 根目录中的第一张图片作为图章，只盖在题目卷上；输出文件名会追加 `_难题`。

## 项目结构

```text
addition_subtraction_quiz.py
multiplication_quiz.py
vertical_arithmetic_quiz.py
convert_to_pdf.py
logging_config.py
config.yaml
common.env.example
src/kid_math_generator/
├── config_loader.py
├── context.py
├── modules/
│   ├── addition_subtraction.py
│   ├── multiplication.py
│   ├── document_builder.py
│   ├── vertical_arithmetic.py
│   ├── vertical_document_builder.py
│   └── pdf_converter.py
└── flows/
    ├── addition_subtraction_flow.py
    ├── multiplication_flow.py
    ├── vertical_arithmetic_flow.py
    └── _quiz_flow.py
docs/
tests/
logs/
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

`docs/COMMON_PROJECT_SKILLS.md` 属于项目文档，应与代码一同提交。本地
`common.env`、`logs/`、`output/` 和测试临时产物不提交；代码、README、
`config.yaml`、`requirements.txt` 和 `common.env.example` 应正常同步。

提交信息统一使用以下格式：

```text
feat: 中文功能描述
refactor: 中文重构描述
```

```powershell
git pull --rebase origin main
git push origin main
```
