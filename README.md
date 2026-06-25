# Kid Math Generator

用于生成小学生口算题 Word 文档，并通过 Microsoft Word 转换为 PDF。配置项集中在 `config.yaml` 中。

## 运行方式

```bash
python PrimarySchoolMathematics.py
```

脚本会先生成题目卷和答案卷 DOCX，再调用 `convert_to_pdf.py` 转成 PDF。PDF 转换依赖本机已安装 Microsoft Word。

## 基础配置

主要配置位于 `config.yaml` 的 `quiz` 节：

```yaml
quiz:
  pages: 15
  count: 20
  columns: 2
  title: "小学生口算题"
  output_file: "小学口算题.docx"
  output_file_answer: "小学口算题_答案.docx"
  orientation: "landscape"
  font_name: "黑体"
  font_size: 22
  info_font_size: 16
  margin_cm: 1.0
```

## 难题图章

当 `hard_label: true` 时，会从 `src` 目录中读取第一张图片作为难题图章，只盖在题目卷上，不盖在答案卷上。

图章处理方式：

- 自动按图片左上角背景色抠透明，去掉底色。
- 每页生成独立图章图片，随机旋转和轻微位置抖动。
- 图章作为浮动图片按页面绝对坐标放置，不使用页眉，不参与正文排版，避免挤出多余页面。
- 输出文件名会自动追加 `_难题`，例如 `小学口算题_难题.pdf`、`小学口算题_答案_难题.pdf`。

推荐配置示例：

```yaml
quiz:
  hard_label: true
  hard_label_width_cm: 5.8
  hard_label_offset_x_cm: 1.2
  hard_label_offset_y_cm: 0.35
  hard_label_rotation_min: -45
  hard_label_rotation_max: 45
  hard_label_jitter_x_cm: 0.15
  hard_label_jitter_y_cm: 0.08
  hard_label_bg_tolerance: 45
```

字段说明：

- `hard_label_width_cm`: 图章宽度，单位厘米。
- `hard_label_offset_x_cm`: 图章距离页面左侧的基准位置，单位厘米。
- `hard_label_offset_y_cm`: 图章距离页面顶部的基准位置，单位厘米。
- `hard_label_rotation_min` / `hard_label_rotation_max`: 每页随机旋转角度范围。
- `hard_label_jitter_x_cm` / `hard_label_jitter_y_cm`: 每页随机位置抖动范围。
- `hard_label_bg_tolerance`: 背景色透明处理容差，图章底色残留明显时可适当调大。

## 出题配置

`settings` 控制题目范围、步数、运算符和结果范围：

```yaml
settings:
  - steps: 1
    term1_min: 11
    term1_max: 99
    term2_min: 11
    term2_max: 99
    operators1: ["+", "-"]
    operator_ratios1:
      "+": 65
      "-": 35
      "*": 0
      "/": 0
    result_min: 1
    result_max: 150
```

支持 `+`、`-`、`*`、`/`，显示时乘除会转换为 `×`、`÷`。

## 输出文件

普通配置：

- `小学口算题.pdf`
- `小学口算题_答案.pdf`

启用难题图章：

- `小学口算题_难题.pdf`
- `小学口算题_答案_难题.pdf`

`convert_to_pdf.py` 在确认 PDF 生成成功后，会删除中间 DOCX 文件。

## 依赖

Python 依赖包括：

- `python-docx`
- `PyYAML`
- `Pillow`
- `comtypes`

PDF 转换还需要 Windows 环境安装 Microsoft Word。
