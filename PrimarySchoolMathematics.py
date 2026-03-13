import random
import yaml
import os
import logging
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from logging_config import setup_logger


class MathQuizGenerator:
    def __init__(self, config_path="config.yaml"):
        # 加载配置
        with open(config_path, "r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f)

        # 初始化日志
        log_level_str = self.config.get("log_level", "INFO").upper()
        log_level = getattr(logging, log_level_str, logging.INFO)
        self.logger = setup_logger(log_level=log_level)
        self.logger.info("初始化小学口算题生成器...")

    def generate_problem(self):
        """根据配置随机生成一道题目"""
        settings = self.config["quiz"]["settings"]
        setting = random.choice(settings)

        steps = setting.get("steps", 1)
        t1_min = setting.get("term1_min", 1)
        t1_max = setting.get("term1_max", 100)
        t2_min = setting.get("term2_min", 1)
        t2_max = setting.get("term2_max", 100)
        ops_pool = setting.get("operators", ["+"])
        r_min = setting.get("result_min", 0)
        r_max = setting.get("result_max", 1000)

        # 尝试生成符合结果范围的题目
        for _ in range(100):  # 最多重试100次
            a = random.randint(t1_min, t1_max)
            problem_str = str(a)
            current_value = a

            valid = True
            for i in range(steps):
                op = random.choice(ops_pool)
                b = random.randint(t2_min, t2_max)

                if op == "+":
                    current_value += b
                elif op == "-":
                    if current_value < b:  # 简单防止负数
                        valid = False
                        break
                    current_value -= b
                elif op == "*":
                    current_value *= b
                elif op == "/":
                    if b == 0 or current_value % b != 0:
                        valid = False
                        break
                    current_value //= b

                # 转换符号显示
                display_op = op.replace("*", "×").replace("/", "÷")
                problem_str += f" {display_op} {b}"

            if valid and r_min <= current_value <= r_max:
                result_text = f"{problem_str} ="
                return result_text

        return "1 + 1 ="  # 保底方案

    def create_docx(self):
        """生成 Word 文档"""
        quiz_cfg = self.config["quiz"]
        count = quiz_cfg.get("count", 100)
        pages = quiz_cfg.get("pages", 1)
        columns = quiz_cfg.get("columns", 4)
        title = quiz_cfg.get("title", "小学生口算题")
        output_file = quiz_cfg.get("output_file", "小学口算题_v2.docx")

        self.logger.info(
            f"开始生成 {pages} 页，每页 {count} 道题目，并导出到 {output_file}"
        )

        doc = Document()
        all_unique_problems = set()

        # 读取字体配置
        font_name = quiz_cfg.get("font_name", "黑体")
        font_size = quiz_cfg.get("font_size", 22)
        info_font_size = quiz_cfg.get("info_font_size", 16)
        margin_cm = quiz_cfg.get("margin_cm", 1.0)
        orientation = quiz_cfg.get("orientation", "landscape")

        for page in range(pages):
            if page > 0:
                doc.add_page_break()

            section = doc.sections[-1]  # 获取当前最后一节

            # 设置页面方向
            if orientation == "landscape":
                new_width, new_height = section.page_height, section.page_width
                section.orientation = WD_ORIENT.LANDSCAPE
                section.page_width = new_width
                section.page_height = new_height

            # 设置页边距
            section.top_margin = Cm(margin_cm)
            section.bottom_margin = Cm(margin_cm)
            section.left_margin = Cm(margin_cm)
            section.right_margin = Cm(margin_cm)

            # 添加标题
            heading = doc.add_heading(title, 0)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 添加 姓名、日期、时间 这一行
            info_para = doc.add_paragraph()
            info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            info_run = info_para.add_run(
                "姓名：__________ 日期：____月____日 时间：________ 对题：____道"
            )
            info_run.font.size = Pt(info_font_size)
            info_run.font.name = font_name
            # 设置中文字体，确保 rFonts 存在
            rPr = info_run._element.get_or_add_rPr()
            rFonts = rPr.get_or_add_rFonts()
            rFonts.set(qn("w:eastAsia"), font_name)

            rows = (count + columns - 1) // columns
            table = doc.add_table(rows=rows, cols=columns)
            table.autofit = True

            current_page_problems = []
            max_retries_per_prob = 1000  # 增加单次尝试次数

            for _ in range(count):
                found_new = False
                # 优先尝试在全局范围内去重
                for _retry in range(max_retries_per_prob):
                    prob = self.generate_problem()
                    if prob not in all_unique_problems:
                        all_unique_problems.add(prob)
                        current_page_problems.append(prob)
                        found_new = True
                        break

                # 如果全局重复但在本页没出现过，允许出现在不同页（解决组合不足的问题）
                if not found_new:
                    for _retry in range(max_retries_per_prob):
                        prob = self.generate_problem()
                        if prob not in set(current_page_problems):
                            current_page_problems.append(prob)
                            all_unique_problems.add(prob)  # 再次标注
                            found_new = True
                            self.logger.warning(
                                f"警告：第 {page + 1} 页尝试生成全局唯一题目失败，已改用页内唯一模式。"
                            )
                            break

                if not found_new:
                    self.logger.error(
                        f"严重警告：第 {page + 1} 页范围极窄，已无法维持页内题目唯一性。"
                    )
                    current_page_problems.append(self.generate_problem())

            for i in range(len(current_page_problems)):
                row = i // columns
                col = i % columns
                cell = table.cell(row, col)
                paragraph = cell.paragraphs[0]
                run = paragraph.add_run(current_page_problems[i])

                # 设置字体和大小
                run.font.size = Pt(font_size)
                run.font.name = font_name
                # 兼容中文字体设置
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                rFonts.set(qn("w:eastAsia"), font_name)

        doc.save(output_file)
        self.logger.info(f"成功生成文档: {os.path.abspath(output_file)}")


if __name__ == "__main__":
    generator = MathQuizGenerator()
    generator.create_docx()
