import os
import comtypes.client
import logging


def convert_docx_to_pdf():
    # 初始化日志
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
    )
    logger = logging.getLogger(__name__)

    # 获取当前目录
    current_dir = os.path.abspath(os.getcwd())
    logger.info(f"正在扫描目录: {current_dir}")

    # 获取 word 实例
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False

    try:
        for filename in os.listdir(current_dir):
            if filename.endswith(".docx") and not filename.startswith("~$"):
                docx_path = os.path.join(current_dir, filename)
                pdf_filename = filename.replace(".docx", ".pdf")
                pdf_path = os.path.join(current_dir, pdf_filename)

                logger.info(f"正在转换: {filename} -> {pdf_filename}")

                # 打开文档
                doc = word.Documents.Open(docx_path)
                # 保存为 PDF (17 是 Word 中 PDF 的常量值)
                doc.SaveAs(pdf_path, FileFormat=17)
                doc.Close()

        logger.info("转换任务全部完成。")
    except Exception as e:
        logger.error(f"转换过程中出现错误: {e}")
    finally:
        word.Quit()


if __name__ == "__main__":
    convert_docx_to_pdf()
