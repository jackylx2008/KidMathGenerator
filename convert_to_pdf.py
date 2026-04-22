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
                doc = None

                try:
                    # 打开文档
                    doc = word.Documents.Open(docx_path)
                    # 保存为 PDF (17 是 Word 中 PDF 的常量值)
                    doc.SaveAs(pdf_path, FileFormat=17)
                    doc.Close()
                    doc = None

                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                        os.remove(docx_path)
                        logger.info(f"已删除中间文件: {filename}")
                    else:
                        logger.warning(
                            f"PDF 未正常生成，保留 DOCX 文件: {filename}"
                        )
                except Exception as file_error:
                    logger.error(f"转换文件失败 {filename}: {file_error}")
                    if doc is not None:
                        doc.Close(False)

        logger.info("转换任务全部完成。")
    except Exception as e:
        logger.error(f"转换过程中出现错误: {e}")
    finally:
        word.Quit()


if __name__ == "__main__":
    convert_docx_to_pdf()
