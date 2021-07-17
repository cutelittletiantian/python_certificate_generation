from win32com.client import Dispatch
import os

# 读取word的文件夹（注意：在word转pdf脚本中，调用路径必须是绝对路径！！！相对路径需要用函数转一道变成绝对路径才行）
wordListPath = os.path.abspath(path=r"resources/证书_自动生成")
if not os.path.exists(path=wordListPath):
    raise FileExistsError("未检测到有证书，请先生成证书然后稍后再试")

# 保存pdf结果的文件夹
pdfListPath = os.path.abspath(path=r"resources/转换_证书文件_pdf")
if not os.path.exists(path=pdfListPath):
    os.mkdir(path=pdfListPath)


def doc2pdf(word_item_path, word_item):
    print(f"正在将文件【{word_item}】转换为pdf文件......")
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(word_item_path)
    # 生成pdf文件路径名称
    generate_pdf_item_path = pdfListPath + "\\" + word_item.split('.')[0] + ".pdf"
    doc.SaveAs(generate_pdf_item_path, FileFormat=17)
    doc.Close()
    word.Quit()
    print(f"文件【{word_item}】转换成功！")


if __name__ == "__main__":
    wordItemList = os.listdir(wordListPath)
    for wordItem in wordItemList:
        if (wordItem.endswith(".doc") or wordItem.endswith(".docx")) and ("~$" not in wordItem):
            filePath = f"{wordListPath}/{wordItem}"
            doc2pdf(filePath, wordItem)
    print("所有word文件转PDF文件已完成！！！")
