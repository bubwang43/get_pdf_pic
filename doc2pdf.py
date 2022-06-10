from win32com.client import Dispatch
from os import walk
import os

wdFormatPDF = 17


def doc2pdf(input_file, filename, save_path):
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    save_path = str(save_path + "\\" + filename.replace(".docx", ".pdf").replace(".doc", ".pdf"))
    doc.SaveAs(save_path, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    print('%s转换完成' % filename)


if __name__ == "__main__":
    doc_files = []
    directory = "D:\\cicc\\word\\2018年\\原版"
    for root, dirs, filenames in walk(directory):
        for filename in filenames:
            if filename.endswith(".doc") or filename.endswith(".docx"):
                file_path = str(root + "\\" + filename)
                doc2pdf(file_path, filename, "D:\\cicc\\word2pdf\\2018")
        print("转换结束!")
