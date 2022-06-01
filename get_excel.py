import pytesseract
from PIL import Image
import pandas as pd

pytesseract.pytesseract.tesseract_cmd = 'D:\\Tesseract-OCR\\tesseract.exe'


def getExcelFromPic(file_dir):
    contents = pytesseract.image_to_string(Image.open(file_dir), lang='chi_sim+eng')
    # print(contents)
    contents = contents.replace(' ', '')
    save("D:\\cicc\\contents1.txt", contents)
    print('success')
    # contents = contents.split('\n')
    # print(contents)
    # while '' in tiqu:  # 不能使用for
    #     tiqu.remove('')
    #     first = tiqu[:6]
    #     second = tiqu[6:12]
    #     third = tiqu[12:]
    #     df = pd.DataFrame()
    #     df[first[0]] = first[1:]
    #     df[second[0]] = second[1:]
    #     df[third[0]] = third[1:]
    # df.to_excel('图片型表格.xlsx')  #转为xlsx文件

def OCR_demo(file_dir):
    # 导入OCR安装路径，如果设置了系统环境，就可以不用设置了
    # pytesseract.pytesseract.tesseract_cmd = r"D:\Program Files\Tesseract-OCR\tesseract.exe"
    # 打开要识别的图片

    image = Image.open(file_dir)
    # 使用pytesseract调用image_to_string方法进行识别，传入要识别的图片，lang='chi_sim'是设置为中文识别，
    text = pytesseract.image_to_string(image, lang='chi_sim')
    # text = pytesseract.image_to_string(image)
    text = text.replace(' ','')
    save("D:\\cicc\\contents2.txt", text)
    print('success')

def save(filename, contents):
    file_handle = open(filename, mode='w', encoding='utf-8')
    file_handle.write(contents)
    file_handle.close()

    # file_handle = open(filename, mode='a')
    # contents = contents.split('\n')
    # for i in range(len(contents)):
    #     if(len(contents[i]) != 0):
    #         file_handle.write(contents[i] + '\n')
    # file_handle.close()

if __name__ == '__main__':
    file_dir = 'D:\\cicc\\图表 16：国企央企“双碳”行动追踪.png'
    getExcelFromPic(file_dir)
    OCR_demo(file_dir)
