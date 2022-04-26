import os
import sys
import time

import xlrd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image


class Action:

    def __init__(self, file_path):
        if os.path.exists(file_path):
            self.file_path = os.path.abspath(file_path)
            xlrd_wb = xlrd.open_workbook(self.file_path)
            self.xlrd_ws = xlrd_wb.sheets()[0]
            self.open_wb = load_workbook(self.file_path)
            self.open_sheet = self.open_wb.worksheets[0]
        else:
            for i in range(5, 0, -1):
                print(f"文件路径错误,{i}秒后即将自动退出")
                time.sleep(1)
            sys.exit()

    def read_excel(self):
        ex_list = []
        for i in range(1, self.xlrd_ws.nrows):
            ex_list.append(self.xlrd_ws.row_values(i))
        print(f"读取后输出列表：{ex_list}")
        return ex_list

    # 写入图片
    def write_image(self, image_path, location):
        img = Image(image_path)
        self.open_sheet.add_image(img, location)
        self.open_wb.save(self.file_path)

    # 写入单元格
    def write_value(self, value, location):
        self.open_sheet[location] = value
        self.open_wb.save(self.file_path)


if __name__ == '__main__':
    a = Action(r"test_excel/test.xlsx")
    # a.write_image(r"./image.png", "A6")
    a.write_value("hello", "A6")
