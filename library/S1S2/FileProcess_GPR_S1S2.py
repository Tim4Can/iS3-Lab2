# coding=utf-8
from docx import Document
import os
import csv
# sys.path.append("..")
from ..FileProcessBasic import FileProcessBasic
from win32com.client import Dispatch


class Processor(FileProcessBasic):
    def __init__(self):
        self.followed = []
        self.checked = []
        self.dic = {}

    # 查找关键字
    def findKeywords(self):
        # 单元格后紧跟结果所在行，列
        self.followed = [[0, 2], [1, 0], [1, 2]]
        # 单元格后选择结果所在行
        self.checked = [2, 13, 14, 15, 16, 17]

    # 寻找上下文
    def findContent(self, table):
        # 获取单元格后紧跟结果类型
        for i in range(len(self.followed)):
            row = self.followed[i][0]
            name = self.followed[i][1]
            tmp = list(table.rows[row].cells)
            cols = sorted(set(tmp), key=tmp.index)
            self.dic[cols[name].text.replace(' ', '')] = cols[name + 1].text

        # 获取单元格后选择结果类型
        for i in self.checked:
            tmp = list(table.rows[i].cells)
            cols = sorted(set(tmp), key=tmp.index)
            name = cols[0].text.replace(' ', '')
            for col in cols:
                if col.text.find('√') > 0:
                    col.text = col.text.replace('√', '')
                    self.dic[name] = col.text
                    break

    # 后续的处理
    def subsequentProcess(self):
        for key, value in self.dic.items():
            print(key, ' ', value)

    def check_output_file(self, output_path):
        if not os.path.exists(output_path):
            with open(output_path, "w", encoding="utf_8_sig", newline="") as f:
                w = csv.DictWriter(f, self.dic.keys())
                w.writeheader()

    def save(self, output):
        output_path = os.path.join(output, "GPR_S1S2.csv")
        self.check_output_file(output_path)
        with open(output_path, "a+", encoding="utf_8_sig", newline="") as f:
            w = csv.DictWriter(f, self.dic.keys())
            w.writerow(self.dic)

    def run(self, inputpath, outputpath):
        transformed = False
        if inputpath.endswith("doc"):
            transformed = True
            try:
                word = Dispatch('Word.Application')
                doc = word.documents.Open(inputpath)
                doc.SaveAs("{}x".format(inputpath), 12)
                doc.Close()
                word.Quit()
            except IOError:
                print("读取文件异常：" + inputpath)
            inputpath = inputpath + "x"
        elif not inputpath.endswith("docx"):
            return

        table = Document(inputpath).tables[2]
        self.findKeywords()
        self.findContent(table)
        self.subsequentProcess()
        self.save(outputpath)

        print("提取完成" + inputpath)
        if transformed and os.path.exists(inputpath):
            os.remove(inputpath)


if __name__ == "__main__":
    test = Processor()
    inputpath = "D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/task3/GPRS1S2.docx"
    outputpath = "D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/task3"
    test.run(inputpath, outputpath)
