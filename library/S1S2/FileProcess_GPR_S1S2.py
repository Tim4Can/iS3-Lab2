# coding=utf-8
from docx import Document
import os
import csv
# sys.path.append("..")
from ..FileProcessBasic import FileProcessBasic
from win32com.client import Dispatch

def check_output_file(output_path, header):
    if not os.path.exists(output_path):
        with open(output_path, "w", encoding="utf_8_sig", newline="") as f:
            w = csv.DictWriter(f, header)
            w.writeheader()

class Record:
    def __init__(self, docx):
        # self.face_station = None  # 掌子面桩号
        # self.number_interval = None  # 桩号区间
        # self.GPR_description = None  # 地质雷达描述
        # self.ground_water_description = None  # 地下水状态描述
        # self.ground_water_level = None  # 地下水对应等级
        # self.lithology = None  # 岩性
        # self.weathering_degree = None  # 风化程度
        # self.structure = None  # 结构构造
        # self.fault = None  # 断层
        # self.stability = None  # 稳定性
        # self.designed_rock_level = None  # 设计围岩级别
        # self.predict_rock_level = None  # 预报围岩级别
        self.dict = {
            "掌子面桩号": None,
            "桩号区间": None,
            "地质雷达描述": None,
            "地下水状态描述": None,
            "地下水对应等级": None,
            "岩性": None,
            "风化程度": None,
            "结构构造": None,
            "断层": None,
            "稳定性": None,
            "设计围岩级别": None,
            "预报围岩级别": None
        }
        self.get_face_station(docx.tables[2])


    def get_face_station(self, table):
        followed = [[0, 2], [1, 0], [1, 2]]
        checked = [2, 13, 14, 15, 16, 17]
        for i in range(len(followed)):
            row = followed[i][0]
            name = followed[i][1]
            tmp = list(table.rows[row].cells)
            cols = sorted(set(tmp), key=tmp.index)
            self.dict[cols[name].text.replace(' ', '')] = cols[name + 1].text

        # 获取单元格后选择结果类型
        for i in checked:
            tmp = list(table.rows[i].cells)
            cols = sorted(set(tmp), key=tmp.index)
            name = cols[0].text.replace(' ', '')
            for col in cols:
                if col.text.find('√') > 0:
                    col.text = col.text.replace('√', '')
                    self.dict[name] = col.text
                    break
        # return ""

    def to_string(self):
        res = ""
        for name, value in vars(self).items():
            if value is None:
                value = ""
            res = res + value + ","
        return res[: -1]

class Processor(FileProcessBasic):
    def save(self, output, record):
        output_path = os.path.join(output, "GPR_S1S2.csv")
        header = record.dict.keys()
        check_output_file(output_path, header)

        with open(output_path, "a+", encoding="utf_8_sig", newline="") as f:
            w = csv.DictWriter(f, record.dict.keys())
            w.writerow(record.dict)

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

        docx = Document(inputpath)
        record = Record(docx)
        self.save(outputpath, record)

        print("提取完成" + inputpath)
        if transformed and os.path.exists(inputpath):
            os.remove(inputpath)


if __name__ == "__main__":
    test = Processor()
    inputpath = "D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/task3/GPRS1S2.docx"
    outputpath = "D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/task3"
    test.run(inputpath, outputpath)
