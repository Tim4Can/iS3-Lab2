from docx import Document
import os
import csv
import util
from library.SCHE.FileProcess_SCHE import Record
import xlrd
import pandas as pd


class Test:

    def compareResult(self, record):
        dataset = record.dataset
        results = []
        details = []
        flag = False
        start = dataset[0]["标段号"]
        PROG_NAME = "标段号"
        PROG_LORR = "左右幅"
        PROG_DATE = "记录时间"
        PROG_END = "桩号区间"
        PROG_SGID = "开挖进度"
        PROG_CQJD = "超前支护进度"
        PROG_CCJD = "初衬进度"
        PROG_ECJD = "二衬进度"
        FILE_FSET = "关联文件"

        df = pd.read_csv('./standard/SCHE/standard_SCHE.csv', header=0)
        nos = df['标段号'].tolist()
        pos = nos.index(start)
        if pos < 0:
            print("找不到起始标段号！")
            results.append(not flag)
            details.append("找不到起始标段号！")
            return results, details
        

        for i in range (len(dataset)):
            row = df.iloc[pos + i]
            data = dataset[i]

            if data[PROG_NAME] == row[PROG_NAME]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(PROG_NAME + "\t程序输出：" + str(data[PROG_NAME]) + " \t标准输出：" + str(row[PROG_NAME]))
            
            if data[PROG_LORR] == row[PROG_LORR]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(PROG_LORR + "\t程序输出：" + str(data[PROG_LORR]) + " \t标准输出：" + str(row[PROG_LORR]))

            if data[PROG_DATE] == row[PROG_DATE]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(PROG_DATE + "\t程序输出：" + str(data[PROG_DATE]) + " \t标准输出：" + str(row[PROG_DATE]))
            
            if data[PROG_END] == row[PROG_END]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(PROG_END + "\t程序输出：" + str(data[PROG_END]) + " \t标准输出：" + str(row[PROG_END]))

            if data[PROG_SGID] == row[PROG_SGID]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(PROG_SGID + "\t程序输出：" + str(data[PROG_SGID]) + " \t标准输出：" + str(row[PROG_SGID]))
            
            if data[PROG_CQJD] == row[PROG_CQJD]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(PROG_CQJD + "\t程序输出：" + str(data[PROG_CQJD]) + " \t标准输出：" + str(row[PROG_CQJD]))

            if data[PROG_CCJD] == row[PROG_CCJD]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(PROG_CCJD + "\t程序输出：" + str(data[PROG_CCJD]) + " \t标准输出：" + str(row[PROG_CCJD]))
            
            if data[PROG_ECJD] == row[PROG_ECJD]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(PROG_ECJD + "\t程序输出：" + str(data[PROG_ECJD]) + " \t标准输出：" + str(row[PROG_ECJD]))

            if data[FILE_FSET] == row[FILE_FSET]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(FILE_FSET + "\t程序输出：" + str(data[FILE_FSET]) + " \t标准输出：" + str(row[FILE_FSET]))
        
        return results, details


class Execute:

    def run(self, input_path):
        result_set = []
        count = 0

        if os.path.isfile(input_path):
            if input_path.endswith(".xlsx") or input_path.endswith(".xls"):
                excel = xlrd.open_workbook(input_path)
                record = Record(excel)

                test = Test()
                result, detail = test.compareResult(record)
                fault = self.filter(result, detail)
                if len(fault) != 0:
                    result_set.append(fault)
                count = count + 1
                print("测试完成SCHE excel文件" + input_path)
        

        return count, result_set
    
    def filter(self, results, details):
        if len(results) != len(details):
            print("长度不一致！")
            return
        fault = []
        for i in range(len(details)):
            if results[i] == True:
                continue
            fault.append(details[i])
        return fault