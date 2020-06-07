from docx import Document
import os
import csv
import util
from library.CHAG.FileProcess_CHAG import Record
import xlrd
import pandas as pd


class Test:

    def compareResult(self, record):
        print("yes")
        dataset = record.dataset
        results = []
        details = []
        flag = False
        start = dataset[0]["处理卡编号"]
        GSI_PCN = "处理卡编号"
        GSI_PRJ = "工程名称"
        GSI_MPN = "里程桩号（变更位置）"
        GSI_TYPE = "变更类型"
        GSI_BEFORE = "原衬砌类别"
        GSI_AFTER = "变更后衬砌类别"
        GSI_CR = "变更原因"
        GSI_SUG = "处理意见"
        GSI_DATE1 = "处理卡签发日期"
        GSI_DATE2 = "变更令签发日期"

        df = pd.read_csv('./standard/CHAG/standard_CHAG_S3S4.csv',header = 0)
        nos = df['处理卡编号'].tolist()
        pos = nos.index(start)
        if pos < 0:
            print("找不到起始处理卡编号！")
            results.append(not flag)
            details.append("找不到起始处理卡编号！")
            return results, details

        print(pos)

        for i in range(len(dataset)):
            row = df.iloc[pos + i]
            data = dataset[i]

            if data[GSI_PCN] == row[GSI_PCN]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_PCN + "\t程序输出：" + str(data[GSI_PCN]) + " \t标准输出：" + str(row[GSI_PCN]))

            if data[GSI_PRJ] == row[GSI_PRJ]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_PRJ + "\t程序输出：" + str(data[GSI_PRJ]) + " \t标准输出：" + str(row[GSI_PRJ]))

            if data[GSI_DATE2] == row[GSI_DATE2]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_DATE2 + "\t程序输出：" + str(data[GSI_DATE2]) + " \t标准输出：" + str(row[GSI_DATE2]))

            if data[GSI_DATE1] == row[GSI_DATE1]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_DATE1 + "\t程序输出：" + str(data[GSI_DATE1]) + " \t标准输出：" + str(row[GSI_DATE1]))

            if data[GSI_SUG] == row[GSI_SUG]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_SUG + "\t程序输出：" + str(data[GSI_SUG]) + " \n标准输出：" + str(row[GSI_SUG]))

            if data[GSI_CR] == row[GSI_CR]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_CR + "\t程序输出：" + str(data[GSI_CR]) + " \n标准输出：" + str(row[GSI_CR]))

            if data[GSI_AFTER] == row[GSI_AFTER]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_AFTER + "\t程序输出：" + str(data[GSI_AFTER]) + " \n标准输出：" + str(row[GSI_AFTER]))

            if data[GSI_BEFORE] == row[GSI_BEFORE]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_BEFORE + "\t程序输出：" + str(data[GSI_BEFORE]) + " \n标准输出：" + str(row[GSI_BEFORE]))

            if data[GSI_MPN] == row[GSI_MPN]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_MPN + "\t程序输出：" + str(data[GSI_MPN]) + " \n标准输出：" + str(row[GSI_MPN]))

            if data[GSI_TYPE] == row[GSI_TYPE]:
                results.append(not flag)
                details.append("")
            else:
                results.append(flag)
                details.append(GSI_TYPE + "\t程序输出：" + str(data[GSI_TYPE]) + " \n标准输出：" + str(row[GSI_TYPE]))

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
                print("测试完成CHAG_S3S excel文件" + input_path)


        print(count)
        print(result_set)
        return count, result_set

    def filter(self, results, details):
        if len(results) != len(details):
            print("长度不一致！")
            print(len(results))
            print(len(details))
            return
        fault = []
        for i in range(len(details)):
            if results[i] == True:
                continue
            fault.append(details[i])
        return fault

