from docx import Document
import os
import csv
import util
from library.PS.FileProcess_PS_S3S4 import Record


class Test:

    def compareResult(self, record):
        GSI_CHAI = record.dict["掌子面桩号"]
        GSI_INTE = record.dict["桩号区间"]
        GSI_WATE = record.dict["地下水状态描述"]
        GSI_WATG = record.dict["地下水对应等级"]
        GSI_LITH = record.dict["岩性"]
        GSI_RKAT = record.dict["岩层产状"]
        GSI_WEA = record.dict["风化程度"]
        GSI_JTNB = record.dict["节理数"]
        GSI_JTAG= record.dict["节理倾角"]
        GSI_ITGT = record.dict["完整性"]
        GSI_IGDG = record.dict["完整性对应等级"]
        GSI_GPR = record.dict["地质雷达描述"]


        with open('../../standard/PS/standard_PS_S3S4.csv', 'r', encoding='utf-8') as csv_file:
            csv_read = csv.reader(csv_file)
            results = []
            details = []
            flag = False
            for row in csv_read:
                if GSI_CHAI:
                    results.append(not flag)
                    if GSI_INTE == row[1]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_INTE) + " \n标准输出：" + str(row[1]))

                    if GSI_GPR.replace(" ", "") == row[2].replace(" ", ""):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_GPR) + " \n标准输出：" + str(row[2]))

                    if GSI_WATE.replace(" ", "") == row[3].replace(" ", ""):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_WATE) + " \n标准输出：" + str(row[3]))

                    if GSI_WATG == row[4]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_WATG) + " \n标准输出：" + str(row[4]))

                    if GSI_LITH == row[5]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_LITH) + " \n标准输出：" + str(row[5]))

                    if GSI_WEA == row[6]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_WEA) + " \n标准输出：" + str(row[6]))

                    if GSI_RKAT == row[7]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_RKAT) + " \n标准输出：" + str(row[7]))

                    if GSI_JTNB == row[8]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_JTNB) + " \n标准输出：" + str(row[8]))

                    if GSI_JTAG.replace(" ", "") == row[9].replace(" ", ""):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_JTAG) + " \n标准输出：" + str(row[9]))

                    if GSI_ITGT.replace(" ", "") == row[10].replace(" ", ""):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_ITGT) + " \n标准输出：" + str(row[10]))

                    if GSI_IGDG.replace(" ", "") == row[11].replace(" ", ""):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_IGDG) + " \n标准输出：" + str(row[11]))


        # print(results)
        # print(details)
        return results, details


class Execute:

    def run(self, input_path):
        files_to_process = set()
        files_to_transform = set()
        # 存储所有文件的测试结果
        result_set = []
        count = 0

        for file in os.listdir(input_path):
            absolute_file_path = os.path.join(input_path, file)
            if file.endswith(".doc"):
                files_to_transform.add(absolute_file_path)
            elif file.endswith(".docx"):
                files_to_process.add(absolute_file_path)

        files_to_delete = util.batch_doc_to_docx(files_to_transform)
        files_to_process = files_to_process.union(files_to_delete)

        for file in files_to_process:
            docx = Document(file)
            record = Record(docx)
            test = Test()
            result, detail = test.compareResult(record)
            fault = self.filter(result, detail)
            if len(fault) != 0:
                result_set.append(fault)
            count = count + 1
            print("测试完成PS_S3S4文件" + file)

        for file in files_to_delete:
            if os.path.exists(file):
                os.remove(file)

        # print(count)
        # print(result_set)
        return count, result_set

    def filter(self, results, details):
        if len(results) - 1 != len(details):
            print("长度不一致！")
            return
        fault = []
        for i in range(len(details)):
            if results[i + 1] == True:
                continue
            fault.append(details[i])
        return fault


if __name__ == "__main__":
    inputpath = "e:/study/is3/PS2"
    a = Execute()
    a.run(inputpath)
