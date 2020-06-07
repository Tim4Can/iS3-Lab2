from docx import Document
import os
import csv
import util
from library.PS.FileProcess_PS_S1S2 import Record


class Test:

    def compareResult(self, record):

        SKTH_CHAI = record.dict["掌子面桩号"]
        SKTH_INTE = record.dict["桩号区间"]
        SKTH_WATE = record.dict["地下水状态描述"]
        SKTH_WATG = record.dict["地下水对应等级"]
        SKTH_LITH = record.dict["岩性"]
        SKTH_FORM = record.dict["岩层产状"]
        SKTH_WEA = record.dict["风化程度"]
        SKTH_JOIQ = record.dict["节理数"]
        SKTH_JOIN = record.dict["节理倾角"]
        SKTH_INTE2 = record.dict["完整性"]
        SKTH_INTG = record.dict["完整性对应等级"]
        SKTH_SURR = record.dict["围岩级别"]
        SKTH_STRU = record.dict["结构面形状"]
        SKTH_FAUL = record.dict["断层"]
        SKTH_STRE = record.dict["高应力、特殊地质"]
        SKTH_STAB = record.dict["围岩稳定情况"]

        with open('../../standard/PS/standard_PS_S1S2.csv', 'r', encoding='utf-8') as csv_file:
            csv_read = csv.reader(csv_file)
            results = []
            details = []
            flag = False
            for row in csv_read:
                if SKTH_CHAI == row[0]:
                    results.append(not flag)
                    if SKTH_INTE == row[1]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_INTE) + " \n标准输出：" + str(row[1]))

                    if SKTH_WATE == row[2]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_WATE) + " \n标准输出：" + str(row[2]))

                    if SKTH_WATG == row[3]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_WATG) + " \n标准输出：" + str(row[3]))

                    if SKTH_LITH == row[4]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_LITH) + " \n标准输出：" + str(row[4]))

                    if SKTH_FORM == row[5]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_FORM) + " \n标准输出：" + str(row[5]))

                    if SKTH_WEA == row[6]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_WEA) + " \n标准输出：" + str(row[6]))

                    if SKTH_JOIQ == row[7]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_JOIQ) + " \n标准输出：" + str(row[7]))

                    if SKTH_JOIN == row[8]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_JOIN) + " \n标准输出：" + str(row[8]))

                    if SKTH_INTE2 == row[9]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_INTE2) + " \n标准输出：" + str(row[9]))

                    if SKTH_INTG == row[10]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_INTG) + " \n标准输出：" + str(row[10]))

                    if SKTH_SURR == row[11]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_SURR) + " \n标准输出：" + str(row[11]))

                    if SKTH_STRU == row[12]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_STRU) + " \n标准输出：" + str(row[12]))

                    if SKTH_FAUL == row[13]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_FAUL) + " \n标准输出：" + str(row[13]))

                    if SKTH_STRE == row[14]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_STRE) + " \n标准输出：" + str(row[14]))

                    if SKTH_STAB == row[15]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(SKTH_STAB) + " \n标准输出：" + str(row[15]))

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
            print("测试完成PS_S1S2文件" + file)

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
    inputpath = "E:/Education/409iS3/土木土木/数据提取文件/数据提取文件示例/源数据/掌子面地质素描/S1-S2标数据"
    a = Execute()
    a.run(inputpath)
