from docx import Document
import os
import csv
import util
from library.GPR.FileProcess_GPR_S1S2 import Record, RecordPDF


class Test:

    def compareResult(self, record):

        GSI_CHAI = record.dict["掌子面桩号"]
        GSI_INTE = record.dict["桩号区间"]
        GSI_GPR = record.dict["地质雷达描述"]
        GSI_WATE = record.dict["地下水状态描述"]
        GSI_WATG = record.dict["地下水对应等级"]
        GSI_LITH = record.dict["岩性"]
        GSI_WEA = record.dict["风化程度"]
        GSI_STRU = record.dict["结构构造"]
        GSI_FAUL = record.dict["断层"]
        GSI_STAB = record.dict["稳定性"]
        GSI_DSCR = record.dict["设计围岩级别"]
        GSI_PSRL = record.dict["预报围岩级别"]

        with open('../../standard/GPR/standard_GPR_S1S2.csv', 'r', encoding='utf-8') as csv_file:
            csv_read = csv.reader(csv_file)
            results = []
            details = []
            flag = False

            for row in csv_read:
                if GSI_CHAI == row[0]:
                    results.append(not flag)
                    if GSI_INTE == row[1]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_INTE) + " \n标准输出：" + str(row[1]))

                    if GSI_GPR == row[2]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_GPR) + " \n标准输出：" + str(row[2]))

                    if GSI_WATE == row[3]:
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

                    if GSI_STRU == row[7]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_STRU) + " \n标准输出：" + str(row[7]))

                    if GSI_FAUL == row[8]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_FAUL) + " \n标准输出：" + str(row[8]))

                    if GSI_STAB == row[9]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_STAB) + " \n标准输出：" + str(row[9]))

                    if GSI_DSCR == row[10]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_DSCR) + " \n标准输出：" + str(row[10]))

                    if GSI_PSRL == row[11]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GSI_PSRL) + " \n标准输出：" + str(row[11]))

        print(results)
        print(details)
        return results, details


class Execute:

    def run(self, input_path):
        files_to_process = set()
        files_to_transform = set()
        pdf_to_process = set()
        # 存储所有文件的测试结果
        result_set = []
        count = 0

        for file in os.listdir(input_path):
            absolute_file_path = os.path.join(input_path, file)
            if file.endswith(".doc"):
                files_to_transform.add(absolute_file_path)
            elif file.endswith(".docx"):
                files_to_process.add(absolute_file_path)
            elif file.endswith(".pdf"):
                pdf_to_process.add(file)

        files_to_delete = util.batch_doc_to_docx(files_to_transform)
        files_to_process = files_to_process.union(files_to_delete)

        # Word
        for file in files_to_process:
            docx = Document(file)
            record = Record(docx)
            test = Test()
            result, detail = test.compareResult(record)
            fault = self.filter(result, detail)
            if len(fault) != 0:
                result_set.append(fault)
            count = count + 1
            print("测试完成GPR_S1S2 Word文件" + file)

        # PDF
        for file in pdf_to_process:
            record = RecordPDF(file)
            test = Test()
            result, detail = test.compareResult(record)
            fault = self.filter(result, detail)
            if len(fault) != 0:
                result_set.append(fault)
            count = count + 1
            print("测试完成GPR_S1S2 PDF文件" + file)

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
    inputpath = "E:/Education/409iS3/土木土木/数据提取文件/数据提取文件示例/源数据/GPR地质预报/S1-S2标数据"
    a = Execute()
    a.run(inputpath)
