from docx import Document
import pdfplumber as plb
import os
import csv
import util
from library.TSP.FileProcess_TSP_S1S2 import Record, RecordPDF, get_INTE, get_INTE_PDF


class Test:

    # def __init__(self, record):
        # self.results, self.details = self.compareResult(record)

    def compareResult(self, record):

        GPRF_INTE = record.dict["桩号区间"]
        GPRF_LITH = record.dict["岩性"]
        GPRF_MD = record.dict["密度ρ（g/cm3）"]
        GPRF_ZPS = record.dict["纵波速度Vp（m/s）"]
        GPRF_HPS = record.dict["横波速度Vs（m/s）"]
        GPRF_PSB = record.dict["泊松比σ"]
        GPRF_EM = record.dict["动态杨氏模量（GPa）"]
        GPRF_FORE = record.dict["预报结果描述"]
        GPRF_WEA = record.dict["风化程度"]
        GPRF_WATE = record.dict["地下水"]
        GPRF_INTE2 = record.dict["完整性"]
        GPRF_STAB = record.dict["稳定性"]
        GPRF_FAUL = record.dict["特殊地质情况"]
        GPRF_PSRL = record.dict["推测围岩级别"]
        GPRF_STRE = record.dict["设计围岩级别"]

        with open('./standard/TSP/standard_TSP_S1S2.csv', 'r', encoding='utf-8') as csv_file:
            csv_read = csv.reader(csv_file)
            results = []
            details = []
            flag = False
            # print(csv_read)
            for row in csv_read:
                if GPRF_INTE == row[0]:
                    results.append(not flag)
                    if GPRF_LITH != "" and (GPRF_LITH in row[1] or row[1] in GPRF_LITH):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_LITH) + " \n标准输出：" + str(row[1]))

                    if GPRF_MD == row[2]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_MD) + " \n标准输出：" + str(row[2]))

                    if GPRF_ZPS == row[3]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_ZPS) + " \n标准输出：" + str(row[3]))

                    if GPRF_HPS == row[4]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_HPS) + " \n标准输出：" + str(row[4]))

                    if GPRF_PSB == row[5]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_PSB) + " \n标准输出：" + str(row[5]))

                    if GPRF_EM == row[6]:
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_EM) + " \n标准输出：" + str(row[6]))

                    if GPRF_FORE != "" and (GPRF_FORE in row[7] or row[7] in GPRF_FORE):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_FORE) + " \n标准输出：" + str(row[7]))

                    if GPRF_WEA != "" and (GPRF_WEA in row[8] or row[8] in GPRF_WEA):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_WEA) + " \n标准输出：" + str(row[8]))

                    if GPRF_WATE != "" and (GPRF_WATE in row[9] or row[9] in GPRF_WATE):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_WATE) + " \n标准输出：" + str(row[9]))

                    if GPRF_INTE2 != "" and (GPRF_INTE2 in row[10] or row[10] in GPRF_INTE2):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_INTE2) + " \n标准输出：" + str(row[10]))

                    if GPRF_STAB != "" and (GPRF_STAB in row[11] or row[11] in GPRF_STAB):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_STAB) + " \n标准输出：" + str(row[11]))

                    if GPRF_FAUL != "" and (
                            GPRF_FAUL in row[12].replace(" ", "") or row[12].replace(" ", "") in GPRF_FAUL):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_FAUL) + " \n标准输出：" + str(row[12]))

                    if GPRF_PSRL != "" and (GPRF_PSRL in row[13] or row[13] in GPRF_PSRL):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_PSRL) + " \n标准输出：" + str(row[13]))

                    if GPRF_STRE != "" and (GPRF_STRE in row[14] or row[14] in GPRF_STRE):
                        results.append(not flag)
                        details.append("")
                    else:
                        results.append(flag)
                        details.append("程序输出：" + str(GPRF_STRE) + " \n标准输出：" + str(row[14]))

        #print(results)
        #print(details)
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
            INTE = get_INTE(docx.tables[1])
            for i in range(len(INTE)):
                record = Record(docx, INTE[i])
                test = Test()
                result, detail = test.compareResult(record)
                fault = self.filter(result, detail)
                if len(fault) != 0:
                    result_set.append(fault)
            count = count + 1
            print("测试完成Word文件" + file)

        # PDF
        for file in pdf_to_process:
            path = os.path.join(input_path, file)
            para_conclusion = ""
            para_suggestion = ""
            get_conclusion = False
            get_suggestion = False
            table_analysis = []
            table_result = []
            with plb.open(path) as pdf:
                for page in pdf.pages:
                    page_content = page.extract_text()
                    sentences = page_content.split("\n")

                    for s in sentences:
                        s = s.strip()
                        if s.startswith("第") and s.endswith("页"):
                            continue

                        if s.startswith("7.1"):
                            get_conclusion = True
                            continue
                        elif s.startswith("7.2"):
                            get_conclusion = False
                        if get_conclusion:
                            para_conclusion += s
                            continue

                        if s.startswith("7.2"):
                            get_suggestion = True
                            continue
                        elif s.startswith("8"):
                            get_suggestion = False
                            break
                        if get_suggestion:
                            para_suggestion += s
                        continue

                    if "力学指标汇总表" in page_content:
                        tables = page.extract_tables()
                        table_result = tables[0]
                    if "地质预报汇总表" in page_content:
                        tables = page.extract_tables()
                        table_analysis = tables[0]

            INTE_PDF = get_INTE_PDF(table_result)
            for i in range(len(INTE_PDF)):
                record = RecordPDF(pdf, INTE_PDF[i], table_result, table_analysis, para_conclusion)
                test = Test()
                result, detail = test.compareResult(record)
                fault = self.filter(result, detail)
                result_set.append(fault)
            count = count + 1
            print("测试完成PDF文件" + file)

        for file in files_to_delete:
            if os.path.exists(file):
                os.remove(file)

        # print(count)
        # print(result_set)
        return count, result_set


    def filter(self,results, details):
        if len(results) - 1 != len(details):
            print("长度不一致！")
            return
        fault = []
        for i in range(len(details)):
            if results[i+1] == True:
                continue
            fault.append(details[i])
        return fault





if __name__ == "__main__":
    inputpath = "D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/iS3-Lab2/test/suite/TSP地质预报/S1标"
    a = Execute()
    a.run(inputpath)
