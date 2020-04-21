from docx import Document
import os
import csv
import re
from library.FileProcessBasic import FileProcessBasic
import util

class Record:

    def __init__(self, inte):
        # 稳定性
        GSI_STAB = ""
        # 设计围岩级别
        GSI_DSCR = ""
        # 推测围岩级别
        GSI_ESRG = ""
        #
        # # GSI_STRU = self.get_GSI_STRU(para_suggestion)
        #
        # 桩号区间
        GSI_INTE = inte
        # 岩性
        GSI_LITH = ""
        # 风化程度
        GSI_WEA = ""
        # GSI_FAUL = self.get_GSI_FAUL(appendix)
        # GSI_FAUL = ""
        # GSI_WATG = self.get_GSI_WATG(appendix)
        # 密度
        GSI_DENST=""
        # 纵波速度
        GSI_PWL=""
        # 横波速度
        GSI_SWL=""
        #泊松比
        GSI_PR=""
        # 动态杨氏模量
        GSI_DYM=""
        # 预报结果描述
        GSI_RESULT=""
        # 完整性
        GSI_ITGT=""
        # 地下水
        GSI_WATER=""
        # 特殊地质情况
        GSI_SGS=""


        self.dict = {

            "桩号区间": GSI_INTE,
            "岩性": GSI_LITH,
            "风化程度": GSI_WEA,
            "密度": GSI_DENST,
            "纵波速度":GSI_PWL,
            "横波速度":GSI_SWL,
            "泊松比":GSI_PR,
            "动态杨氏模量":GSI_DYM,
            "预报结果描述":GSI_RESULT,
            "地下水":GSI_WATER,
            "特殊地质情况":GSI_SGS,
            "稳定性": GSI_STAB,
            "设计围岩级别": GSI_DSCR,
            "推测围岩级别": GSI_ESRG,
            "完整性":GSI_ITGT
        }
    def get_attribute(self,table,i):
        self.get_GSI_LITH(table,i)
        self.get_GSI_RESULT(table,i)
        self.get_GSI_STAB(table,i)
        self.get_GSI_ITGT(table,i)
        self.get_GSI_DSCR(table,i)
        self.get_GSI_PSRL(table,i)


    # 预报结果描述
    def get_GSI_RESULT(self,table,j):
        for col in table.columns:
            if col.cells[0].text.strip().replace("\n", "").replace(" ", "") =="围岩主要工程地质条件"  and  col.cells[1].text.strip().replace("\n", "").replace(" ", "") =="主要工程地质特征":
                cvalue=col.cells[j].text.strip()
                # print(cvalue)
                keywords="风化"
                prewords="岩"
                pre=cvalue.find(prewords)+len(prewords)
                start=cvalue.find("，", pre)+len("，")
                end=cvalue.find("风化", start)
                # print(end)
                if end>0:
                    self.dict["风化程度"]=cvalue[start:end+len(keywords)]
                else:
                    self.dict["风化程度"] = "无"
                self.dict["预报结果描述"]=cvalue
                if self.dict["预报结果描述"]== "":
                    self.dict["预报结果描述"]="无"
                break


    # 岩性
    def get_GSI_LITH(self, table, j):
        for col in table.columns:
            cvalue = col.cells[0].text.strip().replace("\n", "").replace(" ", "")
            if cvalue == "岩性":
                cvalue = col.cells[j].text.strip().replace("\n", "").replace(" ", "")
                # print(self.SB)
                # print(cvalue)
                if not cvalue=="":
                    self.dict["岩性"]=cvalue
                else:
                    self.dict["岩性"]="无"

                break


    # 稳定性
    def get_GSI_STAB(self, table,j):
        for col in table.columns:
            cvalue = col.cells[0].text.strip().replace("\n", "").replace(" ", "")
            if cvalue == "围岩开挖后的稳定状态":
                cvalue = col.cells[j].text.strip().replace("\n", "").replace(" ", "")
                # print(self.SB)
                # print(cvalue)
                if not cvalue == "":
                    self.dict["稳定性"] = cvalue
                else:
                    self.dict["稳定性"] = "无"

                break
    # 完整性
    def get_GSI_ITGT(self,table,j):
        for col in table.columns:
            cvalue = col.cells[0].text.strip().replace("\n", "").replace(" ", "")
            if cvalue == "结构特征和完整状态":
                cvalue = col.cells[j].text.strip().replace("\n", "").replace(" ", "")
                # print(self.SB)
                # print(cvalue)
                if not cvalue == "":
                    self.dict["完整性"] = cvalue
                else:
                    self.dict["完整性"] = "无"

                break
    # 设计围岩级别
    def get_GSI_DSCR(self, table,j):
        for col in table.columns:
            cvalue = col.cells[0].text.strip().replace("\n", "").replace(" ", "")
            if cvalue == "设计围岩级别":
                cvalue = col.cells[j].text.strip().replace("\n", "").replace(" ", "")
                # print(self.SB)
                # print(cvalue)
                if cvalue != "" and cvalue != "/":
                    self.dict["设计围岩级别"] = cvalue[0:cvalue.find("级")]
                else:
                    self.dict["设计围岩级别"] = "无"

                break

    # 推测围岩级别
    def get_GSI_PSRL(self, table,j):
        for col in table.columns:
            cvalue = col.cells[0].text.strip().replace("\n", "").replace(" ", "")
            if cvalue == "预报围岩级别":
                cvalue = col.cells[j].text.strip().replace("\n", "").replace(" ", "")
                # print(self.SB)
                # print(cvalue)
                if cvalue != "" and cvalue != "/":
                    self.dict["推测围岩级别"] = cvalue[0:cvalue.find("级")]
                else:
                    self.dict["推测围岩级别"] = "无"

                break

    # # 地下水状态描述
    # def get_GSI_WATE(self):
    #     return "无"





class Processor(FileProcessBasic):
    def save(self, output, records):
        output_path = os.path.join(output, "TSP_S3S4.csv")
        for record in records:
            header = record.dict.keys()
            util.check_output_file(output_path, header)

            with open(output_path, "a+", encoding="utf_8_sig", newline="") as f:
                w = csv.DictWriter(f, record.dict.keys())
                w.writerow(record.dict)


    def run(self, input_path, output_path):
        files_to_process = set()
        files_to_transform = set()
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
            records = list()
            table2=docx.tables[2]
            table3=docx.tables[3]
            self.get_record_table3(records,table3)
            self.get_record_table2(records,table2)

            self.save(output_path, records)
            print("提取完成" + file)

        for file in files_to_delete:
            if os.path.exists(file):
                os.remove(file)


    def get_record_table2(self,records,table2):
        mileages, vps, vms, prs, dss=self.get_info_table2(table2)
        for record in records:
            inte=record.dict["桩号区间"]
            # print("inte:")
            # print(inte)
            start=inte.find("K")
            if start<0:
                start=inte.find("LS")
                if start >= 0:
                    start = start + len("LS")
                    # print(start)
            else:
                start=start+len("K")


            end=inte.find("～",start)
            max=int(inte[start:end].replace("+",""))
            # print("max:")
            # print(max)
            start=inte.find("K",end)
            if start < 0:
                start = inte.find("LS",end)
                if start >=0:
                    start = start + len("LS")
            else:
                start=start+len("K")

            min=int(inte[start:len(inte)].replace("+",""))
            tvps=[]
            tvms =[]
            tprs =[]
            tdss=[]
            for i in range(0,len(mileages)):

                mileage=mileages[i]

                temp=int(mileage.replace(",","").replace("，",""))
                if temp>=min and temp<=max:
                    tvps.append(vps[i])
                    tvms.append(vms[i])
                    tprs.append(prs[i])
                    tdss.append(dss[i])
            tvps.sort()
            tvms.sort()
            tprs.sort()
            tdss.sort()
            if not len(tvps)==0:
                record.dict["纵波速度"] = tvps[0] + "~" + tvps[len(tvps)-1]
            else:
                record.dict["纵波速度"]="无"
            if not len(tvms)==0:
                record.dict["横波速度"] = str(tvms[0])+ "~" + str(tvms[len(tvms)-1])
            else:
                record.dict["横波速度"]="无"
            if not len(tprs)==0:
                record.dict["泊松比"]=str(tprs[0])+"~"+str(tprs[len(tprs)-1])
            else:
                record.dict["泊松比"]="无"
            if not len(tdss)==0:
                record.dict["密度"]=str(tdss[0])+"~"+str(tdss[len(tdss)-1])
            else:
                record.dict["密度"]="无"

                print(record.dict["纵波速度"])



    def get_info_table2(self,table):
        i=ma=vp=vpm=pr=ds=0
        row=table.rows[0]
        while i< len(table.columns):
            if row.cells[i].text.strip().replace("\n", "").replace(" ", "")=="里程":
                ma=i
            elif row.cells[i].text.strip().replace("\n", "").replace(" ", "")=="Vp(m/s)":
                vp=i
            elif row.cells[i].text.strip().replace("\n", "").replace(" ", "")=="Vp/Vs":
                vpm=i
            elif row.cells[i].text.strip().replace("\n", "").replace(" ", "")=="泊松比":
                pr=i
            elif row.cells[i].text.strip().replace("\n", "").replace(" ", "")=="密度(g/cm3)":
                ds=i
            i=i+1
        mileages = []
        vps = []
        vms = []
        prs = []
        dss = []
        for j in range(2, len(table.rows)):
            row=table.rows[j]
            mileages.append(row.cells[ma].text.strip().replace("-",""))
            # print("mileage:"+row.cells[ma].text.strip().replace("-",""))
            vps.append(row.cells[vp].text.strip().replace(",","").replace("，",""))
            vm=round(float(row.cells[vp].text.strip().replace(",","").replace("，",""))/float(row.cells[vpm].text.strip()),2)
            vms.append(vm)
            prs.append(float(row.cells[pr].text.strip()))
            dss.append(float(row.cells[ds].text.strip()))
        return mileages,vps,vms,prs,dss



    def get_record_table3(self,records,table):
        for col in table.columns:
            cvalue=col.cells[0].text.strip().replace("\n", "").replace(" ", "")
            if cvalue == "预报里程范围":
                # print("预报里程范围")
                for i in range(2, len(col.cells)):
                    cvalue=col.cells[i].text.strip().replace("\n", "").replace(" ", "")
                    # print(cvalue)
                    record=Record(cvalue)
                    record.get_attribute(table, i)
                    records.append(record)
                break



if __name__ == "__main__":
    test = Processor()
    inputpath = "C:/Users/DELL/Desktop/iS3/TSP2"
    outputpath = "C:/Users/DELL/Desktop/iS3"
    test.run(inputpath, outputpath)
