# coding=utf-8
from docx import Document
import sys
import csv
# sys.path.append("..")
from ..FileProcessBasic import FileProcessBasic

class Processor(FileProcessBasic):
    def __init__(self):
        self.followed=[]
        self.checked=[]
        self.dic={}

    # 查找关键字
    def findKeywords(self):
        # 单元格后紧跟结果所在行，列
        self.followed=[[0,2],[1,0],[1,2]]
        # 单元格后选择结果所在行
        self.checked=[2,13,14,15,16,17]

    # 寻找上下文
    def findContent(self,table):
        # 获取单元格后紧跟结果类型
        for i in range(len(self.followed)):
            row=self.followed[i][0]
            name=self.followed[i][1]
            tmp=list(table.rows[row].cells)
            cols=sorted(set(tmp),key=tmp.index)
            self.dic[cols[name].text.replace(' ','')]=cols[name+1].text

        # 获取单元格后选择结果类型
        for i in self.checked:
            tmp=list(table.rows[i].cells)
            cols=sorted(set(tmp),key=tmp.index)
            name=cols[0].text.replace(' ','')
            for col in cols:
                if col.text.find('√')>0:
                    col.text=col.text.replace('√','')
                    self.dic[name]=col.text
                    break

    # 后续的处理
    def subsequentProcess(self,output):
        for key,value in self.dic.items():
            print(key,' ',value)
        with open(output+'/mycsvfile.csv', 'w',encoding='utf8',newline='') as f:  # Just use 'w' mode in 3.x
            w = csv.DictWriter(f, self.dic.keys())
            w.writeheader()
            w.writerow(self.dic)

    def run(self, inputpath, ouputpath):
        table=Document(inputpath).tables[2]
        self.findKeywords();
        self.findContent(table)
        self.subsequentProcess(outputpath)

if  __name__=="__main__":
    test=FileProcess_GPR_S1S2_Table()
    inputpath="D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/task3/GPRS1S2.docx"
    outputpath="D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/task3"
    test.run(inputpath,outputpath)

