import xlrd
import csv
import re
# from library.FileProcessBasic import FileProcessBasic
import os
import datetime
from xlrd import xldate_as_datetime, xldate_as_tuple


class Record:
	def __init__(self,excel):

		self.table = excel.sheets()[0]

		self.header = {
		"标段号":None,
		"左右幅": None,
		"记录时间": None,
		"桩号区间": None,
		"开挖进度": None,
		"超前支护进度": None,
		"初村进度": None,
		"二村进度": None,
		"关联文件": None,
		}

		self.dataset = []
		self.get_dataset(self.table)

	def out_date(self, year, day):
		fir_day = datetime.datetime(year,1,1)
		zone = datetime.timedelta(days=day-1)
		return datetime.datetime.strftime(fir_day + zone, "%Y/%m/%d")

	def get_dataset(self, table):
		# 记录时间
		PROG_DATE = ""

		for i in range (table.nrows):
			if "保泸高速公路隧道完成情况明细表" in table.cell_value(i,20):
				date = table.cell_value(i,20)
				year = int(date[16:20])
				for k in range (len(date)):
					if date[k] == "周":
						month = int(date[22:k-1])
				day = month * 7
				PROG_DATE = self.out_date(year,day)


			if table.cell_value(i,1) == "" or table.cell_value(i,1) == "标段":
				continue
			
			# 标段号
			PROG_NAME = table.cell_value(i,1)
			print("标段号:" + PROG_NAME)

			# 左幅/右幅/斜井
			PROG_LORR = ""
			if "左幅" in table.cell_value(i,3):
				PROG_LORR = "左幅"
			else:
				if "右幅" in table.cell_value(i,3):
					PROG_LORR = "右幅"
				else:
					if "斜井" in table.cell_value(i,3):
						PROG_LORR = "斜井"
					else:
						continue
			print("左右幅:" + PROG_LORR)

			# 记录时间
			print("记录时间：" + PROG_DATE)

			# 桩号区间
			PROG_END = table.cell_value(i,4)
			print("桩号区间:" + PROG_END)

			data = [PROG_NAME, PROG_LORR, PROG_DATE, PROG_END]
			print(data)
			self.dataset.append(data)


class Processor():

    def save(self, output, record):
        output_path = os.path.join(output, "COPR.csv")
        headers = record.header.keys()

        with open(output_path, "a+", encoding="utf_8_sig", newline="") as f:
        	f_csv = csv.writer(f)
        	f_csv.writerow(headers)
        	f_csv.writerows(record.dataset)

    def run(self, input_path, output_path):
    	'''
    	file_to_process = set()
    	for file in os.listdir(input_path):
    		absolute_file_path = os.path.join(input_path, file)

    		if file.endswith(".xlsx") or file.endswith(".xls"):
    			file_to_process.add(absolute_file_path)

    	for file in file_to_process:
    		excel = xlrd.open_workbook(file)
    		record = Record(excel)

    		self.save(output_path, record)
    		print("提取完成" + file)
    	'''
    	if os.path.isfile(input_path):
    		if input_path.endswith(".xlsx") or input_path.endswith(".xls"):
    			excel = xlrd.open_workbook(input_path)
    			record = Record(excel)

    			self.save(output_path, record)
    			print("提取完成" + input_path)







if __name__ == "__main__":
	input_path = '/Users/budi/Desktop/未命名文件夹/源数据/施工进度/隧道明细.xls'
	output_path = "/Users/budi/Desktop"
	processor = Processor()
	processor.run(input_path, output_path)
	