import xlrd
import csv
import re
from library.FileProcessBasic import FileProcessBasic
import util
import os


class Record:
	def __init__(self,excel):
		self.table = excel.sheets()[-1]

		self.header = {
		"处理卡编号":None,
		"工程名称": None,
		"里程桩号（变更位置）": None,
		"变更类型": None,
		"原衬砌类别": None,
		"变更后衬砌类别": None,
		"变更原因": None,
		"处理意见": None,
		"处理卡签发日期": None,
		"变更令签发日期": None,
		}
		self.start_row = 0

		self.get_header()
		self.dataset = self.get_dataset()

	# get header position
	def get_header(self):
		n_rows = self.table.nrows
		n_col = self.table.ncols

		for row in range(n_rows):
			cells = self.table.row_values(row)

			if cells[0] != '序号':
				continue

			self.start_row = row + 2
			for col in range(n_col):
				text = cells[col].replace('\n','')
				if text in self.header:
					self.header[text] = col
			break


	# get all row data
	def get_dataset(self):
		n_row = self.table.nrows

		# save all row data
		dataset = []
		for row in range(self.start_row, n_row):
			row_data = self.get_data(row)
			dataset.append(row_data)

		return dataset


	# get row data
	def get_data(self, row):
		data = self.header.copy()
		cells = self.table.row_values(row)

		# get cells directly
		for key in data:
			index = data[key]
			if index != None:
				content = cells[index]
				if isinstance(content, str):
					data[key] = content.replace('\n','')
				else:
					data[key] = content

		# get CHAG_TYPE1, CHAG_TYPE1, CHAG_TYPE1
		text = data["处理意见"]
		text_split = text.split("变更",1)
		if len(text_split) == 2:
			before = re.findall('[a-zA-Z0-9]+',text_split[0])
			after = re.findall('[a-zA-Z0-9]+',text_split[1])

			if len(before) != 0 and len(before[-1]) >= 3 :
				data['原衬砌类别'] = before[-1]
			if len(after) != 0 and len(after[0]) >=3:
				data['变更后衬砌类别'] = after[0]

			pos = text_split[1].find('类型')
			if pos >=2:
				data['变更类型'] = text_split[1][pos-2 : pos]
			else:
				pos = text_split[0].find('类型')
				if pos >= 2:
					data['变更类型'] = text_split[0][pos-2 : pos]

		return data


class Processor(FileProcessBasic):

    def save(self, output, record):
        output_path = os.path.join(output, "CHAG.csv")
        header = record.header.keys()
        util.check_output_file(output_path, header)

        with open(output_path, "a+", encoding="utf_8_sig", newline="") as f:
        	for data in record.dataset:
        		w = csv.DictWriter(f, data.keys())
        		w.writerow(data)

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
	input_path = 'D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/土木数据/数据提取文件示例/源数据/施工变更'
	output_path = "D:/Death in TJU/Junior_2nd/iS3 Lab2/tasks/task3"
	processor = Processor()
	processor.run(input_path, output_path)
	