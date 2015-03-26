#-*- coding: utf8 -*-

from excel import Excel
import os, json


SUFFIX = "xls"

# Todo: 通过配置读取excel和导出.
# class Setting():
# 	def __init__(self, path):
# 		fp = open(path, "r")
# 		print fp
# 		# self.data = json.read(fp.read())
# 		fp.close()
# 
# setting = Setting("config.json")


def get_xls_files(path = "."):
	files = []
	for cell in os.listdir(path):
		if SUFFIX in cell:
			files.append(cell)

	return files

def write_json(file_name, data):
	encodedjson = json.dumps(data, sort_keys=True)
	# open json file
	fp = open(file_name, "w")
	fp.write(encodedjson)
	fp.close()


def main():
	for cell in get_xls_files():
		excel_data = Excel(cell).get_data4json()
		
		split_pre = cell.find(SUFFIX) - 1
		name = cell[:split_pre] + ".json"
		write_json(name, repr(excel_data))
		

if __name__ == '__main__':
	main()