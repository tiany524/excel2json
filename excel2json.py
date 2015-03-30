# -*- coding: utf8 -*-

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

def get_json_keys(excel):
	return excel.get_row_values(0)

def unic2str(s):
	if type(s) == unicode:
		s = s.encode("utf-8")
	s = str(s) 
	return s

def get_data4json(excel):
	keys = get_json_keys(excel)
	dic = {}
	for uid in xrange(1 , excel.get_rows()):
		primary_key = str(uid)

		dic[primary_key] = {}
		row_value = excel.get_row_values(uid)
		for i in xrange(1, excel.get_cols()):
			key = unic2str(keys[i])
			value = unic2str(row_value[i])
			
			dic[primary_key][key] = value

	return dic


def main():
	for path in get_xls_files():
		excel = Excel(path)
		data = get_data4json(excel)
		
		split_pre = path.find(SUFFIX) - 1
		name = path[:split_pre] + ".json"
		write_json(name, repr(data))

if __name__ == '__main__':
	main()