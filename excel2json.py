# -*- coding: utf8 -*-

from excel import Excel
import os, json

SUFFIX = "xls"
DEFAULT_PATH = "./"

class Setting():

	def __init__(self, path):
		self.data = self.load(path)
		self.create_folder(self.work_dir())
		self.create_folder(self.output_dir())

	def load(self, path):
		fp = open(path, "r")
		content = fp.read()
		fp.close()
		return json.loads(content)

	def create_folder(self, path):
		if os.path.exists(path):
			return
		else:
			os.mkdir(path)

	def work_dir(self):
		return self.data["work_dir"] + "/" or DEFAULT_PATH

	def output_dir(self):
		return self.data["output_dir"] + "/" or DEFAULT_PATH


setting = Setting("config.json")


def get_xls_files(path = "."):
	files = []
	for cell in os.listdir(path):
		if SUFFIX in cell:
			files.append(cell)

	return files


def write_json(file_name, data):
	file_name = setting.output_dir() + file_name
	print file_name

	fp = open(file_name, "w")
	encodedjson = json.dumps(data, sort_keys=True)
	fp.write(encodedjson)
	fp.close()


def unic2str(s):
	if type(s) == unicode:
		s = s.encode("utf-8")
	return str(s)


def get_data4json(excel):
	keys = excel.get_row_values(0)
	Ids  = excel.get_col_values(0)

	dic = {}
	for uid in xrange(1 , excel.get_rows()):
		primary_key = str(int(Ids[uid]))

		dic[primary_key] = {}
		row_value = excel.get_row_values(uid)
		for i in xrange(1, excel.get_cols()):
			key   = unic2str(keys[i])
			value = unic2str(row_value[i])
			
			dic[primary_key][key] = value

	return dic


def main():
	for path in get_xls_files(setting.work_dir()):
		excel = Excel(setting.work_dir() + path)
		data = get_data4json(excel)
		
		split_pre = path.find(SUFFIX) - 1
		name = path[:split_pre] + ".json"

		write_json(name, data)


if __name__ == '__main__':
	main()
