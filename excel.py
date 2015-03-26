#-*- coding: utf8 -*-
import xlrd

class Excel():

	def __init__(self, name):
		self.bk = None
		# 打开文件.
		try:
			self.bk = xlrd.open_workbook(name)
		except Exception, e:
			print(e)

		self.switch_sheet(0)

	# 切换当前读取的sheet
	def switch_sheet(self, index = 0):
		self.sh = self.bk.sheet_by_index(index)

	# sheet的数量
	def get_sheets_num(self):
		return self.bk.nsheets

	# 获取(x, y)的数据.
	def get_cell_value(self, x, y):
		return self.sh.cell_value(x, y)

	# 获取行数
	def get_rows(self, sheet_index = 0):
		return self.sh.nrows

	# 获取列数
	def get_cols(self, sheet_index = 0):
		return self.sh.ncols

	# 获取某行
	def get_row_values(self, row_index):
		return self.sh.row_values(row_index)

	# 获取某列
	def get_col_values(self, col_index):
		return self.sh.col_values(col_index)

	def get_json_keys(self):
		return self.get_row_values(0)
		
	def get_data4json(self):
		keys = self.get_json_keys()
		dic = {}
		for x in xrange(1 , self.get_rows()):
			dic[x] = {}
			row_value = self.get_row_values(x)
			for i in xrange(1, self.get_cols()):
				print row_value[i]
				dic[x][keys[i]] = row_value[i]
			print

		return dic
