#!/usr/bin/env python 
 # -*- coding: utf-8 -*- 
from easyExcel import easyExcel 
def parseExcelToLua(src_path,des_path):
	book = easyExcel.easyExcel(src_path)
	sheet = book.getSheet(1)
	info = sheet.UsedRange
	column = info.Columns.Count
	row =info.Rows.Count
	file_path = des_path
	header_row = 1
	type_row =2
	header_list = []
	'''number string '''
	type_list = []
	value_list = []
	for i in range(1,row+1):
		single_value_list = []
		is_wrong = False
		for j in range(1,column+1):
			temp_value = sheet.Cells(i,j).Value
			if i==header_row :
				header_list.append(temp_value)
			elif i==type_row:
				if temp_value !="number" and temp_value !="string" :
					print("类型错误")
					is_wrong = True
					break
				type_list.append(temp_value)
			else: 
				single_value_list.append(temp_value)
		if is_wrong :
			break
		if i>=3 :
			value_list.append(single_value_list)	
		
	book.close()
	if len(type_list) != len(header_list):
		print("格式错误")
		return
		
	fo = open(file_path,"w")
	fo.write('local t ={\n')
	for i in range(0,len(value_list)):
		fo.write('\t[%d]={\n'%(i))
		for j in range(0,len(header_list)):
			content = parseType(type_list[j],value_list[i][j])
			fo.write('\t\t{0}={1},'.format(header_list[j],content))
			fo.write('\n')
		fo.write('\t},\n')	
	fo.write('}\nreturn t')
	fo.close()

def parseType(type,value):
	if type=="number" :
		if value =='None' :
			return '0'
		template = '{0}'
	else :
		if value =='None':
			value =''
		template ='\"{0}\"'
	return template.format(value)

	
if __name__ == "__main__": 
	src_path =r'E:\python_learn_bran\python_learn\excel_to_lua_tool\test.xlsx'
	des_path =r'E:\python_learn_bran\python_learn\excel_to_lua_tool\Sheet1.lua'
	parseExcelToLua(src_path,des_path)		
	