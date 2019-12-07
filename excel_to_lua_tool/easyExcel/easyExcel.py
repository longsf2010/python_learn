#!/usr/bin/env python 
# -*- coding: utf-8 -*- 
from win32com.client import Dispatch 
import win32com.client 
class easyExcel: 
	"""A utility to make it easier to get at Excel.  Remembering 
	to save the data is your problem, as is  error handling. 
	Operates on one workbook at a time.""" 
	def __init__(self, filename=None): 
		self.xlApp = win32com.client.Dispatch('Excel.Application') 
		if filename: 
			self.filename = filename 
			self.xlBook = self.xlApp.Workbooks.Open(filename) 
		else: 
			self.xlBook = self.xlApp.Workbooks.Add() 
			self.filename = ''  
	def close(self): 
		self.xlBook.Close(SaveChanges=0) 
		del self.xlApp 

	def getSheet(self,index):
		sht = self.xlBook.Worksheets(index)
		return sht
		

if __name__ == "__main__": 
	xls = easyExcel(r'e:\PythonWork\test.xlsx') 
	xls.close()