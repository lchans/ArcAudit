#!/usr/bin/python
from flask import Flask
import xlrd
import re
app = Flask(__name__)

@app.route('/')
def hello_world():
	executive = xlrd.open_workbook('audit.xlsx')
	sheet = executive.sheet_by_index(0)

	societies = [] 
	students = {} 

	string = ""

	for row in range (0, sheet.nrows): 
		society = sheet.cell(row, 0).value
		students = {}
		column = 4
		while column < sheet.ncols:
			if re.search("\w", sheet.cell(row, column).value):
				students[sheet.cell(row, column).value] = sheet.cell(row, column + 1).value
				societies.append(students)
			column += 4


	for society in societies: 
		for student in society:
			string += society[student]
			string += student
			string += "<br>"

	return string

if __name__ == '__main__':
    app.run()
