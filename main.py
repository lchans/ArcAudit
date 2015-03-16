"""`main` is the top level module for your Flask application."""

# Import the Flask Framework
from flask import Flask
import xlrd 
import re
from flask import redirect
from flask import url_for
from flask import render_template

app = Flask(__name__)
# Note: We don't need to call run() since our application is embedded within
# the App Engine WSGI application server.

@app.route('/')
def my_form():
    return render_template('main.html')


@app.route('/', methods=['POST'])
def hello():
	executive = xlrd.open_workbook('audit.xlsx')
	members = xlrd.open_workbook('export.xlsx')

	sheet = executive.sheet_by_index(0)
	memberList = members.sheet_by_index(0)

	societies = {}
	fail = {}
	students = {} 
	non = {}
	member = {}

	string = ""
	members = ""

	for row in range (0, memberList.nrows):
		members += str(memberList.cell(row, 3).value)

	for row in range (1, sheet.nrows): 
		society = sheet.cell(row, 0).value
		fail[society] = "Passed"
		students = {}
		column = 4
		while column < sheet.ncols:
			if re.search("\w", sheet.cell(row, column).value):
				students[sheet.cell(row, column).value] = sheet.cell(row, column + 1).value
				societies[society] = students
				member[society] = {}
			column += 4


	for society in societies: 
		for students in societies[society]:
			club = society 
			name = students 
			number = societies[society][students]

			if not re.search(number, members): 
				member[club][name] = number
				fail[club] = "Failed"

	return render_template('index.html', fail=fail, member=member)


@app.errorhandler(404)
def page_not_found(e):
    """Return a custom 404 error."""
    return 'Sorry, Nothing at this URL.', 404


@app.errorhandler(500)
def page_not_found(e):
    """Return a custom 500 error."""
    return 'Sorry, unexpected error: {}'.format(e), 500
