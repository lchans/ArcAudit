# Import the Flask Framework
from flask import Flask
import xlrd 
import re
from flask import redirect
from flask import url_for
from flask import render_template
from flask import request
from flask import make_response
import csv
import io
from StringIO import StringIO 


app = Flask(__name__)

@app.route('/import', methods= ['POST']) 
def generate_page():
	audit = request.files['audit']
	members = request.files['members']
	passed, failed, clubs, memberships = generate_audit(audit, members)
	return render_template('index.html', passed=passed, failed=failed, clubs=clubs)

@app.route('/generate', methods= ['POST'])
def generate_csv():
	audit = request.files['audit']
	members = request.files['members']
	audit = StringIO(audit.read()) 
	audit = csv.reader(audit, delimiter=',') 
	members = StringIO(members.read()) 
	members = csv.DictReader(members, delimiter=',')

	generated_rows = [] 
	generated_columns = []
	rows = [] 
	columns = []

	memberships = generate_membership_list(members)

	excel = " "
	for row in audit: 
		count = 0
		generated_columns = []
		for col in row: 
			arc = count % 4
			if count >= 13 and arc == 1 and not re.search("\w", col):
				if re.search(row[count-1], memberships) and re.search("\w", row[count-1]):
				 	generated_columns.append("Yes")
				elif not re.search(row[count-1], memberships): 
					generated_columns.append("No")
			else: 
				generated_columns.append(col)
			count = count + 1
		generated_rows.append(generated_columns)

	for row in generated_rows: 
		for col in row: 
			excel = excel + col + ","
		excel = excel +  "\n"

	response = make_response(excel)
	response.headers["Content-Disposition"] = "attachment; filename=books.csv"
	return response


def generate_membership_list (members):
	memberships = ""
	for row in members: 
		group = row["Groups"]
		number = row["Student ID"]
		if re.search("2015 Active Member", group):
			memberships += number
	return memberships

def generate_audit(audit, members):
	memberships = ""
	failed = {}
	passed = {}
	clubs = {}

	audit = StringIO(audit.read()) 
	audit = csv.DictReader(audit, delimiter=',') 
	members = StringIO(members.read()) 
	members = csv.DictReader(members, delimiter=',')

	memberships = generate_membership_list (members)

	for row in audit: 
		name = row["Portal"]
		failed[name] = {}
		passed[name] = {}
		clubs[name] = {}
		for response in row: 
			if re.search("Executive [0-9]+: Full Name", response): 
				digit = re.search("[0-9]+", response).group(0)
				full = "Executive " + digit + ": Full Name"
				number = "Executive " + digit + ": Student Number"
				if re.search("\w", row[full]) and re.search("\w", row[number]):	
					clubs[name][row[number]] = row[full]
					if not re.search(row[number], memberships): 
						failed[name][row[number]] = row[full]
					else: 
						passed[name][row[number]] = row[full]
	return passed, failed, clubs, memberships

@app.route('/download', methods= ['POST'])
def download():
	si = StringIO()
	cw = csv.writer(si)
	for row in request.form['clubs']: 
		cw.writerow([row["Submission ID"], row["Portal"]])
	response = make_response(cw)
	response.headers["Content-Disposition"] = "attachment; filename=books.csv"
	return response

@app.route('/')
def my_form():
    return render_template('main.html')
