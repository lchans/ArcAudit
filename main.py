# Import the Flask Framework
import xlrd 
import re
import csv
import io
from flask import Flask
from flask import redirect
from flask import url_for
from flask import render_template
from flask import request
from flask import make_response
from StringIO import StringIO 


app = Flask(__name__)

@app.route('/import', methods= ['POST']) 
def generate_page():
	audit = request.files['audit']
	members = request.files['members']
	if not allowed_file(audit.filename) or not allowed_file(members.filename):
		return render_template('main.html', error="true")
	passed, failed, clubs, memberships = generate_audit(audit, members)
	return render_template('index.html', error ="false", passed=passed, failed=failed, clubs=clubs)

def allowed_file(filename):
	return re.search(".csv", filename)

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
			if count >= 13 and count % 4 == 1 and not re.search("\w", col):
				print row[count-1].lower()
				if re.search(row[count-1].lower(), memberships) and re.search("\w", row[count-1]):
				 	generated_columns.append("Yes")
				elif not re.search(row[count-1].lower(), memberships): 
					generated_columns.append("No")
			else: 
				generated_columns.append(col)
			count = count + 1
		generated_rows.append(generated_columns)

	response = make_response(generate_csv_string(generated_rows))
	response.headers["Content-Disposition"] = "attachment; filename=books.csv"
	return response


def generate_csv_string(document): 
	excel = ""
	for row in document: 
		for col in row: 
			excel = excel + col + ","
		excel = excel +  "\n"
	return excel


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
					if not re.search(row[number].lower(), memberships): 
						failed[name][row[number]] = row[full]
					else: 
						passed[name][row[number]] = row[full]
	return passed, failed, clubs, memberships

@app.route('/')
def my_form():
    return render_template('main.html', error="false")
