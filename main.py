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
def import_objects():
	audit = request.files['audit'].read() 
	members = request.files['members'].read()
	passed, failed, clubs = generate_audit(audit, members)
	return render_template('index.html', passed=passed, failed=failed, clubs=clubs)

def generate_audit(audit, members):
	memberships = ""
	failed = {}
	passed = {}
	clubs = {}

	audit = StringIO(audit) 
	audit = csv.DictReader(audit, delimiter=',') 
	members = StringIO(members) 
	members = csv.DictReader(members, delimiter=',')

	for row in members: 
		group = row["Groups"]
		number = row["Student ID"]
		if re.search("2015 Active Member", group):
			memberships += number

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
	return passed, failed, clubs

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
