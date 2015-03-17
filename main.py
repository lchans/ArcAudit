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
	audit = StringIO(audit) 
	audit = csv.DictReader(audit, delimiter=',') 
	members = StringIO(members) 
	members = csv.DictReader(members, delimiter=',')




	societies = {}
	fail = {}
	students = {} 
	non = {}
	

	clubs = {}
	memberships = ""

	for row in members: 
		group = row["Groups"]
		number = row["Student ID"]
		if re.search("2015 Active Member", group):
			memberships += number

	failed = {}
	passed = {}
	clubs = {}

	string = "" 


	for row in audit: 
		name = row["Portal"]
		failed[name] = {}
		passed[name] = {}
		clubs[name] = {}
		for response in row: 
			if re.search("Executive [0-9]+: Full Name", response): 
				link = re.match("Executive [0-9]+", response)
				digit =  re.search("[0-9]+", response).group(0)
				full = link.group(0) + ": Full Name"
				number = link.group(0) + ": Student Number" 
				if re.search("\w", row[full]) and re.search("\w", row[number]):	
					clubs[name][row[number]] = row[full]
					if not re.search(row[number], memberships): 
						failed[name][row[number]] = row[full]
					else: 
						passed[name][row[number]] = row[full]


	for club in clubs: 
		name = club 
		string = string + "<h3>" + name + "</h3>"
		for students in failed[name]: 
			string = string + "Failed: " + failed[name][students] + students + "<br>"
		for students in passed[name]: 
			string = string + "Passed: " + passed[name][students] + students + "<br>"

	'''
	if re.search("Student Number", response): 
		if re.search("\w", row[response]):
			clubs[name][]
			number = row[response]
			if not re.search(number, memberships): 
					print number
	'''

	si = StringIO()
	cw = csv.writer(si)
	for row in audit: 
		cw.writerow([row["Submission ID"], row["Portal"]])






	#print members
	return render_template('index.html', passed=passed, failed=failed, clubs=clubs)


@app.route('/download')
def download():
    csv = """"REVIEW_DATE","AUTHOR","ISBN","DISCOUNTED_PRICE"
"1985/01/21","Douglas Adams",0345391802,5.95
"1990/01/12","Douglas Hofstadter",0465026567,9.95
"1998/07/15","Timothy ""The Parser"" Campbell",0968411304,18.99
"1999/12/03","Richard Friedman",0060630353,5.95
"2004/10/04","Randel Helms",0879755725,4.50"""
    # We need to modify the response, so the first thing we 
    # need to do is create a response out of the CSV string
    response = make_response(csv)
    # This is the key: Set the right header for the response
    # to be downloaded, instead of just printed on the browser
    response.headers["Content-Disposition"] = "attachment; filename=books.csv"
    return response

@app.route('/')
def my_form():
    return render_template('main.html')
