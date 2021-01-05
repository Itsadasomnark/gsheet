from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pprint import pprint
import pandas as pd
################# install gspread ##################
######## install or upgrade oauth2client ###########

data = {	
	'Boss': {'hours': 8, 'department': 'intern'},
	'Big': {'hours': 8, 'department': 'intern'},
	'Ta': {'hours': 7, 'department':'pipeline'},
	'Nunu': {'hours': 7, 'department':'pipeline'},
	'Pia': {'hours': 7, 'department':'pipeline'},
	'Tor': {'hours': 9, 'department':'pipeline'},
	}

def createAccount(keysfile):
	scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
	credentials = ServiceAccountCredentials.from_json_keyfile_name(keysfile,scope)
	client = gspread.authorize(credentials)
	return client

def generate_graph(title, project, data):
	keysfile = 'D:/scripts/gsheet/key/credentials.json'
	
	client = createAccount(keysfile)
	try:
		sheet = client.open(title)
	except:
		sheet = sheet = client.create(title)
		email = 'bossdevil62@gmail.com'
		sheet.share(email, 'user', 'owner')
	try:
		wsh = sheet.worksheet(project)
	except:
		sheet.add_worksheet(title=project, rows=1000, cols=1000,index=None)
		wsh = sheet.worksheet(project)
		wsh.append_row(['User', 'Department', 'Hours'])
######################add infomation#########################	
	for key in data.keys():
		department = data[key]['department']
		hours = data[key]['hours']
		wsh.append_row([key, department, hours])

if __name__ == '__main__':
	generate_graph('new', 'd', data)
