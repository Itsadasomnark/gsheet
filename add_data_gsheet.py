from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pprint import pprint
import pandas as pd
from googleapiclient.discovery import build


def createAccount(keysfile):
	scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
	credentials = ServiceAccountCredentials.from_json_keyfile_name(keysfile, scope)
	service = build('sheets', 'v4', credentials=credentials)
	client = gspread.authorize(credentials)
	return client, service

def add_data(keysfile, title, project, data):
	client, service = createAccount(keysfile)
	gsheet = client.open(title)
	service.spreadsheets().values().append(
	        spreadsheetId = gsheet.id,
	        range = "{}!A:Z".format(project),
	        body = {
		            "majorDimension": "ROWS",
		            "values": data
		        	},
		        	valueInputOption="USER_ENTERED"
		    		).execute()

if __name__ == '__main__':
	data = [
			['user', 'work type', 'hour', 'date'],
			['A', 'A', 'A', 'A'],
			['B', 'B' ,'B' ,'B']
			]
	keysfile = 'D:/scripts/gsheet/key/credentials.json'
	title = 'Boss' # gsheet name
	project = 'Himmapan' # worksheet name
	add_data(keysfile, title, project, data)