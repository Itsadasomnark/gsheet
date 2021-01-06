from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pprint import pprint
import pandas as pd
from googleapiclient.discovery import build

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
	service = build('sheets', 'v4', credentials=credentials)
	return client, service

def generate_graph(title, project, data):
	keysfile = 'D:/scripts/gsheet/key/credentials.json'
	
	client, service = createAccount(keysfile)
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
	for key in data.keys():
		department = data[key]['department']
		hours = data[key]['hours']
		wsh.append_row([key, department, hours])
	request_body = {
					    'requests' : [
					        {
					            'addChart':{
					                'chart':{
					                    'spec':{
					                        'title': project, #title of chart
					                        'basicChart' : {
					                            'chartType': 'COLUMN',
					                            'legendPosition' : 'BOTTOM_LEGEND',
					                            'axis': [
					                                # X-AIXS
					                                {
					                                    'position': "BOTTOM_AXIS",
					                                    'title': 'User' #title of x-axis
					                                },
					                                # Y-AIXS
					                                {
					                                    'position': "LEFT_AXIS",
					                                    'title': 'Time' #title of y-axis
					                                }
					                            ],
					                            
					                            #chart lable
					                            'domains': [
					                                {
					                                    'domain': {
					                                        'sourceRange':{
					                                            'sources':[
					                                                {
					                                                    'sheetId': wsh.id,
					                                                    'startRowIndex': 1, # set start Row here!
					                                                    'endRowIndex': 8, # set end Row here!
					                                                    'startColumnIndex': 0, # set start Column here!
					                                                    'endColumnIndex': 1  # set end Column here!
					                                                }
					                                            ]
					                                        },

				                                        }

			                                        }
		                                        ],


				                            	# chart data
					                            'series': [
					                                {
					                                    'series': {
					                                        'sourceRange':{
					                                            'sources':[
					                                                {			
					                                                    'sheetId': wsh.id,
					                                                    'startRowIndex': 1, # set start Row here!
					                                                    'endRowIndex': 8, # set end Row here!
					                                                    'startColumnIndex': 2, # set start Column here!
					                                                    'endColumnIndex': 3  # set end Column here!
					                                                },

					                                            ]
					                                        }    
					                                    },
			                                    		'targetAxis': 'LEFT_AXIS',
			                                    		"dataLabel": {
														    "type": 'CUSTOM',
														    "placement": 'CENTER',
														    "customLabelData": {
																		    'sourceRange':{
									                                            'sources':[
									                                                {			
									                                                    'sheetId': wsh.id,
									                                                    'startRowIndex': 1, # set start Row here!
									                                                    'endRowIndex': 8, # set end Row here!
									                                                    'startColumnIndex': 1, # set start Column here!
									                                                    'endColumnIndex': 2  # set end Column here!
									                                                },

									                                            ]
									                                        }
																		}
														},

					                                }
					                            ],
					                        }
					                    },
					                    'position': {
					                        'newSheet': True
					                    }
					                }
					            }
					        }
					    ]
					}
	response = service.spreadsheets().batchUpdate(
    spreadsheetId=sheet.id,
    body=request_body
	).execute()

if __name__ == '__main__':
	generate_graph('new', 'himmapan', data)
