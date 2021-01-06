from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pprint import pprint
import pandas as pd
from googleapiclient.discovery import build
import argparse

def createAccount(keysfile):
	scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
	credentials = ServiceAccountCredentials.from_json_keyfile_name(keysfile,scope)
	client = gspread.authorize(credentials)
	service = build('sheets', 'v4', credentials=credentials)
	return client, service

def generate_graph(credentials_file,title, project, data):
	
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
	########################################## Create Chart ###############################################	
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
					                                    'title': 'Hours' #title of y-axis
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
					                                                    'startRowIndex': 0, # set start Row here!
					                                                    'endRowIndex': len(data)+1, # set end Row here!
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
					                                                    'startRowIndex': 0, # set start Row here!
					                                                    'endRowIndex': len(data)+1, # set end Row here!
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
									                                                    'startRowIndex': 0, # set start Row here!
									                                                    'endRowIndex': len(data)+1, # set end Row here!
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
					                        'overlayPosition': {
												    'anchorCell':{
												        'sheetId': wsh.id,
												        'rowIndex': 1,  #position
												        'columnIndex': 6
												    },
												    'offsetXPixels': 0,#scale
												    'offsetYPixels': 0,
												    'widthPixels': 700,
												    'heightPixels': 500
												}	
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
	###################################################################################################################################
parser = argparse.ArgumentParser()
parser.add_argument('-c', dest='credential_file', help='credential file (str)', type=str, default='')
parser.add_argument('-t', dest='title', help='title (str)', type=str, default='')
parser.add_argument('-p', dest='project', help='project (str)', type=str, default='')
parser.add_argument('-d', dest='data', help='data (dict)', type=dict, default='')


if __name__ == '__main__':
	params = parser.parse_args()
	data = {	
	'Boss': {'hours': 8, 'department': 'intern'},
	'Big': {'hours': 8, 'department': 'intern'},
	'Ta': {'hours': 7, 'department':'pipeline'},
	'Nunu': {'hours': 7, 'department':'pipeline'},
	'Pia': {'hours': 7, 'department':'pipeline'},
	'Tor': {'hours': 9, 'department':'pipeline'},
	}
	keysfile = 'D:/scripts/gsheet/key/credentials.json'
	generate_graph(credentials_file=keysfile, title='new', project='wee', data=data)
	#generate_graph(credentials_file=params.credential_file, title=params.title, project=params.project, data=params.data)
