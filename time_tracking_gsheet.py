from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pprint import pprint
import pandas as pd
from googleapiclient.discovery import build
import argparse
import webbrowser


def createAccount(keysfile):
	scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
	credentials = ServiceAccountCredentials.from_json_keyfile_name(keysfile,scope)
	client = gspread.authorize(credentials)
	service = build('sheets', 'v4', credentials=credentials)
	token = credentials.token_uri
	return client, service, token

def generate_graph(credentials_file,title, project, data):
	client, service, token = createAccount(keysfile)
	try:
		gsheet = client.open(title)
	except:
		gsheet = client.create(title)
		email = 'bossdevil62@gmail.com'
		gsheet.share(email, 'user', 'owner')
	try:
		wsh = gsheet.worksheet(project)
	except:
		gsheet.add_worksheet(title=project, rows=1000, cols=1000,index=None)
		wsh = gsheet.worksheet(project)
		wsh.append_row(['User', 'Department', 'Hours'])
	for key in data.keys():
		department = data[key]['department']
		hours = data[key]['hours']
		wsh.append_row([key, department, hours])
	# Create Chart #
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
												        'columnIndex': 4
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
    spreadsheetId=gsheet.id,
    body=request_body
	).execute()
	# Download Spreadsheet as PDF file #
	url = ('https://docs.google.com/spreadsheets/d/' + str(gsheet.id) + '/export?'
       + 'format=pdf'  # export as PDF
       + '&portrait=false'  # landscape
       + '&top_margin=0.00'  # Margin
       + '&bottom_margin=0.00'  # Margin
       + '&left_margin=0.00'  # Margin
       + '&right_margin=0.00'  # Margin
       + '&pagenum=RIGHT'  # Put page number to right of footer
       + '&gid=' + str(wsh.id)  # sheetId
       + '&access_token=' + token)  # access token
	webbrowser.open_new_tab(url)
	##################################################################################################################################
if __name__ == '__main__':
	keysfile = 'D:/scripts/gsheet/key/credentials.json'
	client, service, token = createAccount(keysfile)
	gsheet = client.open('Boss')
	list_data = [
				['user','work type','hour','date'],
				['a','a','a','a']
				]
	
	service.spreadsheets().values().append(
        spreadsheetId = gsheet.id,
        range="Himmapan!A:Z",
        body={
            "majorDimension": "ROWS",
            "values": list_data
        },
        valueInputOption="USER_ENTERED"
    ).execute()