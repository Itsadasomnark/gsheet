#from Google import Create_Service

CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
spreadsheet_id = '11asdfl3sd5DSDLEADFIHY73f6e3gd-x3YadlkjY7C8'
sheet_id = '651526025' #put sheet id 
"""
single Bar Chart 
"""
request_body = {
    'requests' : [
        {
            'addchart':{
                'chart':{
                    'spec':{
                        'title': 'title of chart', #title of chart
                        'basicChart' :{
                            'chartType': 'COLUMN',
                            'legendPosition' : 'BOTTOM_LEGEND'
                            'axis': [
                                # X-AIXS
                                {
                                    'position': "BOTTOM_AXIS"
                                    'title': 'title of X-AIXS' #title of x-axis
                                },
                                # Y-AIXS
                                {
                                    'position': "LEFT_AXIS"
                                    'title': 'title of Y-AIXS' #title of y-axis
                                }
                            ],

                            'series': [
                                {
                                    'series': {
                                        'sourcesRange':{
                                            'sources':[
                                                {
                                                    'sheetId': sheet_id,
                                                    'startRowIndex': 0, # set start Row here!
                                                    'endRowIndex': 11, # set end Row here!
                                                    'startColumnIndex': 1, # set start Column here!
                                                    'endColumnIndex': 2  # set end Column here!
                                                }
                                            ]
                                        }
                                    },
                                    'targetAxis': 'LEFT_AXIS'
                                }
                            ]
                        }
                    },
                    'position': {
                        ''
                    }
                }
            }
        }
    ]
}

response = service.spreadsheets().batchUpate(
    spreadsheetId=spreadsheet_id,
    body=request_body
).execute()


                                            ,###########################################

                                                #data lable
                                                'totalDataLabel' : {
                                                    'type' : 'DATA',
                                                    'placement' : 'CENTER',
                                                    'customLabelData' : {
                                                        'sourceRange' :{
                                                            'sources':[
                                                                {
                                                                    'sheetId': wsh.id,
                                                                    'startRowIndex': 1, # set start Row here!
                                                                    'endRowIndex': 8, # set end Row here!
                                                                    'startColumnIndex': 2, # set start Column here!
                                                                    'endColumnIndex': 3  # set end Column here!
                                                                }
                                                            ]

                                                        }
                                                    }
                                                }#############################################