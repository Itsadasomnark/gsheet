from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pprint import pprint
import pandas as pd

###pip install --upgrade oauth2client PyOpenSSL gspread####

scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name('D:/googlesheet/key/credentials.json',scope)

client = gspread.authorize(credentials)

worksheet = client.open('test2')

sheet = worksheet.worksheet('A worksheet')

#print (sheet.get('A1:B2',))

#worksheet.add_worksheet(title='title', rows=100, cols=20,index=0)

#worksheet.share('bossdevil62@gmail.com', perm_type='user', role='owner')

#w2 = worksheet.get_worksheet(1)

#cell = w2.find("Boss")
#val = w2.acell('A1').value

#w2.update('A1:B2', [[1, 2], [3, 4]])

#dataframe = pd.DataFrame(w2.get_all_records())
#print (dataframe.columns.values.tolist())
#worksheet.remove_permissions('i.somnark@gmail.com')
#print ()

cell_list = sheet.get('B1:B4')
val = sheet.acell('B1').value

print (cell_list,val)