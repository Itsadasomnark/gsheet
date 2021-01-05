from oauth2client.service_account import ServiceAccountCredentials
import gspread
from pprint import pprint
import pandas as pd
################# install gspread ##################
######## install or upgrade oauth2client ###########


##############################------SCOPE------###################################
'''	sheets readonly      = 'https://www.googleapis.com/auth/spreadsheets.readonly'
	sheets read/write    = 'https://www.googleapis.com/auth/spreadsheets'
	drive readonly       = 'https://www.googleapis.com/auth/drive.readonly'
	drive created/opened = 'https://www.googleapis.com/auth/drive.file'
	drive full access    = 'https://www.googleapis.com/auth/drive' '''
##################################################################################

def createAccount(keysfile,scope):
	credentials = ServiceAccountCredentials.from_json_keyfile_name(keysfile,scope)
	client = gspread.authorize(credentials)

	return client

def openByTitle(client,title):
	sheet = client.open(title)

	return sheet

def openByURL(client,url):
	sheet = client.open_by_url(url)

	return sheet

def createNewSheet(client,name):
	sheet = client.create(name)

	return sheet

def shareSheet(sheet,email,account_type,role):
	############## account_type #################
	      #user,group,domain,anyone
   	################### role ####################
   			#owner, writer, reader
	sheet.share(email, perm_type=account_type, role=role)

def list_permissions(sheet):
	permissions = sheet.list_permissions()

	return permissions

def remove_permissions(sheet,email):
	sheet.remove_permissions(email)

def select_worksheetByIndex(sheet,index):
	worksheet = sheet.get_worksheet(index)

	return worksheet	

def select_worksheetByTitle(sheet,title):
	worksheet = sheet.worksheet(title)

	return worksheet

def new_worksheet(sheet,title,row,column,index=None):
	sheet.add_worksheet(title=title, rows=row, cols=column,index=index)

def update_val(worksheet,acell,val):
	worksheet.update(acell,val)

def update_val_cell(worksheet,row,column,val):
	worksheet.update_cell(row, column, val)

def get_val_acell(worksheet,acell):
	val = worksheet.get(acell)

	return val

def get_val_cell(worksheet,row,column):
	val = worksheet.cell(row, column).value

	return val

def get_all_val_dicts(worksheet):
	val_dicts = worksheet.get_all_records()

	return val_dicts

def get_all_val_list(worksheet):
	val_list = worksheet.get_all_values()

	return val_list

def get_col_val(worksheet,col):
	values_list = worksheet.col_values(col)

	return values_list
def append_row(worksheet,val):
	worksheet.append_row(val)

def finditem(worksheet,val):
	list_cell = worksheet.findall(val)

	return list_cell

def add_column(worksheet,val):
	worksheet.add_cols(val)

def add_row(worksheet,val):
	worksheet.add_rows(val)

def list_worksheet(sheet):
	ws_list = sheet.worksheets()
	title_list = []

	for wsh in ws_list:
		title_list.append(wsh.title)
	return title_list

def test():
	scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

	keysfile = 'D:/googlesheet/key/credentials.json'

	client = createAccount(keysfile,scope)

	sheet = openByURL(client, 'https://docs.google.com/spreadsheets/d/1c8F2L0ndzgv1qzf5on7-J3t0GHLaQ1f3gznfJ6Ad7AQ/edit?usp=sharing')

	sheet2 = openByTitle(client, 'test')

	sheet3 = createNewSheet(client, 'test3')

	email = 'bossdevil62@gmail.com'
	shareSheet(sheet3, email, 'user', 'owner')

	premissions = list_permissions(sheet)

	remove_permissions(sheet2, email) #only owner

	ws = select_worksheetByIndex(sheet, 0)

	ws2 = select_worksheetByTitle(sheet, 'A worksheet')

	ws3 = new_worksheet(sheet, 'New Worksheet', 10, 10)

	update_val(ws, 'A1', 'Itsada')

	update_val_cell(ws,1,2,'Somnark')

	val = get_val_acell(ws, 'A1')

	val2 = get_val_cell(ws, 1, 1)

	val3 = get_all_val_dicts(ws)

	val4 = get_all_val_list(ws)


	info = [100, 200, 300, 400]
	append_row(ws, info)

	add_column(ws, 5)

	add_row(ws, 5)

	list_cell = finditem(ws, 'Itsada')
	row = list_cell[0].row
	column = list_cell[0].col

def add_record(project, user, path, duration=15):
	from datetime import date
	scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
	keysfile = 'D:/googlesheet/key/credentials.json'
	client = createAccount(keysfile, scope)
	sheet = openByURL(client, 'https://docs.google.com/spreadsheets/d/1fcSg0yy5g5Dx9oBnYD4Rdj5Hl9WtjBj5l_fyutcKDZk/edit#gid=0')
	
	today = date.today()
	date = today.strftime("%d/%m/%Y")

	ws_list = list_worksheet(sheet)
	if project not in ws_list:
		new_worksheet(sheet, project, 10, 10)
		ws = select_worksheetByTitle(sheet, project)
		append_row(ws,['Date', 'User', 'Duration', 'Path'])
	else:
		ws = select_worksheetByTitle(sheet, project)

	user_old = get_col_val(ws, 2)
	if user not in user_old:
		append_row(ws,[date, user, duration, path])
	else:
		user_cell = finditem(ws, user)
		old_date_list = {}
		for cell in user_cell:
			old_date = get_val_acell(ws,'A{}'.format(cell.row))
			old_date_list[old_date[0][0]] = cell
		
		if date not in old_date_list.keys():
			append_row(ws,[date, user, duration, path])
		else:
			for old in old_date_list.keys():
				if date == old:
					current_cell = old_date_list[old]
					old_dul = get_val_cell(ws, current_cell.row, 3)
					new_dul = int(old_dul)+duration
					update_val_cell(ws, current_cell.row, 3, new_dul)
					old_path = get_val_cell(ws, current_cell.row, 4)
					old_path_list = old_path.split(',')
					if path not in old_path_list:
						new_path = old_path + ',{}'.format(path)
						update_val_cell(ws, current_cell.row, 4, new_path)

