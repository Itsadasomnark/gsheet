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

def update_val(worksheet,cell,val):
	worksheet.update(cell,val)

def get_val(worksheet,cell):
	val = worksheet.acell(cell).value

	return val

def get_all_val_dicts(worksheet):
	val_dicts = worksheet.get_all_records()

	return val_dicts

def get_all_val_list(worksheet):
	val_list = sheet.get_all_values()

	return val_list

#scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

#Account = createAccount('D:/googlesheet/key/credentials.json',scope)

#sheet = openByURL(Account,'https://docs.google.com/spreadsheets/d/1c8F2L0ndzgv1qzf5on7-J3t0GHLaQ1f3gznfJ6Ad7AQ/edit?usp=sharing')

#new_worksheet(sheet,'test2',10,20)