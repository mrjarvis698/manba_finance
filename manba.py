import os
from os import path
import shutil
import json
import tkinter
from tkinter import filedialog
import pandas as pd
import json
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time

root = tkinter.Tk()
root.withdraw()

# Open xlsx file
open_sheet = path.exists("cache/opened_sheet.json")
if open_sheet == True :
  opened_sheet_file_path = "cache/opened_sheet.json"
  json_file = open(opened_sheet_file_path)
  data = json.load(json_file)
  xlsx_sheet_check = path.exists(data ['xlsx_file_path'])
  if xlsx_sheet_check == True :
    xlsx_file_path = data ['xlsx_file_path']
  else :
    shutil.rmtree('cache', ignore_errors=True)
    xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
    cache_path = os.path.join(str(os.getcwd()), "cache")
    dictionary = {"xlsx_file_path" : xlsx_file_path}
    json_object = json.dumps(dictionary, indent = 1)
    with open("cache/opened_sheet.json", "w") as outfile:
      outfile.write(json_object)
else :
  xlsx_file_path = filedialog.askopenfilename(title="Open Excel-XLSX File")
  cache_path = os.path.join(str(os.getcwd()), "cache")
  os.mkdir(cache_path)
  dictionary = {"xlsx_file_path" : xlsx_file_path}
  json_object = json.dumps(dictionary, indent = 1)
  with open("cache/opened_sheet.json", "w") as outfile:
    outfile.write(json_object)

# read imported xlsx file path using pandas
input_workbook = pd.read_excel(xlsx_file_path, sheet_name = 'Sheet1', usecols = 'E:I', dtype=str)
input_workbook.head()

# read total number of rows present in xlsx
number_of_rows = len(input_workbook.index)

# Opening JSON file & returns JSON object as a dictionary
json_file = open('settings.json')
settings_data = json.load(json_file)

input_workbook_cc_number = input_workbook['Card Number'].values.tolist()
input_workbook_cvv_number = input_workbook['CVV'].values.tolist()
input_workbook_expiry_number = input_workbook['Expiry'].values.tolist()
input_workbook_atm_pin = input_workbook['ATM pin'].values.tolist()
input_workbook_desk_number = input_workbook['Desk'].values.tolist()

# get-output sheet to append output
output_sheet = path.exists("Output.xlsx")
if output_sheet == True :
  output_sheet_file_path = "Output.xlsx"
else :
  output_headers= ['FirstName','LastName', 'Mobile', 'Email','Amount', 'CardNumber', 'CVV', 'Expiry', 'ATM pin', 'No.of Transactions', 'Desk', "Desk Holder"]
  overall_output = Workbook()
  page = overall_output.active
  page.append(output_headers)
  overall_output.save(filename = 'Output.xlsx')
  output_sheet_file_path = "Output.xlsx"

def cal():
  global output_cc_number
  global done_transactions_wb_1
  global h
  output_load_wb_2 = pd.read_excel(output_sheet_file_path, sheet_name = 'Sheet', usecols = 'F', dtype=str)
  output_load_wb_2.head()
  output_cc_number = output_load_wb_2['CardNumber'].values.tolist()
  output_load_wb_1 = pd.read_excel(output_sheet_file_path, sheet_name = 'Sheet', usecols = 'J', dtype=int)
  output_load_wb_1.head()
  done_transactions_wb_1 = output_load_wb_1['No.of Transactions'].values.tolist()
  h = len(output_load_wb_1.index) - 1
  print (output_cc_number[h],done_transactions_wb_1[h])

def cc_expiry():
  global expiry_month
  global expiry_year
  workbook_expiry_month = input_workbook_expiry_number[x]
  workbook_expiry_year = input_workbook_expiry_number[x]
  expiry_month = workbook_expiry_month[:2]
  expiry_year = workbook_expiry_year[3:]

def start_link():
  driver.get("https://mfq.manbafinance.com/paymentwebsite")

def output_save():
  entry_list = [[settings_data['first_name'], settings_data['last_name'], settings_data['registered_mobile_no'], settings_data['email_id'], settings_data['payable_amount'], input_workbook_cc_number[x], input_workbook_atm_pin[x], input_workbook_cvv_number[x], input_workbook_expiry_number[x], z+1, int(input_workbook_desk_number[x]), settings_data["desk_holder"]]]
  output_wb = load_workbook(output_sheet_file_path)
  page = output_wb.active
  for info in entry_list:
      page.append(info)
  output_wb.save(filename='Output.xlsx')

def whole_work():
    start_link()

caps = DesiredCapabilities().CHROME
#caps["pageLoadStrategy"] = "none"
#caps["pageLoadStrategy"] = "eager"
caps["pageLoadStrategy"] = "normal"
driver=webdriver.Chrome(desired_capabilities=caps, executable_path="chromedriver.exe")
driver.maximize_window()
try:
  cal()
except IndexError:
  for x in range (0 , number_of_rows):
    for z in range (0, int(settings_data['number_of_time_transactions_per_card'])):
      whole_work()
else:
  last_txncard =  input_workbook[input_workbook['Card Number'] == output_cc_number[h]].index[0]
  for x in range (last_txncard , number_of_rows):
    for z in range (done_transactions_wb_1[h], int(settings_data['number_of_time_transactions_per_card'])):
      whole_work()
    done_transactions_wb_1[h] = 0

driver.quit()
