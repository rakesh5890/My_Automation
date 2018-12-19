#Getting website link as input from excel sheet 1st coloumn, loading it in a website, screenshot of the page and saving it locally
import pyautogui
import time
import webbrowser
import xlrd

loc = input('Provide the excel sheet as location')
savelocation=input('Provide the location to save the file')


wb = xlrd.open_workbook(loc) # To open Workbook 
sheet = wb.sheet_by_index(0)
time.process_time()
for i in range(sheet.nrows):
    url_1 = sheet.cell_value(i,0)
    webbrowser.open_new(url_1)
    time.sleep(10)
    pic = pyautogui.screenshot()
    image_name=url_1[4:15].replace('.','') #fecting only certain words and replacing dot
    pic.save((savelocation+image_name+(str(i))+'.jpg'))
