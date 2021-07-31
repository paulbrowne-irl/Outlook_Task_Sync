import logging
import os.path
import shutil

import win32com.client
import pandas as pd

from openpyxl import Workbook
from openpyxl import load_workbook
    
'''
Synchonize Outlook Tasks with an Excel file
Look at Readme.md for an overview of what this script does
'''
#Constants
EXCEL_TASK_FILE="task-data.xlsx"
EXCEL_COL_NAMES={
    "Importance":1,
    "Role":2,
    "Categories":3,
    "Subject":4,
    "Team":5,
    "DueDate":6,
    "EntryID":7,
    "CreatedDate":8,
    "Modified":9
    }
LOG_FILE="task.log"

#Handle TO Outlook, Logs and other objects we will need later
OUTLOOK = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
logging.basicConfig(filename=LOG_FILE, encoding='utf-8', level=logging.DEBUG)

'''
Read Task from Excel, update into excel if possible
'''
def read_tasks_into_outlook():
    
    logging.info ("READING TASKS INTO OUTLOOK")

    # Read our Excel tasks into a pandas dataframe
    task_df = pd.read_excel(EXCEL_TASK_FILE, index_col=False)    

    #Iterate through the rows (tasks) in this dataframe
    for index, row in task_df.iterrows():
        print(row)


'''
Clear the tasks output file, so we can reuse the formatting
'''
def clear_excel_output_file():
    logging.info ("CLEARING EXCEL TASK FILE")

    #Make a backup of the original file
    counter =1
    while(os.path.exists(str(counter)+EXCEL_TASK_FILE)):
        logging.debug("Backup file "+str(counter)+EXCEL_TASK_FILE+" exists, increment and try again")
        counter +=1

    shutil.copyfile(EXCEL_TASK_FILE, str(counter)+EXCEL_TASK_FILE)
    logging.debug("Created new backup file:"+str(counter)+EXCEL_TASK_FILE)

    #Open Sheet using Python
    workbook = load_workbook(filename=EXCEL_TASK_FILE)
    sheet = workbook.active

    #Now delete everything until we are only left with the header row
    # continuously delete row 2 until there
    # is only a single row left over 
    # that contains column names 
    while(sheet.max_row > 1):
        # this method removes the row 2
        logging.debug("deleting row")
        sheet.delete_rows(2)

    #Save the result
    workbook.save(filename=EXCEL_TASK_FILE)
    workbook.close


'''
Output Tasks from Outlook Into Excel
'''
def export_tasks_to_excel():
    thisFolder = OUTLOOK.GetDefaultFolder(13)

    folderItems = thisFolder.items
    logging.info ("EXPORTING TASKS TO EXCEL")
    
 
    #Open Excel Sheet using Python
    workbook = load_workbook(filename=EXCEL_TASK_FILE)
    sheet = workbook.active

    for task in folderItems:
        logging.debug("Outputting task:"+task.Subject)
        
        #insert a new clear line (shifting other tasks downwards)
        sheet.insert_rows(2)

        #Update the values
        sheet.cell(row=2,column=EXCEL_COL_NAMES["Importance"]).value=task.Importance
        sheet.cell(row=2,column=EXCEL_COL_NAMES["Role"]).value=task.Role 
        sheet.cell(row=2,column=EXCEL_COL_NAMES["Categories"]).value=task.Categories # make comma safe?
        sheet.cell(row=2,column=EXCEL_COL_NAMES["Subject"]).value=task.Subject # make comma safe?
        sheet.cell(row=2,column=EXCEL_COL_NAMES["Team"]).value=task.TeamTask 
        
        #update Due Date only if it is not default
        tmpDate = str(task.DueDate)
        if(tmpDate!="4501-01-01 00:00:00+00:00"):
            sheet.cell(row=2,column=EXCEL_COL_NAMES["DueDate"]).value=tmpDate

        sheet.cell(row=2,column=EXCEL_COL_NAMES["EntryID"]).value=task.EntryID 
        sheet.cell(row=2,column=EXCEL_COL_NAMES["CreatedDate"]).value=str(task.CreationTime)
        sheet.cell(row=2,column=EXCEL_COL_NAMES["Modified"]).value=str(task.LastModificationTime) 


    #Save the result
    workbook.save(filename=EXCEL_TASK_FILE)

# simple code to run from command line
if __name__ == '__main__':
    
    # Carry out the steps to sync excel adn outlook
    read_tasks_into_outlook()
    #clear_excel_output_file()
    #export_tasks_to_excel()
    
