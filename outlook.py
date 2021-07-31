import win32com.client
import pandas as pd
import logging

'''
@TODO readme of how this works

'''
#Constants
EXCEL_TASK_FILE="task-data.xlsx"
LOG_FILE="task.log"

#Handle TO Outlook, Logs and other objects we will need later
OUTLOOK = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
logging.basicConfig(filename=LOG_FILE, encoding='utf-8', level=logging.DEBUG)

'''
Importance	Role	Categories	Subject	Team	EntryID	DueDate	CreatedDate	Modified

'''

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

'''
Output Tasks from Outlook Into Excel
'''
def export_tasks_to_excel():
    thisFolder = OUTLOOK.GetDefaultFolder(13)

    folderItems = thisFolder.items
    logging.info ("EXPORTING TASKS TO EXCEL")
    logging.debug("number of tasks"+str(folderItems.count))
 
    for task in folderItems:
        logging.debug(task.Subject)


# simple code to run from command line
if __name__ == '__main__':
    
    # Carry out the steps to sync excel adn outlook
    read_tasks_into_outlook()
    clear_excel_output_file()
    export_tasks_to_excel()
    


# Next open xl file

