
Sync Sequence
1) Read in Excel Tasks to memory
2) Iterate over *existing* outlook tasks. Where Excel Task is fresher, update the outlook task
    a) Fresher = Same ID, different text, where modified = "Y"
    b) Does not
        i) remove any task in outlook, that has been deleted in excel
        ii) add any task in outlook, that has been added in excel


Files
* Outlook.py - the actual sync script
* task-data.xlsx - the file that syncs with outlook


Note about how to use with OneDrive