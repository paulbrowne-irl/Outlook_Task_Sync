
Why would I want to do this?
1) More power than microsoft todo
2) onedrive, bulk edit in excel

Other pluses
* anywhere

Of course not as good - simplcit


Ohter suggestoins
* Save to Onedrive
* cron job to run automatically

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


** Useful technical docs **
https://www.add-in-express.com/creating-addins-blog/2013/06/12/outlook-tasks-create-get-delete/
https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mapifolder?view=outlook-pia
