***
The TestingForms App has two buttons
1. Check Downloads button that downloads files to a folder c:\web. Make sure you have this folder created and empty.
2. The Processing Files button that processes the files.
3 THe Application uses the DailyReports connectionstring as with a database called PsecDb.

How to process.

1 Double click the Check Downloads button. It will download files on a monthly basis to the folder c:\web. This button has a method 
AllDatesInMonth(2021, 2)) that takes a year currently 2021 and month being the month of February. Empty files will be downloaded as they will be deleted at processing.
2. The Process Files button will process all files that are downloaded to the download folder. This process will process the files and move them to
 C:\Web\ProcessedFolder where all files processed will be found.

3 Each download will check if the files exists in the ProcessedFolder to avoid multiple downs.

4 The stored procedure was created and tested.

Thanks.