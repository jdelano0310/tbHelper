# tbHelper
## Small utility to update or revert twinBASIC on your system (it is a work in progress)

Features:
* Check for a new version available by reading the page https://github.com/twinbasic/twinbasic/releases
  +  If settings specify to check when the app loads it performs this automatically, else you must click the button to check.
  +  Displays the change log of each version newer than the installed version
* Keep a log of app usage written to a text file called log.text (it writes the contents of the log listbox on the mail form)
  +  View this log view the Log View form
* Revert twinBASIC to a previous version
  +  View Revert form, when loaded the app fills a dropdown with local zip files found in the specified downloads folder (if with in 10 versions of the current version)
  +  Allows the search of Page 1 of the releases GitHub page to give you access to each version a revert can use. (If required, will download the zip file from GitHub to revert twinBASIC)
  +  Displays the change log of the version selected to revert to

***<ins>If you have installed twinBASIC in a Programs Files folder, you will need to run this as admin for it to delete the sub-folders in the twinBASIC folder and then extract the new files from the newly downloaded twinBASIC zip file.</ins>*** 

### Some screenshots

The main form:

![image](https://github.com/user-attachments/assets/943af5b2-3b3c-4402-9a22-15cdc1087b27)

View Log:

![image](https://github.com/user-attachments/assets/7cd11ccc-ff8f-414b-9b52-fa11e4c4aac3)

Use the dropdown to view loags for a specific day

Revert:\
On first load:

![image](https://github.com/user-attachments/assets/8f2bd5fc-5890-4a8c-a2ee-0f3acb5eba8d)

After clicking GitHub button:

![image](https://github.com/user-attachments/assets/15eab249-8577-4805-ac6b-19163f30f584)
When selecting a version to revert to:\
![image](https://github.com/user-attachments/assets/1bc2b461-ae92-469e-b00d-be1ec6151756)\


