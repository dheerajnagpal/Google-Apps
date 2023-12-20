# Google-Apps
This is a group of Google Apps to automate a few tasks on Gmail. 

## Creating the project and copying the script

* To create a new Apps Script project, go to script.new.
* Click Untitled project.
* Rename the Apps Script project to the corresponding project. E.g. Future Delete Messages
* Next to the Code.gs file, click More more_vert > Rename. Rename to the script file. e.g. futureDeletemessages.gs
* Replace the contents of each file with the corresponding code from the file
* Save the script. Select the function of the file to run it once to make sure everything is configured properly and file can be executed. If prompted, authorize the script. 

## Set up a trigger

Once a script is installed and executed, you need to set up a trigger to execute the script automatically. 
* On the project created, on the left tab, choose triggers. 
* Select the method to be executed by trigger, "Head" as the deployment to run. 
* On the event source, select Calendar driven (for items to be executed on calendar changes) or Time driven (for items that need to be executed periodically like hourly, daily, weekly etc.) 
* Select the interval and save the trigger. 
* The executions tab on the left would demonstrate the status of the executions (successful and failed).

