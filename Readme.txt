


WindowNUTPS - Powershell based Windows NUT client

Does not use NUT auth,  config file needs IP and UPS name updated.  

If any config threashold is hit, shutdown will be triggered. NUT "Low Battery" status triggers a more urgent shutdown


To deploy: 
.\Deploy-PSTask.ps1 -BaseScript "WindowNUTPS" -TargetFolder "NUT"



ToDo: 
 Docs / Help
 Adv config options


Deploy-PSTask - Deployment script for Scheduled Task

1. Creates folder under program files and 
2. Creates logs sub folder
3. Registers Scheduled Task