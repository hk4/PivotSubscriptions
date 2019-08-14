# PivotSubscriptions

<p align="center"><img src="https://github.com/hk4/PivotSubscriptions/blob/dev/Assets/PivotSubscriptions_BookCover.png"></p>

PivotSubsctions allows users to subscribe to a pivot table and recieve a scheduled email containing the specified pivot table. 

## Accessing
Once installation is complete, a new "Pivot Subscriptions" Link will be added to the InterSystems IRIS Business Intelligence User Portal. The User Portal can be found at the Management Portal -> Analytics -> User Portal.

## Installation
1. Use the Download ZIP option for this project
2. Extract the files and copy path
	* This is the path to the directory that contains README.md and LICENSE
3. Open terminal and ZN to desired namespace
4. Run the following commands:
```
	set path="<PATH FROM STEP 2>"
	do $system.OBJ.Load(path_"/PivotSubscriptions/Installer.cls","ck",,1)
	do ##class(PivotSubscriptions.Installer).RunInstaller(path)
```
5. Follow the Configuration steps

## Configuration steps
### Enabling Analytics on InterSystems IRIS
To ensure that links to the SubscriptionManager page are correctly generated, the Default Application of the namespace you are working in is correctly configured. 
1. Go to Management Portal
2. Go to System Administration > Security > Applications > Web Applications
3. Click on  /csp/??? where ??? is the namespace you are working in
4. Check the checkbox next to Namespace Default Application and Save
You can see more information in InterSystems IRIS Business Intelligence's <a href="https://docs.intersystems.com/irislatest/csp/docbook/DocBook.UI.Page.cls?KEY=D2IMP_setup_web_app">documentation</a>.

### Configure Task Manager Email Settings
Subscriptions are delivered by Email. The Task Manager Email must be configured to allow alerts to be delivered by Email. At a minimum, the SMTP Server must be assigned in the Task Manager Email Settings (Management Portal -> System Administration -> Configuration -> Additional Settings -> Task Manager Email). For more information, please see the <a href="http://docs.intersystems.com/irislatest/csp/docbook/DocBook.UI.Page.cls?KEY=RACS_Category_TaskManagerEmail">documentation</a>.

### Setting the Configuration Page
The Configuration Page enables users to set an email address for error reporting and for responses to the email subscriptions. 

In addition, the link can be manually configured by setting:
1. Whether the link should be using HTTP or HTTPS
2. The server address
3. The port number

If a web server is specified without a port, the port number will be removed from the generated link. Thus, the application's link can be referenced through a web server.
