'===========================================================
'This script will navigate to the SAP hosted demo instance of solution manager and create a new incident,
'	then delete/withdraw the incident, using AI primarily.  There are cases where traditional OR is used,
'	where the contrast of the text on the button leads to unstable OCR usage, or hidden UI5 tiles cause 
'	the object to have unstable AI recognition.
'You can edit the IncidentPrefix value in the datatable to be any value you would like
'This script was developed and tested ONLY with the EN - English language.  It is very likely that the script
'	will require updates if you chose a different language as the identifying properties are almost exclusively
'	OCR recognized English text.
'===========================================================


'===========================================================
'Function for creating a number at run time based on current time down to the second, to allow for a unique number each time the script is run
'===========================================================

Function fnRandomNumberWithDateTimeStamp()

'Find out the current date and time
Dim sDate : sDate = Day(Now)
Dim sMonth : sMonth = Month(Now)
Dim sYear : sYear = Year(Now)
Dim sHour : sHour = Hour(Now)
Dim sMinute : sMinute = Minute(Now)
Dim sSecond : sSecond = Second(Now)

'Create Random Number
fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)

'======================== End Function =====================
End Function

Dim IncidentNumber

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
SystemUtil.Run "CHROME.exe" ,"","","",3														'launch Chrome, could be data drive to launch other browser (e.g. Firefox)
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

AIUtil("text_box", "User").Type DataTable.Value("UserName")									'Enter the UserName from the datatable
AIUtil("text_box", "Password").Type DataTable.Value("Password")								'Enter the password from the datatable
AIUtil("combobox", "Language").Select DataTable.Value("Language")							'Enter the language from the datatable
AIUtil("button", "Log On").Click															'Click the log on button

if AIUtil("button", "Logon cookie check failed; repeat" + vbLf + "logon").Exist Then		'Sometimes the application will throw an error upon login, typically from being too long on the page, but if it doesn't retry to login
	AIUtil("text_box", "User").Type DataTable.Value("UserName")
	AIUtil("text_box", "Password").Type DataTable.Value("Password")
	AIUtil("combobox", "Language").Select DataTable.Value("Language")
	AIUtil("button", "Log On").Click
End If

AIUtil("down_triangle", micNoText, micFromTop, 1).Click										'Click the down triangle to bring up the navigation menu
Browser("Home").Page("Home").WebElement("IT Service Management").Click						'Click on the IT Service Management link, isn't always displayed in the list visibly, so use traditional OR to just click no matter where it is
AIUtil.FindTextBlock("Create Incident").Click												'Click the Create Incident menu item in the right part of the menu
AIUtil.FindTextBlock("Select incident type").Exist											'Wait for the text to show up
AIUtil.FindTextBlock("SMIN").Click															'Click the text to create a standard incident type
IncidentNumber = DataTable.Value("IncidentPrefix")											'Build a custom incident name to ensure it is unique, you can use whatever prefix you want in the datatable to ensure you can find it
IncidentNumber = IncidentNumber & fnRandomNumberWithDateTimeStamp
AIUtil("text_box", "*Title:").Type IncidentNumber											'Enter the unique incident name
AIUtil("button", "").Click																	'Click the submit button

If AIUtil.FindTextBlock(IncidentNumber).Exist = False Then									'Use this if you need to write to the reporter or some other notification of a failure to create the incident
	msgbox "Something is wrong"
End If

Browser("Home").Page("Home").SAPUIButton("Withdraw Button").Click							'Click the Withdraw button, the white on blue is not very high contrast, so using traditional OR
AIUtil.FindTextBlock("Yes").Click															'Click the Yes text
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application
'Navigate to the home page, start and stop from the same place to make it can iterate
If AIUtil.FindTextBlock("Home").Exist = False Then											'Sometimes the click on the application doesn't register, because the application is still processing
	AIUtil.FindTextBlock("My Incidents").Click													'Click the menu area to bring up the navigation menu
End If
AIUtil.FindTextBlock("Home").Click															'Click the Home text
Browser("Home").Page("Home").SAPUIButton("Me Button").Click									'Click the button to bring up the user menu
AIUtil.FindText("Sign Out").Click															'Click the Sign Out text
Browser("Home").Page("Home").SAPUIButton("OK").Click										'Click the OK button

AppContext.Close																			'Close the application at the end of your script

