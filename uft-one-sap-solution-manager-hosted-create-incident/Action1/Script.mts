'===========================================================
'This script will navigate to the SAP hosted demo instance of solution manager and create a new incident,
'	then delete/withdraw the incident, using AI primarily.  There are cases where traditional OR is used,
'	where the contrast of the text on the button leads to unstable OCR usage, or hidden UI5 tiles cause 
'	the object to have unstable AI recognition.
'You can edit the IncidentPrefix value in the datatable to be any value you would like
'This script was developed and tested ONLY with the EN - English language.  It is very likely that the script
'	will require updates if you chose a different language as the identifying properties are almost exclusively
'	OCR recognized English text.
'The data table has been extended to include data driving which Theme should be used.  Some competitor solutions
'	use image pictures with pixel variance to identify objects.  That methdology would break with different themes.
'	This is to show proof that our solution doesn't break when you change the theme.
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

Dim IncidentNumber, CurrentTime, SAPTheme, Counter

'================================================================================================
'This code will make it so that the script will be able to be run in both 15.0.1 and 15.0.2+ environment
If isempty(micAnyText) and not isempty(micNoText) Then
    micAnyText = micNoText    
End If
'================================================================================================

While Browser("CreationTime:=0").Exist(0)   													'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
SystemUtil.Run "CHROME.exe" ,"","","",3													'launch Chrome, could be data drive to launch other browser (e.g. Firefox)
Set AppContext=Browser("CreationTime:=0")												'Set the variable for what application (in this case the browser) we are acting upon

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")												'Navigate to the application URL
AppContext.Maximize																		'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																			'Wait for the browser to stop spinning
AIUtil.SetContext AppContext																'Tell the AI engine to point at the application

AIUtil("text_box", "User").Type DataTable.Value("UserName")								'Enter the UserName from the datatable
AIUtil("text_box", "Password").Type DataTable.Value("Password")							'Enter the password from the datatable
AIUtil("combobox", "Language").Select DataTable.Value("Language")							'Enter the language from the datatable
AIUtil("button", "Log On").Click																'Click the log on button

if AIUtil("button", "Logon cookie check failed; repeat" + vbLf + "logon").Exist Then			'Sometimes the application will throw an error upon login, typically from being too long on the page, but if it doesn't retry to login
	AIUtil("text_box", "User").Type DataTable.Value("UserName")							'Enter the UserName from the datatable
	AIUtil("text_box", "Password").Type DataTable.Value("Password")						'Enter the password from the datatable
	AIUtil("combobox", "Language").Select DataTable.Value("Language")						'Enter the language from the datatable
	AIUtil("button", "Log On").Click															'Click the log on button
End If

'===========================================================
'Select the theme, cycling through (data driven) all of the themes to show that AI works regardless of theme
'===========================================================
Browser("Home").Page("Home").SAPUIButton("Me Button").Click							'Click the button to bring up the user menu
AIUtil.FindTextBlock("Settings").Click															'Click the Settings Text
AIUtil.FindTextBlock("Appearance").Click														'Click the Appearance text
Browser("Home").Page("Home").WebElement("Theme List Value").Click						'Select the value from the variable, linked in the OR
If AIUtil("button", "Save").Exist Then
	AIUtil("button", "Save").Click
ElseIf AIUtil.FindTextBlock("Save").Exist Then
	AIUtil.FindTextBlock("Save").Click
Else
	Browser("Home").Page("Home").SAPUIButton("Save").Click									'Click the Save button
End If
'15.0.2 Browser("Home").Page("Home").SAPUIButton("More groups").Click							'Click the down arrow to bring up the jump to functionality
AIUtil("down_triangle", micAnyText, micFromBottom, 1).Click
if AIUtil.FindTextBlock("IT Service Management").Exist then
	AIUtil.FindTextBlock("IT Service Management").Click
Else
	'============================================================================================================
	'The below statement was for when the environment had the menu item text changed.
	Browser("Home").Page("Home").SAPUIList("SAPUIList").Select "IT Service Management"		'Click the IT Service Management section @@ script infofile_;_ZIP::ssf1.xml_;_
End If
Browser("Home").Page("Home").SAPUITile("Create Incident Tile").Click						'Click the Create Incident Tile
'=============================================================================================================
'Wait for the correct tile to show up, if it doesn't show up within a minute, abort the script
'=============================================================================================================
Counter = 0
Do
	Counter = Counter + 1
	wait(1)
	If Counter >= 3 Then
		Reporter.ReportEvent micFail, "Wait on Text", "The Incident (IT Service Management) text didn't display within " & Counter & " attempts.  Aborting run"
		Browser("Home").Page("Home").SAPUIButton("Me Button").Click					'Click the button to bring up the user menu
		AIUtil.FindText("Sign Out").Click														'Click the Sign Out text
		Browser("Home").Page("Home").SAPUIButton("OK").Click							'Click the OK button
		While Browser("CreationTime:=0").Exist(0)   											'Loop to close all open browsers
			Browser("CreationTime:=0").Close 
		Wend
		ExitAction
	End If
Loop Until AIUtil.FindTextBlock("Incident (IT Service Management)").Exist
AIUtil.FindTextBlock("Incident (IT Service Management)").Click								'Click the text to create a standard incident type
AppContext.Sync																			'Wait for the browser to stop spinning
IncidentNumber = DataTable.Value("IncidentPrefix")											'Build a custom incident name to ensure it is unique, you can use whatever prefix you want in the datatable to ensure you can find it
CurrentTime = fnRandomNumberWithDateTimeStamp										'Set the variable to be a unique number
IncidentNumber = IncidentNumber & CurrentTime											'Add the prefix from the datatable to the unique number
Browser("Home").Page("Home").WebEdit("Title").Set IncidentNumber						'Type in the unique incident title value
Browser("Home").Page("Home").SAPUIButton("Create").Click								'Click the create button
Browser("Home").Page("Home").SAPUIButton("Withdraw Button").Click						'Click the Withdraw button, the white on blue is not very high contrast, so using traditional OR
Browser("Home").Page("Home").SAPUIButton("Yes").Click									'Click the Yes button
Do																							'Wait for the system to finish processing before going on
	wait(1)	
Loop Until AIUtil.FindText("Withdrawing Incident").Exist(0) = False
'15.0.2 Browser("Home").Page("Home").WebElement("My Incidents").Click							'CLick the My Incidents object
AIUtil.FindTextBlock("My Incidents").Click
AIUtil.FindTextBlock("Home").Click															'Click the Home text
Browser("Home").Page("Home").SAPUIButton("Me Button").Click							'Click the button to bring up the user menu
AIUtil.FindTextBlock("Sign Out").Click														'Click the Sign Out text
Browser("Home").Page("Home").SAPUIButton("OK").Click									'Click the OK button

AppContext.Close																			'Close the application at the end of your script

