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

'===========================================================
'Application script goes here
'===========================================================
AIUtil("text_box", "User").Type DataTable.Value("UserName")
AIUtil("text_box", "Password").Type DataTable.Value("Password")
AIUtil("combobox", "Language").Select DataTable.Value("Language")
AIUtil("button", "Log On").Click
if AIUtil("button", "Logon cookie check failed; repeat" + vbLf + "logon").Exist Then
	AIUtil("text_box", "User").Type DataTable.Value("UserName")
	AIUtil("text_box", "Password").Type DataTable.Value("Password")
	AIUtil("combobox", "Language").Select DataTable.Value("Language")
	AIUtil("button", "Log On").Click
End If

AIUtil("down_triangle", micNoText, micFromTop, 1).Click
Browser("Home").Page("Home").WebElement("IT Service Management").Click
AIUtil.FindTextBlock("Create Incident").Click
AIUtil.FindTextBlock("Select incident type").Exist


Browser("Home").Page("Home").SAPUIButton("Me Button").Click
AIUtil.FindText("Sign Out").Click
AIUtil.FindTextBlock("OK").Click


AppContext.Close																			'Close the application at the end of your script

