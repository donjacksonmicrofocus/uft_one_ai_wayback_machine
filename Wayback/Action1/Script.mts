'Declare variables that will be used in the script
Dim BrowserExecutable, oShell, CurrentYear, CurrentMonth, CurrentURL, Iterator

'Statements to ensure that the OCR service that the AI Object Detection (AIOD) utilizes is running on the machine
Set oShell = CreateObject ("WSCript.shell")
oShell.run "powershell -command ""Start-Service mediaserver"""
Set oShell = Nothing

'Loop to close all open browsers
While Browser("CreationTime:=0").Exist(0)   													
	Browser("CreationTime:=0").Close 
Wend

'Set the BrowserExecutable variable to be the .exe for the browser declared in the datasheet
BrowserExecutable = Parameter.Item("BrowserName") &  ".exe"

'Launch the browser specified in the data table
SystemUtil.Run BrowserExecutable,"","","",3												

'Set the variable for what application (in this case the browser) we are acting upon
Set AppContext=Browser("CreationTime:=0")												

'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.ClearCache																		

'Navigate to the application URL
AppContext.Navigate "about:blank"

'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Maximize																		

'Wait for the browser to stop spinning
AppContext.Sync																			

'Tell the AI engine to point at the application
AIUtil.SetContext AppContext																
'https://web.archive.org/web/200601/http://facebook.com/

'===========================================================================================
'BP:  Login
'===========================================================================================
CurrentYear = Parameter.Item("StartYear")
CurrentMonth = Parameter.Item("StartMonth")

For Iterator = 1 To Parameter.Item("NumberOfIterations") Step 1
	If CurrentMonth <= 9 Then
		CurrentURL = "https://web.archive.org/web/" & CurrentYear & "0" & CurrentMonth & "/" & Parameter.Item("URL")
	Else
		CurrentURL = "https://web.archive.org/web/" & CurrentYear & CurrentMonth & "/" & Parameter.Item("URL")
	End If
	AppContext.Navigate CurrentURL
	AppContext.Sync	
	If AIUtil.FindText("Utilisez").Exist(0) = True Then
		Reporter.ReportEvent micWarning, "Navigating to " & CurrentURL, "The WayBack Machine loaded the French version of the page."
	Else
		AIUtil("text_box", "E-mail:").Highlight	
		AIUtil.Context.Freeze 
		AIUtil("text_box", "E-mail:").SetText "user@domain.com"
		If AIUtil("text_box", "Password:").Exist(0) = True Then
			AIUtil("text_box", "Password:").Highlight
			AIUtil("text_box", "Password:").SetText "Password"
			Set PasswordField = AIUtil("text_box", "Password:")
		ElseIf AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")) = True Then	
			AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")).Highlight
			AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")).SetText "Password"
			Set PasswordField = AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password"))
		Else
			msgbox "Can't find Password field"
		End If
		AIUtil("button", micAnyText, micWithAnchorAbove, PasswordField).Highlight
		AIUtil("button", micAnyText, micWithAnchorAbove, PasswordField).CheckExists True
		'https://web.archive.org/web/200703/http://facebook.com/ has a problem, looks like it's identifying the Register box as the login?
	'	If AIUtil("button", "Login").Exist(0) = True Then
	'		Set LoginProperties = AIUtil("button", "Login").GetAllProperties
	'		For i = 0 To LoginProperties.count - 1
	'			print LoginProperties.keys()(i) & ":" & LoginProperties(LoginProperties.keys()(i))
	'		Next
	'		AIUtil("button", "Login").Highlight
	'		AIUtil("button", "Login").CheckExists True
	'	ElseIf AIUtil("button", "Login", micFromTop, 1).Exist(0) = True Then
	'		AIUtil("button", "Login", micFromTop, 1).Highlight
	'		AIUtil("button", "Login", micFromTop, 1).CheckExists True
	'	Else
	'		AIUtil("button", "Login", micWithAnchorAbove, AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password"))).Highlight
	'		AIUtil("button", "Login", micWithAnchorAbove, AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password"))).CheckExists True
	'	End If
	
	End If	
	
	CurrentMonth = CurrentMonth + 1
	If CurrentMonth >= 13 Then
		CurrentYear = CurrentYear + 1
		CurrentMonth = 1
	End If
	If CurrentYear >= Year(Now) Then
		If CurrentMonth >= Month(Now) Then
			msgbox "Date Error"
		End If
	End If
	AIUtil.Context.Unfreeze
Next

'Close the application at the end of your script
'AppContext.Close											
