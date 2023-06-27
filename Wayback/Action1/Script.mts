'Declare variables that will be used in the script
Dim BrowserExecutable, oShell, CurrentYear, CurrentMonth, CurrentYearPlusMonth, CurrentURL, booForeignLanguage, Iterator

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
		CurrentYearPlusMonth = CurrentYear & "0" & CurrentMonth
		CurrentURL = "https://web.archive.org/web/" & CurrentYear & "0" & CurrentMonth & "/" & Parameter.Item("URL")
	Else
		CurrentYearPlusMonth = CurrentYear & CurrentMonth
		CurrentURL = "https://web.archive.org/web/" & CurrentYear & CurrentMonth & "/" & Parameter.Item("URL")
	End If
	If (((CurrentYearPlusMonth >= 201007) and (CurrentYearPlusMonth < 201102)) or ((CurrentYearPlusMonth >= 201109) and (CurrentYearPlusMonth < 201111)) _
	or (CurrentYearPlusMonth = 201201) or (CurrentYearPlusMonth = 201204) or (CurrentYearPlusMonth = 201210) or (CurrentYearPlusMonth = 201507) _
	or (CurrentYearPlusMonth = 201508) or (CurrentYearPlusMonth = 201510) or (CurrentYearPlusMonth = 201602) or (CurrentYearPlusMonth = 201607) _
	or (CurrentYearPlusMonth = 201608) or (CurrentYearPlusMonth >= 201611 and CurrentYearPlusMonth < 201702) or (CurrentYearPlusMonth = 201703) _
	or (CurrentYearPlusMonth = 201704) or (CurrentYearPlusMonth = 201706) or (CurrentYearPlusMonth = 201707) or (CurrentYearPlusMonth = 201801) _
	or (CurrentYearPlusMonth = 201806) or (CurrentYearPlusMonth = 201812) or (CurrentYearPlusMonth = 201902) or (CurrentYearPlusMonth = 201908) _
	or (CurrentYearPlusMonth = 201909) or (CurrentYearPlusMonth = 202001) or (CurrentYearPlusMonth = 202005) or (CurrentYearPlusMonth = 202006) _
	or (CurrentYearPlusMonth = 202009) or (CurrentYearPlusMonth = 202203) or (CurrentYearPlusMonth = 202303) ) Then
		Reporter.ReportEvent micWarning, "Navigating to " & CurrentURL, "The WayBack Machine is broken on this page, skipping."
	Else
		AppContext.Navigate CurrentURL
		AppContext.Sync	
		AIUtil.Context.SetBrowserScope(BrowserWindow)
		booForeignLanguage = AIUtil.FindText("Google Translate").Exist(0)
		AIUtil.Context.SetBrowserScope(WebPage)
		AIUtil.Context.Freeze
		print AppContext.GetROProperty("OpenURL")
		print instr(1, AppContext.GetROProperty("OpenURL"), "m.facebook.com")
		If (instr(1, AppContext.GetROProperty("OpenURL"), "m.facebook.com") >= 1) Then
			Reporter.ReportEvent micWarning, "Navigating to " & CurrentURL, "The WayBack Machine loaded the  mobile version of the page, aborting."
		ElseIf AIUtil.FindText("Mobile number").Exist(0) = True Then
			Reporter.ReportEvent micWarning, "Navigating to " & CurrentURL, "The WayBack Machine loaded the  mobile version of the page, aborting."
		ElseIf booForeignLanguage Then
			Reporter.ReportEvent micWarning, "Navigating to " & CurrentURL, "The WayBack Machine loaded a different language version of the page."
		ElseIf ((AIUtil.FindText("est un").Exist(0) = True) or (AIUtil.FindText("ouvert").Exist(0) = True) or (AIUtil.FindText("Het is gratis").Exist(0) = True) _ 
		or (AIUtil.FindText("Homme").Exist(0) = True) or (AIUtil.FindText("nyt ja").Exist(0) = True) or (AIUtil.FindText("e sempre").Exist(0) = True) _ 
		or (AIUtil.FindText("und blei").Exist(0) = True) or (AIUtil.FindText("y fácil").Exist(0) = True) or (AIUtil.FindText("sisällön").Exist(0)= True) _
		or (AIUtil.FindText("yritykselle").Exist(0)) or (AIUtil.FindText("vastaavia").Exist(0) = True) ) Then
			Reporter.ReportEvent micWarning, "Navigating to " & CurrentURL, "The WayBack Machine loaded a different language version of the page."
			print AIUtil.FindText("est un").Exist(0) & " " & AIUtil.FindText("ouvert").Exist(0) & " " & AIUtil.FindText("Het is gratis").Exist(0) & " " _
			& AIUtil.FindText("Homme").Exist(0) & " " & AIUtil.FindText("nyt ja").Exist(0) & " " & AIUtil.FindText("e sempre").Exist(0) & " " _
			& AIUtil.FindText("und blei").Exist(0) & " " & AIUtil.FindText("y fácil").Exist(0) & " " & AIUtil.FindText("sisällön").Exist(0) & " " _
			& AIUtil.FindText("yritykselle").Exist(0) & " " & AIUtil.FindText("vastaavia").Exist(0)
		ElseIf AIUtil.FindText("302 response").Exist(0) Then
			Reporter.ReportEvent micWarning, "Navigating to " & CurrentURL, "The WayBack Machine got an HTTP 302 response at crawl time."
		ElseIf AIUtil.FindText("something went wrong").Exist(0) Then
			Reporter.ReportEvent micWarning, "Navigating to " & CurrentURL, "The WayBack Machine got the something went wrong message at crawl time"
		Else
			'######################################################################################################
			'Change 1 https://web.archive.org/web/200810/http://facebook.com/ Labels for the e-mail and password fields gone, modification needed
			'######################################################################################################
			'Change 2 https://web.archive.org/web/200909/http://facebook.com/ Labels reappear, now as pre-populated values in the boxes, modification needed
			'######################################################################################################
			'Change 3 https://web.archive.org/web/201001/http://facebook.com/ Password label removed again
			'######################################################################################################
			'Change 4 https://web.archive.org/web/201006/http://facebook.com/ Password and E-mail labels above the fields again
			'######################################################################################################
			'Change 5 https://web.archive.org/web/201205/http://facebook.com/ login label changed to be "Email or Phone"
			'######################################################################################################
			'Change 6 https://web.archive.org/web/202011/http://facebook.com/ login label changed to be "Email or Phone Number", page redesigned, but if label hadn't changed it 
			'	wouldn't have mattered
			'######################################################################################################
			'Change 7 https://web.archive.org/web/202110/http://facebook.com/ Cookies popup screen added, have to handle
			'######################################################################################################
			'Change 8 https://web.archive.org/web/202204/http://facebook.com/ Login field label changed
			'######################################################################################################
			'Change 9 https://web.archive.org/web/202205/http://facebook.com/ Cookies popup label changed, and the Login field label changed back to what it was before 202204
			'######################################################################################################
			'Starting https://web.archive.org/web/201007/http://facebook.com/ errors, gets into a loop of partially loading the page, then starts a reload
			'Resolved after https://web.archive.org/web/201102/http://facebook.com/ 
			'Starting https://web.archive.org/web/201109/http://facebook.com/, same errors, not resolved until https://web.archive.org/web/201111/http://facebook.com/
			'https://web.archive.org/web/201201/http://facebook.com/ is broken too
			'https://web.archive.org/web/201204/http://facebook.com/ is broken too
			'https://web.archive.org/web/201210/http://facebook.com/ is broken too
			'https://web.archive.org/web/201507/http://facebook.com/ is broken too
			'https://web.archive.org/web/201510/http://facebook.com/ is broken too
			'https://web.archive.org/web/201602/http://facebook.com/ is broken too
			'https://web.archive.org/web/201607/http://facebook.com/ is broken too
			'https://web.archive.org/web/201608/http://facebook.com/ is broken too
			'https://web.archive.org/web/201611/http://facebook.com/ is broken too
			'https://web.archive.org/web/201703/http://facebook.com/ and 04 is broken too
			'https://web.archive.org/web/201706/http://facebook.com/ and 07 is broken too (crawl had an already logged in user)
			'https://web.archive.org/web/201801/http://facebook.com/ is broken too (crawl had an already logged in user)
			'https://web.archive.org/web/201806/http://facebook.com/ is broken too (looks like it loaded the mobile version)
			'https://web.archive.org/web/201812/http://facebook.com/ is broken too
			'https://web.archive.org/web/201902/http://facebook.com/ is broken too
			'https://web.archive.org/web/201908/http://facebook.com/ and 09 is broken too
			'https://web.archive.org/web/202001/http://facebook.com/ is broken too
			'https://web.archive.org/web/202005/http://facebook.com/ is broken too
			'https://web.archive.org/web/202006/http://facebook.com/ is broken too
			'https://web.archive.org/web/202009/http://facebook.com/ is broken too
			'https://web.archive.org/web/202203/http://facebook.com/ is redirecting
			'https://web.archive.org/web/202303/http://facebook.com/ is broken too
			'######################################################################################################
			Select Case True
				Case (CurrentYearPlusMonth >= 202204)
					If AIUtil("button", "Allow All Cookies").Exist(0) = True Then
						AIUtil("button", "Allow All Cookies").Click
					ElseIf AIUtil("button", "Allow essential and optional cookies").Exist(0) = True Then
						AIUtil("button", "Allow essential and optional cookies").Click
					End If
					AIUtil.Context.UnFreeze
					If AIUtil("text_box", "Email or Phone Number").Exist(0) = True Then
						AIUtil.Context.Freeze
						AIUtil("text_box", "Email or Phone Number").Highlight
						AIUtil("text_box", "Email or Phone Number").SetText "user@domain.com"
					Else
						AIUtil.Context.Freeze
						AIUtil("text_box", "Email address or Phone Number").Highlight
						AIUtil("text_box", "Email address or Phone Number").SetText "user@domain.com"
					End If
					AIUtil("text_box", "Password").Highlight
					AIUtil("text_box", "Password").SetText "Password"
					AIUtil("button", "Log In").Highlight
					AIUtil("button", "Log In").CheckExists True
				Case ((CurrentYearPlusMonth >= 202011) and (CurrentYearPlusMonth < 202204))
					If AIUtil("button", "Allow All Cookies").Exist(0) = True Then
						AIUtil("button", "Allow All Cookies").Click
					End If
					AIUtil.Context.UnFreeze
'					AIUtil("text_box", "Email or Phone Number").Highlight
					AIUtil.Context.Freeze
					AIUtil("text_box", "Email or Phone Number").SetText "user@domain.com"
'					AIUtil("text_box", "Password").Highlight
					AIUtil("text_box", "Password").SetText "Password"
					AIUtil("button", "Log In").Highlight
					AIUtil("button", "Log In").CheckExists True
				Case ((CurrentYearPlusMonth >= 201205) and (CurrentYearPlusMonth < 202011))
					AIUtil.FindTextBlock("facebook").Hover
'					AIUtil("text_box", "Email or Phone").Highlight
					AIUtil("text_box", "Email or Phone").SetText "user@domain.com"
'					AIUtil("text_box", "Password").Highlight
					AIUtil("text_box", "Password").SetText "Password"
					AIUtil("button", micAnyText, micWithAnchorOnLeft, AIUtil("text_box", "Password")).Highlight
					AIUtil("button", micAnyText, micWithAnchorOnLeft, AIUtil("text_box", "Password")).CheckExists True
				Case ((CurrentYearPlusMonth >= 201006) and (CurrentYearPlusMonth < 201205))
					AIUtil.FindTextBlock("facebook").Hover
					AIUtil.Context.Unfreeze
					AIUtil.FindText("It's free").Click
					AIUtil.Context.Freeze
'					AIUtil("text_box", "Email").Highlight
					AIUtil("text_box", "Email").SetText "user@domain.com"
'					AIUtil("text_box", "Password").Highlight
					AIUtil("text_box", "Password").SetText "Password"
					AIUtil("button", micAnyText, micWithAnchorOnLeft, AIUtil("text_box", "Password")).Highlight
					AIUtil("button", micAnyText, micWithAnchorOnLeft, AIUtil("text_box", "Password")).CheckExists True
				Case ((CurrentYearPlusMonth >= 201001) and (CurrentYearPlusMonth < 201006))
					AIUtil.Context.Unfreeze
					AIUtil.FindText("It's free").Click
					AIUtil.Context.Freeze
	'				AIUtil("text_box", "Email").Highlight
					AIUtil("text_box", "Email").SetText "user@domain.com"
	'				AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?")).Highlight
					AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?")).SetText "Password"
					AIUtil("button", micAnyText, micWithAnchorOnLeft, AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?"))).Highlight
					AIUtil("button", micAnyText, micWithAnchorOnLeft, AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?"))).CheckExists True
				Case ((CurrentYearPlusMonth >= 200909) and (CurrentYearPlusMonth < 201001))
					AIUtil.Context.Unfreeze
					AIUtil.FindText("It's free").Click
					AIUtil.Context.Freeze
	'				AIUtil("text_box", "Email").Highlight
					AIUtil("text_box", "Email").SetText "user@domain.com"
	'				AIUtil("text_box", "Password").Highlight
					AIUtil("text_box", "Password").SetText "Password"
					AIUtil("button", "Login").Highlight
					AIUtil("button", "Login").CheckExists True
				Case ((CurrentYearPlusMonth >= 200810) and (CurrentYearPlusMonth < 200909	))
	'				AIUtil("text_box", micAnyText, micWithAnchorOnRight, AIUtil.FindTextBlock("Remember Me")).Highlight
					AIUtil("text_box", micAnyText, micWithAnchorOnRight, AIUtil.FindTextBlock("Remember Me")).SetText "user@domain.com"
	'				AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?")).Highlight
					AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?")).SetText "Password"
					AIUtil("button", "Login").Highlight
					AIUtil("button", "Login").CheckExists True
				Case Else
	'				AIUtil("text_box", "E-mail:").Highlight	
					AIUtil("text_box", "E-mail:").SetText "user@domain.com"
					If AIUtil("text_box", "Password:").Exist(0) = True Then
	'					AIUtil("text_box", "Password:").Highlight
						AIUtil("text_box", "Password:").SetText "Password"
						Set PasswordField = AIUtil("text_box", "Password:")
					ElseIf AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")) = True Then	
	'					AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")).Highlight
						AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")).SetText "Password"
						Set PasswordField = AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password"))
					Else
						msgbox "Can't find Password field"
					End If
					AIUtil("button", micAnyText, micWithAnchorAbove, PasswordField).Highlight
					AIUtil("button", micAnyText, micWithAnchorAbove, PasswordField).CheckExists True
			End Select
		End If
	End If
	CurrentMonth = CurrentMonth + 1
	If CurrentMonth >= 13 Then
		CurrentYear = CurrentYear + 1
		CurrentMonth = 1
	End If
	If CurrentYear >= Year(Now) Then
		If CurrentMonth >= Month(Now) Then
			ExitActionIteration
		End If
	End If
	AIUtil.Context.Unfreeze
Next

'Close the application at the end of your script
'AppContext.Close											








'########################################################################################
'Old code after end select
'		If CurrentYearPlusMonth >= 200810 Then
'			If CurrentYearPlusMonth >= 200909 Then
'				If CurrentYearPlusMonth >= 201001 Then '>= 201001
'					AIUtil.FindText("It's free").Click
'					AIUtil("text_box", "Email").Highlight
'					AIUtil("text_box", "Email").SetText "user@domain.com"
'					AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?")).Highlight
'					AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?")).SetText "Password"
'					AIUtil("button", "Login").Highlight
'					AIUtil("button", "Login").CheckExists True
'				Else 
'					AIUtil.FindText("It's free").Click
'					AIUtil("text_box", "Email").Highlight
'					AIUtil("text_box", "Email").SetText "user@domain.com"
'					AIUtil("text_box", "Password").Highlight
'					AIUtil("text_box", "Password").SetText "Password"
'					AIUtil("button", "Login").Highlight
'					AIUtil("button", "Login").CheckExists True
'				End If
'			Else		
'				AIUtil("text_box", micAnyText, micWithAnchorOnRight, AIUtil.FindTextBlock("Remember Me")).Highlight
'				AIUtil("text_box", micAnyText, micWithAnchorOnRight, AIUtil.FindTextBlock("Remember Me")).SetText "user@domain.com"
'				AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?")).Highlight
'				AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindTextBlock("Forgot your password?")).SetText "Password"
'				AIUtil("button", "Login").Highlight
'				AIUtil("button", "Login").CheckExists True
'			End If
'		Else
'			AIUtil("text_box", "E-mail:").Highlight	
'			AIUtil("text_box", "E-mail:").SetText "user@domain.com"
'			If AIUtil("text_box", "Password:").Exist(0) = True Then
'				AIUtil("text_box", "Password:").Highlight
'				AIUtil("text_box", "Password:").SetText "Password"
'				Set PasswordField = AIUtil("text_box", "Password:")
'			ElseIf AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")) = True Then	
'				AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")).Highlight
'				AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password")).SetText "Password"
'				Set PasswordField = AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password"))
'			Else
'				msgbox "Can't find Password field"
'			End If
'			'######################################################################################################
'			'https://web.archive.org/web/200810/http://facebook.com/ Login button no longer below the Password field, it's now to the right, modification needed
'			'######################################################################################################
'			AIUtil("button", micAnyText, micWithAnchorAbove, PasswordField).Highlight
'			AIUtil("button", micAnyText, micWithAnchorAbove, PasswordField).CheckExists True
'			'https://web.archive.org/web/200703/http://facebook.com/ has a problem, looks like it's identifying the Register box as the login?
'		'	If AIUtil("button", "Login").Exist(0) = True Then
'		'		Set LoginProperties = AIUtil("button", "Login").GetAllProperties
'		'		For i = 0 To LoginProperties.count - 1
'		'			print LoginProperties.keys()(i) & ":" & LoginProperties(LoginProperties.keys()(i))
'		'		Next
'		'		AIUtil("button", "Login").Highlight
'		'		AIUtil("button", "Login").CheckExists True
'		'	ElseIf AIUtil("button", "Login", micFromTop, 1).Exist(0) = True Then
'		'		AIUtil("button", "Login", micFromTop, 1).Highlight
'		'		AIUtil("button", "Login", micFromTop, 1).CheckExists True
'		'	Else
'		'		AIUtil("button", "Login", micWithAnchorAbove, AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password"))).Highlight
'		'		AIUtil("button", "Login", micWithAnchorAbove, AIUtil("text_box", micAnyText, micWithAnchorAbove, AIUtil.FindText("Password"))).CheckExists True
'		'	End If
'		
'		End If	
'########################################################################################

