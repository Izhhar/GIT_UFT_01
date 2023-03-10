Public Sub OpenBrowser(params)
	'Declaration
	Dim MainTimer, SubMainTimer, counterRepeat, limitRepeat, OuputVerif, GetLastError, GetFullFolderName
	'Declaration Browser
	Dim BROWSER_TYPE, BROWSER_URL
	
	MainTimer = params("MainTimer")
	SubMainTimer = params("SubMainTimer")
	counterRepeat = params("counterRepeat")
	limitRepeat = params("limitRepeat")
	GetFullFolderName = params("GetFullFolderName")
	OuputVerif = params("OuputVerif")
	GetLastError = ""
	
	'Input Parameter	
	BROWSER_TYPE = params("BROWSER_TYPE")
	BROWSER_URL = params("BROWSER_URL")
	
	If BROWSER_TYPE = "IE" Then
		SystemUtil.Run "iexplorer.exe", BROWSER_URL
		OutputVerif = "Success"
	ElseIf BROWSER_TYPE = "Edge" Then
		SystemUtil.Run "msedge.exe", BROWSER_URL
		OutputVerif = "Success"
	ElseIf BROWSER_TYPE = "Firefox" Then
		SystemUtil.Run "firefox.exe", BROWSER_URL
		OutputVerif = "Success"
	ElseIf BROWSER_TYPE = "Chrome" Then
		SystemUtil.Run "chrome.exe", BROWSER_URL
		OutputVerif = "Success"
	Else  'if no browser provide, by default open Firefox
		'SystemUtil.Run "firefox.exe", BROWSER_URL
		OutputVerif = "Failed"
		GetLastError = "Open Browser Failed"
	End If
	
	'Capture Screen
	CaptureScreenBlank()
	
	'Check Default Browser
	If Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonNotNow").Exist(0) Then
		Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonNotNow").Click
	End If
	
	'Capture Screen
	CaptureScreenBlank()
	
	'Check Warning: Potential Security
	'If Browser("Warning: Potential Security").Page("Warning: Potential Security").Exist(0) Then
	If Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonAdvance").Exist(0) Then
		Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonAdvance").Click
		'Capture Screen
		CaptureScreenImage()
		Dim mySendKeys
		set mySendKeys = CreateObject("WScript.shell")
		mySendKeys.SendKeys("{PGDN}")
		'Capture Screen
		CaptureScreenImage()
		Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonAcceptAndContinue").Click
	End If
	
	'Capture Screen
	CaptureScreenBlank()
	
	Browser("CreationTime:=0").Sync
	Browser("CreationTime:=0").Maximize
	
	'Output
	params("OutputVerif") = OutputVerif
	params("GetLastError") = GetLastError
End Sub
