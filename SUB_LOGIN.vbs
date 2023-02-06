Public Sub Login(params)
	'Declaration 
	Dim MainTimer, SubMainTimer, counterRepeat, limitRepeat, OuputVerif, GetLastError, GetFullFolderName
	'Declaration Login
	Dim USERNAME, PASSWORD
	
	MainTimer = params("MainTimer")
	SubMainTimer = params("SubMainTimer")
	counterRepeat = params("counterRepeat")
	limitRepeat = params("limitRepeat")
	GetFullFolderName = params("GetFullFolderName")
	OuputVerif = params("OuputVerif")
	GetLastError = ""
	
	'Input Parameter Login
	USERNAME = params("USERNAME")
	PASSWORD = params("PASSWORD")
	
	'Login
	'Capture Screen
	CaptureScreenBlank()
	'Browser("Warning: Potential Security").Page("Retail Credit System").WebEdit("userId").Set "CCDE01"
	Browser("Warning: Potential Security").Page("Retail Credit System").WebEdit("userId").Set USERNAME
	'Capture Screen
	CaptureScreenBlank()
	'Browser("Warning: Potential Security").Page("Retail Credit System").WebEdit("password").SetSecure "63b50637d9bcf6118b204c01425e98580c6a594ff4f5"
	Browser("Warning: Potential Security").Page("Retail Credit System").WebEdit("password").Set PASSWORD
	'Capture Screen
	CaptureScreenBlank()
	Browser("Warning: Potential Security").Page("Retail Credit System").WebButton("Login").Click
	'Capture Screen
	CaptureScreenBlank()
	
	'Check PopUp
	If Window("Mozilla Firefox").InsightObject("InsightObject").Exist(0) Then
		Window("Mozilla Firefox").InsightObject("InsightObject").Click
		'Capture Screen
		CaptureScreenBlank()
		Window("Mozilla Firefox").InsightObject("InsightObject_2").Click
		'Capture Screen
		CaptureScreenBlank()
	End If
	
	Browser("Browser").Maximize
	'Capture Screen
	CaptureScreenBlank()
	
	'Verifikasi Login
	If Browser("Browser").Page("Digital Credit Management").Frame("main").WebElement("Welcome").Exist(0) Then
		reporter.ReportEvent micPass, "Result Verifikasi Login Succes", "Login Succes"
		OutputVerif = "Succes"
	Else
		reporter.ReportEvent micFail, "Result Verifikasi Login Failed", "Login Failed"
		OutputVerif = "Failed"
		GetLastError = "Login Failed"
	End If
	
	'Output
	params("OutputVerif") = OutputVerif
	params("GetLastError") = GetLastError
End Sub
