Public Sub Logout(params)
	'Declaration 
	Dim MainTimer, SubMainTimer, counterRepeat, limitRepeat, OuputVerif, GetLastError, GetFullFolderName
	
	'Input Parameter
	MainTimer = params("MainTimer")
	SubMainTimer = params("SubMainTimer")
	counterRepeat = params("counterRepeat")
	limitRepeat = params("limitRepeat")
	GetFullFolderName = params("GetFullFolderName")
	OuputVerif = params("OuputVerif")
	GetLastError = ""
	
	'Logout
	Browser("Browser").Page("Digital Credit Management").Sync
	Browser("Browser").Page("Digital Credit Management").Link("Link").Click
	'Capture Screen
	CaptureScreenBlank()
	Browser("Browser").HandleDialog micOK
	'Capture Screen
	CaptureScreenBlank()
	
	'Browser("Browser").Page("Retail Credit System").WebElement("Welcome").Click
	'Verifikasi Logout
	If Browser("Browser").Page("Retail Credit System").WebButton("Login").Exist(0) Then
		reporter.ReportEvent micPass, "Result Verifikasi Logout Succes", "Logout Succes"
	Else
		reporter.ReportEvent micFail, "Result Verifikasi Logout Failed", "Logout Failed"
	End If
	
	'Capture Screen
	CaptureScreenBlank()
	
	'Output
	params("OutputVerif") = OutputVerif
	params("GetLastError") = GetLastError
End Sub
