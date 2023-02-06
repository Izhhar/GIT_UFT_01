'Minimize UFT
'MinimizeUFTWindow()

'Get Project Name Dan Asset
Dim TestName, ActionName
TestName = Environment.Value("TestName")
ActionName = Environment.Value("ActionName")

'Define Active Data From Parameter
Dim StartRow, EndRow, ExcelName, SheetName

'Define variable from paramter
Dim MainTime, SubMainTimer, counterRepeat, limitRepeat
MainTimer = 80000
SubMainTimer = 1000
counterRepeat = 0
limitRepeat = 5

'Define Parameter With Type Dictionary
Dim params   ' Create a variable.
Set params = CreateObject("Scripting.Dictionary")

'Sub to playback
'Input Parameter
Dim BROWSER_TYPE, BROWSER_URL

'Local Parameter
Dim OutputVerif, GetLastError
OutputVerif = "Failed"
GetLastError = ""

'Active Data Parameter

'Create Result Folder
'params("GetTestName") = TestName
'params("GetActionName") = ActionName
'CreateResultFolder(params)
'PutUploadFolderName = params("PutUploadFolderName")
'PutRunFolderName = params("PutRunFolderName")
CreateFolderResult()

'Define Capture Screenshots Report
'RegisterUserFunc "Page", "CaptureScreenshot", "CaptureScreenshot"
RegisterUserFunc "Page", "CaptureReport", "CaptureReport"
'RegisterUserFunc "Browser", "CaptureScreenshot", "CaptureScreenshot"
RegisterUserFunc "Browser", "CaptureReport", "CaptureReport"
'RegisterUserFunc "Frame", "CaptureScreenshot", "CaptureScreenshot"
RegisterUserFunc "Frame", "CaptureReport", "CaptureReport"
'RegisterUserFunc "Dialog", "CaptureScreenshot", "CaptureScreenshot"
RegisterUserFunc "Dialog", "CaptureReport", "CaptureReport"
'RegisterUserFunc "swfWindow", "CaptureScreenshot", "CaptureScreenshot"
RegisterUserFunc "swfWindow", "CaptureReport", "CaptureReport"
'RegisterUserFunc "Window", "CaptureScreenshot", "CaptureScreenshot"
RegisterUserFunc "Window", "CaptureReport", "CaptureReport"

'Read Data Form Excel
'Source Excel Declaration
StartRow = 1
EndRow = 0
ExcelName = "C:\ActiveData\TestAD.xlsx"
SheetName = "BROWSER"

DataTable.AddSheet SheetName
DataTable.ImportSheet ExcelName, SheetName, SheetName

If DataTable.GetSheet(SheetName).GetRowCount > 0 Then
	For i  = StartRow To DataTable.GetSheet(SheetName).GetRowCount - EndRow
		DataTable.SetCurrentRow(i)
		
		'Declaration Variable Active Data
		BROWSER_TYPE = DataTable.Value("BROWSER_TYPE",SheetName)
		BROWSER_URL = DataTable.Value("BROWSER_URL",SheetName)
		ACTIVE = DataTable.Value("ACTIVE",SheetName)
		
		If ACTIVE = "Y" Then
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
			CaptureScreen()			
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			'Browser("micclass:=Browser").Page("micclass:=Page").CaptureScreenshot micPass, "Page Screenshot."
			
			'Check Default Browser
			If Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonNotNow").Exist(0) Then
				Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonNotNow").Click
			End If
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Check Warning: Potential Security
			'If Browser("Warning: Potential Security").Page("Warning: Potential Security").Exist(0) Then
			If Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonAdvance").Exist(0) Then
				Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonAdvance").Click
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
				Dim mySendKeys
				set mySendKeys = CreateObject("WScript.shell")
				mySendKeys.SendKeys("{PGDN}")
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
				Browser("Warning: Potential Security").Page("Warning: Potential Security").InsightObject("InsightObject_ButtonAcceptAndContinue").Click
			End  If
			
			If Browser("Warning: Potential Security").Page("Warning: Potential Security_2").InsightObject("InsightObject").Exist(0) Then	
				Browser("Warning: Potential Security").Page("Warning: Potential Security_2").InsightObject("InsightObject").Click @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf131.xml_;_
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
				set mySendKeys = CreateObject("WScript.shell")
				mySendKeys.SendKeys("{PGDN}")
				Browser("Warning: Potential Security").Page("Warning: Potential Security_2").InsightObject("InsightObject_2").Click @@ hightlight id_;_8_;_script infofile_;_ZIP::ssf132.xml_;_
			End If
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			Browser("CreationTime:=0").Sync
			Browser("CreationTime:=0").Maximize
			
			If OutputVerif = "Failed" Then
				'GoTo NextData
			End If
		End If
	Next
End If
 @@ hightlight id_;_394714_;_script infofile_;_ZIP::ssf1.xml_;_
'Source Excel Declaration
StartRow = 1
EndRow = 0
ExcelName = "C:\ActiveData\TestAD.xlsx"
SheetName = "ACTIVEDATA"

DataTable.AddSheet SheetName
DataTable.ImportSheet ExcelName, SheetName, SheetName

If DataTable.GetSheet(SheetName).GetRowCount > 0 Then
	For i  = StartRow To DataTable.GetSheet(SheetName).GetRowCount - EndRow
		DataTable.SetCurrentRow(i)
		
		'Declaration Variable Active Data
		TESTSCRIPT = DataTable.Value("TESTSCRIPT",SheetName)
		ACTIVE = DataTable.Value("ACTIVE",SheetName)
		'Login
		USERNAME = DataTable.Value("USERNAME",SheetName)
		PASSWORD = DataTable.Value("PASSWORD",SheetName)
		'Data
		SALES_NAME = DataTable.Value("SALES_NAME",SheetName)
		BRANCH = DataTable.Value("BRANCH",SheetName)
		CHANNEL = DataTable.Value("CHANNEL",SheetName)
		MARKETING_CODE = DataTable.Value("MARKETING_CODE",SheetName)
		CAMPAIGN_CODE_1 = DataTable.Value("CAMPAIGN_CODE_1",SheetName)
		CAMPAIGN_CODE_2 = DataTable.Value("CAMPAIGN_CODE_2",SheetName)
		CAMPAIGN_CODE_3 = DataTable.Value("CAMPAIGN_CODE_3",SheetName)
		FIRST_NAME = DataTable.Value("FIRST_NAME",SheetName)
		MIDDLE_NAME = DataTable.Value("MIDDLE_NAME",SheetName)
		LAST_NAME = DataTable.Value("LAST_NAME",SheetName)
		FULL_NAME = DataTable.Value("FULL_NAME",SheetName)
		GENDER = DataTable.Value("GENDER",SheetName)
		PLACE_OF_BIRTH = DataTable.Value("PLACE_OF_BIRTH",SheetName)
		ID_TYPE = DataTable.Value("ID_TYPE",SheetName)
		CITY = DataTable.Value("CITY",SheetName)
		KODE_TUJUAN_PERMINTAAN = DataTable.Value("KODE_TUJUAN_PERMINTAAN",SheetName)
		DATE_OF_BIRTH = DataTable.Value("DATE_OF_BIRTH",SheetName)
		ID_NOMOR = DataTable.Value("ID_NO",SheetName)
		ADDRESS = DataTable.Value("ADDRESS",SheetName)
		EMPLOYMENT_CATEGORY = DataTable.Value("EMPLOYMENT_CATEGORY",SheetName)
		IS_BUNDLING = DataTable.Value("IS_BUNDLING",SheetName)
		BUNDLING_CODE = DataTable.Value("BUNDLING_CODE",SheetName)
		CARD_LIMIT = DataTable.Value("CARD_LIMIT",SheetName)
		
		If ACTIVE = "Y" Then
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Check Update Browser
			If Window("Mozilla Firefox").InsightObject("InsightObject_5").Exist(0) Then
				Window("Mozilla Firefox").InsightObject("InsightObject_3").Click
			End If
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Browser("Warning: Potential Security").Page("Retail Credit System").WebEdit("userId").Set "CCDE01"
			'Browser("Warning: Potential Security").Page("Retail Credit System").WebEdit("userId").Set USERNAME
			Browser("Browser").Page("Retail Credit System").WebEdit("userId").Set USERNAME
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Browser("Warning: Potential Security").Page("Retail Credit System").WebEdit("password").SetSecure "63b50637d9bcf6118b204c01425e98580c6a594ff4f5"
			'Browser("Warning: Potential Security").Page("Retail Credit System").WebEdit("password").Set PASSWORD
			Browser("Browser").Page("Retail Credit System").WebEdit("password").Set PASSWORD
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Browser("Warning: Potential Security").Page("Retail Credit System").WebButton("Login").Click
			Browser("Browser").Page("Retail Credit System").WebButton("Login").Click
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Browser("Browser").Page("Retail Credit System").WebEdit("userId").Set "user" @@ hightlight id_;_329346_;_script infofile_;_ZIP::ssf116.xml_;_
			'Browser("Browser").Page("Retail Credit System").WebEdit("password").SetSecure "63c676e879feaf463a5e" @@ script infofile_;_ZIP::ssf117.xml_;_
			'Browser("Browser").Page("Retail Credit System").WebButton("Login").Click @@ script infofile_;_ZIP::ssf118.xml_;_
			
			'Check PopUp
			If Window("Mozilla Firefox").InsightObject("InsightObject").Exist(0) Then
				Window("Mozilla Firefox").InsightObject("InsightObject").Click
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
				Window("Mozilla Firefox").InsightObject("InsightObject_2").Click
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			End  If
						
			If Window("Mozilla Firefox").InsightObject("InsightObject_6").Exist(0) Then
				Window("Mozilla Firefox").InsightObject("InsightObject_7").Click @@ hightlight id_;_15_;_script infofile_;_ZIP::ssf140.xml_;_
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
				Window("Mozilla Firefox").InsightObject("InsightObject_8").Click @@ hightlight id_;_20_;_script infofile_;_ZIP::ssf142.xml_;_
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			End If
			
			'Browser("CreationTime:=0").Sync
			Browser("Browser").Maximize			
			'Browser("Browser").Sync
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			Browser("Browser").RefreshObject
			Browser("Browser").Page("Digital Credit Management").Frame("main").WebElement("Welcome").WaitProperty "innertext", "Welcome", 10000
			
			'Verifikasi Login
			If Browser("Browser").Page("Digital Credit Management").Frame("main").WebElement("Welcome").Exist(0) Then
				reporter.ReportEvent micPass, "Result Verifikasi Login Succes", "Login Succes"
				OutputVerif = "Succes"
			Else
				reporter.ReportEvent micFail, "Result Verifikasi Login Failed", "Login Failed"
				OutputVerif = "Failed"
				GetLastError = "Login Failed"
				'GoTo
			End If
			
			'wait(10)
			Browser("Browser").RefreshObject
			'wait(300)
			'Click Menu
			Browser("Browser").Page("Digital Credit Management").Image("menu2").Click
			'Browser("Browser").Page("Digital Credit Management").Frame("main").Image("menu2").Click
			'Browser("Browser").Page("Digital Credit Management").Image("menu2").Click
			'Browser("Browser").Page("Digital Credit Management").Image("xpath:=//DIV[@id='menuNav']/INPUT[1]").Click
			'If Window("Digital Credit Management").InsightObject("InsightObject_3").Exist(0) Then @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf120.xml_;_
			'	Window("Digital Credit Management").InsightObject("InsightObject_3").Click
			'End  If
			'If Window("Digital Credit Management").InsightObject("InsightObject_5").Exist(0) Then @@ hightlight id_;_3_;_script infofile_;_ZIP::ssf145.xml_;_
			'	Window("Digital Credit Management").InsightObject("InsightObject_5").Click
			'End  If
			
			'Browser("Browser").Page("Digital Credit Management").Image("menu2_2").Click 11,53 @@ script infofile_;_ZIP::ssf143.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Image("menu2_2").Click
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			Browser("Browser").Page("Digital Credit Management").Link("New Application - New").Click
			'Browser("Browser").Page("Digital Credit Management").Link("xpath:=//DIV[@id='domRoot']/DIV[22]/TABLE[1]/TBODY[1]/TR[1]/TD[3]/A[1]")
			'Window("Digital Credit Management").InsightObject("InsightObject_4").Click @@ hightlight id_;_36_;_script infofile_;_ZIP::ssf128.xml_;_
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Wait for
			'Browser("Browser").Page("Digital Credit Management").Frame("main").WebElement("New Application - New").WaitProperty "innertext", "New Application - New", 10000
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("New Application - New").WaitProperty "innertext", "New Application - New", 10000 @@ script infofile_;_ZIP::ssf112.xml_;_
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			'wait(300)
			'Sales Name
			'Browser("Browser").Page("Digital Credit Management").Frame("main").WebEdit("salesCode1Desc").Set "00019" @@ script infofile_;_ZIP::ssf37.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("salesCode1Desc").Set SALES_NAME
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link").Click @@ script infofile_;_ZIP::ssf38.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672816455260").Link("00019").Click
			'wait(1)
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(SALES_NAME).Click
			Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link("text:="&SALES_NAME).Click
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link("title:=SelectedLink").Click
			'AIUtil.FindTextBlock("00019").Click
			'AIUtil.FindTextBlock(SALES_NAME).Click
			'AIUtil.SetContext Window("regexpwndtitle:=Digital Credit Management System — Mozilla Firefox", "regexpwndclass:=MozillaWindowClass", "is owned window:=False", "is child window:=False")
			'AIUtil.FindTextBlock("00019", micFromBottom, 1).Click
			'AIUtil.FindTextBlock(SALES_NAME, micFromBottom, 1).Click
			
			'Branch
			'Browser("Browser").Page("Digital Credit Management").Frame("main_2").WebEdit("originatorBranchDescription").Set "AABTOJ" @@ script infofile_;_ZIP::ssf41.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("originatorBranchDescription").Set BRANCH
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_2").Click @@ script infofile_;_ZIP::ssf44.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672817098182").Link("AABTOJ").Click @@ script infofile_;_ZIP::ssf43.xml_;_
			'wait(1)
			Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(BRANCH).Click
			
			'Channel
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("channelDescription").Set "AH" @@ script infofile_;_ZIP::ssf45.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("channelDescription").Set CHANNEL
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_3").Click @@ script infofile_;_ZIP::ssf46.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820281009").Link("AH").Click @@ script infofile_;_ZIP::ssf47.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(CHANNEL).Click
			
			'Marketing Code
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("marketingCodeDescription").Set "AARG" @@ script infofile_;_ZIP::ssf48.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("marketingCodeDescription").Set MARKETING_CODE
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_4").Click @@ script infofile_;_ZIP::ssf49.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820296604").Link("AARG").Click @@ script infofile_;_ZIP::ssf50.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(MARKETING_CODE).Click
			
			'Campaign Code 1
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode1").Select "BCC - Bundling Payroll & CC" @@ script infofile_;_ZIP::ssf51.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode1").Select CAMPAIGN_CODE_1
			
			'Campaign Code 2
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode2").Select "KR3 - Bundling Mortgage dan CC (Non CP)" @@ script infofile_;_ZIP::ssf52.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode2").Select CAMPAIGN_CODE_2
			
			'Campaign Code 3
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode3").Select "000000 - Non Program" @@ script infofile_;_ZIP::ssf53.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode3").Select CAMPAIGN_CODE_3
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'First Name
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstName").Set "asd" @@ script infofile_;_ZIP::ssf54.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstName").Set FIRST_NAME
			
			'MIddle Name
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("midName").Set "asd" @@ script infofile_;_ZIP::ssf55.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("midName").Set MIDDLE_NAME
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("id1box").Click @@ script infofile_;_ZIP::ssf56.xml_;_
			
			'Last Name
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("lastName").Set "asd" @@ script infofile_;_ZIP::ssf57.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("lastName").Set LAST_NAME
			
			'Full Name
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("genFullNameWithoutAbbrLine1").Set "asd" @@ script infofile_;_ZIP::ssf58.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("genFullNameWithoutAbbrLine1").Set FULL_NAME
			
			'Gender
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("genderCode").Select "F - Female" @@ script infofile_;_ZIP::ssf59.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("genderCode").Select GENDER
			
			'Place of Birth
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("placeOfBirth").Set "jkt" @@ script infofile_;_ZIP::ssf60.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("placeOfBirth").Set PLACE_OF_BIRTH
			
			'ID Type
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("firstIdentificationTypeCode").Select "KT2 - KTP Seumur Hidup" @@ script infofile_;_ZIP::ssf61.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("3").Click @@ script infofile_;_ZIP::ssf62.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("firstIdentificationTypeCode").Select ID_TYPE
			
			'Set Number Area
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("mobileNumberAreaCode").Set "08" @@ script infofile_;_ZIP::ssf64.xml_;_
			'Browser("Browser").HandleDialog micOK @@ hightlight id_;_526168_;_script infofile_;_ZIP::ssf65.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("mobileNumberAreaCode").Set "62" @@ script infofile_;_ZIP::ssf67.xml_;_
			'Browser("Browser").HandleDialog micOK @@ hightlight id_;_526168_;_script infofile_;_ZIP::ssf68.xml_;_
			
			'City
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_5").Click @@ script infofile_;_ZIP::ssf69.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820520763").WebButton("Close").Click @@ script infofile_;_ZIP::ssf70.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("townDesc").Set "0111" @@ script infofile_;_ZIP::ssf71.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("townDesc").Set CITY
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("id1box_2").Click @@ script infofile_;_ZIP::ssf72.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_5").Click @@ script infofile_;_ZIP::ssf73.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820535716").WebElement("0111").Click @@ script infofile_;_ZIP::ssf74.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820535716").Link("0111").Click @@ script infofile_;_ZIP::ssf75.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(CITY).Click
			
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("townDesc").Set "0001" @@ script infofile_;_ZIP::ssf89.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("id1box_3").Click @@ script infofile_;_ZIP::ssf90.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_5").Click @@ script infofile_;_ZIP::ssf91.xml_;_
			
			'Kode Tujuan Permintaan			
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("requestPurpose").Select "01 - Penilaian calon debitur" @@ script infofile_;_ZIP::ssf100.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("requestPurpose").Select KODE_TUJUAN_PERMINTAAN
						
			'Date of Birth
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("4").Click @@ script infofile_;_ZIP::ssf96.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("dateOfBirth").Set "04/01/2023" @@ script infofile_;_ZIP::ssf97.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("dateOfBirth").Set DATE_OF_BIRTH
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("dateOfBirth").Submit @@ script infofile_;_ZIP::ssf98.xml_;_
			
			'ID No
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstIdentificationNumber").Set "12345" @@ script infofile_;_ZIP::ssf76.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstIdentificationNumber").Click
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstIdentificationNumber").Set ID_NOMOR
						
			'Address
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("addressLine1").Set "alamat" @@ script infofile_;_ZIP::ssf77.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("addressLine1").Set ADDRESS
			
			'Employment Category
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("employmentCategoryCode").Select "F - PROFESIONAL" @@ script infofile_;_ZIP::ssf78.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("employmentCategoryCode").Select EMPLOYMENT_CATEGORY
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Is Bundling
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebRadioGroup("bundlingFlag").Select "Y" @@ script infofile_;_ZIP::ssf79.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebRadioGroup("bundlingFlag").Select IS_BUNDLING
			
			wait(5)
			'Dim mySendKeys
			set mySendKeys = CreateObject("WScript.shell")
			mySendKeys.SendKeys("{PGDN}")
			
			'Bundling Code
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("bundlingTypeCode").Set "BPP1"
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("bundlingTypeCode").Set BUNDLING_CODE
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_6").Click @@ script infofile_;_ZIP::ssf81.xml_;_
			wait(5)
			Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820629562").WebButton("Go").Click @@ script infofile_;_ZIP::ssf82.xml_;_
			wait(5) @@ script infofile_;_ZIP::ssf83.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").WebElement("Code").Link("BPP1").Click
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").WebElement("Code").DoubleClick
'			Dim counter, limit
'			counter = 0
'			limit = 10
'			Do While counter < limit
'				'myNum = myNum - 1
'				If Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link("BPP1").Exist(0) Then
'					set mySendKeys = CreateObject("WScript.shell")
'					mySendKeys.SendKeys("{TAB}")
'					Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link("BPP1").Click
'				Else
'					Exit Do
'				End If				
'				counter = counter + 1
'			Loop
'			counter = 0
'			wait(2)
			
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").WebElement("Code").Link("BPP1").Click
			'If BROWSER_TYPE = "Firefox" Then
				'AIUtil.SetContext Window("regexpwndtitle:=Digital Credit Management System — Mozilla Firefox", "regexpwndclass:=MozillaWindowClass", "is owned window:=False", "is child window:=False")
				'AIUtil.FindTextBlock(BUNDLING_CODE, micFromBottom, 1).Click
			'ElseIf BROWSER_TYPE = "Chrome" Then
				
			'End If
			'AIUtil.SetContext Window("regexpwndtitle:=Digital Credit Management System — Mozilla Firefox", "regexpwndclass:=MozillaWindowClass", "is owned window:=False", "is child window:=False")
			'AIUtil.FindTextBlock("BPP1", micFromBottom, 1).Click @@ script infofile_;_ZIP::ssf92.xml_;_
			'AIUtil.FindTextBlock(BUNDLING_CODE, micFromBottom, 1).Click
			'AIUtil.FindTextBlock("BPP1").Click
			'AIUtil.FindTextBlock(BUNDLING_CODE).Click
			
			AIUtil.SetContext Browser("creationtime:=1")
			AIUtil.FindTextBlock(BUNDLING_CODE, micFromBottom, 1).Click
			'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link("BPP1").Click
			
			wait(5)
			mySendKeys.SendKeys("{PGDN}")
			wait(5)
			
			'Card Limit
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("cardLimit").Set "800000"
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("cardLimit").Set CARD_LIMIT
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Is Bundling * Yes No Group").Click @@ script infofile_;_ZIP::ssf86.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Is Bundling * Yes No Group").Click @@ script infofile_;_ZIP::ssf87.xml_;_
			
			'Click Proceed
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Proceed").Click @@ script infofile_;_ZIP::ssf88.xml_;_
			'wait(5)
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("New Application - New").WaitProperty "innertext", "New Application - New", 10000
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Verifikasi
			'If Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("We are currently unable").Exist(0) Then
			If Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("New Application - New").Exist(0) Then
				reporter.ReportEvent micPass, "Result Verifikasi Succes", "Succes"
			Else
				reporter.ReportEvent micFail, "Result Verifikasi Failed", "Failed"
			End If
			
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("We are currently unable").Click @@ script infofile_;_ZIP::ssf101.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("If the problem persists,").Click @@ script infofile_;_ZIP::ssf102.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Reference No : ED7F5DED2D49").Click @@ script infofile_;_ZIP::ssf103.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("We apologise for the inconveni").Click @@ script infofile_;_ZIP::ssf104.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Thank you.").Click @@ script infofile_;_ZIP::ssf108.xml_;_
			
			'get
			'Dim GetNoApp
			'GetNoApp1 = Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("GetNoApp").RefreshObject
			'GetNoApp1 = Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("GetNoApp").GetROProperty("innertext")
			'AIUtil.FindTextBlock("Application No", micWithAnchorBelow, AIUtil.FindTextBlock("0423020000014")).Click
			
			'AIUtil.SetContext Browser("creationtime:=1")
			'AIUtil.FindTextBlock("0423020000014").Click
			'AIUtil.FindTextBlock("Application No").Click
			'GetNoApp1 = AIUtil.FindTextBlock("Application No", micWithAnchorBelow, AIUtil.FindTextBlock("0423020000014")).GetText
			'GetNoApp1 = AIUtil.FindTextBlock("*", micWithAnchorBelow, AIUtil.FindTextBlock("Application No")).GetText
			'MsgBox GetNoApp1
			'GetNoApp2 = Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Name:=GetNoApp","xpath:=//TD/H3[normalize-space()='*']").GetTOProperties("innertext")
			'GetNoApp = Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebTable("html tag:=TABLE").ColumnCount(2)
			'IntGetNoApp = GetNoApp.GetCellData("H3",1)
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("0423020000009").Click
			'MsgBox GetNoApp2
			
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("0423010000181").Click @@ script infofile_;_ZIP::ssf146.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("1234567890").Click @@ script infofile_;_ZIP::ssf147.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Update & Next").Click @@ script infofile_;_ZIP::ssf148.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Update").WaitProperty "text", "Update", 10000
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'File Upload
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_2").Click @@ script infofile_;_ZIP::ssf160.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("documentCategoryCode").Select "B.01 - KTP" @@ script infofile_;_ZIP::ssf161.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("documentStatus").Select "O - TBO" @@ script infofile_;_ZIP::ssf162.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("docNumber").Set "1" @@ script infofile_;_ZIP::ssf163.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("5").Click @@ script infofile_;_ZIP::ssf164.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("strReceivedDate").Set "05/02/2023"
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("strReceivedDate").Submit
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("5").Click @@ script infofile_;_ZIP::ssf165.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("strTboDueDate").Set "05/02/2023"
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("strTboDueDate").Submit
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("remark").Set "OK" @@ script infofile_;_ZIP::ssf166.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebButton("Attach").Click @@ script infofile_;_ZIP::ssf167.xml_;_
			wait(3)
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			Browser("File Upload").Page("File Upload").WebElement("File Upload").Click @@ script infofile_;_ZIP::ssf168.xml_;_
			wait(3)
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			Browser("File Upload").Page("File Upload").WebFile("txtFile").Set "C:\Users\UFT\Pictures\upload.png" @@ script infofile_;_ZIP::ssf169.xml_;_
			'Browser("File Upload").Page("File Upload").WebFile("txtFile").Submit
			wait(3)
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
'			AIUtil.SetContext Browser("creationtime:=2")
'			AIUtil.FindTextBlock("File Upl JlOad").Click
'			AIUtil.FindTextBlock("Browse").Click
'			AIUtil.SetContext Window("text:=File Upload - Mozilla Firefox", "regexpwndtitle:=Mozilla Firefox", "regexpwndclass:=MozillaWindowClass", "is owned window:=False", "is child window:=False")
'			AIUtil.SetContext Window("text:=File Upload - Mozilla Firefox", "regexpwndtitle:=Mozilla Firefox", "regexpwndclass:=MozillaWindowClass", "is owned window:=False", "is child window:=False")
'			AIUtil.FindTextBlock("Upload").Click
'			AIUtil.FindTextBlock("Completed").Click
'			AIUtil.FindTextBlock("Close").Click
'			AIUtil.SetContext Browser("creationtime:=1")
			
			Browser("File Upload").Page("File Upload").WebButton("Upload").Click @@ script infofile_;_ZIP::ssf170.xml_;_
			wait(3)
			'Browser("File Upload").Page("File Upload").WebElement("progressBarText").Click @@ script infofile_;_ZIP::ssf171.xml_;_
			If AIUtil.FindTextBlock("Completed").Exist Then
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
				Browser("File Upload").Page("File Upload").Sync
				Browser("File Upload").Page("File Upload").WebButton("Close").Click
				'Browser("File Upload").Page("File Upload").Sync
				'Browser("File Upload").CloseAllTabs
			Else
				'Capture Screen
				CaptureScreen()
				'Capture Report
				Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
				Browser("File Upload").Page("File Upload").Sync
				Browser("File Upload").Page("File Upload").WebButton("Close").Click
				'Browser("File Upload").Page("File Upload").Sync
				'Browser("File Upload").CloseAllTabs
			End  If
			
			'Popup
			'If AIUtil.FindTextBlock("Resend").Exist Then
				'AIUtil.FindTextBlock("Resend").Click
				wait(3)
				Browser("Browser").HandleDialog micOK
			'End  If
			
			'Browser("Browser").Page("Digital Credit Management").Sync
			'If Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("View").Exist Then			
			'End  If
			
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("View").WaitProperty "text", "View", 10000
			
			'Browser("File Upload").CloseAllTabs @@ hightlight id_;_2819226_;_script infofile_;_ZIP::ssf172.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("View").Click @@ hightlight id_;_2819226_;_script infofile_;_ZIP::ssf174.xml_;_
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Update").WaitProperty "text", "Update", 10000
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("BAPAK AA BB CC").Click @@ script infofile_;_ZIP::ssf149.xml_;_
			Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Update").Click @@ script infofile_;_ZIP::ssf153.xml_;_
			wait(3)
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			Browser("Browser").HandleDialog micOK @@ hightlight id_;_1181018_;_script infofile_;_ZIP::ssf154.xml_;_
			wait(3)
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Browser("Browser").Page("Digital Credit Management").Frame("main_2").WebElement("0423010000181").Click @@ script infofile_;_ZIP::ssf155.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_2").Link("Supp.Doc.").Click @@ script infofile_;_ZIP::ssf156.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_2").Link("Simplified Data Entry").Click @@ script infofile_;_ZIP::ssf157.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link").Click @@ script infofile_;_ZIP::ssf158.xml_;_
			'Browser("Browser").Page("Digital Credit Management").Frame("main_3").WebElement("fancybox-close").Click @@ script infofile_;_ZIP::ssf159.xml_;_
			
			'Logout
			Browser("Browser").Page("Digital Credit Management").Sync
			Browser("Browser").Page("Digital Credit Management").Link("Link").Click
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			Browser("Browser").HandleDialog micOK
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			'Browser("Browser").Page("Retail Credit System").WebElement("Welcome").Click
			'Verifikasi Logout
			If Browser("Browser").Page("Retail Credit System").WebButton("Login").Exist(0) Then
				reporter.ReportEvent micPass, "Result Verifikasi Logout Succes", "Logout Succes"
			Else
				reporter.ReportEvent micFail, "Result Verifikasi Logout Failed", "Logout Failed"
			End If
			
			'Capture Screen
			CaptureScreen()
			'Capture Report
			Browser("micclass:=Browser").Page("micclass:=Page").CaptureReport micPass, "Page Screenshot."
			
			If OutputVerif = "Failed" Then
				'GoTo NextData
			End If
			
			'Create Result
			'CreateFolder()
			CreateFolderResult()
			CreateSubFolder(TESTSCRIPT)
			'CreateWordFile(TESTSCRIPT)
			CreateWordFileResult(TESTSCRIPT)
'			'CreatePDFFile(TESTSCRIPT)
			MoveFileResult(TESTSCRIPT)
			MoveImageResult(TESTSCRIPT)
		End If
	Next
End If

'Close Browser
Browser("CreationTime:=0").Sync
Browser("CreationTime:=0").Close
Browser("CreationTime:=0").Sync
Browser("CreationTime:=0").Close

'Goto Nextdata
'NextData:
If Err.Number<>0 Then
	params("MainTimer") = MainTimer
	params("SubMainTimer") = SubMainTimer
	params("GetFullFolderName") = PutSubFolderName
	params("counterRepeat") = counterRepeat
	params("limitRepeat") = limitRepeat
	'Validate_Error(params)
	OutputVerif = params("OutputVerif")
	GetLastError = params("GetLastError")
End If
