Public Sub RCS(params)
	'Declaration
	
	'Input Parameter
	
	USERNAME = params("USERNAME")
	PASSWORD = params("PASSWORD")
	
	'Click Menu
	'Browser("Browser").Page("Digital Credit Management").Image("menu2").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main").Image("menu2").Click
	Browser("Browser").Page("Digital Credit Management").Image("menu2").Click
	
	'Capture Screen
	CaptureScreenBlank()
	Browser("Browser").Page("Digital Credit Management").Link("New Application - New").Click
	'Capture Screen
	CaptureScreenBlank()
	
	'Wait for
	'Browser("Browser").Page("Digital Credit Management").Frame("main").WebElement("New Application - New").WaitProperty "innertext", "New Application - New", 10000
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("New Application - New").WaitProperty "innertext", "New Application - New", 10000
	
	'Capture Screen
	CaptureScreenBlank()
	
	'Sales Name
	'Browser("Browser").Page("Digital Credit Management").Frame("main").WebEdit("salesCode1Desc").Set "00019"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("salesCode1Desc").Set SALES_NAME
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672816455260").Link("00019").Click
	'wait(1)
	Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(SALES_NAME).Click
	'AIUtil.FindTextBlock("00019").Click
	'AIUtil.FindTextBlock(SALES_NAME).Click
	'AIUtil.SetContext Window("regexpwndtitle:=Digital Credit Management System — Mozilla Firefox", "regexpwndclass:=MozillaWindowClass", "is owned window:=False", "is child window:=False")
	'AIUtil.FindTextBlock("00019", micFromBottom, 1).Click
	'AIUtil.FindTextBlock(SALES_NAME, micFromBottom, 1).Click
	
	'Branch
	'Browser("Browser").Page("Digital Credit Management").Frame("main_2").WebEdit("originatorBranchDescription").Set "AABTOJ"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("originatorBranchDescription").Set BRANCH
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_2").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672817098182").Link("AABTOJ").Click
	'wait(1)
	Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(BRANCH).Click
	
	'Channel
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("channelDescription").Set "AH"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("channelDescription").Set CHANNEL
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_3").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820281009").Link("AH").Click
	Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(CHANNEL).Click
	
	'Marketing Code
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("marketingCodeDescription").Set "AARG"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("marketingCodeDescription").Set MARKETING_CODE
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_4").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820296604").Link("AARG").Click
	Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(MARKETING_CODE).Click
	
	'Campaign Code 1
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode1").Select "BCC - Bundling Payroll & CC"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode1").Select CAMPAIGN_CODE_1
	
	'Campaign Code 2
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode2").Select "KR3 - Bundling Mortgage dan CC (Non CP)"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode2").Select CAMPAIGN_CODE_2
	
	'Campaign Code 3
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode3").Select "000000 - Non Program"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("campaignCode3").Select CAMPAIGN_CODE_3
	
	'Capture Screen
	CaptureScreenBlank()
	
	'First Name
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstName").Set "asd"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstName").Set FIRST_NAME
	
	'MIddle Name
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("midName").Set "asd"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("midName").Set MIDDLE_NAME
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("id1box").Click
	
	'Last Name
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("lastName").Set "asd"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("lastName").Set LAST_NAME
	
	'Full Name
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("genFullNameWithoutAbbrLine1").Set "asd"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("genFullNameWithoutAbbrLine1").Set FULL_NAME
	
	'Gender
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("genderCode").Select "F - Female"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("genderCode").Select GENDER
	
	'Place of Birth
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("placeOfBirth").Set "jkt"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("placeOfBirth").Set PLACE_OF_BIRTH
	
	'ID Type
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("firstIdentificationTypeCode").Select "KT2 - KTP Seumur Hidup"
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("3").Click
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("firstIdentificationTypeCode").Select ID_TYPE
	
	'Set Number Area
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("mobileNumberAreaCode").Set "08"
	'Browser("Browser").HandleDialog micOK
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("mobileNumberAreaCode").Set "62"
	'Browser("Browser").HandleDialog micOK
	
	'City
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_5").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820520763").WebButton("Close").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("townDesc").Set "0111"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("townDesc").Set CITY
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("id1box_2").Click
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_5").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820535716").WebElement("0111").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820535716").Link("0111").Click
	Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frameLink").Link(CITY).Click
	
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("townDesc").Set "0001"
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("id1box_3").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_5").Click
	
	'Kode Tujuan Permintaan			
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("requestPurpose").Select "01 - Penilaian calon debitur"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("requestPurpose").Select KODE_TUJUAN_PERMINTAAN
				
	'Date of Birth
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("4").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("dateOfBirth").Set "04/01/2023"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("dateOfBirth").Set DATE_OF_BIRTH
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("dateOfBirth").Submit
	
	'ID No
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstIdentificationNumber").Set "12345"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstIdentificationNumber").Click
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("firstIdentificationNumber").Set ID_NOMOR
				
	'Address
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("addressLine1").Set "alamat"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("addressLine1").Set ADDRESS
	
	'Employment Category
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("employmentCategoryCode").Select "F - PROFESIONAL"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebList("employmentCategoryCode").Select EMPLOYMENT_CATEGORY
	
	'Capture Screen
	CaptureScreenBlank()
	
	'Is Bundling
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebRadioGroup("bundlingFlag").Select "Y"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebRadioGroup("bundlingFlag").Select IS_BUNDLING
	
	wait(5)
	'Dim mySendKeys
	set mySendKeys = CreateObject("WScript.shell")
	mySendKeys.SendKeys("{PGDN}")
	
	'Bundling Code
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("bundlingTypeCode").Set "BPP1"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("bundlingTypeCode").Set BUNDLING_CODE
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Link_6").Click
	wait(5)
	Browser("Browser").Page("Digital Credit Management").Frame("fancybox-frame1672820629562").WebButton("Go").Click
	wait(5)
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
	
	AIUtil.SetContext Window("regexpwndtitle:=Digital Credit Management System — Mozilla Firefox", "regexpwndclass:=MozillaWindowClass", "is owned window:=False", "is child window:=False")
	'AIUtil.FindTextBlock("BPP1", micFromBottom, 1).Click
	AIUtil.FindTextBlock(BUNDLING_CODE, micFromBottom, 1).Click
	'AIUtil.FindTextBlock("BPP1").Click
	'AIUtil.FindTextBlock(BUNDLING_CODE).Click
	
	'Card Limit
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("cardLimit").Set "800000"
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebEdit("cardLimit").Set CARD_LIMIT
	
	'Capture Screen
	CaptureScreenBlank()
	
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Is Bundling * Yes No Group").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Is Bundling * Yes No Group").Click
	
	'Click Proceed
	Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").Link("Proceed").Click
	
	'Capture Screen
	CaptureScreenBlank()
	
	'Verifikasi
	If Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("We are currently unable").Exist(0) Then
		reporter.ReportEvent micPass, "Result Verifikasi Succes", "Succes"
	Else
		reporter.ReportEvent micFail, "Result Verifikasi Failed", "Failed"
	End If
	
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("We are currently unable").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("If the problem persists,").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Reference No : ED7F5DED2D49").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("We apologise for the inconveni").Click
	'Browser("Browser").Page("Digital Credit Management").Frame("main_new application - new").WebElement("Thank you.").Click

	
End Sub
