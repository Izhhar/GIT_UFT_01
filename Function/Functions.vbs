'------------------------------------------------------------------------------------------------------------------------------------
'Function Name                    : SearchTextAndClickLink(ByVal TestObject, ByVal SearchText, ByVal SearchTextInColumn, ByVal LinkColumn)
'Function Description             : Function to click a link in a web table based on value present in same or another column in the same row.
'Data Parameters                  : TestObject:- Specify the WebTable. eg: Browser("-----").Page("-----").WebTable("-----")
'                                   SearchText:- Specify the text to search in the WebTable. eg: "Cruises"
'                                   SearchTextInColumn:- Specify the column in which text is to be searched as long. eg: 2
'                                   LinkColum:- Specify the column where the link resides in the WebTable as long. eg: 3
'------------------------------------------------------------------------------------------------------------------------------------
Function SearchTextAndClickLink(ByVal TestObject, ByVal SearchText, ByVal SearchTextInColumn, ByVal LinkColumn)
 'in case of any errors
   On Error Resume Next
   'checking whether the object exist
   ObjectExist = TestObject.Exist(10)
   If ObjectExist Then
    'setting MatchFound as false
    MatchFound = FALSE
    'finding the total number of rows in the web table
    TotalRows = TestObject.RowCount
    'looping in the web table until a match is found
    For CRow = 1 to TotalRows
     'checking whether Searched Text is found
     If SearchText = Trim(TestObject.GetCellData(CRow, SearchTextInColumn)) Then
      'saving the row where the value is found to a variable
      FoundRow = CRow
      'setting MatchFound as true
      MatchFound = TRUE
      Exit For
     End If
    Next
    'in case a match is found
    If MatchFound Then
     'setting the object link
   Set LinkObject = TestObject.ChildItem(FoundRow, LinkColumn, "Link", 0)
   'clicking the link
   LinkObject.Click
   'reporting successful link click
   Reporter.ReportEvent micPass, "Specified Link Clicked Successfully.", "Successfully clicked the link " & LinkObject.GetROProperty("Text") & "."
   'in case match is not found
       ElseIf Not MatchFound Then
    'reporting error if specified text is not found in the web table
   Reporter.ReportEvent micFail, "Cannot Find Specified Text", "Cannot find the text " & SearchText & " in the object specified in function call."
    End If    
   ElseIf Not ObjectExist Then
   'reporting error if the specified object is not found
  Reporter.ReportEvent micFail, "Cannot Find The Object", "Object specified in function call cannot be found."
   End If
End Function

'Usage
'Set HomeTable = Browser("Browser").Page("Page").WebTable("Sl. No")
'Call SearchTextAndClickLink(HomeTable, "Mahesh", 2, 3)

'-------------------------------------------------------------------------------------
'Test Reporting
'-------------------------------------------------------------------------------------
Function ReportStep(ByVal TestStep, ByVal DescriptionText, ByVal Status, ByVal Screenshot)
	If Screenshot Then
		Screenshot = CaptureScreen(TestStep & "-" & DescriptionText)
		Reporter.ReportEvent micDone, TestStep, "Screenshot save at"& Screenshot
	End If
	If Status  = True Then
		Reporter.ReportEvent micPass, TestStep, DescriptionText
	ElseIf Status = False Then
		Reporter.ReportEvent micFail, TestStep, DescriptionText
	Else
		Reporter.ReportEvent micDone, TestStep, DescriptionText
	End If
End Function

'-------------------------------------------------------------------------------------
'Error Handling
'-------------------------------------------------------------------------------------
Function ReportStep(ByVal TestStep, ByVal DescriptionText, ByVal Status, ByVal Screenshot)
	If Screenshot Then
		Screenshot = CaptureScreen(TestStep & "-" & DescriptionText)
		Reporter.ReportEvent micDone, TestStep, "Screenshot save at"& Screenshot
	End If
	If Status  = True Then
		Reporter.ReportEvent micPass, TestStep, DescriptionText
	ElseIf Status = False Then
		Reporter.ReportEvent micFail, TestStep, DescriptionText
	Else
		Reporter.ReportEvent micDone, TestStep, DescriptionText
	End If
End Function

Sub MinimizeUFTWindow ()
    Set qtApp = getObject("","QuickTest.Application")
    qtApp.WindowState = "Minimized"
    Set qtApp = Nothing
End Sub
