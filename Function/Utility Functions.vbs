'===UTILITY FUNCTION===
'-------------------------------------------------------------------------------------
'Capture Screen Shots
'-------------------------------------------------------------------------------------
Public Sub CaptureScreen()
	Dim strFileName, SCREENSHOT_COUNTER, CaptureScreen
	SCREENSHOT_COUNTER = 1
	'CurrentTime = Day(Now)&""&Month(Now)&""&Year(Now)&""&Hour(Now)&""&Minute(Now)&""&Second(Now)
	CurrentTime = DotNetFactory.CreateInstance("System.DateTime").Now.ToString("yyyyMMddHHmmss")
	
	'Create Foder Screenshot
'	Dim FolderPath
'	'FolderPath = "C:\Result"
'	FolderPath = Environment.Value("ResultDir")& "\" & "Result"
'	Set ODirectory = DotNetFactory.CreateInstance("System.IO.Directory","System")
'	If ODirectory.Exists(FolderPath) Then
'		'Msgbox "Folder Already Exists"
'	'CreateFolder = false
'	else
'	ODirectory.CreateDirectory FolderPath
'	'CreateFolder = true
'	End If
	
	'MsgBox(Environment.Value("ResultDir"))
	strFile = Environment.Value("ResultDir") & "\" & "Result" & "\" & CurrentTime & ".png"
	'Desktop.CaptureBitmap strFileName, True
	Desktop.CaptureBitmap strFile, True
	'CaptureScreen = strFileName
	CaptureScreen = strFile
End Sub

'-------------------------------------------------------------------------------------
'Capture Screen Shots UFT Report
'-------------------------------------------------------------------------------------
Public Sub CaptureReport( ByRef Sender, ByVal micStatus, ByVal descriptionStr )  
	Dim dateTimeNow, fileNameStr, divDesc, caption
	Dim dicMetaDescription, qtp
	dateTimeNow = DotNetFactory.CreateInstance( "System.DateTime" ).Now.ToString("yyyyMMddHHmmss")
	
	Dim FolderPath
	'FolderPath = "C:\Result"
	FolderPath = Environment.Value("ResultDir")& "\" & "Result" & "\" & "CaptureReport"
	Set ODirectory = DotNetFactory.CreateInstance("System.IO.Directory","System")
	If ODirectory.Exists(FolderPath) Then
		'Msgbox "Folder Already Exists"
		'CreateFolder = false
	Else
		ODirectory.CreateDirectory FolderPath
		'CreateFolder = true
	End If
	
	fileNameStr = Environment.Value("ResultDir") & "\" & "Result" & "\" & "CaptureReport" & "\" & dateTimeNow & ".png"
	
	Set qtp = CreateObject( "QuickTest.Application" )
	qtp.Visible = False
	
	Wait 0, 500
	
'	If IsObject( sender ) Then      
'		Sender.CaptureBitmap fileNameStr, True
'		caption = Sender.ToString & " - Capture Bitmap"
'	Else
		Desktop.CaptureBitmap fileNameStr, True
		caption = "Desktop - Capture Bitmap"
'	End If
	
	qtp.Visible = True
	
	divDesc =	"<table align='center' border='5' cellpadding='1' cellspacing='1' width='100%' title='" & fileNameStr & "' frame='hsides'>" & _ 
				"<caption>" & caption & "</caption>" & _ 
				"<thead><tr><th>RCS</th></tr></thead>" & _ 
				"<tfoot><tr><td align='center'><img border='2px' src='" & fileNameStr & "' /></td></tr></tfoot>" & _ 
				"<tbody><tr><td>" & descriptionStr & "</td></tr></tbody></table>"
	
	Set dicMetaDescription = CreateObject( "Scripting.Dictionary" )
	dicMetaDescription( "Status" ) = micStatus
	dicMetaDescription( "PlainTextNodeName" ) = "RCS"
	dicMetaDescription( "StepHtmlInfo" ) = "<DIV align=center>" & divDesc & "</DIV>"
	dicMetaDescription( "DllIconIndex" ) = 205
	dicMetaDescription( "DllIconSelIndex" ) = 205
	dicMetaDescription( "DllPAth" ) = EnVironment( "ProductDir" ) & "\bin\ContextManager.dll"
	Call Reporter.LogEvent( "User", dicMetaDescription, Reporter.GetContext )
	
'	Dim fso
'	Set fso = CreateObject("Scripting.FileSystemObject")
'	fso.DeleteFile(fileNameStr)
End Sub

Public Sub CaptureScreenshot( ByRef Sender, ByVal micStatus, ByVal descriptionStr )  
	Dim dateTimeNow, fileNameStr, divDesc, caption
	Dim dicMetaDescription, qtp
	dateTimeNow = DotNetFactory.CreateInstance( "System.DateTime" ).Now.ToString( "ddMMyyHHmmss" )
	fileNameStr = Reporter.ReportPath & "\" & dateTimeNow & ".png"
	Set qtp = CreateObject( "QuickTest.Application" )
	qtp.Visible = False
	
	Wait 0, 500
	
	If IsObject( sender ) Then      
		Sender.CaptureBitmap fileNameStr, True
		caption = Sender.ToString & " - Capture Bitmap"
	Else
		Desktop.CaptureBitmap fileNameStr, True
		caption = "Desktop - Capture Bitmap"
	End If
	
	qtp.Visible = True
	
	divDesc =	"<table align='center' border='5' cellpadding='1' cellspacing='1' width='100%' title='" & fileNameStr & "' frame='hsides'>" & _ 
				"<caption>" & caption & "</caption>" & _ 
				"<thead><tr><th>Application Exception Description</th></tr></thead>" & _ 
				"<tfoot><tr><td align='center'><img border='2px' src='" & fileNameStr & "' /></td></tr></tfoot>" & _ 
				"<tbody><tr><td>" & descriptionStr & "</td></tr></tbody></table>"
	
	Set dicMetaDescription = CreateObject( "Scripting.Dictionary" )
	dicMetaDescription( "Status" ) = micStatus
	dicMetaDescription( "PlainTextNodeName" ) = "ApplicationException"
	dicMetaDescription( "StepHtmlInfo" ) = "<DIV align=center>" & divDesc & "</DIV>"
	dicMetaDescription( "DllIconIndex" ) = 205
	dicMetaDescription( "DllIconSelIndex" ) = 205
	dicMetaDescription( "DllPAth" ) = EnVironment( "ProductDir" ) & "\bin\ContextManager.dll"
	Call Reporter.LogEvent( "User", dicMetaDescription, Reporter.GetContext )
End Sub

'-------------------------------------------------------------------------------------
'Create Folder
'-------------------------------------------------------------------------------------
'@description This function will create folders, this function will accept two parameters. First is Folder path and second is the folder name
'Public Function CreateFolder(Byval FolderPath)
Public Function CreateFolder()
	Dim FolderPath
	'FolderPath = "C:\Result"
	FolderPath = Environment.Value("ResultDir")& "\Result"
	Set ODirectory = DotNetFactory.CreateInstance("System.IO.Directory","System")
	If ODirectory.Exists(FolderPath) Then
		'Msgbox "Folder Already Exists"
		'CreateFolder = false
	Else
		ODirectory.CreateDirectory FolderPath
		'CreateFolder = true
	End If
End Function

'-------------------------------------------------------------------------------------
'Create Folder Result
'-------------------------------------------------------------------------------------
Public Function CreateFolderResult()
	Dim FolderPath
	'FolderPath = "C:\Result"
	FolderPath = Environment.Value("ResultDir")& "\Result"
	Set ODirectory = DotNetFactory.CreateInstance("System.IO.Directory","System")
	If ODirectory.Exists(FolderPath) Then
		'Msgbox "Folder Already Exists"
		'CreateFolder = false
	Else
		ODirectory.CreateDirectory FolderPath
		'CreateFolder = true
	End If
End Function

'-------------------------------------------------------------------------------------
'Create Sub Folder
'-------------------------------------------------------------------------------------
Public Function CreateSubFolder(TESTSCRIPT)
	Dim FolderPath
	'FolderPath = "C:\Result\"& TESCRIPT &""
	FolderPath = Environment.Value("ResultDir")& "\Result\"& TESTSCRIPT &""
	Set ODirectory = DotNetFactory.CreateInstance("System.IO.Directory","System")
	If ODirectory.Exists(FolderPath) Then
		'Msgbox "Folder Already Exists"
		'CreateFolder = false
	Else
		ODirectory.CreateDirectory FolderPath
		'CreateFolder = true
	End If
End Function

'-------------------------------------------------------------------------------------
'Create Word File Result
'-------------------------------------------------------------------------------------
Public Function CreateWordFileResult(TESTSCRIPT)
	Set objWord = CreateObject("Word.Application")
	
	'Run Woird In background
	objWord.Visible = False
	
	'Close Alert
	objWord.DisplayAlerts = False
	
	Set objDoc = objWord.Documents.Add()
	Set objSelection = objWord.Selection
	
	' Folder to process
	'strFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
	'strFolder = Environment.Value("ResultDir")
	strFolder = Environment.Value("ResultDir") & "\" & "Result"
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Access the folder to process
	Set objFolder = objFSO.GetFolder(strFolder)
	'strScreenshot = Environment.Value("ResultDir")& "\Result\Screenshot\"
	'Set objFolder = objFSO.GetFolder(strScreenshot)
	
	For Each objFile In objFolder.Files
		' Only import PNG files
		If LCase(objFSO.GetExtensionName(objFile)) = LCase("PNG") Then
			objSelection.InlineShapes.AddPicture (objFile.Path)
			objSelection.TypeText (vbCrLf)   
			objselection.InsertCaption "Figure", vbTab & objFSO.GetBaseName(objFile) , "", wdCaptionPositionBelow
			objSelection.TypeText (vbCrLf)   
			objSelection.TypeText (vbCrLf) 
		Else
			'  Wscript.Echo "No PNG files in """ & objFile.Path & """"
		End If
	Next
	
	DOCXFilePath = objFSO.BuildPath(objFolder, ""&TESTSCRIPT&".docx")
	'DOCXFilePath = objFSO.BuildPath(strFolder,""&TESTSCRIPT&".docx")
	'objDoc.Save(DOCXFilePath)
	objDoc.SaveAs(DOCXFilePath)
	'objDoc.ActiveDocument.SaveAs(DOCXFilePath)
	'objDoc.vbYes
	objDoc.Close()
	'objDoc.ActiveDocument.Close()
	objWord.Quit()
	Set obj = Nothing
	Set objDoc = Nothing
	Set objWord = Nothing
	Set objFSO = Nothing
	SystemUtil.CloseProcessByName("WINWORD.exe")
	Wait(2)
End Function

'-------------------------------------------------------------------------------------
'Move File Result
'-------------------------------------------------------------------------------------
Public Function MoveFileResult(TESTSCRIPT)
	Dim objFso, FilePath, MoveFilePath	
	FilePath = Environment.Value("ResultDir") & "\" & "Result" & "\" & TESTSCRIPT & ".docx"
	MoveFilePath = Environment.Value("ResultDir") & "\" & "Result" & "\" & TESTSCRIPT & "\"
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	'If objFso.FileExists(strSourcePath) then
	'	objFso.MoveFile strSourcePath, strDestPath
	'End If
	If objFso.FileExists(FilePath) then
		objFso.MoveFile FilePath, MoveFilePath
	End If
	Set objFso = Nothing
End Function

'-------------------------------------------------------------------------------------
'Move Image Result
'-------------------------------------------------------------------------------------
Public Function MoveImageResult(TESTSCRIPT)
	'Dim objFso, FilePath, MoveFilePath
	Dim FilePath, MoveFilePath, objFSO
	FilePath = Environment.Value("ResultDir") & "\" & "Result" & "\"
	MoveFilePath = Environment.Value("ResultDir") & "\" & "Result" & "\" & TESTSCRIPT & "\"
	
	'Set objFso = CreateObject("Scripting.FileSystemObject")
	'If objFso.FileExists(strSourcePath) then
	'	objFso.MoveFile strSourcePath, strDestPath
	'End If
'	If objFso.FileExists(FilePath) then
'		objFso.MoveFile FilePath, MoveFilePath
'	End If
	'Set objFso = Nothing
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'strFolder = Environment.Value("ResultDir") & "\" & "Result"
	
	' Access the folder to process
	Set objFolder = objFSO.GetFolder(FilePath)
	'strScreenshot = Environment.Value("ResultDir")& "\Result\Screenshot\"
	'Set objFolder = objFSO.GetFolder(strScreenshot)
	
	For Each objFile In objFolder.Files
		' Only import PNG files
		If LCase(objFSO.GetExtensionName(objFile)) = LCase("PNG") Then
			FilePath = objFile.Path
			If objFSO.FileExists(FilePath) then
				objFSO.MoveFile FilePath, MoveFilePath
			End If
		Else
			'Wscript.Echo "No PNG files in """ & objFile.Path & """"
		End If
	Next
	
	Set objFSO = Nothing
End Function

'-------------------------------------------------------------------------------------
'Move Folder Result
'-------------------------------------------------------------------------------------
Public Function MoveFolderResult(TESTSCRIPT)
	Dim objFso, FilePath, MoveFilePath	
	FilePath = Environment.Value("ResultDir") & "\" & "Result" & "\" & TESTSCRIPT & ".docx"
	MoveFilePath = Environment.Value("ResultDir") & "\" & "Result" & "\" & TESTSCRIPT & "\"
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	'If objFso.FileExists(strSourcePath) then
	'	objFso.MoveFile strSourcePath, strDestPath
	'End If
	If objFso.FileExists(FilePath) then
		objFso.MoveFile FilePath, MoveFilePath
	End If
	Set objFso = Nothing
End Function

Public Function DeleteFileResult(TESTSCRIPT)
	Dim objFso, FilePath, MoveFilePath	
	FilePath = Environment.Value("ResultDir") & "\" & "Result" & "\" & TESTSCRIPT & ".docx"
	MoveFilePath = Environment.Value("ResultDir") & "\" & "Result" & "\" & TESTSCRIPT & "\"
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	'If objFso.FileExists(strSourcePath) then
	'	objFso.MoveFile strSourcePath, strDestPath
	'End If
	If objFso.FileExists(FilePath) then
		objFso.MoveFile FilePath, MoveFilePath
	End If
	Set objFso = Nothing
End Function

Function CreateWordFile(TESTSCRIPT)
    Set objWord = CreateObject("Word.Application")
'Run Woird In background
	objWord.Visible = False
	
    'Close Alert
    objWord.DisplayAlerts = False
    
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

' Folder to process
'Dim ScriptFullName
'ScriptFullName = Environment.Value("ResultDir") & "\" & "Result" & "\"
'ScriptFullName = Environment.Value("ResultDir") & "\"
'strFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strFolder = Environment.Value("ResultDir")
'strFolder = Environment.Value("ResultDir")& "\Result\" & TESTSCRIPT &""
'strFolder = Environment.Value("ResultDir") & "\" & "Result" & "\"
'strFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(ScriptFullName)

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Access the folder to process
Set objFolder = objFSO.GetFolder(strFolder)
'strScreenshot = Environment.Value("ResultDir")& "\Result\Screenshot\"
'Set objFolder = objFSO.GetFolder(strScreenshot)

For Each objFile In objFolder.Files

    ' Only import PNG files
    If LCase(objFSO.GetExtensionName(objFile)) = LCase("PNG") Then
       objSelection.InlineShapes.AddPicture (objFile.Path)
       objSelection.TypeText (vbCrLf)   
       objselection.InsertCaption "Figure", vbTab & objFSO.GetBaseName(objFile) , "", wdCaptionPositionBelow
        objSelection.TypeText (vbCrLf)   
       objSelection.TypeText (vbCrLf) 
    Else
    '  Wscript.Echo "No PNG files in """ & objFile.Path & """"
    End If

Next

DOCXFilePath = objFSO.BuildPath(objFolder, ""&TESTSCRIPT&".docx")
'DOCXFilePath = objFSO.BuildPath(strFolder,""&TESTSCRIPT&".docx")
'objDoc.Save(DOCXFilePath)
objDoc.SaveAs(DOCXFilePath)
'objDoc.ActiveDocument.SaveAs(DOCXFilePath)
'objDoc.vbYes
objDoc.Close()
'objDoc.ActiveDocument.Close()
objWord.Quit()
    Set obj = Nothing
    Set objDoc = Nothing
    Set objWord = Nothing
    Set objFSO = Nothing
    SystemUtil.CloseProcessByName("WINWORD.exe")
    Wait(2)
End Function

'Public Sub Validate_Error(params)
'	Capture(GetFullFolderName & "\" & stdImageExtension)
'	'Verif
'	'tidak error, err.description, err.description
'	
'	'output
'	'params("GetLastError") = "Error #" & str(Err.Number) & " " & Err.Description & "was generated by " & Err.Source & Chr(13) & Err.Description
'	params("GetLastError") = "Error #" & Err.Number & ". " & Err.Description
'	params("OutputVerif") = OutputVerif
'End Sub
