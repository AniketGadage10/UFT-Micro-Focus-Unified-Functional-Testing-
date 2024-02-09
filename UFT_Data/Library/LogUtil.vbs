Option Explicit

'=======================================================================================================================================================
' Function List
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'						Function Name											|					Created By
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'1. Fn_Create_Folder(strFolderPath)												|	Samir Thosar (samir.thosar@siemens.com)
'2. Fn_Share_Folder(strCompName, strFolderPath, strShareName)					|	Samir Thosar (samir.thosar@siemens.com)
'3. Fn_Update_EnvXML(sFilePath, sEnvVal)										|	Archana Deshpande (archana.deshpande@siemens.com)
'4. Fn_WriteLogFile()															|	Vallari Shimpukade (vallari.shimpukade@siemens.com)
'5. Fn_CreateLogFile(sLogFileName)												|	Vallari Shimpukade (vallari.shimpukade@siemens.com)
'6. Fn_Create_SummarySheet														|	Samir Thosar (samir.thosar@siemens.com)
'7. Fn_Update_TestDetail														|	Samir Thosar (samir.thosar@siemens.com)
'8. Fn_Update_TestResult														|	Samir Thosar (samir.thosar@siemens.com)
'9. Fn_Update_SetupDetail														|	Samir Thosar (samir.thosar@siemens.com)
'10. Fn_UpdateLogFiles(sTestLogComment, sBatchLogComment)						|	Vallari Shimpukade (vallari.shimpukade@siemens.com)
'11. Fn_QTPSettings																|	Vallari Shimpukade (vallari.shimpukade@siemens.com)
'12. Fn_PrintQARTLog															|	Vallari Shimpukade (vallari.shimpukade@siemens.com)
'13. Fn_UpdateEnvXMLNode														|	Samir Thosar (samir.thosar@siemens.com)
'14. Fn_ExcelSearch																|	Samir Thosar (samir.thosar@siemens.com)
'15. Fn_GetBatchName															|	Samir Thosar (samir.thosar@siemens.com)
'16. Fn_ExcelSearchForFail														|	Samir Thosar (samir.thosar@siemens.com)
'17. Fn_strUtil_SubField														|	Samir Thosar (samir.thosar@siemens.com)
'18. Fn_ExcelGetResultDetail													|	Samir Thosar (samir.thosar@siemens.com)
'19. Fn_BatchDuration															|	Samir Thosar (samir.thosar@siemens.com)
'20. Fn_BatchDate																|	Samir Thosar (samir.thosar@siemens.com)	
'21. Fn_ExcelColumnHide															|	Samir Thosar (samir.thosar@siemens.com)
'22. Fn_GetXMLNodeValue															|	Samir Thosar (samir.thosar@siemens.com)
'23. Fn_BatchResultCleanup														| 	Samir Thosar (samir.thosar@siemens.com)
'24. Fn_GetTestArea																|	Samir Thosar (samir.thosar@siemens.com)
'25. Fn_GetEnvValue																| 	Samir Thosar (samir.thosar@siemens.com)
'27. Fn_KillProc																|	Mallikarjun Mastamardi (mallikarjun.mastamardi@siemens.com)
'28. Fn_SaveItemDetailsInDataTable												|	Koustubh Watwe (koustubh.watwe.ext@siemens.com)
'29. fn_splm_util_file_operation												|	Samir Thosar (samir.thosar@siemens.com)
'30. fn_splm_util_testduration													|	Samir Thosar (samir.thosar@siemens.com)
'31. fn_splm_excel_update_testduration											|	Samir Thosar (samir.thosar@siemens.com)								
'32. Fn_SetEnvValue																|	Sunny Ruparel (sunny.ruparel.ext@siemens.com)	
'33. Fn_LogUtil_GetXMLPath														|	Sandeep Navghane (sandeep.navghane.ext@siemens.com)
'34. fn_SISW_util_folder_operation												|	Shweta Rathod (shwetambari.rathod.ext@siemens.com)	
'35. Fn_ExcelGetQARTStatusDetail												|	Shweta Joshi (shweta.joshi@siemens.com)	
'=======================================================================================================================================================

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 1                                                                                
' Function Name     	: Fn_Create_Folder
' Function Description  : Create folder
' Function Usage    	: Result = Fn_Create_Folder(strFolderPath)
'							strFolderPath	- Path for folder creation on machine
'                     		return Folder Path on success
'--------------------------------------------------------------------------------------------------------------------
Public Function Fn_Create_Folder(strFolderPath)

	Dim objFSO
	Dim objFolder

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FolderExists(strFolderPath) Then
		Set objFolder = objFSO.GetFolder(strFolderPath)
	Else
		Set objFolder = objFSO.CreateFolder(strFolderPath)
	End If

	Set objFolder = Nothing
	Set objFSO = Nothing

	Fn_Create_Folder = strFolderPath

End Function


Public Function Fn_Share_Folder(strCompName, strFolderPath, strShareName)

	Const FILE_SHARE = 0
	Const MAXIMUM_CONNECTIONS = 25
	Dim objWMIService, objNewShare, errReturn
	Dim objFSO

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FolderExists(strFolderPath) Then

		Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strCompName & "\root\cimv2")
		Set objNewShare = objWMIService.Get("Win32_Share")
		errReturn = objNewShare.Create(strFolderPath, strShareName, FILE_SHARE, MAXIMUM_CONNECTIONS, strShareName)

	End if

	Fn_Share_Folder = errReturn

End Function


Public Function Fn_Update_EnvXML(sFilePath, sEnvVal )
   
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Dim fso, f, ts,TextStreamTest
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	Set f = fso.GetFile(sFilePath)
	
	Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)
	TextStreamTest =   ts.ReadAll
	
	TextStreamTest = Replace(TextStreamTest, "Empty",sEnvVal)  
	ts.Close
	Set ts = f.OpenAsTextStream(ForWriting, TristateUseDefault)
	ts.Write TextStreamTest
	ts.Close
	Fn_Update_EnvXML = True

End Function

'*********************************************************		Function to Write in Log File .		***********************************************************************
'Function Name		:				Fn_WriteLogFile(sFileName, sText)

'Description			 :		 		 Writes log in log file.

'Parameters			   :	 			1. sFileName :  Name of the file to write in
'													 2.	sText	- Text to be written to the file

'Return Value		   : 			Nothing

'Pre-requisite			:		 	Nothing

'Examples				:			Call Fn_WriteLogFile("C:\Automation\test.log,"Lgoin Successful")

'History					 :		
'	Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vallari		 													   22/04/2010			              1.0										Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WriteLogFile(sFileName, sText)
   Dim objFSO, objFile

	On Error Resume Next 
	
	if lcase(Environment.Value("DetailLog")) = "true" Then
		Set objFSO = CreateObject("Scripting.FileSystemObject")	
		
		If sFileName = "" Then
			sFileName = Environment.Value("TestLogFile")
		End If

		Set objFile = objFSO.OpenTextFile(sFileName,8)

		objFile.Write sText
		objFile.Write vblf

		Set objFile = Nothing
		Set objFSO = Nothing
	End If

End Function
'*********************************************************		Function to Creates Log File and folder.		***********************************************************************
'Function Name		:				Fn_CreateLogFile(sLogFileName)

'Description			 :		 		 Creates test log file under Batch File locatio

'Parameters			   :	 			1. sLogFileName : Name with which batch file needs to be creatde

'Return Value		   : 			sFileName : Returns File Path

'Pre-requisite			:		 	Nothing

'Examples				:			Call Fn_CreateLogFile(Environment.Value("TestName"))

'History					 :		
'	Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari		 													   22/04/2010			              1.0										Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_CreateLogFile(sLogFileName)
	Dim objFSO, objLogFile, bReturn
	Dim sBatchFldr, sFilePath
	Dim sAutoDir, sBatchFilePath, sBatchFile, sBatchFullPath, objFile, objShell

	On Error Resume Next 

	sBatchFldr = Environment.Value("BatchFldName")

    Set objFSO = CreateObject("Scripting.FileSystemObject")	

	sFilePath = sBatchFldr + "\" + sLogFileName
	If not (objFSO.FileExists(sFilePath)) Then
		Set objLogFile = objFSO.CreateTextFile(sFilePath)
	Else
		objFSO.DeleteFile sFilePath,True
		Set objLogFile = objFSO.CreateTextFile(sFilePath)
	End If
	
	
	'Progress Monitor Launch: Start:

	If CBool(Environment.Value("ProgressMonitorStatus")) = True Then

		'Kill Process of Pre-Existing Consoles:
		bReturn = Fn_KillProc("ProgressMonitor.EXE")
		bReturn = Fn_KillProc("cmd.exe")	
	
		'Initialise Requisite Variables:
		sAutoDir = Environment.Value("sPath")
		sUtilityLibPath = sAutoDir	 + "\Utilities"	
		sBatchFile = "ProgressMonitor.bat"
		sBatchFullPath = sUtilityLibPath + "\" + sBatchFile
	
		If not (objFSO.FileExists(sBatchFullPath)) Then	
			Set objFile = objFSO.CreateTextFile(sBatchFullPath, True)
			objFile.WriteLine ("Set PATH=%PATH%;" &sUtilityLibPath)
			objFile.WriteLine ("ProgressMonitor -f " & sFilePath)
			objFile.Close
			'Execute Batch File:
			Set objShell = CreateObject("WScript.Shell")
			objShell.Run "cmd /c " & sBatchFullPath, 8, False
			
			Set objShell = Nothing
		Else
			objFSO.DeleteFile(sBatchFullPath)
			Set objFile = objFSO.CreateTextFile(sBatchFullPath, True)
			objFile.WriteLine ("Set PATH=%PATH%;" &sUtilityLibPath)
			objFile.WriteLine ("ProgressMonitor -f " & sFilePath)
			objFile.Close
			'Execute Batch File:
			Set objShell = CreateObject("WScript.Shell")
			objShell.Run "cmd /c " & sBatchFullPath, 8, False
			Set objShell = Nothing		
		End If
		
	End If
					
	'Progress Monitor Launch: End	

	Fn_CreateLogFile = sFilePath

	Set objFSO = Nothing
	Set objLogFile = Nothing
End Function

'**********************************************************************************************************************
' Function Number   	: 6                                                                                
' Function Name     	: Fn_Create_SummarySheet
' Function Description  : Create batch result excel
' Function Usage    	: bReturn = Fn_Create_SummarySheet(strFolderLocation)
'							strFolderLocation	- Location of batch result folder
'                     	return batch excel path on success, False on failuer
' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  9-June-2010	| 	2.0	 |									|				|
'**********************************************************************************************************************

Public Function Fn_Create_SummarySheet(strFolderLocation)

   Const xlCenter = -4108
   Const xlLeft = -4131
   Const xlUnderlineStyleNone = -4142
   Const xlAutomatic = -4105
   Const xlNone = -4142
   Const xlContinuous = 1
   Const xlThin = 2
   Const xlDiagonalDown = 5
   Const xlDiagonalUp = 6
   Const xlEdgeLeft = 7
   Const xlEdgeTop = 8
   Const xlEdgeBottom = 9
   Const xlEdgeRight = 10
   Const xlInsideVertical = 11 
   Const xl2003=56

   Dim objExcel
   Dim objWorkBook
   Dim objWorkSheet
   Dim objRange
   Dim strExcelPath
   Dim strExcelVersion

	On Error Resume Next

	Set objExcel = CreateObject("Excel.Application")

	If Err.Number = 0 Then

		strExcelVersion = objExcel.Version
		
		Set objWorkBook = objExcel.Workbooks.Add
		
		If Err.Number <> 0 Then
			Fn_Create_SummarySheet = False
			Exit Function
		End If
			
		objExcel.Visible = False
		objExcel.DisplayAlerts = False 
		
		Set objWorkBook = objExcel.Application.ActiveWorkbook
		
		If objWorkbook.Worksheets.Count > 2 Then
			Do While objWorkbook.Worksheets.Count > 2
				objWorkBook.Worksheets(objWorkbook.Worksheets.Count).delete
			Loop
		Elseif objWorkbook.Worksheets.Count < 2 Then
			objExcel.ActiveWorkbook.Worksheets.Add
		End If
		
		Set objWorkSheet = objWorkBook.WorkSheets(1)
		
			objWorkSheet.Name = "Test Details"
		
			objWorkSheet.Cells(1,1).Value = "Sr No."
			objWorkSheet.Cells(1,2).Value = "Product"
			objWorkSheet.Cells(1,3).Value = "Release"
			objWorkSheet.Cells(1,4).Value = "Area"
			objWorkSheet.Cells(1,5).Value = "Category"
			objWorkSheet.Cells(1,6).Value = "Test Case"
			objWorkSheet.Cells(1,7).Value = "Date"    
			objWorkSheet.Cells(1,8).Value = "Result"    
			objWorkSheet.Cells(1,9).Value = "Comments"    
			objWorkSheet.Cells(1,10).Value = "Logs"
			objWorkSheet.Cells(1,11).Value = "Test Duration"
			objWorkSheet.Cells(1,12).Value = "Start Time"
			objWorkSheet.Cells(1,13).Value = "End Time"
			objWorkSheet.Cells(1,14).Value = "QART Upload Status"
			objWorkSheet.Cells(1,15).Value = "Failed Function Name"
			
			objWorkSheet.Columns(1).ColumnWidth = 10
			objWorkSheet.Columns(2).ColumnWidth = 10
			objWorkSheet.Columns(3).ColumnWidth = 20
			objWorkSheet.Columns(4).ColumnWidth = 35
			objWorkSheet.Columns(5).ColumnWidth = 25
			objWorkSheet.Columns(6).ColumnWidth = 30
			objWorkSheet.Columns(7).ColumnWidth = 20
			objWorkSheet.Columns(8).ColumnWidth = 10
			objWorkSheet.Columns(9).ColumnWidth = 30
			objWorkSheet.Columns(10).ColumnWidth = 70
			objWorkSheet.Columns(11).ColumnWidth = 25
			objWorkSheet.Columns(12).ColumnWidth = 25
			objWorkSheet.Columns(13).ColumnWidth = 25
			objWorkSheet.Columns(14).ColumnWidth = 25
			objWorkSheet.Columns(15).ColumnWidth = 35
			
			Set objRange = objWorkSheet.Range("A1:O1")
			objRange.HorizontalAlignment = xlCenter
			objRange.VerticalAlignment = xlCenter 
			objRange.Font.Name = "Arial"
			objRange.Font.Size = "12"
			objRange.Font.Bold = True
			objRange.Interior.ColorIndex = "37"
			objRange.Borders.LineStyle = xlContinuous   
		
		Set objWorkSheet = objWorkBook.WorkSheets(2)
		
			objWorkSheet.Name = "Setup Details"
		
			objWorkSheet.Cells(1,1).Value = "Teamcenter Release"
			objWorkSheet.Cells(2,1).Value = "Build"
			objWorkSheet.Cells(3,1).Value = "TC Server Host"
			objWorkSheet.Cells(4,1).Value = "OS Version"
			objWorkSheet.Cells(5,1).Value = "Application Server"
			objWorkSheet.Cells(6,1).Value = "TC DB Host"
			objWorkSheet.Cells(7,1).Value = "Database Type"    
			objWorkSheet.Cells(8,1).Value = "Database Version"    
			objWorkSheet.Cells(9,1).Value = "Client Host"    
			objWorkSheet.Cells(10,1).Value = "Client OS Version"
			objWorkSheet.Cells(11,1).Value = "OTW Location"
		
			objWorkSheet.Columns(1).ColumnWidth = 30
			objWorkSheet.Columns(2).ColumnWidth = 60
		
			Set objRange = objWorkSheet.UsedRange
			objRange.HorizontalAlignment = xlLeft
			objRange.VerticalAlignment = xlCenter 
			objRange.Font.Name = "Arial"
			objRange.Font.Size = "12"
			objRange.Font.Bold = True
			objRange.Interior.ColorIndex = "37"
			objRange.Borders.LineStyle = xlContinuous
		
		Set objWorkSheet = objWorkBook.WorkSheets(1)
		objWorkSheet.Activate
		Set objRange = objWorkSheet.UsedRange
		objRange.Range("A1").Activate
		
		strExcelPath = strFolderLocation & "\BatchRunDetails.xlsx"
		
		'If strExcelVersion = "12.0" Then
		'	objWorkBook.SaveAs(strExcelPath), xl2003
		'Else
			objWorkBook.SaveAs(strExcelPath)
		'End If
		objExcel.Quit
		
		Set objRange = Nothing
		Set objWorkSheet = Nothing
		Set objWorkBook = Nothing
		Set objExcel = Nothing
		
		Fn_Create_SummarySheet = strExcelPath
	Else
		Fn_Create_SummarySheet = False
	End If

End Function

'**********************************************************************************************************************
' Function Number   	: 7                                                                                
' Function Name     	: Fn_Update_TestDetail
' Function Description  : Update test details in batch result excel
' Function Usage    	: bReturn = Fn_Update_TestDetail(strResultSheetLocation, arrTestData(), intSheetNumber)
'							strResultSheetLocation	- Location of batch result excel
'							arrTestData				- Test array data 
'							intSheetNumber			- Sheet number in excel
'                     	return True on success, False on failuer
' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  9-June-2010	| 	2.0	 |									|				|
'**********************************************************************************************************************

Public Function Fn_Update_TestDetail(strResultSheetLocation, arrTestData(), intSheetNumber)

	Const xlCellTypeLastCell = 11
	Const xlContinuous = 1
	Const xlCenter = -4108
	Const xlLeft = -4131	

	Dim objFile
	Dim objExcel
	Dim objWorkbook
	Dim objWorkSheet
	Dim objRange
	Dim iCount
    Dim sLinkAddress
	Dim objFSO
	Dim intRowId
	Dim sBatchSetupFilePath,sBatchFIlePath,bReturn,sSearchStr,objWorkbookBatchRun,objWorksheetBatchRun,objfind,iRow
	On Error Resume Next

	If  strResultSheetLocation = "" Then
		strResultSheetLocation=Environment.Value("BatchFldName") +"\BatchRunDetails.xlsx"
	End If


		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(strResultSheetLocation) Then

			Set objExcel = CreateObject("Excel.Application")

			objExcel.Visible = False
			objExcel.AlertBeforeOverwriting = False
			objExcel.DisplayAlerts = False
	
			Set objWorkbook = objExcel.Workbooks.Open(strResultSheetLocation)
	
			If intSheetNumber = "" Then
				 intSheetNumber = 1
			End If   
	
			Set objWorkSheet = objWorkbook.Worksheets(intSheetNumber)
			objWorkSheet.Activate
			
			Set objRange = objWorkSheet.UsedRange
			objRange.SpecialCells(xlCellTypeLastCell).Activate
			
			intRowId = objExcel.ActiveCell.Row + 1
	
			objWorkSheet.Cells(intRowId, 1).Value = intRowId - 1
			objExcel.Cells(intRowId, 1).HorizontalAlignment = xlCenter
			objExcel.Cells(intRowId, 1).VerticalAlignment = xlCenter
			
			sBatchSetupFilePath = Fn_LogUtil_GetXMLPath("batchsetup")
			sBatchFIlePath= Fn_GetXMLNodeValue(sBatchSetupFilePath ,"BatchFilePath" )
			bReturn = fn_splm_util_file_operation("exist", sBatchFIlePath)
			If bReturn = True Then
				sSearchStr = Environment.Value("TestName")
				Set objWorkbookBatchRun = objExcel.Workbooks.Open(sBatchFIlePath)
				Set objWorksheetBatchRun = objWorkbookBatchRun.Worksheets(1)
				objWorksheetBatchRun.Activate
				set objfind = objWorksheetBatchRun.Range("H:H").Find(sSearchStr)
				wait 1
				iRow = objfind.Row
				arrTestData(3) = objWorksheetBatchRun.cells(iRow, 5 )
				arrTestData(4) = objWorksheetBatchRun.cells(iRow, 6 )
				set objWorksheetBatchRun = nothing
				objWorkbookBatchRun.Close
				set objWorkbookBatchRun = nothing
			End If
				
			For iCount = 1 to UBound(arrTestData)
				objWorkSheet.Cells(intRowId, (iCount + 1)).Value = arrTestData(iCount)
			Next
			'QART Upload Status
			objWorkSheet.Cells(intRowId, 14).Font.Bold = True
			objWorkSheet.Cells(intRowId, 14).HorizontalAlignment = xlCenter
			objWorkSheet.Cells(intRowId, 14).VerticalAlignment = xlCenter	
			objWorkSheet.Cells(intRowId, 14).Value = "Not Uploaded"
			objWorkSheet.Cells(intRowId, 14).Interior.Color = RGB(255,128,128)
			
			objRange.Borders.LineStyle = xlContinuous
			objRange.Range("A1").Activate
	
			objWorkbook.Save
			objExcel.Quit

			Fn_Update_TestDetail = True

		Else
			Fn_Update_TestDetail = False
		End If

    Set objRange = Nothing
    Set objWorkSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
	Set objFSO = Nothing

End Function

'**********************************************************************************************************************
' Function Number   	: 8                                                                                
' Function Name     	: Fn_Update_TestResult
' Function Description  : Update test result in batch result excel
' Function Usage    	: bReturn = Fn_Update_TestResult(strResultSheetLocation, sResultData, intSheetNumber)
'							strResultSheetLocation	- Location of batch result excel
'							sResultData				- Test result data (PASS / FAIL and Verification Point Details) 
'							intSheetNumber			- Sheet number in excel
'                     	return True on success, False on failuer
' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  9-June-2010	| 	2.0	 |									|				|
'**********************************************************************************************************************

Public Function Fn_Update_TestResult(strResultSheetLocation, sResultData, intSheetNumber)

	Const xlCellTypeLastCell = 11
	Const xlContinuous = 1
	Const xlCenter = -4108

	Dim objFile
	Dim objExcel
	Dim objWorkbook
	Dim objWorkSheet
	Dim objRange
	Dim iCount
	Dim objFSO
	Dim sColumnId
	Dim arrResult, sShardPath, sBatFldPath
	Dim arrResultData(2)
	Dim bCommentFlag:bCommentFlag=False
		If Instr(sResultData,"|")>0 Then
			sResultData = Replace(sResultData,"|",":")
		End If
		arrResult = Split(sResultData, ":", -1,1)
		If Ubound(arrResult) > 1 Then

			arrResultData(0) = arrResult(0)
			arrResultData(1) = ""
            
			For iCount = 1 to Ubound(arrResult)
				arrResultData(1) = arrResultData(1) & arrResult(iCount) & " "
			Next
		Else
			arrResultData(0) = arrResult(0)
			arrResultData(1) = arrResult(1)
		End If

		If  strResultSheetLocation = "" Then
			strResultSheetLocation=Environment.Value("BatchFldName") +"\BatchRunDetails.xlsx"
		End If

		sColumnId = Fn_ExcelSearch(strResultSheetLocation, "Result", intSheetNumber)
		sColumnId = CInt(Fn_strUtil_SubField(sColumnId, ":", 1))

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(strResultSheetLocation) Then

			Set objExcel = CreateObject("Excel.Application")
			objExcel.AlertBeforeOverwriting = False
	
			objExcel.Visible = False
			objExcel.DisplayAlerts = False
	
			Set objWorkbook = objExcel.Workbooks.Open(strResultSheetLocation)
	
			If intSheetNumber = "" Then
				 intSheetNumber = 1
			End If  
	
			Set objWorkSheet = objWorkbook.Worksheets(intSheetNumber)
			objWorkSheet.Activate
			
			Set objRange = objWorkSheet.UsedRange
			objRange.SpecialCells(xlCellTypeLastCell).Activate
			For iCount = 0 to (UBound(arrResultData)-1)
				If bCommentFlag=True And GBL_EXPECTED_MESSAGE <> "" And GBL_ACTUAL_MESSAGE <>"" Then
					objWorkSheet.Cells(objExcel.ActiveCell.Row, sColumnId).Value ="Fail to verify message :  Actual Message:  "& GBL_ACTUAL_MESSAGE &"  Expected Message: "&GBL_EXPECTED_MESSAGE
					objExcel.Cells(objExcel.ActiveCell.Row, sColumnId).Interior.Color = RGB(255,255,0)
					GBL_EXPECTED_MESSAGE=""
					GBL_ACTUAL_MESSAGE=""
					bCommentFlag=False
				Else
					objWorkSheet.Cells(objExcel.ActiveCell.Row, sColumnId).Value = Trim(arrResultData(iCount))
				End If
				If iCount = 0 Then
					objExcel.Cells(objExcel.ActiveCell.Row, sColumnId).Font.Bold = True
					objExcel.Cells(objExcel.ActiveCell.Row, sColumnId).HorizontalAlignment = xlCenter
					objExcel.Cells(objExcel.ActiveCell.Row, sColumnId).VerticalAlignment = xlCenter	
					If InStr(LCase(arrResultData(iCount)),"pass") <> 0 Then
						objExcel.Cells(objExcel.ActiveCell.Row, sColumnId).Interior.ColorIndex = 35
					ElseIf InStr(LCase(arrResultData(iCount)),"fail") <> 0 Then
						bCommentFlag=True
						objExcel.Cells(objExcel.ActiveCell.Row, sColumnId).Interior.Color = RGB(255,128,128)
						If GBL_FAILED_FUNCTION_NAME<>"" Then
							objWorkSheet.Cells(objExcel.ActiveCell.Row, 15).Value = Trim(GBL_FAILED_FUNCTION_NAME)
							objExcel.Cells(objExcel.ActiveCell.Row, 15).Font.Bold = True
							GBL_FAILED_FUNCTION_NAME = ""
						End If
					End If				
				End If
				sColumnId = sColumnId + 1
			Next
			sShardPath = "\\" & Environment.Value("LocalHostName") & "\" & Fn_GetBatchName(Environment.Value("BatchFldName"))
			objWorkSheet.Cells(objExcel.ActiveCell.Row, sColumnId).Value = sShardPath + "\" + Environment.Value("TestName") +".log"

			objRange.Borders.LineStyle = xlContinuous
			objRange.Range("A1").Activate
	
			objWorkbook.Save
			objExcel.Quit

			Fn_Update_TestResult = True

		Else
			Fn_Update_TestResult = False

		End IF

    Set objRange = Nothing
    Set objWorkSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
	Set objFSO = Nothing

End Function

'**********************************************************************************************************************
' Function Number   	: 9                                                                                
' Function Name     	: Fn_Update_SetupDetail
' Function Description  : Update Result sheet with the test setup details
' Function Usage    	: Result = Fn_Update_SetupDetail(strResultSheetLocation, arrSetupData(), intSheetNumber)
'							strResultSheetLocation	- Location of Result excel on test machine
'							arrSetupData			- Test setup details
'							intSheetNumber			- Sheet number in excel
'                     return True on success, False on failuer
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  9-June-2010	| 	2.0	 |									|				|
'**********************************************************************************************************************


Public Function Fn_Update_SetupDetail(strResultSheetLocation, arrSetupData(), intSheetNumber)

	Const xlCellTypeLastCell = 11
	Const xlContinuous = 1
	Const xlLeft = -4131
	Const xlCenter = -4108

	Dim objFile
	Dim objExcel
	Dim objWorkbook
	Dim objWorkSheet
	Dim objRange
	Dim iCount
	Dim objFSO

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(strResultSheetLocation) Then

			Set objExcel = CreateObject("Excel.Application")
			objExcel.AlertBeforeOverwriting = False
			objExcel.DisplayAlerts = False
			objExcel.Visible = False
	
			Set objWorkbook = objExcel.Workbooks.Open(strResultSheetLocation)
	
			If intSheetNumber = "" Then
				 intSheetNumber = 1
			End If 
	
			Set objWorkSheet = objWorkbook.Worksheets(intSheetNumber)
			objWorkSheet.Activate
					
			For iCount = 0 to UBound(arrSetupData)
				objWorkSheet.Cells((iCount + 1), 2).Value = arrSetupData(iCount)
			Next
	
			Set objRange = objWorkSheet.UsedRange
			objRange.SpecialCells(xlCellTypeLastCell).Activate
			objRange.HorizontalAlignment = xlLeft
			objRange.VerticalAlignment = xlCenter 
			objRange.Font.Name = "Arial"
			objRange.Font.Size = "12"
			objRange.Borders.LineStyle = xlContinuous
			Set objWorkSheet = objWorkBook.WorkSheets(1)
			objWorkSheet.Activate
	
			objWorkbook.Save
			objExcel.Quit

			Fn_Update_SetupDetail = True

		Else

			Fn_Update_SetupDetail = False

		End If

    Set objRange = Nothing
    Set objWorkSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
End Function

'**********************************************************************************************************************
' Function Number   	: 10                                                                                
' Function Name     	: Fn_UpdateLogFiles()
' Function Description  : Update log file and batch excel
' Function Usage    	: bReturn = Fn_UpdateLogFiles(sTestLogComment, sBatchLogComment)
'							sTestLogComment	- Log statements to be entered in test log file
'							sBatchLogComment - Log statements to be entered in batch excel file
'             				
' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  9-June-2010	| 	2.0	 |									|				|
'**********************************************************************************************************************
Public Function Fn_UpdateLogFiles(sTestLogComment, sBatchLogComment)
   
	Dim objFSO, objFile, sFileName
   
	On Error Resume Next 
   
	If sTestLogComment <> "" Then
	
		Set objFSO = CreateObject("Scripting.FileSystemObject")	
	
		sFileName = Environment.Value("TestLogFile")
	
		Set objFile = objFSO.OpenTextFile(sFileName,8)
	
		objFile.WriteLine sTestLogComment
	
		Set objFile = Nothing
		Set objFSO = Nothing
	
	End If
		
	If sBatchLogComment <> "" Then
		Call Fn_Update_TestResult("", sBatchLogComment, 1)
	End If

End Function
'*********************************************************		Function to  set the qtp setting***********************************************************************
'Function Name		:				Fn_QTPSettings

'Description			 :		 		This function is used to set the qtp setting.
'                                                   Tools->Options->Java : Tree path value is set to :
'                                                   File->Setting->Run  When error occur during  run session proceed to next step.
'							Note : In this function Tools->Options->Run  Checkbox of Allow other HP products to run tests and components need to check manually. Its not handled as it is not supported by QTP for security reason.
'Parameters			   :	 		  NA

'Return Value		   : 			NA

'Pre-requisite			:		    NA

'Examples				:		  Fn_QTPSettings()

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	Reviewed Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							29-Mar-2010		1.0																	Santosh			30-Mar-10
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
Public Function Fn_QTPSettings()
	Dim objAppSetng
	Set objAppSetng = CreateObject("QuickTest.Application")

	'Tools->Options->Java : Tree path value is set to :
	objAppSetng.Options.Java.TreePathSeparator = ":"
	'File->Setting->Run  When error occur during  run session proceed to next step.
	objAppSetng.Test.Settings.Run.OnError ="NextStep"
	set objAppSetng = nothing
End Function

'*************************************************************************************************************************************************
' Function Number   	: 12                                                                                
' Function Name     	: Fn_PrintQARTLog()
' Function Description  : Prepares the log file for test case
' Function Usage    	: bReturn = Fn_PrintQARTLog()
'							BatchFolderPath	- Location of Excel
'             				
' Function History
'------------------------------------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done									|	Reviewer	|	Reviewed Date
'------------------------------------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  9-June-2010	| 	2.0	 |															|				|
'------------------------------------------------------------------------------------------------------------------------------------------------
' 	Shweta Rathod		| 26-Aug-2016	| 	1.0	 |  Added code to print Briefcase browser information		| Koustubh Watwe | 26-Aug-2016
'*************************************************************************************************************************************************

Public Function Fn_PrintQARTLog()

	Dim bReturn,sParamName,iParamCount,i

	bReturn = Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
	bReturn = Fn_UpdateLogFiles("Setup Details", "")
	bReturn = Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
	bReturn = Fn_UpdateLogFiles("Teamcenter Release" & Fn_Tab(1) & " - " & Environment.Value("TcRelease"), "")
	bReturn = Fn_UpdateLogFiles("Teamcenter Build" & Fn_Tab(1) & " - " & Environment.Value("TcBuild"), "")
	bReturn = Fn_UpdateLogFiles("Teamcenter Server Host" &  "   - " & Environment.Value("TcServer"), "")
	bReturn = Fn_UpdateLogFiles("Teamcenter RAC Host" & Fn_Tab(1) &	" - " & Environment.Value("LocalHostName"), "")
	bReturn = Fn_UpdateLogFiles("Tester Name" & Fn_Tab(3) & " - " & Environment.Value("UserName"), "")
	'Start of Modification - Added code to print the briefcase browser version in log file [Briefcase Browser 11.2.3 (P11000.2.0.30_20160608.00) 64-bit]
	If lcase(Environment.Value("BBFlag")) = "true" then
		bReturn = Fn_UpdateLogFiles("Briefcase Browser Build" & space(1) & " - " & Environment.Value("BBVersion"), "")
		bReturn = Fn_UpdateLogFiles("NX Version" & space(14) & " - " & Environment.Value("NXVersion"), "")
	End if
	'End of Modification - Added code to print the briefcase browser version in log file [Briefcase Browser 11.2.3 (P11000.2.0.30_20160608.00) 64-bit]
	bReturn = Fn_UpdateLogFiles(vblf, "")
	bReturn = Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
	bReturn = Fn_UpdateLogFiles("Test Case Information", "")
	bReturn = Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
	bReturn = Fn_UpdateLogFiles("QART Root" & Fn_Tab(3) & " - " & Environment.Value("QARTRoot"), "")
	bReturn = Fn_UpdateLogFiles("QART Release" & Fn_Tab(2) & " - " & Environment.Value("QARTRelease"), "")
	bReturn = Fn_UpdateLogFiles("QART Feature" & Fn_Tab(2) & " - " & DataTable("Feature", dtGlobalSheet), "")
	bReturn = Fn_UpdateLogFiles("QART Category" & Fn_Tab(2) & " - " & DataTable("Category", dtGlobalSheet), "")
    iParamCount = DataTable.GetSheet("Global").GetParameterCount
	For i = 1 To iParamCount
		sParamName = DataTable.GetSheet("Global").GetParameter(i).Name
		If strComp("Version",sParamName)=0 Then
			bReturn = Fn_UpdateLogFiles("QART Version" & Fn_Tab(2) & " - " & DataTable("Version", dtGlobalSheet), "")
		End if 
	Next
	bReturn = Fn_UpdateLogFiles("QART TestCase" & Fn_Tab(2) & " - " & Environment.Value("TestName"), "")
	bReturn = Fn_UpdateLogFiles("Test Run Date" & Fn_Tab(2) & " - " & MonthName(Month(Date)) & "/" & Day(Date) & "/" & Year(Date), "")
	bReturn = Fn_UpdateLogFiles("Test Run Time" & Fn_Tab(2) & " - " & TimeSerial(Hour(Time) , Minute(Time), Second(Time)), "")
	'----- Set TC StartTime [PoonamC_26Dec2016]--------------
	Environment.Value("TCStartTime") = now()
	bReturn = Fn_UpdateLogFiles(vblf, "")
	bReturn = Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
	bReturn = Fn_UpdateLogFiles("Test Script Logs", "")
	bReturn = Fn_UpdateLogFiles("-----------------------------------------------------------------------------------------------", "")
	bReturn = Fn_UpdateLogFiles(vblf, "")

End Function

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 13                                                                                
' Function Name     	: Fn_UpdateEnvXMLNode
' Function Description  : Update QTP Environment XML with the BatchResult folder path
' Function Usage    	: Result = Fn_UpdateEnvXMLNode(XMLDataFile, sNodeName, sNodeValue)
'							XMLDataFile	- Location of QTP Environment XML on test machine
'							sNodeName	- Node Name in XML (e.g. BatchFldName in QTP Environment XML)
'							sNodeValue	- Node Value for the sNodeName (e.g. BatchResult folder path)
'                     return True on success, False on failuer
'--------------------------------------------------------------------------------------------------------------------

Public Function Fn_UpdateEnvXMLNode(XMLDataFile, sNodeName, sNodeValue)

Dim objXMLDoc
Dim objChildNodes
Dim objSelectNode
Dim intNodeLength
Dim intNodeCount
Dim intChildNodeCount
Dim strNodeSting

set objXMLDoc=CreateObject("Microsoft.XMLDOM")												' Create XMLDOM object
objXMLDoc.async="false"
objXMLDoc.load(XMLDataFile)																	' Loading QTP Environment XML

If (objXMLDoc.parseError.errorCode <> 0) Then
	Fn_UpdateEnvXMLNode = False
Else
	intNodeLength = objXMLDoc.getElementsByTagName("Variable").length
	For intNodeCount = 0 to (intNodeLength - 1)
		Set objChildNodes = objXMLDoc.documentElement.childNodes.item(intNodeCount).childNodes
			strNodeSting = ""
			For intChildNodeCount = 0 to (objChildNodes.length - 1)
					strNodeSting = strNodeSting & objChildNodes(intChildNodeCount).text 
			Next
			If Instr(strNodeSting, sNodeName) Then
				Set objSelectNode = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount &"]/Value")
				objSelectNode.Text = sNodeValue
				Exit For
			End If
	Next
	objXMLDoc.Save(XMLDataFile)
	Set objSelectNode = nothing 
	Set objChildNodes = nothing
	Set objXMLDoc = nothing
	Fn_UpdateEnvXMLNode = True
End if	

End Function

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 14                                                                                
' Function Name     	: Fn_ExcelSearch(sExcelPath, sSearchStr, iSheetNumber)
' Function Description  : Search for a specific string in the excel
' Function Usage    	: Result = Fn_ExcelSearch(sExcelPath, sSearchStr, intSheetNumber)
'							sExcelPath		- Location of Excel
'							sSearchStr		- Search string to be searched in Excel
'							intSheetNumber	- Excel sheet number
'                     return Cell id for the searched text (e.g. for "A1" it will return 1:1) on success, False on failuer       
'--------------------------------------------------------------------------------------------------------------------
Public Function Fn_ExcelSearch(sExcelPath, sSearchStr, iSheetNumber)

	Const xlCellTypeLastCell = 11
	
	Dim objExcel
	Dim objWorkbook
	Dim objWorksheet
	Dim iRowId, iColId
	
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)
	objExcel.Visible = false
	objExcel.DisplayAlerts = False
	
	If iSheetNumber = "" Then
		 iSheetNumber = 1
	End If 
	
	Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
	objWorksheet.Activate
	objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate

	For iRowId = 1 To objExcel.ActiveCell.Row
        For iColId = 1 to objExcel.ActiveCell.Column
			If LCase(objExcel.Cells(iRowId, iColId).Value) = LCase(sSearchStr) Then
				Fn_ExcelSearch = iRowId &":"& iColId
				Exit For
			End If
		Next
	Next

	If Fn_ExcelSearch = "" Then
		Fn_ExcelSearch = False
	End If

	objExcel.Quit

	Set objWorksheet = Nothing
	Set objWorkbook = Nothing
	Set objExcel = Nothing

End Function
'**********************************************************************************************************************
' Function Number   	: 15                                                                                
' Function Name     	: Fn_GetBatchName(BatchFolderPath)
' Function Description  : Return the folder name created runtime
' Function Usage    	: Result = Fn_GetBatchName(BatchFolderPath)
'							BatchFolderPath	- Location of Excel
'             				return Foldername on success, False on failuer
' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  30-Apr-2010	| 	1.0	 |									|				|
'**********************************************************************************************************************
Public Function Fn_GetBatchName(sFolderPath)

	Dim objFSO
	Dim BatchName

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(sFolderPath) Then
		BatchName = objFSO.GetBaseName(sFolderPath)
		Fn_GetBatchName = BatchName
	Else
		Fn_GetBatchName = False
	End If

End Function

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 16                                                                                
' Function Name     	: Fn_ExcelSearchForFail(sExcelPath, iSheetNumber)
' Function Description  : Search for a failuer in the result sheet
' Function Usage    	: Result = Fn_ExcelSearch(sExcelPath, intSheetNumber)
'							sExcelPath		- Location of Excel
'							intSheetNumber	- Excel sheet number
'                     return True on success, False on failuer     
' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  03-May-2010	| 	1.0	 |									|				|  
'--------------------------------------------------------------------------------------------------------------------
Public Function Fn_ExcelSearchForFail(sExcelPath, iSheetNumber)

	Const xlCellTypeLastCell = 11
	
	Dim objExcel
	Dim objWorkbook
	Dim objWorksheet
	Dim iRowId, iColId

	iColId = Fn_ExcelSearch(sExcelPath, "Result", iSheetNumber)
	iColId = Fn_strUtil_SubField(iColId, ":", 1)
	
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)
	objExcel.Visible = false
	objExcel.DisplayAlerts = False
	
	If iSheetNumber = "" Then
		 iSheetNumber = 1
	End If 
	
	Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
	objWorksheet.Activate
	objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate

	For iRowId = 1 To objExcel.ActiveCell.Row
		If LCase(objExcel.Cells(iRowId, CInt(iColId)).Value) = LCase("FAIL") Then
			Fn_ExcelSearchForFail = True
			Exit For
		Else
			Fn_ExcelSearchForFail = False
		End If
	Next

	objExcel.Quit

	Set objWorksheet = Nothing
	Set objWorkbook = Nothing
	Set objExcel = Nothing

End Function

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 17                                                                                
' Function Name     	: Fn_strUtil_SubField(expression, delimiter, intCount)
' Function Description  : Returns a substring
' Function Usage    	: Result = Fn_strUtil_SubField(expression, delimiter, intCount)
'							expression		- String expression containing substrings and delimiters. If expression is a zero-length string
'							delimiter		- String character used to identify substring limits
'							count 			- Number of delimiter after which substring is present in expression
'                     return Sub String on success, False on failuer     
' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  03-May-2010	| 	1.0	 |									|				|  
'--------------------------------------------------------------------------------------------------------------------

Function Fn_strUtil_SubField(expression, delimiter, intCount)

	Dim strSubString

	If len(expression) = 0 Then
		Fn_strutil_SubField = False
	ElseIf Instr(expression, delimiter) = 0 Then
		Fn_strutil_SubField = False
	Else
		strSubString = split(expression, delimiter, -1, 1)
		If Ubound(strSubString) >= intCount Then
			Fn_strutil_SubField = strSubString(intCount)
		else
			Fn_strutil_SubField = False
		End If
	End If

End Function

'--------------------------------------------------------------------------------------------------------------------
' Function Number   	: 18                                                                                
' Function Name     	: Fn_ExcelGetResultDetail(sExcelPath, sActionType, iSheetNumber)
' Function Description  : Returns required data from result sheet
' Function Usage    	: Result = Fn_strUtil_SubField(sExcelPath, sActionType, iSheetNumber)
'							sExcelPath		- Location of Excelsheet
'							sActionType		- 4 types of action supported
'													i. FailExist
'												   ii. NoTestCases
'												  iii. FailCount
'												   iv. PassCount									
'							intSheetNumber	- Excel sheet number
'                     return Sub String on success, False on failuer     
' Function History
'----------------------------------------------------------------------------------------------------------------------
'	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
'----------------------------------------------------------------------------------------------------------------------
' 	Samir Thosar		|  06-May-2010	| 	1.0	 |									|				|  
'----------------------------------------------------------------------------------------------------------------------

Public Function Fn_ExcelGetResultDetail(sExcelPath, sActionType, iSheetNumber)

	Const xlCellTypeLastCell = 11
	Dim objFSO, objExcel, objWorkbook, objWorkSheet, objRange
	Dim usedRange, iColId , PassCnt, FailCnt, TotalCnt
    
	iColId = Fn_ExcelSearch(sExcelPath, "Result", iSheetNumber)
	iColId = Fn_strUtil_SubField(iColId, ":", 1)    
	
	iColId = Chr(64 + iColId)
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(sExcelPath) Then
	
		Set objExcel = CreateObject("Excel.Application")
			objExcel.Visible = False
			objExcel.AlertBeforeOverwriting = False
			objExcel.DisplayAlerts = False
	
			Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)	

			If iSheetNumber = "" Then
				 iSheetNumber = 1
			End If 			
			
			Set objWorkSheet = objExcel.ActiveWorkbook.Worksheets(iSheetNumber)
			Set objRange = objWorkSheet.UsedRange
			objRange.SpecialCells(xlCellTypeLastCell).Activate	
			
			usedRange = iColId & "2:" & iColId & objExcel.ActiveCell.Row
			
			if  LCase(sActionType) = LCase("FailExist") Then
				objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "FAIL" & Chr(34) & ")"
				FailCnt = objExcel.Cells(10000, 1).value
				if  FailCnt > 0 then
					Fn_ExcelGetResultDetail = True
				else
					TotalCnt = objExcel.ActiveCell.Row - 1
					objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "PASS" & Chr(34) & ")"
					PassCnt = objExcel.Cells(10000, 1).value
					objExcel.Cells(10001, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "FAIL" & Chr(34) & ")"
					FailCnt = objExcel.Cells(10001, 1).value
					if  PassCnt + FailCnt <> TotalCnt then
						Fn_ExcelGetResultDetail = True
					else
						Fn_ExcelGetResultDetail = False
					end if
				end if																																			
			elseif LCase(sActionType) = LCase("NoTestCases") Then
				Fn_ExcelGetResultDetail = objExcel.ActiveCell.Row - 1				
				
			elseif LCase(sActionType) = LCase("PassCount") then			
				objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "PASS" & Chr(34) & ")"
				Fn_ExcelGetResultDetail = objExcel.Cells(10000, 1).value
				
			elseif LCase(sActionType) = LCase("FailCount") then
				TotalCnt = objExcel.ActiveCell.Row - 1
				objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "PASS" & Chr(34) & ")"
				PassCnt = objExcel.Cells(10000, 1).value
				objExcel.Cells(10001, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "FAIL" & Chr(34) & ")"
				FailCnt = objExcel.Cells(10001, 1).value
				if  PassCnt + FailCnt <> TotalCnt then
					Fn_ExcelGetResultDetail = 	TotalCnt - PassCnt
				else
					Fn_ExcelGetResultDetail = 	FailCnt																	      	
				end if								
			end if				
			
							
	objExcel.Quit			
			
    Set objRange = Nothing
    Set objWorkSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
	Set objFSO = Nothing			
			
	End If	

End Function

''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 19                                                                                
'' Function Name     	: Fn_BatchDuration(sExcelPath)
'' Function Description  : Returns test duration of the batch run
'' Function Usage    	: Result = Fn_BatchDuration(sExcelPath)
''							sExcelPath		- Location of Excelsheet
''                     return batch test duration on success, False on failuer     
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  06-May-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------

Function Fn_BatchDuration(sExcelPath)

	Dim objFSO
	Dim objFile
	Dim sfileCreate
	Dim sfileLastModified
	Dim timeDiff
	Dim timeDuration
	Dim timeTaken

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists (sExcelPath) Then

		Set objFile = objFSO.GetFile(sExcelPath)

		sfileCreate = objFile.DateCreated
		sfileLastModified = objFile.DateLastModified

		timeDiff = Formatnumber((DateDiff("s", sfileCreate, sfileLastModified)/3600), 2, 0, -1)

		If timeDiff => 1 Then
			Fn_BatchDuration = Fn_strUtil_SubField(timeDiff,".",0) & " hours, " & Formatnumber((Fn_strUtil_SubField(timeDiff,".",1) * 0.6), 0, 0, -1) & " mins"
		Else
			Fn_BatchDuration = Formatnumber((Fn_strUtil_SubField(timeDiff,".",1) * 0.6), 0, 0, -1) & " mins"
		End If
	Else
		Fn_BatchDuration = False
	End if

	Set objFSO = nothing
	Set objFile = nothing

End Function

''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 20                                                                               
'' Function Name     	: Fn_BatchDate(sExcelPath)
'' Function Description  : Returns date of the batch run
'' Function Usage    	: Result = Fn_BatchDate(sExcelPath)
''							sExcelPath		- Location of Excelsheet
''                     		return batch date on success, False on failuer     
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  20-May-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Function Fn_BatchDate(sExcelPath)

	Dim objFSO
	Dim objFile
	Dim sfileCreate
	Dim sfileLastModified
	Dim timeDiff
	Dim timeDuration
	Dim timeTaken

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FileExists (sExcelPath) Then

		Set objFile = objFSO.GetFile(sExcelPath)
		sfileCreate = objFile.DateCreated
		Fn_BatchDate = datevalue(sfileCreate)
		
	Else
		Fn_BatchDate = False
	End if

	Set objFSO = nothing
	Set objFile = nothing

End Function

''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 21                                                                               
'' Function Name     	: Fn_ExcelColumnHide(sExcelPath, sColumnRange, iSheetNumber)
'' Function Description : Function used to hide the range of columne provided by user
'' Function Usage    	: Result = Fn_ExcelColumnHide(sExcelPath, sColumnRange, iSheetNumber)
''							sExcelPath		- Location of Excelsheet
''							sColumnRange	- Range of column
''							intSheetNumber	- Excel sheet number
''                     		return True on success, False on failuer     
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  07-June-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------

Public Function Fn_ExcelColumnHide(sExcelPath, sColumnRange, iSheetNumber)
	
	Dim objFSO
	Dim objExcel
	Dim objWorkbook
	Dim objWorkSheet
	Dim objSelection
	
	On Error Resume Next
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FileExists(sExcelPath) Then
	
		If err = 0 then
			Set objExcel = CreateObject("Excel.Application")
			Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)
			objExcel.Visible = false
			objExcel.DisplayAlerts = False
		
			If iSheetNumber = "" Then
				iSheetNumber = 1
			End If 
		
			Set objWorkSheet = objWorkbook.Worksheets(iSheetNumber)
			objWorkSheet.Activate		
			
			objExcel.Range(sColumnRange).Select
			Set objSelection = objExcel.Selection
			objSelection.EntireColumn.Hidden = True
			Fn_ExcelColumnHide = True
			
			objWorkbook.Save
			objExcel.Quit
			
			Set objSelection = Nothing
			Set objWorkSheet = Nothing
			Set objWorkbook = Nothing
			Set objExcel = Nothing		
			
		Else
			
			Fn_ExcelColumnHide = False
			Set objExcel = Nothing	

		End If
		
	Else
	
		Fn_ExcelColumnHide = False
		
	End If
	
	Set objFSO = Nothing

End Function


''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 22                                                                               
'' Function Name     	: Fn_GetXMLNodeValue(XMLDataFile, sNodeName)
'' Function Description : Function used to get the value of a tag from .xml file
'' Function Usage    	: Result = Fn_GetXMLNodeValue("C:\Test.xml", "Var1")
''							XMLDataFile	- Location of xml file
''							sNodeName	- Variable name for which value needs to be fetched   
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  16-Aug-2010	| 	1.0	 |									|				|  
'' 	Samir Thosar		|  04-Apr-2012	| 	2.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Public Function Fn_GetXMLNodeValue(XMLDataFile, sNodeName)

	Dim objXMLDoc	
	Dim intNodeLength
	Dim intNodeCount	
	Dim objNodeName
	Dim objNodeVal

	set objXMLDoc=CreateObject("Microsoft.XMLDOM")												' Create XMLDOM object
	objXMLDoc.async="false"
	objXMLDoc.load(XMLDataFile)																	' Loading QTP Environment XML

	If (objXMLDoc.parseError.errorCode <> 0) Then
		Fn_GetXMLNodeValue = False
	Else
		intNodeLength = objXMLDoc.getElementsByTagName("Variable").length
		For intNodeCount = 0 to (intNodeLength - 1)
			Set objNodeName = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount &"]/Name")			
			Set objNodeVal = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount &"]/Value")
			If LCase(objNodeName.Text) = LCase(sNodeName) Then				
				Fn_GetXMLNodeValue = objNodeVal.Text
				Exit For
			ElseIF instr(LCase(objNodeVal.Text),  LCase(sNodeName) ) > 0 Then
				Fn_GetXMLNodeValue = objNodeVal.Text
				Exit For
			End IF
			Set objNodeVal = Nothing
			Set objNodeName = Nothing
		Next
		Set objNodeVal = nothing 
		Set objNodeName = nothing
		Set objXMLDoc = nothing

		If Fn_GetXMLNodeValue = "" Then
			Fn_GetXMLNodeValue = False
		End If
	
	End if	

End Function

''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 23                                                                               
'' Function Name     	: Fn_BatchResultCleanup(sExcelPath, iSheetNumber)
'' Function Description : Function used to cleanup the batch result sheet before sending the batch 
'' Function Usage    	: Result = Fn_BatchResultCleanup(sExcelPath, iSheetNumber)
''							sExcelPath		- Location of Excelsheet
''							intSheetNumber	- Excel sheet number 
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  20-Aug-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Function Fn_BatchResultCleanup(sExcelPath, iSheetNumber)

	Const xlCellTypeLastCell = 11
	Const xlCenter = -4108
	
	Dim objExcel, objWorkbook, objWorksheet, objRange
	Dim iRowId, iColId

	iColId = Fn_ExcelSearch(sExcelPath, "Result", iSheetNumber)
	iColId = CInt(Fn_strUtil_SubField(iColId, ":", 1))
    'iColId = 8

	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)
	objExcel.Visible = false
	objExcel.DisplayAlerts = false
	
	If iSheetNumber = "" Then
		 iSheetNumber = 1
	End If 	
	
	Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
	objWorksheet.Activate
	objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate
	
	Set objRange = objWorkSheet.UsedRange
	
	For iRowId = 2 To objExcel.ActiveCell.Row 
		If LCase(objExcel.Cells(iRowId, iColId).Value) <> LCase("PASS") AND LCase(objExcel.Cells(iRowId, iColId).Value) <> LCase("FAIL") then
			If Instr(LCase(objExcel.Cells(iRowId, iColId).Value), LCase("PASS")) <> 0 then
				objExcel.Cells(iRowId, iColId).Value = "PASS"
				objExcel.Cells(iRowId, iColId).Interior.ColorIndex = 35
			Else
				objExcel.Cells(iRowId, iColId).Value = "FAIL"
				objExcel.Cells(iRowId, iColId).Interior.Color = RGB(255,128,128)						
			End if		
		Else
			If LCase(objExcel.Cells(iRowId, iColId).Value) = LCase("PASS") Then
				objExcel.Cells(iRowId, iColId).Interior.ColorIndex = 35
			Else
				objExcel.Cells(iRowId, iColId).Interior.Color = RGB(255,128,128)
			End If
		End If
		objExcel.Cells(iRowId, iColId).Font.Bold = True
		objExcel.Cells(iRowId, iColId).HorizontalAlignment = xlCenter
		objExcel.Cells(iRowId, iColId).VerticalAlignment = xlCenter			
	Next
	
	objRange.Range("A1").Activate

	objWorkbook.Save
	objExcel.Quit

	Set objWorksheet = Nothing
	Set objWorkbook = Nothing
	Set objExcel = Nothing		

    Fn_BatchResultCleanup = True
    
End Function
''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 24                                                                              
'' Function Name     	: Fn_GetTestArea(sExcelPath, iSheetNumber)
'' Function Description : Function used to fetch the test area from batch result excel
'' Function Usage    	: Result = Fn_GetTestArea(sExcelPath, iSheetNumber)
''							sExcelPath		- Location of Excelsheet
''							intSheetNumber	- Excel sheet number 
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  20-Aug-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Function Fn_GetTestArea(sExcelPath, iSheetNumber)

		Const xlCellTypeLastCell = 11
		
		Dim objExcel, objWorkbook, objWorksheet, objRange
		Dim iRowId, iColId
		Dim CellValue, UniqueVal, i , j
		Dim uniqueValueArr, testAreaArr
		Dim arrtestArea()
		Dim sTestArea

		iColId = Fn_ExcelSearch(sExcelPath, "Area", iSheetNumber)
		iColId = CInt(Fn_strUtil_SubField(iColId, ":", 1))
'		iColId = 4
		
		Set objExcel = CreateObject("Excel.Application")
		Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)
		objExcel.Visible = false
		objExcel.DisplayAlerts = false
		
		If iSheetNumber = "" Then
			 iSheetNumber = 1
		End If 	
		
		Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
		objWorksheet.Activate
		objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate
		
		CellValue = ""
		UniqueVal = ""
		
		For iRowId = 2 To objExcel.ActiveCell.Row 
			 if CellValue <>  objExcel.Cells(iRowId, iColId).Value AND Instr(UniqueVal, objExcel.Cells(iRowId, iColId).Value) = 0 then
			 	UniqueVal = UniqueVal + objExcel.Cells(iRowId, iColId).Value + "/" 
			 end if
			 CellValue = objExcel.Cells(iRowId, iColId).Value
		Next               		    
	    uniqueValueArr = Split(UniqueVal, "/", -1,1)
		ReDim arrtestArea(Ubound(uniqueValueArr)-1)    
    	For i = 0 to Ubound(uniqueValueArr) - 1
    		if Instr(LCase(uniqueValueArr(i)), LCase("Teamcenter")) <> 0 then		' My Teamcenter and My Teamcenter 2007
    			arrtestArea(i) = "My Teamcenter" 
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Search")) <> 0 then		' Search
				arrtestArea(i) = "Search"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Admin")) <> 0 then		' Admin
				arrtestArea(i) = "Admin"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("RequirementsManagement")) <> 0 then	' Requirements Manager
				arrtestArea(i) = "Requirements Manager"		
			elseif Instr(LCase(uniqueValueArr(i)), LCase("WEB")) <> 0 then	' Web
				arrtestArea(i) = "Web Client"							
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Project")) <> 0 AND Instr(LCase(uniqueValueArr(i)), LCase("Program")) <> 0 then	' Project and Program Manager
				arrtestArea(i) = "Project and Program Manager"		
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Project")) <> 0 AND Instr(LCase(uniqueValueArr(i)), LCase("Work")) <> 0 then	' Projects and Work Context
				arrtestArea(i) = "Projects and Work Context"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Change")) <> 0 then	' Change Management
				arrtestArea(i) = "Change Management"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Product")) <> 0 AND Instr(LCase(uniqueValueArr(i)), LCase("Structure")) <> 0 then ' Structure Manager
				arrtestArea(i) = "Structure Manager"	
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Workflow")) <> 0 then	' Workflow
				arrtestArea(i) = "Workflow"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("ADS")) <> 0 then	' ADS
				arrtestArea(i) = "ADS"	
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Classification")) <> 0 then	' Classification
				arrtestArea(i) = "Classification"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Vendor Management")) <> 0 then	' Vendor Management
				arrtestArea(i) = "Vendor Management"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Import and Export")) <> 0 then	' Import and Export
				arrtestArea(i) = "Import and Export"																														
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Audit Manager")) <> 0 OR Instr(LCase(uniqueValueArr(i)), LCase("Subscription Manager")) <> 0 then	' Subscribe and Audit
				arrtestArea(i) = "Subscribe and Audit"																														
			elseif Instr(LCase(uniqueValueArr(i)), LCase("MRO")) <> 0 then	' MRO
				arrtestArea(i) = "MRO"		
            elseif Instr(LCase(uniqueValueArr(i)), LCase("RDV")) <> 0 then	'RDV
				arrtestArea(i) = "RDV"
            elseif Instr(LCase(uniqueValueArr(i)), LCase("Business Modeler")) <> 0 then	'Business Modeler
				arrtestArea(i) = "Business Modeler"	                		
            elseif Instr(LCase(uniqueValueArr(i)), LCase("CPD")) <> 0 OR Instr(LCase(uniqueValueArr(i)), LCase("4GD")) <> 0 then'CPD
				arrtestArea(i) = "4GD"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Mechatronics")) <> 0 then	'Mechatronics
				arrtestArea(i) = "Mechatronics"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Content")) <> 0 then	'Content Management
				arrtestArea(i) = "Content Management"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("DIPRO")) <> 0 then	'DIPRO
				arrtestArea(i) = "DIPRO"
			elseif Instr(LCase(uniqueValueArr(i)), LCase("Classic Multi-Site")) <> 0 then 'Classic Multi-Site
				arrtestArea(i) = "Classic Multi-Site"
			end if
		Next   
		
 		for j = 0 to Ubound(arrtestArea)	    
 			if  InStr(sTestArea, arrtestArea(j)) = 0 then
 		    	sTestArea = sTestArea & " " & arrtestArea(j)
			end if 		    	
 		next		
		objExcel.Quit
		
		Fn_GetTestArea = sTestArea
		
		Set objWorksheet = Nothing
		Set objWorkbook = Nothing
		Set objExcel = Nothing
		
		if Fn_GetTestArea = "" then
			Fn_GetTestArea = False
			objExcel.Quit
			Set objWorksheet = Nothing
			Set objWorkbook = Nothing
			Set objExcel = Nothing
		End if						
		
End Function
''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 25                                                                              
'' Function Name     	: Fn_GetEnvValue(sType, sEnvName)
'' Function Description : Function used to get the env values for the system
'' Function Usage    	: Result = Fn_GetEnvValue(sType, sEnvName)
''							sType		- User / System
''							sEnvName	- Name of the env variable (e.g. FMS_HOME, AutomationDir)
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  20-Sep-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Function Fn_GetEnvValue(sType, sEnvName)

	Dim objShell
	Dim objUsrEnv
	Dim UserVar
	
	UserVar = ""
	
	Set objShell = CreateObject("WScript.Shell")
	Set objUsrEnv = objShell.Environment(sType)
		UserVar = objUsrEnv(sEnvName)
    Set objUsrEnv = Nothing
	Set objShell = Nothing
	
	if UserVar <> "" then
		Fn_GetEnvValue = UserVar
	else
		Fn_GetEnvValue = False
	end if				
			

End Function

''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 26                                                                             
'' Function Name     	: Fn_Tab(tabCount)
'' Function Description : Function used for text formating 
'' Function Usage    	: Result = Fn_Tab(2)
''							tabCount	- Number of tabs require
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  20-Sep-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Function Fn_Tab(tabCount)
	Dim sTabOpr , i

	For i = 0 to tabCount
		sTabOpr = sTabOpr & vbTab
	Next
	Fn_Tab = sTabOpr
End Function

''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 26                                                                             
'' Function Name     	: LoadEnvXML()
'' Function Description : Function used load external env xml
'' Function Usage    	: Result = LoadEnvXML()

'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  20-Sep-2010	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Function LoadEnvXML()
	Dim sAutoDir
	sAutoDir = Fn_GetEnvValue("User", "AutomationDir")
	Environment.Value("sPath") = sAutoDir
	Environment.LoadFromFile(sAutoDir + "\TestData\EnvVar_Ext.xml")
	LoadEnvXML = True
End Function
''*********************************************************		Function to Kill all the Current User Processes		***********************************************************************
'Function Name		:				Fn_KillProc

'Description			 :		 		 The function is used to kill all the preferred processes owned by current user

'Parameters			   :	 			sProcessNames
											
'Return Value		   : 				Boolean

'Pre-requisite			:		 		None

'Examples				:				 Fn_KillProc("iexplore.exe:notepad.exe") 

'History:
'										Developer Name			Date					Revision			Changes Done			Reviewer	Reviewed  Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Mallikarjun		 		23-Dec-2010       		1.0					
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_KillProc(sProcessNames)

			Dim objNetwork, currUser, User, Domain                                
			Dim sArrData,  iCount, strComputer,sImgPath
			Dim objWMIService, objProcess, colProcess
										
			sArrData = split(sProcessNames, ":",-1,1)

			Set objNetwork = CreateObject("Wscript.Network")
			currUser = objNetwork.UserName

			strComputer = "." 
			Set objWMIService = GetObject("winmgmts:" _
					    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

			For iCount = 0 to ubound(sArrData)
				Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name ='"+ sArrData(iCount) +"'")  

				For Each objProcess in colProcess 
					If objProcess.GetOwner ( User, Domain ) = 0 Then
 						If LCase(User) = currUser then
							objProcess.Terminate()
						End If
					End If
				Next 
			Next

			Set objWMIService = Nothing
			Set objNetwork = Nothing
			Set colProcess = Nothing

			If Err.Number <> 0 Then
				Fn_KillProc =False
			Else
				Fn_KillProc =True
			End If
End Function
'********************************************************************************************************************************************************************************************************
'Description		 :	This function is used to set the item details in Data Table using SOA file soaoutput.xml.

'Parameters			:	1. sAction : Action to perform
'						2. sIDColName : Name of the column where Item ID is to be stored
'						3. sRevisionColName : Name of the column where Item Revision ID is to be stored
'						4. sItemNameColName : Name of the column where Item Name is to be stored
'							Default Column names are ItemID, ItemRevID, ItemName											
'Return Value		: 	True / False

'Pre-requisite		:	soaoutput.xml must be present in SOA folder

'Examples			:	Call Fn_SaveItemDetailsInDataTable("Overwrite","ItemID","ItemRevisionID","Item_Name")
'Examples			:	Call Fn_SaveItemDetailsInDataTable("Append","","","")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Examples			:	Environment.value("PLMXML_FilePath") = "c:\plm.xml" ( - File path is mandetory )
'					:	Call Fn_SaveItemDetailsInDataTable("PLMXML_Overwrite","ItemID","ItemRevisionID","Item_Name")
'					:	Call Fn_SaveItemDetailsInDataTable("PLMXML_Append","","","")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'History:
'					Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					Koustubh				    3-Jan-2011			    1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					Koustubh				    22-Jun-2011			    1.0			Modified code to extract item id, rev id, name from item definition
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					Koustubh				    20-Jan-2012			    1.0			Added code to extract item id, rev id, name from PLM XML
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SaveItemDetailsInDataTable(sAction, sIDColName, sRevisionColName, sItemNameColName)
		On error resume next
		Dim objXMLDoc, objChildNodes, objSelectNode, intNodeCount, intChildNodeCount, obj
		Dim ItemIDFlag, ItemRevIDFlag, ItemNameFlag, iRowCounter, aItemArr
		Dim iCounter, iColumnCount, aItemData, iRows,aIDArr, sDataString, i

		Fn_SaveItemDetailsInDataTable = False
		ItemIDFlag = False
		ItemRevIDFlag= False
		ItemNameFlag= False
		
		if sIDColName = "" then sIDColName = "ItemID"
		if sRevisionColName = "" then sRevisionColName = "ItemRevID"
		if sItemNameColName = "" then sItemNameColName = "ItemName"
		
		' checking existance of columns
		For iCounter = 1 to iColumnCount
			If Datatable.GetSheet("Global").GetParameter(iCounter).Name = sIDColName Then
				ItemIDFlag = True
			End If
			If Datatable.GetSheet("Global").GetParameter(iCounter).Name = sRevisionColName Then
				ItemRevIDFlag= True
			End If
			If Datatable.GetSheet("Global").GetParameter(iCounter).Name = sItemNameColName Then
				ItemNameFlag= True
			End If		
		Next
		' adding ItemID column to datatable in Global datasheet
		If ItemIDFlag = False then
			Datatable.GetSheet("Global").AddParameter sIDColName,""
		End If
		' adding ItemRevID column to datatable in Global datasheet
		If ItemRevIDFlag = False then
			Datatable.GetSheet("Global").AddParameter sRevisionColName,""
		End If
		' adding ItemName column to datatable in Global datasheet
		If ItemNameFlag = False then
			Datatable.GetSheet("Global").AddParameter sItemNameColName,""
		End If
	
		If inStr(sAction,"PLMXML_") > 0 Then
			' import from PLM XML
			set objXMLDoc=CreateObject("Microsoft.XMLDOM")												' Create XMLDOM object
			objXMLDoc.async="false"
			objXMLDoc.load(Environment.value("PLMXML_FilePath"))																	' Loading QTP Environment XML
			sDataString = ""
			If (objXMLDoc.parseError.errorCode <> 0) Then
				Fn_SaveItemDetailsInDataTable = False
				exit function
			Else
				For intNodeCount = 0 to (objXMLDoc.getElementsByTagName("ProductRevision").length)
					Set objChildNodes = objXMLDoc.documentElement.childNodes.item(intNodeCount).childNodes
					For intChildNodeCount = 0 to (objChildNodes.length - 1)
						If lcase(cstr(objChildNodes(intChildNodeCount).nodeName)) = "userdata" Then
							Set obj = objChildNodes(intChildNodeCount).childNodes
							For i = 0 to obj.length -1
								If lcase(obj(i).getAttribute("title")) = "object_string"  Then
									If sDataString = "" Then
										sDataString =  obj(i).getAttribute("value")
									Else
										sDataString = sDataString & "|" & obj(i).getAttribute("value")
									End If
									Fn_SaveItemDetailsInDataTable = True
								End If
							Next
						End If
					Next
				Next
				sDataString = sDataString & "|"
				Set objSelectNode = nothing 
				Set objChildNodes = nothing
				Set objXMLDoc = nothing
			End IF
		Else
				' checking whether structure has been generated or not
				sAutomationDir = Fn_GetEnvValue("User", "AutomationDir")
				If sAutomationDir = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SaveItemDetailsInDataTable ] Can not find environment variable AutomationDir.")
					Exit function
				End If
				If cBool(trim(Fn_GetXMLNodeValue(sAutomationDir & "\SOA\soaoutput.xml", "SOAResult"))) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SaveItemDetailsInDataTable ] Failed to generate Items using SOA.")
					Exit function
				End If
				sDataString = trim(Fn_GetXMLNodeValue(sAutomationDir&"\SOA\soaoutput.xml", "SOAData"))
		End If
		aItemArr = split(sDataString ,"|")
		Select Case sAction
				Case "Overwrite", "True", true, "PLMXML_Overwrite"
					' overwrite data
					iRowCounter = 1
					For iCounter = 0 to uBound(aItemArr ) - 1
						' repalce ;1 with empty string
						aItemArr(iCounter) = replace(aItemArr(iCounter) ,";1","")
						'split string with /
						aIDArr = split(aItemArr(iCounter),"/")
						'split second ellement with -
						aItemData = split(aIDArr (1),"-")
						Datatable.SetCurrentRow( iRowCounter)
						Datatable.Value(sIDColName, "Global") = "'" &aIDArr(0)
						Datatable.Value(sRevisionColName,"Global") = "'" &aItemData(0)
						Datatable.Value(sItemNameColName,"Global") = trim(aItemData(1))
						iRowCounter = iRowCounter + 1
					Next
					Fn_SaveItemDetailsInDataTable = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "", "False", "Append", False, "PLMXML_Append"
					' do not overwrite data
					iRows =  DataTable.GetRowCount  'rows
					For iCounter = 0 to uBound(aItemArr ) - 1
						' repalce ;1 with empty string
						aItemArr(iCounter) = replace(aItemArr(iCounter) ,";1","")
						'split string with /
						aIDArr = split(aItemArr(iCounter),"/")
						'split second ellement with -
						aItemData = split(aIDArr (1),"-")
						For iRowCounter = 1 to iRows
							Datatable.SetCurrentRow( iRowCounter)
							' mapping item names 
							If trim(Datatable.Value(sItemNameColName,"Global")) = trim(aItemData(1)) Then
								If Datatable.Value(sIDColName, "Global") = "" Then
									' saving item data to data table
									Datatable.Value(sIDColName, "Global") = "'" &aIDArr(0)
									Datatable.Value(sRevisionColName,"Global") = "'" &aItemData(0)
									Datatable.Value(sItemNameColName,"Global") = trim(aItemData(1))
									Exit for
								End If
							End If
						Next

						If iRowCounter = (iRows + 1) Then
							iRowCounter =  iRowCounter + 1
							
						End If
						 
						If iRowCounter  = (iRows + 2) Then
						'appending item data to data table
							Datatable.SetCurrentRow( iRowCounter)
							' saving item data to data table
							Datatable.Value(sIDColName, "Global") = "'" &aIDArr(0)
							Datatable.Value(sRevisionColName,"Global") = "'" &aItemData(0)
							Datatable.Value(sItemNameColName,"Global") = trim(aItemData(1))
							iRows = iRows + 1
						End If
					Next
					Fn_SaveItemDetailsInDataTable = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -'

				Case "AppendExt"
					' do not overwrite data
					iRows =  DataTable.GetRowCount  'rows
					For iCounter = 0 to uBound(aItemArr ) - 1
						' repalce ;1 with empty string
						aItemArr(iCounter) = replace(aItemArr(iCounter) ,";1","")
						'split string with /
						aIDArr = split(aItemArr(iCounter),"/")
						'split second ellement with -
						aItemData = split(aIDArr (1),"-")
						For iRowCounter = 1 to iRows
							Datatable.SetCurrentRow( iRowCounter)
							' mapping item names 
							If trim(Datatable.Value(sItemNameColName,"Global")) = trim(aItemData(1)) Then
								If Datatable.Value(sIDColName, "Global") = "" Then
									' saving item data to data table
									Datatable.Value(sIDColName, "Global") = "'" &aIDArr(0)
									Datatable.Value(sRevisionColName,"Global") = "'" &aItemData(0)
									Datatable.Value(sItemNameColName,"Global") = trim(aItemData(1))
									Exit for
								End If
							End If
						Next

						If iRowCounter = (iRows + 1) Then
							iRowCounter =  iRowCounter + 1
							
						
							Datatable.SetCurrentRow( iRowCounter)
							' saving item data to data table
							Datatable.Value(sIDColName, "Global") = "'" &aIDArr(0)
							Datatable.Value(sRevisionColName,"Global") = "'" &aItemData(0)
							Datatable.Value(sItemNameColName,"Global") = trim(aItemData(1))
							iRows = iRows + 1
						End If
					Next
					Fn_SaveItemDetailsInDataTable = True
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SaveItemDetailsInDataTable ] Invalid case [ " & sAction & " ].")
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		End Select
		If Fn_SaveItemDetailsInDataTable = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SaveItemDetailsInDataTable ] Executed successfully with  case [ " & sAction & " ].")
		End If
End Function
''---------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 29                                                                            
'' Function Name     	: fn_splm_util_file_operation(sOpCase, sFilePath)
'' Function Description : Function used to get
''						: 1. File create date & time
''						: 2. File last modified date & time
'' Function Usage    	: Result = fn_splm_util_file_operation()

'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  17-Feb-2011	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
'' 	shweta Rathod		|  31-Aug-2016	| 	1.0	 |	added case:	exist,delete		|koustubh watwe|  31-aug-16
''----------------------------------------------------------------------------------------------------------------------
Function fn_splm_util_file_operation(sOpCase, sFilePath)
	
	Dim objFSO
	Dim objFile
	Dim sReturnVal
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Select Case sOpCase
	Case "CreateDate"
		if objFSO.FileExists(sFilePath) then
			Set objFile = objFSO.GetFile(sFilePath)
				sReturnVal = objFile.DateCreated
			Set objFile = Nothing
		else
			sReturnVal = False
		end if
		Set objFile = Nothing
	Case "LastModified"
		if objFSO.FileExists(sFilePath) then
			Set objFile = objFSO.GetFile(sFilePath)
				sReturnVal = objFile.DateLastModified
			Set objFile = Nothing
		else
			sReturnVal = False
		end if
		Set objFile = Nothing
	Case "exist"
		if objFSO.FileExists(sFilePath) then
			sReturnVal = true
		else
			sReturnVal = False
		End if
	Case "delete"
		objFSO.DeleteFile sFilePath, True
		sReturnVal = true
	End Select
	Set objFSO = Nothing
	fn_splm_util_file_operation = sReturnVal  

End Function

''---------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 30                                                                            
'' Function Name     	: fn_splm_util_testduration(sCreateDate, sLastModified)
'' Function Description : Function used to get test case duration
'' Function Usage    	: Result = fn_splm_util_file_operation()

'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  17-Feb-2011	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Function fn_splm_util_testduration(sCreateDate, LastModified)

	Dim timeDiff
	Dim timeDuration
	Dim timeTaken
	Dim sRetrunVal
	
	timeDiff = Formatnumber((DateDiff("s", sCreateDate, LastModified)/3600), 2, 0, -1)

	If timeDiff => 1 Then
		sRetrunVal = Fn_strUtil_SubField(timeDiff,".",0) & " hours, " & Formatnumber((Fn_strUtil_SubField(timeDiff,".",1) * 0.6), 0, 0, -1) & " mins"
	Else
		sRetrunVal = Formatnumber((Fn_strUtil_SubField(timeDiff,".",1) * 0.6), 0, 0, -1) & " mins"
	End If

	fn_splm_util_testduration = sRetrunVal
	If Fn_strUtil_SubField(sRetrunVal, " ", 0) = "" Then
		fn_splm_util_testduration= False
	End If

End Function

''---------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 31                                                                            
'' Function Name     	: fn_splm_excel_update_testduration(sFilePath, sSheetNum)
'' Function Description : Function used update the result sheet with test duration / start time / end time
'' Function Usage    	: Result = fn_splm_excel_update_testduration()

'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Samir Thosar		|  17-Feb-2011	| 	1.0	 |									|				|  
''----------------------------------------------------------------------------------------------------------------------
Function fn_splm_excel_update_testduration(sExcelPath, iSheetNum)

	Const xlCellTypeLastCell = 11
	Const xlContinuous = 1
	Const xlCenter = -4108
	Const xlLeft = -4131	
	
	Dim objFSO
	Dim objExcel
	Dim objWorkBook
	Dim objWorkSheet
	Dim objRange
	
	Dim iRowCnt
	Dim sFilePath
	Dim sLastRowCnt
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(sExcelPath) Then
		
			Set objExcel = CreateObject("Excel.Application")
			objExcel.Visible = False
			objExcel.AlertBeforeOverwriting = False
			objExcel.DisplayAlerts = False
			
			Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)
			
			If iSheetNum = "" Then
				iSheetNum = 1
			End If			
			
			Set objWorksheet = objWorkbook.Worksheets(iSheetNum)
			objWorkSheet.Activate	
			Set objRange = objWorkSheet.UsedRange
			objRange.SpecialCells(xlCellTypeLastCell).Activate

			sLastRowCnt = objExcel.ActiveCell.Row
'			
'			objExcel.Cells(1,11).Value = "Test Duration"			
'			objExcel.Cells(1,12).Value = "Start Time"
'			objExcel.Cells(1,13).Value = "End Time"
'			
'			objWorkSheet.Columns(11).ColumnWidth = 10
'			objWorkSheet.Columns(12).ColumnWidth = 25
'			objWorkSheet.Columns(13).ColumnWidth = 25			
'			
'			For iRowCnt = 2 To objExcel.ActiveCell.Row
'				sFilePath = objExcel.Cells(iRowCnt,10).Value
'				objExcel.Cells(iRowCnt,11).Value = fn_splm_util_testduration(fn_splm_util_file_operation("CreateDate", sFilePath), fn_splm_util_file_operation("LastModified", sFilePath))												
'				objExcel.Cells(iRowCnt,12).Value = "st-" & fn_splm_util_file_operation("CreateDate", sFilePath) 								
'				objExcel.Cells(iRowCnt,13).Value = "et-" & fn_splm_util_file_operation("LastModified", sFilePath)
'			Next
			
			Set objRange = objWorkSheet.Range("K1:M1")
			objRange.HorizontalAlignment = xlCenter
			objRange.VerticalAlignment = xlCenter 
			objRange.Font.Name = "Arial"
			objRange.Font.Size = "12"
			objRange.Font.Bold = True
			objRange.Interior.ColorIndex = "37"
			objRange.Borders.LineStyle = xlContinuous
			
			Set objRange = objWorkSheet.Range("K2:M" & sLastRowCnt)
			objRange.SpecialCells(xlCellTypeLastCell).Activate
			objRange.Borders.LineStyle = xlContinuous
			objRange.HorizontalAlignment = xlLeft
			objRange.VerticalAlignment = xlCenter

			objRange.Range("A1").Activate							
						
			objWorkbook.Save
			objExcel.Quit
			
			Set objRange = Nothing
		    Set objWorkSheet = Nothing
		    Set objWorkbook = Nothing
		    Set objExcel = Nothing											
			
			fn_splm_excel_update_testduration = True
		Else
			fn_splm_excel_update_testduration = False
		End If
		
		Set objFSO = Nothing

End Function

''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 32                                                                              
'' Function Name     	: Fn_SetEnvValue(sType, sEnvName, sEnvValue)
'' Function Description : Function used to set the env values for the system
'' Function Usage    	: Result = Fn_SetEnvValue(sType, sEnvName, sEnvValue)
''							sType		      - User / System
''							sEnvName	- Name of the env variable (e.g. FMS_HOME, AutomationDir)
''							sEnvValue     - Value to be set for sEnvName  
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Sunny Ruparel    		|  25-Mar-2011	| 	1.0	 |													 |	Sandeep N 	| 25-Mar-2011  
''----------------------------------------------------------------------------------------------------------------------
Function Fn_SetEnvValue(sType, sEnvName, sEnvValue)

	Dim objShell
	Dim objUsrEnv
	
	Set objShell = CreateObject("WScript.Shell")
	Set objUsrEnv = objShell.Environment(sType)
		objUsrEnv(sEnvName) = sEnvValue
    Set objUsrEnv = Nothing
	Set objShell = Nothing
	
	Fn_SetEnvValue = True

End Function
''--------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 33                                                                              
'' Function Name     	: Fn_LogUtil_GetXMLPath(xmlName)
'' Function Description : Function used to get XML paths
'' Function Usage    	: Result = Fn_LogUtil_GetXMLPath(xmlName)
''							xmlName		      - XML file name without extension
'  Return Value : XML path or False
'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Sandeep N   		|  15-Apr-2011	| 	1.0	 |									|	Sunny R		  | 15-Apr-2011  
''----------------------------------------------------------------------------------------------------------------------
'' 	Koustubh W   		|  11-May-2011	| 	1.0	 |Added code to handle PSE Menu XML
''----------------------------------------------------------------------------------------------------------------------
'' 	Shantan   		|  22-Mar-2016	| 	1.0	 |Added case "SM_popupMenu" for structure manager | Shweta R | 22-Mar-2016
''---------------------------------------------------------------------------------
Function Fn_LogUtil_GetXMLPath(xmlName)
   Fn_LogUtil_GetXMLPath=False
	Environment.Value("sPath") = Fn_GetEnvValue("User", "AutomationDir")
	Select Case xmlName
		Case "Web_User"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\WebConfig\Web_Users.xml"
		Case "Web_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\WEB_Menu.xml"
		Case "WebMyTc_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\WEBMyTc_Menu.xml"
		Case "Web_ErrorMsg"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\WebConfig\Web_ErrorMsg.xml"
		Case "WebMyTc_ErrorMsg"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\WebConfig\WebMyTc_ErrorMsg.xml"
		Case "WebChangeMgr_Users"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\WebConfig\WebChangeMgr_Users.xml"
        Case "WebChangeMgr_ErrorMsg"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\WebConfig\WebChangeMgr_ErrorMsg.xml"
        Case "WEBChange_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\WEBChange_Menu.xml"
		Case "RAC_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\RAC_Menu.xml"
		Case "WEB_PSE_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\WEB_PSE_Menu.xml"
		Case "WEB_Classification_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\WEB_Classification_Menu.xml"
		Case "LifecycleViewer_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\LifecycleViewer_Menu.xml"
		Case "Viz_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\Viz_Menu.xml"
		Case "WEB_SE_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\WEB_SE_Menu.xml"
		Case "WEB_DC_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\WEB_DC_Menu.xml"
		Case "RAC_Toolbar"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\RAC_Toolbar.xml"
		Case "RAC_PopupMenu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\PopupMenuXML\RAC_PopupMenu.xml"
        Case "RAC_CPD_Messages"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\WebConfig\RAC_CPD_Messages.xml"
		Case "MyTc_ErrorMsg"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\MyTc_ErrorMsg.xml"	
		Case "Workflow_ErrorMessage"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\Workflow_ErrorMessage.xml"
		Case "Workflow_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\Workflow_Menu.xml"
		Case "MyTc_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\MyTc_Menu.xml"	
		Case "StructureManager_ErrorMessage"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\StructureManager_ErrorMessage.xml"
		Case "PSM_Toolbar"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\PSM_Toolbar.xml"
		Case "CPD_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\CPD_Menu.xml"
		Case "PSE_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\PSE_Menu.xml"
        Case "Classification_Toolbar"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\Classification_Toolbar.xml"	
		Case "CollaborativeProductDevelopmentError"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\CollaborativeProductDevelopmentError.xml"	
		Case "ManufacturingProcessPlanner_ErrorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\ManufacturingProcessPlanner_ErrorMessage.xml"	
        Case "MRO_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\MRO_Menu.xml"
		Case "4GD_Toolbar"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\4GD_Toolbar.xml"
		Case "Classification_ErorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\Classification_ErorMessage.xml"
        Case "MRO_PopupMenu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\PopupMenuXML\MRO_PopupMenu.xml"
        Case "WebADS_ErrorMsg"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")&"\TestData\AutomationXML\WebConfig\WebADS_ErrorMsg.xml"
        Case "ADS_ErrorMsg"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")&"\TestData\AutomationXML\ErrorMessageXML\ADS_ErrorMsg.xml"
		Case "RDV_Toolbar"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\RDV_Toolbar.xml"
        Case "RDV_ErrorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")&"\TestData\AutomationXML\ErrorMessageXML\RDV_ErrorMessage.xml"
		Case "WebDC_ErrorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")&"\TestData\AutomationXML\WebConfig\WebDC_ErrorMessage.xml"
		Case "VIZ_ErrorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")&"\TestData\AutomationXML\ErrorMessageXML\VIZ_ErrorMessage.xml"
		Case "VISIO_ShapeData"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")&"\TestData\AutomationXML\Visio\ShapeData.xml"
		Case "RM_ErrorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")&"\TestData\AutomationXML\ErrorMessageXML\RM_ErrorMessage.xml"
		Case "ServiceScheduler"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\ServiceScheduler.xml"
		Case "MRO_Toolbar"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\MRO_Toolbar.xml"
		Case "MRO_AdvancedSearch"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\AdvancedSearchXML\MRO_AdvancedSearch.xml"
		Case "MRO_ErrorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\MRO_ErrorMessage.xml"
		Case "SearchErrorMessage"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\SearchErrorMessage.xml"
		Case "WEB_MyTc_ErrorMsg"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\WEB_MyTc_ErrorMsg.xml"
		Case "Workflow_Users"	
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\Workflow_Users.xml"
		Case "Mechatronics_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\Mechatronics\Mechatronics_Menu.xml"	
		Case "Mechatronics_ErrorMsg"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\Mechatronics\Mechatronics_ErrorMsg.xml"
		Case "ChangeManager"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\ChangeManager.xml"
		Case "ChangeMgr_Users"	
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath") + "\TestData\ChangeMgr_Users.xml"  
		Case "SP_PopupMenu"	
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath") + "\TestData\AutomationXML\PopupMenuXML\SP_PopupMenu.xml" 			
		Case "RM_PopupMenu"	
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath") + "\TestData\AutomationXML\PopupMenuXML\RM_PopupMenu.xml"
		Case "ContentMgmt_PopupMenu"	
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath") + "\TestData\AutomationXML\PopupMenuXML\ContentMgmt_PopupMenu.xml"
		Case "RM_Menu"	
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\RM_Menu.xml"
		Case "Dataset_Type"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\TcDataTypes\Dataset_Types.xml"
		Case "RM_Toolbar"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\RM_Toolbar.xml"
		Case "AdminITAR_ErrorMsg"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\AdminITAR_ErrorMsg.xml"
		Case "PPM_ErrorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\PPM_ErrorMessage.xml"	
		Case "Project_ErrorMessage"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\Project_ErrorMessage.xml"		
		Case "PWC_PopupMenu"	
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath") + "\TestData\AutomationXML\PopupMenuXML\PWC_PopupMenu.xml"
		Case "ContentManagement_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml"
		Case "VIZ_Toolbar"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\VIZ_Toolbar.xml"	
		Case "Configurator_Menu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\Configurator_Menu.xml"
		Case "Configurator_Toolbar"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\Configurator_Toolbar.xml"
		Case "CPD_PopupMenu"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\PopupMenuXML\CPD_PopupMenu.xml"
		Case "ResourceMgr_Toolbar"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\ResourceMgr_Toolbar.xml"
		Case "SM_popupMenu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\PopupMenuXML\SM_popupMenu.xml"
		Case "Subscribe_ErrMsg"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\SubscribeAudit_ErrMsg.xml"
		Case "OfficeClient_ErrorMessage"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\OfficeClient_ErrorMessage.xml"
		Case "OfficeClient_PopupMenu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\PopupMenuXML\OfficeClient_PopupMenu.xml"
		Case "OfficeClientDropdownMenu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarDropdownMenu\OfficeClientDropdownMenu.xml"
		Case "BriefcaseBrowser_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\BriefcaseBrowser_Menu.xml"
		Case "BriefcaseBrowser_Toolbar"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarXML\BriefcaseBrowser_Toolbar.xml"
		Case "BriefcaseBrowser_Envvar"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\BBEnvVar.xml"
		Case "NX_Menu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\MenuXML\NX_Menu.xml"
		Case "BB_PopupMenu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\PopupMenuXML\BB_PopupMenu.xml"
		Case "BB_ErrorMessage"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\BriefcaseBrowser_ErrMsg.xml"
		Case "NX_PopupMenu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\PopupMenuXML\NX_PopupMenu.xml"
		Case "Key_Numbers"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\KeyXML\Key_Numbers.xml"
		Case "RAC_NewItemValues_APL"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ApplicationInformationXML\RAC_NewItemValues_APL.xml"
		Case "RAC_TabValues_APL"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ApplicationInformationXML\RAC_TabValues_APL.xml"
		Case "RAC_NewFolderValues_APL"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ApplicationInformationXML\RAC_NewFolderValues_APL.xml"
		Case "RAC_NewPartValues_APL"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ApplicationInformationXML\RAC_NewPartValues_APL.xml"
		Case "RAC_NewFormValues_APL"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ApplicationInformationXML\RAC_NewFormValues_APL.xml"
		Case "RAC_ProjectValues_APL"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ApplicationInformationXML\RAC_ProjectValues_APL.xml"
		Case "RAC_GroupValues_APL"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ApplicationInformationXML\RAC_GroupValues_APL.xml"	
		Case "RAC_LoadQueryValues_APL"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ApplicationInformationXML\RAC_LoadQueryValues_APL.xml"	
		Case "PIE_Messages"
			Fn_LogUtil_GetXMLPath=Environment.Value("sPath")+ "\TestData\AutomationXML\WebConfig\PIE_Messages.xml"
		Case "ConfiguratorErrorMessage"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\ConfiguratorErrorMessage.xml"
		Case "TcViz_PopupMenu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\PopupMenuXML\TcViz_PopupMenu.xml"
		Case "RM_ToolbarDropdownMenu"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ToolbarDropdownMenu\RM_ToolbarDropdownMenu.xml"
		Case "CPD_DisplayName"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectRealNames\CPD_DisplayName.xml"	
		Case "batchsetup"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\BatchRun\batchsetup.xml"
		Case "Administrator_ErrMessage"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\Administrator_ErrMessage.xml"	
		Case "SiteInfo"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\Sites.xml"	
		Case "ClassicMultiSite_ErrorMessage"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\ClassicMultiSite_ErrorMessage.xml"	
		Case "UnabletocreateBusinessObject"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ErrorMessageXML\MyTc_ErrorMsg.xml"
		Case "BriefcaseBrowser"
			Fn_LogUtil_GetXMLPath = Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectXML\BriefcaseBrowser.xml"			
End Select
End Function

''---------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 34                                                                            
'' Function Name     	: fn_SISW_util_folder_operation(sAction,sFolderPath)
'' Function Description : Function used to perform various operation on folder
''						: 1. create folder
''						: 2. delete folder
'' Function Usage    	: Result = fn_SISW_util_folder_operation("createfolder","c:\foldername")

'' Function History
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name		|	  Date		|Rev. No.|		    Changes Done			|	Reviewer	|	Reviewed Date
''----------------------------------------------------------------------------------------------------------------------
'' 	Shweta Rathod		|  26-Aug-2016	| 	1.0	 |			Added function			| Kaustubh watwe|  26-Aug-2016
''----------------------------------------------------------------------------------------------------------------------
Function fn_SISW_util_folder_operation(sAction,sFolderPath)
	fn_SISW_util_folder_operation = false
	Dim objFSO,objFolder
	Dim sPath
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	sPath = sFolderPath
	Select Case lcase(sAction)
		Case "exist"
				fn_SISW_util_folder_operation = objFSO.FolderExists(sPath)
		Case "createfolder"
			If objFSO.FolderExists(sPath) = false Then
				Set objFolder = objFSO.CreateFolder(sPath)
			End If
			fn_SISW_util_folder_operation = true

		Case "deletefolder"
			If objFSO.FolderExists(sPath) Then
				objFSO.DeleteFolder(sPath)
			End if
			fn_SISW_util_folder_operation = true

	End Select
	Set objFolder = Nothing
	Set objFSO = Nothing
End Function

''---------------------------------------------------------------------------------------------------------------------
'' Function Number   	: 35                                                                            
'' Function Name     	: Fn_ExcelGetQARTStatusDetail(sExcelPath, sActionType, iSheetNumber)
Public Function Fn_ExcelGetQARTStatusDetail(sExcelPath, sActionType, iSheetNumber)

	Const xlCellTypeLastCell = 11
	Dim objFSO, objExcel, objWorkbook, objWorkSheet, objRange
	Dim usedRange, iColId ,TotalCnt, NotUploadedCnt, ResultsSavedCnt
    
	iColId = Fn_ExcelSearch(sExcelPath, "QART Upload Status", iSheetNumber)
	iColId = Fn_strUtil_SubField(iColId, ":", 1)    
	
	iColId = Chr(64 + iColId)
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(sExcelPath) Then
	
		Set objExcel = CreateObject("Excel.Application")
			objExcel.Visible = False
			objExcel.AlertBeforeOverwriting = False
			objExcel.DisplayAlerts = False
	
			Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)	

			If iSheetNumber = "" Then
				 iSheetNumber = 1
			End If 			
			
			Set objWorkSheet = objExcel.ActiveWorkbook.Worksheets(iSheetNumber)
			Set objRange = objWorkSheet.UsedRange
			objRange.SpecialCells(xlCellTypeLastCell).Activate	
			
			usedRange = iColId & "2:" & iColId & objExcel.ActiveCell.Row
			
			if  LCase(sActionType) = LCase("Results Saved") then	
				objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "Results Saved" & Chr(34) & ")"
				Fn_ExcelGetQARTStatusDetail = objExcel.Cells(10000, 1).value
				
			elseif LCase(sActionType) = LCase("Not Uploaded") then
				TotalCnt = objExcel.ActiveCell.Row - 1
				objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "Results Saved" & Chr(34) & ")"
				ResultsSavedCnt = objExcel.Cells(10000, 1).value
				objExcel.Cells(10001, 1).Formula = "=COUNTIF(" & usedRange & "," & Chr(34) & "Not Uploaded" & Chr(34) & ")"
				NotUploadedCnt = objExcel.Cells(10001, 1).value
				if  ResultsSavedCnt + NotUploadedCnt <> TotalCnt then
					Fn_ExcelGetQARTStatusDetail = 	TotalCnt - ResultsSavedCnt
				else
					Fn_ExcelGetQARTStatusDetail = 	NotUploadedCnt
				end if								
			end if				
			

		objExcel.Quit			
			
		Set objRange = Nothing
		Set objWorkSheet = Nothing
		Set objWorkbook = Nothing
		Set objExcel = Nothing
		Set objFSO = Nothing
	End If
End Function
