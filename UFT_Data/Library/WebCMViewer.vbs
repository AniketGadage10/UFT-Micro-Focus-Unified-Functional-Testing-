Option Explicit
																								'Function List
'************************************************************************************************************************************************************************************************************
'001. Fn_CMV_TableOperations()
'************************************************************************************************************************************************************************************************************
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_CMV_TableOperations
'@@
'@@    Description				 :	Function Used to perform Operation On BOM Table
'@@
'@@    Parameters			   :	1.sAction : Action Name
'@@												  2.sNodeName : Node Path Or Item Name
'@@												  3.sColumn : Column Name
'@@												  4.sCellValue : Expected Value
'@@
'@@    Return Value		   	   : 	True Or False Or Column Names Or Image Name
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client And PSEperspective should be open
'@@
'@@    Examples					:	
'@@				cases		NodeSelect / Select 		- Call Fn_CMV_TableOperations("NodeSelect","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@												 		- Call Fn_CMV_TableOperations("Select","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View) @2:000057/A;1-asm","","")
'@@							NodeDeSelect / Deselect 	- Call Fn_CMV_TableOperations("NodeDeSelect","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							NodeVerify / Exist /Exists 	- Call Fn_CMV_TableOperations("NodeVerify","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View):000056/A;1-SubItem2","","")
'@@							CellVerify					- Call Fn_CMV_TableOperations("CellVerify","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View):000056/A;1-SubItem2","Item Type","Item")
'@@							Collapse					- Call Fn_CMV_TableOperations("Collapse","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							MultiSelect					- Call Fn_CMV_TableOperations("MultiSelect","000054/A;1-TopItem (View)~000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							Expand						- Call Fn_CMV_TableOperations("Expand","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							ColumnExist					- Call Fn_CMV_TableOperations("ColumnExist","","Name~BOM Line~Find No.","")
'@@							ColumnClick					- Call Fn_CMV_TableOperations("ColumnClick","","Name~BOM Line~Find No.","")
'@@							CellEdit					- Call Fn_CMV_TableOperations("CellEdit","000015/A;1-top (View):000016/A;1-sub","Find No.","20") 
'@@							ClearSelection				- Call Fn_CMV_TableOperations("ClearSelection","","","")
'@@
'@@	   History:				Developer Name										Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe										17-nov-2011						1.0							Created
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CMV_TableOperations(sAction, sNodeName, sColumn, sCellValue)
	GBL_FAILED_FUNCTION_NAME="Fn_CMV_TableOperations"
	' Declaration of an Variable
	Dim objDialog, objImg, objWebChk, objLink, objButton
	Dim aElements, aSubElement, iCounter, bFlag, jCounter, iRowCnt, iColPos, iOuterCnt, sText, iCounter2, iRowCnt2
	Dim objTRs, objTDs, objElements
	' Initialization of an Variable
	Set objDialog = Browser("Teamcenter Web - Change").Page("Teamcenter Web - Change").WebTable("CMViewerTable")
	bFlag = False
	Fn_CMV_TableOperations = False
	
	'Operations Case
	Select Case sAction
		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "ClearSelection"
			iRowCnt = objDialog.RowCount
			for iCounter = 1 to iRowCnt -1
				Set objWebChk = objDialog.ChildItem(iCounter, 1, "WebCheckBox", 0)
				If TypeName(objWebChk) <> "Nothing" Then
					If objWebChk.GetROProperty("checked") = "1" Then
						objWebChk.Click 1, 1, micLeftBtn
					End If
				End If
			next
			Fn_CMV_TableOperations = True
		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "NodeSelect", "Select"
			iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
			If iRowCnt <> -1 Then
				Call Fn_CMV_TableOperations("ClearSelection", "", "", "")
				Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
				If TypeName(objWebChk) <> "Nothing" Then
					If objWebChk.GetROProperty("checked") = "0" Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Node ["+CStr(sNodeName)+"] found.")
						objWebChk.Click 1, 1, micLeftBtn
						bFlag = True
					elseIf objWebChk.GetROProperty("checked") = "1" Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Node ["+CStr(sNodeName)+"] found.")
						objWebChk.Click 1, 1, micLeftBtn
						objWebChk.Click 1, 1, micLeftBtn
						bFlag = True
					End If
				End If
				Set objWebChk = Nothing
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
			End If
			' Write the Log of Success or Failure
			If bFlag = True Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Node ["+CStr(sNodeName)+"] Selected Successfully ")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to Select the Nod ["+CStr(sNodeName)+"] . ")
			End If

		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "NodeVerify", "Exist", "Exists"
			' Write the Log of Success or Failure
			iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
			If iRowCnt <> -1 Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Node  ["+CStr(sNodeName)+"] Verified Successfully. ")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to Verify the Node  ["+CStr(sNodeName)+"] . ")
			End If

		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "CellVerify"
			' For Getting the Column Header Position
			If sColumn = "" Then
				iColPos = Fn_WebUI_TableColumnIndex(objDialog,"Name")
			else
				iColPos = Fn_WebUI_TableColumnIndex(objDialog, sColumn)
			End If
			iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
			If iColPos <> -1 and iRowCnt <> -1 Then
				Set objLink = objDialog.ChildItem(iRowCnt, iColPos, "WebEdit", 0)
				If TypeName(objLink) <> "Nothing" Then
					If trim(objLink.GetROProperty("value")) = Trim(sCellValue) then 
						bFlag = True
					end if
				ElseIf Trim(objDialog.GetCellData(iRowCnt, iColPos)) = Trim(sCellValue) Then
						bFlag = True
				End If
			End If
			' Write the Log of Success or Failure
			If bFlag = True Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations :  Value [ " & sCellValue & "]  is successfully verified for column [ " & sColumn & " ] at Node ["+CStr(sNodeName)+"]. ")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations :  Failed to verify value [ " & sCellValue & "]  for column [ " & sColumn & " ] at Node ["+CStr(sNodeName)+"]. ")
			End If

		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -				
		Case "Expand"
			iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
			iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
			If iRowCnt <> -1 Then
				Set objImg = objDialog.ChildItem(iRowCnt, iColPos, "Image", 0)
				If TypeName(objImg) <> "Nothing" Then
					If objImg.GetROProperty("file name") = "plus.png" Then
						objImg.Click 1,1, micLeftBtn
						bFlag = True
					elseIf objImg.GetROProperty("file name") = "minus.png" Then
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations: node  ["+CStr(sNodeName)+"] was already expanded.")
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations: can not expand node  ["+CStr(sNodeName)+"].")
					End If
				End If						
				Set objImg = Nothing
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations: node  ["+CStr(sNodeName)+"] does not exist in BOM table.")
			End If
			' Write the Log of Success or Failure
			If bFlag = True Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Node  ["+CStr(sNodeName)+"] expanded successfully. ")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to expand node  ["+CStr(sNodeName)+"]")
			End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "Collapse"
			iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
			iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
			If iRowCnt <> -1 and iColPos <> -1 Then
				Set objImg = objDialog.ChildItem(iRowCnt, iColPos, "Image", 0)
				If TypeName(objImg) <> "Nothing" Then
					If objImg.GetROProperty("file name") = "minus.png" Then
						objImg.Click 1,1, micLeftBtn
						bFlag = True
					End If
				End If
			end if
			' Write the Log of Success or Failure
			If bFlag = True Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Node  ["+CStr(sNodeName)+"] collapsed Successfully. ")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to collapse the Node  ["+CStr(sNodeName)+"] . ")
			End If
		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -		
		Case "MultiSelect"
			aElements = Split(sNodeName, "~", -1, 1)
			For iOuterCnt = 0 To UBound(aElements)
				iRowCnt = Fn_WebUI_TableRowIndex(objDialog, aElements(iOuterCnt), "")
				If iRowCnt <> -1 and iColPos <> -1 Then
					Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
					If TypeName(objWebChk) <> "Nothing" Then
						If objWebChk.GetROProperty("checked") = "0" Then
							objWebChk.Click 
							bFlag = True
						End If
					End If
					Set objWebChk = Nothing
				else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to find node ["+CStr(aElements(iOuterCnt))+"] . ")	
					Exit for
				End If
			Next
			If bFlag = True Then
				' For Success Log
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Multiple Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] selected Successfully. ")
			Else
				' For Failure Log
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to Select the Multiple Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] . ")	
			End If
		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "NodeDeSelect", "Deselect"
			aElements = Split(sNodeName, "~", -1, 1)
			jCounter = 0
			For iOuterCnt = 0 To UBound(aElements)
				iRowCnt = Fn_WebUI_TableRowIndex(objDialog, aElements(iOuterCnt), "")
				If iRowCnt <> -1 and iColPos <> -1 Then
					Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
					If TypeName(objWebChk) <> "Nothing" Then
						If objWebChk.GetROProperty("checked") = "1" Then
							objWebChk.Click 
							bFlag = True
						End If
					End If
					Set objWebChk = Nothing
				else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to find node ["+CStr(aElements(iOuterCnt))+"] . ")	
					Exit for
				End If
			Next
			' Write the Log of Success or Failure
			If bFlag = True Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Deselected  Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] successfully. ")
			Else
				' For Failure Log
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to deselect  the Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] . ")	
			End If
		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "NodeClick"
			' For Getting the Column Header Position
			jCounter = objDialog.ColumnCount(1)
			aSubElement = Split(sNodeName, ":", -1, 1)
			If Trim(sColumn) <> "" Then ' here
				iRowCnt = 1
			Else
				iRowCnt = 2
			End If
			For iCounter = iRowCnt To jCounter
				Set objLink = objDialog.ChildItem(1, iCounter, "Link", 0)
				If TypeName(objLink) <> "Nothing" Then
					If Trim(objLink.GetROProperty("text")) = Trim(sColumn) Then
						iColPos = iCounter
						Exit For
					End If
				Else
					sText = objDialog.GetCellData(1, iCounter)
					If Trim(sText) = Trim(sColumn) Then
						iColPos = iCounter
						Exit For
					End If
					sText = ""
				End If
				Set objLink = Nothing
			Next
			jCounter = 0
			iRowCnt = objDialog.RowCount
			iCounter = 1
			If Trim(sCellValue) = "" and Trim(sColumn) = "" Then
				iColPos = iColPos - 1
			End If
			Do While 1 = 1
				objDialog.RefreshObject
				iRowCnt = objDialog.RowCount
				' For Last Node of an Element
				If jCounter=UBound(aSubElement) Then
					If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
						iRowCnt2 = objDialog.ChildItemCount(iCounter, iColPos, "Link")
						 For iCounter2 = 0 To  iRowCnt2 -1
							Set objLink = objDialog.ChildItem(iCounter, iColPos, "Link", iCounter2)
							If TypeName(objLink) <> "Nothing" Then
								If Trim(objLink.GetROProperty("text")) = Trim(sCellValue) Then
									bFlag = True
									objLink.Click
									Exit For
								End If
							End If
						 Next
						 If Trim(sCellValue) = "" Then
							Set objLink = objDialog.ChildItem(iCounter, iColPos, "Link", 0)
							If TypeName(objLink) <> "Nothing" Then
								bFlag = True
								objLink.Click
							End If
						 End If
						jCounter = jCounter + 1
						Set objLink = Nothing
						Exit Do
					End If
				Else
					' For the Node Hierarchy of an Element
					If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
						jCounter = jCounter + 1
					End If
				End If
				' Exit from the Loop when Total Rows Finished without finding an Node
				If iRowCnt = iCounter  Then
					Exit Do
				Else
					' Increment the Counter for next level
					iCounter = iCounter + 1
				End If
			Loop
			' Write the Log of Success or Failure
			If bFlag = True Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Node  ["+CStr(sNodeName)+"]  Clicked Successfully. ")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click the Node  ["+CStr(sNodeName)+"] . ")
			End If

		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "ColumnExist"
			If sColumn <> "" Then
				aElements = Split(sColumn, "~", -1, 1)
				For iCounter = 0 To UBound(aElements)
					iColPos = Fn_WebUI_TableColumnIndex(objDialog,aElements(iCounter))
					If iColPos <> -1  Then
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Column  [ " & aElements(iCounter) & " ]  exists in BOM table. ")
					else
						bFlag = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to check existence of column  [ " & aElements(iCounter) & " ]. ")
						Exit for
					End If
				Next
			End If
			' Write the Log of Success or Failure
			If bFlag = True Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : All Columns  ["+CStr(replace(sColumn,"~", ", "))+"]  exists in BOM table. ")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to check existence of column(s) ["+CStr(replace(sColumn,"~", ", "))+"]. ")
			End If
	' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "ColumnClick"
			' For Getting the Column Header Position
			If sColumn <> "" Then
				iColPos = Fn_WebUI_TableColumnIndex(objDialog, sColumn)' For Row Number
				iOuterCnt = 0
				If iColPos <> -1 Then
					iRowCnt = objDialog.ChildItemCount(1, iColPos, "Link")
					If iRowCnt > 0 Then
						Set objLink = objDialog.ChildItem(1, iColPos, "Link", 0)
						If TypeName(objLink) <> "Nothing" Then
							If Trim(objLink.GetROProperty("text")) = Trim(sColumn) Then
								objLink.Click 
								bFlag = True
							End If
						End If
						Set objLink = Nothing
					End If
				End If
			End If
			If bFlag = True Then
				' Write the Log of Success or Failure
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Column Heading  ["+CStr(sColumn)+"]  Clicked Successfully. ")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Column Heading  ["+CStr(sColumn)+"]  Not Found to Click. ")
			End If	
		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
		Case "CellEdit"
			objDialog.RefreshObject
			iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
			iColPos = Fn_WebUI_TableColumnIndex(objDialog, sColumn)' For Row Number
			If iRowCnt <> -1 Then
					Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
					If TypeName(objWebChk) <> "Nothing" Then
						If objWebChk.GetROProperty("checked") = "0" Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Node ["+CStr(sNodeName)+"] found.")
							objWebChk.Click 1, 1, micLeftBtn
						End If

						Set objElements = Description.Create
						objElements("html tag").value = "TR"
						Set objTRs =  objDialog.ChildObjects(objElements)
						objElements("html tag").value = "TD"
						Set objTDs = objTRs(iRowCnt - 1 ).ChildObjects(objElements)
						objTDs(iColPos -1).click

						Set objLink = objDialog.ChildItem(iRowCnt, iColPos, "WebEdit", 0)
						If TypeName(objLink) <> "Nothing" Then
							objLink.set sCellValue
							bFlag = True
						End If
					End If
'						End If
					Set objWebChk = Nothing
			else
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
			End If
			' Write the Log of Success or Failure
			If bFlag = True Then
				Fn_CMV_TableOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_CMV_TableOperations : Modified Cell [ "+CStr(sNodeName)+" ][ " & sColumn & " ] to value [ " & sCellValue & " ] Successfully.")
			Else
				Fn_CMV_TableOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_CMV_TableOperations : Failed to modify Cell [ "+CStr(sNodeName)+" ][ " & sColumn & " ] to value [ " & sCellValue & " ].")
			End If
	End Select
	Set objDialog = Nothing
	Set objImg = Nothing
	Set objWebChk = Nothing
	Set objLink = Nothing
End Function
