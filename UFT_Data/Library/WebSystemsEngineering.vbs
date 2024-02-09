Option Explicit
																								'Function List
'************************************************************************************************************************************************************************************************************
'001. Fn_WebSE_BOMTableOperations()
'002. Fn_WebSE_CreateBaseline()
'003. Fn_WebSE_TraceLinkVerifications()
'004. Fn_WebSE_BOMTable_ViewEditOperations()
'************************************************************************************************************************************************************************************************************
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebSE_BOMTableOperations
'@@
'@@    Description				 :	Function Used to perform Operation On Syst6em Engineering BOM Table
'@@
'@@    Parameters			   :	1.sAction : Action Name
'@@												  2.sNodeName : Node Path Or Item Name
'@@												  3.sColumn : Column Name
'@@												  4.sCellValue : Expected Value
'@@
'@@    Return Value		   	   : 	True Or False Or Column Names Or Image Name
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client And SE perspective should be open
'@@
'@@    Examples					:	
'@@											Msgbox Fn_WebSE_BOMTableOperations("GetImage","000098/A;1-Requierment_Data_File (View)","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("NodeSelect","000098/A;1-Requierment_Data_File (View):000207/A;1-Para2","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("Select","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View) @2:000057/A;1-asm","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("NodeDeSelect","000098/A;1-Requierment_Data_File (View):000207/A;1-Para2","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("NodeVerify","000098/A;1-Requierment_Data_File (View):000207/A;1-Para2","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("CellVerify","000210/A;1-Spec1 (View):000211/A;1-Para1","Item Type","Paragraph")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("Collapse","000210/A;1-Spec1 (View)","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("MultiSelect","000210/A;1-Spec1 (View)~000210/A;1-Spec1 (View):000212/A;1-Para2","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("Expand","000210/A;1-Spec1 (View)","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("ExpandBelow","000054/A;1-TopItem (View)","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("GetImage","","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("FirstElement","","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("ColumnExist","","Name~BOM Line~Find No.","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("ColumnClick","","Name~BOM Line~Find No.","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("CellEdit","000015/A;1-top (View):000016/A;1-sub","Find No.","20") 
'@@											Msgbox  Fn_WebSE_BOMTableOperations("ClearSelection","","","")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("NodeClick","000210/A;1-Spec1 (View):000212/A;1-Para2","","View")
'@@											Msgbox  Fn_WebSE_BOMTableOperations("GetNextElement", "000078/A;1-Alenia_123 (View):REQ-000130/A;1-Program Context (View)", "", "")
'@@
'@@	   History:				Developer Name										Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Sandeep Navghane									25-Sep-2011						1.0																		Sunny Ruparel
'@@							Sandeep Navghane									22-Dec-2011						1.1				Added Case "GetNextElement"														Swati K
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_WebSE_BOMTableOperations(sAction, sNodeName, sColumn, sCellValue)
		GBL_FAILED_FUNCTION_NAME="Fn_WebSE_BOMTableOperations"
		' Declaration of an Variable
		Dim objDialog, objImg, objWebChk, objLink, objButton
        Dim aElements, aSubElement, iCounter, bFlag, jCounter, iRowCnt, iColPos, iOuterCnt, sText, iCounter2, iRowCnt2
		Dim objTRs, objTDs, objElements
		Dim iRowPos
		' Initialization of an Variable
		bFlag = False
		For iCounter=0 to 3
			If Browser("TeamcenterWeb").Page("SystemsEngineering").WebTable("SEBOMTable").Exist(4) Then
				Set objDialog = Browser("TeamcenterWeb").Page("SystemsEngineering").WebTable("SEBOMTable")
				wait 2
				bFlag = True
				Exit for
			End If
		Next
		If bFlag = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:  SE BOM Table not exist")
			Exit function
		End If
		bFlag = False
		Fn_WebSE_BOMTableOperations = False
		
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
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
        		Case "NodeSelect", "Select"
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
							Call Fn_WebSE_BOMTableOperations("ClearSelection", "", "", "")
							Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
							If TypeName(objWebChk) <> "Nothing" Then
								If objWebChk.GetROProperty("checked") = "0" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = True
								elseIf objWebChk.GetROProperty("checked") = "1" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = True
								End If
							End If
							Set objWebChk = Nothing
					else
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] Selected Successfully ")
					Else
						Fn_WebSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to Select the Nod ["+CStr(sNodeName)+"] . ")
					End If

			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "NodeVerify", "Exist", "Exists"
					' Write the Log of Success or Failure
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
						Fn_WebSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node  ["+CStr(sNodeName)+"] Verified Successfully. ")
					Else
						Fn_WebSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to Verify the Node  ["+CStr(sNodeName)+"] . ")
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
						Fn_WebSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations :  Value [ " & sCellValue & "]  is successfully verified for column [ " & sColumn & " ] at Node ["+CStr(sNodeName)+"]. ")
					Else
						Fn_WebSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations :  Failed to verify value [ " & sCellValue & "]  for column [ " & sColumn & " ] at Node ["+CStr(sNodeName)+"]. ")
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
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations: node  ["+CStr(sNodeName)+"] was already expanded.")
									else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations: can not expand node  ["+CStr(sNodeName)+"].")
									End If
							End If						
							Set objImg = Nothing
					else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations: node  ["+CStr(sNodeName)+"] does not exist in BOM table.")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node  ["+CStr(sNodeName)+"] expanded successfully. ")
					Else
						Fn_WebSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to expand node  ["+CStr(sNodeName)+"]")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -				
			Case "ExpandBelow"
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
							Call Fn_WebSE_BOMTableOperations("ClearSelection", "", "", "")
							Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
							If TypeName(objWebChk) <> "Nothing" Then
								    If objWebChk.GetROProperty("checked") = "0" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = Fn_Web_MenuOperation("Select",Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewExpandBelow"))
								elseIf objWebChk.GetROProperty("checked") = "1" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = Fn_Web_MenuOperation("Select",Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewExpandBelow"))
								End If
							End If
							Set objWebChk = Nothing
					else
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node  ["+CStr(sNodeName)+"] expanded successfully. ")
					Else
						Fn_WebSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to expand node  ["+CStr(sNodeName)+"]")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "GetImage"
'				objDialog.RefreshObject
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
					If iRowCnt <> -1 and iColPos <> -1 Then
							Set objImg = objDialog.ChildItem(iRowCnt, iColPos, "Image", 0)
							If TypeName(objImg) <> "Nothing" Then
									strImageName=Split(objImg.GetROProperty("file name"),".")
									Fn_WebSE_BOMTableOperations=strImageName(0)
									bFlag = True
							Else
									Fn_WebSE_BOMTableOperations=False
							End If						
							Set objImg = Nothing
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_WebSE_BOMTableOperations = strImageName(0)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Image [ " & Fn_WebSE_BOMTableOperations & " ]  associated with node  ["+CStr(sNodeName)+"].")
					Else
									Fn_WebSE_BOMTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations :  Failed to retrieve image name of node  ["+CStr(sNodeName)+"].")
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
									Fn_WebSE_BOMTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node  ["+CStr(sNodeName)+"] collapsed Successfully. ")
					Else
									Fn_WebSE_BOMTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to collapse the Node  ["+CStr(sNodeName)+"] . ")
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
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to find node ["+CStr(aElements(iOuterCnt))+"] . ")	
										Exit for
								End If
						Next
						If bFlag = True Then
								' For Success Log
								Fn_WebSE_BOMTableOperations = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Multiple Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] selected Successfully. ")
						Else
								' For Failure Log
								Fn_WebSE_BOMTableOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to Select the Multiple Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] . ")	
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
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to find node ["+CStr(aElements(iOuterCnt))+"] . ")	
										Exit for
								End If
						Next
					' Write the Log of Success or Failure
					If bFlag = True Then
							Fn_WebSE_BOMTableOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Deselected  Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] successfully. ")
					Else
							' For Failure Log
							Fn_WebSE_BOMTableOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to deselect  the Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] . ")	
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
									Fn_WebSE_BOMTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Node  ["+CStr(sNodeName)+"]  Clicked Successfully. ")
					Else
									Fn_WebSE_BOMTableOperations = False
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
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Column  [ " & aElements(iCounter) & " ]  exists in BOM table. ")
							else
								bFlag = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to check existence of column  [ " & aElements(iCounter) & " ]. ")
								Exit for
							End If
						Next
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_WebSE_BOMTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : All Columns  ["+CStr(replace(sColumn,"~", ", "))+"]  exists in BOM table. ")
					Else
									Fn_WebSE_BOMTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to check existence of column(s) ["+CStr(replace(sColumn,"~", ", "))+"]. ")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "FirstElement"
						' For Getting the Column Header Position
						iRowCnt = objDialog.GetROProperty("rows")
						iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
						If iRowCnt > 1 Then
								iRowCnt2 = objDialog.GetCellData(2, iColPos)
						End If
						If iRowCnt2 <> "" Then
								Fn_WebSE_BOMTableOperations = iRowCnt2
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : First Element ["+CStr(iRowCnt2)+"] Present in the BOM Table ")
						Else
								Fn_WebSE_BOMTableOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : No First Element Found in BOM Table ")
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
							Fn_WebSE_BOMTableOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Column Heading  ["+CStr(sColumn)+"]  Clicked Successfully. ")
						Else
							Fn_WebSE_BOMTableOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Column Heading  ["+CStr(sColumn)+"]  Not Found to Click. ")
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
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
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
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : Modified Cell [ "+CStr(sNodeName)+" ][ " & sColumn & " ] to value [ " & sCellValue & " ] Successfully.")
					Else
						Fn_WebSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : Failed to modify Cell [ "+CStr(sNodeName)+" ][ " & sColumn & " ] to value [ " & sCellValue & " ].")
					End If
'			 - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "GetNextElement"
						
						iRowCnt = objDialog.GetROProperty("rows")
						iRowPos = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "Name")
						If iRowCnt > 1 Then
								iRowCnt2 = objDialog.GetCellData(iRowPos+1,2)
						End If
						If iRowCnt2 <> "" Then
								Fn_WebSE_BOMTableOperations =Trim(iRowCnt2)
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTableOperations : First Element ["+CStr(iRowCnt2)+"] Present in the BOM Table ")
						Else
								Fn_WebSE_BOMTableOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTableOperations : No First Element Found in BOM Table ")
						End If
		End Select
		Set objDialog = Nothing
		Set objImg = Nothing
		Set objWebChk = Nothing
		Set objLink = Nothing
End Function
''*********************************************************		Function to Apply Revision Rule	***********************************************************************
'Function Name		:			  Fn_WebSE_CreateBaseline

'Description			 :		 	Function to crate Baseline.	

'Parameters			   :	 		     Function To Add Component using Menu (Edit;Add)
'													1.sAction       : Action to perform
'													2.sBOMLine  : BOM Line to select 
'													3.sDescription  : Description text
'													4.sBaselineTemplate  : Baseline Template text
'													5.sBaselineLabel  : Baseline Label text
'													6.sJobDescription  : Job Description
'													7.bOpenOnCreate  : True / False / "" values for Open On Creat checkbox
'													8.bDryRun  : True / False / "" values for Dry Run checkbox
'													9.bPrecise  : True / False / "" values for Precise checkbox
'													10.sField  : for future use
'													11.sValue  : for future use
'Return Value		   : 			     TRUE \ FALSE

'Pre-requisite			:		 	SE window should be displayed.

'Examples				: 			Call  Fn_WebSE_CreateBaseline("CreateBaseline", "", "Desc", "TC Default Baseline Process", "", "", "", False, False, "", "")
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Koustubh Watwe			19-Oct-2011			1.0				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebSE_CreateBaseline(sAction, sBOMLine, sDescription, sBaselineTemplate, sBaselineLabel, sJobDescription, bOpenOnCreate, bDryRun, bPrecise, sField, sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_WebSE_CreateBaseline"
	Dim objCreateBLTable, objWebChk, objWebEdit, bReturn, sMenu
	Dim objButtonTable
	Fn_WebSE_CreateBaseline = False
	Set objCreateBLTable = Browser("TeamcenterWeb").Page("SystemsEngineering").WebTable("CreateBaseLine")
	Set objButtonTable = Browser("TeamcenterWeb").Page("SystemsEngineering").WebTable("ButtonTable")

	If Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectExist", objCreateBLTable) = False Then
		' select bom line
		If sBOMLine <> "" Then
                  bReturn = Fn_WebSE_BOMTableOperations("Select", sBOMLine,"","")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to select BOM line [ "& sBOMLine & " ].")	
				Exit function
			End If
		End If
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_SE_Menu"), "ToolsBaseLine")
		bReturn = Fn_Web_MenuOperation("Select", sMenu)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to perform menu operation [ "& sMenu & " ].")	
			Exit function
		End If

		'chcek the existence of dialog box
		If Fn_Web_UI_ObjectExist("Fn_Web_UI_ObjectExist", objCreateBLTable) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to open [ Create Baseline ] web dialog.")	
				Exit function
		End If
	End If

	Select Case sAction
'		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "CreateBaseline"
				If sDescription <> "" Then
						' 4
						Set objWebEdit = objCreateBLTable.ChildItem(4, 2, "WebEdit", 0)
						If TypeName(objWebEdit) <> "Nothing" Then
							objWebEdit.Set sDescription
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to get Editbox [ Description ].")	
							Exit function
						End IF
				End If
				If sBaselineTemplate <> ""  Then
						' 5
						Set objWebEdit = objCreateBLTable.ChildItem(5, 2, "WebEdit", 0)
						If TypeName(objWebEdit) <> "Nothing" Then
							objWebEdit.Set sBaselineTemplate
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to get Editbox [ Baseline Template ].")	
							Exit function
						End IF
				End If
				If sBaselineLabel <> ""  Then
						' 7
						Set objWebEdit = objCreateBLTable.ChildItem(7, 2, "WebEdit", 0)
						If TypeName(objWebEdit) <> "Nothing" Then
							objWebEdit.Set sBaselineLabel
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to get Editbox [ Baseline Label ].")	
							Exit function
						End IF
				End If
				If sJobDescription <> ""  Then
						' 8
						Set objWebEdit = objCreateBLTable.ChildItem(5, 2, "WebEdit", 0)
						If TypeName(objWebEdit) <> "Nothing" Then
							objWebEdit.Set sJobDescription
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to get Editbox [ Job Description ].")	
							Exit function
						End IF
				End If
				If bOpenOnCreate <> "" Then
						' 9
						Set objWebChk = objCreateBLTable.ChildItem(9, 2, "WebCheckBox", 0)
						If TypeName(objWebChk) <> "Nothing" Then
							If cBool(bOpenOnCreate) Then
								If objWebChk.GetROProperty("checked") = "0" Then objWebChk.Click 1, 1, micLeftBtn
							Else 								
								If objWebChk.GetROProperty("checked") = "1" Then objWebChk.Click 1, 1, micLeftBtn
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to get Checkbox [ Open On Create ].")	
							Exit function
						End IF
				End If
				If bDryRun <> "" Then
						' 10
						Set objWebChk = objCreateBLTable.ChildItem(10, 2, "WebCheckBox", 0)
						If TypeName(objWebChk) <> "Nothing" Then
							If cBool(bDryRun) Then
								If objWebChk.GetROProperty("checked") = "0" Then objWebChk.Click 1, 1, micLeftBtn
							Else 								
								If objWebChk.GetROProperty("checked") = "1" Then objWebChk.Click 1, 1, micLeftBtn
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to get Checkbox [ Dry Run ].")	
							Exit function
						End IF
				End If
				If bPrecise <> "" Then
						' 11
						Set objWebChk = objCreateBLTable.ChildItem(11, 2, "WebCheckBox", 0)
						If TypeName(objWebChk) <> "Nothing" Then
							If cBool(bPrecise) Then
								If objWebChk.GetROProperty("checked") = "0" Then objWebChk.Click 1, 1, micLeftBtn
							Else 								
								If objWebChk.GetROProperty("checked") = "1" Then objWebChk.Click 1, 1, micLeftBtn
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Failed to get Checkbox [ Precise ].")	
							Exit function
						End IF
				End If
				' clicking on OK button
				If objButtonTable.WebButton("OK").Exist(2) Then
					objButtonTable.WebButton("OK").Click 1,1,micLeftBtn
				Else
					 Browser("TeamcenterWeb").Page("SystemsEngineering").WebButton("OK").Click 1,1,micLeftBtn
				End If
                Fn_WebSE_CreateBaseline = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
                  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_CreateBaseline : Invalid case [ "& sAction & " ].")	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select
	If Fn_WebSE_CreateBaseline = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_CreateBaseline : Executed successfully with case  [ "& sAction & " ].")	
	End If
	Set objCreateBLTable = nothing
	Set objButtonTable = Nothing
	Set objWebChk = nothing
	Set objWebEdit = nothing
End Function


''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_WebSE_TraceLinkVerifications(sAction,sSelectNode,sVerificationNode,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will verify Nodes in Defining and Complying Objects in Tracelinks Tab
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  The Document mapping tree should be present
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sSelectNode : Valid Node name to be selected  { Note : (:) seperated part will contain the Parent node & (,) seperated part contains nodes to select
''''/$$$$										sVerificationNode : Node to be Verified
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          12/01/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			 12/01/2011            1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_WebSE_TraceLinkVerifications("VerifyTracelink","REQ-000001/A;1-RootReq (View):REQ-000003/A;1-2","Defining:REQ-000002/A;1-1","","")
''''/$$$$									  bReturn=Fn_WebSE_TraceLinkVerifications("VerifyTracelink","REQ-000001/A;1-RootReq (View):REQ-000002/A;1-1","Complying:REQ-000001/A;2-1","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_WebSE_TraceLinkVerifications(sAction,sSelectNode,sVerificationNode,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_WebSE_TraceLinkVerifications"
   Dim bReturn,sValue,objPage,aValues
   Set objPage=Browser("TeamcenterWeb").Page("SystemsEngineering")
	Fn_WebSE_TraceLinkVerifications=false

		Select Case sAction

						Case "VerifyTracelink"

										If sSelectNode<>"" Then
												'Select the Desired Node from the BOM Table
												bReturn=Fn_WebSE_BOMTableOperations("Select",sSelectNode,"","")
												If bReturn=false Then
														Fn_WebSE_TraceLinkVerifications = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_TraceLinkVerifications : Failed to Select the Node ["+CStr(sSelectNode)+"] . ")
														Exit Function
											 Else
														Fn_WebSE_TraceLinkVerifications = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_TraceLinkVerifications : Node ["+CStr(sSelectNode)+"] Selected Successfully ")		   
											 End If
									End If

								  'Activate the Tracelink Tab
									objPage.WebElement("Tracelink").Click 0,0,micLeftBtn
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_TraceLinkVerifications : Successfully Activated Tracelinks Tab")	
									objPage.Sync
									wait 2

								 'Synchronise
								 call Fn_Web_ReadyStatusSync(1)

								 'verify the Desired Values
								 aValues=split(sVerificationNode,":",-1,1)
							If aValues(0)="Defining" Then
										objPage.WebElement("ObjectsValue").SetTOProperty "innertext",aValues(1)
										 If objPage.WebElement("ObjectsValue").Exist(5) Then
												Fn_WebSE_TraceLinkVerifications = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_TraceLinkVerifications : Node ["+CStr(aValues(1))+"] Verified successfully in Defining Objects Field")	
										Else
												Fn_WebSE_TraceLinkVerifications = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_TraceLinkVerifications : Failed to verify the Node ["+CStr(aValues(1))+"] in Defining Objects Field")
												Exit Function	
										 End If
							Elseif aValues(0)="Complying" then
										objPage.WebElement("ObjectsValue").SetTOProperty "innertext",aValues(1)
										 If objPage.WebElement("ObjectsValue").Exist(5) Then
												Fn_WebSE_TraceLinkVerifications = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_TraceLinkVerifications : Node ["+CStr(aValues(1))+"] Verified successfully in Complying Objects Field")	
										Else
												Fn_WebSE_TraceLinkVerifications = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_TraceLinkVerifications : Failed to verify the Node ["+CStr(aValues(1))+"] in Complying Objects Field")
												Exit Function	
										 End If
							End If

		End Select
		Set objPage=nothing
End Function 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Fn_WebSE_BOMTableOperations(sAction, sNodeName)

'Function Name		:	Fn_WebSE_BOMTable_ViewEditOperations
'@@
'@@    Description				 :	Function Used to perform Operation On  View/Edit Link 
'@@
'@@    Parameters			   :	1.sAction : Action Name
'@@												  2.sNodeName : Node Path Or Item Name [where you want  to  click on link]
'@@                                             					 
'@@
'@@    Return Value		   	   : 	True Or False 
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client And SE perspective should be open
'@@
'@@    Examples					:	
'@@										
'@@  									:      Fn_WebSE_BOMTable_ViewEditOperations("View","REQ-191988/A;1-Req1")
'@@  									:	 Fn_WebSE_BOMTable_ViewEditOperations("Edit","REQ-191988/A;1-Req1")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Avinash Jagdale 			26-April-2012		1.0				 											Koustubh Watwe										
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebSE_BOMTable_ViewEditOperations(sAction, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_WebSE_BOMTable_ViewEditOperations"
		' Declaration of an Variable
Dim iColPos,objDialog,bFlag,ICount,iRowIndex,objLink
		' Initialization of an Variable
		Set objDialog = Browser("TeamcenterWeb").Page("SystemsEngineering").WebTable("SEBOMTable")
		bFlag = False
		Fn_WebSE_BOMTable_ViewEditOperations = False
		
		'Operations Case
		Select Case sAction
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
					Case "View", "Edit"

								iRowIndex = Fn_WebUI_TableRowIndex(objDialog,sNodeName, "")

								iColPos=Fn_WebUI_TableRowIndex(objDialog,sNodeName, "")
								'iColPos=CInt(iColPos)+1
                                If  iRowIndex <> -1 and iColPos <> -1 Then
											For ICount=0 to 1
													 Set objLink = objDialog.ChildItem(iRowIndex,iColPos,"Link",ICount)
														If TypeName(objLink) <> "Nothing" Then
																	If Trim(objLink.GetROProperty("text")) = Trim(sAction) Then
                                                                        objLink.Click
																		Fn_WebSE_BOMTable_ViewEditOperations = True
																		Exit For
																	End If
													   End If
										 Next
								End If
			
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
                 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebSE_BOMTable_ViewEditOperations : Invalid case [ "& sAction & " ].")	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select
	If  Fn_WebSE_BOMTable_ViewEditOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebSE_BOMTable_ViewEditOperations : Executed successfully with case  [ "& sAction & " ].")	
	End If
	Set objDialog = nothing
	Set objLink = Nothing
    
End Function