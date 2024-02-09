Option Explicit
																								'Function List
'************************************************************************************************************************************************************************************************************
'000. Fn_SISW_WebPSE_GetObject(sObjectName)
'001. Fn_WebPSE_BOMTableOperations()
'002. Fn_WebPSE_AddComponent()
'003. Fn_WebPSE_SetRevisionRule()
'004. Fn_WebPSE_PackUnPackBOMLine()
'005. Fn_WebPSE_CreateSnapshot()
'006. Fn_WebPSE_VariantConfigurationOperations()
'007. Fn_WebPSE_FindComponentInDisplayOperations()
'008. Fn_PSE_AddSubstitute()
'009. Fn_WebPSE_ListSubstituteOperations()
'010. Fn_PSE_ReplaceBOMline()
'************************************************************************************************************************************************************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_WebPSE_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_WebPSE_GetObject("TeamcenterWebStructure")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sachin Joshi		 27-Sept-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_WebPSE_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\WebStructureMananger.xml"
	Set Fn_SISW_WebPSE_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebPSE_BOMTableOperations
'@@
'@@    Description				 :	Function Used to perform Operation On BOM Table
'@@
'@@    Parameters			   :	1.sAction : Action Name
'@@												  2.sNodeName : Node Path Or Item Name
'@@												  3.sColumn : Column Name
'@@												  4.sCellValue : Expected Value ( for case GetImage its an Image Number )
'@@
'@@    Return Value		   	   : 	True Or False Or Column Names Or Image Name
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client And PSEperspective should be open
'@@
'@@    Examples					:	
'@@				cases		NodeSelect / Select 		- Call Fn_WebPSE_BOMTableOperations("NodeSelect","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@												 		- Call Fn_WebPSE_BOMTableOperations("Select","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View) @2:000057/A;1-asm","","")
'@@							NodeDeSelect / Deselect 	- Call Fn_WebPSE_BOMTableOperations("NodeDeSelect","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							NodeVerify / Exist /Exists 	- Call Fn_WebPSE_BOMTableOperations("NodeVerify","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View):000056/A;1-SubItem2","","")
'@@							CellVerify					- Call Fn_WebPSE_BOMTableOperations("CellVerify","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View):000056/A;1-SubItem2","Item Type","Item")
'@@							Collapse					- Call Fn_WebPSE_BOMTableOperations("Collapse","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							MultiSelect					- Call Fn_WebPSE_BOMTableOperations("MultiSelect","000054/A;1-TopItem (View)~000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							Expand						- Call Fn_WebPSE_BOMTableOperations("Expand","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							ExpandBelow					- Call Fn_WebPSE_BOMTableOperations("ExpandBelow","000054/A;1-TopItem (View)","","")
'@@							GetImage					- Call Fn_WebPSE_BOMTableOperations("GetImage","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@														- Call Fn_WebPSE_BOMTableOperations("GetImage","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","1")
'@@														- Call Fn_WebPSE_BOMTableOperations("GetImage","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","2")
'@@							FirstElement				- Call Fn_WebPSE_BOMTableOperations("FirstElement","","","")
'@@							ColumnExist					- Call Fn_WebPSE_BOMTableOperations("ColumnExist","","Name~BOM Line~Find No.","")
'@@							ColumnClick					- Call Fn_WebPSE_BOMTableOperations("ColumnClick","","Name~BOM Line~Find No.","")
'@@							CellEdit					- Call Fn_WebPSE_BOMTableOperations("CellEdit","000015/A;1-top (View):000016/A;1-sub","Find No.","20") 
'@@							ClearSelection				- Call Fn_WebPSE_BOMTableOperations("ClearSelection","","","")
'@@							VerifyBackgroundColour		- Call Fn_WebPSE_BOMTableOperations("VerifyBackgroundColour","000015/A;1-top (View):000016/A;1-sub","","green")
'@@
'@@	   History:				Developer Name				Date				Rev. No.		Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Sandeep Navghane			28-Apr-2011			1.0															Sunny Ruparel
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				10-May-2011			1.0				made function simpler. added reusable code .
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				12-May-2011			1.0				Added case CellEdit
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				12-May-2011			1.0				Added case ClearSelection
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				28-Nov-2011			1.0				Added case VerifyBackgroundColour
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe				14-Dec-2011			1.0				Modifeid case GetImage
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebPSE_BOMTableOperations(sAction, sNodeName, sColumn, sCellValue)
	GBL_FAILED_FUNCTION_NAME="Fn_WebPSE_BOMTableOperations"
		' Declaration of an Variable
		Dim objDialog, objImg, objWebChk, objLink, objButton
        Dim aElements, aSubElement, iCounter, bFlag, jCounter, iRowCnt, iColPos, iOuterCnt, sText, iCounter2, iRowCnt2
		Dim objTRs, objTDs, objElements
		Dim sColour, Str
		' Initialization of an Variable
		Set objDialog = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("BOMTable")
		bFlag = False
		Fn_WebPSE_BOMTableOperations = False
		
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
					Fn_WebPSE_BOMTableOperations = True
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
        		Case "NodeSelect", "Select"
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
							Call Fn_WebPSE_BOMTableOperations("ClearSelection", "", "", "")
							Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
							If TypeName(objWebChk) <> "Nothing" Then
								If objWebChk.GetROProperty("checked") = "0" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = True
								elseIf objWebChk.GetROProperty("checked") = "1" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = True
								End If
							End If
							Set objWebChk = Nothing
					else
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebPSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] Selected Successfully ")
					Else
						Fn_WebPSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to Select the Nod ["+CStr(sNodeName)+"] . ")
					End If

			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "NodeVerify", "Exist", "Exists"
					' Write the Log of Success or Failure
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
						Fn_WebPSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node  ["+CStr(sNodeName)+"] Verified Successfully. ")
					Else
						Fn_WebPSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to Verify the Node  ["+CStr(sNodeName)+"] . ")
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
						Fn_WebPSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations :  Value [ " & sCellValue & "]  is successfully verified for column [ " & sColumn & " ] at Node ["+CStr(sNodeName)+"]. ")
					Else
						Fn_WebPSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations :  Failed to verify value [ " & sCellValue & "]  for column [ " & sColumn & " ] at Node ["+CStr(sNodeName)+"]. ")
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
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations: node  ["+CStr(sNodeName)+"] was already expanded.")
									else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations: can not expand node  ["+CStr(sNodeName)+"].")
									End If
							End If						
							Set objImg = Nothing
					else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations: node  ["+CStr(sNodeName)+"] does not exist in BOM table.")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebPSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node  ["+CStr(sNodeName)+"] expanded successfully. ")
					Else
						Fn_WebPSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to expand node  ["+CStr(sNodeName)+"]")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -				
			Case "ExpandBelow"
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
							Call Fn_WebPSE_BOMTableOperations("ClearSelection", "", "", "")
							Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
							If TypeName(objWebChk) <> "Nothing" Then
								    If objWebChk.GetROProperty("checked") = "0" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = Fn_Web_MenuOperation("Select",Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewExpandBelow"))
								elseIf objWebChk.GetROProperty("checked") = "1" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = Fn_Web_MenuOperation("Select",Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewExpandBelow"))
								End If
							End If
							Set objWebChk = Nothing
					else
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebPSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node  ["+CStr(sNodeName)+"] expanded successfully. ")
					Else
						Fn_WebPSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to expand node  ["+CStr(sNodeName)+"]")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "GetImage"
'				objDialog.RefreshObject
					If sCellValue = "" then 
						sCellValue = 1
					Else 
						sCellValue = cInt("" & sCellValue)
					End If
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
					If iRowCnt <> -1 and iColPos <> -1 Then
							Set objImg = objDialog.ChildItem(iRowCnt, iColPos, "Image", (sCellValue - 1))
							If TypeName(objImg) <> "Nothing" Then
									strImageName=Split(objImg.GetROProperty("file name"),".")
									Fn_WebPSE_BOMTableOperations=strImageName(0)
									bFlag = True
							Else
									Fn_WebPSE_BOMTableOperations=False
							End If						
							Set objImg = Nothing
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_WebPSE_BOMTableOperations = strImageName(0)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Image [ " & Fn_WebPSE_BOMTableOperations & " ]  associated with node  ["+CStr(sNodeName)+"].")
					Else
									Fn_WebPSE_BOMTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations :  Failed to retrieve image name of node  ["+CStr(sNodeName)+"].")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "Collapse"
'		
'					aSubElement = Split(sNodeName, ":", -1, 1)
'					jCounter = 0
'					iRowCnt = objDialog.RowCount
'					iCounter = 1
'					Do While 1 = 1
'							objDialog.RefreshObject
'							iRowCnt = objDialog.RowCount
'							' For Last Node of an Element
'							If jCounter=UBound(aSubElement) Then
'									If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
'												Set objImg = objDialog.ChildItem(iCounter, 2, "Image", 0)
'												If TypeName(objImg) <> "Nothing" Then
'														
'													If objImg.GetROProperty("file name") = "minus.png" Then
'																objImg.Click 1,1, micLeftBtn
'													End If
'														
'												End If
'												bFlag = True
'												jCounter = jCounter + 1
'												Set objImg = Nothing
'												Exit Do 
'									End If
'							Else
'									' For the Node Hierarchy of an Element
'									If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
'												Set objImg = objDialog.ChildItem(iCounter, 2, "Image", 0)
'												If objImg.GetROProperty("file name") = "plus.png" Then
'															objImg.Click 1,1, micLeftBtn
'												End If
'												jCounter = jCounter + 1
'												Set objImg = Nothing
'									End If
'							End If
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
									Fn_WebPSE_BOMTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node  ["+CStr(sNodeName)+"] collapsed Successfully. ")
					Else
									Fn_WebPSE_BOMTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to collapse the Node  ["+CStr(sNodeName)+"] . ")
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
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to find node ["+CStr(aElements(iOuterCnt))+"] . ")	
										Exit for
								End If
						Next
						If bFlag = True Then
								' For Success Log
								Fn_WebPSE_BOMTableOperations = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Multiple Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] selected Successfully. ")
						Else
								' For Failure Log
								Fn_WebPSE_BOMTableOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to Select the Multiple Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] . ")	
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
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to find node ["+CStr(aElements(iOuterCnt))+"] . ")	
										Exit for
								End If
						Next
					' Write the Log of Success or Failure
					If bFlag = True Then
							Fn_WebPSE_BOMTableOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Deselected  Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] successfully. ")
					Else
							' For Failure Log
							Fn_WebPSE_BOMTableOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to deselect  the Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] . ")	
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
									Fn_WebPSE_BOMTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Node  ["+CStr(sNodeName)+"]  Clicked Successfully. ")
					Else
									Fn_WebPSE_BOMTableOperations = False
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
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Column  [ " & aElements(iCounter) & " ]  exists in BOM table. ")
							else
								bFlag = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to check existence of column  [ " & aElements(iCounter) & " ]. ")
								Exit for
							End If
						Next
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_WebPSE_BOMTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : All Columns  ["+CStr(replace(sColumn,"~", ", "))+"]  exists in BOM table. ")
					Else
									Fn_WebPSE_BOMTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to check existence of column(s) ["+CStr(replace(sColumn,"~", ", "))+"]. ")
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
								Fn_WebPSE_BOMTableOperations = iRowCnt2
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : First Element ["+CStr(iRowCnt2)+"] Present in the BOM Table ")
						Else
								Fn_WebPSE_BOMTableOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : No First Element Found in BOM Table ")
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
							Fn_WebPSE_BOMTableOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Column Heading  ["+CStr(sColumn)+"]  Clicked Successfully. ")
						Else
							Fn_WebPSE_BOMTableOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Column Heading  ["+CStr(sColumn)+"]  Not Found to Click. ")
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
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Node ["+CStr(sNodeName)+"] found.")
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
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebPSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Modified Cell [ "+CStr(sNodeName)+" ][ " & sColumn & " ] to value [ " & sCellValue & " ] Successfully.")
					Else
						Fn_WebPSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to modify Cell [ "+CStr(sNodeName)+" ][ " & sColumn & " ] to value [ " & sCellValue & " ].")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "VerifyBackgroundColour", "VerifyCellBackgroundColour"
					objDialog.RefreshObject
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If sAction <> "VerifyCellBackgroundColour" Then
						iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")' For Row Number
					Else
						If sColumn = "" Then sColumn = "Name"
						iColPos = Fn_WebUI_TableColumnIndex(objDialog, sColumn )' For Row Number
					End If
					
					bFlag = False
					If iRowCnt <> -1 Then
						Set objTDs = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("BOMTable").ChildItem(iRowCnt, iColPos,"WebElement","0")
						Set objTRs = objTDs.object.parentNode
						If InStr(1,Environment.Value("WebBrowserName"),"IE")>0 Then
							If sAction <> "VerifyCellBackgroundColour" Then
								do while lcase(trim(objTRs.NodeName)) <> "tr"
									Set objTRs = objTRs.parentNode
								loop
							End If
							Style = objTRs.style.cssText
							Style = lcase(trim(Style))
						ElseIf InStr(1,Environment.Value("WebBrowserName"),"FF")>0 Then
							Style = objTRs.getAttribute("style")
						End If
						If instr(Style,"background-color:") > 0 OR objTRs.getAttribute("bgcolor") <> "" Then
							If objTRs.getAttribute("bgcolor") <> "" then
								' not yet implemented
							Else
								bgCol = instr(Style,"background-color:")
								If inStr(bgCol,Style,";") > 0 Then
									Str = trim(mid (Style, bgCol +  len("background-color:"),  instr(bgCol, Style,";") - bgCol -  len("background-color:")))
								Else
									Str = trim(mid (Style, bgCol +  len("background-color:"),  len(style) - bgCol +  len("background-color:")))
								End If
							End If
						End If
						Select Case lCase(Str)
							Case "green", "rgb(64, 224, 208)","#40e0d0", "rgb(0, 139, 139)","#008b8b","#dee7ff","rgb(222, 231, 255)"
								sColour = "green"
							Case "transparent","white", "#ffffff"
								sColour = "white"
'							Case "#dee7ff"
'								sColour = "light blue"
						End Select
						If lCase(sCellValue) = sColour then bFlag = True
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_WebPSE_BOMTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_WebPSE_BOMTableOperations : Successfully  verified background colour [ " & sCellValue & " ] of Cell [ "+CStr(sNodeName)+" ].")
					Else
						Fn_WebPSE_BOMTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WebPSE_BOMTableOperations : Failed to verify background colour [ " & sCellValue & " ] of Cell [ "+CStr(sNodeName)+" ].")
					End If					
		End Select
		Set objDialog = Nothing
		Set objImg = Nothing
		Set objWebChk = Nothing
		Set objLink = Nothing
		Set objTDs = Nothing
		Set objTRs = Nothing
End Function
''*********************************************************		Function for adding PSE Compnent		***********************************************************************
'Function Name		:				Fn_WebPSE_AddComponent

'Description			 :		 		 Enter Details for Adding Component to Assembly.

'Parameters			   :	 			Function To Add Component using Menu (Edit;Add)
'													1.sAction       : Action to perform
'													2.sBOMLine  : BOM Line to select
'													3.sItemID       : Item Id
'													4.sFindNo	   : Find No.
'													5.sQuantity	    : No Of Quantity for component to add
'													6.sItem		 :  Specify the action(i.e. QuickAdd, MenuAdd)
'													7.sName	    : Fin No of the Newly Added Component.
'													8.sDescription:  Specify the action(i.e. QuickAdd, MenuAdd)
'												      9.sType		:  Item Type
'												    10.sOwningUser : Owning User name
'												    11.sOwningGroup : owning user group
'												    12.sBtnName    :  Button name
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		PSE window should be displayed BOMLine Should be selected. Item to add is already created..

'Examples				:			call Fn_WebPSE_AddComponent("Add", "000061/A;1-top (View)", "000063", "", "1", "", "", "", "", "", "", "")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			11-May-2011			1.0				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_WebPSE_AddComponent(sAction, sBOMLine, sItemID, sFindNo, sQuantity, sItem, sName, sDescription, sType, sOwningUser, sOwningGroup, sBtnName)
		GBL_FAILED_FUNCTION_NAME="Fn_WebPSE_AddComponent"
		Dim bFlag, objDialog, objButtonTable, sMenu
		Fn_WebPSE_AddComponent = False
		Set objDialog = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("AddNewBOMLine")
		Set objButtonTable = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ButtonTable")
		' selecting BOM Line.
		If sBOMLine <> "" Then
			bFlag = Fn_WebPSE_BOMTableOperations("Select", sBOMLine ,"","")
			If NOT(bFlag) Then
                        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_AddComponent : Failed to select BOM Line [ " & sBOMLine  & " ].")
				Exit function
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_AddComponent : Successfully selected BOM Line [ " & sBOMLine  & " ].")
			End If
		End If

		'performing menu operation
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "EditAdd")
		If sMenu <> "False" Then
			bFlag = Fn_Web_MenuOperation("Select", sMenu)
			If NOT(bFlag) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_AddComponent : Failed to perform menu operation [ " & sMenu & " ].")
				Exit function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_AddComponent : Successfully performed menu operation [ " & sMenu & " ].")
			End If
		else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_AddComponent : Failed to get menu from XML.")
				Exit function
		End If
	   	Select Case sAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "Add"
						' set item id
						If sItemID <> "" then
							 Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_AddComponent", objDialog, "ItemID", sItemID)
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_AddComponent : Field Item ID is compulsory.")
								Exit function
						End if

						' set find no.
						If sFindNo <> "" then
							Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_AddComponent", objDialog, "FindNo", sFindNo)
						end if

						' setting quantity
						If sQuantity <> "" then
							Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_AddComponent", objDialog, "Quantity", sQuantity)
						end if

						'clickin on OK button
						If sBtnName = "" Then
							Call Fn_Web_UI_Button_Click("Fn_WebPSE_AddComponent", objButtonTable, "OK")
						else
							Call Fn_Web_UI_Button_Click("Fn_WebPSE_AddComponent", objButtonTable, sBtnName)
						End If
						Fn_WebPSE_AddComponent = True
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "FindAndAdd"
					'For future use
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_AddComponent : Invalid case [ " & sAction  & " ].")
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		End Select
		If Fn_WebPSE_AddComponent Then
                  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_AddComponent : executed successfully with case [ " & sAction  & " ].")
		End If
		Set objDialog = nothing
		Set objButtonTable = nothing
End Function

''*********************************************************		Function to Apply Revision Rule	***********************************************************************
'Function Name		:			  Fn_WebPSE_SetRevisionRule

'Description			 :		 	Function to Apply Revision Rule.	

'Parameters			   :	 		     Function To Add Component using Menu (Edit;Add)
'													1.sAction       : Action to perform
'													2.dicSetWebRevRule  : Dictionary Object
'Return Value		   : 			     TRUE \ FALSE

'Pre-requisite			:		 	PSE window should be displayed.

'Examples				:			   Dim dicSetWebRevRule
'										Set dicSetWebRevRule = CreateObject("Scripting.Dictionary")
'										
'										With dicSetWebRevRule
'											.Add "sRevisionRule","Latest Working" 
'											.Add"sDate",""
'											.Add"sUnitNumber","" 
'											.Add"sEndItem", "000305"
'											.Add"sEndItemBy",""
'											.Add"sOverrideFolder",""
'											.Add"sOverrideFolderBy",""
'										End With
'										call Fn_WebPSE_SetRevisionRule("Set", dicSetWebRevRule )

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			13-May-2011			1.0				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebPSE_SetRevisionRule(sAction, dicSetWebRevRule )
	GBL_FAILED_FUNCTION_NAME="Fn_WebPSE_SetRevisionRule"
	Dim objButtonPanel, objRevRulePanel, sMenu, bFlag
	Fn_WebPSE_SetRevisionRule = False
	Set objButtonPanel = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ButtonTable")
	Set objRevRulePanel = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("SetRevisionRule")

		'performing menu operation
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewRevisionRule")
		If sMenu <> "False" Then
			bFlag = Fn_Web_MenuOperation("Select", sMenu)
			If NOT(bFlag) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_SetRevisionRule : Failed to perform menu operation [ " & sMenu & " ].")
				Exit function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_SetRevisionRule : Successfully performed menu operation [ " & sMenu & " ].")
			End If
		else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_SetRevisionRule : Failed to get menu from XML.")
				Exit function
		End If
		If Fn_Web_UI_ObjectExist("Fn_WebPSE_SetRevisionRule", objRevRulePanel) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_SetRevisionRule : Failed to open [ Set Transient Revision Rule ] panel.")
				Exit function
		end if
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Set"
				' setting Revision Rule
				If dicSetWebRevRule("sRevisionRule") <> "" Then
					Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_SetRevisionRule", objRevRulePanel, "RevisionRule", dicSetWebRevRule("sRevisionRule"))
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_SetRevisionRule : successfully set [ Revision Rule ] to [ " & dicSetWebRevRule("sRevisionRule")   & " ].")
				End If

				'setting date
				If dicSetWebRevRule("sDate") <> "" Then
					Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_SetRevisionRule", objRevRulePanel, "Date", dicSetWebRevRule("sDate"))
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_SetRevisionRule : successfully set [ Date ] to [ " & dicSetWebRevRule("sDate")   & " ].")
				End If

				'setting unit
				If dicSetWebRevRule("sUnitNumber") <> "" Then
					Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_SetRevisionRule", objRevRulePanel, "Unit", dicSetWebRevRule("sUnitNumber") )
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_SetRevisionRule : successfully set [ Unit ] to [ " & dicSetWebRevRule("sUnitNumber")   & " ].")
				End If

				'setting end item
				Select Case dicSetWebRevRule("sEndItemBy")
					Case "Find"
							' for future use
					Case else
							If dicSetWebRevRule("sEndItem") <> "" Then
									Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_SetRevisionRule", objRevRulePanel, "EndItem", dicSetWebRevRule("sEndItem"))
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_SetRevisionRule : successfully set [ End Item ] to [ " & dicSetWebRevRule("sEndItem")   & " ].")
							End If
				End Select

				'setting override folder
				Select Case dicSetWebRevRule("sOverrideFolderBy") 
					Case "Find"
							' for future use
					Case else
						If dicSetWebRevRule("sOverrideFolder") <> "" Then
									Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_SetRevisionRule", objRevRulePanel, "OverrideFolder", dicSetWebRevRule("sOverrideFolder") )
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_SetRevisionRule : successfully set [ Override Folder ] to [ " & dicSetWebRevRule("sOverrideFolder")   & " ].")
						End If
				End Select
				
				'clicking on OK
				Call Fn_Web_UI_Button_Click("Fn_WebPSE_SetRevisionRule", objButtonPanel, "OK")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_SetRevisionRule : successfully clicked [ OK ] button.")
				Fn_WebPSE_SetRevisionRule = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_SetRevisionRule : Invalid case [ " & sAction  & " ].")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	If Fn_WebPSE_SetRevisionRule Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_SetRevisionRule : executed successfully with case [ " & sAction  & " ].")
	End If
	Set objButtonPanel = nothing
	Set objRevRulePanel = nothing
End Function
''*********************************************************		Function to Pack and Unpack select BOM Line ***********************************************************************
'Function Name		:			  Fn_WebPSE_PackUnPackBOMLine

'Description			 :		 	Function to Pack and Unpack select BOM Line

'Parameters			   :	 		     Function To Add Component using Menu (Edit;Add)
'													1.sAction       : Action to perform
'													2.sBOMLine  : BOM Line to select
'Return Value		   : 			     TRUE \ FALSE

'Pre-requisite			:		 	PSE window should be displayed.

'Examples				:			   
'										Call Fn_WebPSE_PackUnPackBOMLine("Pack", "000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)")
'										Call Fn_WebPSE_PackUnPackBOMLine("PackAll", "")
'										Call Fn_WebPSE_PackUnPackBOMLine("Unpack", "000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)")
'										Call Fn_WebPSE_PackUnPackBOMLine("UnpackAll", "")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			13-May-2011			1.0				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebPSE_PackUnPackBOMLine(sAction, sBOMLine)
		GBL_FAILED_FUNCTION_NAME="Fn_WebPSE_PackUnPackBOMLine"
		Dim bReturn, sMenu 
		Fn_WebPSE_PackUnPackBOMLine = False

		' fetching menu from XML file
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewPackAll")
		If sMenu = "False" Then		
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_PackUnPackBOMLine : Failed to get menu from XML.")
			Exit function
		End If
		' selecting BOM Line
		If sBOMLine <> ""  Then
			bReturn = Fn_WebPSE_BOMTableOperations("Select",sBOMLine,"","")
			If NOT( bReturn ) Then
				'failed to select bom Line
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_PackUnPackBOMLine : Failed to select BOM Line [ " & sBOMLine  & " ].")
				Exit function
			End If
		End If

		Select Case sAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "PackAll", "Pack"
						'performing menu operation
						If NOT(Fn_Web_MenuOperation("IsChecked", sMenu)) Then
								bReturn = Fn_Web_MenuOperation("Select", sMenu)
								If NOT(bReturn) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_PackUnPackBOMLine : Failed to perform menu operation [ " & sMenu & " ].")
									Exit function
								End If
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_PackUnPackBOMLine : Successfully performed menu operation [ " & sMenu & " ].")
						Fn_WebPSE_PackUnPackBOMLine = True
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "UnpackAll", "Unpack"
						'performing menu operation
						If Fn_Web_MenuOperation("IsChecked", sMenu) Then
								bReturn = Fn_Web_MenuOperation("Select", sMenu)
								If NOT(bReturn) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_PackUnPackBOMLine : Failed to perform menu operation [ " & sMenu & " ].")
									Exit function
								End If
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_PackUnPackBOMLine : Successfully performed menu operation [ " & sMenu & " ].")
						Fn_WebPSE_PackUnPackBOMLine = True
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_PackUnPackBOMLine : Invalid case [ " & sAction  & " ].")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			End Select
			If Fn_WebPSE_PackUnPackBOMLine Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_PackUnPackBOMLine : executed successfully with case [ " & sAction  & " ].")
			End If

End Function

''*********************************************************		Function to create snapshot ***********************************************************************
'Function Name		:			  Fn_WebPSE_CreateSnapshot

'Description			 :		 	Function to create snapshot

'Parameters			   :	 		     Function To Add Component using Menu (Edit;Add)
'													1.sAction       : Action to perform
'													2.sSnapshotName  : snapshot name
'													2.sDescription  : description
'Return Value		   : 			     TRUE \ FALSE

'Pre-requisite			:		 	PSE window should be displayed.

'Examples				:			   
'										Call Fn_WebPSE_CreateSnapshot("Create", "name2", "desc")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			13-May-2011			1.0				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebPSE_CreateSnapshot(sAction, sSnapshotName, sDescription)
		GBL_FAILED_FUNCTION_NAME="Fn_WebPSE_CreateSnapshot"
		Dim bReturn, sMenu, objCreateSnapshotPanel, objButtonPanel 
		Fn_WebPSE_CreateSnapshot = False
		Set objCreateSnapshotPanel = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("CreateSnapshot")
		Set objButtonPanel = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ButtonTable")
		' fetching menu from XML file
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ToolsSnapshot")
		If sMenu = "False" Then		
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_CreateSnapshot : Failed to get menu from XML.")
			Exit function
		End If
            'performing menu operation
		bReturn = Fn_Web_MenuOperation("Select", sMenu)
		If NOT(bReturn) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_CreateSnapshot : Failed to perform menu operation [ " & sMenu & " ].")
			Exit function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_CreateSnapshot : Successfully performed menu operation [ " & sMenu & " ].")
		Select Case sAction
				Case"Create"
						' setting snapshot
						If sSnapshotName <> "" Then
                                          Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_CreateSnapshot", objCreateSnapshotPanel, "SnapshotName", sSnapshotName)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_CreateSnapshot : successfully set [ Snapshot Name ] to [ " & sSnapshotName   & " ].")
						End If

						'setting dscription
						If sDescription <> "" Then
                                          Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_CreateSnapshot", objCreateSnapshotPanel, "Description", sDescription)
                                          Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_CreateSnapshot : successfully set [ Description ] to [ " & sDescription   & " ].")
						End If

						' clicking no OK
                                    Call Fn_Web_UI_Button_Click("Fn_WebPSE_CreateSnapshot", objButtonPanel, "OK")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_CreateSnapshot : successfully clicked [ OK ] button.")
						Fn_WebPSE_CreateSnapshot = True
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_CreateSnapshot : Invalid case [ " & sAction  & " ].")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		End Select
		If Fn_WebPSE_CreateSnapshot Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_CreateSnapshot : executed successfully with case [ " & sAction  & " ].")
		End If
		Set objCreateSnapshotPanel = nothing
		Set objButtonPanel = nothing
End Function
''*********************************************************		Function to perform operations on Variant Configuration window***********************************************************************
'Function Name		:			  Fn_WebPSE_VariantConfigurationOperations

'Description		    :		 	Function to perform operations on Variant Configuration window

'Parameters			   :	 		  1.sAction       : Action to perform
'									2.sVariants    : 
'									3.sValues      :
'									4.sVarName  :
'									5.sVarDescription	:
'									6.sSavedConfiguration	:

'Return Value		   : 			     TRUE \ FALSE

'Pre-requisite		    :		 	PSE window should be displayed.

'Examples			    :			   
'										Call Fn_WebPSE_VariantConfigurationOperations("SetVariant", "001142:Levl1 (String)", "B", "", "", "")
'										Call Fn_WebPSE_VariantConfigurationOperations("LoadVariantConfiguration", "001142:Levl1 (String)", "B", "", "", "C")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			1-Jun-2011			1.0				
'										Sachin Joshi			03-Jun-2011			1.0				Modified code to select Value from Drop Down List
'										Sachin Joshi			08-Jun-2011			1.0				To compare variant Label
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			26-Sept-2012			1.0				Added case VerifyVariantConfiguration
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebPSE_VariantConfigurationOperations(sAction, sVariants, sValues, sVarName, sVarDescription, sSavedConfiguration)
	GBL_FAILED_FUNCTION_NAME="Fn_WebPSE_VariantConfigurationOperations"
	Dim bFlag, sMenu, objVarConfigPanel, objButtonPanel,objWebButton
	Dim iCount, iRows, aVariants, iRowCnt, objElement, aValues,arr
	Dim objSavedConfig 
	Fn_WebPSE_VariantConfigurationOperations = False
	bFlag = False
	Set objVarConfigPanel = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("SetVariants")
	Set objButtonPanel = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ButtonTable")
'		' fetching menu from XML file
	sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewVariantConfiguration")
	If sMenu = "False" Then		
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_VariantConfigurationOperations : Failed to get menu from XML.")
		Exit function
	End If
	'performing menu operation
	If not objVarConfigPanel.Exist(8) Then
		bReturn = Fn_Web_MenuOperation("Select", sMenu)
		If NOT(bReturn) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_VariantConfigurationOperations : Failed to perform menu operation [ " & sMenu & " ].")
			Exit function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_VariantConfigurationOperations : Successfully performed menu operation [ " & sMenu & " ].")
	End If

		' setting values to variant options
		If sVariants <> "" AND sValues <> ""  Then
			aVariants = split(sVariants,"~")
			aValues = split(sValues,"~")
			If UBound(aVariants) = UBound(aValues) Then
				iRows = cInt(objVarConfigPanel.RowCount)
				For iCount = 0 to uBound(aVariants)
						bFlag = False
						For iRowCnt = 1 to iRows 
								set objElement = objVarConfigPanel.ChildItem(iRowCnt, 0,"WebElement",0)
									If Instr(1, objElement.GetROProperty("innerhtml"), aVariants(iCount)) Then
										'	------------------------------------------------
										'   Commented Code as set is not working
										'	set objElement = objVarConfigPanel.ChildItem(iRowCnt, 2,"WebEdit",0)
										'	objElement.set aValues(iCount)
										'	------------------------------------------------
											'Added by Sachin
											wait 3
											iRowCnt=objVarConfigPanel.GetRowWithCellText(objElement.GetROProperty("innerhtml"))
											set objWebButton = objVarConfigPanel.ChildItem(iRowCnt, 2,"WebButton",0)
											objWebButton.Click 1,1
											Call Fn_WEB_UI_Object_SetTOProperty("",objVarConfigPanel.WebElement("VarValue"),"innertext",aValues(iCount))
											wait 3
											objVarConfigPanel.WebElement("VarValue").Click 1,1,micLeftBtn
											wait 3
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_VariantConfigurationOperations : Value [ " & aValues(iCount) & " ] set agains [ " & aVariants(iCount) & " ].")
											bFlag = True
											Exit for
									End If
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_VariantConfigurationOperations : Failed to find variant option [ " & aVariants(iCount)  & " ].")
							Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", objButtonPanel, "Cancel")
							Set objVarConfigPanel = nothing
							Set objButtonPanel = nothing
							Set objSavedConfig = nothing
							Exit function
						End If
				Next
			End If
		End If

		Select Case sAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "SetVariant", "LoadVariantConfiguration"
					' do nothing
					' clickin on Ok button
					Select Case sAction
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "SetVariant"
								' Do Nothing
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "LoadVariantConfiguration"
								Set objSavedConfig = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("SavedConfiguration")

								If objButtonPanel.WebButton("Load").Exist(5) Then
									objButtonPanel.WebButton("Load").Click
									wait 2
								End If
								If  objSavedConfig.Exist(5)  = False Then
									Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure"), "Load")
								End If
								
								If objSavedConfig.Exist(15) Then
										If sSavedConfiguration <> ""  Then
											'Added by Sachin
											iRowCnt=objSavedConfig.GetRowWithCellText("Saved Configuration:")
											set objWebButton = objSavedConfig.ChildItem(iRowCnt, 2,"WebButton",0)
											objWebButton.Click 1,1
											wait 2
											Call Fn_WEB_UI_Object_SetTOProperty("",objSavedConfig.WebElement("ConfigType"),"innertext",sSavedConfiguration)
											objSavedConfig.WebElement("ConfigType").Click 1,1,micLeftBtn
											'	------------------------------------------------
											'   Commented Code as set is not working
											'	Call Fn_Web_UI_WebEdit_Set("Fn_WebPSE_VariantConfigurationOperations", objSavedConfig, "SavedConfigName", sSavedConfiguration)
											'	objSavedConfig.WebEdit("SavedConfigName").Set 
											'	------------------------------------------------
										End If
										Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("2ndLevelButtonPanel"), "OK")
								End If
					End Select
								wait 2
								If objButtonPanel.Exist(5) Then
									Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", objButtonPanel, "OK")
								Else
									Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure"), "OK")
								End If
								Fn_WebPSE_VariantConfigurationOperations = True
	        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyVariantConfiguration"
				' setting values to variant options
				If sVariants <> "" AND sValues <> ""  Then
					aVariants = split(sVariants,"~")
					aValues = split(sValues,"~")
					If UBound(aVariants) = UBound(aValues) Then
						iRows = cInt(objVarConfigPanel.RowCount)
						For iCount = 0 to uBound(aVariants)
							bFlag = False
							For iRowCnt = 1 to iRows
								set objElement = objVarConfigPanel.ChildItem(iRowCnt, 0,"WebElement",0)
								If Instr(1, objElement.GetROProperty("innerhtml"), aVariants(iCount)) Then
									set objElement = objVarConfigPanel.ChildItem(iRowCnt, 2,"WebEdit",0)
									If objElement.getROProperty("value") = aValues(iCount) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_VariantConfigurationOperations : Value [ " & aValues(iCount) & " ] set agains [ " & aVariants(iCount) & " ].")
										bFlag = True
										Exit for
									End IF
								End If
							Next
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_VariantConfigurationOperations : Failed to find variant option [ " & aVariants(iCount)  & " ].")
								If objButtonPanel.Exist(5) Then
									Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", objButtonPanel, "Cancel")
								Else
									Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure"), "Cancel")
								End If
								Set objVarConfigPanel = nothing
								Set objButtonPanel = nothing
								Set objSavedConfig = nothing
								Exit function
							End If
						Next
					End If
				End If
				If objButtonPanel.Exist(5) Then
					Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", objButtonPanel, "Cancel")
				Else
					Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure"), "Cancel")
				End If
				Fn_WebPSE_VariantConfigurationOperations = True
'----------------------------------------------------------------------------------------------------------------------------------------
			Case "ClearVariantConfiguration"
						If sVariants <> "" Then
							aVariants = split(sVariants,"~")
							iRows = cInt(objVarConfigPanel.RowCount)
							For iCount = 0 to uBound(aVariants)
								bFlag = False
								For iRowCnt = 1 to iRows
									set objElement = objVarConfigPanel.ChildItem(iRowCnt, 0,"WebElement",0)
									If Instr(1, objElement.GetROProperty("innerhtml"), aVariants(iCount)) Then
										set objElement = objVarConfigPanel.ChildItem(iRowCnt, 2,"WebEdit",0)
										objElement.Set ""
                                        bFlag = True
										Exit For
									End If
								Next
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_VariantConfigurationOperations : Failed to find variant option [ " & aVariants(iCount)  & " ].")
									If objButtonPanel.Exist(5) Then
										Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", objButtonPanel, "Cancel")
									Else
										Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure"), "Cancel")
									End If
									Set objVarConfigPanel = nothing
									Set objButtonPanel = nothing
									Set objSavedConfig = nothing
									Exit function
								End If
							Next
						End If
					If objButtonPanel.Exist(5) Then
						Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", objButtonPanel, "OK")
					Else
						Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure"), "OK")
					End If
					Fn_WebPSE_VariantConfigurationOperations = True
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "VerifySavedRuleList"
					Set objSavedConfig = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("SavedConfiguration")
					If objButtonPanel.WebButton("Load").Exist(5) Then
						objButtonPanel.WebButton("Load").Click
						wait 2
					End If
					If  objSavedConfig.Exist(5)  = False Then
						Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure"), "Load")
					End If

					If objSavedConfig.Exist(15) Then
						If sSavedConfiguration <> "" Then
							aSavedConfiguration = split(sSavedConfiguration,"~")
							iRowCnt=objSavedConfig.GetRowWithCellText("Saved Configuration:")
							set objWebButton = objSavedConfig.ChildItem(iRowCnt, 2,"WebButton",0)
							objWebButton.Click 1,1
							bFlag = True
							wait 2
							For iCount = 0 to uBound(aSavedConfiguration)
								Call Fn_WEB_UI_Object_SetTOProperty("",objSavedConfig.WebElement("ConfigType"),"innertext",aSavedConfiguration(iCount))
								If objSavedConfig.WebElement("ConfigType").Exist = False Then
									bFlag = False
									Exit For
								End If		
							Next
						End IF
					End IF
					If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_VariantConfigurationOperations : Failed to find variant option [ " & aSavedConfiguration(iCount)  & " ].")
								Fn_WebPSE_VariantConfigurationOperations = False
					End If
						Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("2ndLevelButtonPanel"), "Cancel")
						If objButtonPanel.Exist(5) Then
									Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", objButtonPanel, "Cancel")
								Else
									Call Fn_Web_UI_Button_Click("Fn_WebPSE_VariantConfigurationOperations", Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure"), "Cancel")
					End If
					Fn_WebPSE_VariantConfigurationOperations = True
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_VariantConfigurationOperations : Invalid case [ " & sAction  & " ].")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		End Select
		If Fn_WebPSE_VariantConfigurationOperations Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_VariantConfigurationOperations : executed successfully with case [ " & sAction  & " ].")
		End If
		Set objVarConfigPanel = nothing
		Set objButtonPanel = nothing
		Set objSavedConfig = nothing
End Function
''*********************************************************		Function to perform operations on Variant Configuration window***********************************************************************
'Function Name		:			  Fn_WebPSE_FindComponentInDisplayOperations

'Description		    :		 	Function to perform operations on Find and Display window

'Parameters			   :	 		  1.sAction       : Action to perform
'									2. sAttributeName    : 
'									3. sRelationalOperator      :
'									4. sAttributeValue  :
'									5. sLogicalOperator	:
'									6. sExpression	:
'									7. sResult :
'									8. bCloseDialog :

'Return Value		   : 			     TRUE \ FALSE

'Pre-requisite		    :		 	PSE window should be displayed.

'Examples			    :			   

'						Call Fn_WebPSE_FindComponentInDisplayOperations("Find", "Name~Find No.", "=~=", "Top~10", "AND~OR", "", "Number of BOM lines found: 1", "True")
'						Call Fn_WebPSE_FindComponentInDisplayOperations("Up", "", "", "", "", "", "Number of BOM lines found: 0", "")
'						Call Fn_WebPSE_FindComponentInDisplayOperations("Down", "", "", "", "", "", "Number of BOM lines found: 0", "True")
'						Call Fn_WebPSE_FindComponentInDisplayOperations("Verify", "", "", "", "", "", "Number of BOM lines found: 0", "")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			1-Jun-2011			1.0				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebPSE_FindComponentInDisplayOperations(sAction, sAttributeName, sRelationalOperator, sAttributeValue, sLogicalOperator, sExpression, sResult, bCloseDialog)
	GBL_FAILED_FUNCTION_NAME="Fn_WebPSE_FindComponentInDisplayOperations"
	Dim objFindComponentInDisplay, iCnt, objEdit, objAddBtn,sMenu, bReturn
	Dim arrAttribNames, arrRelOperators, arrAttribValues, arrLogicalOperators

	Fn_WebPSE_FindComponentInDisplayOperations = False
	Set objFindComponentInDisplay = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("FindComponentInDisplay")

	If objFindComponentInDisplay.Exist(5) = False Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ActionsFindInDisplay")
		If sMenu = "False" Then		
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : Failed to get menu from XML.")
			Exit function
		End If

		'performing menu operation
		bReturn = Fn_Web_MenuOperation("Select", sMenu)
		If NOT(bReturn) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : Failed to perform menu operation [ " & sMenu & " ].")
			Exit function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_FindComponentInDisplayOperations : Successfully performed menu operation [ " & sMenu & " ].")

	End If
	If objFindComponentInDisplay.Exist(5) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : Failed to launch [ Find and Display ] dialog.")
		Exit function
	End If
	Select Case sAction
		Case "Find"
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			arrAttribNames = Split(sAttributeName, "~")
			arrRelOperators = Split(sRelationalOperator, "~") 
			arrAttribValues = Split(sAttributeValue, "~")
			arrLogicalOperators = Split(sLogicalOperator, "~")
			' object on Add button
			Set objAddBtn = objFindComponentInDisplay.ChildItem(5, 1, "WebButton","0")

			' setting data
			For iCnt = 0 to UBound(arrAttribNames)
				'Attr Name: 1
				If arrAttribNames(iCnt) <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(1, 2, "WebEdit","0")
				End If
				objEdit.set arrAttribNames(iCnt)

				'Relational Operator: 2
				If arrRelOperators(iCnt) <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(2, 2, "WebEdit","0")
					objEdit.set arrRelOperators(iCnt)
				End If

				'Attr Value: 3
				If arrAttribValues(iCnt) <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(3, 2, "WebEdit","0")
					objEdit.set arrAttribValues(iCnt)
				End If

				'Logical Operator: 4
				If arrLogicalOperators(iCnt) <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(4, 2, "WebEdit","0")
					objEdit.set arrLogicalOperators(iCnt)
				End If

				'clicking on Add button
				objAddBtn.Click

				'Expression: 6
				' do nothing
			Next
			' clicking on Find button
			Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ButtonTable").WebButton("Find").Click
			Fn_WebPSE_FindComponentInDisplayOperations = True
		     '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Clear", "Up", "Down"
			Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ButtonTable").WebButton(sAction).Click
			Fn_WebPSE_FindComponentInDisplayOperations = True
		     '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
				Fn_WebPSE_FindComponentInDisplayOperations = True
				'Attr Name: 1
				If sAttributeName <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(1, 2, "WebEdit","0")
					If sAttributeName <> trim(objEdit.GetROProperty("value")) then 
						Fn_WebPSE_FindComponentInDisplayOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sAttributeName   & " ] is not present in Field [ Attr Name ].")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sAttributeName   & " ] is present in Field [ Attr Name ].")
					End If
				End If

				'Relational Operator: 2
				If sRelationalOperator <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(2, 2, "WebEdit","0")
					If sRelationalOperator <> trim(objEdit.GetROProperty("value")) then 
						Fn_WebPSE_FindComponentInDisplayOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sRelationalOperator   & " ] is not present in Field [ Relational Operator ].")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sRelationalOperator   & " ] is present in Field [ Relational Operator ].")
					End If
				End If

				'Attr Value: 3
				If sAttributeValue <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(3, 2, "WebEdit","0")
					If sAttributeValue <> trim(objEdit.GetROProperty("value")) then 
						Fn_WebPSE_FindComponentInDisplayOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sAttributeValue   & " ] is not present in Field [ Attr Value ].")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sAttributeValue   & " ] is present in Field [ Attr Value ].")
					End If
				End If

				'Logical Operator: 4
				If sLogicalOperator <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(4, 2, "WebEdit","0")
					If sLogicalOperator <> trim(objEdit.GetROProperty("value")) then 
						Fn_WebPSE_FindComponentInDisplayOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sLogicalOperator   & " ] is not present in Field [ Logical Operator ].")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sLogicalOperator   & " ] is present in Field [ Logical Operator ].")
					End If
				End If

				'Expression: 6
				If sExpression <> "" Then
					Set objEdit = objFindComponentInDisplay.ChildItem(6, 2, "WebEdit","0")
					If sExpression <> trim(objEdit.GetROProperty("value")) then 
						Fn_WebPSE_FindComponentInDisplayOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sExpression & " ] is not present in Field [ Expression ].")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sExpression & " ] is present in Field [ Expression ].")
					End If
				End If
				'Number of BOM lines found: 
				If sResult <> "" then
					If sResult <> trim(objFindComponentInDisplay.GetCellData(7,2)) Then
						Fn_WebPSE_FindComponentInDisplayOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_FindComponentInDisplayOperations : [ " & sResult & " ] is not present in results.")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_FindComponentInDisplayOperations : Value [ " & sResult & " ] is present in results.")
					End If
				End If
				If bCloseDialog = "" Then bCloseDialog = True
	End Select

	If bCloseDialog <> "" Then
		If bCloseDialog Then
			Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ButtonTable").WebButton("Cancel").Click
		End If
	End If
      If Fn_WebPSE_FindComponentInDisplayOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_FindComponentInDisplayOperations : executed successfully with case [ " & sAction  & " ].")
	End If
	Set objEdit = Nothing
	Set objAddBtn = Nothing
	Set objFindComponentInDisplay = Nothing
End Function
''*********************************************************	 Function to add substitute  ***********************************************************************
'Function Name		:			    Fn_PSE_AddSubstitute

'Description		    :		 	Function to add substitute  

'Parameters			   :	 		1.sAction       : Action to perform
'									2. sItemId    : Item Id
'									3. bFindItem  : Boolean value to open find Item dialog
'									4. sBtnName   : Name of the button to be clicked. 

'Return Value		   : 			     TRUE \ FALSE

'Pre-requisite		    :		 	PSE window should be displayed.

'Examples			    :		Call Fn_PSE_AddSubstitute("AddSubstitute", "000032", "", "Apply")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			12-12-2011			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PSE_AddSubstitute(sAction, sItemId, bFindItem, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_PSE_AddSubstitute"
	Dim objAddSubstitute, objEdit, sMenu, bReturn 
	Set objAddSubstitute = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("AddSubstitute")
	Fn_PSE_AddSubstitute = False
	If objAddSubstitute.Exist(5) = False Then
		' getting menu operation command.
            sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "EditAddSubstitute")
		If sMenu = "False" Then		
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_PSE_AddSubstitute : Failed to get menu from XML.")
			Exit function
		End If

		'performing menu operation
		bReturn = Fn_Web_MenuOperation("Select", sMenu)
		If NOT(bReturn) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_PSE_AddSubstitute : Failed to perform menu operation [ " & sMenu & " ].")
			Exit function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_PSE_AddSubstitute : Successfully performed menu operation [ " & sMenu & " ].")
	End If

	If objAddSubstitute.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_PSE_AddSubstitute : Failed to open [ Add Substitute ] window.")
			Set objAddSubstitute = Nothing
			Exit function
	End If
	Select Case sAction
		Case "AddSubstitute"
			'setting item id
			If sItemId <> "" Then
				Set objEdit = objAddSubstitute.ChildItem( 1, 2,"WebEdit", 0)
				If TypeName(objEdit) <> "Nothing" Then
					objEdit.set sItemId
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_PSE_AddSubstitute : Successfully set [ Item ID ] = [ " & sItemId & " ] ")
				End If
			End If
			If bFindItem <> "" Then
				' click on find button
				 'for future use.
			End If
			Fn_PSE_AddSubstitute = true
	End Select
	If sBtnName <> "" Then
		Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebButton(sBtnName).Click 1,1,micLeftBtn
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_PSE_AddSubstitute : Executed successfully")
	Set objAddSubstitute = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_WebPSE_ListSubstituteOperations
'@@
'@@    Description				 :	Function perform operations on List substitute 
'@@
'@@    Parameters			   :	1. sAction : Action to be performed
'@@									2. sOpenDialogBy : Open dialog by method Menu / BOMNodeIcon
'@@									3. sPath : Nav tree node Path / BOM Table node path
'@@									5. sListSubstitute : List substitute Item name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Structure Manager perspective shuld be activated.
'@@
'@@    Examples					:	
'@@									Call Fn_WebPSE_ListSubstituteOperations("Verify", "Menu", "", "000034/A;1-sub") 
'@@									Call Fn_WebPSE_ListSubstituteOperations("Prefer", "", "000031/A;1-top (View):000034/A;1-sub", "000032/A;1-item") 
'@@									Call Fn_WebPSE_ListSubstituteOperations("Remove", "BOMNodeIcon", "000031/A;1-top (View):000032/A;1-item", "000034/A;1-sub") 
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			12-12-2011			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WebPSE_ListSubstituteOperations(sAction, sOpenDialogBy, sPath, sListSubstitute)
	GBL_FAILED_FUNCTION_NAME="Fn_WebPSE_ListSubstituteOperations"
	Dim objListSub, iCnt, bReturn, sMenu
	Dim objDialog, iRowCnt, objImg, arrListSubItems
	Dim iCount, bFlag
	bFlag = False
	Fn_WebPSE_ListSubstituteOperations = False

	Set objListSub = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ListSubstitute")
	' checking existence of dialog
	If objListSub.Exist(5) = False Then
			Select Case sOpenDialogBy
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				' opening by performing menu operation
				Case "Menu", ""
						' selecting BOM node.
						If sPath <> "" Then
							bReturn =  Fn_WebPSE_BOMTableOperations("Select",sPath,"","")
							If NOT(bReturn) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_ListSubstituteOperations : Failed to select node [ " & sPath & " ].")
								Exit function
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_ListSubstituteOperations : Successfully select node [ " & sPath & " ].")
						End If
						sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "EditRemovePreferSubstitute")
						If sMenu = "False" Then		
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_ListSubstituteOperations : Failed to get menu from XML.")
							Exit function
						End If
				
						'performing menu operation
						bReturn = Fn_Web_MenuOperation("Select", sMenu)
						If NOT(bReturn) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_ListSubstituteOperations : Failed to perform menu operation [ " & sMenu & " ].")
							Exit function
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_ListSubstituteOperations : Successfully performed menu operation [ " & sMenu & " ].")
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				Case "BOMNodeIcon"
						Set objDialog = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("BOMTable")
						iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sPath, "")
						If iRowCnt <> -1 Then
								Set objImg = objDialog.ChildItem(iRowCnt, 2, "Image", 1)
								If TypeName(objImg) <> "Nothing" Then
									If objImg.GetROProperty("file name") = "listalternates.png" Then
											objImg.Click 1,1, micLeftBtn
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_ListSubstituteOperations : Successfully clicked on [ Global Alternate ] icon of node [ " & sPath & " ].")
									End If
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_ListSubstituteOperations : Failed to find [ Global Alternate ] icon for node [ " & sPath & " ].")
									Exit function
								End IF
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_ListSubstituteOperations : Failed to find node [ " & sPath & " ].")
								Exit function
						End IF
			End Select
	End If

	' checking existence of dialog.
	If objListSub.Exist(5) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_ListSubstituteOperations : Failed to open [ Manage Global Alternates ] Window.")
		Set objListSub = Nothing
		Exit Function
	End If

	iRowCnt = cInt(objListSub.RowCount)

	Select Case sAction
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Remove", "Prefer"
						'selecting alternate
						For iCnt = 1 to iRowCnt 
							If trim(objListSub.GetCellData(iCnt, 1)) = sListSubstitute then
								objListSub.Object.rows(iCnt - 1).click 1,1,"LEFT"
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_ListSubstituteOperations : Successfully selected [ " & sListSubstitute  & " ].")
								bFlag = True
								Exit for
							End If
						Next
						wait(1)
						Fn_WebPSE_ListSubstituteOperations = bFlag
						If bFlag Then
							If sAction = "Prefer" Then
								' clicking on prefer button
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("SetPreferred").Click 1,1,micLeftBtn
							Else
								'clicking on remove button
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("Remove").Click 1,1,micLeftBtn
								'handling delete confirmation box
								If Browser("TeamcenterWeb").Dialog("Dialog").Exist(5) then
									Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").Click 1,1,micLeftBtn
								End If
							End If
						End If
						' closing dialog
						wait(1)
						Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebButton("Close").Click 1,1,micLeftBtn
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Verify", "VerifyCheckMark"
					arrListSubItems = split(sListSubstitute,"~")
					For iCount = 0 to UBound(arrListSubItems)
						bFlag = False
						For iCnt = 1 to iRowCnt 
							If trim(objListSub.GetCellData(iCnt, 1)) = arrListSubItems(iCount) then
								bFlag = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_ListSubstituteOperations : Successfully verified existence of [ " & arrListSubItems(iCount)  & " ].")
								If sAction = "VerifyCheckMark" Then
									bFlag = False
									Set objImg = objListSub.ChildItem(iCnt, 3, "Image", 0)
									If TypeName(objImg) <> "Nothing" Then
										If objImg.GetROProperty("file name") = "checkmark.png" Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_ListSubstituteOperations : Successfully verified check mark against [ " & arrListSubItems(iCount)  & " ].")
											bFlag = True
										End If
									End IF
								End If
								Exit for
							End If
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_ListSubstituteOperations : [ " & arrListSubItems(iCount)  & " ] does not exists.")
							Exit for
						End If
					Next
					' closing dialog
					Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebButton("Close").Click 1,1,micLeftBtn

					Fn_WebPSE_ListSubstituteOperations =  bFlag
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_WebPSE_ListSubstituteOperations : Invalid case [ " & sAction  & " ].")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		End Select

	If Fn_WebPSE_ListSubstituteOperations Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_WebPSE_ListSubstituteOperations : executed successfully with case [ " & sAction  & " ].")
	End If
	Set objListSub = Nothing
	Set objImg = Nothing
End Function
''*********************************************************	 Function to replace BOM line ***********************************************************************
'Function Name		:			    Fn_PSE_ReplaceBOMline

'Description		    :		 	Function to add substitute  

'Parameters			   :	 		1.sAction       : Action to perform
'									2. sItemId    : Item Id
'									3. bFindItem  : Boolean value to open find Item dialog
'									4. sBtnName   : Name of the button to be clicked. 

'Return Value		   : 			     TRUE \ FALSE

'Pre-requisite		    :		 	PSE window should be displayed.

'Examples			    :		Call Fn_PSE_ReplaceBOMline("ReplaceBOMline", "000032", "", "Apply")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh Watwe			14-12-2011			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PSE_ReplaceBOMline(sAction, sItemId, bFindItem, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_PSE_ReplaceBOMline"
	Dim objReplaceBOM, objEdit, sMenu, bReturn 
	Set objReplaceBOM = Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ReplaceBOMline")

	Fn_PSE_ReplaceBOMline = False
	If objReplaceBOM.Exist(5) = False Then
		' getting menu operation command.
            sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "EditReplaceBOMline")
		If sMenu = "False" Then		
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_PSE_AddSubstitute : Failed to get menu from XML.")
			Exit function
		End If

		'performing menu operation
		bReturn = Fn_Web_MenuOperation("Select", sMenu)
		If NOT(bReturn) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_PSE_ReplaceBOMline : Failed to perform menu operation [ " & sMenu & " ].")
			Exit function
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Fn_PSE_ReplaceBOMline : Successfully performed menu operation [ " & sMenu & " ].")
	End If

	If objReplaceBOM.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fn_PSE_ReplaceBOMline : Failed to open [ Replace BOM Line ] window.")
			Set objReplaceBOM = Nothing
			Exit function
	End If
	Select Case sAction
		Case "ReplaceBOMline"
			'setting item id
			If sItemId <> "" Then
				Set objEdit = objReplaceBOM.ChildItem( 1, 2,"WebEdit", 0)
				If TypeName(objEdit) <> "Nothing" Then
					objEdit.set sItemId
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_PSE_ReplaceBOMline : Successfully set [ Item ID ] = [ " & sItemId & " ] ")
				End If
			End If
			If bFindItem <> "" Then
				' click on find button
				 'for future use.
			End If
			Fn_PSE_ReplaceBOMline = true
	End Select
	If sBtnName <> "" Then
		Browser("TeamcenterStructureWeb").Page("TeamcenterWebStructure").WebTable("ButtonTable").WebButton(sBtnName).Click 1,1,micLeftBtn
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_PSE_ReplaceBOMline : Executed successfully")
	Set objReplaceBOM = Nothing
End Function