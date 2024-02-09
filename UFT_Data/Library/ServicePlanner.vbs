Option Explicit

iTimeOut = 240

''*********************************************************	Function List		***********************************************************************'
'	0. Fn_SISW_SP_GetObject()
'	1. Fn_SISW_SP_BOMTable_ColIndex()
'	2. Fn_SISW_SP_TableRowIndex()
'	3. Fn_SISW_SP_BOMTable_NodeOperation()
'	4. Fn_SISW_SP_TabOperations()
'	5. Fn_SISW_SP_NewServicePlanCreate()
'	6. Fn_SISW_SP_NewServicePartitionCreate()
'	7. Fn_SISW_SP_NewServiceRequirementCreate()
'	8. Fn_SISW_SP_NewWorkCardCreate()
'	9. Fn_SISW_SP_NewNoticeCreate()
'	10.Fn_SISW_SP_NewActivityCreate()
'	11 Fn_SISW_SP_NewSkillCreate()
'	12.Fn_SISW_SP_ResolvedFaultsOperation()
'	13.Fn_SISW_SP_SetupRequiresRelationOperation()
'	14.Fn_SISW_PSE_PlanDetailsOperation()
'	15.Fn_SISW_SP_BOMTable_ColumnOperation()
'   16.Fn_SISW_SP_SystemErrorHandle()
'   17 Fn_SISW_SP_NavTree_NodeOperation()
'   18 Fn_SISW_SP_AssignCharacteristics()
'   19 Fn_SISW_SP_CreateFaultCodeTypeOperations()
'   20 Fn_SISW_SP_FrequencyOperations()
'	21 Fn_SISW_SP_ActivitiesTableOperations()
'	22 Fn_SISW_SP_TimePanelOperations()
'   23 Fn_SISW_SP_ActivityAssignmentsOperations()
'	24 Fn_SISW_SP_PopulateAllocatedTimeOperations()
''*********************************************************	Function List		***********************************************************************'

'****************************************    Function to return required Object ***************************************
'
''Function Name		 	:	Fn_SISW_SP_GetObject
'
''Description		    :  	Function to get objects of Service Manager

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_SP_GetObject("NewServiceRequirement")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Veena		 			06-Nov-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_SP_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ServicePlanner.xml"
	Set Fn_SISW_SP_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'*********************************************************		Function to Get BOM Table Column Index into Service Planner		***********************************************************************
'Function Name		:					Fn_SISW_SP_BOMTable_ColIndex

'Description			 :		 		  This function is used to get the BOM Table Column Index.

'Parameters			   :	 			1.  sColName : Name of the Column
											
'Return Value		   : 				 Column index

'Pre-requisite			:		 		Service Planner Perspective should be Open.

'Examples				:				Fn_SISW_SP_BOMTable_ColIndex("BOM Line")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							29-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_BOMTable_ColIndex(sColName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_BOMTable_ColIndex"
	On Error Resume Next
	Dim iCols, iCounter, objTable, bFound, sColumn
	
	If Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").Exist(iTimeOut) Then
		iCols = Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").GetROProperty("cols")

		Set objTable = Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").Object
	
		For iCounter = 0 to iCols -1
			sColumn = objTable.getColumnName(iCounter)
        	If Trim(sColumn) = Trim(sColName) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: The Column Index for Column [" + sColName + "] is [" +iCounter +"] in Service Planner BOMTable")
				Fn_SISW_SP_BOMTable_ColIndex = iCounter
				Exit For
			End If
		Next

		If Cint(iCounter) = Cint(iCols) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: The Column [" + sColName + "] dose not exist in Service Planner BOMTable.")
			Fn_SISW_SP_BOMTable_ColIndex=-1
		End If

	   Set objTable = Nothing

	End If
End Function

'*********************************************************		Function to Get BOM Table Node Index into Service Planner		***********************************************************************

'Function Name		:					Fn_SISW_SP_TableRowIndex

'Description			 :		 		  This function is used to get the BOM Table Node Index.

'Parameters			   :	 			1. objTable - Table Object 
'										        2. sNodeName:Name of the Node to retrieve Index for.
											
'Return Value		   : 				 Node index

'Pre-requisite			:		 		Service Planner window should be displayed .

'Examples				:				 Call Fn_SISW_SP_TableRowIndex(objTable, "518611/A;1-Item_518611 (view):001270/A;1-ffff")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin					29-Oct-2012				1.0					
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_TableRowIndex(objTable, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_TableRowIndex"
	Dim aRowNode, iColIndex, sArr
	Dim iRows, iRowCnt, iInstance, iNodeCnt, iPathCnt
	Dim sNode, sNodePath, aPath, sPath, bFound
	Dim objComponent
	sPath = ""

	If Fn_UI_ObjectExist("Fn_SISW_SP_TableRowIndex", objTable) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_TableRowIndex ] Table does not exist.")	
		Fn_SISW_SP_TableRowIndex = -1
		Exit function
	End If
	iColIndex = 0
	bFound = False
	If sNodeName <> "" Then
		iRows = cInt(objTable.GetROProperty ("rows"))
		sArr = split(sNodeName , ":")
		iRowCnt = 0
		For iNodeCnt=0 to UBound(sArr)
				aRowNode = split(trim((sArr(iNodeCnt))),"@")
				If sPath = "" Then
							sPath =  trim(aRowNode(0))
				Else
							sPath = sPath &":"& trim(aRowNode(0))
				End If
		Next
		For iNodeCnt=0 to UBound(sArr)
			If iRowCnt = iRows  Then
				Exit for
			End If
			aRowNode = split(trim((sArr(iNodeCnt))),"@")
			iInstance = 0
			bFound = False
			do While iRowCnt < iRows
				If uBound(aRowNode) > 0 Then
							sNodePath = objTable.object.getValueAt(iRowCnt, iColIndex).toString()
							If trim(sNodePath) = trim(aRowNode(0)) then
	                                Set objComponent = ObjTable.object.getComponentForRow(iRowCnt)
									sNodePath = ""
									Do while NOT (objComponent is Nothing)
										If sNodePath = "" Then
											sNodePath = objComponent.getProperty("bl_indented_title")
										Else
											sNodePath =objComponent.getProperty("bl_indented_title") & ", " & StrNodePath
										End If
										'set objComponent = objComponent.parent()
										If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
											If IsObject(objComponent.parent()) = True Then
												set objComponent = objComponent.parent()
											Else
												Exit Do
											End If
										Else
											set objComponent = objComponent.parent()
											If  objComponent is Nothing Then
												Exit do
											End If
										End If
									Loop
'									sNodePath =objTable.Object.getPathForRow(iRowCnt).toString()
'									sNodePath = Right(sNodePath, (Len(sNodePath)-Instr(1, sNodePath, ",", 1)))					
'									sNodePath = Left(sNodePath, Len(sNodePath)-1)
									If instr(sNodePath, "@BOM::") > 0 Then
										sNodePath = trim(replace(sNodePath,"""",""))
										aPath = split(sNodePath,",")
										sNodePath = ""
										For icnt = 0 to uBound(aPath)
											aPath(iCnt) = Left(aPath(iCnt), instr(aPath(iCnt),"@")-1)
											If sNodePath = "" Then
												sNodePath = trim(aPath(iCnt))
											else
												sNodePath = sNodePath & ", " & trim(aPath(iCnt))
											End If
										Next
									End If

									sNodePath = trim(replace(sNodePath,", ",":"))
									If instr(sPath, sNodePath ) > 0 Then
										iInstance = iInstance +1
										If iInstance = cInt(aRowNode(1)) Then 
												If UBound(sArr) = iNodeCnt Then
														bFound = True
												End If
												Exit do
										End If
									End If
							End if
				Else
					If objTable.object.getPathForRow(iRowCnt).getLastPathComponent().getClass().toString() <> "class com.teamcenter.rac.treetable.HiddenSiblingNode" Then
						sNodePath = objTable.object.getValueAt(iRowCnt, iColIndex).toString()
					Else
						sNodePath = ""
					End If
					If trim(sNodePath) = trim(aRowNode(0)) then
                        Set objComponent = ObjTable.object.getComponentForRow(iRowCnt)
						sNodePath = ""
						Do while NOT (objComponent is Nothing)
							If sNodePath = "" Then
								sNodePath = objComponent.getProperty("bl_indented_title")
							Else
								sNodePath =objComponent.getProperty("bl_indented_title") & ", " & StrNodePath
							End If
							'set objComponent = objComponent.parent()
							
							If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
								If IsObject(objComponent.parent()) = True Then
									set objComponent = objComponent.parent()
								Else
									Exit Do
								End If
							Else
								set objComponent = objComponent.parent()
								If  objComponent is Nothing Then
									Exit do
								End If
							End If										
						Loop
'						sNodePath =objTable.Object.getPathForRow(iRowCnt).toString()
'						sNodePath = Right(sNodePath, (Len(sNodePath)-Instr(1, sNodePath, ",", 1)))					
'						sNodePath = Left(sNodePath, Len(sNodePath)-1)
						If instr(sNodePath, "@BOM::") > 0 Then
							sNodePath = trim(replace(sNodePath,"""",""))
							aPath = split(sNodePath,",")
							sNodePath = ""
							For icnt = 0 to uBound(aPath)
								aPath(iCnt) = Left(aPath(iCnt), instr(aPath(iCnt),"@")-1)
								If sNodePath = "" Then
									sNodePath = trim(aPath(iCnt))
								else
									sNodePath = sNodePath & ", " & trim(aPath(iCnt))
								End If
							Next
						End If
						sNodePath = trim(replace(sNodePath,", ",":"))
						If instr(sPath, sNodePath ) > 0 Then
								If UBound(sArr) = iNodeCnt Then
									bFound = True
								End If
								Exit do
						End if
					End if
				End If
				iRowCnt = iRowCnt + 1
			loop
		Next
	End If
	If bFound Then
				Fn_SISW_SP_TableRowIndex = iRowCnt
	Else
				Fn_SISW_SP_TableRowIndex = -1
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SP_TableRowIndex ] executed successfully.")
End Function

'******************************************************************Function to perform BOM Table Node operations************************************************************************************************************

'Function Name:				Fn_SISW_SP_BOMTable_NodeOperation

'Description: 				   1. This function is used to perform all Operation on BOM Table in Service Planner.

'Parameters:				  1. sAction: Action to be performed (Eg : Select/Exist/CellEdit etc.)
'						  			2. sNodeName: Fully qualified path of the BOM Table Node (Node delimiter as ':') (Multi-Nodes delimiter as ',')
'									3. sColName: Name of the BOM Table Column
'						  			4. sValue: BOM Table cell value for Edit or Verify actions
'						  			5. sPopupMenu: BOM Table Node context menu to be selected
'											  

'Return Value:				TRUE \ FALSE

'Pre-requisite:				Service Planner Window should be displayed with BOM Table loaded.

'Examples:				   
'					1 . Call Fn_SISW_SP_BOMTable_NodeOperation("Select", "000050/A;1-ffff (View):000052/A;1-Part2 (View):000053/A;1-ChildPart1 @2", "", "", "","Product")
'					2 . Call Fn_SISW_SP_BOMTable_NodeOperation("Deselect", "000050/A;1-ffff (View):000052/A;1-Part2 (View):000053/A;1-ChildPart1 @2", "", "", "","Product")
'					3 . Call Fn_SISW_SP_BOMTable_NodeOperation("MultiSelect", "000050/A;1-ffff (View):000052/A;1-Part2 (View):000053/A;1-ChildPart1~000050/A;1-ffff (View):000052/A;1-Part2 (View):000053/A;1-ChildPart1 @2", "", "", "","Product")
'					4 . Call Fn_SISW_SP_BOMTable_NodeOperation("SelectAll", "", "", "", "","")
'					5 . Call Fn_SISW_SP_BOMTable_NodeOperation("CellVerify", "000050/A;1-ffff (View):000052/A;1-Part2 (View):000053/A;1-ChildPart1 x 2", "Quantity", "3", "","Product")
'					6 . Call Fn_SISW_SP_BOMTable_NodeOperation("Exists", "000050/A;1-ffff (View):000052/A;1-Part2 (View):000053/A;1-ChildPart1 x 3", "", "", "","Product")
'					7 . Call Fn_SISW_SP_BOMTable_NodeOperation("Exists", "000050/A;1-ffff (View):000052/A;1-Part2 (View):000053/A;1-ChildPart3", "", "", "","Product")
'					8 . Call Fn_SISW_SP_BOMTable_NodeOperation("Expand", "000050/A;1-ffff (View)", "", "", "","Product")
'					9 . Call Fn_SISW_SP_BOMTable_NodeOperation("ExpandBelow", "000050/A;1-ffff (View)", "", "", "","Product")
'					10. Call Fn_SISW_SP_BOMTable_NodeOperation("Collapse", "000050/A;1-ffff (View)", "", "", "","Product")
'					11. Call Fn_SISW_SP_BOMTable_NodeOperation("CellEdit", "000050/A;1-ffff (View):000052/A;1-Part2 (View):000053/A;1-ChildPart1", "Quantity", "5", "","Product")
'					12. Call Fn_SISW_SP_BOMTable_NodeOperation("PopupSelect", "000050/A;1-ffff (View)", "", "", "Expand","Product")
'					13.Call Fn_SISW_SP_BOMTable_NodeOperation("Select", "000205/A;1-SP1 (View)", "", "", "","Service Plan")
'					14. Call Fn_SISW_SP_BOMTable_NodeOperation("Select", "SP1", "", "", "","Service Plan")
'					15. Call Fn_SISW_SP_BOMTable_NodeOperation("Select", "000201/A;1-Part1 (View)", "", "", "","Product:000201-Part1")
'					16. Call Fn_SISW_SP_BOMTable_NodeOperation("Select", "000205/A;1-SP1 (View)", "", "", "","Service Plan:000205-SP1:Base View")
'					17. Call Fn_SISW_SP_BOMTable_NodeOperation("Select", "SP1", "", "", "","Service Plan:000205-SP1:SP1")
'					18. Call Fn_SISW_SP_BOMTable_NodeOperation("MultiSelectPopupMenuSelect", "SP:dcxscxz:000095/A;1-SR (View)~SP:dcxscxz:000253/A;1-Req1 (View)", "Process Structure", "", "Expand","Service Plan:000093-SP:SP")
'History:
'										Developer Name			Date				Rev. No.			Changes Done												Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Sachin 				29-Oct-2012			1.0  									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_BOMTable_NodeOperation(sAction, sNodeName, sColName, sValue, sPopupMenu,sViewType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_BOMTable_NodeOperation"
	Dim objTable, objContextMenu, sArr
	Dim sNodes, bFound
	Dim strMenu, aMenu,  aNodeNames,iSubMenu
	Dim iCount, iRowCnt, iCounter, iColIndex, iRows, iCnt
	'Added code by Anumol: In 10.1 Column Name "Process Structure" change to "BOM Line"
	If sColName="Process Structure" Then
		sColName="BOM Line"
	End If
	If sViewType <> "" Then
		sArr = Split(sViewType,":")
		Call Fn_TabFolder_Operation("Select", sArr(1), "")
	
		If lcase(sArr(0)) = "product" Then
			For iCnt = 0 To 10
				Window("ServicePlannerWindow").JavaApplet("SPApplet").SetTOProperty "index", iCnt
				If Fn_SISW_UI_Object_Operations("Fn_SISW_SP_BOMTable_NodeOperation","Exist", Window("ServicePlannerWindow").JavaApplet("SPApplet"), 2) = True Then
					If Instr(1, sNodeName, Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").Object.getComponentForRow(0).getProperty("bl_indented_title") ) > 0 Then
						Exit For
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SP_BOMTable_NodeOperation] SPApplet does not exists.")
					Fn_SISW_SP_BOMTable_NodeOperation = False
					Exit Function
				End If
			Next
			'Window("ServicePlannerWindow").JavaApplet("SPApplet").SetTOProperty "Index","0"
		ElseIf sArr(0) = "Service Plan" Then
				If instr(sArr(2),"/") > 0 Then
					For iCnt = 0 To 10
						Window("ServicePlannerWindow").JavaApplet("SPApplet").SetTOProperty "index", iCnt
						If Fn_SISW_UI_Object_Operations("Fn_SISW_SP_BOMTable_NodeOperation","Exist", Window("ServicePlannerWindow").JavaApplet("SPApplet"), 2) = True Then
							If Instr(1, sNodeName, Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").Object.getComponentForRow(0).getProperty("bl_indented_title") ) > 0 Then
								Exit For
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SP_BOMTable_NodeOperation] SPApplet does not exists.")
							Fn_SISW_SP_BOMTable_NodeOperation = False
							Exit Function
						End If
					Next
					'Window("ServicePlannerWindow").JavaApplet("SPApplet").SetTOProperty "Index","1"
					Call Fn_SISW_SP_TabOperations("Activate", sArr(2))
				Else
					For iCnt = 0 To 10
						Window("ServicePlannerWindow").JavaApplet("SPApplet").SetTOProperty "index", iCnt
						If Fn_SISW_UI_Object_Operations("Fn_SISW_SP_BOMTable_NodeOperation","Exist", Window("ServicePlannerWindow").JavaApplet("SPApplet"), 2) = True Then
							If Instr(1, sNodeName, Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").Object.getComponentForRow(0).getProperty("bl_indented_title") ) > 0 Then
								Exit For
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SP_BOMTable_NodeOperation] SPApplet does not exists.")
							Fn_SISW_SP_BOMTable_NodeOperation = False
							Exit Function
						End If
					Next
					'Window("ServicePlannerWindow").JavaApplet("SPApplet").SetTOProperty "Index","2"
					Call Fn_SISW_SP_TabOperations("Activate", sArr(2))
				End If
		End If
	End If

	If Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").Exist = True Then
		Set objTable = Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable")
		Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaObject("BOMPanel").Click 1,1,"LEFT" 

	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SP_BOMTable_NodeOperation] BOM Table does not exists.")
		Set objTable = nothing
		Fn_SISW_SP_BOMTable_NodeOperation = False
		Exit function
	End if

	Select Case sAction
		Case "Select"
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName) 
				
				If iRowCnt <> -1 Then
                    objTable.Object.clearSelection  
					objTable.SelectRow iRowCnt 
					Fn_SISW_SP_BOMTable_NodeOperation = True
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False					
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

		Case "Deselect"
			Fn_SISW_SP_BOMTable_NodeOperation = False
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName) 
				
				If iRowCnt <> -1 Then
                    objTable.DeselectRow iRowCnt 
					Fn_SISW_SP_BOMTable_NodeOperation = True
				End If
			End If

		Case "MultiSelect"
			aNodeNames = split(sNodeName , "~")
			objTable.Object.clearSelection' Clear All Selection
			For iCounter = 0 to UBound(aNodeNames)
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,trim(aNodeNames(iCounter)))
				If iRowCnt <> -1 Then
					objTable.ExtendRow iRowCnt 
					Fn_SISW_SP_BOMTable_NodeOperation = True
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
					objTable.Object.clearSelection
					Exit for
				End If
			Next

		Case "SelectAll"
			objTable.Object.clearSelection' Clear All Selection
			iRows = cInt(objTable.GetROProperty ("rows"))
			For iCounter = 0 to iRows - 1
                objTable.ExtendRow iCounter 
			Next
			Fn_SISW_SP_BOMTable_NodeOperation = True

		Case "Exist", "Exists"
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName)

				If iRowCnt <> -1 Then
					Fn_SISW_SP_BOMTable_NodeOperation = True
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

		Case "Expand"
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName)
				If iRowCnt <> -1 Then
					objTable.SelectRow iRowCnt 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Selected  SP BOM Table Node [" + sNodeName + "]")
					If Fn_MenuOperation("WinMenuSelect", "View:Expand Options:Expand") = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Expanded SP BOM Table Node [" + sNodeName + "]")
						Fn_SISW_SP_BOMTable_NodeOperation = True
					Else							
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SP_BOMTable_NodeOperation] Failed to Expanded SP BOM Table Node [" + sNodeName + "]")
						Fn_SISW_SP_BOMTable_NodeOperation = False
					End If						
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

		Case "ExpandBelow"
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName) 

				If iRowCnt <> -1 Then
					objTable.SelectRow iRowCnt 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Selected  SP BOM Table Node [" + sNodeName + "]")
					If Fn_MenuOperation("WinMenuSelect", "View:Expand Options:Expand Below") = True Then
						If  Window("ServicePlannerWindow").JavaWindow("SPWindow").JavaDialog("ExpandBelow").Exist(5) then
							 Window("ServicePlannerWindow").JavaWindow("SPWindow").JavaDialog("ExpandBelow").JavaButton("Yes").Click
						End if
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Expanded Below SP BOM Table Node [" + sNodeName + "]")
						Fn_SISW_SP_BOMTable_NodeOperation = True
					Else							
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SP_BOMTable_NodeOperation] Failed to Expand Below SP BOM Table Node [" + sNodeName + "]")
						Fn_SISW_SP_BOMTable_NodeOperation = False
					End If						
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

		Case "Collapse"
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName)
				If iRowCnt <> -1 Then
					objTable.SelectRow iRowCnt 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Selected  SP BOM Table Node [" + sNodeName + "]")
					If Fn_MenuOperation("WinMenuSelect", "View:Collapse Below") = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Collapsed Below SP BOM Table Node [" + sNodeName + "]")
						Fn_SISW_SP_BOMTable_NodeOperation = True
					Else							
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SP_BOMTable_NodeOperation] Failed to Collapsed Below SP BOM Table Node [" + sNodeName + "]")
						Fn_SISW_SP_BOMTable_NodeOperation = False
					End If						
					
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

		Case "CellEdit"
			Fn_SISW_SP_BOMTable_NodeOperation = False
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName) 
				
				If iRowCnt <> -1 Then				
					objTable.SelectRow iRowCnt
					objTable.ClickCell iRowCnt,sColName, "LEFT" 
					wait 1
					If Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaEdit("BOMEdit_1").exist(3) Then
						Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaEdit("BOMEdit_1").Set sValue
						Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaEdit("BOMEdit_1").Activate
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Cell Edited of SP BOM Table Node [" + sNodeName + "]")
						Fn_SISW_SP_BOMTable_NodeOperation = True
					elseif Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaEdit("BOMComboEdit").exist(2) Then
						Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaEdit("BOMComboEdit").Set sValue
						Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaEdit("BOMComboEdit").Activate						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Cell Edited of SP BOM Table Node [" + sNodeName + "]")
						Fn_SISW_SP_BOMTable_NodeOperation = True
					End If
				End If
			End If

		Case "CellVerify"
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName) 
				If iRowCnt <> -1 Then
					bFound = Trim(cstr(objTable.GetCellData( iRowCnt,sColName)))

					If isNumeric(bFound) Then
						bFound=CStr(CInt(bFound))
					End If

					If isNumeric(sValue) Then
						sValue=CStr(CInt(sValue))
					End If

					If bFound = Trim(cstr(sValue)) Then
						Fn_SISW_SP_BOMTable_NodeOperation = True
					Else
						Fn_SISW_SP_BOMTable_NodeOperation = False
						If isNumeric(bFound) Then
							 bFound = Abs(bFound)
							 If cstr(bFound) = Trim(cstr(sValue)) Then
								 Fn_SISW_SP_BOMTable_NodeOperation = True
							End  If
						End If
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Cell verified of SP BOM Table Node [" + sNodeName + "]")
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

		Case "PopupSelect"


			objTable.Object.clearSelection  
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName)
				If iRowCnt <> -1 Then

					aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowCnt ,"BOM Line", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCnt ,sColName, "RIGHT","NONE"
					End If
					wait 3
					bFound = Fn_UI_JavaMenu_Select("Fn_MenuItem_select_Operation",JavaWindow("ServicePlanner"), sPopupMenu)
					If bFound=False Then
						Fn_SISW_SP_BOMTable_NodeOperation = False
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
					Fn_SISW_SP_BOMTable_NodeOperation = True
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If
		Case "MultiSelectPopupMenuSelect"
			If sNodeName <> "" Then
				aNodeNames = split(sNodeName , "~")
				objTable.Object.clearSelection
				For iCounter = 0 to UBound(aNodeNames)
					iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,trim(aNodeNames(iCounter)))
					If iRowCnt <> -1 Then
						objTable.ExtendRow iRowCnt 
						Fn_SISW_SP_BOMTable_NodeOperation = True
					Else
						Fn_SISW_SP_BOMTable_NodeOperation = False
						objTable.Object.clearSelection
						Exit for
					End If
				Next
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,trim(aNodeNames(UBound(aNodeNames))))
				If iRowCnt <> -1 Then
					aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowCnt ,"BOM Line", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCnt ,sColName, "RIGHT","NONE"
					End If
					wait 1
					Select Case Ubound(aMenu)
						Case "0"
							strMenu = Window("ServicePlannerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
							Window("ServicePlannerWindow").WinMenu("ContextMenu").Select strMenu
						Case "1"
							strMenu = Window("ServicePlannerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
							Window("ServicePlannerWindow").WinMenu("ContextMenu").Select strMenu
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [ Fn_SISW_SP_BOMTable_NodeOperation ] Context Menu Case NOT Exists for Supplied Menu [" + StrPopupMenu + "]")
							Fn_SISW_SP_BOMTable_NodeOperation = False
					End Select
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
					Fn_SISW_SP_BOMTable_NodeOperation = True
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

        Case "PopupMenuSelectWithCtrlKeyMultiSelect"
			If sNodeName <> "" Then
				aNodeNames = split(sNodeName , "~")
				objTable.Object.clearSelection
				For iCounter = 0 to UBound(aNodeNames)
					iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,trim(aNodeNames(iCounter)))
					If iRowCnt <> -1 Then
					    If iCounter = 0 Then
					    	objTable.ClickCell iRowCnt,"Item Description"
					    Else
							objTable.ClickCell iRowCnt,"Item Description","LEFT" ,"CONTROL"
					    End If
						Fn_SISW_SP_BOMTable_NodeOperation = True
					Else
						Fn_SISW_SP_BOMTable_NodeOperation = False
						objTable.Object.clearSelection
						Exit for
					End If
				Next
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,trim(aNodeNames(UBound(aNodeNames))))
				If iRowCnt <> -1 Then
					aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowCnt ,"BOM Line", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCnt ,sColName, "RIGHT","NONE"
					End If
					wait 1
					Select Case Ubound(aMenu)
						Case "0"
							strMenu = Window("ServicePlannerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
							Window("ServicePlannerWindow").WinMenu("ContextMenu").Select strMenu
						Case "1"
							strMenu = Window("ServicePlannerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
							Window("ServicePlannerWindow").WinMenu("ContextMenu").Select strMenu
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [ Fn_SISW_SP_BOMTable_NodeOperation ] Context Menu Case NOT Exists for Supplied Menu [" + StrPopupMenu + "]")
							Fn_SISW_SP_BOMTable_NodeOperation = False
					End Select
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SP_BOMTable_NodeOperation] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
					Fn_SISW_SP_BOMTable_NodeOperation = True
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

		Case "VerifyRowIsSelected"
			If sNodeName <> "" Then
				iRowCnt = Fn_SISW_SP_TableRowIndex(objTable,sNodeName)
				If iRowCnt <> -1 Then
					If objTable.Object.isRowSelected(iRowCnt)  Then
						Fn_SISW_SP_BOMTable_NodeOperation = True
					Else
						Fn_SISW_SP_BOMTable_NodeOperation = False
					End If 
				Else
					Fn_SISW_SP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_SISW_SP_BOMTable_NodeOperation = False
			End If

		Case Else
			Fn_SISW_SP_BOMTable_NodeOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_BOMTable_NodeOperation ] Invalid Action [ " & sAction & " ].")
			Set objTable = nothing
			exit function
			
	End Select
	If Fn_SISW_SP_BOMTable_NodeOperation <>FALSE then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SP_BOMTable_NodeOperation ] executed successfully with Action [ " & sAction & " ].")	
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to execute Function [ Fn_SISW_SP_BOMTable_NodeOperation ] with Action [ " & sAction & " ].")
	End if
	Set objTable = nothing
End Function

'*********************************************************		Function to Get BOM Table Column Index into Service Planner		***********************************************************************
'Function Name		:					Fn_SISW_SP_TabOperations

'Description			 :		 		  This function is used to Activate tab.

'Parameters			   :	 			1.  sAction : Action to Perform
'												  2. StrTabName :Tab to Activate.
											
'Return Value		   : 				 True/False

'Pre-requisite			:		 		Service Planner Perspective should be Open.

'Examples				:				Call Fn_SISW_SP_TabOperations("Activate", "SP1")
'												Call Fn_SISW_SP_TabOperations("Activate", "Base View")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							30-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_TabOperations(sAction, StrTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_TabOperations"
	Dim objTab,objWindow, bReturn, objTabWidget, iIndexCounter, iXposition
	Fn_SISW_SP_TabOperations = False
	Select Case sAction
		Case "Activate"
			Set objTab = JavaWindow("ServicePlanner").JavaTab("BaseView")
			If objTab.Exist(iTimeOut) Then
				bReturn = Fn_UI_JavaTab_Select("Fn_SISW_SP_TabOperations",JavaWindow("ServicePlanner"),"BaseView",StrTabName)
				If bReturn Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SP_TabOperations ] SPTab set to ["+StrTabName+"] for case ["+sAction+"].")
					Fn_SISW_SP_TabOperations = True
				End If
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: For Function [ Fn_SISW_SP_TabOperations ] case ["+sAction+"] is Invalid.")
			Fn_SISW_SP_TabOperations = True
	End Select

Set objTab = nothing
End Function


'*********************************************************		Function to Create Service Plan in Service Planner		***********************************************************************
'Function Name		:					Fn_SISW_SP_NewServicePlanCreate

'Description			 :		 		  This function is used to Create Service Plan.

'Parameters			   :	 			1.  sAction : Action to Perform
'												 2. sSPType : Service Plan Type
'												 3. sSPName : Service Plan Name
'												 4. sSPDescription : Service Plan Desc
'												 5. sSPID : Service Plan ID
'												 6. sSPRevID : Service Plan Rev ID
											
'Return Value		   : 				 ID-RevID / False

'Pre-requisite			:		 		Service Planner Perspective should be Open.

'Examples				:				Call Fn_SISW_SP_NewServicePlanCreate("Create", "Service Plan", "abc", "xyz","","")
' 
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							30-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_SP_NewServicePlanCreate(sAction, sSPType, sSPName, sSPDescription,sSPID,sSPRevID)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_NewServicePlanCreate"
	Dim objSerPlan,sItemId,sRevId
	Dim arrType, aSPType, sMRUPath, sCmplitListPath
	Set objSerPlan = JavaWindow("ServicePlanner").JavaWindow("NewServicePlan")
	Fn_SISW_SP_NewServicePlanCreate = False
			
	If objSerPlan.Exist(5) = False Then
		bReturn = Fn_MenuOperation("Select","File:New:Service Plan...")
		Call Fn_ReadyStatusSync(5)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > Service Plan... ] of Function Fn_SISW_SP_NewServicePlanCreate.")
			Set objSerPlan = nothing
			Exit Function
		End If
		If objSerPlan.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServicePlanCreate ] Failed to display Service Plan window.")
			Set objSerPlan = nothing
			Exit Function
		End if 
	End If
       	Select Case sAction
		Case "Create"
			If objSerPlan.JavaTree("ServicePlanType").Exist(3) Then
				If Trim(sSPType) <> "" Then
					aSPType = Split(sSPType,":",-1,1)
					sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
					sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServicePlanCreate",objSerPlan,"ServicePlanType", "Complete List" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewServicePlanCreate",objSerPlan,"ServicePlanType", "Complete List")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServicePlanCreate",objSerPlan,"ServicePlanType", "Most Recently Used" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewServicePlanCreate",objSerPlan,"ServicePlanType", "Most Recently Used")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServicePlanCreate",objSerPlan,"ServicePlanType", sMRUPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewServicePlanCreate", objSerPlan, "ServicePlanType",sMRUPath)
					elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServicePlanCreate",objSerPlan,"ServicePlanType", sCmplitListPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewServicePlanCreate", objSerPlan, "ServicePlanType",sCmplitListPath)
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServicePlanCreate ] Service Plan Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
						Set objSerPlan = nothing
						Fn_SISW_SP_NewServicePlanCreate = False
						Exit function
					end if
				End If
	
				'Call Fn_Button_Click("Fn_SISW_SP_NewServicePlanCreate", objSerPlan, "Next")
				objSerPlan.JavaButton("Next").Click
				'Call Fn_ReadyStatusSync(5)
				Wait (3)
			End If

			If sSPID <> "" Then
                 'Call Fn_Edit_Box("Fn_SISW_SP_NewServicePlanCreate",objSerPlan,"ID", sSPID)
				 objSerPlan.JavaEdit("ID").Set sSPID
			Else
				'Call Fn_Button_Click("Fn_SISW_SP_NewServicePlanCreate", objSerPlan, "AssignID")
				objSerPlan.JavaButton("AssignID").Click
				'Call Fn_ReadyStatusSync(5)
				Wait (3)
				sItemId = objSerPlan.JavaEdit("ID").GetROProperty("value")
			End If

			If sSPRevID <> "" Then
                 'Call Fn_Edit_Box("Fn_SISW_SP_NewServicePlanCreate",objSerPlan,"Revision", sSPRevID)
				 objSerPlan.JavaEdit("Revision").Set sSPRevID
			Else
				'Call Fn_Button_Click("Fn_SISW_SP_NewServicePlanCreate", objSerPlan, "AssignRevision")
				objSerPlan.JavaButton("AssignRevision").Click
				'Call Fn_ReadyStatusSync(5)
				Wait (3)
                sRevId = objSerPlan.JavaEdit("Revision").GetROProperty("value")
			End If

			If sSPName <> "" Then
				'Call Fn_Edit_Box("Fn_SISW_SP_NewServicePlanCreate", objSerPlan,"Name", sSPName )
				objSerPlan.JavaEdit("Name").Set sSPName
				If cInt(objSerPlan.JavaButton("Finish").getROProperty("enabled")) <> 1 Then
					objSerPlan.JavaEdit("Name").Activate
				End If
			End If

			If sSPDescription <> "" Then
				'Call Fn_Edit_Box("Fn_SISW_SP_NewServicePlanCreate", objSerPlan,"Description", sSPDescription )
				objSerPlan.JavaEdit("Description").Set sSPDescription
			End If
            
			'Call Fn_Button_Click("Fn_SISW_SP_NewServicePlanCreate", objSerPlan, "Finish")
			objSerPlan.JavaButton("Finish").Click
			Call Fn_ReadyStatusSync(2)

			If objSerPlan.exist(5) then
				'Call Fn_Button_Click("Fn_SISW_SP_NewServicePlanCreate", objSerPlan, "Cancel")
				objSerPlan.JavaButton("Cancel").Click
				Call Fn_ReadyStatusSync(1)
			End if

            Fn_SISW_SP_NewServicePlanCreate =  sItemId & "-" & sRevId

		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServicePlanCreate ] Invalid case [ " & sAction & " ].")
			Set objSerPlan = nothing
			Exit Function
	End Select
	If Fn_SISW_SP_NewServicePlanCreate <> false Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SP_NewServicePlanCreate ] Executed successfully with case [ " & sAction & " ].")
	End If
	Set objSerPlan = nothing
End Function

'*********************************************************		Function to Create Service Partition in Service Planner		***********************************************************************
'Function Name		:					Fn_SISW_SP_NewServicePartitionCreate

'Description			 :		 		  This function is used to Create Service Plan.

'Parameters			   :	 			1.  sAction : Action to Perform
'												 2. sSPType : Service Partition Type
'												 3. sSPName : Service Partition Name
'												 4. sSPDescription : Service Partition Desc
											
'Return Value		   : 				 True / False

'Pre-requisite			:		 		Service Planner Perspective should be Open.

'Examples				:				Call Fn_SISW_SP_NewServicePartitionCreate("Create", "Service Partition", "abc", "xyz")
' 
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							30-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_SP_NewServicePartitionCreate(sAction, sSPType, sSPName, sSPDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_NewServicePartitionCreate"
	Dim objSerPart, aSPType, sMRUPath, sCmplitListPath

	Set objSerPart = JavaWindow("ServicePlanner").JavaWindow("NewServicePartition")
	Fn_SISW_SP_NewServicePartitionCreate = False
			
	If objSerPart.Exist(5) = False Then
		bReturn = Fn_MenuOperation("Select","File:New:Service Partition...")
		Call Fn_ReadyStatusSync(5)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > Service Partition... ] of Function Fn_SISW_SP_NewServicePartitionCreate.")
			Set objSerPart = nothing
			Exit Function
		End If
		If objSerPart.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServicePartitionCreate ] Failed to display Service Plan window.")
			Set objSerPart = nothing
			Exit Function
		End if 
	End If
       	Select Case sAction
		Case "Create"
			If objSerPart.JavaTree("ServicePartitionType").Exist(3) Then
				If Trim(sSPType) <> "" Then
					aSPType = Split(sSPType,":",-1,1)
					sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
					sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServicePartitionCreate",objSerPart,"ServicePartitionType", "Complete List" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewServicePartitionCreate",objSerPart,"ServicePartitionType", "Complete List")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServicePartitionCreate",objSerPart,"ServicePartitionType", "Most Recently Used" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewServicePartitionCreate",objSerPart,"ServicePartitionType", "Most Recently Used")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServicePartitionCreate",objSerPart,"ServicePartitionType", sMRUPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewServicePartitionCreate", objSerPart, "ServicePartitionType",sMRUPath)
					elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServicePartitionCreate",objSerPart,"ServicePartitionType", sCmplitListPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewServicePartitionCreate", objSerPart, "ServicePartitionType",sCmplitListPath)
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServicePartitionCreate ] Service Plan Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
						Set objSerPart = nothing
						Fn_SISW_SP_NewServicePartitionCreate = False
						Exit function
					end if
				End If
	
				Call Fn_Button_Click("Fn_SISW_SP_NewServicePartitionCreate", objSerPart, "Next")
				Call Fn_ReadyStatusSync(5)
			End If

			If sSPName <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServicePartitionCreate", objSerPart,"Name", sSPName )
				If cInt(objSerPart.JavaButton("Finish").getROProperty("enabled")) <> 1 Then
					objSerPart.JavaEdit("Name").Activate
				End If
			End If

			If sSPDescription <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServicePartitionCreate", objSerPart,"Description", sSPDescription )
			End If
            
			Call Fn_Button_Click("Fn_SISW_SP_NewServicePartitionCreate", objSerPart, "Finish")

			If objSerPart.exist(5) then
				Call Fn_Button_Click("Fn_SISW_SP_NewServicePartitionCreate", objSerPart, "Cancel")
			End if

            Fn_SISW_SP_NewServicePartitionCreate =  True

		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServicePartitionCreate ] Invalid case [ " & sAction & " ].")
			Set objSerPart = nothing
			Exit Function
	End Select
	If Fn_SISW_SP_NewServicePartitionCreate =True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SP_NewServicePartitionCreate ] Executed successfully with case [ " & sAction & " ].")
	End If
	Set objSerPart = nothing
End Function


'*********************************************************		Function to Create Service Requirement in Service Requirementner		***********************************************************************
'Function Name		:					Fn_SISW_SP_NewServiceRequirementCreate

'Description			 :		 		  This function is used to Create Service Requirement.

'Parameters			   :	 			1.  sAction : Action to Perform
'												 2. sSRType : Service Requirement Type
'												 3. sSRID : Service Requirement ID
'												 4. sSRRevID : Service Requirement Rev ID
'												 5. sSRName : Service Requirement Name
'												 6. sSRDescription : Service Requirement Desc
'												 7. sSRCategory: Service Requirement Category 
'												 8. sSRReqType: Service Requirement Type [on Second window]
'												 9. sSRUpgrade: Service Requirement Upgrade [Radio]
											
'Return Value		   : 				 ID-RevID / False

'Pre-requisite			:		 		Service Requirementner Perspective should be Open.

'Examples				:				Call Fn_SISW_SP_NewServiceRequirementCreate("Create", "Service Requirement", "","","dasdsa", "adsdsa","","Repair","True")
'											    Call Fn_SISW_SP_NewServiceRequirementCreate("VerifyError", "Service Requirement", "000062","","dasdsa", "adsdsa","The instance cannot be saved because it contains at least one attribute that violates a unique attribute rule.","Repair","True")
' 
'History:
'	Developer Name			Date			Rev. No.		Reviewer			Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sachin					30-Oct-2012		1.0
'	Sandeep N					31-jan-2013		1.1			Anumol			Modified function to select Requirement type as in 10.1 object is changed
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_SP_NewServiceRequirementCreate(sAction, sSRType, sSRID,sSRRevID,sSRName, sSRDescription,sSRCategory,sSRReqType,sSRUpgrade)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_NewServiceRequirementCreate"
	Dim objSerReq,sItemId,sRevId
	Dim arrType, aSPType, sMRUPath, sCmplitListPath
	Dim bFlag,iCounter
	Dim objTable,objChild

	Set objSerReq = JavaWindow("ServicePlanner").JavaWindow("NewServiceRequirement")
	Fn_SISW_SP_NewServiceRequirementCreate = False
			
	If objSerReq.Exist(5) = False Then
		bReturn = Fn_MenuOperation("Select","File:New:Service Requirement...")
		Call Fn_ReadyStatusSync(5)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > Service Requirement... ] of Function Fn_SISW_SP_NewServiceRequirementCreate.")
			Set objSerReq = nothing
			Exit Function
		End If
		If objSerReq.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServiceRequirementCreate ] Failed to display Service Requirement window.")
			Set objSerReq = nothing
			Exit Function
		End if 
	End If
	objSerReq.Maximize
	Select Case sAction
		Case "Create"
			If objSerReq.JavaTree("ServiceRequirementType").Exist(3) Then
				If Trim(sSRType) <> "" Then
					aSPType = Split(sSRType,":",-1,1)
					sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
					sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", "Complete List" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", "Complete List")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", "Most Recently Used" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", "Most Recently Used")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", sMRUPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "ServiceRequirementType",sMRUPath)
					elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", sCmplitListPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "ServiceRequirementType",sCmplitListPath)
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServiceRequirementCreate ] Service Requirement Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
						Set objSerReq = nothing
						Fn_SISW_SP_NewServiceRequirementCreate = False
						Exit function
					end if
				End If
	
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "Next")
	
				Call Fn_ReadyStatusSync(5)
			End If

			If sSRID <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ID", sSRID)
			Else
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "AssignID")
				Call Fn_ReadyStatusSync(5)
				sItemId = objSerReq.JavaEdit("ID").GetROProperty("value")
			End If

			If sSRRevID <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"Revision", sSRRevID)
			Else
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "AssignRevision")
				Call Fn_ReadyStatusSync(5)
				sRevId = objSerReq.JavaEdit("Revision").GetROProperty("value")
			End If

			If sSRName <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq,"Name", sSRName )
				If cInt(objSerReq.JavaButton("Finish").getROProperty("enabled")) <> 1 Then
					objSerReq.JavaEdit("Name").Activate
				End If
			End If

			If sSRDescription <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq,"Description", sSRDescription )
			End If

			If sSRCategory <> "" Then
				' do nothing
				' this Option is disabled Currently
			End If

			If sSRReqType <> "" Then
				bFlag=False
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "RequirementType")
				wait 2
				Set objTable=Description.Create()
				objTable("Class Name").value="JavaTable"
				Set objChild=objSerReq.ChildObjects(objTable)
				For iCounter=0 to objChild(0).GetROProperty("rows")-1
					If objChild(0).GetCellData(iCounter,0)=sSRReqType Then
						objChild(0).ActivateRow iCounter
						bFlag=True
						Exit for
					End If
				Next
				Set objTable=Nothing
				Set objChild=Nothing
				If bFlag=False Then
					Exit function
				End If
'				Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq,"ReqType", sSRReqType )
			End If
			wait 2
			If lcase(Cstr(sSRUpgrade)) = "true" Then
				objSerReq.JavaRadioButton("True").Set "ON"
			Else
				objSerReq.JavaRadioButton("False").Set "ON"
			End If
			objSerReq.Restore
			Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "Finish")

			If objSerReq.exist(5) then
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "Cancel")
			End if

            Fn_SISW_SP_NewServiceRequirementCreate =  sItemId & "-" & sRevId

		Case "VerifyError"
		
			If Trim(sSRType) <> "" Then
				If objSerReq.JavaTree("ServiceRequirementType").Exist(3) = True Then	'Condtion added to check existence of object selection tree
						aSPType = Split(sSRType,":",-1,1)
						sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
						sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
						If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", "Complete List" , "", "") <> -1 then
							Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", "Complete List")
						end if
						
						If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", "Most Recently Used" , "", "") <> -1 then
							Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", "Most Recently Used")
						end if
						
						If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", sMRUPath , "", "") <> -1 then
							Call Fn_JavaTree_Select("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "ServiceRequirementType",sMRUPath)
						elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ServiceRequirementType", sCmplitListPath , "", "") <> -1 then
							Call Fn_JavaTree_Select("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "ServiceRequirementType",sCmplitListPath)
						else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServiceRequirementCreate ] Service Requirement Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
							Set objSerReq = nothing
							Fn_SISW_SP_NewServiceRequirementCreate = False
							Exit function
						end if
						Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "Next")

						Call Fn_ReadyStatusSync(5)
				End If
			End If



			If sSRID <> "" Then
                 Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"ID", sSRID)
			Else
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "AssignID")
				Call Fn_ReadyStatusSync(5)
				sItemId = objSerReq.JavaEdit("ID").GetROProperty("value")
			End If

			If sSRRevID <> "" Then
                 Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate",objSerReq,"Revision", sSRRevID)
			Else
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "AssignRevision")
				Call Fn_ReadyStatusSync(5)
                sRevId = objSerReq.JavaEdit("Revision").GetROProperty("value")
			End If

			If sSRName <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq,"Name", sSRName )
				If cInt(objSerReq.JavaButton("Finish").getROProperty("enabled")) <> 1 Then
					objSerReq.JavaEdit("Name").Activate
				End If
			End If

			If sSRDescription <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq,"Description", sSRDescription )
			End If

            If sSRReqType <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq,"ReqType", sSRReqType )
			End If

			If lcase(Cstr(sSRUpgrade)) = "true" Then
				 JavaWindow("ServicePlanner").JavaWindow("NewServiceRequirement").JavaRadioButton("True").Set "ON"
			Else
				JavaWindow("ServicePlanner").JavaWindow("NewServiceRequirement").JavaRadioButton("False").Set "ON"
			End If
			objSerReq.Restore
			Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "Finish")

			If objSerReq.JavaWindow("Error").Exist(5) Then
				objSerReq.JavaWindow("Error").JavaStaticText("Static").SetTOProperty "label",sSRCategory

				If trim(objSerReq.JavaWindow("Error").JavaStaticText("Static").GetROProperty("attached text")) = trim(sSRCategory) Then
					Fn_SISW_SP_NewServiceRequirementCreate =  True
				Else
					Fn_SISW_SP_NewServiceRequirementCreate =  False
				End If
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq.JavaWindow("Error"), "OK")
			End If
			

			If objSerReq.exist(5) then
				Call Fn_Button_Click("Fn_SISW_SP_NewServiceRequirementCreate", objSerReq, "Cancel")
			End if

            Fn_SISW_SP_NewServiceRequirementCreate =  True
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewServiceRequirementCreate ] Invalid case [ " & sAction & " ].")
			Set objSerReq = nothing
			Exit Function
	End Select
	
	If Fn_SISW_SP_NewServiceRequirementCreate <> false Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SP_NewServiceRequirementCreate ] Executed successfully with case [ " & sAction & " ].")
	End If
	Set objSerReq = nothing
End Function

'*********************************************************		Function to Create Work Card in Work Card	***********************************************************************
'Function Name		:					Fn_SISW_SP_NewWorkCardCreate

'Description			 :		 		  This function is used to Create Work Card.

'Parameters			   :	 			1.  sAction : Action to Perform
'												 2. sWCType : Work Card Type
'												 3. sWCID : Work Card ID
'												 4. sWCRevID : Work Card Rev ID
'												 5. sWCName : Work Card Name
'												 6. sWCDescription : Work Card Desc
'												 7. sWCExeType : Work Card Activity Execution type
'												 8. sWCLaborCost : Work Card Labor Cost
'												 9. sMaterialCost: Work Card Material Cost
'												 10. sNerrative : Work Card  Nerrative
											
'Return Value		   : 				 ID-RevID / False

'Pre-requisite			:		 		Work Cardner Perspective should be Open.

'Examples				:				Call Fn_SISW_SP_NewWorkCardCreate("Create", "Work Card", "","","test", "sdd","Perform","23","4566","dsdsd")
' 
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							30-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_SP_NewWorkCardCreate(sAction, sWCType, sWCID,sWCRevID,sWCName, sWCDescription,sWCExeType,sWCLaborCost,sMaterialCost,sNerrative)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_NewWorkCardCreate"

   Dim objWC,sItemId,sRevId
	Dim arrType, aSPType, sMRUPath, sCmplitListPath
	Set objWC = JavaWindow("ServicePlanner").JavaWindow("NewWorkCard")
	Fn_SISW_SP_NewWorkCardCreate = False
			
	If objWC.Exist(5) = False Then
		bReturn = Fn_MenuOperation("Select","File:New:Work Card...")
		Call Fn_ReadyStatusSync(5)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > Work Card... ] of Function Fn_SISW_SP_NewWorkCardCreate.")
			Set objWC = nothing
			Exit Function
		End If
		If objWC.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewWorkCardCreate ] Failed to display Work Card window.")
			Set objWC = nothing
			Exit Function
		End if 
	End If
       	Select Case sAction
		Case "Create"
			If objWC.JavaTree("WorkCardType").Exist(3) Then
				If Trim(sWCType) <> "" Then
					aSPType = Split(sWCType,":",-1,1)
					sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
					sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewWorkCardCreate",objWC,"WorkCardType", "Complete List" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewWorkCardCreate",objWC,"WorkCardType", "Complete List")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewWorkCardCreate",objWC,"WorkCardType", "Most Recently Used" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewWorkCardCreate",objWC,"WorkCardType", "Most Recently Used")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewWorkCardCreate",objWC,"WorkCardType", sMRUPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewWorkCardCreate", objWC, "WorkCardType",sMRUPath)
					elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewWorkCardCreate",objWC,"WorkCardType", sCmplitListPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewWorkCardCreate", objWC, "WorkCardType",sCmplitListPath)
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewWorkCardCreate ] Work Card Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
						Set objWC = nothing
						Fn_SISW_SP_NewWorkCardCreate = False
						Exit function
					end if
				End If
	
				Call Fn_Button_Click("Fn_SISW_SP_NewWorkCardCreate", objWC, "Next")
	
				Call Fn_ReadyStatusSync(5)
			End If

			If sWCID <> "" Then
                 Call Fn_Edit_Box("Fn_SISW_SP_NewWorkCardCreate",objWC,"ID", sWCID)
			Else
				Call Fn_Button_Click("Fn_SISW_SP_NewWorkCardCreate", objWC, "AssignID")
				Call Fn_ReadyStatusSync(5)
				sItemId = objWC.JavaEdit("ID").GetROProperty("value")
			End If

			If sWCRevID <> "" Then
                 Call Fn_Edit_Box("Fn_SISW_SP_NewWorkCardCreate",objWC,"Revision", sWCRevID)
			Else
				Call Fn_Button_Click("Fn_SISW_SP_NewWorkCardCreate", objWC, "AssignRevision")
				Call Fn_ReadyStatusSync(5)
                sRevId = objWC.JavaEdit("Revision").GetROProperty("value")
			End If

			If sWCName <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewWorkCardCreate", objWC,"Name", sWCName )
				If cInt(objWC.JavaButton("Finish").getROProperty("enabled")) <> 1 Then
					objWC.JavaEdit("Name").Activate
				End If
			End If

			If sWCExeType <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewWorkCardCreate", objWC,"ActivityExecutionType", sWCExeType )
			End If

			If sWCDescription <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewWorkCardCreate", objWC,"Description", sWCDescription )
			End If

			If sWCLaborCost <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewWorkCardCreate", objWC,"LaborCost", sWCLaborCost )
			End If

			If sMaterialCost <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewWorkCardCreate", objWC,"MaterialCost", sMaterialCost )
			End If

			If sNerrative <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewWorkCardCreate", objWC,"Nerrative", sNerrative)
			End If

			Call Fn_Button_Click("Fn_SISW_SP_NewWorkCardCreate", objWC, "Finish")

			If objWC.exist(5) then
				Call Fn_Button_Click("Fn_SISW_SP_NewWorkCardCreate", objWC, "Cancel")
			End if

            Fn_SISW_SP_NewWorkCardCreate =  sItemId & "-" & sRevId

		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewWorkCardCreate ] Invalid case [ " & sAction & " ].")
			Set objWC = nothing
			Exit Function
	End Select
	If Fn_SISW_SP_NewWorkCardCreate <> false Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SP_NewWorkCardCreate ] Executed successfully with case [ " & sAction & " ].")
	End If
	Set objWC = nothing
End Function

'*********************************************************		Function to Create Notice in Notice	***********************************************************************
'Function Name		:					Fn_SISW_SP_NewNoticeCreate

'Description			 :		 		  This function is used to Create Notice.

'Parameters			   :	 			1.  sAction : Action to Perform
'												 2. sNoticeType : Notice Type
'												 3. sNoticeName : Notice Name
'												 4. sNoticeDescription : Notice Desc
'												 5. sOptionNoticeType : Notice Type to select on 2nd page
											
'Return Value		   : 				 ID-RevID / False

'Pre-requisite			:		 		Noticener Perspective should be Open.

'Examples				:				Call Fn_SISW_SP_NewNoticeCreate("Create", "Notice", "Notice1", "Notice1dsd","Note")
' 
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							30-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_SP_NewNoticeCreate(sAction, sNoticeType, sNoticeName, sNoticeDescription,sOptionNoticeType)
	
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_NewNoticeCreate"
   	Dim objNotice, sItemId, sRevId
	Dim arrType, aSPType, sMRUPath, sCmplitListPath
	
	If Instr(JavaWindow("DefaultWindow").GetROProperty("title"), "Service Scheduler") > 0  Then
		Set objNotice = JavaWindow("ServiceScheduler").JavaWindow("NewNotice")
	Else
		Set objNotice = JavaWindow("ServicePlanner").JavaWindow("NewNotice")
	End If

	Fn_SISW_SP_NewNoticeCreate = False
			
	If objNotice.Exist(5) = False Then
		bReturn = Fn_MenuOperation("Select","File:New:Notice...")
		Call Fn_ReadyStatusSync(5)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > Notice... ] of Function Fn_SISW_SP_NewNoticeCreate.")
			Set objNotice = nothing
			Exit Function
		End If
		If objNotice.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewNoticeCreate ] Failed to display Notice window.")
			Set objNotice = nothing
			Exit Function
		End if 
	End If
       	Select Case sAction
		Case "Create"
			If objNotice.JavaTree("NoticeType").Exist(3) Then
				If Trim(sNoticeType) <> "" Then
					aSPType = Split(sNoticeType,":",-1,1)
					sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
					sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewNoticeCreate",objNotice,"NoticeType", "Complete List" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewNoticeCreate",objNotice,"NoticeType", "Complete List")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewNoticeCreate",objNotice,"NoticeType", "Most Recently Used" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewNoticeCreate",objNotice,"NoticeType", "Most Recently Used")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewNoticeCreate",objNotice,"NoticeType", sMRUPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewNoticeCreate", objNotice, "NoticeType",sMRUPath)
					elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewNoticeCreate",objNotice,"NoticeType", sCmplitListPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewNoticeCreate", objNotice, "NoticeType",sCmplitListPath)
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewNoticeCreate ] Notice Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
						Set objNotice = nothing
						Fn_SISW_SP_NewNoticeCreate = False
						Exit function
					end if
				End If
	
				Call Fn_Button_Click("Fn_SISW_SP_NewNoticeCreate", objNotice, "Next")
				Call Fn_ReadyStatusSync(5)
			End If

			If sNoticeName <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewNoticeCreate", objNotice,"Name", sNoticeName )
			End If

			If sNoticeDescription <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewNoticeCreate", objNotice,"Description", sNoticeDescription )
			End If

			If sOptionNoticeType <> "" Then
'				Call Fn_Edit_Box("Fn_SISW_SP_NewNoticeCreate", objNotice,"NoticeType", sOptionNoticeType )
				Call Fn_Button_Click("Fn_SISW_SP_NewNoticeCreate", objNotice, "NoticeType")
				wait 2
				objNotice.JavaTree("Tree").Activate sOptionNoticeType
'				Call Fn_Edit_Box("Fn_SISW_SP_NewNoticeCreate", objNotice,"NoticeType", sOptionNoticeType )
			End If

			Call Fn_Button_Click("Fn_SISW_SP_NewNoticeCreate", objNotice, "Finish")

			If objNotice.exist(5) then
				Call Fn_Button_Click("Fn_SISW_SP_NewNoticeCreate", objNotice, "Cancel")
			End if

            Fn_SISW_SP_NewNoticeCreate =  True

		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewNoticeCreate ] Invalid case [ " & sAction & " ].")
			Set objNotice = nothing
			Exit Function
	End Select
	If Fn_SISW_SP_NewNoticeCreate = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SP_NewNoticeCreate ] Executed successfully with case [ " & sAction & " ].")
	End If
	Set objNotice = nothing
End Function

'*********************************************************		Function to Create Activity in Activity	***********************************************************************
'Function Name		:					Fn_SISW_SP_NewActivityCreate

'Description			 :		 		  This function is used to Create Activity.

'Parameters			   :	 			1.  sAction : Action to Perform
'												 2. sActivityType : Activity Type
'												 3. sActivityName : Activity Name
'												 4. sActivityExecutionType : Activity Ececution Type
'												 5. sActivityStartTime : Activity Start Time
'												 6. sActivityDuration : Activity Duration
'												 7. sActivityDescription : Activity  Desc
											
'Return Value		   : 				 True / False

'Pre-requisite			:		 		Service Planner Perspective should be Open.

'Examples				:				Call Fn_SISW_SP_NewActivityCreate("Create", "MEActivity", "Act", "Perform","0.0","0.0","desc")
' 
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							30-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_SP_NewActivityCreate(sAction, sActivityType, sActivityName, sActivityExecutionType,sActivityStartTime,sActivityDuration,sActivityDescription)
  GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_NewActivityCreate"
   Dim objActivity,sItemId,sRevId
	Dim arrType, aSPType, sMRUPath, sCmplitListPath
	Set objActivity = JavaWindow("ServicePlanner").JavaWindow("NewActivity")
	Fn_SISW_SP_NewActivityCreate = False
			
	If objActivity.Exist(5) = False Then
		bReturn = Fn_MenuOperation("Select","File:New:Activity...")
		Call Fn_ReadyStatusSync(5)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > Activity... ] of Function Fn_SISW_SP_NewActivityCreate.")
			Set objActivity = nothing
			Exit Function
		End If
		If objActivity.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewActivityCreate ] Failed to display Activity window.")
			Set objActivity = nothing
			Exit Function
		End if 
	End If
       	Select Case sAction
		Case "Create"
			If objActivity.JavaTree("ActivityType").Exist(3) Then
				If Trim(sActivityType) <> "" Then
					aSPType = Split(sActivityType,":",-1,1)
					sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
					sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewActivityCreate",objActivity,"ActivityType", "Complete List" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewActivityCreate",objActivity,"ActivityType", "Complete List")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewActivityCreate",objActivity,"ActivityType", "Most Recently Used" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewActivityCreate",objActivity,"ActivityType", "Most Recently Used")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewActivityCreate",objActivity,"ActivityType", sMRUPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewActivityCreate", objActivity, "ActivityType",sMRUPath)
					elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewActivityCreate",objActivity,"ActivityType", sCmplitListPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewActivityCreate", objActivity, "ActivityType",sCmplitListPath)
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewActivityCreate ] Activity Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
						Set objActivity = nothing
						Fn_SISW_SP_NewActivityCreate = False
						Exit function
					end if
				End If
	
				Call Fn_Button_Click("Fn_SISW_SP_NewActivityCreate", objActivity, "Next")
				Call Fn_ReadyStatusSync(5)
			End If

			If sActivityName <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewActivityCreate", objActivity,"Name", sActivityName )
			End If

			If sActivityExecutionType <> "" Then
				'Call Fn_Edit_Box("Fn_SISW_SP_NewActivityCreate", objActivity,"ActivityType", sActivityExecutionType )
				Call Fn_Button_Click("Fn_SISW_SP_NewActivityCreate", objActivity, "ActivityInformation")
				bReturn = Fn_JavaTree_Select("Fn_SISW_SP_NewActivityCreate", objActivity.JavaWindow("Shell"), "ActivityTypeTree",sActivityExecutionType)
				objActivity.JavaEdit("Name").Click 50, 5, "LEFT"
				If bReturn = False Then
					Exit Function
				End If
				Wait 2
			End If

			If sActivityStartTime <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewActivityCreate", objActivity,"StartTime", sActivityStartTime )
			End If

			If sActivityDuration <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewActivityCreate", objActivity,"Duration", sActivityDuration )
			End If
			
			If sActivityDescription <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewActivityCreate", objActivity,"Description", sActivityDescription )
			End If

			Call Fn_Button_Click("Fn_SISW_SP_NewActivityCreate", objActivity, "Finish")

			If objActivity.exist(5) then
				Call Fn_Button_Click("Fn_SISW_SP_NewActivityCreate", objActivity, "Cancel")
			End if

            Fn_SISW_SP_NewActivityCreate =  True

		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewActivityCreate ] Invalid case [ " & sAction & " ].")
			Set objActivity = nothing
			Exit Function
	End Select
	If Fn_SISW_SP_NewActivityCreate = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SP_NewActivityCreate ] Executed successfully with case [ " & sAction & " ].")
	End If
	Set objActivity = nothing
End Function


'*********************************************************		Function to Create New Skill in New Skillner		***********************************************************************
'Function Name		:					Fn_SISW_SP_NewSkillCreate

'Description			 :		 		  This function is used to Create New Skill.

'Parameters			   :	 			1.  sAction : Action to Perform
'												 2. sSKType : New Skill Type
'												 3. sSKID : New Skill ID
'												 4. sSKRevID : New Skill Rev ID
'												 5. sSKName : New Skill Name
'												 6. sSKDescription : New Skill Desc
'												 7. sSKDiscipline: New Skill Descipline 
											
'Return Value		   : 				 ID-RevID / False

'Pre-requisite			:		 		New Service Planner Perspective should be Open.

'Examples				:				Call Fn_SISW_SP_NewSkillCreate("Create", "Skill", "","","Skill1", "desc","AutoDisp1")
' 
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							30-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SISW_SP_NewSkillCreate(sAction, sSKType, sSKID,sSKRevID,sSKName, sSKDescription,sSKDiscipline)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_NewSkillCreate"
	Dim objSkill,sItemId,SKRevId
	Dim arrType, aSPType, sMRUPath, sCmplitListPath
	Set objSkill = JavaWindow("ServicePlanner").JavaWindow("NewSkill")
	Fn_SISW_SP_NewSkillCreate = False
			
	If objSkill.Exist(5) = False Then
		bReturn = Fn_MenuOperation("Select","File:New:Skill...")
		Call Fn_ReadyStatusSync(5)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > New Skill... ] of Function Fn_SISW_SP_NewSkillCreate.")
			Set objSkill = nothing
			Exit Function
		End If
		If objSkill.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewSkillCreate ] Failed to display New Skill window.")
			Set objSkill = nothing
			Exit Function
		End if 
	End If
       	Select Case sAction
		Case "Create"
			If objSkill.JavaTree("SkillType").Exist(3) Then
				If Trim(sSKType) <> "" Then
					aSPType = Split(sSKType,":",-1,1)
					sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
					sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewSkillCreate",objSkill,"SkillType", "Complete List" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewSkillCreate",objSkill,"SkillType", "Complete List")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewSkillCreate",objSkill,"SkillType", "Most Recently Used" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SP_NewSkillCreate",objSkill,"SkillType", "Most Recently Used")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewSkillCreate",objSkill,"SkillType", sMRUPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewSkillCreate", objSkill, "SkillType",sMRUPath)
					elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SP_NewSkillCreate",objSkill,"SkillType", sCmplitListPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SP_NewSkillCreate", objSkill, "SkillType",sCmplitListPath)
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewSkillCreate ] New Skill Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
						Set objSkill = nothing
						Fn_SISW_SP_NewSkillCreate = False
						Exit function
					end if
				End If
	
				Call Fn_Button_Click("Fn_SISW_SP_NewSkillCreate", objSkill, "Next")
				Call Fn_ReadyStatusSync(5)
			End If

			If sSKID <> "" Then
                 Call Fn_Edit_Box("Fn_SISW_SP_NewSkillCreate",objSkill,"ID", sSKID)
			Else
				Call Fn_Button_Click("Fn_SISW_SP_NewSkillCreate", objSkill, "AssignID")
				Call Fn_ReadyStatusSync(5)
				sItemId = objSkill.JavaEdit("ID").GetROProperty("value")
			End If

			If sSKRevID <> "" Then
                 Call Fn_Edit_Box("Fn_SISW_SP_NewSkillCreate",objSkill,"Revision", sSKRevID)
			Else
				Call Fn_Button_Click("Fn_SISW_SP_NewSkillCreate", objSkill, "AssignRevision")
				Call Fn_ReadyStatusSync(5)
                SKRevId = objSkill.JavaEdit("Revision").GetROProperty("value")
			End If

			If sSKName <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewSkillCreate", objSkill,"Name", sSKName )
				If cInt(objSkill.JavaButton("Finish").getROProperty("enabled")) <> 1 Then
					objSkill.JavaEdit("Name").Activate
				End If
			End If

			If sSKDescription <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewSkillCreate", objSkill,"Description", sSKDescription)
			End If

			If sSKDiscipline <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SP_NewSkillCreate", objSkill,"Discipline", sSKDiscipline)
				wait 2
				Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
			End If

			Call Fn_Button_Click("Fn_SISW_SP_NewSkillCreate", objSkill, "Finish")

			If objSkill.exist(5) then
				Call Fn_Button_Click("Fn_SISW_SP_NewSkillCreate", objSkill, "Cancel")
			End if

            Fn_SISW_SP_NewSkillCreate =  sItemId & "-" & SKRevId

		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SP_NewSkillCreate ] Invalid case [ " & sAction & " ].")
			Set objSkill = nothing
			Exit Function
	End Select
	If Fn_SISW_SP_NewSkillCreate <> false Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SP_NewSkillCreate ] Executed successfully with case [ " & sAction & " ].")
	End If
	Set objSkill = nothing
End Function
'*********************************************************		Function to Verify Resolved Faults into Service Planner		***********************************************************************
'Function Name		:					Fn_SISW_SP_ResolvedFaultsOperation

'Description			 :		 		  This function is used to perform Operations on Resolved Faults into Service Planner.

'Parameters			   :	 			1.  sAction :Action to Perform

											
'Return Value		   : 				 Column index

'Pre-requisite			:		 		Resolved Faults Dialog should be Opened.

'Examples				:				

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin							31-Oct-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_ResolvedFaultsOperation(sAction,sFaults)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_ResolvedFaultsOperation"
Dim objResolvFault, aFaults, iCount, bFlag
Set objResolvFault = JavaWindow("ServicePlanner").JavaWindow("CreateResolvesRelation")

If objResolvFault.Exist(5) Then
	Select Case sAction
	Case "Verify"
		If sFaults <> "" Then
			If Instr(sFaults,"~") > 0 Then
				aFaults = Split(sFaults,"~")
			Else
				sFaults = sFaults +"~1"
				aFaults = Split(sFaults,"~")
			End If
		End If
		If aFaults(1) <> "1" Then
			iCount = objResolvFault.JavaList("FaultCodes").GetROProperty("items count")
				For iCnt=0 to iCount-1
					If Trim(lcase(objResolvFault.JavaList("FaultCodes").GetItem(iCnt))) = Trim(lcase(aFaults(iRowData))) then
						bFlag = True
					End If
				Next
				If iCnt = iCount and bFlag = False Then
					Fn_SISW_SP_ResolvedFaultsOperation = False
				Else
					Fn_SISW_SP_ResolvedFaultsOperation = True
				End If
		Else
			If Trim(lcase(objResolvFault.JavaList("FaultCodes").GetItem(0))) = Trim(lcase(aFaults(0))) then
				Fn_SISW_SP_ResolvedFaultsOperation = True
			End If
		End If
		Call Fn_Button_Click("Fn_SISW_SP_ResolvedFaultsOperation", objResolvFault, "Cancel")

	Case "Select"
			iCount = objResolvFault.JavaList("FaultCodes").GetROProperty("items count")
				For iCnt=0 to iCount-1
					If Trim(lcase(objResolvFault.JavaList("FaultCodes").GetItem(iCnt))) = Trim(lcase(sFaults)) then
						objResolvFault.JavaList("FaultCodes").Select(iCnt)
						bFlag = True
						Exit for
					End If
				Next
				If iCnt = iCount and bFlag = False Then
					Fn_SISW_SP_ResolvedFaultsOperation = False
				Else
					Fn_SISW_SP_ResolvedFaultsOperation = True
				End If
		wait 2
	'---  temperoary solution to set focus on ok button
		objResolvFault.JavaButton("OK").highlight
		Call Fn_Button_Click("Fn_SISW_SP_ResolvedFaultsOperation", objResolvFault, "OK")
	Case Else
		Fn_SISW_SP_ResolvedFaultsOpration = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_ResolvedFaultsOperation ] Invalid Action [ " & sAction & " ].")
		Set objResolvFault = nothing
		Exit Function
	End Select
Else
	Fn_SISW_SP_ResolvedFaultsOpration = False
	Set objResolvFault = nothing
	Exit Function
End If

End Function
'*********************************************************		Function to Perform Operations on Setup Requires Relation and Setup Satisfies Relation into Service Planner.	***********************************************************************
'Function Name		:					Fn_SISW_SP_SetupRequiresRelationOperation

'Description			 :		 		  This function is used to perform Operations on Setup Requires Relation and Setup Satisfies Relation into Service Planner.

'Parameters			   :	 			1.  sAction :Action to Perform
'													 2. sServiceReq: Service Requirement
'													3. sRequiresServiceReq: RequiresServiceReq
'													4. sButton
											
'Return Value		   : 				 True\False

'Pre-requisite			:		 		Setup Requires Relation OR  Setup Satisfies Relation Dialog should be Open.

'Examples				:				'msgbox Fn_SISW_SP_SetupRequiresRelationOperation("Verify","000152-REQ1","000155-REQ2","Cancel")
													'msgbox Fn_SISW_SP_SetupRequiresRelationOperation("Swap","000195-Ra1","000196-J1","OK")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Veena						05-Nov-2012		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_SetupRequiresRelationOperation(sAction,sServiceReq,sRequiresServiceReq,sButton)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_SetupRequiresRelationOperation"
Dim objSetupRequires, iCounter, bFlag
Set objSetupRequires = JavaWindow("ServicePlanner").JavaWindow("SetupRequiresRelation")
Fn_SISW_SP_SetupRequiresRelationOperation = False
If objSetupRequires.Exist(5) Then
	Select Case sAction
	Case "Swap"
		If sServiceReq <> "" and sRequiresServiceReq <> ""Then
			Call Fn_Button_Click("Fn_SISW_SP_SetupRequiresRelationOperation", objSetupRequires, "Swap")
			Call Fn_Button_Click("Fn_SISW_SP_SetupRequiresRelationOperation", objSetupRequires, sButton)
			Fn_SISW_SP_SetupRequiresRelationOperation = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [ Fn_SISW_SP_SetupRequiresRelationOperation ] Performed Action [ " & sAction & " ].")
		End If

	Case "Verify"
		bFlag = False
			If sRequiresServiceReq <> "" Then
					If objSetupRequires.JavaObject("RequiresServicereqquirementLink").Object.getText() = sRequiresServiceReq Then
						bFlag = True
					End If
			End If
			If bFlag=False Then
				Fn_SISW_SP_SetupRequiresRelationOperation = False
				Exit Function
			End If

			If sServiceReq <> "" Then
				If objSetupRequires.JavaObject("ServiceReqLink").Object.getText() = sServiceReq Then
						bFlag = True
				End If
				If bFlag=False Then
					Fn_SISW_SP_SetupRequiresRelationOperation = False
					Exit Function
				End If
			End If
		Call Fn_Button_Click("Fn_SISW_SP_SetupRequiresRelationOperation", objSetupRequires, sButton)

	Case Else
		Fn_SISW_SP_SetupRequiresRelationOperation = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_SetupRequiresRelationOperation ] Invalid Action [ " & sAction & " ].")
		Set objSetupRequires = nothing
		Exit Function
	End Select

	If Fn_SISW_SP_SetupRequiresRelationOperation <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SP_SetupRequiresRelationOperation ] Successfully executed with case [ " & sAction & " ].")
	End If
	Fn_SISW_SP_SetupRequiresRelationOperation = True
	Set objSetupRequires = nothing
	Exit Function
End If
End Function
'********************************************************* To Perform Operations in plan details dialog.*************************
'Function Name		:	Fn_SISW_PSE_PlanDetailsOperation

'Description		:	This function is used to Perform Operations on Plan details dialog

'Parameters			:	01. sAction         - Action to perform
'								02. dicPlanDetails   - Dictionary object 

'Return Value		: 	TRUE \ FALSE

'Pre-requisite		:	Plan Details should be Opened

'Examples			:	
'								Dim dicPlanDetails
'								Set dicPlanDetails = CreateObject( "Scripting.Dictionary")
'								With dicPlanDetails  
'									.Add "Static","Satisfies Service Requirements"
'									.Add "Value","000044-SR66"
'								End with
'								Call Fn_SISW_PSE_PlanDetailsOperation("Verify", dicPlanDetails)
'--------------------------------------------------------------------------------------------------------------------------------
'								Dim dicPlanDetails
'								Set dicPlanDetails = CreateObject( "Scripting.Dictionary")
'								With dicPlanDetails  
'									.Add "Static","Satisfies Service Requirements"
'									.Add "Value","000044-SR66"	
'									.Add "MenuPath","Send To:Classification"
'								End with
'								Call Fn_SISW_PSE_PlanDetailsOperation("PopupMenuSelect", dicPlanDetails)
'--------------------------------------------------------------------------------------------------------------------------------
'History:
'		Developer Name			Date			Version		Reviewer		Changes
'--------------------------------------------------------------------------------------------------------------------------------
'		Sachin					06/11/2012		1.0		
'--------------------------------------------------------------------------------------------------------------------------------
'		Koustubh Watwe			30/11/2012		1.0							Added case PopupMenuSelect
'--------------------------------------------------------------------------------------------------------------------------------
'		Shweta Rathod			13/06/2016		1.0							Added case Verify_ext,AddNew,paste
'--------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_PSE_PlanDetailsOperation(sAction, dicPlanDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PSE_PlanDetailsOperation"
	Dim objServicePlanner,objSearch,oCurrentNode,sArr,sArrSelect,StrName,iCount,sText,iRowData,aValue,iCtr,sArrStatic
	Dim sTabList,iPath
	sTabList = "Required Resulting Information~Satisfies Service Requirements~Satisfied By Service Requirements~Requires Service Requirements~Required By Service Requirements~Fault~Part Applicabilty~References~Specifications"
	sTabList = sTabList & "~Measurement Characteristics~Notices~Upgrades~Properties~Related Datasets"
	Set objServicePlanner = JavaWindow("ServicePlanner")
	sArrStatic = Split(sTabList,"~")


	For i=0 to uBound(sArrStatic)
		If sArrStatic(i) <> dicPlanDetails("Static") Then
			objServicePlanner.JavaStaticText("StaticText").SetTOProperty "label",sArrStatic(i)
			If objServicePlanner.JavaStaticText("StaticText").Exist(1) Then
				Call Fn_SISW_UI_Twistie_Operations("", "Collapse", JavaWindow("ServicePlanner"), "Twistie", sArrStatic(i),"StaticText")
			End If
		Else
			objServicePlanner.JavaStaticText("StaticText").SetTOProperty "label",sArrStatic(i)
			If objServicePlanner.JavaStaticText("StaticText").Exist(1) Then
				Call Fn_SISW_UI_Twistie_Operations("", "Expand", JavaWindow("ServicePlanner"), "Twistie", sArrStatic(i),"StaticText")
			End If
		End If
	Next

	Select Case lcase(sAction)
		Case "verify"
			Call Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", "Plan Details", "")
            objServicePlanner.JavaStaticText("StaticText").SetTOProperty "label",dicPlanDetails("Static")
			wait 2
			objServicePlanner.JavaTable("Table").SetTOProperty "attached text",dicPlanDetails("Static")

			Call Fn_Button_Click ("Fn_SISW_PSE_PlanDetailsOperation",objServicePlanner,"Table")
			Call Fn_ReadyStatusSync(5)
			
			If objServicePlanner.JavaTable("Table").Exist(5) = false Then
				Call Fn_Button_Click ("Fn_SISW_PSE_PlanDetailsOperation",objServicePlanner,"Table")
				Call Fn_ReadyStatusSync(5)
			End If

			If dicPlanDetails("Value") <> "" Then
                iRows = objServicePlanner.JavaTable("Table").GetROProperty("rows")
				aValue = split(dicPlanDetails("Value"), "~",-1,1)
				iCount = Ubound(aValue)
				iCtr = 0
				For iRowData=0 to iCount
					For iCnt=0 to iRows-1
						If Trim(lcase(objServicePlanner.JavaTable("Table").Object.getItem(iCnt).getData().toString())) = Trim(lcase(aValue(iRowData))) then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aValue(iRowData) &" Sucessfully found in Table")
							iCtr = iCtr + 1
							Exit For
						End If
					Next
				Next
				If iCtr=iCount+1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completion of function Fn_SISW_PSE_PlanDetailsOperation")
					Fn_SISW_PSE_PlanDetailsOperation = TRUE
					Call Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", "Plan Details", "")
					Set objServicePlanner = nothing 
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" failed to Verify Item exist in Table.")
					Fn_SISW_PSE_PlanDetailsOperation = False
					Call Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", "Plan Details", "")
					Set objServicePlanner = Nothing
					Exit Function
				End If
			End If
		'--------------------------------------------------------------------------------------------------------------------------------
		Case "verify_ext"
			Set objSearch = JavaWindow("ServicePlanner").JavaTree("RequiredResultingInformation")
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_PSE_PlanDetailsOperation", objSearch, dicPlanDetails("Value") , sDelimiter, "@")
			If iPath=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + dicPlanDetails("Value") + "] Not exist in Required Resulting Information Tree")
				Fn_SISW_PSE_PlanDetailsOperation = False
			Else
				sArr = split(replace(iPath,"#",""),":")
				Fn_SISW_PSE_PlanDetailsOperation = true
				Set oCurrentNode = objSearch.Object
				For iCnt = 0 to UBound(sArr) -1
					Set oCurrentNode = oCurrentNode.GetItem(sArr(iCnt))
					If cBool(oCurrentNode.getExpanded()) = False Then
						Fn_SISW_PSE_PlanDetailsOperation = false
						Exit for
					End If
				Next
				If Fn_SISW_PSE_PlanDetailsOperation Then
					call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + dicPlanDetails("Value") + "] Exist in Required Resulting Information Tree")
				End If
			End If
			Set oCurrentNode = Nothing
			Set objSearch = nothing
		'--------------------------------------------------------------------------------------------------------------------------------
		Case "addnew","paste"
			objServicePlanner.JavaButton("Button").SetTOProperty "label",dicPlanDetails("ButtonName")
			Call Fn_Button_Click ("Fn_SISW_PSE_PlanDetailsOperation",objServicePlanner,"Button")
			Fn_SISW_PSE_PlanDetailsOperation = true
			Call Fn_ReadyStatusSync(5)
			If lcase(sAction) = "addnew" Then
				Set dicAddNew = CreateObject( "Scripting.Dictionary" )
				dicAddNew("Object Type") = dicPlanDetails("ItemType")
				dicAddNew("Name") = dicPlanDetails("ItemName")
				bRet = Fn_SISW_CreateNewBusinessObject("CreateNewBusinessObject",dicAddNew)
				If bRet = false Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_PSE_PlanDetailsOperation -> Fn_SISW_CreateNewBusinessObject ] Failed to create business object.")
					Fn_SISW_PSE_PlanDetailsOperation = false
					Exit Function
				End If
			End If	
			
		'--------------------------------------------------------------------------------------------------------------------------------
		Case "popupmenuselect"
            objServicePlanner.JavaStaticText("StaticText").SetTOProperty "label",dicPlanDetails("Static")
			wait 2
			objServicePlanner.JavaTable("Table").SetTOProperty "attached text",dicPlanDetails("Static")

			Call Fn_Button_Click ("Fn_SISW_PSE_PlanDetailsOperation",objServicePlanner,"Table")
			Call Fn_ReadyStatusSync(5)
			
			If objServicePlanner.JavaTable("Table").Exist(5) = false Then
				Call Fn_Button_Click ("Fn_SISW_PSE_PlanDetailsOperation",objServicePlanner,"Table")
				Call Fn_ReadyStatusSync(5)
			End If

			If dicPlanDetails("Value") <> "" Then
				iRows = objServicePlanner.JavaTable("Table").GetROProperty("rows")
				aValue = split(dicPlanDetails("Value"), "~",-1,1)
				iCount = Ubound(aValue)
				For iCnt=0 to iRows-1
					If Trim(lcase(objServicePlanner.JavaTable("Table").Object.getItem(iCnt).getData().toString())) = Trim(lcase(aValue(iRowData))) then
						objServicePlanner.JavaTable("Table").ClickCell 0,0,"RIGHT"
						wait 1
						Fn_SISW_PSE_PlanDetailsOperation = Fn_UI_JavaMenu_Select("Fn_SISW_PSE_PlanDetailsOperation",objServicePlanner,dicPlanDetails("MenuPath"))
						Exit For
					End If
				Next
				If Fn_SISW_PSE_PlanDetailsOperation = TRUE Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completion of function Fn_SISW_PSE_PlanDetailsOperation")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" failed to Verify Item exist in Table.")
					Fn_SISW_PSE_PlanDetailsOperation = False
				End If
			End If
	End Select
	Set objServicePlanner = Nothing
End Function
'*********************************************************		Function to Get BOM Table Column Index into Service Planner		***********************************************************************

'Function Name		:		Fn_SISW_SP_BOMTable_ColumnOperation

'Description		:			This function is used to Perform Operations on BOMTable Columns

'Parameters			:		01. StrAction         - Action to perform
'										02. StrColName	- Column Name   
'										03. sIndex				- Applet Index

'Return Value		: 	TRUE \ FALSE

'Pre-requisite		:	Service Planner  window should be displayed .

'Examples			:		Call Fn_SISW_SP_BOMTable_ColumnOperation("Add","All Notes","2")
'History:
'		Developer Name			Date			Version		Reviewer		Changes
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Sachin					18/11/2012		1.0		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_BOMTable_ColumnOperation(StrAction,StrColName,sIndex)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_BOMTable_ColumnOperation"
		Dim PopUpMenu,iColIndex,ObjTable,ArrCol,iIndex,sColToAdd, objList, intCol, objChangeColumnDialog,IntCounter,StrName,bFlag,IntCols
		Window("ServicePlannerWindow").JavaApplet("SPApplet").SetToProperty "Index",sIndex
		wait 2
		Set ObjTable = Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").Object
		
		Fn_SISW_SP_BOMTable_ColumnOperation = False
		Select Case StrAction
				Case "Add"
						ArrCol = Split(StrColName,":",-1,1)
                        For iIndex = 0 To Ubound(ArrCol)
								'Check that Column is present in the BOMTable.
								iColIndex =  Fn_SISW_SP_BOMTable_ColIndex(ArrCol(iIndex))
								If iColIndex = -1 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Warning: Column does not  exist in the Application.Need to Add Column ["& ArrCol(iIndex) &"]." )
										sColToAdd = sColToAdd +":"+ArrCol(iIndex)
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Column ["& ArrCol(iIndex) &"] exists in the Application" )
										Fn_SISW_SP_BOMTable_ColumnOperation =TRUE
								End if
						Next
						If sColToAdd <>""  Then
								sColToAdd = Mid(sColToAdd, 2,Len(sColToAdd))
								ArrCol = Split(sColToAdd,":",-1,1)

								Set objChangeColumnDialog = Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaDialog("Change Columns")
								If NOT objChangeColumnDialog.Exist( 1)  Then
										Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaTable("BOMTable").SelectColumnHeader "#1","RIGHT"       	
										Window("ServicePlannerWindow").JavaApplet("SPApplet").JavaMenu("label:=Insert column\(s\) ...").Select 										       
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
										Set objList = objChangeColumnDialog.JavaList("ListAvailableCols").Object
								End If
                				Call Fn_ReadyStatusSync(2)

								For iIndex = 0 To Ubound(ArrCol)							
										intCol = objChangeColumnDialog.JavaList("ListAvailableCols").GetItemIndex(ArrCol(iIndex))
										objList.ensureIndexIsVisible intCol
										objChangeColumnDialog.JavaList("ListAvailableCols").ExtendSelect ArrCol(iIndex)
								Next
								objChangeColumnDialog.JavaButton("Add").Click

								objChangeColumnDialog.JavaButton("Apply").Click

								objChangeColumnDialog.JavaButton("Cancel").Click

								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Added  Column  ["& sColToAdd &"] in BOMTable")									
								Fn_SISW_SP_BOMTable_ColumnOperation = TRUE					
						End If

		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid case  ["& StrAction &"].")
				Fn_SISW_SP_BOMTable_ColumnOperation = False
		End Select

		 Window("ServicePlannerWindow").JavaApplet("SPApplet").SetToProperty "Index","0"
		wait 2
		Set objList = Nothing
		Set ObjTable = Nothing
		Set objChangeColumnDialog = nothing
 End Function

'*********************************************************		Function to Verify Error msg .	***********************************************************************
'Function Name		:					Fn_SISW_SP_SystemErrorHandle

'Description			 :		 		  This function is used to perform Operations on Setup Requires Relation and Setup Satisfies Relation into Service Planner.

'Parameters			   :	 			1. ErrMsg
'													
											
'Return Value		   : 				 True\False

'Pre-requisite			:		 		Requires Service Requirment Relation Creation Failed dialog should be open

'Examples				:				'msgbox Fn_SISW_SP_SystemErrorHandle("The instance cannot be saved because it contains at least one attribute that violates a unique attribute rule.")
													

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ashok kakade		20-Nov-2012				1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_SystemErrorHandle(ErrMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_SystemErrorHandle"
	GBL_EXPECTED_MESSAGE=ErrMsg
	Dim objErrWindow, sMsg
	Set objErrWindow = JavaWindow("ServicePlanner").JavaWindow("RequiresServiceRequirement")
	Fn_SISW_SP_SystemErrorHandle = False
	If objErrWindow.Exist(5) Then
	If objErrWindow.JavaStaticText("The instance cannot be").exist(1) Then
		sMsg = objErrWindow.JavaStaticText("The instance cannot be").GetROProperty("label")
    ElseIf objErrWindow.JavaEdit("ErrorMsg").Exist(1) Then
         sMsg = objErrWindow.JavaEdit("ErrorMsg").GetROProperty("value")
	End If
		Wait 2
		

			If ErrMsg = sMsg Then
				Call Fn_Button_Click("Fn_SISW_SP_SetupRequiresRelationOperation", objErrWindow, "OK")			
				Fn_SISW_SP_SystemErrorHandle = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [ Fn_SISW_SP_SystemErrorHandle ] Error Message [ " & sMsg & " ] is Valid.")
			Else
				GBL_ACTUAL_MESSAGE=sMsg
				Call Fn_Button_Click("Fn_SISW_SP_SetupRequiresRelationOperation", objErrWindow, "OK")			
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_SystemErrorHandle ] Error Message [ " & sMsg & " ] is not Valid.")
				Exit Function
			End If
	End If
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'*********************************************************		Function to action perform on NavTree of  Service Manager ***********************************************************************
'Function Name		:				Fn_SISW_SP_NavTree_NodeOperation

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'																	2. Node multi-select
'																	3. Node Expand
'																	4. Node Collapse
'																	5. Node Popup menu select
'																	6. Node double-click
'																	7. Node Deselect
'																	8. Node Exist
'																	9. Node SelectRange

'Parameters			   :	 			1. StrAction: Action to be performed
'													2. StrNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
'												   3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Service Planner module window should be displayed

'Examples				:				   1) Call Fn_ASB_NavTree_NodeOperation("PopupMenuSelect","Home:Newstuff","Copy")
'													   2) Call Fn_ASB_NavTree_NodeOperation("Exist","Home:Newstuff","")		

'History					 :		
'	Developer Name				Date						Rev. No.			Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vrushali     Wani				28-Nov-2012				0001																	Rupali   Palhad
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_SP_NavTree_NodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_NavTree_NodeOperation"
	Dim intCount, aMenuList,iCounter, iRows
	Dim objJavaWindowMyTc, objJavaTreeNav,ArrNodeName
	Dim sPath, sEle,arr
	Dim arrStrNode,intNodeCount,oCurrentNode,sReturn

	Fn_SISW_SP_NavTree_NodeOperation = FALSE
	Set objJavaWindowMyTc = JavaWindow("ServicePlanner")
	Set objJavaTreeNav = JavaWindow("ServicePlanner").JavaTree("NavTree")

	Select Case StrAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
					sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SP_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					If sPath <> False Then
						objJavaTreeNav.Select sPath
						Fn_SISW_SP_NavTree_NodeOperation = True
					End If
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------

		Case "Deselect"
					sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SP_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					If sPath <> False Then
						objJavaTreeNav.Deselect sPath
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deselected Node [" + StrNodeName + "] of NavTree")
						Fn_SISW_SP_NavTree_NodeOperation = True
					End If
		'----------------------------------------------------------------------- For selecting multiple node at a time -------------------------------------------------------------------------
		Case "Multiselect"
					Set objJavaTreeNav = JavaWindow("ServicePlanner").JavaTree("NavTree")
					sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SP_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					If sPath <> False Then
						Call Fn_UI_JavaTree_ExtendSelect("Fn_SISW_SP_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree", sPath)
						Fn_SISW_SP_NavTree_NodeOperation = TRUE
					End If
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
					sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SP_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					If sPath <> False Then
						objJavaTreeNav.Expand sPath
						Fn_SISW_SP_NavTree_NodeOperation = True
					End If
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
			sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SP_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
			If sPath <> False Then
				objJavaTreeNav.Collapse sPath
				Fn_SISW_SP_NavTree_NodeOperation = True
			End If
		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuSelect"
			sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SP_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
			If sPath <> False Then
					'Select node
                    Call Fn_JavaTree_Select("Fn_SISW_SP_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree",sPath )
					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_SP_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree",sPath)
					wait 3
					Fn_SISW_SP_NavTree_NodeOperation = Fn_UI_JavaMenu_Select("Fn_SISW_SP_NavTree_NodeOperation",objJavaWindowMyTc,StrMenu)
			End If		
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "DoubleClick"
			sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SP_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
			If sPath <> False Then
				JavaWindow("ServicePlanner").JavaTree("NavTree").Activate sPath 
				Fn_SISW_SP_NavTree_NodeOperation = TRUE
			End If
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "Exist"
				sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SP_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
				If sPath <> False Then
					Fn_SISW_SP_NavTree_NodeOperation = TRUE
				End If
         '----------------------------------------------------------------------- For  Select Range of  Nav tree -------------------------------------------------------------------------
		Case "SelectRange"
			ReDim ArrNodeName(2)
					ArrNodeName = Split(StrNodeName,"|")
					JavaWindow("ServicePlanner").JavaTree("NavTree").SelectRange ArrNodeName(0),ArrNodeName(1)
					If err.number < 0 Then
						Fn_SISW_SP_NavTree_NodeOperation = False
					else
						Fn_SISW_SP_NavTree_NodeOperation = True
					End If

'- - - - - - - - - - - -  Retruns All Childs of any given Node in the tree in form of an array - - - - - - - - - - - - - - -
				Case "GetChildrenList"
						sReturn=""
						If Fn_SISW_SP_NavTree_NodeOperation("Expand",StrNodeName,"")=True Then
							arrStrNode = Split (StrNodeName, ":")
							If UBound(arrStrNode)=0 Then
								Set oCurrentNode = JavaWindow("ServicePlanner").JavaTree("NavTree").Object.getItem(0)
								intNodeCount = oCurrentNode.getItemCount()
								For iCount=0 To intNodeCount-1
									If iCount=0 Then
										sReturn=oCurrentNode.getItem(iCount).getData().toString()
									Else
										sReturn=sReturn+","+oCurrentNode.getItem(iCount).getData().toString()
									End If
								Next
								arr = Split(sReturn,",")
								Fn_SISW_SP_NavTree_NodeOperation = arr
								Set oCurrentNode=Nothing
								Exit Function
							Else
								Set oCurrentNode = JavaWindow("ServicePlanner").JavaTree("NavTree").Object.getItem(0)
								intNodeCount=0
								For each echStrNode In arrStrNode
									iRows = oCurrentNode.getItemCount()
									For iCounter = 0 to iRows - 1
										If oCurrentNode.getItem(iCounter).getData().toString() = echStrNode Then
											Set oCurrentNode=oCurrentNode.getItem(iCounter)
											intNodeCount = oCurrentNode.getItemCount()
											Exit For
										End If
									Next
								Next 
								For iCount=0 To intNodeCount-1
									If iCount=0 Then
										sReturn=oCurrentNode.getItem(iCount).getData().toString()
									Else
										sReturn=sReturn+","+oCurrentNode.getItem(iCount).getData().toString()
									End If
								Next
								arr = Split(sReturn,",")
								Fn_SISW_SP_NavTree_NodeOperation = arr
								Set oCurrentNode=Nothing
							End If
						Else
							Fn_SISW_SP_NavTree_NodeOperation = False
						End If

		'****************************************************************************************	
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail :[ Fn_SISW_SP_NavTree_NodeOperation ] Invalid case [ " & StrAction &" ].")
				Exit function
	End Select

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_SISW_SP_NavTree_NodeOperation")
	Set objJavaWindowMyTc = nothing
	Set objJavaTreeNav = nothing
End Function

'********************** Function to perform operations on Create Fault Code Type dialog in Service Planner ***************************************
'
''Function Name		 	:	Fn_SISW_SP_CreateFaultCodeTypeOperations
'
''Description		    :	Function to perform operations on Create Fault Code Type dialog in Service Planner
'
''Parameters		    :	1. sAction : Action need to perform
'					  					2. dicFaultCodeType : Dictionary object to set Fault Code Type data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Planner perspective should be opened.

''Examples  			:	Dim dicFaultCodeType
'					  								Set dicFaultCodeType = CreateObject("Scripting.Dictionary")
'													dicFaultCodeType("FaultCodeType")  = "Fault Code"
'					  								dicFaultCodeType("Name") = "name1"
'					  								dicFaultCodeType("Description") = "name desc"

'							
'					  		msgbox 	Fn_SISW_SP_CreateFaultCodeTypeOperations("Create", dicFaultCodeType)

'History:
'	Developer Name			Date			           Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali    wani              28-Nov-2012			0001					Created		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_CreateFaultCodeTypeOperations(sAction, dicFaultCodeType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_CreateFaultCodeTypeOperations"
	Dim objDialog, strNodePathC
	Fn_SISW_SP_CreateFaultCodeTypeOperations = False
	Set objDialog = JavaWindow("ServicePlanner").JavaWindow("NewFaultCode")
	
	'Select menu [ File -> New -> Fault Code.. ]
	If Fn_UI_ObjectExist("Fn_SISW_SP_CreateFaultCodeTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Fault Code...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SP_CreateFaultCodeTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_CreateFaultCodeTypeOperations ] Failed to find [ Fault Code Type ] dialog.")
			Exit Function
		End If
	End If

	Select Case sAction
		Case "Create"

			If objDialog.JavaTree("FaultCodeTypeTree").Exist(5) Then
				Call Fn_UI_JavaTree_Expand( "Fn_SISW_SP_CreateFaultCodeTypeOperations" , objDialog, "FaultCodeTypeTree" , "Complete List")

				If Fn_UI_JavaTree_NodeExist("Fn_SISW_SP_CreateFaultCodeTypeOperations",objDialog.JavaTree("FaultCodeTypeTree"),"Complete List:"+dicFaultCodeType("FaultCodeType")) Then
							strNodePathC="Complete List:"+dicFaultCodeType("FaultCodeType")
				Else
							strNodePathC="Most Recently Used:"+dicFaultCodeType("FaultCodeType")
				End If
			
				'Select [ Fault Code ]
				Call Fn_JavaTree_Select( "Fn_SISW_SP_CreateFaultCodeTypeOperations", objDialog, "FaultCodeTypeTree", strNodePathC )				
				Call Fn_ReadyStatusSync(2)
						
				'Click on NEXT button to navigate ahead
				Call Fn_Button_Click("Fn_SISW_SP_CreateFaultCodeTypeOperations", objDialog, "Next")			
				Call Fn_ReadyStatusSync(2)
			End If
			
			' Name
			If dicFaultCodeType("Name") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SP_CreateFaultCodeTypeOperations", objDialog,"Name",dicFaultCodeType("Name"))
			End If

			' Description
			If dicFaultCodeType("Description") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SP_CreateFaultCodeTypeOperations", objDialog,"Description",dicFaultCodeType("Description"))
			End If

			Call Fn_Button_Click("Fn_SISW_SP_CreateFaultCodeTypeOperations", objDialog , "Finish")
			Fn_SISW_SP_CreateFaultCodeTypeOperations = Fn_Button_Click("Fn_SISW_SP_CreateFaultCodeTypeOperations", objDialog , "Cancel")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_CreateFaultCodeTypeOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SP_CreateFaultCodeTypeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SP_CreateFaultCodeTypeOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function

'*********************************************************		Function to Get BOM Table Column Index into Service Planner		***********************************************************************
'Function Name		:					Fn_SISW_SP_AssignCharacteristics

'Description			 :		 		  This function is used to assign characteristics to selected node (e.g. Work Card)

'Parameters			   :	 			1.  sNodeName : Name of the Characteristic which needs to be selected
'													 2. sColName : Name of Column (for further use)
'													 3. sButton : Name of button 
											
'Return Value		   : 				 True/ False

'Pre-requisite			:		 		Assign Characteristics Window should be open in Service Planner Perspective

'Examples				:				'bReturn =  Fn_SISW_SP_AssignCharacteristics("obs110", "", "OK") 	' to select one characteristic
'													'bReturn =  Fn_SISW_SP_AssignCharacteristics("obs112~date112~life113", "", "OK") 	' to select multiple characteristic

'History:
'										Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Pallavi Jadhav			28-Nov-2012			1.0																	Koustubh 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_AssignCharacteristics(sNodeName, sColName, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_AssignCharacteristics"
	Dim iRows, iCnt, objAssignChar, bFlag, sPropName, sArr, iCounter
	Set objAssignChar = Fn_SISW_SP_GetObject("AssignCharacteristics")
	bFlag = false
	sArr = split(sNodeName, "~", -1, 1)
	If Fn_UI_ObjectExist("Fn_SISW_SP_AssignCharacteristics", objAssignChar) = True Then
		iRows = cInt(objAssignChar.JavaTable("CharacteristicsDefinitions").GetROProperty("rows"))
		For iCnt = 0 to iRows-1
			For iCounter = 0 to Ubound(sArr)	
				sPropName = objAssignChar.JavaTable("CharacteristicsDefinitions").Object.getItem(iCnt).getData().toString()
				If sPropName = sArr(iCounter) Then
					wait 2
					If iCnt = 0 Then
						objAssignChar.JavaTable("CharacteristicsDefinitions").ClickCell iCnt, 0
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully selected Characteristic [ "+sArr(iCounter)+" ]." )
					Else
						objAssignChar.JavaTable("CharacteristicsDefinitions").ExtendRow(iCnt)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully selected Characteristic [ "+sArr(iCounter)+" ]." )
					End If
					wait 1
				End If
			Next
		Next
		If sButton = "" Then
			sButton = "OK"			
		End If
		If cInt(objAssignChar.JavaButton(sButton).getROProperty("enabled")) = 0 Then
			objAssignChar.JavaButton(sButton).object.setEnabled true
			 bFlag = True
		End If
		Call Fn_Button_Click("Fn_SISW_SP_AssignCharacteristics", objAssignChar, sButton)
		wait 2
		Fn_SISW_SP_AssignCharacteristics = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully Assigned Characteristics [ "+Replace(sNodeName, "~", ", ")+" ]." )
	Else
		If bFlag = False Then
			Fn_SISW_SP_AssignCharacteristics = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to assign Characteristics [ "+Replace(sNodeName, "~", ", ")+" ]." )
			Exit Function
		End If
	End If
End Function
'********************** Function to create frequency on service requirement in Service Planner ***************************************
'
''Function Name		 	:	Fn_SISW_SP_FrequencyOperations
'
''Description		    :	Function to create Frequency in Service Planner
'
''Parameters		    :	1. sAction : Action need to perform
'							2. sFrequency :frequency to worked on  	
'					  		3. dicNewFrequency : Dictionary object to create new frequency
'							4. sButton : button name OK/Cancel									
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	New Frequency/ Edit Frequency dialog must be opened .

''Examples  			:	Dim dicNewFrequency
							'Set dicNewFrequency = createObject("Scripting.Dictionary")
							'dicNewFrequency("ID") = "ID145"
							'dicNewFrequency("Name") = "Frequency2"
							'dicNewFrequency("Description") = "Frequency"
							'dicNewFrequency("Keywords") = "At"
							'dicNewFrequency("Value1") = "1"
							'dicNewFrequency("Characteristics") = "Calendar"
							'dicNewFrequency("Date") = "today"
							'dicNewFrequency("ToleranceOperator") = "+"
							'dicNewFrequency("ToleranceValue") = "2"
							'dicNewFrequency("ToleranceType") = "%"
							'dicNewFrequency("After/Unit") = "Until"
							'dicNewFrequency("AdvancedValue1") = "4"
							'dicNewFrequency("AdvancedCharacteristics") = "Days"

							'Call Fn_SISW_SP_FrequencyOperations("Create","",dicNewFrequency,"OK")

'History:
'	Developer Name			Date		Rev. No.	Reviewer			Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali wani		3-Dec-2012		  01		Koustubh W			Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_FrequencyOperations(sAction, sFrequency, dicNewFrequency, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_FrequencyOperations"
	Dim objDialog, strNodePathC, dictItems, dictKeys, iCounter,bFlag
	Dim iRows,iCnt,arrDateTime,i,sFrequencyValue,iRowSelect
	Fn_SISW_SP_FrequencyOperations = False

	Set objDialog = JavaWindow("ServicePlanner").JavaWindow("Frequency")
	If Fn_UI_ObjectExist("Fn_SISW_SP_FrequencyOperations", objDialog) = False Then
		Exit Function
	End If
	Select Case sAction
		Case "Create"
			dictItems = dicNewFrequency.Items
			dictKeys = dicNewFrequency.Keys
			
			For iCounter = 0 to dicNewFrequency.Count - 1
				If IsNull(DictKeys(iCounter)) = False Then
					If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
						Select Case DictKeys(iCounter)
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "Keywords", "Characteristics","Operators","PhraseSeparator"
								Call Fn_List_Select("Fn_SISW_SP_FrequencyOperations", objDialog, DictKeys(iCounter),dictItems(iCounter))
								'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "ToleranceOperator", "AdvancedOperators","AdvancedCharacteristics","ToleranceType","After/Unit"
								Call Fn_SISW_UI_Twistie_Operations("Fn_SISW_SP_FrequencyOperations", "Expand", objDialog, "Twistie", "Advanced","Advanced")
								wait 1
								Call Fn_List_Select("Fn_SISW_SP_FrequencyOperations", objDialog, DictKeys(iCounter),dictItems(iCounter))
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "AdvancedDate", "AdvancedValue1","AdvancedValue2","PhraseSeparator"
								Call Fn_SISW_UI_Twistie_Operations("Fn_SISW_SP_FrequencyOperations", "Expand", objDialog, "Twistie", "Advanced","Advanced")
								wait 1
								Call Fn_Edit_Box("Fn_SISW_SP_FrequencyOperations", objDialog , DictKeys(iCounter),dictItems(iCounter))
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "Date"
								Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog, "Date")
								If lcase(DictKeys(iCounter)) = "today" then
									Call Fn_UI_SetDateAndTime("Fn_SISW_SP_FrequencyOperations","Today","")
								Else
									arrDateTime = Split(DictKeys(iCounter)," ")
									Call Fn_UI_SetDateAndTime("Fn_SISW_SP_FrequencyOperations",arrDateTime(0),arrDateTime(1))
								End If
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case  "AdvancedDate"
								Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog, "AdvancedDate")
								If lcase(DictKeys(iCounter)) = "today" then
									Call Fn_UI_SetDateAndTime("Fn_SISW_SP_FrequencyOperations","Today","")
								Else
									arrDateTime = Split(DictKeys(iCounter)," ")
									Call Fn_UI_SetDateAndTime("Fn_SISW_SP_FrequencyOperations",arrDateTime(0),arrDateTime(1))
								End If
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case Else
								' Edit Box 
								bReturn = Fn_Edit_Box("Fn_SISW_SP_FrequencyOperations", objDialog , DictKeys(iCounter),dictItems(iCounter))
								If bReturn = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_FrequencyOperations ] Failed to set Editbox [ " & DictKeys(iCounter) & " = " & dictItems(iCounter) & " ].")
									Exit Function
								End If
						End Select
					End If
				End If
			Next
			Call Fn_SISW_UI_Twistie_Operations("Fn_SISW_SP_FrequencyOperations", "Collapse", objDialog, "Twistie", "Advanced","Advanced")
			wait 1
			Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog, "Add")
			Fn_SISW_SP_FrequencyOperations = True
		Case "Edit"
			
			iRows = cInt(objDialog.JavaTable("FrequencyExpression").GetROProperty("rows"))
			For iCnt= 0 to iRows - 1
				sData = objDialog.JavaTable("FrequencyExpression").GetCellData(iCnt,0)
				If sData = sFrequency	Then
					objDialog.JavaTable("FrequencyExpression").SelectCell iCnt,"0"
					Exit for 
				End If
			Next
			Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog, "Edit")

			dictItems = dicNewFrequency.Items
			dictKeys = dicNewFrequency.Keys
			For iCounter = 0 to dicNewFrequency.Count - 1
				If IsNull(DictKeys(iCounter)) = False Then
					If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
						Select Case DictKeys(iCounter)
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "Keywords", "Characteristics","Operators","PhraseSeparator"
								Call Fn_List_Select("Fn_SISW_SP_FrequencyOperations", objDialog, DictKeys(iCounter),dictItems(iCounter))
								'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "ToleranceOperator", "AdvancedOperators","AdvancedCharacteristics","ToleranceType","After/Unit"
								Call Fn_SISW_UI_Twistie_Operations("Fn_SISW_SP_FrequencyOperations", "Expand", objDialog, "Twistie", "Advanced","Advanced")
								wait 1
								Call Fn_List_Select("Fn_SISW_SP_FrequencyOperations", objDialog, DictKeys(iCounter),dictItems(iCounter))
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "AdvancedDate", "AdvancedValue1","AdvancedValue2","PhraseSeparator"
								Call Fn_SISW_UI_Twistie_Operations("Fn_SISW_SP_FrequencyOperations", "Expand", objDialog, "Twistie", "Advanced","Advanced")
								wait 1
								Call Fn_Edit_Box("Fn_SISW_SP_FrequencyOperations", objDialog , DictKeys(iCounter),dictItems(iCounter))
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "Date"
								Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog, "Date")
								If lcase(dictItems(iCounter)) = "today" then
									Call Fn_UI_SetDateAndTime("Fn_SISW_SP_FrequencyOperations","Today","")
								Else
									arrDateTime = Split(DictKeys(iCounter)," ")
									Call Fn_UI_SetDateAndTime("Fn_SISW_SP_FrequencyOperations",arrDateTime(0),arrDateTime(1))
								End If
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case  "AdvancedDate"
								Call Fn_SISW_UI_Twistie_Operations("Fn_SISW_SP_FrequencyOperations", "Expand", objDialog, "Twistie", "Advanced","Advanced")
								Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog, "AdvancedDate")
								If lcase(dictItems(iCounter)) = "today" then
									Call Fn_UI_SetDateAndTime("Fn_SISW_SP_FrequencyOperations","Today","")
								Else
									arrDateTime = Split(DictKeys(iCounter)," ")
									Call Fn_UI_SetDateAndTime("Fn_SISW_SP_FrequencyOperations",arrDateTime(0),arrDateTime(1))
								End If
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case Else
								' Edit Box 
								bReturn = Fn_Edit_Box("Fn_SISW_SP_FrequencyOperations", objDialog , DictKeys(iCounter),dictItems(iCounter))
								If bReturn = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_FrequencyOperations ] Failed to set Editbox [ " & DictKeys(iCounter) & " = " & dictItems(iCounter) & " ].")
									Exit Function
								End If
						End Select
					End If
				End If
			Next
			Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog, "Save")
			Fn_SISW_SP_FrequencyOperations = True
		Case "GroupFrequency"
			sFrequencyValue = Split(dicNewFrequency("Group"),"~")
			iRows = cInt(objDialog.JavaTable("FrequencyExpression").GetROProperty("rows"))
			For i = 0 to UBound(sFrequencyValue)
				For iCnt= 0 to iRows-1
					sData = objDialog.JavaTable("FrequencyExpression").GetCellData(iCnt,0)
					If sData = sFrequencyValue(i) Then
						Call objDialog.JavaTable("FrequencyExpression").ExtendRow(iCnt)
						Exit For
					End If
				Next
			Next
			Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog,  "Group")
			Fn_SISW_SP_FrequencyOperations = True
		Case "UnGroupFrequency"
			sFrequencyValue = Split(dicNewFrequency("UnGroup"),"~")
			iRows = cInt(objDialog.JavaTable("FrequencyExpression").GetROProperty("rows"))
			For i = 0 to UBound(sFrequencyValue)
				For iCnt = 0 to iRows-1
					sData = objDialog.JavaTable("FrequencyExpression").GetCellData(iCnt,0)
					If sData =sFrequencyValue(i) Then
						Call objDialog.JavaTable("FrequencyExpression").ExtendRow(iCnt)
						Exit For
					End If
				Next
			Next
			Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog,  "UnGroup")
			Fn_SISW_SP_FrequencyOperations = True
		Case "RemoveFrequency"
			iRows = objDialog.JavaTable("FrequencyExpression").GetROProperty("rows")
			For iCnt= 0 to iRows-1
				sData = objDialog.JavaTable("FrequencyExpression").GetCellData(iCnt,0)
				If sData = dicNewFrequency("RemoveFrequency") Then
					objDialog.JavaTable("FrequencyExpression").SelectCell iCnt, 0
					Exit for 
				End If
			Next
			Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog, "Remove")
			Fn_SISW_SP_FrequencyOperations = True
		Case "Verify"
    		iRows = cInt(objDialog.JavaTable("FrequencyExpression").GetROProperty("rows"))
			bFlag=False
			For iCnt= 0 to iRows - 1
				sData = objDialog.JavaTable("FrequencyExpression").GetCellData(iCnt,0)
				If sData = sFrequency	Then
					bFlag=True
					Exit for 
				End If
			Next
			If bFlag = False Then
				Exit Function
			End If
			If dicNewFrequency<>"" Then
				dictItems = dicNewFrequency.Items
				dictKeys = dicNewFrequency.Keys
				For iCounter = 0 to dicNewFrequency.Count - 1
					If IsNull(DictKeys(iCounter)) = False Then
						If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
							Select Case DictKeys(iCounter)
								'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								Case "Keywords", "Characteristics","Operators","PhraseSeparator"
									'JavaList
									If Fn_SISW_UI_JavaList_Operations("Fn_SISW_SP_FrequencyOperations", "GetText", objDialog,DictKeys(iCounter),"", "", "")<>dictItems(iCounter) Then
										Fn_SISW_SP_FrequencyOperations=False
										Exit Function
									End If
								'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
								Case Else
									' Edit Box 
									If Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SP_FrequencyOperations", "GetText",  objDialog, DictKeys(iCounter), "" )<> dictItems(iCounter) Then
										Fn_SISW_SP_FrequencyOperations=False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_FrequencyOperations ] Failed to set Editbox [ " & DictKeys(iCounter) & " = " & dictItems(iCounter) & " ].")
										Exit Function
									End If
							End Select
						End If
					End If
				Next
			End If
			Fn_SISW_SP_FrequencyOperations = True
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_FrequencyOperations ] Invalid case [ " & sAction & " ].")
	End Select
	If sButton<>"" Then
		Call Fn_Button_Click("Fn_SISW_SP_FrequencyOperations", objDialog , sButton)
	End If
			
	If Fn_SISW_SP_FrequencyOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SP_FrequencyOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function
'*************************************************************
''Function Name		 	:	Fn_SISW_SP_ActivitiesTableOperations
'
''Description		    :	Function to perform operations on Activity Table in Service Planner
'
''Parameters		    :	1. sAction 		: Action need to perform
'							2. sRow 		: Row Path
'					  		3. sColumn 		: Column Name
'							4. sValue 		: Value
'							5. sPopupMenu 	: Popup Menu select
'								
'Return Value		    :  	True / False, -1 / Row Number / Colum Number
'
''Examples  			:	Call Fn_SISW_SP_ActivitiesTableOperations("GetPathForRow", 2,"","","")
''			  			:	Call Fn_SISW_SP_ActivitiesTableOperations("GetRowIndex", "000234/A:Activity","","","")
''			  			:	Call Fn_SISW_SP_ActivitiesTableOperations("GetColumnIndex", "","Line","","")
''			  			:	Call Fn_SISW_SP_ActivitiesTableOperations("Select", "000234/A:Activity","","","")
''			  			:	Call Fn_SISW_SP_ActivitiesTableOperations("Expand", "000234/A:Activity","","","")
''			  			:	Call Fn_SISW_SP_ActivitiesTableOperations("CellVerify", "000234/A:Activity","Frequency","1","")
''			  			:	Call Fn_SISW_SP_ActivitiesTableOperations("CellEdit", "000234/A:Activity","Frequency","2","")

'History:
'	Developer Name			Date		Rev. No.	Reviewer			Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Kousubh Watwe		4-Dec-2012		  01		Koustubh W			Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_ActivitiesTableOperations(sAction, sRow, sColumn, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_ActivitiesTableOperations"
	Dim iCnt, iRows, objActivitiesTable, objComponent, sNodePath, iCols
	Dim objSelectType, intNoOfObjects,bFlag,iCounter
	Fn_SISW_SP_ActivitiesTableOperations = False
	Set objActivitiesTable = JavaWindow("ServicePlanner").JavaTable("ActivityTable")
	If Fn_UI_ObjectExist("Fn_SISW_SP_ActivitiesTableOperations", objActivitiesTable) = False Then
		Exit Function
	End If
	Select Case sAction
		Case "GetPathForRow"
			set objComponent = objActivitiesTable.Object.getNodeForRow(sRow)
			sNodePath = False
			Do while Not (objComponent is nothing)
				If sNodePath = False Then
					sNodePath = objComponent.getProperty("me_cl_display_string")
				Else
					sNodePath = objComponent.getProperty("me_cl_display_string") & ":" & sNodePath
				End If
				'Set objComponent = objComponent.parent()
				If Environment.Value("ProductName") = sQTPProductName OR Environment.Value("ProductName") = sUFTProductName Then
					If objComponent.parent().getProperty("me_cl_display_string") <> "me_cl_display_string" Then
					'If IsObject(objComponent.parent()) = True Then
						set objComponent = objComponent.parent()
					Else
						Exit Do
					End If
				Else
					set objComponent = objComponent.parent()
					If  objComponent is Nothing Then
						Exit do
					End If
				End If
				If objComponent is nothing Then
					Exit do
				End If
			Loop
			Fn_SISW_SP_ActivitiesTableOperations = sNodePath
			set objComponent = Nothing
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetRowIndex"
			' Needs Modifications in future
			Fn_SISW_SP_ActivitiesTableOperations = -1
			iRows = cInt(JavaWindow("ServicePlanner").JavaTable("ActivityTable").GetROProperty("rows"))
			For iCnt = 0 to iRows -1
					sNodePath = Fn_SISW_SP_ActivitiesTableOperations("GetPathForRow", iCnt,"","","")
					If sNodePath = False Then
						Exit for
					Else
						If Instr(sRow, ":") = 0 Then
							If Instr(sNodePath, sRow) > 0 Then
								Fn_SISW_SP_ActivitiesTableOperations = iCnt
								Exit for	
							End If
						ElseIf sNodePath = sRow Then
							Fn_SISW_SP_ActivitiesTableOperations = iCnt
							Exit for
						End If
					End If
			Next
		Case "GetColumnIndex"
			Fn_SISW_SP_ActivitiesTableOperations = -1
			iCols = cInt(objActivitiesTable.GetROProperty("cols"))
			For iCnt = 0 to iCols - 1
				If objActivitiesTable.GetColumnName(iCnt) = sColumn Then
					Fn_SISW_SP_ActivitiesTableOperations = iCnt
					Exit for
				End If
			Next
		Case "Exist"
			iCnt = Fn_SISW_SP_ActivitiesTableOperations("GetRowIndex", sRow,"","","")
			If iCnt <> -1 Then
				Fn_SISW_SP_ActivitiesTableOperations = True
			End If
		' select row
		Case "Select"
			iCnt = Fn_SISW_SP_ActivitiesTableOperations("GetRowIndex", sRow,"","","")
			If iCnt <> -1 Then
				Call objActivitiesTable.SelectRow(iCnt)
				Fn_SISW_SP_ActivitiesTableOperations = True
			End If
		Case "PopupMenuSelect"
			iCnt = Fn_SISW_SP_ActivitiesTableOperations("GetRowIndex", sRow,"","","")
			If iCnt <> -1 Then
				objActivitiesTable.ClickCell CInt(iCnt),1,"RIGHT"
				wait 3
				Fn_SISW_SP_ActivitiesTableOperations = Fn_UI_JavaMenu_Select("Fn_SrvMgr_NavTree_NodeOperation",JavaWindow("ServicePlanner"),sPopupMenu)
			End If
		Case "Expand"
			iCnt = Fn_SISW_SP_ActivitiesTableOperations("GetRowIndex", sRow,"","","")
			If iCnt <> -1 Then
				Call objActivitiesTable.Object.expandNode(objActivitiesTable.Object.getNodeForRow(iCnt))
				Fn_SISW_SP_ActivitiesTableOperations = True
			End If
		Case "CellVerify"
			iCnt = Fn_SISW_SP_ActivitiesTableOperations("GetRowIndex", sRow,"","","")
			iCols = Fn_SISW_SP_ActivitiesTableOperations("GetColumnIndex", "", sColumn,"","")
			If iCnt <> -1 AND iCols <> -1 Then
				If objActivitiesTable.Object.getValueAt(iCnt, iCols).toString() = sValue Then
					Fn_SISW_SP_ActivitiesTableOperations = True
				End If
			End If
		Case "CellEdit"
            bFlag=False
			iCnt = Fn_SISW_SP_ActivitiesTableOperations("GetRowIndex", sRow,"","","")
			iCols = Fn_SISW_SP_ActivitiesTableOperations("GetColumnIndex", "", sColumn,"","")
			If iCnt <> -1 AND iCols <> -1 Then
				objActivitiesTable.ClickCell iCnt, iCols
				' press F2
				Call Fn_KeyBoardOperation("SendKeys", "{F2}")
				Set objSelectType = description.Create()
				objSelectType("Class Name").value = "JavaButton"
				objSelectType("attached text").value = "dropdown_16"
				objSelectType("tagname").value = "dropdown_16"
				objSelectType("path").value = ".*ActivitiesPanel.*"
				objSelectType("path").RegularExpression = True
				Set  intNoOfObjects = JavaWindow("ServicePlanner").ChildObjects(objSelectType)
				If intNoOfObjects.Count > 0 Then
					intNoOfObjects(0).Click 
					wait 1
					Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					objSelectType("label").value = sValue
					Set  intNoOfObjects = JavaWindow("ServicePlanner").ChildObjects(objSelectType)
					If intNoOfObjects.Count > 0 Then
						intNoOfObjects(0).Click 1, 1, "LEFT"
						Fn_SISW_SP_ActivitiesTableOperations = True
					'*Added By Nilesh Gadekar on 30-May-2013 
                    ElseIf Window("ServicePlannerWindow").JavaApplet("JApplet").JavaWindow("JWindow").JavaTable("LOVTreeTable").Exist(5) Then    
						For iCounter=0 to Window("ServicePlannerWindow").JavaApplet("JApplet").JavaWindow("JWindow").JavaTable("LOVTreeTable").GetROProperty("rows")-1
							If sValue=trim(Window("ServicePlannerWindow").JavaApplet("JApplet").JavaWindow("JWindow").JavaTable("LOVTreeTable").Object.getValueAt(iCounter,0).getDisplayableValue())  Then
								Window("ServicePlannerWindow").JavaApplet("JApplet").JavaWindow("JWindow").JavaTable("LOVTreeTable").SelectRowsRange iCounter,iCounter
								bFlag = True
								Exit For
							Else
								bFlag = False
							End If
						Next
						If bFlag = True Then
							Fn_SISW_SP_ActivitiesTableOperations=True
						Else
							Fn_SISW_SP_ActivitiesTableOperations=False
						End If''*End
					Else
						Fn_SISW_SP_ActivitiesTableOperations = False
					End If
					Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				Else
					objActivitiesTable.SetCellData iCnt,iCols,sValue		
					Fn_SISW_SP_ActivitiesTableOperations = True
					Call Fn_KeyBoardOperation("SendKeys", "{TAB}")
				End If
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_ActivitiesTableOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SP_ActivitiesTableOperations <> False AND Fn_SISW_SP_ActivitiesTableOperations <> -1 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SP_ActivitiesTableOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objActivitiesTable = Nothing
End Function 
'*************************************************************
''Function Name		 	:	Fn_SISW_SP_TimePanelOperations
'
''Description		    :	Function to perform operations on Time panel in Service Planner
'
''Parameters		    :	1. sAction 		: Action need to perform
'							2. dicTime 		: Dictionary Object
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	

''Examples  			:	Dim dicTime
'							Set dicTime = CreateObject( "Scripting.Dictionary")
'							With dicTime  
'								.Add "Activity", "000234/A:Activity"
'							End with
''			  			:	Call Fn_SISW_SP_TimePanelOperations("AddActivityAfter", dicTime)
''			  			:	Call Fn_SISW_SP_TimePanelOperations("AddActivityBelow", dicTime)
''			  			:	Call Fn_SISW_SP_TimePanelOperations("OpenDataCard", dicTime)
''			  			:	Call Fn_SISW_SP_TimePanelOperations("OpenTiConSearch", dicTime)
''			  			:	Call Fn_SISW_SP_TimePanelOperations("RemoveActivity", dicTime)

'History:
'	Developer Name			Date		Rev. No.	Reviewer			Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Kousubh Watwe		4-Dec-2012		  01		Koustubh W			Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_TimePanelOperations(sAction, dicTime)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_TimePanelOperations"
   Dim bResult, objDialog
	Fn_SISW_SP_TimePanelOperations = False
	Set objDialog = JavaWindow("ServicePlanner")
    If Fn_SISW_UI_RACTabFolderWidget_Operation("Select", "Time", "") = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_TimePanelOperations ] Failed to select [ Time  ].")
		Exit function
	End If
	Select Case sAction
		Case "AddActivityAfter"
			bResult = Fn_SISW_SP_ActivitiesTableOperations("Select", dicTime("Activity"), "", "", "")
			If bResult Then
				Fn_SISW_SP_TimePanelOperations = Fn_Button_Click("Fn_SISW_SP_TimePanelOperations", objDialog, "AddActivityAfter")
			End If
		Case "AddActivityBelow"
			bResult = Fn_SISW_SP_ActivitiesTableOperations("Select", dicTime("Activity"), "", "", "")
			If bResult Then
				Fn_SISW_SP_TimePanelOperations = Fn_Button_Click("Fn_SISW_SP_TimePanelOperations", objDialog, "AddActivityBelow")
			End If
		Case "OpenDataCard"
			bResult = Fn_SISW_SP_ActivitiesTableOperations("Select", dicTime("Activity"), "", "", "")
			If bResult Then
				Fn_SISW_SP_TimePanelOperations = Fn_Button_Click("Fn_SISW_SP_TimePanelOperations", objDialog, "OpenDataCard")
			End If
		Case "OpenTiConSearch"
			bResult = Fn_SISW_SP_ActivitiesTableOperations("Select", dicTime("Activity"), "", "", "")
			If bResult Then
				Fn_SISW_SP_TimePanelOperations = Fn_Button_Click("Fn_SISW_SP_TimePanelOperations", objDialog, "OpenTiConSearch")
			End If
		Case "RemoveActivity"
			bResult = Fn_SISW_SP_ActivitiesTableOperations("Select", dicTime("Activity"), "", "", "")
			If bResult Then
				Fn_SISW_SP_TimePanelOperations = Fn_Button_Click("Fn_SISW_SP_TimePanelOperations", objDialog, "RemoveActivity")
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SP_TimePanelOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SP_TimePanelOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SP_TimePanelOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function 

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''Function Name		 	:	Fn_SISW_SP_ActivityAssignmentsOperations
'
''Description		    :	Function to perform operations on Activity Table in Service Planner
'
''Parameters		    :	1. sAction 	    	: Action need to perform
'										2. sRow 			 :  Row of table
'					  					3. sColumn 		  : Column Name
'										4. sValue 			: Value
'										5. sButtonName 	: Name of button to be pressed at last of function (It should only be 'OK' or 'Cancel' or blank "" )
'								
'Return Value		    :  	True / False
'
'Pre-Requisite			: Activity Assignments dialog should be visible
'
''Examples  			:	Call Fn_SISW_SP_ActivityAssignmentsOperations("Add", "000020/A;1-Part2","","","OK")
''			  						:	Call Fn_SISW_SP_ActivityAssignmentsOperations("Remove", "000020/A;1-Part2","","","Cancel")
''			  						:	Call Fn_SISW_SP_ActivityAssignmentsOperations("VerifyInOperationTable", "000020/A;1-Part2","Find No.","10","")
''			  						:	Call Fn_SISW_SP_ActivityAssignmentsOperations("VerifyInActivityTable", "000020/A;1-Part2","BOM Line","000020/A;1-Part2","")
'History:
'	Developer Name					Date				Rev. No.					Changes Done					Reviewer			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar					14-Oct-2013		  		01								Created 							Pranav Ingle	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_ActivityAssignmentsOperations(sAction, sRow, sColumn, sValue,sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_ActivityAssignmentsOperations"
	Dim objActivityAssignmentsDialog,bResult
	Fn_SISW_SP_ActivityAssignmentsOperations = False
	Set objActivityAssignmentsDialog = JavaWindow("ServicePlanner").JavaWindow("WEmbeddedFrame").JavaDialog("ActivityAssignments")
	If Fn_SISW_UI_Object_Operations("Fn_SISW_SP_ActivityAssignmentsOperations","Exist", objActivityAssignmentsDialog,"")  = False Then Exit Function
	Select Case sAction
		Case "Add"
			bResult= Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SP_ActivityAssignmentsOperations", "ClickCell", objActivityAssignmentsDialog , "OccurencesOnOperationTable", "", "BOM Line", sRow, sColumn, "", "", "")
			If bResult = False Then
				Exit Function
			End If
			bResult=Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SP_ActivityAssignmentsOperations", "Click", objActivityAssignmentsDialog,"GoRight")
			If bResult = False Then
				Exit Function
			End If
		Case "Remove"
			bResult= Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SP_ActivityAssignmentsOperations", "ClickCell", objActivityAssignmentsDialog , "OccurencesOnActivityTable", "", "BOM Line", sRow, sColumn, "", "", "")
			If bResult = False Then
				Exit Function
			End If
			bResult=Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SP_ActivityAssignmentsOperations", "Click", objActivityAssignmentsDialog,"GoLeft")
			If bResult = False Then
				Exit Function
			End If
		Case "VerifyInOperationTable"
			bResult= Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SP_ActivityAssignmentsOperations", "VerifyCellData", objActivityAssignmentsDialog , "OccurencesOnOperationTable", "", "BOM Line", sRow, sColumn, sValue, "", "")
			If bResult = False Then
				Exit Function
			End If
		Case "VerifyInActivityTable"
			bResult= Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SP_ActivityAssignmentsOperations", "VerifyCellData", objActivityAssignmentsDialog , "OccurencesOnActivityTable", "", "BOM Line", sRow, sColumn, sValue, "", "")
			If bResult = False Then
				Exit Function
			End If
		End Select
		If sButtonName<>"" Then
			Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SP_ActivityAssignmentsOperations", "Click", objActivityAssignmentsDialog,sButtonName)
		End If
     Fn_SISW_SP_ActivityAssignmentsOperations=True
	Set objActivityAssignmentsDialog = Nothing
End Function 

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''Function Name		 	:	Fn_SISW_SP_PopulateAllocatedTimeOperations
'
''Description		    :	Function to perform operations on PopulateAllocatedTime dialog in Service Planner
'
''Parameters		    :	1. sAction 	    	 : Action need to perform
'										 2. dicInputs 		 :  Dictionary of values for action
'										 3. sButtonName 	: Name of button to be pressed at last of function (It should only be 'OK' or 'Cancel' or blank "" )
'								
'Return Value		    :  	True / False
'
'Pre-Requisite			: Populate Allocated Time dialog should be visible
'
''Examples  			:	set dic2 = CreateObject( "Scripting.Dictionary" )
'										dic2("PopulatedTime") = "Simulated Time"
'										dic2("PopulateZeroValues") = "ON"
'										dic2("PopulateUptoLevel") = "10"

'										Call Fn_SISW_SP_PopulateAllocatedTimeOperations("Select",dic2,"OK")
'History:
'	Developer Name					Date				Rev. No.					Changes Done					Reviewer			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar					16-Oct-2013		  		01					    Created 						Pritam Shikare	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SP_PopulateAllocatedTimeOperations(sAction , dicInputs,sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SP_PopulateAllocatedTimeOperations"
	   Dim objDialog, dicItem,dicValue,iCount,bResult
	   Fn_SISW_SP_PopulateAllocatedTimeOperations = False

	   Set objDialog = Fn_SISW_SP_GetObject("PopulateAllocatedTime")
		If Fn_SISW_UI_Object_Operations("Fn_SISW_SP_PopulateAllocatedTimeOperations","Exist", objDialog,"") = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: PopulateAllocatedTime Dialog doesn't Exist.")
			Exit Function
		End If

	    dicItem = dicInputs.Keys
		dicValue = dicInputs.Items
		Select Case sAction
    		Case "Select"
				For iCount = 0 to dicInputs.Count - 1
					Select Case trim(dicItem(iCount))
						Case "PopulatedTime"
							'Radio button
							objDialog.JavaRadioButton("CommonRadioButton").SetTOProperty "attached text",trim(dicValue(iCount))
							bResult =Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_SP_PopulateAllocatedTimeOperations", "Set", objDialog,"CommonRadioButton", "ON")
							If bResult = False Then Exit Function
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						 Case "PopulateUptoLevel"
						 'Check Box
							bResult = Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_SP_PopulateAllocatedTimeOperations", "Set", objDialog, "PopulateUptoLevelChkBox", "ON")
							If bResult = False Then Exit Function
							bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SP_PopulateAllocatedTimeOperations", "Set",  objDialog, "PopulateUptoLevelSpinner",trim(dicValue(iCount)))
							If bResult = False Then Exit Function
							Wait 3
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "PopulateZeroValues"
						'Check Box
							bResult = Fn_SISW_UI_JavaCheckBox_Operations("Fn_SISW_SP_PopulateAllocatedTimeOperations", "Set", objDialog, "PopulateZeroValuesChkBox", trim(dicValue(iCount)))
							If bResult = False Then Exit Function
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "InsertIntoPopulatedTimeProperties"
							'Not Done yet
						Case "RemoveFromPopulatedTimeProperties"
							'Not Done yet
					End Select
				Next
			Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Not a valid action.")
				Exit Function
		End Select

		If sButtonName <> "" Then
			Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SP_PopulateAllocatedTimeOperations", "Click", objDialog,sButtonName)
		End If

		Fn_SISW_SP_PopulateAllocatedTimeOperations=True
		Set objDialog = Nothing
End Function
