'----------------------'Global variables for Teamcenter Perspective Names----------------------------------------------------------------
Public GBL_PERSPECTIVE_MULTI_STRUCTURE_MANAGER
GBL_PERSPECTIVE_MULTI_STRUCTURE_MANAGER="Multi-Structure Manager"
'----------------------'Global variables for Teamcenter Perspective Names----------------------------------------------------------------

'*********************************************************		Function Names ***********************************************************************
'0. Fn_SISW_MSM_GetObject
'1. Fn_MSM_TableRowIndex()
'2. Fn_MSM_BOMTableNodeOpeations()
'3. Fn_MSM_BOMTabPanelOperation()
'4. Fn_MSM_TableColumnIndex()
'5. Fn_MSM_LineTableNodeOperation()
'6. Fn_MSM_AttachmentsTableNodeOpration()
'7. Fn_MSM_OccurrenceGroup()
'8. Fn_MSM_VariantConditionOperations()
'9. Fn_MSM_VariantCreate()
'10 Fn_MSM_CreateAllocationMap()
'11 Fn_MSM_CreateAllocationContext()
'12 Fn_MSM_ErrorWindowVerify()
'13 Fn_MSM_CreateAllocation()
'14 Fn_MSM_SetViewConfigurationToContextOperations()
'15 Fn_MSM_ErrorDialogVerify(sMesssage, sButton)
'16 Fn_MSM_AllocationTableNodeOperation()
'17 Fn_SISW_MSM_ErrorVerify()
'18 Fn_MSM_DataPanelTabOperations
'19 Fn_MSM_BottomContextOperations()
'20 Fn_MSM_Publish_UnPublish_Operation
'21 Fn_MSM_ConfigurationContextOperations
'***************** Functions can be used in Multi-Stucture Manager from StructureManager vbs ***********************************************
'1  Fn_PSE_ICInfoTabOperations()
'2  Fn_PSE_UnDOICChanges()
'3  Fn_PSE_SaveNewConfigContext()
'4	Fn_PSE_RevRuleSetEffectivityGroup()

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			Function to get Object hierarchy  		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_SISW_MSM_GetObject
'
''Description		  	 :  	Function to get Object hierarchy

''Parameters		   :	1. sObjectName : Object Handle name
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_MSM_GetObject("JApplet")

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'									Sonal Padmawar		 				4-July-2012				1.0					Sunny
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_MSM_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\MultiStructureMananger.xml"
	Set Fn_SISW_MSM_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'*********************************************************		Function to Get Table Node Index in Multi-Structure Manager		***********************************************************************

'Function Name		        :				Fn_MSM_TableRowIndex

'Description			  :		 		This function is used to get the RootStructures Table Node Index.

'Parameters			 :	 			1. objTable : Object of Java Table
'									2. StrNodeName : Name of the Node to retrieve Index for.
											
'Return Value		   	: 					Node index / -1

'Pre-requisite			:				 Multi-Structure Manager window should be displayed .

'Examples				:				Fn_MSM_TableRowIndex(objTable,"001729/A;1-top (View):001736/A;1-asm (View) @2:001737/A;1-sub", "BOM Line")
'								         Fn_MSM_TableRowIndex(objTable,"001729/A;1-top (View):001736/A;1-asm (View) @2:001737/A;1-sub", "")


'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep				18-June-2010				1.0																Tushar
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh				 25-Oct-2010			 	2.0			Changed Name of the function
'																								Changed Logic
'																								Added parameters. objTable, sCol		
'																								Made generic function of getting row index
'																								Added call to Fn_MSM_TableColumnIndex		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_TableRowIndex(objTable, sNodeName, sCol)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_TableRowIndex"
	Dim nodeArr, aRowNode, iColIndex
	Dim iRowCounter, sNode, iInstance, iNodeCounter, iPathCounter, bFound 
	Dim iRows, sNodePath

	If Fn_UI_ObjectExist("Fn_MSM_TableRowIndex", objTable) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MSM_TableRowIndex ] Table does not exist.")	
		Fn_MSM_TableRowIndex = -1
		Exit function
	End If

	If sCol <> "" Then
		iColIndex = cInt( Fn_MSM_TableColumnIndex(objTable, sCol) )
		If iColIndex = -1 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MSM_TableRowIndex ] specified Column [ " & sCol & " ]  does not exist in table.")	
			Fn_MSM_TableRowIndex = -1
			Exit function
		End If
	Else 
		iColIndex = 0
	End If
	bFound = False
	If sNodeName <> "" Then
		' identifying RowId
		iRows = cInt(objTable.GetROProperty ("rows"))
		nodeArr = split(sNodeName , ":")
		iRowCounter = 0
		For iNodeCounter=0 to UBound(nodeArr)
			If iRowCounter = iRows  Then
				Exit for
			End If
			aRowNode = split(trim((nodeArr(iNodeCounter))),"@")
			iInstance = 0
			bFound = False
			do While iRowCounter < iRows
				If uBound(aRowNode) > 0 Then
					' instance number exist in name
					' initialize instance num
					' ith row matches with aRowNode(0) then
					sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
					If trim(sNodePath) = trim(aRowNode(0)) then
						iInstance = iInstance +1
						If iInstance = cInt(aRowNode(1)) Then 
							If UBound(nodeArr) = iNodeCounter Then
								bFound = True
							End If
							Exit do
						End If
						'exit loop
					End if
				Else
					'ith row matches with aRowNode(0) then
					sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
					If trim(sNodePath) = trim(aRowNode(0)) then
						If UBound(nodeArr) = iNodeCounter Then
							bFound = True
						End If
						Exit do
						'exit loop
					End if
				End If
				iRowCounter = iRowCounter + 1
				' increment counter
			loop
		Next
	End If
	If bFound Then
		Fn_MSM_TableRowIndex = iRowCounter
	Else
		Fn_MSM_TableRowIndex = -1
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MSM_TableRowIndex ] executed successfully.")
End Function

'*********************************************************		Function to Get BOM Table Node operation in MultiStructure Manager		***********************************************************************

'	Function Name			:			Fn_MSM_BOMTableNodeOpeations

'	Description				:		 	This function is used to get the ReqMgr Table Node operation.

'	Parameters			   	:			1.	strAction = "Select"
'										2. StrNodeName:Name of the Node. 
'										3. strColName
'										4. strColValue
'										5. strPopupMenu
										
'	Return Value		   	: 			True/ False/TableNodeIndex

'	Pre-requisite			:			MultiStructure Manager window should be displayed .

'	Examples				:			Fn_MSM_BOMTableNodeOpeations("Select","001729/A;1-top (View):001736/A;1-asm (View) @2:001737/A;1-sub","","","")
'										Fn_MSM_BOMTableNodeOpeations("Exist","001729/A;1-top (View):001736/A;1-asm (View) @2:001737/A;1-sub","","","")
'										Fn_MSM_BOMTableNodeOpeations("ExpandBelow","001729/A;1-top (View):001736/A;1-asm (View) @2:001737/A;1-sub","","","")
'										Fn_MSM_BOMTableNodeOpeations("DeSelect","003465/A;1-Test (View):003498/A;1-Tes","","","")
'										Fn_MSM_BOMTableNodeOpeations("VerifyNode","003465/A;1-Test (View):003498/A;1-Tes","","","")
'										Fn_MSM_BOMTableNodeOpeations("PopupMenuSelect","","","","Copy")
'										Fn_MSM_BOMTableNodeOpeations("AddColumns","","Relation~Status","","")
'										Fn_MSM_BOMTableNodeOpeations("RemoveColumns","","Relation","","")
'									 	Fn_MSM_BOMTableNodeOpeations("LoadInViewer","000041/A;1-TopAssm (View):000042/A;1-ChildAssm","","","")
'									 	Fn_MSM_BOMTableNodeOpeations("VerifyLoadInViewerState","000041/A;1-TopAssm (View):000042/A;1-ChildAssm","",True,"")
'									 	Fn_MSM_BOMTableNodeOpeations("VerifyLoadInViewerState","000041/A;1-TopAssm (View):000042/A;1-ChildAssm","",False,"")
'									 	Fn_MSM_BOMTableNodeOpeations("UnloadFromViewer","000041/A;1-TopAssm (View):000042/A;1-ChildAssm","","","")
'										Fn_MSM_BOMTableNodeOpeations("VerifyRootStructureList", "", "", "", "")
'										Fn_MSM_BOMTableNodeOpeations("VerifyRootStructureListValues", "000027-item1~000028/A;1-item2", "", "", "")
'										Fn_MSM_BOMTableNodeOpeations("SelectRootStructureFromList", "000028/A;1-item2", "", "", "")
'	History:
'
'	Developer Name			Date				Rev. No.										Changes Done																					Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sandeep				19-June-2010			1.0																																				Tushar								
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh			25-Oct-2010				1.0			Changed function call to get Row index.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh			18-Apr-2011				1.0			Added case ExpandBelow, removed duplicate code of case Exist 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sonal				11-Apr-2012				1.1			Added case CellEdit and modified case CellVerify
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Naveen				10-Sept-2012			1.2			Added case MultiSelect
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
'	Koustubh			19-Jul-2013				1.2			Added case AddColumns, RemoveColumns
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
'	Poonam				13-Apr-2015				1.2			Added condition to handle Requirement Manager test cases (Column name changed from BOM line to Requirement)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam			15-Feb-2016				1.2			Added new Case "SelectRootStructureFromList", "VerifyRootStructureListValues", "VerifyRootStructureList"           [Tc1122:2016011300:15Feb2016:VivekA:NeDevelopment]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Madhura P			03-May-2016				1.2			Added new Case "GetSelected" - returns selected BOM Line elements      	[TC1122-20160420-03_05_2016-VivekA-NeDevlopment] 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_BOMTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_BOMTableNodeOpeations"
	Dim iRowNo, sMenu, iNodeNo, iColNo, iStart, strName, aColumn
	Dim ObjRootStrcTable, objEditQuan, dicColumnMgnt, objMSMApplet
	Dim objSelectType, intNoOfObjects, iCnt, iItr
	
	Fn_MSM_BOMTableNodeOpeations=False
	'Verify RootStructures Table
	If Fn_UI_ObjectExist("Fn_MSM_BOMTableNodeOpeations",JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaTable("RootStructures"))=True Then
		Set ObjRootStrcTable=Fn_UI_ObjectCreate("Fn_MSM_BOMTableNodeOpeations", JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaTable("RootStructures"))
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		' temp solution for setting focus - Need to rework
'        JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaObject("RootStructurePanel").DblClick 0, 0,"LEFT"
'		JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaObject("RootStructurePanel").Click 0, 0,"LEFT"
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		If Instr(strAction,"_OnBelowTable") > 0 Then
			ObjRootStrcTable.SetTOProperty"Index", 1
		Else
			ObjRootStrcTable.SetTOProperty"Index", 0
		End If
		
		Select Case StrAction

			Case "Select", "Select_OnBelowTable"		'("Select"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				If strColName = "Requirement" Then
						iRowNo =Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"Requirement")
				Else
						iRowNo = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
                End if 
				If iRowNo <> -1 Then
					ObjRootStrcTable.SelectRow iRowNo
					ObjRootStrcTable.ActivateRow iRowNo
					Fn_MSM_BOMTableNodeOpeations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSM_BOMTableNodeOpeations: Succesfully Selected the Node" & strNodeName)
				End if
			
			Case "Expand", "Expand_OnBelowTable"
				If strColName = "Requirement" Then
						iRowNo =Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"Requirement")
				Else
						iRowNo = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
                End if 
				If iRowNo <> -1 Then
					ObjRootStrcTable.SelectRow iRowNo
					ObjRootStrcTable.ActivateRow iRowNo
					StrReturn = Fn_MenuOperation("Select", "View:Expand")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_MSM_BOMTableNodeOpeations] Expanded PSE BOM Table Node [" + StrNodeName + "]")							
					Fn_MSM_BOMTableNodeOpeations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSM_BOMTableNodeOpeations: Succesfully expanded of Node" & strNodeName)
				End if
			
			Case "ExpandBelow", "Expand_OnBelowTable"
				iRowNo = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
				If iRowNo <> -1 Then
					ObjRootStrcTable.SelectRow iRowNo
					ObjRootStrcTable.ActivateRow iRowNo
					StrReturn = Fn_MenuOperation("Select", "View:Expand Below")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_MSM_BOMTableNodeOpeations] Expanded PSE BOM Table Node [" + StrNodeName + "]")							
					If  JavaWindow("StructureManager").JavaWindow("WEmbeddedFrame").JavaDialog("Expand Below").Exist(15) then
							 JavaWindow("StructureManager").JavaWindow("WEmbeddedFrame").JavaDialog("Expand Below").JavaButton("Yes").Click
					End if
					Fn_MSM_BOMTableNodeOpeations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSM_BOMTableNodeOpeations: Succesfully expanded of Node" & strNodeName)
				End if
			
			Case "DeSelect"		'("DeSelect"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				iRowNo = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
				If iRowNo <> -1 Then

					ObjRootStrcTable.DeselectRow iRowNo
					Fn_MSM_BOMTableNodeOpeations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSM_BOMTableNodeOpeations: Succesfully DeSelected the Node" & strNodeName)

				End if

			Case "VerifyNode", "Exist","Exist_OnBelowTable"		'("VerifyNode"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")		
        			'Verify Node Exist
					iRowNo = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
					If iRowNo <> -1 Then

						Fn_MSM_BOMTableNodeOpeations=True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSM_BOMTableNodeOpeations: Succesfully Verified the Node" & strNodeName)

					End if
        
			Case "PopupMenuSelect","PopupMenuSelect_OnBelowTable"	'("PopupMenuSelect","","","","Trace Link:Start Trace Link")
				'Pre-requisite = Row should be selected
				strPopupMenu=Replace(strPopupMenu,":",";")
				iRowNo=ObjRootStrcTable.Object.getSelectedRow()
				If iRowNo <> -1 Then
					
					Call Fn_UI_JavaTable_CellRightClick("Fn_MSM_BOMTableNodeOpeations",JavaWindow("MultiStructManager").JavaWindow("MSWindow"),"RootStructures",iRowNo,"BOM Line","RIGHT","")
					wait 1
					sMenu = JavaWindow("MultiStructManager").WinMenu("ContextMenu").BuildMenuPath(strPopupMenu)
					wait 1
					JavaWindow("MultiStructManager").WinMenu("ContextMenu").Select sMenu
					Fn_MSM_BOMTableNodeOpeations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSM_BOMTableNodeOpeations: Succesfully Selected the PopUp Menu" & strPopupMenu)
				End if

			Case "CellVerify","CellVerify_OnBelowTable"	
				If strNodeName <> "" Then
					If instr(1,strColName, "Requirement") > 0 Then
						iRowNo =Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"Requirement")
						aColumn = Split(strColName,"~")
						strColName = aColumn(1)
					Else
						iRowNo =Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")					
                    End if 
					If iRowNo<> -1 Then
						ObjRootStrcTable.SelectRow iRowNo 
						If Lcase(cstr(ObjRootStrcTable.GetCellData( iRowNo,strColName))) = LCase(strColValue) Then
							Fn_MSM_BOMTableNodeOpeations = True
						Else
							Fn_MSM_BOMTableNodeOpeations = False
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_BOMTableNodeOpeations] Cell verified of BOM Table Node [" + strNodeName + "]")
					Else
						Fn_MSM_BOMTableNodeOpeations = False
					End If
				Else
					Fn_MSM_BOMTableNodeOpeations = False
				End If		

			Case "CellEdit"
						Fn_MSM_BOMTableNodeOpeations = False
						If strNodeName <> "" Then

							iRowNo = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
							If iRowNo <> -1 Then				
								ObjRootStrcTable.SelectRow iRowNo
								ObjRootStrcTable.ClickCell iRowNo,strColName, "LEFT" 
								wait 1
								If JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaEdit("BOMEdit_1").exist(5) Then
									JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaEdit("BOMEdit_1").Set strColValue
									JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaEdit("BOMEdit_1").Activate
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_BOMTableNodeOpeations] Cell Edited of MSM BOM Table Node [" + strNodeName + "]")
									Fn_MSM_BOMTableNodeOpeations = True
								Elseif strColName = "Quantity" then
									' case when UOM template is deployed.
									ObjRootStrcTable.DoubleClickCell iRowNo,strColName, "LEFT" 
									Set objEditQuan = JavaWindow("StructureManager").JavaWindow("Edit Quantity")
										If Fn_UI_ObjectExist("Fn_MSM_BOMTableNodeOpeations",objEditQuan) = True Then
											Call Fn_Edit_Box("Fn_MSM_BOMTableNodeOpeations",objEditQuan,"Quantity", strColValue)
											Call Fn_Button_Click("Fn_MSM_BOMTableNodeOpeations", objEditQuan, "OK")
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								End If
							End If
						End If
			Case "MultiSelect"
				   aNodeNames = split(strNodeName , "~")
				   'Clear the already selected Nodes
				   ObjRootStrcTable.Object.clearSelection
				   For iCounter = 0 to UBound(aNodeNames)
						iRowCounter = Fn_MSM_TableRowIndex(ObjRootStrcTable,trim(aNodeNames(iCounter)), "")
						If iRowCounter <> -1 Then
							ObjRootStrcTable.ExtendRow iRowCounter 
							Fn_MSM_BOMTableNodeOpeations = True
						Else
							Fn_MSM_BOMTableNodeOpeations = False
							ObjRootStrcTable.Object.clearSelection
							Exit for
						End If
					   Next
			Case "DoubleClickCell"		'("DoubleClickCell"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				iRowNo = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
				If iRowNo <> -1 Then
					ObjRootStrcTable.SelectRow iRowNo
					If strColName="" Then
						strColName=0
					End If
					wait 1
					ObjRootStrcTable.DoubleClickCell iRowNo,strColName,"LEFT"
					Fn_MSM_BOMTableNodeOpeations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSM_BOMTableNodeOpeations: Succesfully Selected the Node" & strNodeName)
				End if
			Case "AddColumns"
				ObjRootStrcTable.SelectColumnHeader "#1","RIGHT"       	
				JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaMenu("label:=Insert column\(s\) ...").Select 										       
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
'				bReturn = Fn_SISW_LoadLibrary(Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")
'				If bReturn = False  Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to load Preference Library : " & Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")	
'					Fn_MSM_BOMTableNodeOpeations = False
'					Exit Function
'				End If
				Set dicColumnMgnt = CreateObject( "Scripting.Dictionary" )
				dicColumnMgnt("Columns") = strColName
				dicColumnMgnt("CloseDialog") = True
				Fn_MSM_BOMTableNodeOpeations = Fn_SISW_RAC_Common_TableColumnManagement( "Add" , dicColumnMgnt )
				Set dicColumnMgnt = Nothing
			Case "RemoveColumns"
				ObjRootStrcTable.SelectColumnHeader "#1","RIGHT"       	
				JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaMenu("label:=Insert column\(s\) ...").Select 										       
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
'				bReturn = Fn_SISW_LoadLibrary(Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")
'				If bReturn = False  Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to load Preference Library : " & Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")	
'					Fn_MSM_BOMTableNodeOpeations = False
'					Exit Function
'				End If
				Set dicColumnMgnt = CreateObject( "Scripting.Dictionary" )
				dicColumnMgnt("Columns") = strColName
				dicColumnMgnt("CloseDialog") = True
				Fn_MSM_BOMTableNodeOpeations = Fn_SISW_RAC_Common_TableColumnManagement( "Remove" , dicColumnMgnt )
				Set dicColumnMgnt = Nothing
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			Case "LoadInViewer"
					Err.Clear
					If strNodeName <> "" Then
						iRowCounter = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
						If iRowCounter <> -1 Then
									If ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked() = False OR lCase(ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked()) = "false"  Then
										ObjRootStrcTable.Object.getNodeForRow(iRowCounter).stateIconClicked()
										 JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").SetTOProperty "title", "Viewer Confirmation"
										 
										 Set objSelectType=description.Create()
										objSelectType("Class Name").value = "JavaDialog"
										objSelectType("to_description").value = "JavaDialog"
										objSelectType("title").value = "Viewer Confirmation"
										
										If  JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").Exist(2)  Then
										 	Set  intNoOfObjects = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").ChildObjects(objSelectType)
										Else 
										   JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title", "Viewer Confirmation"
                                           Set  intNoOfObjects = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").ChildObjects(objSelectType) 
										End If 
'										Set  intNoOfObjects = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").ChildObjects(objSelectType)
										For iCnt = intNoOfObjects.count-1 to 0 Step -1
											If iCnt = 0 Then
												intNoOfObjects(iCnt).JavaButton("label:=Yes").Click
											Else
												intNoOfObjects(iCnt).JavaButton("label:=No").Click
											End If
										Next
										
										If Err.Number < 0 Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to load [ "+strNodeName+" ] in Viewer")
											 Set ObjRootStrcTable = Nothing
										Else
											Fn_MSM_BOMTableNodeOpeations = True											 
											Set ObjRootStrcTable = Nothing
											 Exit Function
										End If		
									End If					
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find the node [ "+strNodeName+" ] in the table ")
								 Set ObjRootStrcTable = Nothing
								Exit Function
							End If
						End If
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------			
				Case "UnloadFromViewer"
					If strNodeName <> "" Then
						iRowCounter = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
						If iRowCounter <> -1 Then
									If ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked() = True OR lCase(ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked()) = "true"  Then
											ObjRootStrcTable.Object.getNodeForRow(iRowCounter).stateIconClicked()
											If Err.Number < 0 Then
												 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to unload [ "+strNodeName+" ] from Viewer")
												 Set ObjRootStrcTable = Nothing
											Else
												Fn_MSM_BOMTableNodeOpeations = True											 
												Set ObjRootStrcTable = Nothing
												 Exit Function
											End If		
									End If					
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find the node [ "+strNodeName+" ] in the table ")
								 Set ObjRootStrcTable = Nothing
								Exit Function
							End If
						End If
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				Case "VerifyLoadInViewerState"
					If strNodeName <> "" Then
						iRowCounter = Fn_MSM_TableRowIndex(ObjRootStrcTable, strNodeName,"BOM Line")
						If iRowCounter <> -1 Then
								If  strColValue = False Then
									If ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked() = False OR lCase(ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked()) = "false"  Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "successfully  verified [ "+strNodeName+" ] node is not loaded in Viewer")
											Fn_MSM_BOMTableNodeOpeations = True
											Set ObjRootStrcTable = Nothing
											 Exit Function
									ElseIf ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked() = True OR lCase(ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked()) = "true" Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  verify  [ "+strNodeName+" ] node is not loaded in Viewer")
											 Set ObjRootStrcTable = Nothing
									End If
								ElseIf  strColValue = True Then
									If ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked() = True OR lCase(ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked()) = "true"  Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "successfully  verified [ "+strNodeName+" ] node is loaded in Viewer")
											Fn_MSM_BOMTableNodeOpeations = True
											 Exit Function
									ElseIf ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked() = False OR lCase(ObjRootStrcTable.Object.getNodeForRow(iRowCounter).getchecked()) = "false" Then
											 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  verify  [ "+strNodeName+" ] node is loaded in Viewer")
											 Set ObjRootStrcTable = Nothing
									End If
								End If
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find the node [ "+strNodeName+" ] in the table ")
								 Set ObjRootStrcTable = Nothing
							End If
					End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				Case "VerifyRootStructureList"					'[Tc1122:2016011300:15Feb2016:AnkitN:NeDevlopment] -- Added Case to Verify Existence of Root Structure List
					Set objMSMApplet = Fn_UI_ObjectCreate("Fn_MSM_BOMTableNodeOpeations", JavaWindow("MultiStructManager").JavaWindow("MSWindow"))
					If Fn_SISW_UI_Object_Operations("Fn_MSM_BOMTableNodeOpeations","Exist", objMSMApplet.JavaList("RootStructures"),"") = True Then
						Fn_MSM_BOMTableNodeOpeations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified Existence of Root Structure List")
						Set ObjRootStrcTable = Nothing						
					End If
					Set objMSMApplet = Nothing
					Set ObjRootStrcTable = Nothing
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------				
				Case "VerifyRootStructureListValues"			'[Tc1122:2016011300:15Feb2016:AnkitN:NeDevlopment] -- Added Case to Verify Root Structure List Values
					Set objMSMApplet = Fn_UI_ObjectCreate("Fn_MSM_BOMTableNodeOpeations", JavaWindow("MultiStructManager").JavaWindow("MSWindow"))
					If Fn_MSM_BOMTableNodeOpeations("VerifyRootStructureList", "", "", "", "") = True Then
						If instr(strNodeName, "~") > 0 Then
							For iItr = 0 To UBound(Split(strNodeName, "~"))
								Fn_MSM_BOMTableNodeOpeations = False
								For iCnt = 0 To objMSMApplet.JavaList("RootStructures").GetROProperty("items count") - 1
									If LCase(Trim(objMSMApplet.JavaList("RootStructures").GetItem(iCnt))) = LCase(Trim(Split(strNodeName, "~")(iItr))) Then
										Fn_MSM_BOMTableNodeOpeations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified Existence of Root Structure List Values ["+strNodeName+"]")
										Exit For
									End If
								Next
								If Fn_MSM_BOMTableNodeOpeations = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify Existence of Root Structure List Values ["+strNodeName+"]")
									Exit For
								End If
							Next
						Else
							For iCnt = 0 To objMSMApplet.JavaList("RootStructures").GetROProperty("items count") - 1
								If LCase(Trim(objMSMApplet.JavaList("RootStructures").GetItem(iCnt))) = LCase(Trim(strNodeName)) Then
									Fn_MSM_BOMTableNodeOpeations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified Existence of Root Structure List Values ["+strNodeName+"]")
									Exit For
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify Existence of Root Structure List Values ["+strNodeName+"]")								
								End If
							Next												
						End If					
					End If
					Set objMSMApplet = Nothing
					Set ObjRootStrcTable = Nothing
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				Case "SelectRootStructureFromList"				'[Tc1122:2016011300:15Feb2016:AnkitN:NeDevlopment] -- Added Case to Select Root Structure List Value
					Set objMSMApplet = Fn_UI_ObjectCreate("Fn_MSM_BOMTableNodeOpeations", JavaWindow("MultiStructManager").JavaWindow("MSWindow"))
					If Fn_MSM_BOMTableNodeOpeations("VerifyRootStructureListValues", strNodeName, "", "", "") = True Then
						If Fn_SISW_UI_JavaList_Operations("Fn_MSM_BOMTableNodeOpeations", "Select", objMSMApplet,"RootStructures", strNodeName, "", "") = True Then
							Fn_MSM_BOMTableNodeOpeations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Root Structure List Values ["+strNodeName+"]")
						Else
							Fn_MSM_BOMTableNodeOpeations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Root Structure List Values ["+strNodeName+"]")							
						End If 
					End If
					Set objMSMApplet = Nothing
					Set ObjRootStrcTable = Nothing					
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------								
				' returns selected BOM Line elements
				Case "GetSelected","GetSelected_OnBelowTable"                      '[TC1122-20160420-03_05_2016-VivekA-NeDevlopment] - Added Case to get selected nodes - By Madhura P
					iRows = CInt(ObjRootStrcTable.Object.getRowCount)
					sSelectedNodes = ""
					iBOMLineColIndex = Fn_MSM_TableColumnIndex(ObjRootStrcTable,"BOM Line")
					
					For iCount =0  to iRows - 1 
						If ObjRootStrcTable.Object.isRowSelected(iCount) Then
								If sSelectedNodes = "" Then
										sSelectedNodes = ObjRootStrcTable.Object.getValueAt(iCount, iBOMLineColIndex ).toString()
								Else
										sSelectedNodes = sSelectedNodes & ":" & ObjRootStrcTable.Object.getValueAt(iCount, iBOMLineColIndex ).toString()
								End If
						End if
					Next
				
					If sSelectedNodes = ""  Then
							Fn_MSM_BOMTableNodeOpeations =False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_BOMTableNodeOpeations] No BOM Table Node is selected.")
					Else
							Fn_MSM_BOMTableNodeOpeations = sSelectedNodes
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_BOMTableNodeOpeations] Returns selected PSE BOM Table Nodes [" + sSelectedNodes + "]")
					End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------								
                   Case "VerifyForegroundColour", "VerifyBackgroundColour"
							Fn_MSM_BOMTableNodeOpeations = False
							If StrNodeName <> "" Then
									StrIndex = Fn_MSM_TableRowIndex(ObjRootStrcTable,StrNodeName, "")
									'StrIndex = Fn_PSE_TableRowIndex(JavaWindow("MultiStructManager").JavaApplet("MSWindowApplet").JavaTable("RootStructures"),StrNodeName) 
							
									If cint(StrIndex) = -1  Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_MSM_BOMTableNodeOpeations] Couldnt find  PSE BOM Table Node [" + StrNodeName + "]")
											Exit function
									End If
							Else
								Exit function
							End If

							'Set  objNodeForRow =  JavaWindow("MultiStructManager").JavaApplet("MSWindowApplet").JavaTable("RootStructures").Object.getNodeForRow(cint(StrIndex))
							' if background colour
							If StrAction = "VerifyBackgroundColour" Then
									sColour = ObjRootStrcTable.Object.getCellRenderer(StrIndex,1).getBackground().toString()
							Else
							' if foreground colour
									sColour = ObjRootStrcTable.Object.getCellRenderer(StrIndex,1).getForeground().toString()
							End If

							sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
							' comparing colour codes RGB
							Select Case strColValue
								Case "DARKGREEN"
										If sColour = "[r=0,g=255,b=0]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "BLACK"
										If sColour = "[r=0,g=0,b=0]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "WHITE"
										If sColour = "[r=255,g=255,b=255]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "GRAY"
										If sColour = "[r=178,g=180,b=191]" OR sColour = "[r=160,g=182,b=192]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "DARKGRAY"
										If sColour = "[r=128,g=128,b=128]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "DARKBLUE"
										If sColour = "[r=0,g=0,b=255]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "GREEN"
										If sColour = "[r=80,g=176,b=128]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "ORANGE"
										If sColour = "[r=255,g=200,b=0]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "RED"
										If sColour = "[r=255,g=0,b=0]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case "YELLOW"
										If sColour = "[r=255,g=255,b=0]" Then
											Fn_MSM_BOMTableNodeOpeations = True
										End If
								Case Else
											Exit function
							End Select
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_BOMTableNodeOpeations] Successfully verified colour [ " & StrValue & " ] for case [" & StrAction & "]")
							Set objNodeForRow = nothing
		End Select
	Else
		'RMTable not displayed in Requirement Manager!
		Fn_MSM_BOMTableNodeOpeations=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"BOM Table is not Exist")
	End if
End Function

'-----------------------------------------------------------------------Function for Tabs in MultiStructure Manager -------------------------------------------------------------------------------------

'Function Name		:			Fn_MSM_BOMTabPanelOperation

'Description		:			This function is used to Activate, Verify Activate,  BOM Table Tab.

'Parameters			:			1.	strAction:
'											2.	strPanelName:
'											3.	strMenuName:"Close Panel" or "Split Panel"
											
'Return Value		:			True/False

'Pre-requisite		:			MSWindowApplet  window should be displayed .

'Examples			:			
									'Call Fn_MSM_BOMTabPanelOperation("Activate","(003499/A;1-y)","")
									'Fn_MSM_BOMTabPanelOperation("VerifyActivate","(003499/A;1-y)","")
									'Fn_MSM_BOMTabPanelOperation("VerifyBackground","(REQ-000023-MyTc002)","98affb")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N			 19-June-2010			1.0											                        Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh			    26-oct-2010				1.0				Added case Verify
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N			09-Apr-2012			1.1				Added case VerifyBackground							                        Sonal P
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N			18-Apr-2012			1.2				Added case PopupMenuSelect							                        Sonal P
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_BOMTabPanelOperation(strAction,strPanelName,strMenuName)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_BOMTabPanelOperation"
	on Error Resume Next
	Dim objName
	Fn_MSM_BOMTabPanelOperation=False	
	If Fn_UI_ObjectExist("Fn_MSM_BOMTabPanelOperation",JavaWindow("MultiStructManager").JavaWindow("MSWindow"))=True Then
		
		JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("PanelName").SetTOProperty "label",strPanelName
		JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("PanelName").Highlight
	Select Case strAction		
		Case "Activate"			'('"Activate","(000772-MyTc002)","")
            'Call Fn_UI_JavaStaticText_Click("Fn_MSM_BOMTabPanelOperation",JavaWindow("MultiStructManager").JavaWindow("MSWindow"),"PanelName", 5,5, "LEFT")
			JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("PanelName").Click 1,1,"LEFT"
			Fn_MSM_BOMTabPanelOperation=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Succesfully Activated" & strPanelName &" Panel")

		Case "VerifyActivate"	'("VerifyActivate","(REQ-000023-MyTc002)","")
			Set objName = JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("PanelName").Object			
			Fn_MSM_BOMTabPanelOperation=objName.selected()

			If Fn_MSM_BOMTabPanelOperation="true" Then
				Fn_MSM_BOMTabPanelOperation=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Succesfully Verify" & strPanelName &" Panel is Activated")					
			End If
			Set objName = Nothing
		Case "Verify"
			Set objName = JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("PanelName")			
			If Fn_UI_ObjectExist("Fn_MSM_BOMTabPanelOperation", objName) = True Then
				Fn_MSM_BOMTabPanelOperation = True
			Else
				Fn_MSM_BOMTabPanelOperation = False
			End If
			Set objName = Nothing

		Case "VerifyBackground"	'("VerifyBackground","(REQ-000023-MyTc002)","98affb")
			Set objName = JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("PanelName")	
			'In this case use parameter : = strMenuName to pass the background number
			If objName.GetROProperty("background")=strMenuName Then
				Fn_MSM_BOMTabPanelOperation=True
			End If
			Set objName = Nothing
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PopupMenuSelect"
			If JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("PanelName").Exist(5) Then
				JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("PanelName").Click 5,5,"RIGHT"
				wait 1
				aMenu = split(strMenuName,":",-1,1)
				Select Case Ubound(aMenu)
						Case "0"
							JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaMenu("label:="&aMenu(0)&"","index:=0").Select
						Case else
							Fn_MSM_BOMTabPanelOperation = False
					End Select
					If Err.number < 0 Then
						'Menu Selection Failed
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MSM_BOMTabPanelOperation] Failed to Select Popup Menu ["+ strMenuName +"]")
						Fn_MSM_BOMTabPanelOperation = False			
					Else
						'Menu Selection Successful
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_MSM_BOMTabPanelOperation] Popup Menu ["+ strMenuName +"]Selected Sucessfully")
						Fn_MSM_BOMTabPanelOperation = True
					End If	
			Else
				Fn_MSM_BOMTabPanelOperation = False
			End If

		Case Else
			Fn_MSM_BOMTabPanelOperation=False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Action name"& strAction &" Pass to the Function")
	End Select
	Else
		Fn_MSM_BOMTabPanelOperation=False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"MSWindowApplet Window is not exist")
	End If
End Function

'*********************************************************		Function to Get Table Column Index in Multi-Structure Manager		***********************************************************************

'Function Name		        :				Fn_MSM_TableColumnIndex

'Description			  :		 		This function is used to get the Table Column Index

'Parameters			 :	 			1. objTable : Object of Java Table
'									2. sCol : Name of the Column
											
'Return Value		   	: 					Node index / -1

'Pre-requisite			:				 Multi-Structure Manager window should be displayed .

'Examples				:				Fn_MSM_TableColumnIndex(objTable, "BOM Line")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh				 26-Oct-2010			 	1.0			Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MSM_TableColumnIndex(objTable, sCol)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_TableColumnIndex"
	Dim iCounter, iCols
	iCols = cInt(objTable.GetROProperty("cols"))
	Fn_MSM_TableColumnIndex = -1
	If Fn_UI_ObjectExist("Fn_MSM_TableColumnIndex",objTable) <> False Then
		For iCounter = 0 to iCols -1
			
			If objTable.GetColumnName(iCounter) = sCol Then
				Fn_MSM_TableColumnIndex = iCounter
				Exit for
			End If
		Next
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MSM_TableColumnIndex ] executed successfully.")	
End Function


'*********************************************************		Function to Get BOM Table Node operation in MultiStructure Manager		***********************************************************************

'Function Name		:				Fn_MSM_LineTableNodeOperation

'Description			 :		 		This function is used to get the Multu-Structure Manager Table Node operation.

'Parameters			   :				1.	strAction = "Select"
'										2. StrNodeName:Name of the Node. 
'										3. strColName
'										4. strColValue
'										5. strPopupMenu
				
											
'Return Value		   : 				True / False / TableNodeIndex

'Pre-requisite			:				MultiStructure Manager window should be displayed .


'Examples				:			Fn_MSM_LineTableNodeOperation("Select","003465/A;1-Test (View):003498/A;1-Tes","","","")
'							        Fn_MSM_LineTableNodeOperation("Expand","003465/A;1-Test (View):003498/A;1-Tes","","","")
'									Fn_MSM_LineTableNodeOperation("Exist","003465/A;1-Test (View):003498/A;1-Tes","","","")
'									Fn_MSM_LineTableNodeOperation("AddColumns","","Relation~Status","","")
'									Fn_MSM_LineTableNodeOperation("RemoveColumns","","Relation","","")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh 				25-Oct-2010			1.0				Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh 				16-Mar-2011			1.0				Modified case Expand
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh 				16-Mar-2011			1.0				Added case Exist
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh			19-Jul-2013			1.2		Added case AddColumns, RemoveColumns
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'	Naveen				10-Sept-2012			1.2		Added case MultiSelect
Public function Fn_MSM_LineTableNodeOperation(sAction, sNodeName,sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_LineTableNodeOperation"
	Dim objTable, iRowCounter, objApplet, bReturn, iRows, dicColumnMgnt
	Set objApplet = JavaWindow("MultiStructManager").JavaWindow("MSWindow")
	Set objTable = objApplet.JavaTable("LineTable")
	If  Fn_UI_ObjectExist("Fn_MSM_LineTableNodeOperation", objTable) = False Then
                    Fn_MSM_LineTableNodeOperation = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_MSM_LineTableNodeOperation ] Line Table does not exist.")
		Set objTable = nothing
		Exit function
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	' temp solution for setting focus - Need to rework
'	JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaObject("LineTablePanel").DblClick 0,0,"LEFT"
'	JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaObject("LineTablePanel").Click 0,0,"LEFT"

	'Commented above code, since its Toolkit property has changed on Build - 111  -- Amit T.

	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		
	Select Case sAction
		Case "Select"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Line")
			Fn_MSM_LineTableNodeOperation = False
			If iRowCounter <> -1 Then
				Call Fn_UI_JavaTable_SelectRow("Fn_MSM_LineTableNodeOperation",objApplet, "LineTable", iRowCounter)
				Fn_MSM_LineTableNodeOperation = true 
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Exist"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Line")
			Fn_MSM_LineTableNodeOperation = False
			If iRowCounter <> -1 Then
				Fn_MSM_LineTableNodeOperation = true 
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Expand"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Line")
			Fn_MSM_LineTableNodeOperation = False
			If iRowCounter <> -1 Then
				'Call Fn_UI_JavaTable_SelectRow("Fn_MSM_LineTableNodeOperation",objApplet, "LineTable", iRowCounter)
				'Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_LineTableNodeOperation] Selected  Line Table Node [" + sNodeName + "]")
				'	If Fn_MenuOperation("Select", "View:Expand") = True Then
				'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_LineTableNodeOperation] Expanded Line Table Node [" + sNodeName + "]")
				'		Fn_MSM_LineTableNodeOperation = True
				'	Else							
				'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_MSM_LineTableNodeOperation] Failed to Expanded Line Table Node [" + sNodeName + "]")
				'		Fn_MSM_LineTableNodeOperation = False
				'	End If	
				objTable.Object.expandNode objTable.Object.getNodeForRow(cint(iRowCounter))					
				Fn_MSM_LineTableNodeOperation = True
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
			
				Case "PopupMenuSelect"	'("PopupMenuSelect","","","","Trace Link:Start Trace Link")
				'Pre-requisite = Row should be selected
				sPopupMenu=Replace(sPopupMenu,":",";")
				iRowNo=objTable.Object.getSelectedRow()
				If iRowNo <> -1 Then
					
					Call Fn_UI_JavaTable_CellRightClick("Fn_MSM_BOMTableNodeOpeations",objApplet,"LineTable",iRowNo,"Line","RIGHT","")
					wait 1
					sMenu = JavaWindow("MultiStructManager").WinMenu("ContextMenu").BuildMenuPath(sPopupMenu)
					wait 1
					JavaWindow("MultiStructManager").WinMenu("ContextMenu").Select sMenu
					Fn_MSM_LineTableNodeOperation=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_MSM_LineTableNodeOperation: Succesfully Selected the PopUp Menu" & sPopupMenu)
				End if
			
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
			Case "AddColumns"
				objTable.SelectColumnHeader "#1","RIGHT"       	
				JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaMenu("label:=Insert column\(s\) ...").Select 										       
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
'				bReturn = Fn_SISW_LoadLibrary(Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")
'				If bReturn = False  Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to load Preference Library : " & Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")	
'					Fn_MSM_LineTableNodeOperation = False
'					Exit Function
'				End If
				Set dicColumnMgnt = CreateObject( "Scripting.Dictionary" )
				dicColumnMgnt("Columns") = sColName
				dicColumnMgnt("CloseDialog") = True
				Fn_MSM_LineTableNodeOperation = Fn_SISW_RAC_Common_TableColumnManagement( "Add" , dicColumnMgnt )
				Set dicColumnMgnt = Nothing

			Case "RemoveColumns"
				objTable.SelectColumnHeader "#1","RIGHT"       	
				JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaMenu("label:=Insert column\(s\) ...").Select
				wait 2						       
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
'				bReturn = Fn_SISW_LoadLibrary(Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")
'				If bReturn = False  Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to load Preference Library : " & Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")	
'					Fn_MSM_LineTableNodeOperation = False
'					Exit Function
'				End If
				Set dicColumnMgnt = CreateObject( "Scripting.Dictionary" )
				dicColumnMgnt("Columns") = sColName
				dicColumnMgnt("CloseDialog") = True
				Fn_MSM_LineTableNodeOperation = Fn_SISW_RAC_Common_TableColumnManagement("Remove" , dicColumnMgnt )
				Set dicColumnMgnt = Nothing
		Case Else
			Fn_MSM_LineTableNodeOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Function [ Fn_MSM_LineTableNodeOperation ] execution failed due to invalid case [ "& sAction &" ]")
			Set objTable = nothing
			Set objApplet = nothing
			Exit function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function [ Fn_MSM_LineTableNodeOperation ] Successfully executed with case [ "& sAction &" ]")
	Set objTable = nothing
	Set objApplet = nothing
End Function

'*********************************************************		Function to Get BOM Table Node operation in MultiStructure Manager		***********************************************************************

'Function Name		:				Fn_MSM_AttachmentsTableNodeOpration

'Description			 :		 		This function is used to get the Multu-Structure Manager Table Node operation.

'Parameters			   :					1. sAction = "Select"
'										2. sNodeName:Name of the Node. 
'										3. sColName
'										4. sValue
'										5. sPopupMenu
				
											
'Return Value		   : 				True / False / TableNodeIndex

'Pre-requisite			:				MultiStructure Manager window should be displayed .


'Examples				:			Fn_MSM_AttachmentsTableNodeOpration("Select","001682/A;1-d1:Representation For","","","")
'							         Fn_MSM_AttachmentsTableNodeOpration("CellDoubleClick","001682/A;1-d1:Representation For","Line","","")
											
' NOTE: 	for second Attachments table set TO Property of AttachmentsTreeTablePanel to 2
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Koustubh 				26-Oct-2010			1.0				Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_MSM_AttachmentsTableNodeOpration(sAction, sNodeName,sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_AttachmentsTableNodeOpration"
	Dim objTable, iRowCounter, objApplet, bReturn, iRows
	Set objApplet = JavaWindow("MultiStructManager").JavaWindow("MSWindow")
	Set objTable = objApplet.JavaTable("AttachmentTable")
	If  Fn_UI_ObjectExist("Fn_MSM_AttachmentsTableNodeOpration", objTable) = False Then
		Call Fn_ToolbatButtonClick("Show/Hide the data panel")
		objApplet.JavaStaticText("DataPanelHeader").SetTOProperty "label","Attachments"
		objApplet.JavaStaticText("DataPanelHeader").Click 0,0,"LEFT"
		wait 3
		If  Fn_UI_ObjectExist("Fn_MSM_AttachmentsTableNodeOpration", objTable) = False Then
			Fn_MSM_AttachmentsTableNodeOpration = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_MSM_AttachmentsTableNodeOpration ] Line Table does not exist.")
			Set objTable = nothing
			Exit function
		End If
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	' temp solution for setting focus - Need to rework
	JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaObject("AttachmentsTreeTablePanel").DblClick 0,0,"LEFT"
	JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaObject("AttachmentsTreeTablePanel").Click 0,0,"LEFT"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

	Select Case sAction
		Case "Select"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Line")
			Fn_MSM_AttachmentsTableNodeOpration = False
			If iRowCounter <> -1 Then
				Call Fn_UI_JavaTable_SelectRow("Fn_MSM_AttachmentsTableNodeOpration",objApplet, "AttachmentTable", iRowCounter)
				Fn_MSM_AttachmentsTableNodeOpration = true 
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Expand"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Line")
			Fn_MSM_AttachmentsTableNodeOpration = False
			If iRowCounter <> -1 Then
				Call Fn_UI_JavaTable_SelectRow("Fn_MSM_AttachmentsTableNodeOpration",objApplet, "AttachmentTable", iRowCounter)
				objTable.Object.expandNode(objTable.Object.getNodeForRow(iRowCounter))
				Fn_MSM_AttachmentsTableNodeOpration = true 
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "CellDoubleClick"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Line")
			Fn_MSM_AttachmentsTableNodeOpration = False
			If iRowCounter <> -1 Then
				Call Fn_UI_JavaTable_SelectRow("Fn_MSM_AttachmentsTableNodeOpration",objApplet, "AttachmentTable", iRowCounter)
				wait 1
				objTable.DoubleClickCell iRowCounter, sColName
				Fn_MSM_AttachmentsTableNodeOpration = true 
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "CellVerify"
				iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Line")
				iColIndex = Fn_MSM_TableColumnIndex(objTable, sColName)
				Fn_MSM_AttachmentsTableNodeOpration = False
				If iRowCounter <> -1 Then
					If trim(objTable.object.getValueAt(iRowCounter, iColIndex).toString()) = sValue Then
						Fn_MSM_AttachmentsTableNodeOpration = true 
					End If
				End If 
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case Else
			Fn_MSM_AttachmentsTableNodeOpration = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_MSM_AttachmentsTableNodeOpration ] execution failed due to invalid case [ "& sAction &" ]")
			Set objTable = nothing
			Set objApplet = nothing
			Exit function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function [ Fn_MSM_AttachmentsTableNodeOpration ] Successfully executed with case [ "& sAction &" ]")
	Set objTable = nothing
	Set objApplet = nothing
End Function
'#############################################################################################################
'###    FUNCTION NAME   :   Fn_MSM_OccurrenceGroup(sOccName, sOccDesc, sOccType, sButtons)
'###
'###    DESCRIPTION        :   This function is used to create a new form.
'###
'###    PARAMETERS      :   1.sOccName: Valid Occurrence Name.
'###			    		                     2.sOccDesc: Valid Occurrence Description.
'###			    		                    3.sOccType: Valid Occurrence Type.
'###			   			                   4.sButtons.
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           23/02/2011         1.0
'###
'###    EXAMPLE          : Msgbox Fn_MSM_OccurrenceGroup("TestOcc", "Testing", "OccurrenceGroup", "OK")
'#############################################################################################################
Public Function Fn_MSM_OccurrenceGroup(sOccName, sOccDesc, sOccType, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_OccurrenceGroup"
	Dim objNewOcc, objSelectType, intNoOfObjects, iCounter, breturn, aButtons, iCount
	Fn_MSM_OccurrenceGroup = False
	
	breturn = Fn_UI_ObjectExist("Fn_MSM_OccurrenceGroup", JavaWindow("MultiStructManager").JavaWindow("TcDefaultApplet").JavaDialog("NewOccurrenceGroup"))	
	If breturn = false Then
		breturn = Fn_UI_ObjectExist("Fn_MSM_OccurrenceGroup",JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaDialog("NewOccurrenceGroup"))
	End If
	
    If breturn = false Then
		Call Fn_MenuOperation("Select","File:New:Occurrence Group...")
	End If
	Set objNewOcc = JavaWindow("MultiStructManager").JavaWindow("TcDefaultApplet").JavaDialog("NewOccurrenceGroup")
	if Fn_UI_ObjectExist("Fn_MSM_OccurrenceGroup", objNewOcc) = False Then
		Set objNewOcc = Fn_UI_ObjectCreate("Fn_MSM_OccurrenceGroup",JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaDialog("NewOccurrenceGroup"))
    Else
     	Set objNewOcc = Fn_UI_ObjectCreate("Fn_MSM_OccurrenceGroup",JavaWindow("MultiStructManager").JavaWindow("TcDefaultApplet").JavaDialog("NewOccurrenceGroup"))
    End if
    
   'Set Name
    call Fn_Edit_Box("Fn_MSM_OccurrenceGroup",objNewOcc,"Name",sOccName)
    'Set description
	call Fn_Edit_Box("Fn_MSM_OccurrenceGroup",objNewOcc,"Description",sOccDesc)
	' Click on more button
	call Fn_CheckBox_Set("Fn_MSM_OccurrenceGroup", objNewOcc,"More","ON")
	Call Fn_ReadyStatusSync(2)
   'Set Form Type
	'/////////////////////////////////////////////
	Set objSelectType=description.Create()
	objSelectType("Class Name").value = "JavaStaticText"
    Set  intNoOfObjects = objNewOcc.ChildObjects(objSelectType)
	For  iCounter = 0 to intNoOfObjects.count-1
		If  intNoOfObjects(iCounter).getROProperty("label") = sOccType Then
			intNoOfObjects(iCounter).Click 1,1
			wait 3
			Exit for
		End If
		'   wait 1
	Next
	'Click on Buttons
	If sButtons<>"" Then
		aButtons = split(sButtons, ":",-1,1)
		For iCount=0 to Ubound(aButtons)
			'Click on Add Button
			Call Fn_Button_Click("Fn_MSM_OccurrenceGroup", objNewOcc, aButtons(iCount))
		Next
	End If
	Fn_MSM_OccurrenceGroup = True
	Set objNewOcc = nothing 
	Set objSelectType = nothing
	Set intNoOfObjects = nothing
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_MSM_VariantConditionOperations

'Description			 :	This function is used to Perform operation on Variant Condition

'Parameters			   :   1.strAction : Action to perform "Add"
'										2. strLogicalOpt:Logical Option (AND , OR) 
'										3. strItem : Item Name
'										4. strOption : Option
'										5. strCondition : Condition
'										6. strValue : Values
				
											
'Return Value		   : 				True / False

'Pre-requisite			:				MultiStructure Manager window should be displayed .


'Examples				:	Fn_MSM_VariantConditionOperations("Add","AND","Computing","Test","!=","One")
'							         	Fn_MSM_VariantConditionOperations("Add","OR","Computer","Test","=","Two")
'									
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep	 					28-Feb-2011			1.0								Created						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_VariantConditionOperations(strAction,strLogicalOpt,strItem,strOption,strCondition,strValue)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_VariantConditionOperations"
   Dim ObjVarCondDialog
   Fn_MSM_VariantConditionOperations=False
	Set ObjVarCondDialog=JavaWindow("MultiStructManager").JavaWindow("TcDefaultApplet").JavaDialog("VariantCondition")
	If Not ObjVarCondDialog.Exist(6) Then
		Call Fn_MenuOperation("Select","Edit:Variant Condition...")
	End If

	Call Fn_Button_Click("Fn_MSM_VariantConditionOperations",ObjVarCondDialog,"LegacySwitch")
	Select Case strAction
		Case "Add"
			If strLogicalOpt<>"" Then
				ObjVarCondDialog.JavaRadioButton("Condition").SetTOProperty "attached text",strLogicalOpt
				Call Fn_UI_JavaRadioButton_SetON("Fn_MSM_VariantConditionOperations",ObjVarCondDialog, "Condition")
			End If
			If strItem<>"" Then
				Call Fn_UI_EditBox_Type("Fn_MSM_VariantConditionOperations",ObjVarCondDialog,"Item",strItem)
			End If
			If strOption<>"" Then
				Call Fn_UI_EditBox_Type("Fn_MSM_VariantConditionOperations",ObjVarCondDialog,"Option",strOption)
			End If
			If strCondition<>"" Then
				Call Fn_Edit_Box("Fn_MSM_VariantConditionOperations",ObjVarCondDialog,"Condition",strCondition)
			End If
			If strValue<>"" Then
				Call Fn_CheckBox_Set("Fn_MSM_VariantConditionOperations", ObjVarCondDialog, "lov_19", "ON")
				Call Fn_List_Select("Fn_MSM_VariantConditionOperations", ObjVarCondDialog, "ValueList",strValue)
			End If
			Call Fn_Button_Click("Fn_MSM_VariantConditionOperations",ObjVarCondDialog,"Append")
			Fn_MSM_VariantConditionOperations=True
	End Select
	Call Fn_Button_Click("Fn_MSM_VariantConditionOperations",ObjVarCondDialog,"OK")
	Set ObjVarCondDialog=Nothing
End Function 

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_MSM_VariantCreate

'Description			 :This function use to create variant

'Parameters			   : 1.  strName: Variant Name
'									  2. strDescription: Variant Description
'									  3. strValues: Values
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:	Multi Structure Manager window should be displayed with DataPanle loaded.

'Examples				:	Fn_MSM_VariantCreate("Test","","One")
'										Fn_MSM_VariantCreate("Test","","A:B:C:D")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep							28-Feb-2011		1.0							Created							Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'strValues : Values could be ( : ) colon separated 
Public Function Fn_MSM_VariantCreate(strName,strDescription,strValues)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_VariantCreate"
   Dim ObjMSMApp
   Dim arrValues,iCounter
   Fn_MSM_VariantCreate=False
	Set ObjMSMApp=JavaWindow("MultiStructManager").JavaWindow("MSWindow")
	If strName<>"" Then
		Call Fn_Edit_Box("Fn_MSM_VariantCreate",ObjMSMApp,"VariantName",strName)
	End If
	If strDescription<>"" Then
		Call Fn_Edit_Box("Fn_MSM_VariantCreate",ObjMSMApp,"VariantDescription",strDescription)
	End If
	If strValues<>"" Then
		arrValues=Split(strValues,":")
		For iCounter=0 To Ubound(arrValues)
			Call Fn_Edit_Box("Fn_MSM_VariantCreate",ObjMSMApp,"VariantValue",arrValues(iCounter))
			Call Fn_Button_Click("Fn_MSM_VariantCreate",ObjMSMApp,"AddVariant")
		Next
	End If
	Call Fn_Button_Click("Fn_MSM_VariantCreate",ObjMSMApp,"CreateVariant")
	Fn_MSM_VariantCreate=True
End Function 

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_MSM_CreateAllocationMap

'Description			 :This function use to create Allocation Map

'Parameters			   : 1.  strID: Allocation Map ID
'									  2.  strRev: Allocation Map Revision
'									  3.  strName: Allocation Map Name
'									  4. strDescription : Allocation Map Description
'									  5. UOM : Unit Of Measures
											
'Return Value		   : ID-Revision \ FALSE

'Pre-requisite			:	Multi Structure Manager window should be displayed.

'Examples				:	Fn_MSM_CreateAllocationMap("","","Test2","TestAllocationMap1","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep							28-Feb-2011		1.0							Created							Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_CreateAllocationMap(strID,strRev,strName,strDescription,UOM)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_CreateAllocationMap"
   Dim ObjAllocMapDialog
   Dim strMapID,strMapRev
	Fn_MSM_CreateAllocationMap=False
'	Set ObjAllocMapDialog=JavaWindow("MultiStructManager").JavaWindow("TcDefaultApplet").JavaDialog("NewAllocationMap")
'		Changing  Hierarchy as per TC10.0   -  Sonal
	Set ObjAllocMapDialog=Window("MSMWindow").JavaDialog("NewAllocationMap")

	If Not ObjAllocMapDialog.Exist(6) Then
		Call Fn_Button_Click("Fn_MSM_CreateAllocationMap", JavaWindow("MultiStructManager").JavaWindow("MSWindow"), "CreateAllocationContext")
	End If
	If JavaDialog("AllocationContextReviseConfr").Exist(6) Then
		Call Fn_Button_Click("Fn_MSM_CreateAllocationMap",JavaDialog("AllocationContextReviseConfr"),"Yes")
	End If
	Call Fn_List_Select("Fn_MSM_CreateAllocationMap", ObjAllocMapDialog, "MapList","Allocation Map")
	Call Fn_Button_Click("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog, "Next")
	If strID<>"" Then
		Call Fn_Edit_Box("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog,"ID",strID)
	End If
	If strRev<>"" Then
		Call Fn_Edit_Box("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog,"Revision",strRev)
	End If
	If strID="" And strRev="" Then
		Call Fn_Button_Click("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog, "Assign")
	End If
	strMapID=Fn_Edit_Box_GetValue("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog,"ID")
	strMapRev=Fn_Edit_Box_GetValue("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog,"Revision")
	If strName<>"" Then
		Call Fn_Edit_Box("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog,"Name",strName)
	End If
	If strDescription<>"" Then
		Call Fn_Edit_Box("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog,"Description",strDescription)
	End If
	Call Fn_Button_Click("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog, "Finish")
	Call Fn_Button_Click("Fn_MSM_CreateAllocationMap",ObjAllocMapDialog, "Close")

	Fn_MSM_CreateAllocationMap="'"&strMapID+"-"+strMapRev
	Set ObjAllocMapDialog=Nothing
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_MSM_CreateAllocationContext

'Description			 :This function use to create Allocation Context (Revise Allocation Context)

'Parameters			   : 1.  strName: Allocation Context Name
'									  2.  strDescription:Allocation Context Description
'									  3. UOM : Unit Of Measures
											
'Return Value		   : ID-Revision \ FALSE

'Pre-requisite			:	Multi Structure Manager window should be displayed.

'Examples				:	Fn_MSM_CreateAllocationContext("NewTest","NewDescription","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep							28-Feb-2011		1.0							Created							Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_CreateAllocationContext(strName,strDescription,UOM)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_CreateAllocationContext"
   Dim ObjAllocCntxtDialog
   Dim strMapID,strMapRev
	Fn_MSM_CreateAllocationContext=False
'	Set ObjAllocCntxtDialog=JavaWindow("MultiStructManager").JavaWindow("TcDefaultApplet").JavaDialog("NewAllocationContext")
'	Changing Object Hierarchy as per TC10.0  - Sonal
	Set ObjAllocCntxtDialog=Window("MSMWindow").JavaDialog("NewAllocationContext")
	If Not ObjAllocCntxtDialog.Exist(6) Then
		Call Fn_Button_Click("Fn_MSM_CreateAllocationContext", JavaWindow("MultiStructManager").JavaWindow("MSWindow"), "ReviseAllocationContext")
	End If
	If JavaDialog("AllocationContextReviseConfr").Exist(6) Then
		Call Fn_Button_Click("Fn_MSM_CreateAllocationContext",JavaDialog("AllocationContextReviseConfr"),"Yes")
	End If
    strMapID=Fn_Edit_Box_GetValue("Fn_MSM_CreateAllocationContext",ObjAllocCntxtDialog,"ID")
	strMapRev=Fn_Edit_Box_GetValue("Fn_MSM_CreateAllocationContext",ObjAllocCntxtDialog,"Revision")
	If strName<>"" Then
		Call Fn_Edit_Box("Fn_MSM_CreateAllocationContext",ObjAllocCntxtDialog,"Name",strName)
	End If
	If strDescription<>"" Then
		Call Fn_Edit_Box("Fn_MSM_CreateAllocationContext",ObjAllocCntxtDialog,"Description",strDescription)
	End If
	Call Fn_Button_Click("Fn_MSM_CreateAllocationContext",ObjAllocCntxtDialog, "Finish")
	Call Fn_Button_Click("Fn_MSM_CreateAllocationContext",ObjAllocCntxtDialog, "Close")

	Fn_MSM_CreateAllocationContext=strMapID+"-"+strMapRev
	Set ObjAllocCntxtDialog=Nothing
End Function 

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_MSM_ErrorWindowVerify

'Description			 :This function use to Handle Error Window And to verify Error Msg

'Parameters			   : 1.  strWindowName: Error Window Name
'									  2.  strControlName:Control Name on which controls Error Msg need to Verify . JavaStaticText , JavaEditBox
'									  3. strErrMsg : Error Message
'									  4. strButton : Button Name
											
'Return Value		   : TURE \ FALSE

'Pre-requisite			:	Error Window should be appear on screen

'Examples				:	Fn_MSM_ErrorWindowVerify("Allocate error","StaticText","Error while Allocating","OK")
'										Fn_MSM_ErrorWindowVerify("Allocate error","EditBox","Invalid Tag  - the requested object does not exist","OK")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep							28-Feb-2011		1.0							Created							Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_ErrorWindowVerify(strWindowName,strControlName,strErrMsg,strButton)

  Dim dicErrorInfo
  Set dicErrorInfo = CreateObject("Scripting.Dictionary")
  dicErrorInfo.Add "Action", "ErrorWindowVerify"
  dicErrorInfo.Add "Title", strWindowName
  dicErrorInfo.Add "ControlName", strControlName
  dicErrorInfo.Add "Message", strErrMsg
  dicErrorInfo.Add "Button", strButton    
  Fn_MSM_ErrorWindowVerify = Fn_SISW_MSM_ErrorVerify(dicErrorInfo)
  Set dicErrorInfo = Nothing

End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_MSM_CreateAllocation

'Description			 :This function use to Create New Allocation

'Parameters			   : 1.  strName: Allocation Name
'									  2.  strReason:Allocation Reason
											
'Return Value		   : TURE \ FALSE

'Pre-requisite			:	New Allocation Dialog Should be Appear On Screen

'Examples				:	Fn_MSM_CreateAllocation("","")
'										Fn_MSM_CreateAllocation("TestAllocation","Allocation For New Rule")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep							28-Feb-2011		1.0							Created							Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_CreateAllocation(strName,strReason)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_CreateAllocation"
   Dim ObjAllocationDialog
	Fn_MSM_CreateAllocation=False
'	Set ObjAllocationDialog=JavaWindow("MultiStructManager").JavaWindow("TcDefaultApplet").JavaDialog("NewAllocation")
'	Changing Hierarchy as per TC10.0    - Sonal	
	Set ObjAllocationDialog=JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaDialog("NewAllocation")

	If strName<>"" Then
		Call Fn_Edit_Box("Fn_MSM_CreateAllocation",ObjAllocationDialog,"Name",strName)
	End If
	If strReason<>"" Then
		Call Fn_Edit_Box("Fn_MSM_CreateAllocation",ObjAllocationDialog,"Reason",strReason)
	End If
	Call Fn_Button_Click("Fn_MSM_CreateAllocation", ObjAllocationDialog, "OK")
	Fn_MSM_CreateAllocation=True
	If ObjAllocationDialog.JavaDialog("AllocateError").Exist(7) Then
		Call Fn_Button_Click("Fn_MSM_CreateAllocation", ObjAllocationDialog.JavaDialog("AllocateError"), "OK")
		Call Fn_Button_Click("Fn_MSM_CreateAllocation", ObjAllocationDialog, "Cancel")
		Fn_MSM_CreateAllocation=False
	End If
	Set ObjAllocationDialog=Nothing
End Function
'******************
'*********************************************************		Function to perform operations on Set View Configuration to Context dialog ***********************************************************************

'Function Name		    :		Fn_MSM_SetViewConfigurationToContextOperations

'Description			:	    This function is used to perform operations on Set View Configuration to Context 

'Parameters				:	1. sAction : Action need to perform.
'                                       			1. sSelectFrom : Home / Clipboard / OpenStructureContextByName / OpenConfigurationByName / OpenCCByName
'							3. sConfigNamePath : Name of configation context
'							4. sSearchCriteria - For future use										
											
'Return Value		   	: 	True/False

'Pre-requisite			:	   

' Note : Case Verify and Options OpenStructureContextByName / OpenConfigurationByName / OpenCCByName are not implemented yet.
'
'Examples				:	    Call Fn_MSM_SetViewConfigurationToContextOperations("Select","Home", "Home:Newstuff:config1_45786", "")
'						        Call Fn_MSM_SetViewConfigurationToContextOperations("Select", "Clipboard", "Clipboard:conf1", "")

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W			19-Apr-2011	           	  1.0			      Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Snehal Salunkhe		18-May-2012	           	  1.0			      Modified object hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MSM_SetViewConfigurationToContextOperations(sAction, sSelectFrom, sConfigNamePath, sSearchCriteria)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_SetViewConfigurationToContextOperations"
	Dim objSetConfigCon, aNamePath, sPath, iCount
	Fn_MSM_SetViewConfigurationToContextOperations = False
	Set objSetConfigCon = JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaDialog("SetViewConfiguration")
	sPath = ""
	' checking existance of dialog window
	If objSetConfigCon.Exist(5) = False Then
		Call  Fn_MenuOperation("Select", "Tools:Set View Configuration...")
	End If
	If objSetConfigCon.Exist(15) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MSM_SetViewConfigurationToContextOperations ] Fail : Failed to find dialog window [ " & "Set View Configuration to Context" &  " ] ")
		Set objSetConfigCon = nothing
		Exit function
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Select"
			Select Case sSelectFrom
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "Home"
					' activating Home tab
					call Fn_UI_JavaRadioButton_SetON("Fn_MSM_SetViewConfigurationToContextOperations",objSetConfigCon,"Home")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MSM_SetViewConfigurationToContextOperations ] Passed : successfully activated TAb [ Home ] ")
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "Clipboard"
					' activating Clipboard tab
					call Fn_UI_JavaRadioButton_SetON("Fn_MSM_SetViewConfigurationToContextOperations",objSetConfigCon,"Clipboard")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MSM_SetViewConfigurationToContextOperations ] Passed : successfully activated TAb [ Home ] ")
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "OpenStructureContextByName"
					' for future use
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "OpenConfigurationByName"
					' for future use
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "OpenCCByName"
					' for future use
				Case else
					Exit function
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			End Select
			aNamePath = split(sConfigNamePath,":")
			' expanding tree path
			For iCount = 0 to UBound(aNamePath) - 1
				If sPath = "" Then
					sPath = aNamePath(iCount)
				else
					sPath = sPath & ":" & aNamePath(iCount)
				End If
				Wait(3)
				Call Fn_UI_JavaTree_Expand("Fn_MSM_SetViewConfigurationToContextOperations",objSetConfigCon,"CCTree",sPath)
			Next
			Wait(2)
			' selecting tree path
			Call Fn_JavaTree_Select("Fn_MSM_SetViewConfigurationToContextOperations",objSetConfigCon,"CCTree",sConfigNamePath)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MSM_SetViewConfigurationToContextOperations ] Passed : successfully tree path [ " & sConfigNamePath &  " ] ")

			' clicking on OK
			Call Fn_Button_Click("Fn_MSM_SetViewConfigurationToContextOperations",objSetConfigCon, "OK")
			Fn_MSM_SetViewConfigurationToContextOperations = true
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Verify"
			' for future use
			Select Case sSelectFrom
				Case "Home"
					' activating Home tab
					call Fn_UI_JavaRadioButton_SetON("Fn_MSM_SetViewConfigurationToContextOperations",objSetConfigCon,"Home")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MSM_SetViewConfigurationToContextOperations ] Passed : successfully activated TAb [ Home ] ")
				Case "Clipboard"
					' activating Clipboard tab
					call Fn_UI_JavaRadioButton_SetON("Fn_MSM_SetViewConfigurationToContextOperations",objSetConfigCon,"Clipboard")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MSM_SetViewConfigurationToContextOperations ] Passed : successfully activated TAb [ Home ] ")
				Case "OpenStructureContextByName"
					' for future use
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "OpenConfigurationByName"
					' for future use
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "OpenCCByName"
					' for future use
				Case else
					Exit function
			End Select
			' clicking on Cancel
			Call Fn_Button_Click("Fn_MSM_SetViewConfigurationToContextOperations",objSetConfigCon, "Cancel")
			Fn_MSM_SetViewConfigurationToContextOperations = true
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case Else
                              Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MSM_SetViewConfigurationToContextOperations ] Failed : Invalid case [ " & sAction &  " ] ")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	End Select
	Set objSetConfigCon = nothing
	If Fn_MSM_SetViewConfigurationToContextOperations = true Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MSM_SetViewConfigurationToContextOperations ] Passed : executed successfully with case [ " & sAction &  " ] ")
	End If
End Function

 '*********************************************************		Function to Verify Error Message ***********************************************************************
'Function Name		:        Fn_MSM_ErrorDialogVerify ()

'Description	    	:        Verifies The Error Message 

'Parameters		     :    		sMesssage: Message to be Verified [Optional]
'			                         		 sButton: Button to be clicked on the doalig

'Return Value		: 			True/False

'Pre-requisite	    :		     Error dialog is diaplyed

'Examples		    :			Call  Fn_MSM_ErrorDialogVerify ("", "OK")

'History		    :		
'													Developer Name				        Date						Rev. No.						Changes Done						Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Vidya Kulkarni						     10/06/2011	              1.0																			Prasanna
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_ErrorDialogVerify(sMessage, sButton)

		 Dim dicErrorInfo
		  Set dicErrorInfo = CreateObject("Scripting.Dictionary")
		  dicErrorInfo.Add "Action", "ErrorDialogVerify"
		  dicErrorInfo.Add "Title",  "Failed to open component"		  
		  dicErrorInfo.Add "Message", sMessage
		  dicErrorInfo.Add "Button", sButton    
		  Fn_MSM_ErrorDialogVerify = Fn_SISW_MSM_ErrorVerify(dicErrorInfo)
		  Set dicErrorInfo = Nothing
	
End Function

'*********************************************************		Function to Get Allocation Table Node operation in MultiStructure Manager		***********************************************************************

'Function Name		:				Fn_MSM_AllocationTableNodeOperation

'Description			 :		 		This function is used to get the Multu-Structure Manager Table Node operation.

'Parameters			   :				1.	strAction = "Select"
'										2. StrNodeName:Name of the Node. 
'										3. strColName
'										4. strColValue
'										5. strPopupMenu
				
											
'Return Value		   : 				True / False / TableNodeIndex

'Pre-requisite			:				MultiStructure Manager window should be displayed .


'Examples				:			Fn_MSM_AllocationTableNodeOperation("Select","003465/A;1-Test :Allocation1","","","")
'							        Fn_MSM_AllocationTableNodeOperation("Expand","003465/A;1-Test :Allocation1","","","")
'									Fn_MSM_AllocationTableNodeOperation("Exist","003465/A;1-Test :Allocation1","","","")
'									
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sonal 				09-April-2010			1.0				Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MSM_AllocationTableNodeOperation(sAction, sNodeName,sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_AllocationTableNodeOperation"
	Dim objTable, iRowCounter, objApplet, bReturn, iRows
	Dim arrCol,arrTypes,iCount,aMenu,strMenu
	Dim sSelectedNodes,iBOMLineColIndex

	If inStr(sColName,"Sources")>0 or inStr(sColName,"Targets")>0 Then
		sColName =Replace(Replace(sColName,"Targets","Targets without Duplicate"),"Sources","Sources without Duplicate")
	End If
	

	Set objApplet = JavaWindow("MultiStructManager").JavaWindow("MSWindow")
	Set objTable = objApplet.JavaTable("AllocationsTable")
	If Fn_SISW_UI_Object_Operations("Fn_MSM_AllocationTableNodeOperation", "Exist", objTable, SISW_MINLESS_TIMEOUT) = False Then
		Call Fn_ToolbatButtonClick("Show/Hide Allocation Navigator Panel")  ' [Mainline-20170830-20_09_2017-JotibaT-Porting]
			If  Fn_UI_ObjectExist("Fn_MSM_AllocationTableNodeOperation", objTable) = False Then
				Fn_MSM_AllocationTableNodeOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_MSM_AllocationTableNodeOperation ] Line Table does not exist.")
				Set objTable = nothing
				Exit function
			End If 
	End If
		
	
	Select Case sAction
		Case "Select"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Name")
			Fn_MSM_AllocationTableNodeOperation = False
			If iRowCounter <> -1 Then
				Call Fn_UI_JavaTable_SelectRow("Fn_MSM_AllocationTableNodeOperation",objApplet, "AllocationsTable", iRowCounter)
				Fn_MSM_AllocationTableNodeOperation = true 
			End If

		Case "PopupMenuSelect"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Name")
			Fn_MSM_AllocationTableNodeOperation = False
			If iRowCounter <> -1 Then
				Call Fn_UI_JavaTable_SelectRow("Fn_MSM_AllocationTableNodeOperation",objApplet, "AllocationsTable", iRowCounter)
				aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowCounter ,"Name", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCounter ,sColName, "RIGHT","NONE"
					End If
					wait 1
					Select Case Ubound(aMenu)
						Case "0"
							strMenu = JavaWindow("MultiStructManager").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
							JavaWindow("MultiStructManager").WinMenu("ContextMenu").Select strMenu
							Fn_MSM_AllocationTableNodeOperation = true
						Case "1"
							strMenu = JavaWindow("MultiStructManager").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
							JavaWindow("MultiStructManager").WinMenu("ContextMenu").Select strMenu
							Fn_MSM_AllocationTableNodeOperation = true
						Case Else
							Fn_MSM_AllocationTableNodeOperation = False
					End Select 
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Exist"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Name")
			Fn_MSM_AllocationTableNodeOperation = False
			If iRowCounter <> -1 Then
				Fn_MSM_AllocationTableNodeOperation = true 
			End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Expand"
			iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Name")
			Fn_MSM_AllocationTableNodeOperation = False
			If iRowCounter <> -1 Then
				objTable.Object.expandNode objTable.Object.getNodeForRow(cint(iRowCounter))					
				Fn_MSM_AllocationTableNodeOperation = True
			End If
		Case "CellVerify"
				If sNodeName <> "" Then
				iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Name")
				If iRowCounter <> -1 Then
					'objTable.SelectRow iRowCounter 
					bFound = Trim(cstr(objTable.GetCellData( iRowCounter,sColName)))
					If bFound = Trim(cstr(sValue)) Then
						Fn_MSM_AllocationTableNodeOperation = True
					Else
						Fn_MSM_AllocationTableNodeOperation = False
						If isNumeric(bFound) Then
							 bFound = Abs(bFound)
							 If cstr(bFound) = Trim(cstr(sValue)) Then
								 Fn_MSM_AllocationTableNodeOperation = True
							end  If
						End If
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_AllocationTableNodeOperation] Cell verified of MSM Allocation Table Node [" + sNodeName + "]")
				Else
					Fn_MSM_AllocationTableNodeOperation = False
				End If
			Else
				Fn_MSM_AllocationTableNodeOperation = False
			End If
       Case "GetCellData"
			If sNodeName <> "" Then
				iRowCounter = Fn_MSM_TableRowIndex(objTable, sNodeName,"Name")
				If iRowCounter <> -1 Then
					'objTable.SelectRow iRowCounter 
					Fn_MSM_AllocationTableNodeOperation = Trim(cstr(objTable.GetCellData( iRowCounter,sColName)))
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MSM_AllocationTableNodeOperation] Cell verified of MSM Allocation Table Node [" + sNodeName + "]")
				Else
					Fn_MSM_AllocationTableNodeOperation = False
				End If
			Else
				Fn_MSM_AllocationTableNodeOperation = False
			End If
		Case "AddColumn"
			
			arrCol=Split(sColName,"~")
			objTable.SelectColumnHeader 0,"RIGHT"
			objTable.JavaMenu("label:=Insert column\(s\) \.\.\.").Select
			wait 2
			If sValue<>"" Then
				arrTypes=Split(sValue,":")
				JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").JavaTree("CategoryAndType").Activate "Types:"+arrTypes(0)
				wait 1
				Call Fn_JavaTree_Select("Fn_MSM_AllocationTableNodeOperation", JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns"), "CategoryAndType","Types:"+sValue)
			End If
			wait 2
			
			' Added code to handle change coloumn dailog 
			If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns").Exist(1) Then
				Set ObjChangeColumns=JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
				sListValueName="ListAvailableCols"
			ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ChangeColumns").Exist(1) Then
				Set ObjChangeColumns=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ChangeColumns")
				sListValueName="AvailableCol"
			End If
			If Not ObjChangeColumns.Exist(1) Then
				Set objTable = nothing
				Exit Function
			End If
				
			For iCount=0 To UBound(arrCol)	
				Call Fn_List_Select("Fn_MSM_AllocationTableNodeOperation", ObjChangeColumns, sListValueName,arrCol(iCount))
				Call  Fn_Button_Click("Fn_MSM_AllocationTableNodeOperation",ObjChangeColumns,"Add")
				bReturn=True
			Next
			
			If bReturn=True Then
				Call  Fn_Button_Click("Fn_MSM_AllocationTableNodeOperation",ObjChangeColumns,"Apply")
				If ObjChangeColumns.JavaButton("Close").Exist(1) Then
					Call  Fn_Button_Click("Fn_MSM_AllocationTableNodeOperation",ObjChangeColumns,"Close")
				ElseIf ObjChangeColumns.JavaButton("Cancel").Exist(1) Then
					Call  Fn_Button_Click("Fn_MSM_AllocationTableNodeOperation",ObjChangeColumns,"Cancel")
				End If
				Fn_MSM_AllocationTableNodeOperation=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "GetSelectedRow"
				iRows = CInt(objTable.Object.getRowCount)
				sSelectedNodes = ""
				iBOMLineColIndex = cInt( Fn_MSM_TableColumnIndex(objTable, "Name") )
				
				For iCount =0  to iRows - 1 
					If objTable.Object.isRowSelected(iCount) Then
							If sSelectedNodes = "" Then
									sSelectedNodes = objTable.Object.getValueAt(iCount, iBOMLineColIndex ).toString()
							Else
									sSelectedNodes = sSelectedNodes & ":" & objTable.Object.getValueAt(iCount, iBOMLineColIndex ).toString()
							End If
					End if
				Next
				
				If sSelectedNodes = ""  Then
						Fn_MSM_AllocationTableNodeOperation =False
				Else
						Fn_MSM_AllocationTableNodeOperation = sSelectedNodes
				End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case Else
			Fn_MSM_AllocationTableNodeOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Function [ Fn_MSM_AllocationTableNodeOperation ] execution failed due to invalid case [ "& sAction &" ]")
			Set objTable = nothing
			Set objApplet = nothing
			Exit function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function [ Fn_MSM_AllocationTableNodeOperation ] Successfully executed with case [ "& sAction &" ]")
	Set objTable = nothing
	Set objApplet = nothing
End Function


'*********************************************************		Function to create basic Item		***********************************************************************
'Function Name		:				Fn_SISW_LegacyItemBasicCreate

'Description			 :		 		 Creats an Item with basic information

'Parameters			   :	 			1.StrItemType: Type of the item.
'													 2.StrConfItem: True or False
'													 2.StrItemID: ID of the item it should be unique.
'													3.StrItemRevID:Revision ID of the item.
'													4.StrItemName:Name of item.
'													5.StrItemDesc: Description of the item.
'													6:StrItemUOM: Unit of measure of item. ( not handling this part)

'Return Value		   : 				Item Id  -  Revision Id

'Pre-requisite			:		 		should be logged in

'Examples				:				 Fn_SISW_LegacyItemBasicCreate("Item","OFF","1213132","A","my","","")

'History					 :		
'	Developer Name			Date			Rev. No.	Reviewer		Changes Done					
'	------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Archana D			05/01/2013		1.0							Created
'	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_LegacyItemBasicCreate(StrItemType,StrConfItem,StrItemID,StrItemRevID,StrItemName,StrItemDesc,StrItemUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_LegacyItemBasicCreate"
	Dim sItemId, sRevId
	Dim objDialogNewItem
	'Select menu [File -> New -> Item...]
	If Fn_UI_ObjectExist("Fn_SISW_LegacyItemBasicCreate",Window("TeamcenterWindow").JavaDialog("New Item"))=False Then
        Call Fn_MenuOperation("Select","File:New:Item...")
       Call  Fn_ReadyStatusSync(3)
	End If
	
	'Check the existence of "New Item " window
	'Set objDialogNewItem=Fn_UI_ObjectCreate("Fn_SISW_LegacyItemBasicCreate",Window("TeamcenterWindow").JavaDialog("New Item"))
	Set objDialogNewItem = Window("TeamcenterWindow").JavaDialog("New Item")
	If objDialogNewItem.Exist(5)= False Then
		Set objDialogNewItem = JavaWindow("MultiStructManager").JavaWindow("New Item")

	End If
			'Select Item Type
			' To handle Application change (Funcitonality is changed to Funciton from Tc 09 0119 Build) Code added by Archana 
			If  StrItemType = "Functionality" Then StrItemType = "Function"
    		wait 1
           
			'checked Configuration item or not
			If StrConfItem <> "" Then
             Call Fn_CheckBox_Set("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"Configuration Item",StrConfItem)
			End If
			
			If objDialogNewItem.JavaTree("Tree").Exist(2) = True Then
				Call Fn_UI_JavaTree_Expand("Fn_ItemBasicCreate", objDialogNewItem, "Tree","Complete List")
				Call Fn_JavaTree_Select("Fn_ItemBasicCreate", objDialogNewItem, "Tree","Complete List")
				Call Fn_JavaTree_Select("Fn_ItemBasicCreate", objDialogNewItem, "Tree","Complete List:"+StrItemType)
			Else
			   Call Fn_List_Select("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"ItemType",StrItemType)
			End If

         ' Wait till  Button is Enabled
          objDialogNewItem.JavaButton("Next").WaitProperty "enabled", 1, 60000

		  Call  Fn_ReadyStatusSync(3)
          	'Click on "Next" button
			objDialogNewItem.JavaButton("Next").Click micLeftBtn
          ' Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"Next")

			If StrItemID <> "" Then
				'Set  Item Id
				objDialogNewItem.JavaStaticText("ID").SetTOProperty "label","ID:"
                 Call Fn_Edit_Box("Fn_SISW_LegacyItemBasicCreate",objDialogNewItem,"ItemID", StrItemID)
			End If

			If StrItemRevID <> "" Then
				'Set Revision ID
                Call Fn_Edit_Box("Fn_SISW_LegacyItemBasicCreate",objDialogNewItem,"RevisionID", StrItemRevID)
			End If

			If  StrItemID = "" Then
				'click on assign button
				If Not objDialogNewItem.JavaButton("Assign").GetROProperty("enabled")="0" Then
					Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem, "Assign")
				End If
			End If
			
			If  StrItemRevID = "" Then
				'click on assign button
				objDialogNewItem.JavaStaticText("ID").SetTOProperty "label","Revision:"
				If objDialogNewItem.JavaButton("Assign").Exist(2) = True Then
				If Not objDialogNewItem.JavaButton("Assign").GetROProperty("enabled")="0" Then
					Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem, "Assign")
				End If
				End If

			End If

			'*****************************************************************
			'Added by TusharB
			wait(3)
			'*****************************************************************
			
			'Extract Creation data
			objDialogNewItem.JavaStaticText("ID").SetTOProperty "label","ID:"
			sItemId = Fn_Edit_Box_GetValue("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"ItemID")
			If objDialogNewItem.JavaEdit("RevisionID").Exist(2)=True Then
				sRevId = Fn_Edit_Box_GetValue("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"RevisionID")
			End If
            
			
			'*****************************************************************
			'Added by Tushar B, In case ItemId and rev field are blank
'			If  sItemId = "" or sRevId = "" Then
'				'click on assign button
'				Call Fn_UpdateLogFiles(Time() & " - " & "WARNING - Assign button need to click again.", "")
'				call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Item ID not shown in ItemId Textbox[" + CStr(sItemId) + "]")
'				Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem, "Assign")
'				sItemId = Fn_Edit_Box_GetValue("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"ItemID")
'				sRevId = Fn_Edit_Box_GetValue("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"RevisionID")
'			End If
			'*****************************************************************
			
			'Set Item name
             Call Fn_Edit_Box("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"ItemName",StrItemName)
			'Set description
			If StrItemDesc <> "" Then
				Call Fn_Edit_Box("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"Description",StrItemDesc)
			End If
			'Set UOM
			If StrItemUOM <> "" Then
              Call Fn_Edit_Box("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"Unit of Measure",StrItemUOM)
			End If

			wait(2)
			objDialogNewItem.JavaButton("Finish").WaitProperty "enabled", 1, 20000
			
			'Click on "Finish" butto
'            Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem, "Finish") 
'			   = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
'				Sandeep : Added Code to handle Negative scenario for BMIDE test cases
				If objDialogNewItem.JavaButton("Finish").GetROProperty("enabled")="0" Then
					If StrItemDesc="" Then
						Call Fn_Edit_Box("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem,"Description","Test")
					End If
					objDialogNewItem.JavaButton("Finish").WaitProperty "enabled", 1, 20000
					Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem, "Finish")
				Else
					Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem, "Finish")
				End If
				'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
				Fn_SISW_LegacyItemBasicCreate = sItemId & "-" & sRevId
				Call Fn_ReadyStatusSync(1)
            
			'Click on Close button
            'Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem, "Close") 
            If Fn_UI_ObjectExist("Fn_SISW_LegacyItemBasicCreate",objDialogNewItem)=True Then
				'Click on Close button
				Call Fn_Button_Click("Fn_SISW_LegacyItemBasicCreate", objDialogNewItem, "Close") 
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Item of ID [" + CStr(sItemId) + "]")

		Set objDialogNewItem=Nothing
End Function

'*********************************************************	Generic function to handle Error dialogs in MSM Module  	***********************************************************************
'Function Name		:		Fn_SISW_MSM_ErrorVerify()

'Description		:	The function is generic function to handle error dialogs. It is created after combining error dialog functions from MultiStructureMananger.vbs
'									Fn_MSM_ErrorWindowVerify
'									Fn_MSM_ErrorDialogVerify

'Parameters			 :	 			1.  dicErrorInfo
											
'Return Value		 : 				True/False

'Pre-requisite		 :		 		NA.

'Examples			 :				  Dim dicErrorInfo
'												  Set dicErrorInfo = CreateObject("Scripting.Dictionary")
'												  dicErrorInfo.Add "Action", "ErrorDialogVerify"
'												  dicErrorInfo.Add "Title", ""Failed to open component"
'												  dicErrorInfo.Add "ControlName", strControlName
'												  dicErrorInfo.Add "Message", "Cannot open this type of object in this application"
'												  dicErrorInfo.Add "Button", "OK"    
'												  Fn_MSM_ErrorWindowVerify = Fn_SISW_MSM_ErrorVerify(dicErrorInfo)
'												  Set dicErrorInfo = Nothing

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sushma Pagare          8-Jul-2013
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
public Function Fn_SISW_MSM_ErrorVerify(dicErrorInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_MSM_ErrorVerify"

	Dim dicKeys, dicItems, iCounter, bReturn
	Dim sAction, sTitle, sErrorMsg,sButton, sAppMsg
	Dim objErrorDialog ,sControlName  
	Dim objJApplet, objWin,  iIndex

	On Error Resume Next
	Fn_SISW_MSM_ErrorVerify = False

	dicKeys = dicErrorInfo.Keys
	dicItems = dicErrorInfo.Items
	For  iCounter=0 to dicErrorInfo.Count-1
		Select Case dicKeys(iCounter)
			Case "Action"
					sAction = dicItems(iCounter)
			Case "Title"
					sTitle= dicItems(iCounter)
			Case "Message"
					sErrorMsg= dicItems(iCounter)
					GBL_EXPECTED_MESSAGE=sErrorMsg
			Case "ControlName"
					sControlName= dicItems(iCounter)
			Case "Button"
					sButton = dicItems(iCounter)
		End Select
	Next
	
	Select Case sAction

		''  This covers  Fn_MSM_ErrorWindowVerify(strWindowName,strControlName,strErrMsg,strButton)
		Case "ErrorWindowVerify"
			
			Set objErrorDialog=JavaWindow("MultiStructManager").JavaWindow("ErrorWindow")
			objErrorDialog.SetTOProperty "title",sTitle
			If objErrorDialog.Exist(5) Then
				If sControlName="" Then
					sControlName="StaticText"
				End If
				Select Case sControlName
					Case "StaticText"
						sAppMsg=objErrorDialog.JavaStaticText("ErrorText").GetROProperty("label")
						If InStr(1,Trim(LCase(sAppMsg)),Trim(LCase(sErrorMsg)))>0 Then
							Fn_SISW_MSM_ErrorVerify=True
						Else
							GBL_ACTUAL_MESSAGE=sAppMsg
						End If
					Case "EditBox"
						sAppMsg=objErrorDialog.JavaEdit("Details").GetROProperty("value")
						If InStr(1,Trim(LCase(sAppMsg)),Trim(LCase(sErrorMsg)))>0 Then
							Fn_SISW_MSM_ErrorVerify=True
						Else
							GBL_ACTUAL_MESSAGE=sAppMsg
						End If
				End Select
				objErrorDialog.JavaButton("OK").SetTOProperty "label",sButton
				Call Fn_Button_Click("Fn_SISW_MSM_ErrorVerify", objErrorDialog,sButton)
			End If			
			Set objErrorDialog=Nothing
			Exit Function
		
		'' 	Case "ErrorDialogVerify": This covers Fn_MSM_ErrorDialogVerify(sMesssage, sButton)
		Case "ErrorDialogVerify"

			If sTitle  = "" Then
				sTitle = "Failed to open component"
			End If
			'If Error Message is blank, then take it from global veriable
			If sErrorMsg = "" Then
				sErrorMsg = sErrorText
			End If
			Set objJApplet = Fn_SISW_MSM_GetObject("JApplet")
			Set objWin = Fn_SISW_MSM_GetObject("ErrorDialog")
			      
			For iCounter = 0 to 10
					objJApplet.SetTOProperty "index", iCounter
					objWin.SetTOProperty "title",sTitle  
					If objWin.Exist(5) Then
						If sErrorMsg <> "" Then
								sActMsg = objWin.JavaEdit("JTextArea").GetROProperty("value")
								If instr(sActMsg, sErrorMsg) <> 0 Then
										Fn_SISW_MSM_ErrorVerify = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Validated Message [" + sErrorMsg + "] On [ Open Component Error] Dialog")
								Else
										GBL_ACTUAL_MESSAGE=sActMsg
										Fn_SISW_MSM_ErrorVerify = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Validate Message [" + sErrorMsg + "] On [Open Component Error] Dialog")
								End If
								objWin.JavaButton("OK").SetTOProperty "label", sButton
								objWin.JavaButton("OK").Click micLeftBtn
								If Err.Number < 0 Then
									Fn_SISW_MSM_ErrorVerify = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Button [" + sButton + "] On [Open Component Error] Dialog")
								End If
								Exit For
						End If						
					End If
			Next
			Set objWin = Nothing
			Set objJApplet = Nothing
			Exit Function
						
	End Select

End Function



'*********************************************************		Function for Tabs in Multi-Structure Manager ***********************************************************************

'Function Name		:			Fn_MSM_DataPanelTabOperations

'Description		:			This function is used to Activate, Verify Activate

'Parameters			:			1.	strAction:
'								2.	strPanelName:
											
'Return Value		:			True/False

'Pre-requisite		:			Data panel in Multi-Structure Manager should be displayed .

'Examples			:			
		'Call Fn_MSM_DataPanelTabOperations("Activate","Graphics","")
		'Call Fn_MSM_DataPanelTabOperations("VerifyActivate","Graphics","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Reema W				24-Jun-2014			1.0										
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_DataPanelTabOperations(strAction,strPanelName,strMenuName)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_DataPanelTabOperations"
	on Error Resume Next
	Dim objName
	Fn_MSM_DataPanelTabOperations=False	
	If strPanelName <> "" Then
		JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("DataPanelHeader").SetTOProperty "label", strPanelName
	End If
	Select Case strAction		
		Case "Activate"			'('"Activate","Graphics","")
			 JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("DataPanelHeader").Click 5,5,"LEFT"
			Fn_MSM_DataPanelTabOperations=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Activated [" + strPanelName + "] Tab from Data Panel")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "VerifyActivate"	'("VerifyActivate","Graphics","")
			Set objName =  JavaWindow("MultiStructManager").JavaWindow("MSWindow").JavaStaticText("DataPanelHeader").Object			
			Fn_MSM_DataPanelTabOperations=objName.selected()	
			If Fn_MSM_DataPanelTabOperations="true" then
				Fn_MSM_DataPanelTabOperations=True	
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified [" + strPanelName + "] Tab is selected.")
			Else
				Fn_MSM_DataPanelTabOperations=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to Verify [" + strPanelName + "] Tab is selected.")
			End If
			Set objName = Nothing
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case Else
			Fn_MSM_DataPanelTabOperations=False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail : Function [ Fn_MSM_DataPanelTabOperations ] execution failed due to invalid case [ "& strAction &" ]")
	End Select

End Function

'************************************Function to perform operations on Bottom Buttons in Multi Structure Manager************************************

'	Function Name		    :		Fn_MSM_BottomContextOperations

'	Description				:	    This function is used to perform operations on Bottom Buttons in Multi Structure Manager

'	Parameters				:		1. sAction 			: Action need to perform.
'                                  	2. sButtomButton 	: Open Structure Context By Name / Open CC By Name ...
'									3. sSearch 			: Name of configation context
'									4. sVerify 			: Verify the Specific value in Configuration Context 										
'									5. sReserve			: For future use
											
'	Return Value		   	: 		True / False

'	Pre-requisite			:	  	Multi-Structure Manager Perspective should be open 

'	Examples				:	    CASE "SearchAndClickCC"	: Call Fn_MSM_BottomContextOperations("SearchAndClickCC","OpenCCByName", "cc123", "", "")

'	History:
'
'	Developer Name			Date			Rev. No.				Changes Done												Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam			16-Feb-2016	         1.0			     	 Created									[Tc1122:2016011300:16Feb2016:VivekA:NewDevelopment]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MSM_BottomContextOperations(sAction, sButtomButton, sSearch, sVerify, sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_BottomContextOperations"

	Dim objMSMApplet
	Dim iCnt
	
	Fn_MSM_BottomContextOperations = False
	
	Set objMSMApplet = Fn_UI_ObjectCreate("Fn_MSM_BOMTableNodeOpeations", JavaWindow("MultiStructManager").JavaWindow("MSWindow"))
	
	Select Case sButtomButton
		Case "OpenCCByName"
			Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_MSM_BottomContextOperations", "Set", objMSMApplet, "OpenCCbyName", "ON")
			wait 1			
			If Fn_UI_Object_SetTOProperty_ExistCheck("Fn_MSM_BottomContextOperations", objMSMApplet.JavaDialog("OpenConfigurationbyName"),"title","Open CC by name") = True Then
				
				Set objOpenCCTable = Fn_UI_ObjectCreate("Fn_MSM_BOMTableNodeOpeations", objMSMApplet.JavaDialog("OpenConfigurationbyName").JavaTable("CCTable"))
				
				Select Case sAction
					Case "SearchAndClickCC"
						If sSearch <> "" Then
							Call Fn_SISW_UI_JavaEdit_Operations("Fn_MSM_BottomContextOperations", "Set",  objMSMApplet.JavaDialog("OpenConfigurationbyName"), "Name", sSearch)
						End If
						Call Fn_SISW_UI_JavaButton_Operations("Fn_MSM_BottomContextOperations", "Click", objMSMApplet.JavaDialog("OpenConfigurationbyName"), "Find")
						wait 2
						If Fn_SISW_UI_Object_Operations("Fn_MSM_BottomContextOperations","Exist", objOpenCCTable,"") = True Then
							For iCnt = 0 To objOpenCCTable.GetROProperty("rows") - 1
								If Fn_SISW_UI_JavaTable_Operations("Fn_MSM_BottomContextOperations", "VerifyCellData", objOpenCCTable, "", "", "", iCnt, "Object", sSearch, "", "") = True Then
									objOpenCCTable.DoubleClickCell 0,"Object","LEFT"
									Fn_MSM_BottomContextOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Searched and Clicked [" + sSearch + "] .")
								Else
									Fn_MSM_BottomContextOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to Search and click [" + sSearch + "] .")	
								End If
							Next
						End If
				End Select
				Set objOpenCCTable = Nothing
			Else
				Fn_MSM_BottomContextOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to Verify Existance of  Open CC by name Dialog .")	
			End If 
	End Select
	
	Set objMSMApplet = Nothing
	
End Function
'*********************************************************		Function to Perform Publish/UnPublish operation in Multi-Structure Manager		***********************************************************************

'Function Name		        :			Fn_MSM_Publish_UnPublish_Operation

'Description			  :		 		This function is used to perform Publish/Unpublish Data in Multi-Structure Manager

'Parameters				 :	 			1. sAction : "Publish" or "UnPublish". (Action to Perform on the Structures)
'									   2. sCheckbox : "transform" or "shape". (Option to select on Publish/Unpublish Data Dialog. Action to select both transform and shape should be implemented.)
'									   3. sButton : "OK" or "Cancel" Button.  (To click on Publish/Unpublish Data Dialog)
'											
'Return Value		   	: 			 True/False

'Pre-requisite			:		 Two different structures should be displayed in two different Panels. One Child Item should be selected from each Structure.

'Examples				:		Fn_MSM_Publish_UnPublish_Operation("Publish","transform","OK")
'							 Fn_MSM_Publish_UnPublish_Operation("UnPublish","transform","OK")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Alok Dew			21-Nov-2017				1.0					Created				TC11.4(20171106.00)_NewDevelopment_PoonamC_22Nov2017							
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_Publish_UnPublish_Operation(sAction,sCheckbox,sButton)
	GBL_FAILED_FUNCTION_NAME = "Fn_MSM_Publish_UnPublish_Operation"
	Dim objPublishDlg,objAssocDlg,sMenu,objConfDlg

	Fn_MSM_Publish_UnPublish_Operation = False

	Set objPublishDlg = Fn_SISW_MSM_GetObject("PublishData")
	Set objAssocDlg = Fn_SISW_MSM_GetObject("PublishDataCreate")
	Set objConfDlg = Fn_SISW_MSM_GetObject("PublishDataCreate")

	'Checking whether Publish Data Dialog Exists
	If Fn_UI_ObjectExist("Fn_MSM_Publish_UnPublish_Operation",objPublishDlg) = False Then 
		If sAction = "Publish" Then
			sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("PSE_Menu"),"ToolsStructureAlignementPublishData")
		ElseIf sAction="Publish_NoButton" Then
			sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("PSE_Menu"),"ToolsStructureAlignementPublishData")
		ElseIf sAction = "UnPublish" Then
			sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"DeletePublishLinkForSource")
		End If
		Call Fn_MenuOperation("Select",sMenu) 'Selecting the option Publish Data/Delete publish link from source
		Call Fn_ReadyStatusSync(1)
	End If

	If Fn_UI_ObjectExist("Fn_MSM_Publish_UnPublish_Operation",objPublishDlg) Then

		Select Case sAction
			Case "Publish"
				'Check checkbox option
				If sCheckbox <> "" Then
					Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_MSM_Publish_UnPublish_Operation","Set",objPublishDlg,sCheckbox,"ON")
					Wait SISW_MICRO_TIMEOUT
				End If
				'Click on button OK/Cancel
				If sButton <> "" Then
					Call Fn_SISW_UI_JavaButton_Operations("Fn_MSM_Publish_UnPublish_Operation","Click",objPublishDlg,sButton)
					Call Fn_ReadyStatusSync(1)
				End If
				If Fn_UI_ObjectExist("Fn_MSM_Publish_UnPublish_Operation",objAssocDlg) Then 'if global association dialog exists then clicks on Yes
					Call Fn_SISW_UI_JavaButton_Operations("Fn_MSM_Publish_UnPublish_Operation","Click",objAssocDlg,"Yes")
					Call Fn_ReadyStatusSync(1)
				End If
				Fn_MSM_Publish_UnPublish_Operation = True
			Case "Publish_NoButton"
				'Check checkbox option
				If sCheckbox <> "" Then
					Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_MSM_Publish_UnPublish_Operation","Set",objPublishDlg,sCheckbox,"ON")
					Wait SISW_MICRO_TIMEOUT
				End If
				'Click on button OK/Cancel
				If sButton <> "" Then
					Call Fn_SISW_UI_JavaButton_Operations("Fn_MSM_Publish_UnPublish_Operation","Click",objPublishDlg,sButton)
					Call Fn_ReadyStatusSync(1)
				End If
				If Fn_UI_ObjectExist("Fn_MSM_Publish_UnPublish_Operation",objAssocDlg) Then 'if global association dialog exists then clicks on Yes
					Call Fn_SISW_UI_JavaButton_Operations("Fn_MSM_Publish_UnPublish_Operation","Click",objAssocDlg,"No")
					Call Fn_ReadyStatusSync(1)
				End If
				Fn_MSM_Publish_UnPublish_Operation = True
			Case "UnPublish"
				   'Check checkbox option
					If sCheckbox <> "" Then
						Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_MSM_Publish_UnPublish_Operation","Set",objPublishDlg,sCheckbox,"ON")
						Wait SISW_MICRO_TIMEOUT
					End If
					'Click on button OK/Cancel
					If sButton <> "" Then
						Call Fn_SISW_UI_JavaButton_Operations("Fn_MSM_Publish_UnPublish_Operation","Click",objPublishDlg,sButton)
						Call Fn_ReadyStatusSync(1)
					End If
					
					If Fn_UI_ObjectExist("Fn_MSM_Publish_UnPublish_Operation",objConfDlg) Then
							Call Fn_SISW_UI_JavaButton_Operations("Fn_MSM_Publish_UnPublish_Operation","Click",objConfDlg,"Yes")
					End If
					Fn_MSM_Publish_UnPublish_Operation = True
		End Select
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MSM_Publish_UnPublish_Operation ] [ "& objPublishDlg.toString &" ] does not exist.")
		Fn_MSM_Publish_UnPublish_Operation = False
	End If

	Set objPublishDlg = Nothing
	Set objAssocDlg = Nothing
	Set objConfDlg = Nothing
	
End Function
'#***********************************************************************************************************************************
'#
'#	Function Name		    :		Fn_MSM_ConfigurationContextOperations
'#
'#	Description				:	    This function is used to perform operations on Configuration Context dialog in Multi Structure Manager
'#
'#	Parameters				:		1. sAction 			: Action need to perform.
'#                                  2. dicdetails 		: dictionary object with config details
'#									3. sButtons 		: Name of button
'#											
'#	Return Value		   	: 		True / False
'#
'#	Pre-requisite			:	  	Multi-Structure Manager Perspective should be open 
'#
'#	Examples				:	    Set dicdetails = creatobject(""Scripting.dictionary")
'#										dicdetails("ConfigType") = "ConfigurationContext"
'#										dicdetails("Name") = "Config1"
'#										dicdetails("Description") = "Desc"
'#										dicdetails("RevisionRule")  = "Working"
'#										dicdetails("VariantRules") = "VarRuleName"
'#										dicdetails("closureRule") = "closerRuleName"
'#									Call Fn_MSM_ConfigurationContextOperations("Create",dicdetails,"OK")
'#	History:				
'#
'#	Developer Name			Date				Rev. No.				Changes Done												Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Poonam Chopade			01-Dec-2017	         1.0			     	 Created									[TC11.4_NewDevelopment_PoonamC_01Dec2017]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MSM_ConfigurationContextOperations(sAction,dicdetails,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_MSM_ConfigurationContextOperations"
	Dim objMSMConfigContext,arrButtons,iCount,sMenu
	
	Set objMSMConfigContext = Fn_SISW_MSM_GetObject("NewConfigContext")
	Fn_MSM_ConfigurationContextOperations = False
	
	'Check Existence of dialog
	If Fn_UI_ObjectExist("Fn_MSM_ConfigurationContextOperations",objMSMConfigContext) = False Then
			sMenu = Fn_GetXMLNodeValue( Fn_LogUtil_GetXMLPath("RAC_Menu"),"FileNewConfigurationContext")
			Call Fn_MenuOperation("Select",sMenu)
      	    Call Fn_ReadyStatusSync(3)
      	    If Fn_UI_ObjectExist("Fn_MSM_ConfigurationContextOperations",objMSMConfigContext) = False Then
      	    	Set objMSMConfigContext = Nothing
      	    	Exit Function
      	    End If
	End If
	
	Select Case sAction
		Case "Create"
			'Select Type
			If dicdetails("ConfigType") <> "" Then
				 	Call Fn_UI_JavaStaticText_Click("Fn_MSM_ConfigurationContextOperations",objMSMConfigContext,"SelectOption",1,1,"LEFT")
					Wait 1
					Call Fn_UI_JavaMenu_Select("Fn_MSM_ConfigurationContextOperations",objMSMConfigContext,dicdetails("ConfigType"))
			End If
			'Enter name
			If dicdetails("Name") <> "" Then
				 	Call Fn_SISW_UI_JavaEdit_Operations("Fn_MSM_ConfigurationContextOperations", "Set", objMSMConfigContext,"Name", dicdetails("Name"))
			End If
			'Enter description
			If dicdetails("Description") <> "" Then
				 	Call Fn_SISW_UI_JavaEdit_Operations("Fn_MSM_ConfigurationContextOperations", "Set", objMSMConfigContext,"Description", dicdetails("Description"))
			End If
			'select revision rule
			If dicdetails("RevisionRule") <> "" Then
				 	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_MSM_ConfigurationContextOperations", "Set", objMSMConfigContext, "revisionrule_16","ON")
				 	wait 1
				 	Call Fn_OpenByNameOperations("CellDoubleClick",dicdetails("RevisionRule"),"", "", "", "")
				 	wait 1
			End If
			'select Variant rules
			If dicdetails("VariantRules") <> "" Then
				 	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_MSM_ConfigurationContextOperations", "Set", objMSMConfigContext, "Add","ON")
				 	wait 1
				 	Call Fn_OpenByNameOperations("CellDoubleClick",dicdetails("VariantRules"),"", "", "", "")
				 	wait 1
			End If
			'select closure rule
			If dicdetails("closureRule") <> "" Then
				 	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_MSM_ConfigurationContextOperations", "Set", objMSMConfigContext, "find_16","ON")
				 	wait 1
				 	Call Fn_OpenByNameOperations("CellDoubleClick",dicdetails("closureRule"),"", "", "", "")
				 	wait 1
			End If
			Fn_MSM_ConfigurationContextOperations = True
	End Select
	
	'click on button on dialog
	If sButtons <> "" Then
		arrButtons = Split(sButtons,"~")
		For iCount = 0 To UBound(arrButtons)
			Call Fn_Button_Click("Fn_MSM_ConfigurationContextOperations",objMSMConfigContext,arrButtons(iCount))
			Call Fn_ReadyStatusSync(1)
		Next
	End If
	
	Set objMSMConfigContext = Nothing
	
End Function 
