Option Explicit
iTimeOut = 40
'---------------------------'Global variables for Teamcenter Perspective Names-----------------------------------------------------------
Public GBL_PERSPECTIVE_MANUFACTURING_PROCESS_PLANNER
GBL_PERSPECTIVE_MANUFACTURING_PROCESS_PLANNER="Manufacturing Process Planner"
'---------------------------'Global variables for Teamcenter Perspective Names-----------------------------------------------------------
'*********************************************************	Function List		***********************************************************************
'		0. Fn_SISW_MPP_GetObject(sObjectName)
'		1. Fn_MPP_ProcessBasicCreate()
'		2. Fn_MPP_BOMTable_RowIndex()
'		3. Fn_MPP_BOMTable_NodeOperation()
'		4. Fn_MPP_RevRuleSetEffectivityGroup()
'		5. Fn_MPP_SaveNewConfigContext()
'		6. Fn_MPP_ApplyConfigurationContext()
'		7. Fn_MPP_ConfigurationInformationOperations()
'		8. Fn_MPP_ViewMenuOperations()
'		9. Fn_MPP_OccurrenceGroupCreate()
'	   10. Fn_MPP_TabOperations()
'	   11. Fn_MPP_CCTreeOperations()
'	   12. Fn_MPP_BOMTable_ColIndex()
'	   13. Fn_MPP_BOMTable_ColumnOperations()
'	   14. Fn_MPP_OperationDetailCreateDic()
'	   15. Fn_MPP_AttachmentTableNodeOperation()
'	   16. Fn_MPP_RemoveLevel()
'	   17. Fn_MPP_ObjectReplaceSpecial()
'	   18. Fn_MPP_NoteCreate()
'	   19. Fn_MPP_VerifyActiveIC()
'	   20. Fn_SISW_MPP_SetIndexMPPAppletFromTab() - Supporting function
'	   21. Fn_MPP_EndItemAssemblyStateOperation
'	   22. Fn_MPP_WorkareaCreate
'	   23. Fn_MPP_CreatePublishLink_Operation
'	   24. Fn_MPP_AdvancedAccountabilityCheck_Operation	
'      25. Fn_MPP_VariantRuleOperations()	
'***************** Functions can be used in Manufacturing Process Planner from StructureManager vbs ***********************************************

'	   01. Fn_PSE_RevRuleSetEffectivityGroup()

'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_MPP_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_MPP_GetObject("PSEApplet")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sachin Joshi		 14-June-2012		1.0	
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_MPP_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ManufacturingProcessPlanner.xml"
	Set Fn_SISW_MPP_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'*********************************************************	Function List		***********************************************************************

 '*********************************************************		Function to create basic Process		***********************************************************************
'Function Name		:				Fn_MPP_ProcessBasicCreate

'Description			 :		 		 Creats an Process with basic information

'Parameters			   :	 			1.StrItemType: Type of the Process.
'													 2.StrConfItem: True or False
'													 2.StrItemID: ID of the item it should be unique.
'													3.StrItemRevID:Revision ID of the Process.
'													4.StrItemName:Name of Process.
'													5.StrItemDesc: Description of the Process.
'													6:StrItemUOM: Unit of measure of Process.

'Return Value		   : 				Item Id  -  Revision Id

'Pre-requisite			:		 		should be Prespective to Manufacturing Process Planner

'Examples				:				 Call Fn_MPP_ProcessBasicCreate("MEProcess","","","","Name","Desc","")

'History				 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sachin Joshi 												03-March-2011			              1.0										Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_ProcessBasicCreate(StrProcessAreaType,StrConfProcess,StrProcessID,StrProcessRevID,StrProcessName,StrProcessDesc,StrProcessUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_ProcessBasicCreate"
	Dim sItemId, sRevId
	Dim objDialogNewProcess,objDialogNewProcessNew
	'Select menu [File -> New -> Item...]
	'Check the existence of "New Process " window
	Set objDialogNewProcess = Window("MPPWindow").JavaDialog("New Process")
	Set objDialogNewProcessNew = JavaWindow("Manufacturing Process").JavaWindow("New Process")

	If objDialogNewProcess.Exist(3) = False And objDialogNewProcessNew.Exist(3) = False Then
        Call Fn_MenuOperation("Select","File:New:Process...")
	End If
	
    If objDialogNewProcessNew.Exist(3) And objDialogNewProcess.Exist(3) = False Then
    	Set objDialogNewProcess = JavaWindow("Manufacturing Process").JavaWindow("New Process")
    End If
    		
    		If objDialogNewProcess.Exist(3) = False And objDialogNewProcessNew.Exist(3) = False Then
        		 Call Fn_MenuOperation("Select","File:New:Item...")
        		 'Select Process Type
		    	'Call Fn_List_Select("Fn_MPP_ProcessBasicCreate", objDialogNewProcess,"Enter Additional Process",StrItemType)
		    	 'Call Fn_JavaTree_Select("Fn_MPP_ProcessBasicCreate", objDialogNewProcess, "Enter Additional Process","Complete List:"&StrProcessAreaType)
'		   		'Click on "Next" button
              	  'Call Fn_Button_Click("Fn_MPP_ProcessBasicCreate", objDialogNewProcess,"Next")
			End If
			
			'added code to handle Plant BOP selection from Complete List
			If objDialogNewProcess.JavaTree("Enter Additional Process").Exist(3) = True Then
				Call Fn_JavaTree_Select("Fn_MPP_ProcessBasicCreate", objDialogNewProcess, "Enter Additional Process","Complete List:"&StrProcessAreaType)
				Call Fn_Button_Click("Fn_MPP_ProcessBasicCreate", objDialogNewProcess,"Next")
			End If
			
			'Verify Id is Empty
			If StrItemID <> "" Then
				'Set  Item Id
                 'Call Fn_Edit_Box("Fn_MPP_ProcessBasicCreate",objDialogNewProcess,"ID", StrItemID)
				  Call Fn_Edit_Box("Fn_MPP_ProcessBasicCreate",objDialogNewProcess,"NewWorkAreaID", StrItemID)
			End If
			'Verify RevId is Empty
			If StrProcessRevID <> "" Then
				'Set Revision ID
                'Call Fn_Edit_Box("Fn_MPP_ProcessBasicCreate",objDialogNewProcess,"RevID", StrItemRevID)
                Call Fn_Edit_Box("Fn_MPP_ProcessBasicCreate",objDialogNewProcess,"NewWorkAreaRev", StrProcessRevID)
			End If

			'Click on Assign Button
			If  StrItemID = "" or StrProcessRevID = "" Then
				'click on assign button
                  objDialogNewProcess.JavaButton("Assign").SetTOProperty "Index","0"
                  Call Fn_Button_Click("Fn_MPP_ProcessBasicCreate", objDialogNewProcess, "Assign")
                  objDialogNewProcess.JavaButton("Assign").SetTOProperty "Index","1"
                  Call Fn_Button_Click("Fn_MPP_ProcessBasicCreate", objDialogNewProcess, "Assign")
			End If

            wait(3)

            'Extract Creation data
			sItemId = Fn_Edit_Box_GetValue("Fn_MPP_ProcessBasicCreate", objDialogNewProcess,"ID")
			'sItemId = Fn_Edit_Box_GetValue("Fn_MPP_ProcessBasicCreate", objDialogNewProcess,"NewWorkAreaID")
	            sRevId = Fn_Edit_Box_GetValue("Fn_MPP_ProcessBasicCreate", objDialogNewProcess,"RevID")
	            'sRevId = Fn_Edit_Box_GetValue("Fn_MPP_ProcessBasicCreate", objDialogNewProcess,"NewWorkAreaRev")
			'Set Process Name
			If StrProcessName <> "" Then
                 Call Fn_Edit_Box("Fn_MPP_ProcessBasicCreate",objDialogNewProcess,"Name", StrProcessName)
			End If

			'Set Process Desc
			If StrProcessDesc <> "" Then
                 Call Fn_Edit_Box("Fn_MPP_ProcessBasicCreate",objDialogNewProcess,"Description", StrProcessDesc)
			End If
			
			'Set Process UOM
			If StrProcessUOM <> "" Then 
				If objDialogNewProcess.JavaEdit("Unit of Measure:").Exist(5) =False Then
					objDialogNewProcess.JavaEdit("Unit of Measure:").SetTOProperty"toolkit class","org.eclipse.swt.custom.StyledText"
				End If
              Call Fn_Edit_Box("Fn_ItemBasicCreate", objDialogNewItem,"Unit of Measure:",StrProcessUOM)
			End If

			wait(2)
			objDialogNewProcess.JavaButton("Finish").WaitProperty "enabled", 1, 20000
			'Click on Finish Button 
			Call Fn_Button_Click("Fn_MPP_ProcessBasicCreate", objDialogNewProcess,"Finish")
			wait(1)
			Fn_MPP_ProcessBasicCreate = "'"&sItemId & "-" & sRevId
			Call Fn_ReadyStatusSync(1)

			 If objDialogNewProcess.Exist(3)Then
				'Click on Close button
				Call Fn_Button_Click("Fn_ItemBasicCreate", objDialogNewProcess, "Close") 
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Process of ID [" + CStr(sItemId) + "]")

		Set objDialogNewProcess=Nothing
End Function
'*********************************************************		Function to Get BOM Table Node Index into Structure Manager		***********************************************************************

'Function Name		:					Fn_MPP_BOMTable_RowIndex

'Description			 :		 		  This function is used to get the BOM Table Node Index.

'Parameters			   :	 			1. objTable - Table Object 
' 				  									2. sNodeName:Name of the Node to retrieve Index for.
											
'Return Value		   : 				 Node index

'Pre-requisite			:		 		Manufacturing Process Planner window should be displayed .

'Examples				:				 Fn_MPP_BOMTable_RowIndex(objTable, "518611/A;1-Item_518611 (view):001270/A;1-ffff")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin Joshi			09-March-2011		1.0	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ashwini P				20-March-2012		1.0	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_BOMTable_RowIndex(objTable, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_BOMTable_RowIndex"
	Dim nodeArr, aRowNode, iColIndex, aPath
	Dim iRowCounter, sNode, iInstance, iNodeCounter, iPathCounter, bFound 
	Dim iRows, sNodePath, sPath, StrNodePath
	Dim iNewCnt,bNewFlag
	Dim objComponent
	sPath = ""

	If Fn_UI_ObjectExist("Fn_MPP_BOMTable_RowIndex", objTable) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_BOMTable_RowIndex ] Table does not exist.")	
		Exit function
	End If
'    objTable.RefreshObject
	iColIndex = 0
	bFound = False
		bNewFlag=false
	If sNodeName <> "" Then
		' identifying RowId
		iRows = cInt(objTable.GetROProperty ("rows"))
		nodeArr = split(sNodeName , ":")
		iRowCounter = 0
		For iNodeCounter=0 to UBound(nodeArr)
				aRowNode = split(trim((nodeArr(iNodeCounter))),"@")
				If sPath = "" Then
							sPath =  trim(aRowNode(0))
				Else
							sPath = sPath &":"& trim(aRowNode(0))
				End If
		Next
		For iNodeCounter=0 to UBound(nodeArr)
			If iRowCounter = iRows  Then
				Exit for
			End If
			bNewFlag=False
			If Instr(1,trim((nodeArr(iNodeCounter))),"@@")>1 Then ' Added By Jotiba To check Same node name under different Parent 
				aRowNode = split(trim((nodeArr(iNodeCounter))),"@@")
				bNewFlag=True
			Else
				aRowNode = split(trim((nodeArr(iNodeCounter))),"@")
			End If
			
			iInstance = 0
			bFound = False
			do While iRowCounter < iRows
				If uBound(aRowNode) > 0 Then
							' instance number exist in name
							' initialize instance num
							' ith row matches with aRowNode(0) then
							sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
							If trim(sNodePath) = trim(aRowNode(0)) then
	                                set objComponent = ObjTable.object.getComponentForRow(iRowCounter)
									StrNodePath = ""
									Do while NOT (objComponent is Nothing)
										If StrNodePath = "" Then
											StrNodePath = objComponent.getProperty("bl_indented_title")
										Else
											StrNodePath =objComponent.getProperty("bl_indented_title") & ", " & StrNodePath
										End If
										
'------------------------------Code modified to identify the QTP version----------------------									
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
									
									If bNewFlag=True Then ' Added Tc12.2_20119021100_Maintenance_JotibaT To check Same node name under different Parent 
										StrNodePath =objTable.Object.getPathForRow(iRowCounter).toString()
										StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
										StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
									End If
									
									If instr(StrNodePath, "@BOM::") > 0 Then
										StrNodePath = trim(replace(StrNodePath,"""",""))
										aPath = split(StrNodePath,",")
										StrNodePath = ""
										For icnt = 0 to uBound(aPath)
											aPath(iCnt) = Left(aPath(iCnt), instr(aPath(iCnt),"@")-1)
											If StrNodePath = "" Then
												StrNodePath = trim(aPath(iCnt))
											else
												StrNodePath = StrNodePath & ", " & trim(aPath(iCnt))
											End If
										Next
									End If

									StrNodePath = trim(replace(StrNodePath,", ",":"))
									If instr(sPath, StrNodePath ) > 0 Then
										iInstance = iInstance +1
										If iInstance = cInt(aRowNode(1)) Then 
												If UBound(nodeArr) = iNodeCounter Then
														bFound = True
												End If
												Exit do
										End If
													'exit loop
									End If
							End if
				Else
					'ith row matches with aRowNode(0) then
					sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
					If trim(sNodePath) = trim(aRowNode(0)) then
                        set objComponent = ObjTable.object.getComponentForRow(iRowCounter)
						StrNodePath = ""
						Do while NOT (objComponent is Nothing)
							If StrNodePath = "" Then
								StrNodePath = objComponent.getProperty("bl_indented_title")
							Else
								StrNodePath =objComponent.getProperty("bl_indented_title") & ", " & StrNodePath
							End If
'------------------------------Code modified to identify the QTP version----------------------									
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
'						StrNodePath =objTable.Object.getPathForRow(iRowCounter).toString()
'						StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
'						StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
						If instr(StrNodePath, "@BOM::") > 0 Then
							StrNodePath = trim(replace(StrNodePath,"""",""))
							aPath = split(StrNodePath,",")
							StrNodePath = ""
							For icnt = 0 to uBound(aPath)
								aPath(iCnt) = Left(aPath(iCnt), instr(aPath(iCnt),"@")-1)
								If StrNodePath = "" Then
									StrNodePath = trim(aPath(iCnt))
								else
									StrNodePath = StrNodePath & ", " & trim(aPath(iCnt))
								End If
							Next
						End If
						StrNodePath = trim(replace(StrNodePath,", ",":"))
						If instr(sPath, StrNodePath ) > 0 Then
								If UBound(nodeArr) = iNodeCounter Then
									bFound = True
								End If
								Exit do
								'exit loop
						'Added by Vallari - above native method 'getPathForRow' does not fetch revised node names (revision B/C/D)
''						Elseif UBound(nodeArr) = iNodeCounter Then
''								strNodePath = ""
''								For iNewCnt = 0 to iRowCounter
''										strNodePath = strNodePath + ":" + objTable.object.getValueAt(iRowCounter, iColIndex).toString()
''								Next
''								strNodePath = Right(strNodePath, Len(strNodePath) - 1)
''								If instr(sPath, StrNodePaths) Then
''									bFound = True
''								End If
								Exit do
						End if
					End if
				End If
				iRowCounter = iRowCounter + 1
				' increment counter
			loop
		Next
	End If
	If bFound Then
				Fn_MPP_BOMTable_RowIndex = iRowCounter
	Else
				Fn_MPP_BOMTable_RowIndex = -1
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_BOMTable_RowIndex ] executed successfully.")
End Function


'******************************************************************Function to perform BOM Table Node operations************************************************************************************************************

'Function Name:				Fn_MPP_BOMTable_NodeOperation

'Description: 				 1. This function is used to select the BOM Table Node.
'						2. This function is used to multi-select BOM Table Nodes.
'						3. This function is used to expand the BOM Table Node.
'						4. This function is used to Collapse the BOM Table Node.
'						5. This function is used to verify the existance the BOM Table Node.
'						6. This function is used to edit the the BOM Table Node cell value.
'						7. This function is used to verify the BOM Table Node cell value.
'						8. This function is used to double-click the BOM Table Node cell.
'						9. This function is used to select the BOM Table Node popup menu.

'Parameters:			  1. sAction / sAction~sTabName : Action to be performed (Eg : Select/Exist/CellEdit etc.)
'						  2. sNodeName: Fully qualified path of the BOM Table Node (Node delimiter as ':') (Multi-Nodes delimiter as ',')
'						  3. sColName: Name of the BOM Table Column
'						  4. sValue: BOM Table cell value for Edit or Verify actions
'						  5. sPopupMenu: BOM Table Node context menu to be selected
'											  

'Return Value:				TRUE \ FALSE

'Pre-requisite:				Manufacturing Process Planner window should be displayed with BOM Table loaded.

'Examples:				 Call Fn_MPP_BOMTable_NodeOperation("Select", "000020/A;1-Top (View):000021/A;1-asm1 (View) @2:000080/A;1-asm3 @2", "", "", "")
'								Call Fn_MPP_BOMTable_NodeOperation("MultiSelect", "000359/A;1-EffGrp1 (View)~000359/A;1-EffGrp1 (View):000489/A;1-eff2 (View)", "", "", "")
'								Call Fn_MPP_BOMTable_NodeOperation("Expand", "000359/A;1-EffGrp1 (View):000489/A;1-eff2 (View)", "", "", "")
'                               Call Fn_MPP_BOMTable_NodeOperation("Expand Below", "000359/A;1-EffGrp1 (View):000489/A;1-eff2 (View)", "", "", "")
'                               Call Fn_MPP_BOMTable_NodeOperation("CellEdit", "000359/A;1-EffGrp1 (View):000489/A;1-eff2 (View) x 5", "Quantity", "3", "")
'								Call Fn_MPP_BOMTable_NodeOperation("CellVerify","001383/A;1-EffGrp1", "Item Description", "EffGrp1", "")
'								Call Fn_MPP_BOMTable_NodeOperation("PopupSelect", "000983/A;1-Top_Assembly (View):000984/A;1-Part1", "", "", "Send To:My Teamcenter")
'								Call Fn_MPP_BOMTable_NodeOperation("Exist", "000983/A;1-Top_Assembly (View):000984/A;1-Part1", "", "", "")

'		Added code to handle Applet index with passing extra parameter info in 'sAction'						
'		Eg.						Call Fn_MPP_BOMTable_NodeOperation("MultiSelect~StrContext-65790", "000359/A;1-EffGrp1 (View)~000359/A;1-EffGrp1 (View):000489/A;1-eff2 (View)", "", "", "")
'								Call Fn_MPP_BOMTable_NodeOperation("Expand~StrContext-65790", "000359/A;1-EffGrp1 (View):000489/A;1-eff2 (View)", "", "", "")
'History:
'										Developer Name			Date				Rev. No.			Changes Done												Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Sachin 	Joshi				09-March-2011			1.0  				
										'  Rajeev Gupta			  18-March-2011                     Case "Expand" Added 
										'  Rajeev Gupta			     21-March-2011                     Case "Expand Below" Added 
									    '  Rajeev Gupta			     21-March-2011                     Case "CellEdit" Added 
'										Sachin Joshi				 25-March-2011						Added Case "CellVerify"	
'										Sachin Joshi				 30-March-2011						Added Case "PopupSelect","Exist"	
'										Amit Talegaonkar		     21-Apr-2011						Added Case "MuliSelect"	
'										Koustubh Watwe			     21-Apr-2011						Commented unnecessary code and added code to clear selection in all cases.
'										Ketan Raje					 16-Sept-2011						Added Case : "IsSelected"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N, Koustubh W	     20-feb-2013						Added code to handle Applet index.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 										Vivek A 					 25-Jun2-2015						Added case "DeSelect" by Vivek to Deselect node in MPP BOM 
'										Call Fn_MPP_BOMTable_NodeOperation("DeSelect", "000020/A;1-Top (View):000021/A;1-asm1 (View):000080/A;1-asm3", "", "", "")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_BOMTable_NodeOperation(sAction, sNodeName, sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_BOMTable_NodeOperation"
	Dim iRowCounter, objTable, aNodeNames, objSelectType, intNoOfObjects,iCnt
	Dim strMenu, aMenu, iCounter,sCellData
	Dim iRows, sSelectedNodes
	Dim iSubMenuCount, objContextMenu, bFound
	Dim iCount, objNodeForRow, sColour, sColourCode
	Dim objWindow,bFlag,iRowCounter1,aNodeName1
	Dim aAction,iTabIndex,sTabName
	
'	If instr(1,sAction,"~") Then
		aAction = split(sAction,"~")
		If UBound(aAction) > 0 Then
			sAction = aAction(0)
		End If

		sTabName = JavaWindow("Manufacturing Process").JavaTab("ViewAll").GetROProperty("value")
		JavaWindow("Manufacturing Process").JavaTab("NewTabName").SetTOProperty "Value",sTabName
		
		If JavaWindow("Manufacturing Process").JavaTab("NewTabName").Exist(2) Then
			If JavaWindow("Manufacturing Process").JavaTab("InnerViewAll").Exist(1) Then
				sTabName = JavaWindow("Manufacturing Process").JavaTab("InnerViewAll").GetROProperty("value")
				JavaWindow("Manufacturing Process").JavaTab("NewInnerViewAll").SetTOProperty "Value",sTabName
			End If
		End If
		
		Set objWindow = JavaWindow("Manufacturing Process")
		Set objSelectType = description.Create()
		objSelectType("Class Name").value = "JavaTab"
		objSelectType("toolkit class").value = "org.eclipse.swt.custom.CTabFolder|com.teamcenter.rac.ms.ui.tab.TabFolderViewer\$1"						
		Set  objIntNoOfObjects = objWindow.ChildObjects(objSelectType)
			For icount = 0 To objIntNoOfObjects.Count-1 Step 1
				bFlag=False			
				iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
				For iCounter = 0 To iItemCount- 1 Step 1
					If trim(sTabName) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
						iTabIndex=iCounter
							If iTabIndex=0 OR iTabIndex=1 Then
								iTabIndex=2
							End If
						bFlag=True
						Exit For 
					End IF
				Next
				If bFlag=True Then Exit For 
			Next
		
			If Instr(1,sAction,"Ext")>0 Then ' Added By Jotiba T  
				For iCounter=iTabIndex+2 to 0 STEP -1
					Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "index",iCounter
					If Fn_SISW_UI_Object_Operations("Fn_MPP_BOMTable_NodeOperation","Exist",Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("NewCMEBOMTreeTable"),SISW_MICRO_TIMEOUT) Then
						Set objTable = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("NewCMEBOMTreeTable")
						aNodeName1=split(sNodeName,":")
						iRowCounter1 = Fn_MPP_BOMTable_RowIndex(objTable,aNodeName1(0))
						If iRowCounter1 <> -1 Then
							Exit for	
						End If
					End If
				Next
			Else
				If bFlag=True Then
					If Not JavaWindow("Manufacturing Process").JavaTab("NewInnerViewAll").Exist(1) Then
						For iCounter=iTabIndex to 0 STEP -1
							Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "index",iCounter
							If Fn_SISW_UI_Object_Operations("Fn_MPP_BOMTable_NodeOperation","Exist",Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("NewCMEBOMTreeTable"),SISW_MICRO_TIMEOUT) Then
								Exit for
							End If
						Next
						Set objTable = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("NewCMEBOMTreeTable")
					Else
						Set objTable = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable")			
						If Fn_SISW_MPP_SetIndexMPPAppletFromTab() = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_MPP_BOMTable_NodeOperation] BOM Table does not exists.")
							Fn_MPP_BOMTable_NodeOperation = False
							Exit function
						End If
					End If
				End If
			End If

'	Set objTable = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable")
	If Fn_UI_ObjectExist("Fn_MPP_BOMTable_NodeOperation", objTable) = False then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_MPP_BOMTable_NodeOperation] BOM Table does not exists.")
		Set objTable = nothing
		Fn_MPP_BOMTable_NodeOperation = False
		Exit function
	End if
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	' temp solution for setting focus
    'JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").JavaObject("BOMTreeTable").Click 1,1,"LEFT" 
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Select Case sAction
		Case "Select","SelectExt"
			If sNodeName <> "" Then
			'	If instr(sNodeName,"@") > 1 Then
			'		iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable, sNodeName)
			'	else
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
			'		If iRowCounter = -1 Then
			'			iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable, sNodeName)
			'		End If
			'	End If
				If iRowCounter <> -1 Then
                    objTable.Object.clearSelection  
					objTable.SelectRow iRowCounter 
					Fn_MPP_BOMTable_NodeOperation = True
				Else
					Fn_MPP_BOMTable_NodeOperation = False					
				End If
			Else
				Fn_MPP_BOMTable_NodeOperation = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

			Case "IsSelected"		
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
				If isNumeric(iRowCounter) Then
					If Cint(objTable.GetROProperty("SelectedRow")) = Cint(iRowCounter) Then						
						Fn_MPP_BOMTable_NodeOperation=True
					Else
						Fn_MPP_BOMTable_NodeOperation=False
					End If
				End if
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------			

		Case "MultiSelect","MultiSelectExt"
			aNodeNames = split(sNodeName , "~")

			'Clear the already selected Nodes
			objTable.Object.clearSelection

			For iCounter = 0 to UBound(aNodeNames)
				'iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,trim(aNodeNames(iCounter)))
			If instr(trim(aNodeNames(iCounter)),"@") > 1 Then
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,trim(aNodeNames(iCounter))) 
			else
			
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,trim(aNodeNames(iCounter))) 
				If iRowCounter = -1 Then
					iRowCounter = Fn_PSE_BOMTable_RowIndex(trim(aNodeNames(iCounter)))
				End If
			End If
				If iRowCounter <> -1 Then
					objTable.ExtendRow iRowCounter 
					Fn_MPP_BOMTable_NodeOperation = True
				Else
					Fn_MPP_BOMTable_NodeOperation = False
					objTable.Object.clearSelection
					Exit for
				End If
			Next
			
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            Case "Expand"
			If  sNodeName <> "" Then
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable, sNodeName)
				objTable.SelectRow iRowCounter
				If iRowCounter < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] Failed to Select PSE BOM Table Node [" + sNodeName + "]")
					Fn_MPP_BOMTable_NodeOperation = FALSE	
				Else
					If Fn_MenuOperation("WinMenuSelect", "View:Expand Options:Expand") = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_PSE_BOMTable_NodeOperation] Expanded PSE BOM Table Node [" + StrNodeName + "]")							
						Fn_MPP_BOMTable_NodeOperation = TRUE
					Else							
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_PSE_BOMTable_NodeOperation] Faile to Expanded PSE BOM Table Node [" + StrNodeName + "]")
						Fn_MPP_BOMTable_NodeOperation = FALSE
					End If						
				End If
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] Faile to Get PSE BOM Table Node [" + sNodeName + "]")
					Fn_MPP_BOMTable_NodeOperation = FALSE		
			End If
       '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		 Case "Expand Below","ExpandBelow"
			  If sNodeName <> "" Then
					iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable, sNodeName)
					If iRowCounter < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] Failed to Select PSE BOM Table Node [" + sNodeName + "]")
						Fn_MPP_BOMTable_NodeOperation = FALSE	
					Else
						'Clear the already selected Nodes
						objTable.Object.clearSelection
						wait 2
						wait 1
'						Call Fn_UI_JavaTable_SelectRow("Fn_MPP_BOMTable_NodeOperation", Window("MPPWindow").JavaWindow("MPPApplet"), "CMEBOMTreeTable",iRowCounter)
						objTable.SelectRow iRowCounter
						wait 2
						If Fn_MenuOperation("WinMenuSelect", "View:Expand Options:Expand Below") = True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_MPP_BOMTable_NodeOperation] Expanded PSE BOM Table Node [" + StrNodeName + "]")							
							Fn_MPP_BOMTable_NodeOperation = TRUE
							If JavaWindow("Manufacturing Process").JavaWindow("Search").JavaDialog("Expand Below").Exist Then
								JavaWindow("Manufacturing Process").JavaWindow("Search").JavaDialog("Expand Below").JavaButton("Yes").Click
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_MPP_BOMTable_NodeOperation] Clicked on Yes Button Under Expand Below Dialog")
								Fn_MPP_BOMTable_NodeOperation = True
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] Failed to Click on Yes Button Under Expand Below Dialog")
								Fn_MPP_BOMTable_NodeOperation = False				
							End If
						Else							
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] Faile to Expanded PSE BOM Table Node [" + StrNodeName + "]")
							Fn_MPP_BOMTable_NodeOperation = FALSE
						End If		
					End If
			  Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] Faile to Get PSE BOM Table Node [" + sNodeName + "]")
					Fn_MPP_BOMTable_NodeOperation = FALSE		
			  End If
     '.---------------------------------------This case is used to edit the BOM Table Node cell.----------------------------------------------
		Case "CellEdit","CellEditExt"
			If sNodeName <> "" Then
					iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable, sNodeName)
					'objTable.SelectRow iRowCounter
					If  iRowCounter < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] Faile to Select PSE BOM Table Node [" + StrNodeName + "]")
							Fn_MPP_BOMTable_NodeOperation = FALSE
					Else
						'Clear the already selected Nodes
						objTable.Object.clearSelection
						objTable.SelectRow iRowCounter
						
						objTable.SetCellData iRowCounter,sColName,sValue
						If  Err.Number < 0  Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Function [Fn_MPP_BOMTable_NodeOperation] Has Been Changed The Value for Column ["+sColName+"] as ["+sValue+"]")
							Fn_MPP_BOMTable_NodeOperation = False
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] to change The Value for Column ["+sColName+"] as ["+sValue+"]")
							Fn_MPP_BOMTable_NodeOperation = True
						End If
 							
					End If
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_BOMTable_NodeOperation] Faile to Get PSE BOM Table Node [" + sNodeName + "]")
					Fn_MPP_BOMTable_NodeOperation = FALSE
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "CellVerify","CellVerifyExt"
			If sNodeName <> "" Then
				'If instr(sNodeName,"@") > 1 Then
					iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName) 
				'else
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
				'	If iRowCounter = -1 Then
				'		iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable, sNodeName)
				'	End If
				'End If
				If iRowCounter <> -1 Then
					'Clear the already selected Nodes
					objTable.Object.clearSelection
					objTable.SelectRow iRowCounter
					
					If cstr(objTable.GetCellData( iRowCounter,sColName)) = cstr(sValue) Then
						Fn_MPP_BOMTable_NodeOperation = True
					Else
						Fn_MPP_BOMTable_NodeOperation = False
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MPP_BOMTable_NodeOperation] Cell verified of PSE BOM Table Node [" + sNodeName + "]")
				Else
					Fn_MPP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_MPP_BOMTable_NodeOperation = False
			End If
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
			Case "LoadInViewer"
				Err.Clear
				If sNodeName <> "" Then
					If instr(sNodeName,"@") > 1 Then
						iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName) 
					else
						iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
						If iRowCounter = -1 Then
							iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName) 
						End If
					End If
					If iRowCounter <> -1 Then
						objTable.Object.getNodeForRow(iRowCounter).stateIconClicked()	
						Set objSelectType=description.Create()
						objSelectType("Class Name").value = "JavaDialog"
						objSelectType("title").value = "Viewer Confirmation"
						Set  intNoOfObjects = JavaWindow("Manufacturing Process").JavaWindow("WEmbeddedFrame").ChildObjects(objSelectType)
						For iCnt = intNoOfObjects.count-1 to 0 Step -1
							If iCnt = 0 Then
								intNoOfObjects(iCnt).JavaButton("label:=Yes").Click
							Else
								intNoOfObjects(iCnt).JavaButton("label:=No").Click
							End If							
					   	Next
					   	Set intNoOfObjects = nothing
						Set objSelectType=nothing
						Fn_MPP_BOMTable_NodeOperation = True
					Else
						Fn_MPP_BOMTable_NodeOperation = False					
					End If
	
				Else
					Fn_MPP_BOMTable_NodeOperation = False
				End If
				
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "GetCellData"
			If sNodeName <> "" Then
					iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName) 
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
				If iRowCounter <> -1 Then

					sCellData = cstr(objTable.GetCellData( iRowCounter,sColName))
					If  sCellData <> ""  Then
						Fn_MPP_BOMTable_NodeOperation = sCellData
					Else
						Fn_MPP_BOMTable_NodeOperation = False
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MPP_BOMTable_NodeOperation] get Cell Data of PSE BOM Table Node [" + sNodeName + "]")
				Else
					Fn_MPP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_MPP_BOMTable_NodeOperation = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Exist", "Exists","ExistsExt"
			If sNodeName <> "" Then
				'If instr(sNodeName,"@") > 1 Then
				'	iRowCounter = Fn_MPP_BOMTable_RowIndex(sNodeName)
				'Else
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
				'	If iRowCounter = -1 Then
				'		iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable, sNodeName)
				'	End If
				'End If
				If iRowCounter <> -1 Then
					Fn_MPP_BOMTable_NodeOperation = True
				Else
					Fn_MPP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_MPP_BOMTable_NodeOperationt = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "PopupSelect","PopupSelectExt"
			If sNodeName <> "" Then
				iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
				If iRowCounter <> -1 Then
					'Split Context menu to Build Path Accordingly
					aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowCounter ,"BOM Line", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCounter ,sColName, "RIGHT","NONE"
					End If
					wait 2
					Select Case Ubound(aMenu)
						Case "0"
							strMenu = JavaWindow("Manufacturing Process").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
							JavaWindow("Manufacturing Process").WinMenu("ContextMenu").Select strMenu
						Case "1"
							strMenu = JavaWindow("Manufacturing Process").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
							JavaWindow("Manufacturing Process").WinMenu("ContextMenu").Select strMenu
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [ Fn_MPP_BOMTable_NodeOperation ] Context Menu Case NOT Exists for Supplied Menu [" + StrPopupMenu + "]")
							Fn_MPP_BOMTable_NodeOperation = False
					End Select
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_MPP_BOMTable_NodeOperation] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
					Fn_MPP_BOMTable_NodeOperation = True
				Else
					Fn_MPP_BOMTable_NodeOperation = False
				End If
			Else
				Fn_MPP_BOMTable_NodeOperation = False
			End If
		Case "DeSelect" ' Case added by Vivek to Deselect node in MPP BOM Table
				If sNodeName <> "" Then
					iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
			
					If iRowCounter <> -1 Then
                    	objTable.Object.clearSelection  
						objTable.DeselectRow iRowCounter 
						Fn_MPP_BOMTable_NodeOperation = True
					Else
						Fn_MPP_BOMTable_NodeOperation = False					
					End If
				Else
					Fn_MPP_BOMTable_NodeOperation = False
				End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'TC11.4(2017120100)_NewDevelopment_PoonamC_01Feb2018 : Added Case to verify Background color of Node
		Case "VerifyBackgroundColour","VerifyBackgroundColourExt"
				If sNodeName <> "" Then
					iRowCounter = Fn_MPP_BOMTable_RowIndex(objTable,sNodeName)
					If iRowCounter <> -1 Then
                    	Set objNodeForRow =  objTable.Object.getNodeForRow(iRowCounter) 
						sColour = objTable.Object.getBackground(objNodeForRow,false).toString()
						sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
						Select Case sValue
							Case "LIGHTGREEN"
								sColourCode = "[r=159,g=255,b=159]"
						End Select
						If sColour = sColourCode  Then
							Fn_MPP_BOMTable_NodeOperation = True
						Else
							Fn_MPP_BOMTable_NodeOperation = False
						End If	
					Else
						Fn_MPP_BOMTable_NodeOperation = False					
					End If
				Else
					Fn_MPP_BOMTable_NodeOperation = False
				End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case Else
			Fn_MPP_BOMTable_NodeOperation = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_BOMTable_NodeOperation ] Invalid Action [ " & sAction & " ].")
			Set objTable = nothing
			exit function
			
	End Select
	If Fn_MPP_BOMTable_NodeOperation <>FALSE then
		If instr(1,sAction,"~") Then
			Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "index","1"
		End If 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_BOMTable_NodeOperation ] executed successfully with Action [ " & sAction & " ].")	
	Else	
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to execute Function [ Fn_MPP_BOMTable_NodeOperation ] with Action [ " & sAction & " ].")
	End if
	Set objTable = nothing
End Function


''****************************************    Function to set Date, End Item, Unit, Effectivity Group for Revision Rule ***************************************
'
''Function Name		    :			  Fn_MPP_RevRuleSetEffectivityGroup
'
''Description			:  	      Function to set Date, End Item, Unit, Effectivity Group for Revision Rule 
'
''Parameters			:	 	1. sAction : Action need to perform
'								2. dicRevRuleInfo : Dictionary object 
'								  	"MRUList" - for future use
'								3. btnName = button to be clicked. button name eg. OK / Cancel
'											      
''											
''Return Value		    :  			True \ False
'
''Pre-requisite			:		 	 BOM Line should be selected
''Examples				:			 
' 
'							dicRevRuleInfo("bUseToday") =True / False / ""
'							dicRevRuleInfo("sEffectivityDate") = "10-Oct-2010 11:17:49" / ""
'							dicRevRuleInfo("sUnitNumber")="9" / ""
'							dicRevRuleInfo("sEndItem")="" / "000065" / "Top~000065"
'							dicRevRuleInfo("sEndItemBy")="" / "OpenByName" / "PasteFromClipboard" ( "MRUList") - for future used. Not yet implemented
'							dicRevRuleInfo("bAnyIntent")= True / False / ""
'							dicRevRuleInfo("sIntentName")= "I1~I2"  - ~ separated list of Intent Names
'							dicRevRuleInfo("sIntentDesc")="i1desc~i2Desc" -   ~ separated list of Intent Descriptions
'							dicRevRuleInfo("sAddIntentBy")="OpenByName" / "CreateNew" / ""
'							dicRevRuleInfo("sRemoveIntents") ="I21~I22"  - ~ separated list of Intent Names
'							dicRevRuleInfo("sEffectivityGrpAction") = "Append" / "Replace" / "Insert" / "Remove" / "View/Edit"
'							dicRevRuleInfo("sEffectivityGrpEntry") = "000006/A;1-test1"
'							dicRevRuleInfo("sEffectivityGrpSearchBy") = "MRUList" / "OpenByName" / "PasteFromClipboard"  ( "MRUList") - for future used. Not yet implemented
'							dicRevRuleInfo("sEffectivityGrp") = "" / "Eff_Grp1~000008" / "000008"
'							dicRevRuleInfo("sEffectivityGrpRev") = "" / "A"

'							Call Fn_MPP_RevRuleSetEffectivityGroup("Set", dicRevRuleInfo,"")
'- - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - 
'Example :
'							1) dicRevRuleInfo("sEffectivityGrp") =  "000643-bb"
'							dicRevRuleInfo("sEffectivityGrpAction") = "Append"
'							Call Fn_MPP_RevRuleSetEffectivityGroup("Set", dicRevRuleInfo,"")
'
'							2)dicRevRuleInfo("sEffectivityGrpEntry") =  "000643/A;1-bb"
'							Call Fn_MPP_RevRuleSetEffectivityGroup("VerifyEffectivityGrpEntry", dicRevRuleInfo,"")
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Sachin Joshi			09-March-2011		     	    1.0			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MPP_RevRuleSetEffectivityGroup(sAction, dicRevRuleInfo, btnName)
		GBL_FAILED_FUNCTION_NAME="Fn_MPP_RevRuleSetEffectivityGroup"
		Dim objRevRuleSetDate, aEndItemArr, aEffeGrp
		Dim iCount
		
		Set objRevRuleSetDate = JavaWindow("Manufacturing Process").JavaWindow("Search").JavaDialog("Set Date/Unit/End Item")
		Fn_MPP_RevRuleSetEffectivityGroup = False
		'creating object of dialog box
		If Fn_UI_ObjectExist("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate) = False Then
			Call Fn_MenuOperation("Select", "Tools:Revision Rule:Set Date/Unit/End Item..." )
			If Fn_UI_ObjectExist("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_MPP_RevRuleSetEffectivityGroup ] Failed to open [ Set Date/Unit/End Item/ Intent ] window.")
				Set objRevRuleSetDate = nothing
				Exit function
			End If
		End If
		Select Case sAction
			'- - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - 
			Case "Set"
                 objRevRuleSetDate.JavaTab("JTabbedPane").Select "Set Date/Unit/End Item"
				If dicRevRuleInfo("bUseToday") <> "" Then
					If cBool(dicRevRuleInfo("bUseToday")) Then
						Call Fn_CheckBox_Select("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "Use Today")
					Else
						Call Fn_CheckBox_Set("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "Use Today","OFF")
						'setting effectivity date
						If dicRevRuleInfo("sEffectivityDate") <> "" Then
							objRevRuleSetDate.JavaCheckBox("Date").Object.setDate dicRevRuleInfo("sEffectivityDate") 
						End If
					End If
				End If

				If dicRevRuleInfo("sUnitNumber") <> "" Then
					Call Fn_Edit_Box("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "EffectiveUnitNumber", dicRevRuleInfo("sUnitNumber"))
				End If

				' setting End Item..
				If dicRevRuleInfo("sEndItemBy") <> "" Then
					'EndItem selection type
					Select Case sEndItemBy
						Case "MRUList"
						Case "OpenByName"
								Call Fn_CheckBox_Select("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "OpenByName" )
								aEndItemArr = split(dicRevRuleInfo("sEndItem"),"~")
								Call Fn_OpenByNameOperations("CellDoubleClick", aEndItemArr(0) , aEndItemArr(1),"","","")
						Case "PasteFromClipboard"
							Call Fn_Button_Click("Fn_MPP_RevRuleSetEffectivityGroup",objRevRuleSetDate, "Paste")
					End Select
				Else
					'set end Item
					If dicRevRuleInfo("sEndItem") <> "" Then
						Call Fn_Edit_Box("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "End Item", dicRevRuleInfo("sEndItem") )
						objRevRuleSetDate.JavaEdit("End Item").Activate
					End If
				End If
				'effectivity group
				objRevRuleSetDate.JavaTab("JTabbedPane").Select "Effectivity Groups"
				wait(2)
				Select Case dicRevRuleInfo("sEffectivityGrpSearchBy")
					Case "MRUList"
					Case "OpenByName"
							Call Fn_CheckBox_Select("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "OpenByName" )
							aEffeGrp = split(dicRevRuleInfo("sEffectivityGrp"),"~")
							Call Fn_OpenByNameOperations("CellDoubleClick", aEffeGrp(0) , aEffeGrp(1),"","","")
					Case "PasteFromClipboard"
							Call Fn_Button_Click("Fn_MPP_RevRuleSetEffectivityGroup",objRevRuleSetDate, "Paste")
					Case else
						If dicRevRuleInfo("sEffectivityGrp") <> "" then
							Call Fn_Edit_Box("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "Effectivity Group", dicRevRuleInfo("sEffectivityGrp") )
							objRevRuleSetDate.JavaEdit("Effectivity Group").Activate
						end if 
				End Select
				' setting rev of effectivity group
				If dicRevRuleInfo("sEffectivityGrpRev") <> "" Then
					call Fn_List_Select("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "Effectivity GroupRev",dicRevRuleInfo("sEffectivityGrpRev"))
				End If
				' selecting effectivity group entry from list.
				If dicRevRuleInfo("sEffectivityGrpEntry") <> "" Then
					call Fn_List_Select("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "EffectivityGrpList",dicRevRuleInfo("sEffectivityGrpEntry"))
				End If

				' clicking on button
				If objRevRuleSetDate.JavaButton(dicRevRuleInfo("sEffectivityGrpAction")).WaitProperty("enabled", 1, 30) Then
					Call Fn_Button_Click("Fn_MPP_RevRuleSetEffectivityGroup",objRevRuleSetDate, dicRevRuleInfo("sEffectivityGrpAction"))
					Fn_MPP_RevRuleSetEffectivityGroup = True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_MPP_RevRuleSetEffectivityGroup ] Failed to Click on button [ " & dicRevRuleInfo("sEffectivityGrpAction") & " ]of [ Set Date/Unit/End Item/ Intent ] window.")
					Fn_MPP_RevRuleSetEffectivityGroup = False
				End If
		' - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - 
		Case "VerifyEffectivityGrpEntry"
			 objRevRuleSetDate.JavaTab("JTabbedPane").Select "Effectivity Groups"
			if dicRevRuleInfo("sEffectivityGrpEntry") <> "" then
				aEffeGrp = split(dicRevRuleInfo("sEffectivityGrpEntry"),"~")
				Fn_MPP_RevRuleSetEffectivityGroup = true
				For iCount = 0 to UBound(aEffeGrp)
					Fn_MPP_RevRuleSetEffectivityGroup = Fn_UI_ListItemExist("Fn_MPP_RevRuleSetEffectivityGroup", objRevRuleSetDate, "JList", aEffeGrp(iCount))
					If Fn_MPP_RevRuleSetEffectivityGroup = false Then
						Exit for
					End If
				Next
			end if
			
	End Select

	If btnName <> ""  Then
		' clicking on button
		Call Fn_Button_Click("Fn_MPP_RevRuleSetEffectivityGroup",objRevRuleSetDate, btnName)
	End If
	Set objRevRuleSetDate = nothing
          Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : [ Fn_MPP_RevRuleSetEffectivityGroup ] executed successfully with case [ " & sAction & " ]")
End function

'*********************************************************		Function to Add/Remove Cloumn into Manufacturing Process Planner		***********************************************************************

'Function Name		:					Fn_MPP_BOMTable_ColumnOperation

'Description			 :		 		  This function is used to Add/Remove Column in Bom Table

'Parameters			   :	 			1.  StrAction:Action to be Performed
'													2.	StrColName : Valid Column Name						
											
'Return Value		   : 				 True/False

'Pre-requisite			:		 		Manufacturing Process Planner window should be displayed .

'Examples				:				Call Fn_MPP_BOMTable_ColumnOperation("Remove","AIE_OCC_ID")
'													Call Fn_MPP_BOMTable_ColumnOperation("Add","AIE_OCC_ID")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin	Joshi						25-March-2011		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MPP_BOMTable_ColumnOperation(StrAction,StrColName)
		GBL_FAILED_FUNCTION_NAME="Fn_MPP_BOMTable_ColumnOperation"
		Dim PopUpMenu,iColIndex,ObjTable,ArrCol,iIndex,sColToAdd, objList, intCol, objChangeColumnDialog,IntCounter,StrName,bFlag,IntCols
		Set ObjTable = JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").Object

		Select Case StrAction
			Case "Add"
					ArrCol = Split(StrColName,":",-1,1)
					For iIndex = 0 To Ubound(ArrCol)
							'Check that Column is present in the BOMTable.
							iColIndex =  Fn_MPP_BOMTable_ColIndex(ArrCol(iIndex))
							If iColIndex = -1 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Warning: Column does not  exist in the Application.Need to Add Column ["& ArrCol(iIndex) &"]." )
									sColToAdd = sColToAdd +":"+ArrCol(iIndex)
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Column ["& ArrCol(iIndex) &"] exists in the Application" )
									Fn_MPP_BOMTable_ColumnOperation =TRUE
							End if
					Next
					If sColToAdd <>""  Then
						sColToAdd = Mid(sColToAdd, 2,Len(sColToAdd))
						ArrCol = Split(sColToAdd,":",-1,1)
						'Invoke Choose Column Window if it is not present on the screen
						Set objChangeColumnDialog = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						If NOT objChangeColumnDialog.Exist( 1)  Then
								JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").SelectColumnHeader "#1","RIGHT"       	
								JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").JavaMenu("label:=Insert column\(s\) ...").Select 										       
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
								Set objList = objChangeColumnDialog.JavaList("ListAvailableCols").Object
						End If
						Call Fn_ReadyStatusSync(2)							
						For iIndex = 0 To Ubound(ArrCol)							
								'Select Col to be added from the lsit
								intCol = objChangeColumnDialog.JavaList("ListAvailableCols").GetItemIndex(ArrCol(iIndex))
								objList.ensureIndexIsVisible intCol
								objChangeColumnDialog.JavaList("ListAvailableCols").ExtendSelect ArrCol(iIndex)
						Next
						' Click on ADD Button
						objChangeColumnDialog.JavaButton("Add").Click
						' Click on Apply Button
						objChangeColumnDialog.JavaButton("Apply").Click
						' Click on Apply Button
						objChangeColumnDialog.JavaButton("Cancel").Click
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Added  Column  ["& sColToAdd &"] in BOMTable")									
						Fn_MPP_BOMTable_ColumnOperation = TRUE					
					End If
		Case "Remove"
			ArrCol = Split(StrColName,":",-1,1)
			For iIndex = 0 To Ubound(ArrCol)										
					'Check that Column is present in the BOMTable
					iColIndex =  Fn_MPP_BOMTable_ColIndex(ArrCol(iIndex))						
					If iColIndex = -1 Then							
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"WARNING:Column dose not  exist in the Application.No Need to Remove Column ["& ArrCol(iIndex) &"]")
							Fn_MPP_BOMTable_ColumnOperation  = FALSE
					Else
							'Remove the given Column From the BOMTable.													
							JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").SelectColumnHeader iColIndex,"RIGHT"
							JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").JavaMenu("label:=Remove this column").Select											
							JavaWindow("Manufacturing Process").JavaWindow("Remove Column").JavaButton("Yes").Click		
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Successfully removed Column  ["& ArrCol(iIndex) &"] from BOMTable.")          																
							Fn_MPP_BOMTable_ColumnOperation  =TRUE										 						
					End if
			Next
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid case  ["& StrAction &"].")
				Fn_MPP_BOMTable_ColumnOperation = False
		End Select
		 
		Set objList = Nothing
		Set ObjTable = Nothing
		Set objChangeColumnDialog = nothing
 End Function

'*********************************************************		Function to Get BOM Table Column Index into Manufacturing Process Planner		***********************************************************************

'Function Name		:					Fn_MPP_BOMTable_ColIndex

'Description			 :		 		  This function is used to get the BOM Table Node Index.

'Parameters			   :	 			1.  StrColName:Name of the Col to retrieve Index for.
											
'Return Value		   : 				 Col index

'Pre-requisite			:		 		Manufacturing Process Planner window should be displayed .

'Examples				:				Fn_MPP_BOMTable_ColIndex("All Notes)

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sachin	Joshi			25-March-2011		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe			21-june-2012		1.0			Modified object hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_BOMTable_ColIndex(StrColName)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_BOMTable_ColIndex"
	On Error Resume Next
	Dim IntCols , IntCounter, ObjTable, StrColIndex, StrName

	'Verify that PSE BOM Table is displayed
	If Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").Exist(iTimeOut) Then

		'Get the No. of cols present in the BOM Table

		IntCols = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").GetROProperty("cols")
		Set ObjTable = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").Object
	
		'Get the Col No. of required Column
		For IntCounter = 0 to IntCols -1
			StrName = ObjTable.getColumnName(IntCounter)
		  
			If Trim(StrName) = Trim(StrColName) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: The Column Index for Column [" + StrColName + "] is [" +IntCounter +"] in MPP BOMTable")
				Fn_MPP_BOMTable_ColIndex = IntCounter
				Exit For
			End If
		Next
		If Cint(IntCounter) = Cint(IntCols) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"WARNING: The Column [" + StrColName + "] dose not exist in MPP BOM table." )
			Fn_MPP_BOMTable_ColIndex=-1
		End If

		'Release the Table object
	   set ObjTable = Nothing

	End If
End Function
''****************************************    Function to save Configuration Context  ***************************************
'
''Function Name		 :			  Fn_MPP_SaveNewConfigContext
'
''Description		     :  	      Function to save Configuration Context  
'
''Parameters		    :	 	1. sAction : Action need to perform
'						2. sConfigType : Configuration Context Type : For future use not yet implemented 
'						3. sName : Configuration Context Name
'						4. sDescription : Configuration Context Description
'						5. sRevisionRule : Closure Rule Name
'						6. sVariantRule : Closure Rule Name
'						7. sClosureRule : Closure Rule Name
'						8. sEffectivityGroups : ~ separated list of Effectivity Group names
'											      
''											
''Return Value		    :  			True \ False
'
''Pre-requisite		     :		BOM Line should be selected
''Examples		     :		Call Fn_MPP_SaveNewConfigContext("Save", "", "name", "dec", "", "", "", "") 
'					         Call Fn_MPP_SaveNewConfigContext("Apply", "", "name", "dec", "", "", "", "")
'					         Call Fn_MPP_SaveNewConfigContext("Verify", "", "name", "dec", "RevRul", "VarRul", "", "")
												
'History:
'		Developer Name			Date				Rev. No.	Reviewer	Changes Done			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Sushma Pagare			12-Apr-2011		     	 1.0			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Vandana Patel			17-July-2012		     1.0		Koustubh	Modified cases according to TC10.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_SaveNewConfigContext(sAction, sConfigType, sName, sDescription, sRevisionRule, sVariantRule, sClosureRule, sEffectivityGroups)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_SaveNewConfigContext"
	Dim objSaveConfigContext, iCount, aEffectivityGroups, bReturn
	Set objSaveConfigContext = JavaWindow("Manufacturing Process").JavaWindow("Save As New Configuration")

	Fn_MPP_SaveNewConfigContext = false
	If objSaveConfigContext.Exist(5) = False Then
			Call Fn_MenuOperation("Select", "File:CC:Save as New Configuration Context")
	End If
	If objSaveConfigContext.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] Failed to open [ Save New Configuration Context ] window.")
			Set objSaveConfigContext = nothing
			Exit function
	End If
	' setting Selct Config type
	If sConfigType <> "" then
		bReturn = Fn_UI_JavaTreeGetItemPathExt("FunctionName", objSaveConfigContext.JavaTree("ConfigurationContext"),"Complete List:" & sConfigType,"","")
		If bReturn <> False Then
			objSaveConfigContext.JavaTree("ConfigurationContext").Select bReturn
		Else
			bReturn = Fn_UI_JavaTreeGetItemPathExt("FunctionName", objSaveConfigContext.JavaTree("ConfigurationContext"),"Most Recently Used:" & sConfigType,"","")
			If bReturn <> False Then
				objSaveConfigContext.JavaTree("ConfigurationContext").Select bReturn
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] Failed to select Configuration type.")
				Set objSaveConfigContext = nothing
				Exit function
			End If
		End If
	End if
	Call Fn_Button_Click("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Next")
	
	Select Case sAction
	'- - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  
		Case "Save", "Apply"
			' setting Name
			If sName <> "" then
				call Fn_Edit_Box("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Name",sName)
			end if
			' setting Description
			If sDescription <> "" then
				call Fn_Edit_Box("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Description",sDescription)
			end if
			' setting Revision Rule
			'If sRevisionRule <> "" then
			'	call Fn_Edit_Box("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Revision Rule",sRevisionRule)
			'end if
			' setting Variant Rule
			'If sVariantRule <> "" then
			'	call Fn_Edit_Box("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Variant Rule",sVariantRule)
			'end if
			' setting Closure Rule
			'If sClosureRule <> "" then
			'	call Fn_Edit_Box("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Closure Rule",sClosureRule)
			'end if
			' selecting Effectivity group
			'If sEffectivityGroups <> "" then
			'	aEffectivityGroups = split(sEffectivityGroups, "~")
			'	For iCount = 0 to UBound(aEffectivityGroups)
			'		If Fn_UI_ListItemExist("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Effectivity Groups", aEffectivityGroups(iCount)) Then
			'			call Fn_UI_JavaList_ExtendSelect("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Effectivity Groups", aEffectivityGroups(iCount))
			'		else
			'			Set objSaveConfigContext = nothing
			'			Exit function
			'		End If
			'	Next
			'end if
			' clicking on apply button
			'If sAction = "Apply" Then
			'	Call Fn_Button_Click("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Apply")
			'End If
			' clicking on OK
			'Call Fn_Button_Click("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"OK")
			Call Fn_Button_Click("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Finish")
			Fn_MPP_SaveNewConfigContext = True
		'- - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  
		Case "Verify"
			' setting Name
			If sName <> "" then
				If Fn_Edit_Box_GetValue("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Name") <> sName then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] Name does not match with [ " & sName & " ] .")
					Set objSaveConfigContext = nothing
					Exit function
				end if
			end if
			' setting Description
			If sDescription <> "" then
				If Fn_Edit_Box_GetValue("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Description") <> sDescription  then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] Description does not match with [ " & sDescription & " ] .")
					Set objSaveConfigContext = nothing
					Exit function
				end if
			end if
'			' setting Revision Rule
'			If sRevisionRule <> "" then
'				If Fn_Edit_Box_GetValue("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Revision Rule") <> sRevisionRule then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] Revision Rule does not match with [ " & sRevisionRule & " ] .")
'					Set objSaveConfigContext = nothing
'					Exit function
'				end if
'			end if
'			' setting Variant Rule
'			If sVariantRule <> "" then
'				If Fn_Edit_Box_GetValue("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Variant Rule") <> sVariantRule then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] Variant Rule does not match with [ " & sVariantRule & " ] .")
'					Set objSaveConfigContext = nothing
'					Exit function
'				end if
'			end if
'			' setting Closure Rule
'			If sClosureRule <> "" then
'				If Fn_Edit_Box_GetValue("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Closure Rule") <> sClosureRule then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] Closure Rule does not match with [ " & sClosureRule & " ] .")
'					Set objSaveConfigContext = nothing
'					Exit function
'				end if 
'			end if
'			If sEffectivityGroups <> "" then
'				aEffectivityGroups = split(sEffectivityGroups, "~")
'				For iCount = 0 to UBound(aEffectivityGroups)
'					If Fn_UI_ListItemExist("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Effectivity Groups", aEffectivityGroups(iCount)) = False Then
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] Effectivity Group [ " & aEffectivityGroups(iCount) & " ] is not present in list.")
'						Set objSaveConfigContext = nothing
'						Exit function
'					End If
'				Next
'			end if
			Call Fn_Button_Click("Fn_MPP_SaveNewConfigContext",objSaveConfigContext,"Cancel")
			Fn_MPP_SaveNewConfigContext = True
		'- - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  - - - - - - - - -  
		Case else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_SaveNewConfigContext ] execution failed. Invalid case [ " & sAction & " ]")
	End Select

	If Fn_MPP_SaveNewConfigContext Then
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : [ Fn_MPP_SaveNewConfigContext ] executed successfully with case [ " & sAction & " ]")
	End If
	Set objSaveConfigContext = nothing
End Function

''****************************************    Function to Apply Configuration Context  ***************************************
'
''Function Name		 :			  Fn_MPP_ApplyConfigurationContext
'
''Description		     :  	      Function to save Configuration Context  
'
''Parameters		    :	 	1. sAction : Action need to perform
'						2. sName : Configuration Context Name
'						3. sSearchName : Type to Search Name : "' / "OpenByName"
'						4. sDescription : Configuration Context Description
'						5. sRevisionRule : Closure Rule Name
'						6. sVariantRule : Closure Rule Name
'						7. sClosureRule : Closure Rule Name
'						8. sEffectivityGroups : ~ separated list of Effectivity Group names
'											      
''											
''Return Value		    :  			True \ False
'
''Pre-requisite		     :		BOM Line should be selected

''Examples		     :		Call Fn_MPP_ApplyConfigurationContext("Apply", "name","", "dec", "", "", "", "") 
'					         Call Fn_MPP_ApplyConfigurationContext("Apply", "name", "OpenByName","dec", "", "", "", "")
												
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Sushma Pagare		12-Apr-2011		     	    1.0			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

public Function Fn_MPP_ApplyConfigurationContext(sAction, sName, sSearchName, sDescription, sRevisionRule, sVariantRule, sClosureRule, sEffectivityGroups)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_ApplyConfigurationContext"
	Dim aEffectivityGroups, objConfigDialog, iCount
	Dim objOpenByName,iRowCount,objTable
	Fn_MPP_ApplyConfigurationContext = False
	Set objConfigDialog = JavaWindow("Manufacturing Process").JavaWindow("Apply Configuration Context")

	 'Verify Existance of Apply Configuration Context Dialog
	If JavaWindow("Manufacturing Process").JavaWindow("Apply Configuration Context").Exist(5) = False Then
		Call Fn_MenuOperation("Select", "File:CC:Apply Configuration Context...")
		If  objConfigDialog.Exist(5) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_ApplyConfigurationContext ] Failed to open [ Apply Configuration Context] window.")
			Set objConfigDialog = nothing
			Exit function
		end if
	End If
	Select Case sAction
		Case "Apply"
				' setting Name
				If sName <> "" then
					Select Case sSearchName
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						Case "OpenByName"
								Call Fn_Button_Click("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"OpenByName")
								Set objOpenByName = JavaWindow("DefaultWindow").JavaWindow("Open by Name")
								If objOpenByName.Exist(15)  = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_ApplyConfigurationContext ] Failed to open [ Open By Name ] window.")
									Set objOpenByName = nothing
									Set objConfigDialog = nothing
									Exit function
								end if	
								'Clear the contents
								objOpenByName.JavaEdit("Name").Set ""
								'set the values
								Call Fn_Edit_Box("Fn_MPP_ApplyConfigurationContext",objOpenByName,"Name",sName)
								
								'Click on Find Button
								wait(1)
								call Fn_Button_Click("Fn_MPP_ApplyConfigurationContext",objOpenByName, "Find")
								wait(3)
								'objOpenByName.JavaTable("SrchResultTable").DoubleClickCell 0,"Object"
								objOpenByName.JavaTable("SrchResultTable").DoubleClickCell 0, 0
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -								
						Case else
							call Fn_Edit_Box("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"Name",sName)
							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
					End Select

				end if
				' setting Description
				If sDescription <> "" then
					call Fn_Edit_Box("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"Description",sDescription)
				end if
				' setting Revision Rule
				If sRevisionRule <> "" then
					call Fn_Edit_Box("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"Revision Rule",sRevisionRule)
				end if
				' setting Variant Rule
				If sVariantRule <> "" then
					call Fn_Edit_Box("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"Variant Rule",sVariantRule)
				end if
				' setting Closure Rule
				If sClosureRule <> "" then
					call Fn_Edit_Box("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"Closure Rule",sClosureRule)
				end if
				If sEffectivityGroups <> "" Then
					aEffectivityGroups = split(sEffectivityGroups, "~")
					For iCount = 0 to UBound(aEffectivityGroups)
						If Fn_UI_ListItemExist("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"Effectivity Groups", aEffectivityGroups(iCount)) Then
							call Fn_UI_JavaList_ExtendSelect("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"Effectivity Groups", aEffectivityGroups(iCount))
						else
							Set objConfigDialog = nothing
							Exit function
						End If
					Next
				End If
				' clicking on OK
				Call Fn_Button_Click("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"OK")
				Fn_MPP_ApplyConfigurationContext = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		Case "VerifyNoObjectFound"
					If sName <> "" then
								Call Fn_Button_Click("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"OpenByName")
								Set objOpenByName = JavaWindow("DefaultWindow").JavaWindow("Open by Name")
								If objOpenByName.Exist(15)  = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_ApplyConfigurationContext ] Failed to open [ Open By Name ] window.")
									Set objOpenByName = nothing
									Set objConfigDialog = nothing
									Exit function
								end if	
								'Clear the contents
								objOpenByName.JavaEdit("Name").Set ""
								'set the values
								Call Fn_Edit_Box("Fn_MPP_ApplyConfigurationContext",objOpenByName,"Name",sName)
								
								'Click on Find Button
								wait(1)
								call Fn_Button_Click("Fn_MPP_ApplyConfigurationContext",objOpenByName, "Find")
								wait(3)

								If objOpenByName.JavaWindow("WEmbeddedFrame").JavaDialog("Nothing found!").exist(15) then
											Call Fn_Button_Click("Fn_MPP_ApplyConfigurationContext",objOpenByName.JavaWindow("WEmbeddedFrame").JavaDialog("Nothing found!"), "OK")
											Call Fn_Button_Click("Fn_MPP_ApplyConfigurationContext",objOpenByName,"Cancel")
											Call Fn_Button_Click("Fn_MPP_ApplyConfigurationContext",objConfigDialog,"Cancel")
											Fn_MPP_ApplyConfigurationContext = True
								end if
				end if
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		Case else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_MPP_ApplyConfigurationContext ] execution failed. Invalid case [ " & sAction & " ]")
	End Select
	If Fn_MPP_ApplyConfigurationContext Then
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : [ Fn_MPP_ApplyConfigurationContext ] executed successfully with case [ " & sAction & " ]")
	End If
	Set objConfigDialog = nothing
End Function
'*********************************************************	Function to perform operations on Configuration Information	window 	***********************************************************************

'Function Name		       :					Fn_MPP_ConfigurationInformationOperations

'Description			 :		 		  This function is used to copy BOM Node.

'Parameters			:	 			1.  String - Action ( Verify )
'									2.  String - sFields - ~ separated list of fields
'									3.  String - sValues - ~ separated list of Values
											
'Return Value		         :			         True / False

'Pre-requisite			:		 	           Manufacturing Process Planner perspective should be already set

'Examples				:				  Call Fn_MPP_ConfigurationInformationOperations("Verify", "Header~Effectivity", "000208-assm50~Today", "False")
'History:
'		Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh W			13-Apr-2011		              1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MPP_ConfigurationInformationOperations(sAction, sFields, sValues, bClose)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_ConfigurationInformationOperations"
	Dim objConfigInfo, aFields, aValues, iCnt, sVal
	Fn_MPP_ConfigurationInformationOperations = False
	Set objConfigInfo = JavaWindow("Manufacturing Process").JavaWindow("Configuration Information")
	If objConfigInfo.exist(5) = False  Then
		'open window
		'Call Fn_ToolbarOperation("Click", "Show Information","")
		Call Fn_ToolbarOperation("Click", "Configuration Information...","")
		If objConfigInfo.exist(15) = False  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_ConfigurationInformationOperations ] Failed to open Configuration Information window." )
			Set objConfigInfo = nothing
			Exit function
		end if
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_ConfigurationInformationOperations ] Successfully open Configuration Information window." )
	Select Case sAction
		Case "Verify"
			aFields = split(sFields,"~")
			aValues = split(sValues,"~")
			Fn_MPP_ConfigurationInformationOperations = True
			For iCnt = 0 to UBound(aFields)
				Select Case lCase(trim(aFields(iCnt)))
					Case "header"
						sVal = trim(objConfigInfo.JavaStaticText("Header").GetROProperty("label"))
					Case "revision rule"
						If objConfigInfo.JavaEdit("Revision rule").Exist(5) = false then
							objConfigInfo.JavaStaticText("AccordionHeader").SetTOProperty "label" , "Configuration Rules"
							objConfigInfo.JavaStaticText("AccordionHeader").Click 1,1,"LEFT"
						end if
						sVal = trim(objConfigInfo.JavaEdit("Revision rule").GetROProperty("value"))
					Case "effectivity"
						If objConfigInfo.JavaEdit("Effectivity").Exist(5) = false then
							objConfigInfo.JavaStaticText("AccordionHeader").SetTOProperty "label" , "Configuration Rules"
							objConfigInfo.JavaStaticText("AccordionHeader").Click 1,1,"LEFT"
						end if
						sVal = trim(objConfigInfo.JavaEdit("Effectivity").GetROProperty("value"))
					Case "variant rule"
						If objConfigInfo.JavaEdit("Variant rule").Exist(5) = false then
							objConfigInfo.JavaStaticText("AccordionHeader").SetTOProperty "label" , "Configuration Rules"
							objConfigInfo.JavaStaticText("AccordionHeader").Click 1,1,"LEFT"
						end if
						sVal = trim(objConfigInfo.JavaEdit("Variant rule").GetROProperty("value"))
					Case "effectivity groups"
						If objConfigInfo.JavaEdit("EffectivityGroups").Exist(5) = false then
							objConfigInfo.JavaStaticText("AccordionHeader").SetTOProperty "label" , "Effectivity Groups"
							objConfigInfo.JavaStaticText("AccordionHeader").Click 1,1,"LEFT"
						end if
						sVal = trim(objConfigInfo.JavaEdit("EffectivityGroups").GetROProperty("value"))
					Case "in context"
						If objConfigInfo.JavaEdit("In context").Exist(5) = false then
							objConfigInfo.JavaStaticText("AccordionHeader").SetTOProperty "label" , "Context"
							objConfigInfo.JavaStaticText("AccordionHeader").Click 1,1,"LEFT"
						end if
						sVal = trim(objConfigInfo.JavaEdit("In context").GetROProperty("value"))
					Case "ic edit context"
						If objConfigInfo.JavaEdit("IC edit context").Exist(5) = false then
							objConfigInfo.JavaStaticText("AccordionHeader").SetTOProperty "label" , "Incremental Change"
							objConfigInfo.JavaStaticText("AccordionHeader").Click 1,1,"LEFT"
						end if
						sVal = trim(objConfigInfo.JavaEdit("IC edit context").GetROProperty("value"))
					Case else
						Fn_MPP_ConfigurationInformationOperations = false
						Exit for
				End Select
				If sVal <> aValues(iCnt) then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_ConfigurationInformationOperations ] Value of field [ " & trim(aFields(iCnt)) & " ] does not match with value [ "&aValues(iCnt)& " ]." )
					Fn_MPP_ConfigurationInformationOperations = false
					Exit for
				end if
				Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_ConfigurationInformationOperations ] Value of field [ " & trim(aFields(iCnt)) & " ] matches with value [ "&aValues(iCnt)& " ]." )
			Next
	End Select

	'Closing dialog
	If bClose <> "" Then
		If cBool(bClose) Then
			Call Fn_Button_Click("Fn_MPP_ConfigurationInformationOperations",objConfigInfo,"Close")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_ConfigurationInformationOperations ] Clicked on Close button." )
		End If
	else
		Call Fn_Button_Click("Fn_MPP_ConfigurationInformationOperations",objConfigInfo,"Close")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_ConfigurationInformationOperations ] Clicked on Close button." )
	End If
          Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_ConfigurationInformationOperations ] executed successfully with case [ " & sAction & " ]." )
	set objConfigInfo = nothing
end function

'*********************************************************		Function to perform operations on View Menu context menu***********************************************************************

'Function Name		    :		Fn_MPP_ViewMenuOperations

'Description			:	    This function is used to perform operations on View Menu context menu

'Parameters				:	1. sAction : Action need to perform.
'                                       			2. sMenu : ( : ) separated menu string.										
											
'Return Value		   	: 	True/False

'Pre-requisite			:	   

'Examples				:	 Call  Fn_MPP_ViewMenuOperations("Select", "Show Information")
'						Call  Fn_MPP_ViewMenuOperations("Exist", "Show Unconfigured Changes")
'						Call  Fn_MPP_ViewMenuOperations("IsEnabled", "Show Unconfigured Changes")
'						Call  Fn_MPP_ViewMenuOperations("IsChecked", "Show Unconfigured Changes")

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W			19-Apr-2011	           	  1.0			      Created
'	Ketan Raje			09-Sept-2011	          1.0			      Modified to handle Menu Operations in SE prespective.
'	Koustubh W			08-Nov-2011	           	  1.0			      Added code to close context menu.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MPP_ViewMenuOperations(sAction, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_ViewMenuOperations"
	Dim bReturn
	Fn_MPP_ViewMenuOperations = false
	Select Case sAction
		Case "Select","SESelect"
			If sAction = "SESelect" Then
				Call Fn_ToolbarButtonClick_Ext(1,"View Menu")
			ElseIf JavaWindow("Manufacturing Process").JavaTree("CCTree").exist(2) = True Then
				Call Fn_ToolbarButtonClick_Ext(2,"View Menu")
			Else
				Call Fn_ToolbarOperation("Click", "View Menu","")
			End If
			bReturn = JavaWindow("DefaultWindow").WinMenu("ContextMenu").CheckItemProperty (sMenu, "Exists", true, 2000)
			If bReturn then 
				JavaWindow("DefaultWindow").WinMenu("ContextMenu"). Select sMenu
				Fn_MPP_ViewMenuOperations = true
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MPP_ViewMenuOperations ] Fail : Failed to find menu [ " & sMenu &  " ] ")
			end if
		Case "IsChecked","SEIsChecked"
			If sAction = "SEIsChecked" Then
				Call Fn_ToolbarButtonClick_Ext(2,"View Menu")
			Else
				Call Fn_ToolbarOperation("Click", "View Menu","")
			End If
			bReturn = JavaWindow("DefaultWindow").WinMenu("ContextMenu").CheckItemProperty (sMenu, "Exists", true, 2000)
			If bReturn then 
				Fn_MPP_ViewMenuOperations = JavaWindow("DefaultWindow").WinMenu("ContextMenu").CheckItemProperty( sMenu, "Checked", 1, 2000)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MPP_ViewMenuOperations ] Pass : Property [ Checked ] = [ " &Fn_MPP_ViewMenuOperations &" ] of menu [ " & sMenu &  " ] ")
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MPP_ViewMenuOperations ] Fail : Failed to find menu [ " & sMenu &  " ] ")
			end if
			call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		Case "IsEnabled","SEIsEnabled"
			If sAction = "SEIsEnabled" Then
				Call Fn_ToolbarButtonClick_Ext(2,"View Menu")
			Else
				Call Fn_ToolbarOperation("Click", "View Menu","")
			End If
			bReturn = JavaWindow("DefaultWindow").WinMenu("ContextMenu").CheckItemProperty (sMenu, "Exists", 1, 2000)
			If bReturn then 
				Fn_MPP_ViewMenuOperations = JavaWindow("DefaultWindow").WinMenu("ContextMenu").CheckItemProperty (sMenu, "Enabled", true, 2000)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MPP_ViewMenuOperations ] Pass : Property [ Enabled ] = [ " &Fn_MPP_ViewMenuOperations &" ] of menu [ " & sMenu &  " ] ")
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MPP_ViewMenuOperations ] Fail : Failed to find menu [ " & sMenu &  " ] ")
			end if
			call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		Case "Exists", "Exist", "SEExist"
			If sAction = "SEExist" Then
				Call Fn_ToolbarButtonClick_Ext(2,"View Menu")
			Else
				Call Fn_ToolbarOperation("Click", "View Menu","")
			End If
			Fn_MPP_ViewMenuOperations = JavaWindow("DefaultWindow").WinMenu("ContextMenu").CheckItemProperty (sMenu, "Exists", 1, 2000)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MPP_ViewMenuOperations ] Pass : Property [ Exists ] = [ " &Fn_MPP_ViewMenuOperations &" ] of menu [ " & sMenu &  " ] ")
			call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		Case else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MPP_ViewMenuOperations ] Fail : Invalid case [ " & sAction &  " ] ")
			Exit function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[ Fn_MPP_ViewMenuOperations ] Pass : Function executed successfully with case [ " & sAction &  " ] ")
End Function
'*********************************************************		Function to perform operations on Occurrence Group ***********************************************************************

'Function Name		    :		Fn_MPP_OccurrenceGroupCreate

'Description			:	    This function is used to perform operations on View Menu context menu

'Parameters		        :		1. sAction : Action need to perform.
'                               2. sOGText : Occurrence Group Type
'                               3. sOGType : Occurrence Group Type
'                               4. bOpenOnCreate : for future use
'                               5. sOGName : name
'                               6. sOGDescription : description
											
'Return Value		   	: 	True / False

'Pre-requisite			:	   

'Examples				:	 Call Fn_MPP_OccurrenceGroupCreate("Create", "", "Complete List:OccurrenceGroup", "","Name", "Desc")

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W			20-Apr-2011	           	  1.0			      Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MPP_OccurrenceGroupCreate(sAction, sOGText, sOGType, bOpenOnCreate,sOGName, sOGDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_OccurrenceGroupCreate"
	Dim objOccGroup
	Dim arrType, iCnt, sPath, iCount
	Dim aOccurrenceGroupType, iRows, aReturn, sReturn
	Dim sMRUPath, sCmplitListPath
	Set objOccGroup = JavaWindow("Manufacturing Process").JavaWindow("New Occurrence Group")
	Fn_MPP_OccurrenceGroupCreate = False
			
	If objOccGroup.Exist(5) = False Then
		bReturn = Fn_MenuOperation("Select","File:New:Occurrence Group...")
		Call Fn_ReadyStatusSync(5)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > Occurrence Group... ] of Function Fn_MPP_OccurrenceGroupCreate.")
			Set objOccGroup = nothing
			Exit Function
		End If
		If objOccGroup.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_MPP_OccurrenceGroupCreate ] Failed to display Occurrence Group window.")
			Set objOccGroup = nothing
			Exit Function
		End if 
	End If
       	Select Case sAction
		Case "Create"
			If sOGText <> "" Then
				objOccGroup.JavaEdit("Occurrence Group Type").Set sOGText
			End If
			' selecting type from tree
			If Trim(sOGType) <> "" Then
				'	aStructureContextType = Split(sOGType,":",-1,1)
				'	objOccGroup.JavaEdit("Occurrence Group Type").Type aStructureContextType(1)
				'	iRows = objOccGroup.JavaTree("Occurrence Group Type").GetROProperty("items count")
				'	For iCount = 0 to iRows-1
				'		If InStr(1,Trim(Lcase(objOccGroup.JavaTree("Occurrence Group Type").GetItem(iCount))),Trim(Lcase(aStructureContextType(1)))) <> 0 Then		
				'			sReturn = objOccGroup.JavaTree("Occurrence Group Type").GetItem(iCount)
				'			aReturn = Split(sReturn,":",-1,1)		
				'			Call Fn_JavaTree_Select("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Occurrence Group Type",aReturn(0))		
				'			Call Fn_JavaTree_Select("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Occurrence Group Type",sReturn)
				'			Call Fn_JavaTree_Select("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Occurrence Group Type",aReturn(0))		
				'				Call Fn_JavaTree_Select("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Occurrence Group Type",sReturn)
				'			Exit For
				'		End If
				'	Next
				aOccurrenceGroupType = Split(sOGType,":",-1,1)
				sMRUPath =  "Most Recently Used:" & aOccurrenceGroupType(UBound(aOccurrenceGroupType))
				sCmplitListPath = "Complete List:" & aOccurrenceGroupType(UBound(aOccurrenceGroupType))
				If Fn_JavaTree_NodeIndexExt("Fn_MPP_OccurrenceGroupCreate",objOccGroup,"Occurrence Group Type", "Complete List" , "", "") <> -1 then
					Call Fn_UI_JavaTree_Expand("Fn_MPP_OccurrenceGroupCreate",objOccGroup,"Occurrence Group Type", "Complete List")
				end if
				
				If Fn_JavaTree_NodeIndexExt("Fn_MPP_OccurrenceGroupCreate",objOccGroup,"Occurrence Group Type", "Most Recently Used" , "", "") <> -1 then
					Call Fn_UI_JavaTree_Expand("Fn_MPP_OccurrenceGroupCreate",objOccGroup,"Occurrence Group Type", "Most Recently Used")
				end if
				
				If Fn_JavaTree_NodeIndexExt("Fn_MPP_OccurrenceGroupCreate",objOccGroup,"Occurrence Group Type", sMRUPath , "", "") <> -1 then
					Call Fn_JavaTree_Select("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Occurrence Group Type",sMRUPath)
				elseif Fn_JavaTree_NodeIndexExt("Fn_MPP_OccurrenceGroupCreate",objOccGroup,"Occurrence Group Type", sCmplitListPath , "", "") <> -1 then
					Call Fn_JavaTree_Select("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Occurrence Group Type",sCmplitListPath)
				else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_MPP_OccurrenceGroupCreate ] Collaboration Context Type [ " & UBound(aOccurrenceGroupType) & " ] is not present in the List tree.")
					Set objOccGroup = nothing
					Fn_MPP_OccurrenceGroupCreate = False
					Exit function
				end if
			End If

			' clicking checkbox open on create
			If bOpenOnCreate <> "" then
				If cbool(bOpenOnCreate)  Then
					call Fn_CheckBox_Set("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Open On Create","ON")
				Else
					call Fn_CheckBox_Set("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Open On Create","OFF")
				End If
			end if 
			' clicking on next
			Call Fn_Button_Click("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Next")
			' set name
			Call Fn_ReadyStatusSync(5)
			If sOGName <> "" Then
				Call Fn_Edit_Box("Fn_MPP_OccurrenceGroupCreate", objOccGroup,"Name", sOGName )
				If cInt(objOccGroup.JavaButton("Finish").getROProperty("enabled")) <> 1 Then
					objOccGroup.JavaEdit("Name").Activate
				End If
			End If
			' set description
			If sOGDescription <> "" Then
				Call Fn_Edit_Box("Fn_MPP_OccurrenceGroupCreate", objOccGroup,"Description", sOGDescription )
			End If
			
			' clicking on finish
			Call Fn_Button_Click("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Finish")

			if objOccGroup.exist(5) then
			' clicking on Close
				Call Fn_Button_Click("Fn_MPP_OccurrenceGroupCreate", objOccGroup, "Close")
			end if

			Fn_MPP_OccurrenceGroupCreate = true
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_MPP_OccurrenceGroupCreate ] Invalid case [ " & sAction & " ].")
			Set objOccGroup = nothing
			Exit Function
	End Select
	If Fn_MPP_OccurrenceGroupCreate = true Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_MPP_OccurrenceGroupCreate ] Executed successfully with case [ " & sAction & " ].")
	End If
	Set objOccGroup = nothing
End Function
'*********************************************************	  Function to select  the Tab into Manufacturing Process Planner ***********************************************************************

'Function Name	      	         :			Fn_MPP_TabOperations

'Description		             :		 	    This function is used to select  the Tab into Manufacturing Process Planner 

'Parameters			  :	 		   1.  String - StrTabName :Name of the Tab to be selected.
											
'Return Value		            :  		             True / False

'Pre-requisite			   :		 	    Manufacturing Process Planner window should be displayed .

'Examples				   :			    Fn_MPP_TabOperations("Activate", "Base View")

'History:
'					Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Koustubh W 				 20-Apr-2010 	          1.0				created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Koustubh W 				 17-Nov-2011 	          1.0				Added new cases Close, Select, DoubleClick
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Koustubh W 				 21-Nov-2011 	          1.0				Added new case Exist, modified cases Select, DoubleClick
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_TabOperations(sAction, StrTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_TabOperations"
'	Dim objTab,objWindow, bReturn, objTabWidget, iIndexCounter, iXposition
'	iXposition = 0
'	' initialization
	Dim objSelectType,objIntNoOfObjects,icount,iItemCount,iCounter, objItem,iIndex,objWindow
	Fn_MPP_TabOperations = False
	bFlag=False
'	Set objWindow = JavaWindow("Manufacturing Process")
'	If sAction = "Activate" Then ' - Not working in TC9.1
'			Set objTab = JavaWindow("Manufacturing Process").JavaTab("ViewAll")
'			'Synchronization
'			If objTab.Exist(iTimeOut) Then
'				bReturn = Fn_UI_JavaTab_Select("Fn_MPP_TabOperations",objWindow,"ViewAll",StrTabName)
'				If bReturn Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_TabOperations ] MPP Tab set to [" + StrTabName + "].")
'					Fn_MPP_TabOperations = True
'				End If
'			End If
'	Else
'		Set objTabWidget = JavaWindow("Manufacturing Process").JavaObject("RACTabFolderWidget")
'		Select Case sAction
'			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'			Case "Select", "DoubleClick"
'				For iIndexCounter = 0 to 2 ' assuming that there are max 3 tabs are opened. 
'					objTabWidget.setTOProperty "Index", iIndexCounter
'					If objTabWidget.Exist(5) Then
'						iTabsCount = cInt(objTabWidget.Object.getTabItemCount)
'						For iCounter = 0 to iTabsCount -1
'								Set objTab = objTabWidget.Object.getItem(iCounter)
'								iXposition = iXposition + (objTab.getWidth)
'								If  trim(objTab.Text) = StrTabName Then
'										iXposition = iXposition - (objTab.getWidth/2)
'											sBounds = objTab.getCloseButtonBounds.toString()
'											sBounds = right(sBounds, Len(sBounds)-instr(sBounds, "{"))
'											aBounds = split(sBounds, ",", -1, 1)
'											iXposition = Cint(trim(aBounds(0))) - 5
'											iYposition = Cint(trim(aBounds(1))) + 5
'										If sAction = "DoubleClick" Then
'											objTabWidget.DblClick iXposition,  iYposition, "LEFT"
'										Else
'											objTabWidget.Click iXposition,   iYposition, "LEFT"
'										End If
'										Fn_MPP_TabOperations = True
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_MPP_TabOperations : Successfully clicked on [ " & sTab & " ] tab.")
'										bFlag = True
'										Exit for
'								End If
'						Next
'					Else
'					' tab does not exist
'						Exit for
'					End If
'					If bFlag = True Then
'						' if Tab found then exit from loop
'						Exit for
'					End If
'				Next
'			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'			Case "Exist"
'				For iIndexCounter = 0 to 2 ' assuming that there are max 3 tabs are opened. 
'					objTabWidget.setTOProperty "Index", iIndexCounter
'					If objTabWidget.Exist(5) Then
'						iTabsCount = cInt(objTabWidget.Object.getTabItemCount)
'						For iCounter = 0 to iTabsCount -1
'								Set objTab = objTabWidget.Object.getItem(iCounter)
'								iXposition = iXposition + (objTab.getWidth)
'								If  trim(objTab.Text) = StrTabName Then
'										Fn_MPP_TabOperations = True
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_MPP_TabOperations : Successfully verified existence [ " & sTab & " ] tab.")
'										bFlag = True
'										Exit for
'								End If
'						Next
'					Else
'					' tab does not exist
'						Exit for
'					End If
'					If bFlag = True Then
'						' if Tab found then exit from loop
'						Exit for
'					End If
'				Next
'			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'			Case "Close"
'				For iIndexCounter = 0 to 2 ' assuming that there are max 3 tabs are opened. 
'					objTabWidget.setTOProperty "Index", iIndexCounter
'					If objTabWidget.Exist(5) Then
'						iTabsCount = cInt(objTabWidget.Object.getTabItemCount)
'						For iCounter = 0 to iTabsCount -1
'								Set objTab = objTabWidget.Object.getItem(iCounter)
'								iXposition = iXposition + (objTab.getWidth - 15)
'								If  trim(objTab.Text) = StrTabName Then
'										' selecting tab
'										iXposition = iXposition - (objTab.getWidth/2)
'										objTabWidget.Click iXposition, (objTab.getHeight/2), "LEFT"
'										' fetching coordinates of close button
'										sBounds = objTab.getCloseButtonBounds.toString()
'										sBounds = right(sBounds, Len(sBounds)-instr(sBounds, "{"))
'										aBounds = split(sBounds, ",", -1, 1)
'										iXposition = Cint(trim(aBounds(0))) + 5
'										iYposition = Cint(trim(aBounds(1))) + 5
'										' clicking on close X 
'										objTabWidget.Click iXposition, iYposition, "LEFT"
'										Fn_MPP_TabOperations = True
'										bFlag = True
'										Exit for
'								End If
'						Next
'					Else
'					' tab does not exist
'						Exit for
'					End If
'					If bFlag = True Then
'						' if Tab found then exit from loop
'						Exit for
'					End If
'				Next
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'		End Select
'	End If

	Set objWindow = JavaWindow("Manufacturing Process")
	Set objSelectType = description.Create()
	objSelectType("Class Name").value = "JavaTab"
	objSelectType("toolkit class").value = "org.eclipse.swt.custom.CTabFolder|com.teamcenter.rac.ms.ui.tab.TabFolderViewer\$1"						
	Set  objIntNoOfObjects = objWindow.ChildObjects(objSelectType)

	Select Case sAction
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select", "Activate"
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(StrTabName) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							objIntNoOfObjects(icount).Select StrTabName
							bFlag=True
							Exit For 
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		Case "DoubleClick"
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
						iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
						For iCounter = 0 To iItemCount- 1 Step 1
							If trim(StrTabName) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
								objIntNoOfObjects(icount).Select StrTabName
								iIndex=objIntNoOfObjects(icount).Object.getSelectionIndex
								Set objItem=objIntNoOfObjects(icount).Object.getItem(iIndex)
										sBounds = objItem.getBounds().toString()
										sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
										aBounds = split(sBounds,",")
										X = cInt(trim(aBounds(0)))
										H = cInt(trim(aBounds(3)))
										sxLen = X + 15
										syLen = (H/2)
									objIntNoOfObjects(icount).DblClick sxLen,syLen,"LEFT"
									wait 2
									bFlag=True
									Exit For 
							End IF
						Next
						If bFlag=True Then Exit For 
					Next
			
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Exist"
				For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(StrTabName) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
						  	bFlag=True
						  	Exit For 
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Close"
			For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
					iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
					For iCounter = 0 To iItemCount- 1 Step 1
						If trim(StrTabName) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
							objIntNoOfObjects(icount).Select StrTabName
							wait 1
							objIntNoOfObjects(icount).CloseTab StrTabName
							bFlag=True
							Exit For 
						End IF
					Next
					If bFlag=True Then Exit For 
				Next
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
	End Select

	If bFlag = True Then
		Fn_MPP_TabOperations = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_TabOperations ] Executed successfully.")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_MPP_TabOperations ] Execution failed.")
	End If
	
	Set  objItem=Nothing
	Set  objIntNoOfObjects=Nothing
	Set objSelectType=Nothing
	Set objWindow = nothing
	
'	Set objTab = nothing
End Function
'*********************************************************	Function to perform operations on Collaboration Context Tree ***********************************************************************

'Function Name		       :	Fn_MPP_CCTreeOperations

'Description			 :  This function is to perform operations on Collaboration Context Tree

'Parameters			:  1.  String - Action ( Verify )
'					   2.  String - sNodeName - ( : ) separated path of the Node
'					   3.  String - sMenu -  ( : ) separated menu
											
'Return Value		         :  True / False

'Pre-requisite			:  Manufacturing Process Planner perspective should be already set

'Examples				:  Call Fn_MPP_CCTreeOperations("Expand", "CC", "")
'				         :  Call Fn_MPP_CCTreeOperations("Exist", "CC:Top_Process", "")
'				         :  Call Fn_MPP_CCTreeOperations("Select", "CC:Top_Process", "")

'History:
'		Developer Name			Date				Rev. No.	Reviewer	Changes Done		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh W			25-Apr-2011		        1.0
'		sachin				27-Apt-2011						
'		Koustubh W			29-Apr-2011		        1.0						Removed code to set attached text to CCTree
'		Koustubh W			17-Jul-2012		        1.0						Modified code to get Tree Path
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MPP_CCTreeOperations(sAction, sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_CCTreeOperations"
	Dim objWindow, bRet
	Set objWindow = JavaWindow("Manufacturing Process")
	Fn_MPP_CCTreeOperations = False
	If objWindow.JavaTree("CCTree").exist(5) =  False Then
		Call Fn_ToolbarOperation("Click", "Open Collaboration Context Tree","")
		Call Fn_ReadyStatusSync(3)
		If objWindow.JavaTree("CCTree").exist(15) = False  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_CCTreeOperations ] Failed to open Collaboration Context Tree." )
			Set objWindow = nothing
			Exit function
		end if
	End If

	Select Case sAction
		' select case
		Case "Select"
			bRet = Fn_UI_JavaTreeGetItemPathExt("Fn_MPP_CCTreeOperations",objWindow.JavaTree("CCTree"), sNodeName, "", "")
			If bRet <> False Then
				Fn_MPP_CCTreeOperations = Fn_JavaTree_Select("Fn_MPP_CCTreeOperations", objWindow, "CCTree",bRet)
			end if
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		' expand case
		Case "Expand"
			bRet = Fn_UI_JavaTreeGetItemPathExt("Fn_MPP_CCTreeOperations",objWindow.JavaTree("CCTree"), sNodeName, "", "")
			If bRet <> False Then
				Fn_MPP_CCTreeOperations = Fn_UI_JavaTree_Expand("Fn_MPP_CCTreeOperations",objWindow,"CCTree",bRet)
			end if
                              
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		' exist case
		Case "Exist"
			bRet = Fn_UI_JavaTreeGetItemPathExt("Fn_MPP_CCTreeOperations",objWindow.JavaTree("CCTree"), sNodeName, "", "")
			If bRet <> False Then
				Fn_MPP_CCTreeOperations = True
			end if
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case else
                              Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_CCTreeOperations ] FAIL : Invalid Case [ " & sAction & " ]" )
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If Fn_MPP_CCTreeOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Fn_MPP_CCTreeOperations ] PASS : Executed successfully with case [ " & sAction & " ]" )
	End If
	Set objWindow = nothing
End Function
'*********************************************************		Function to Get BOM Table Column Index in MPP ***********************************************************************

'Function Name		:					Fn_MPP_BOMTable_ColIndex

'Description			 :		 		  This function is used to get the BOM Table Node Index.

'Parameters			   :	 			1.  StrColName:Name of the Col to retrieve Index for.
											
'Return Value		   : 				 Col index

'Pre-requisite			:		 		Structure Manager window should be displayed .

'Examples				:				Fn_MPP_BOMTable_ColIndex("All Notes")

'History:
'		Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh				05-May-2011			1.0				Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh				21-jun-2012			1.0				Modified object hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_BOMTable_ColIndex(StrColName)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_BOMTable_ColIndex"
	On Error Resume Next
	Dim IntCols , IntCounter, ObjTable, StrColIndex, StrName
	Call Fn_SISW_MPP_SetIndexMPPAppletFromTab()
	'Verify that PSE BOM Table is displayed
	If Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").Exist(iTimeOut) Then

		'Get the No. of cols present in the BOM Table

		IntCols = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").GetROProperty("cols")
		Set ObjTable = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").Object
	
		'Get the Col No. of required Column
		For IntCounter = 0 to IntCols -1
			StrName = ObjTable.getColumnName(IntCounter)
		  
			If Trim(StrName) = Trim(StrColName) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: The Column Index for Column [" + StrColName + "] is [" +IntCounter +"] in MPP BOM Table")
				Fn_MPP_BOMTable_ColIndex = IntCounter
				Exit For
			End If
		Next
		If Cint(IntCounter) = Cint(IntCols) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"WARNING: The Column [" + StrColName + "] dose not exist in MPP BOM table." )
			Fn_MPP_BOMTable_ColIndex=-1
		End If

		'Release the Table object
	   set ObjTable = Nothing
	End If
End Function
'*********************************************************		Function to Get BOM Table Column Index in MPP ***********************************************************************

'Function Name		:					Fn_MPP_BOMTable_ColumnOperations

'Description			 :		 		  This function is used to get the BOM Table Node Index.

'Parameters			   :	 			1.  StrAction : Action to be performed.
'									2.  StrColName : Name of the Col to retrieve Index for.
											
'Return Value		   : 				 Col index

'Pre-requisite			:		 		MPP window should be displayed .

'Examples				:				Fn_MPP_BOMTable_ColumnOperations("Add", "All Notes")
'									Fn_MPP_BOMTable_ColumnOperations("Remove", "All Notes")

'History:
'		Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh				05-May-2011			1.0				Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Koustubh				21-Jun-2012			1.0				modified object hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MPP_BOMTable_ColumnOperations(StrAction,StrColName)
		GBL_FAILED_FUNCTION_NAME="Fn_MPP_BOMTable_ColumnOperations"
		Dim PopUpMenu,iColIndex,ObjTable,ArrCol,iIndex,sColToAdd, objList, intCol, objChangeColumnDialog,IntCounter,StrName,bFlag,IntCols,iColNo,iCnt
		Dim objApplet
		Call Fn_SISW_MPP_SetIndexMPPAppletFromTab()
		Set ObjTable = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable")
		
		Select Case StrAction
				Case "ColumnExists"
						bFlag=False
						iColNo = cInt(ObjTable.GetROProperty("cols"))
						For iCnt = 0 to iColNo -1
							If trim(ObjTable.Object.GetColumnName(iCnt)) = trim(strColName) then
								bFlag=True
								Exit for
							end if
						Next
						If bFlag=True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Column ["& strColName &"] exists in the Application" )
							Fn_MPP_BOMTable_ColumnOperations = True
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Warning: Column ["& strColName &"] does not  exist in the Application" )
							Fn_MPP_BOMTable_ColumnOperations = False
						End If
						
				Case "Add"
						ArrCol = Split(StrColName,":",-1,1)
						For iIndex = 0 To Ubound(ArrCol)
								'Check that Column is present in the BOMTable.
								iColIndex =  Fn_MPP_BOMTable_ColIndex(ArrCol(iIndex))
								If iColIndex = -1 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Warning: Column does not  exist in the Application.Need to Add Column ["& ArrCol(iIndex) &"]." )
										sColToAdd = sColToAdd +":"+ArrCol(iIndex)
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Column ["& ArrCol(iIndex) &"] exists in the Application" )
										Fn_MPP_BOMTable_ColumnOperations =TRUE
								End if
						Next
						If sColToAdd <>""  Then
								sColToAdd = Mid(sColToAdd, 2,Len(sColToAdd))
								ArrCol = Split(sColToAdd,":",-1,1)
								'Invoke Choose Column Window if it is not present on the screen
								Set objApplet = Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame")
								objApplet.SetTOProperty "Index", 1
								Set objChangeColumnDialog = objApplet.JavaDialog("Change Columns")
								If NOT objChangeColumnDialog.Exist( 1)  Then
										ObjTable.SelectColumnHeader "#1","RIGHT"  
										Wait(2)
										JavaWindow("Manufacturing Process").JavaMenu("label:=Insert column\(s\) ...").Select 										       
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
										Set objList = objChangeColumnDialog.JavaList("ListAvailableCols").Object
								End If
								'Synchronization for Ready state
								Call Fn_ReadyStatusSync(2)							
								For iIndex = 0 To Ubound(ArrCol)							
										'Select Col to be added from the lsit
										intCol = objChangeColumnDialog.JavaList("ListAvailableCols").GetItemIndex(ArrCol(iIndex))
										objList.ensureIndexIsVisible intCol
										objChangeColumnDialog.JavaList("ListAvailableCols").ExtendSelect ArrCol(iIndex)
								Next
								' Hit  Add Column  Button after every Column selection
								Call Fn_Button_Click("Fn_MPP_BOMTable_ColumnOperations", objChangeColumnDialog, "Add")

								' Hit  Apply Button after selection
								Call Fn_Button_Click("Fn_MPP_BOMTable_ColumnOperations", objChangeColumnDialog, "Apply")
								
								Call Fn_Button_Click("Fn_MPP_BOMTable_ColumnOperations", objChangeColumnDialog, "Cancel")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Added  Column  ["& sColToAdd &"] in BOMTable")									
								Fn_MPP_BOMTable_ColumnOperations = TRUE	
								objApplet.SetTOProperty "Index", 4								
						End If
						
			Case "Remove"
						ArrCol = Split(StrColName,":",-1,1)
						For iIndex = 0 To Ubound(ArrCol)										
								'Check that Column is present in the BOMTable
								iColIndex =  Fn_MPP_BOMTable_ColIndex(ArrCol(iIndex))						
								If iColIndex = -1 Then							
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"WARNING:Column dose not  exist in the Application.No Need to Remove Column ["& ArrCol(iIndex) &"]")
										Fn_MPP_BOMTable_ColumnOperations  = FALSE
								Else
										'Remove the given Column From the BOMTable.													
										ObjTable.SelectColumnHeader iColIndex,"RIGHT"
										JavaWindow("Manufacturing Process").JavaMenu("label:=Remove this column").Select											
'										JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Remove Column").JavaButton("Yes").Click	
										JavaWindow("Manufacturing Process").JavaWindow("Remove Column").JavaButton("Yes").Click		
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Successfully removed Column  ["& ArrCol(iIndex) &"] from BOMTable.")          																
										Fn_MPP_BOMTable_ColumnOperations  =TRUE										 						
								End if
						Next
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Invalid case  ["& StrAction &"].")
				Fn_MPP_BOMTable_ColumnOperations = False
		End Select
		 
		Set objList = Nothing
		Set ObjTable = Nothing
		Set objChangeColumnDialog = nothing
 End Function
'*********************************************************		Function to create detail Operation		***********************************************************************
'Function Name		:        Fn_MPP_OperationDetailCreateDic  

'Description	    	:        Creates an Operation with detail information

'Parameters		     :    		sOperationType: Operation type to be selected
'			                         	sOperationID: Unique ID for the Operation [if non-empty, then enter]
'							          	sOperationRevID: Revision of the Operation [if non-empty, then enter] - if any one of the fields (id/rev) are blank then click Assign button
'									 	sOperationName: Name of the Operation
'									  	sOperationDesc: Description of the Operation
' 										dicOperationDetailsCreate : Dictionary paramter  for detail creation
'	Example						Set dicOperationDetailsCreate = CreateObject( "Scripting.Dictionary" )
'										dicOperationDetailsCreate.RemoveAll
'											dicOperationDetailsCreate("OperationAddInfo") = "yes"
'											dicOperationDetailsCreate("ProgramID") = "123"
'											dicOperationDetailsCreate("ProcessComm") = "Testing OK"
'											dicOperationDetailsCreate("TargetTime") = "1200"
'											dicOperationDetailsCreate("TargetCost") = "300"	
'										Msgbox Fn_MPP_OperationDetailCreateDic("MEOP", "", "", "TestOperation", "Testing", dicOperationDetailsCreate)
'										Set dicOperationDetailsCreate = Nothing
'Return Value		: 			OperationID-OperationRevID 

'Pre-requisite	    :		 	Should be logged in

'History		    :		
'													Developer Name				Date						Rev. No.			
'--------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje				     05/09/2011			           1.0								
'--------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_OperationDetailCreateDic(sOperationType, sOperationID, sOperationRevID, sOperationName, sOperationDesc, dicOperationDetailsCreate)
		GBL_FAILED_FUNCTION_NAME="Fn_MPP_OperationDetailCreateDic"

		Dim sAssignId, sAssignRevId, aProjectName, objDialogNewOperation, ObjStaticText, bReturn,sTitle

		On Error Resume Next
		'sTitle = Window("TeamcenterWindow").JavaDialog("New Item").GetTOProperty("title")
		sTitle = JavaWindow("Manufacturing Process").JavaWindow("New Process").GetTOProperty("title")
		'Creating Object for New Operation window.
		'Window("TeamcenterWindow").JavaDialog("New Item").SetTOProperty "title","New Operation"
		JavaWindow("Manufacturing Process").JavaWindow("New Process").SetTOProperty "title","New Operation"
		'set objDialogNewOperation = Window("TeamcenterWindow").JavaDialog("New Item")
		set objDialogNewOperation = JavaWindow("Manufacturing Process").JavaWindow("New Process")
		 'Creating Object of links on the left side of the window
		Set ObjStaticText =objDialogNewOperation.JavaStaticText("Stpes")

		If not objDialogNewOperation.Exist (5)  Then
			'Select menu [File -> New -> Operation...]
			bReturn = Fn_MenuOperation("Select","File:New:Operation...")
			Wait(10)
			If bReturn = False Then
					Fn_MPP_OperationDetailCreateDic = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [File:New:Operation]")
					Set objDialogNewOperation = Nothing
					Exit Function
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [File:New:Operation]")
			End If
		End If
	   'Check the existence of the "NewOperation" Window
		If objDialogNewOperation.Exist (20)  Then
					'Select  "Operation Type"
					'objDialogNewOperation.JavaList("ItemType").Select sOperationType
			iItemCount=Fn_UI_Object_GetROProperty("Fn_MPP_OperationDetailCreateDic",objDialogNewOperation.JavaTree("Enter Additional Process"), "items count")
			For iCount=0 To iItemCount-1
				strItem=objDialogNewOperation.JavaTree("Enter Additional Process").GetItem(iCount)
				If Trim(strItem)="Most Recently Used:"+Trim(sOperationType) Then
					bFlag=True
					Exit For
				ElseIf Trim(strItem)="Complete List" Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Call Fn_JavaTree_Select("Fn_MPP_OperationDetailCreateDic", objDialogNewOperation, "Enter Additional Process","Most Recently Used")
				Call Fn_JavaTree_Select("Fn_MPP_OperationDetailCreateDic", objDialogNewOperation, "Enter Additional Process","Most Recently Used:"+sOperationType)
			Else
				Call Fn_UI_JavaTree_Expand("Fn_MPP_OperationDetailCreateDic", objDialogNewOperation, "Enter Additional Process","Complete List")
				Call Fn_JavaTree_Select("Fn_MPP_OperationDetailCreateDic", objDialogNewOperation, "Enter Additional Process","Complete List")
				Call Fn_JavaTree_Select("Fn_MPP_OperationDetailCreateDic", objDialogNewOperation, "Enter Additional Process","Complete List:"+sOperationType)	
			End If
					If Err.Number < 0 Then
							Fn_MPP_OperationDetailCreateDic = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Operation Type [" + sOperationType + "]")
							Set objDialogNewOperation = Nothing
							Exit Function
					End If
					'Click on "Next" button
					objDialogNewOperation.JavaButton("Next").WaitProperty "enabled", 1, 200000
					objDialogNewOperation.JavaButton("Next").Click
					If Err.Number < 0 Then
							Fn_MPP_OperationDetailCreateDic = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
							Set objDialogNewOperation = Nothing
							Exit Function
					End If
					'Enter Operation ID
					If sOperationID <> "" Then
						 'objDialogNewOperation.JavaEdit("ItemID").Set sOperationID
						 objDialogNewOperation.JavaEdit("NewWorkAreaID").Set sOperationID
					End If
					'Enter Revision ID
					If sOperationRevID <> "" Then
						'objDialogNewOperation.JavaEdit("RevisionID").Set sOperationRevID
						objDialogNewOperation.JavaEdit("NewWorkAreaRev").Set sOperationRevID
					End If
					'Check  "Operation Id and Revision ID"
					If sOperationID = "" or sOperationRevID = "" Then
							'Click on "Assign" button
							objDialogNewOperation.JavaButton("Assign").SetTOProperty "Index","0"
							objDialogNewOperation.JavaButton("Assign").WaitProperty "enabled", 1, 20000
							objDialogNewOperation.JavaButton("Assign").Click
							objDialogNewOperation.JavaButton("Assign").SetTOProperty "Index","1"
							objDialogNewOperation.JavaButton("Assign").WaitProperty "enabled", 1, 20000
							objDialogNewOperation.JavaButton("Assign").Click
							Wait 2
					End If
					'Extract Operation Id and Rev Id
					'sAssignId = objDialogNewOperation.JavaEdit("ItemID").GetROProperty("value")
					'sAssignRevId = objDialogNewOperation.JavaEdit("RevisionID").GetROProperty("value")
					sAssignId = objDialogNewOperation.JavaEdit("ID").GetROProperty("value")
					sAssignRevId = objDialogNewOperation.JavaEdit("NewWorkAreaRev").GetROProperty("value")
					If sAssignRevId = "" and sOperationRevID = ""Then					
						sAssignRevId = "A"
					End If
					'Set the Operation Name
					If sOperationName <> "" Then
						'objDialogNewOperation.JavaEdit("ItemName").Set sOperationName
						objDialogNewOperation.JavaEdit("Name").Set sOperationName
					End If
					If sOperationDesc <> "" Then
						objDialogNewOperation.JavaEdit("Description").Set sOperationDesc
					End If
		Else
					Fn_MPP_OperationDetailCreateDic = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[New Operation] Dialog Not Found")
		End If

		'Code for details should be written as required.

		'Click on "Finish" button
		objDialogNewOperation.JavaButton("Finish").WaitProperty "enabled" , 1, 20000
		objDialogNewOperation.JavaButton("Finish").Click
		'Click on "Close" button
		objDialogNewOperation.JavaButton("Close").WaitProperty "enabled" , 1, 20000
		objDialogNewOperation.JavaButton("Close").Click
		Call Fn_ReadyStatusSync(2)
		Fn_MPP_OperationDetailCreateDic = "'"&sAssignId+"-"+sAssignRevId
		objDialogNewOperation.SetTOProperty "title",sTitle
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Operation [" + sOperationName + "] Created Successfully")

End Function
'****************************************		Function to perform operations on Attachments Table.  ***************************************

'Function Name		      :			  Fn_MPP_AttachmentTableNodeOperation  

'Description			     :  	      Function to perform operations on Attachments Table.

'Parameters			   		:	   	  1.  String - sAction
'						          			2.  String - sBOMLine
'	                					         	3.  String - sAttchLine,
' 							          		4.  String - sCol
'										5.  String - sData
											
'Return Value		       : 			 For	Case Select  : True / False
'								   			Case Exist     : True / False
'								   			Case GetCell : Cell Data / False
'								   			Case SetCell : True / False
'											Case Remove: True / False

'Pre-requisite			    :		 	  Manufacturing Process Planner window should be displayed .
'								  

'Examples				    :			Call Fn_MPP_AttachmentTableNodeOperation("Select","000266/A;1-asm1", "000612/A","","")
'						     			Call Fn_MPP_AttachmentTableNodeOperation("Select","000266/A;1-asm1", "000612 @2","","")
'						     			Call Fn_MPP_AttachmentTableNodeOperation("MultiSelect","000266/A;1-asm1", "000612/A~000612 @2","","")
'								 	    Call Fn_MPP_AttachmentTableNodeOperation("Exist","000266/A;1-asm1", "000612/A","Line","")
'								 	    Call Fn_MPP_AttachmentTableNodeOperation("Copy","000266/A;1-asm1", "000612/A","Line","")
'								 	    Call Fn_MPP_AttachmentTableNodeOperation("Paste","000266/A;1-asm1", "000612/A","Line","")
'								 	    Call Fn_MPP_AttachmentTableNodeOperation("GetCell","000266/A;1-asm1","000612/A","Description","")
'								 	    Call Fn_MPP_AttachmentTableNodeOperation("SetCell","000266/A;1-asm1","000612/A","Description","Data")
'									    Call Fn_MPP_AttachmentTableNodeOperation("CellDoubleClick","000266/A;1-asm1","000612/A","","")
'									    Call Fn_MPP_AttachmentTableNodeOperation("PopupSelect","000016/A;1-Top (View)", "000016/A;1-Top @3","","Expand Below")
'									    Call Fn_MPP_AttachmentTableNodeOperation("Expand","000016/A;1-Top (View)", "View","","")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W			  17-Nov-2011		  1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W			  21-Nov-2011		  1.0					Modified function to open and close Attachment tab. it is necessary because 
'																							Attachment table matches with CMEBOMTreeTable.
'																							Added cases Copy and Paste 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Vrushali W			  11-Jun-2013		  1.0					Added case SelectWithoutClose
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_AttachmentTableNodeOperation(sAction, sBOMLine, sAttchLine, sCol, sData)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_AttachmentTableNodeOperation"
	Dim bReturn, objBOMTable, objDataTabPane, objAttachmentTable
	Dim i, iCounter, iRowCount, iColCount, iLineColumnNumber, iColCounter, sTemp
	Dim iColNum,objRemove, aAttchLines, sLine, bFound
	Dim iInstance, aNodePath,  StrNodeName, instanceCnt, aMenu, strMenu

	'Function Return False
	Fn_MPP_AttachmentTableNodeOperation = False

'	JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").setTOProperty "Index", 0
	'Verify Attachement Table
	If Fn_MPP_TabOperations("Exist","Attachments") = True Then
		'Activating Attachement Tab
		Call Fn_ReadyStatusSync(2)
		Call Fn_MPP_TabOperations("Close","Attachments")
	End If

	If sBOMLine <>"" Then
		bReturn = Fn_MPP_BOMTable_NodeOperation("Select", sBOMLine,"","","")
		If bReturn  = False Then
			Exit function
		End If
	End If
	' opening Attachement View
	Call Fn_SetView("Manufacturing:Attachments")
	Call Fn_ReadyStatusSync(2)
	Call Fn_MPP_TabOperations("DoubleClick","Attachments")
	Call Fn_ReadyStatusSync(2)
	'Creating Object of Attachment Table
	Set objAttachmentTable = JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").JavaTable("AttachmentsTreeTable")
	If objAttachmentTable.Exist(10) = False  Then
	'	JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").setTOProperty "Index", 1
		Exit function
	End If

	iCounter = 0
	iRowCount = objAttachmentTable.GetROProperty("rows")
	iColCount =  objAttachmentTable.GetROProperty("cols")

   	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Select", "Copy", "Paste","SelectWithoutClose"	
				'Mapping Object column to column number.
				iLineColumnNumber = 0
				
				If  instr(sAttchLine, "@") > 0 Then
					aNodePath = split(sAttchLine, "@",-1, 1)
					sAttchLine = trim(aNodePath(0))
					iInstance = cint(aNodePath(1))
				Else
					iInstance = 1
				End If
				'selecting specified cell
				instanceCnt = 0
				For iCounter = 0 to iRowCount -1
					sTemp = objAttachmentTable.Object.getValueAt(iCounter, iLineColumnNumber).toString
					If sTemp = sAttchLine Then
						instanceCnt = instanceCnt + 1
						If instanceCnt = iInstance Then
							wait 2
							objAttachmentTable.SelectRow iCounter
							Fn_MPP_AttachmentTableNodeOperation  = True
							Exit for
						End If
					End If
				Next
				If Fn_MPP_AttachmentTableNodeOperation  = True Then
					Select Case sAction
							Case "Copy"
								bReturn = Fn_MenuOperation("Select","Edit:Copy")
								If bReturn = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [Edit:Copy	Ctrl+C] of Function Fn_MPP_AttachmentTableNodeOperation")
										Fn_MPP_AttachmentTableNodeOperation = False
										Exit Function
								End If
							Case "Paste"
								bReturn = Fn_MenuOperation("Select","Edit:Paste")
								If bReturn = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [Edit:Copy	Ctrl+V] of Function Fn_MPP_AttachmentTableNodeOperation")
										Fn_MPP_AttachmentTableNodeOperation = False
										Exit Function
								End If
							Case Else
								' do nothing
					End Select
				End If
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "MultiSelect"
			'Mapping Object column to column number.
			iLineColumnNumber = Fn_PSE_Table_ColumnIndex(objAttachmentTable, "Line")
			aAttchLines = split(sAttchLine,"~")
			objAttachmentTable.Object.clearSelection
			for each sLine in aAttchLines
				bFound = False
				If  instr(sLine, "@") > 0 Then
					aNodePath = split(sLine, "@",-1, 1)
					sLine = trim(aNodePath(0))
					iInstance = cint(aNodePath(1))
				Else
					iInstance = 1
				End If
				'selecting specified cell
				instanceCnt = 0
				For iCounter = 0 to iRowCount -1
					sTemp = objAttachmentTable.Object.getValueAt(iCounter, iLineColumnNumber).toString
					If sTemp = sLine Then
						instanceCnt = instanceCnt + 1
						If instanceCnt = iInstance Then
							wait 2
							objAttachmentTable.ExtendRow iCounter
							bFound = True
							Exit for
						End If
					End If
				Next
				if bFound = False then exit for
			next
			Fn_PSE_AttachmentTableNodeOperation  = bFound

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Expand"
				'Mapping Object column to column number.
				iLineColumnNumber = Fn_PSE_Table_ColumnIndex(objAttachmentTable, "Line")
				
				If  instr(sAttchLine, "@") > 0 Then
					aNodePath = split(sAttchLine, "@",-1, 1)
					sAttchLine = trim(aNodePath(0))
					iInstance = cint(aNodePath(1))
				Else
					iInstance = 1
				End If
				'selecting specified cell
				instanceCnt = 0
				For iCounter = 0 to iRowCount -1
					sTemp = objAttachmentTable.Object.getValueAt(iCounter, iLineColumnNumber).toString
					If sTemp = sAttchLine Then
						instanceCnt = instanceCnt + 1
						If instanceCnt = iInstance Then
							objAttachmentTable.Object.expandNode objAttachmentTable.Object.getNodeForRow(cint(iCounter))
							Fn_MPP_AttachmentTableNodeOperation  = True
							Exit for
						End If
					End If
				Next

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "PopupSelect"
				'Mapping Object column to column number.
			iLineColumnNumber = Fn_PSE_Table_ColumnIndex(objAttachmentTable, "Line")
			
			If  instr(sAttchLine, "@") > 0 Then
				aNodePath = split(sAttchLine, "@",-1, 1)
				sAttchLine = trim(aNodePath(0))
				iInstance = cint(aNodePath(1))
			Else
				iInstance = 1
			End If
			'selecting specified cell
			instanceCnt = 0
			aMenu = split(sData,":",-1,1)

			For iCounter = 0 to iRowCount -1
				sTemp = objAttachmentTable.Object.getValueAt(iCounter, iLineColumnNumber).toString
				If sTemp = sAttchLine Then
					instanceCnt = instanceCnt + 1
					If instanceCnt = iInstance Then
						Fn_MPP_AttachmentTableNodeOperation  = True
						'objAttachmentTable.SelectRow iCounter
						objAttachmentTable.ClickCell iCounter,"Line","RIGHT"
						Select Case Ubound(aMenu)
							Case "0"
								strMenu = JavaWindow("StructureManager").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
								JavaWindow("StructureManager").WinMenu("ContextMenu").Select strMenu
							Case "1"
								strMenu = JavaWindow("StructureManager").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
								JavaWindow("StructureManager").WinMenu("ContextMenu").Select strMenu
							Case Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [Fn_MPP_AttachmentTableNodeOperation] Context Menu Case NOT Exists for Supplied Menu [" + StrPopupMenu + "]")
								Fn_MPP_AttachmentTableNodeOperation = False
						End Select
						Exit for
					End If
				End If
			Next

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Exist"
			'Mapping Object column to column number.
			iLineColumnNumber = Fn_PSE_Table_ColumnIndex(objAttachmentTable, "Line")
			Fn_MPP_AttachmentTableNodeOperation  = False
			If  instr(sAttchLine, "@") > 0 Then
				aNodePath = split(sAttchLine, "@",-1, 1)
				sAttchLine = trim(aNodePath(0))
				iInstance = cint(aNodePath(1))
			Else
				iInstance = 1
			End If
			'selecting specified cell
			instanceCnt = 0
			For iCounter = 0 to iRowCount -1
				sTemp = objAttachmentTable.Object.getValueAt(iCounter, iLineColumnNumber).toString
				If sTemp = sAttchLine Then
					instanceCnt = instanceCnt + 1
					If instanceCnt = iInstance Then
						wait 2
						Fn_MPP_AttachmentTableNodeOperation  = True
						Exit for
					End If
				End If
			Next
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "GetCell"
			'Mapping Object column to column number.
			iLineColumnNumber = Fn_PSE_Table_ColumnIndex(objAttachmentTable, "Line")
			If iLineColumnNumber = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_MPP_AttachmentTableNodeOperation ]  Failed to find column index for [ " &  "Line" &  " ]") 
				Fn_MPP_AttachmentTableNodeOperation = False
				Exit function
			End If
			iColNum = Fn_PSE_Table_ColumnIndex(objAttachmentTable, sCol)
			If iColNum = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_MPP_AttachmentTableNodeOperation ]  Failed to find column index for [ " &  sCol &  " ]") 
				Fn_MPP_AttachmentTableNodeOperation = False
				Exit function
			End If
			
			' Retrieving cell data
			For iCounter = 0 to iRowCount -1
				sTemp = objAttachmentTable.Object.getValueAt(iCounter, iLineColumnNumber).toString
				If sTemp = sAttchLine Then
					Fn_MPP_AttachmentTableNodeOperation  =  objAttachmentTable.Object.getValueAt(iCounter, iColNum).toString
'					Fn_MPP_AttachmentTableNodeOperation  = True
					Exit for
				End If
			Next

		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "SetCell"
			'Mapping Object column to column number.
			iLineColumnNumber = Fn_PSE_Table_ColumnIndex(objAttachmentTable, "Line")
			If iLineColumnNumber = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_MPP_AttachmentTableNodeOperation ]  Failed to find column index for [ " &  "Line" &  " ]") 
				Fn_MPP_AttachmentTableNodeOperation = False
				Exit function
			End If
			iColNum = Fn_PSE_Table_ColumnIndex(objAttachmentTable, sCol)
			If iColNum = -1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: Function [ Fn_MPP_AttachmentTableNodeOperation ]  Failed to find column index for [ " &  sCol &  " ]") 
				Fn_MPP_AttachmentTableNodeOperation = False
				Exit function
			End If
			'setting cell data.
			For iCounter = 0 to iRowCount -1
				sTemp = objAttachmentTable.Object.getValueAt(iCounter, iLineColumnNumber).toString
				If sTemp = sAttchLine Then
					 objAttachmentTable.SetCellData iCounter, iColNum,sData
					 Fn_MPP_AttachmentTableNodeOperation  = True
					 Exit for
				End If
			Next

          ' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "CellDoubleClick"
			If sCol = "" then
				iLineColumnNumber = Fn_PSE_Table_ColumnIndex(objAttachmentTable, "Line")
			Else
				iLineColumnNumber = Fn_PSE_Table_ColumnIndex(objAttachmentTable, sCol)
			End If
				'selecting specified cell
				For iCounter = 0 to iRowCount -1
					sTemp = objAttachmentTable.Object.getValueAt(iCounter, iLineColumnNumber).toString
					If sTemp = sAttchLine Then
					objAttachmentTable.SelectRow iCounter
					wait 1
					If trim(lcase(sAttchLine)) = "view" Then
						Dim objRect, intX, intY
						set objRect = objAttachmentTable.Object.getCellRect(iCounter, iLineColumnNumber,True)
						intX = cint(objRect.getX) + cint(objRect.getWidth) / 2
						intY = cint(objRect.getY) + cint(objRect.getHeight) / 2											
						objAttachmentTable.DblClick intX, intY,"LEFT"
						Set objRect = Nothing
					Else
						objAttachmentTable.DoubleClickCell iCounter, iLineColumnNumber
					End If

					If Err.Number < 0 Then
						Fn_MPP_AttachmentTableNodeOperation  = False
					Else
						Fn_MPP_AttachmentTableNodeOperation  = True
					End If
					Exit for
				End If
			Next
' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	'Release the Table object
	If 	inStr(sAction,"WithoutClose") = 0 Then
		Call Fn_MPP_TabOperations("Close","Attachments")
	End If
'	JavaWindow("Manufacturing Process").JavaWindow("MPPApplet").setTOProperty "Index", 1
'	 Set ObjTable = Nothing
	 Set objAttachmentTable=Nothing
End Function
'****************************************		Function to remove level of the BOMLine.  ***************************************

'Function Name		          :		   Fn_MPP_RemoveLevel  

'Description			    :  		       Remove level of the BOMLine.

'Parameters			   :	   	          1.  String - sAction ( Menu / RemoveToolbar / ShortKey )
'						          					2.  String - sBOMLine
'													3. String - sKeepSubTree ( ON  /  OFF )
											
'Return Value		       :			True / False

'Pre-requisite			    :		 	 Structure Manager window should be displayed .

'Examples				    :			  Fn_MPP_RemoveLevel("Remove","518611/A;1-Item_518611 (view):001270/A;1-ffff", "ON")
'Examples				    :			  Fn_MPP_RemoveLevel("RemoveToolbar","518611/A;1-Item_518611 (view):001270/A;1-ffff","ON")
'Examples				    :			  Fn_MPP_RemoveLevel("ShortKey","518611/A;1-Item_518611 (view):001270/A;1-ffff","OFF")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Koustubh W			  21-Nov-2011		  1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_RemoveLevel(sAction,sBOMLIne, sKeepSubTree)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_RemoveLevel"
	Dim objRemove
	Dim bReturn
	Fn_MPP_RemoveLevel = True

	If sBOMLine <> "" Then
		' selecting BOM Line from BOM Table
		bReturn = Fn_MPP_BOMTable_NodeOperation("Select",sBOMLine, "","","")
		If bReturn = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_RemoveLevel ] BOM Line [ "+ sBOMLine +" ] selected successfully") 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_RemoveLevel ]  Failed to select BOMLine[ "+ sBOMLine +" ]") 
			Fn_MPP_RemoveLevel = False
			Exit function
		End If
	End If

	Select Case sAction
		Case "Menu","Remove"
			bReturn = Fn_MenuOperation("Select", "Edit:Remove")
			If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_RemoveLevel ] Case [ Menu ] Failed to select Edit > Remove option For BOMLine[ "+ sBOMLine +" ]") 
				Fn_MPP_RemoveLevel = False
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_RemoveLevel ] Case [ Menu ] Edit > Remove option selected successfully") 
			End If
		Case "RemoveToolbar"
			bReturn = Fn_ToolbatButtonClick("Remove a line (Ctrl+R)")
			If bReturn <> True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_RemoveLevel ] Case [ RemoveToolbar ] Remove option is selected from toolbar For BOMLine[ "+ sBOMLine +" ].") 
				Fn_MPP_RemoveLevel = False
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_RemoveLevel ] Case [ RemoveToolbar ] Failed to select Remove option  from toolbar For BOMLine[ "+ sBOMLine +" ].") 
			End If
		Case "ShortKey"
			bReturn = Fn_KeyBoardOperation("SendKey", "^(r)")
			If bReturn = True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_RemoveLevel ] Case [ ShortKey ] Shortcut Keys pressed successfully For BOMLine[ "+ sBOMLine +" ].") 
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_RemoveLevel ] Case [ ShortKey ] Failed to press Shortcut Keys For BOMLine[ "+ sBOMLine +" ].") 
				Fn_MPP_RemoveLevel = False
			End If

		Case Else 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_RemoveLevel ]  Invalid Option") 
			Fn_MPP_RemoveLevel = False
	End Select

	' handling remove confirmation
	set objRemove = JavaWindow("Manufacturing Process").JavaWindow("Search").JavaDialog("Remove")
	If objRemove.exist = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_RemoveLevel ]  Failed to open Remove Dialogbox.") 
		Fn_MPP_RemoveLevel = False
		Exit function
	End If

'	If objRemove.JavaCheckBox("RemoveOption").Exist Then
'		Select Case sKeepSubTree
'			Case "ON"
'				objRemove.JavaCheckBox("RemoveOption").Set("ON")
'			Case "OFF"
'				objRemove.JavaCheckBox("RemoveOption").Set("OFF")
'		End Select
'	End If

	' clicking on Yes button
	bReturn = Fn_Button_Click("Fn_MPP_RemoveLevel",objRemove,"Yes")
	If bReturn = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_RemoveLevel ] Clicked on Yes successfully.") 
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_RemoveLevel ] Failed to click on Yes.") 
	End If
	
	' clicking on OK button
	wait(5)

	If objRemove.JavaButton("OK").Exist Then
		bReturn = Fn_Button_Click("Fn_MPP_RemoveLevel",objRemove,"OK")
		If bReturn = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_RemoveLevel ] Clicked on OK successfully.") 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_RemoveLevel ] Failed to click on OK.") 
		End If
	End If

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_RemoveLevel ] executed successfully.") 
	Set BOMTableObj = nothing
	Set objRemove = nothing
End Function
''********************************************************   End of  Fn_MPP_RemoveLevel   ****************************************************************************
'*********************************************************		Function to perform replace special on BOM Node from Manufacturing Process Planner ***********************************************************************

'Function Name		:					Fn_MPP_ObjectReplaceSpecial

'Description			 :		 		  This function is used to perform replace special on BOM Node.

'Parameters			   :	 			   1.  String - sBOMLine, 
'									   2.  String - sItemID, 
'									   3.  String - sRevID, 
'									   4.  String - sName, 
'									   5.  String - sViewType,
'									   6.  String - sReplaceOption, 
'									   7.  String - sErrorMsg
											
'Return Value		   : 				         True / False

'Pre-requisite			:		 	           Manufacturing Process Planner window should be displayed .

'Examples				:		   Fn_MPP_ObjectReplaceSpecial("000266/A;1-asm1", "000263", "", "", "CAEAnalysis","All", "")
'	 		 		:				Fn_MPP_ObjectReplaceSpecial("000266/A;1-asm1", "000263", "", "", "SingleComponent","All", "")

'History:
'									Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Koustubh W			21-Nov-2011		        1.0					Created
'									IKHALAQUE         	04/July/2012
'									Koustubh W			24-Jul-2012		        					Modified object hierarchy
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_ObjectReplaceSpecial(sBOMLine, sItemID, sRevID, sName, sViewType,sReplaceOption, sErrorMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_ObjectReplaceSpecial"
	Dim bReturn, replaceObj, errorObj
	Dim bFlag,iCounter
	bFlag=false
	Fn_MPP_ObjectReplaceSpecial = True
	Set replaceObj = JavaWindow("Manufacturing Process").JavaWindow("Search").JavaDialog("Replace")
	If replaceObj.Exist(5)  Then
		' Do Nothing
	Else
		Set replaceObj = Window("MPPWindow").JavaWindow("MPPApplet").JavaDialog("Replace")
		For iCounter=0 to 10
			Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "index",iCounter
			If Window("MPPWindow").JavaWindow("MPPApplet").JavaDialog("Replace").Exist(2) Then
				Set replaceObj =Window("MPPWindow").JavaWindow("MPPApplet").JavaDialog("Replace")
				bFlag = true
				Exit for
			End If
		Next
		If bFlag=false Then
			Fn_MPP_ObjectReplaceSpecial = false
			Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "index",0
		End If
	End If
	If replaceObj.exist(5) = False Then
		'Start From BOM Line Selection
		If sBOMLine <> "" Then
			bReturn = Fn_MPP_BOMTable_NodeOperation("Select",sBOMLine, "","","")
			If bReturn = True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_ObjectReplaceSpecial ] BOM Line [ "+ sBOMLine +" ] selected successfully") 
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_ObjectReplaceSpecial ]  Failed to select BOMLine[ "+ sBOMLine +" ]") 
				Fn_MPP_ObjectReplaceSpecial = False
				Exit function
			End If
		End If

		'Invoke menu Edit;Replace Special
		bReturn = Fn_MenuOperation("Select", "Edit:Replace...")
		If bReturn <> True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_ObjectReplaceSpecial ] Failed to select Edit > Replace...  option for BOMLine[ "+ sBOMLine +" ]") 
			Fn_MPP_ObjectReplaceSpecial = False
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_ObjectReplaceSpecial ] Edit > Replace... option selected successfully") 
		End If
	
		If JavaWindow("Manufacturing Process").JavaWindow("Search").JavaDialog("Replace").Exist(5)  Then
			Set replaceObj = JavaWindow("Manufacturing Process").JavaWindow("Search").JavaDialog("Replace")
		Else
			bFlag=false
			For iCounter=0 to 10
				Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "index",iCounter
				If Window("MPPWindow").JavaWindow("MPPApplet").JavaDialog("Replace").Exist(2) Then
					Set replaceObj =Window("MPPWindow").JavaWindow("MPPApplet").JavaDialog("Replace")
					bFlag=true
					Exit for
				End If
			Next
			If bFlag=false Then
				Fn_MPP_ObjectReplaceSpecial = false
				Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "index",0
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_ObjectReplaceSpecial ] Failed to open Replace Special dialog box.") 
				Exit Function
			End If
		End If
	End If

	'Fill in Details Like ItemID, Replace Option
	If sItemId <> ""  Then
		Call Fn_SISW_UI_JavaEdit_Operations("Fn_MPP_ObjectReplaceSpecial", "Set", replaceObj, "ItemID", sItemId)
		replaceObj.JavaEdit("ItemID").Activate
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_ObjectReplaceSpecial ] Item Id is empty.") 
		Fn_MPP_ObjectReplaceSpecial = False	
	End If

	'If Error dilog exists then handle it By verifying Error Message.

	' setting revision id
	If sRevId <> "" Then
		replaceObj.JavaList("RevisionID").Select sRevID
	End If

	' setting view type 
	If sViewType <> "" Then
		replaceObj.JavaList("ViewType").Select sViewType
	End If

	' case for replace option
	Select Case sReplaceOption
		Case "SingleComponent"
				replaceObj.JavaRadioButton("ReplaceOption_SingleComp").Set("ON")
		Case "All"
				replaceObj.JavaRadioButton("ReplaceOptionAll").Set("ON")
	End Select

	'Hit OK Button
	bReturn = Fn_Button_Click ("Fn_MPP_ObjectReplaceSpecial", replaceObj,"OK")
	Wait 2
	If Dialog("ErrorDialog").Exist = True Then
		Do 
			Call Fn_PSE_ErrorDialogHandler("","","OK")
		Loop While Dialog("ErrorDialog").Exist = True
		Call Fn_Button_Click("Fn_MPP_ObjectReplaceSpecial", replaceObj,"Cancel")
	End If
	
	If bReturn <> True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_ObjectReplaceSpecial ] Failed to click on OK button.") 
		Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "Index",0
		Fn_MPP_ObjectReplaceSpecial = False	
	Else
		Fn_MPP_ObjectReplaceSpecial = True
		Window("MPPWindow").JavaWindow("MPPApplet").SetTOProperty "Index",0
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_ObjectReplaceSpecial ] executed successfully.")
	End If
	Set replaceObj = nothing
End Function

 '*********************************************************		Function to create Note   ***********************************************************************
'Function Name		:				Fn_MPP_NoteCreate

'Description		:		 		 Creates a Note with basic information

'Parameters			:	 			1.StrNoteType - Type of Note.
'									2.StrConfNote - True or False.
'									3.StrNoteID - ID of the Note should be unique.
'									4.StrNoteRev - Revision of the Note.
'									5.StrNoteName - Name of Note.
'									6.StrNoteDesc - Description of the Note.
'									7.StrNoteUOM - Unit of measure of note.

'Return Value		 : 				StrNoteID~StrNoteRev

'Pre-requisite		 :		 		Manufacturing Process Planner should be open with a selected BOM Node

'Examples			:			   Call Fn_MPP_NoteCreate("Create","Custom Note","","","","Note_Name","Note_Desc","")

'History			:		
'									Developer Name					Date				Rev. No.				
'----------------------------------------------------------------------------------------------------------------------------------------------------
' 								   Amit Talegaonkar 			 22-Nov-2011 	         1.0					
'----------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MPP_NoteCreate( sAction, StrNoteType , StrConfNote , StrNoteID , StrNoteRev , StrNoteName , StrNoteDesc , StrNoteUOM )
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_NoteCreate"
	Dim sItemId, sRevId
	Dim objDialogNewNote
	Fn_MPP_NoteCreate = False
	'If Note TYPE is Blank, function will fail
	If StrNoteType = "" Or StrNoteName = "" Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Note TYPE/NAME is Blank")
		Exit Function
	End If
	
	'Create Object
	Set objDialogNewNote = JavaWindow("Manufacturing Process").JavaWindow("Search").JavaDialog("NewCustomNote")
	
	'Check if [ New Custom Note ] dialog is open or not
	If objDialogNewNote.Exist(2) = False Then
        Call Fn_MenuOperation("Select","File:New:Custom Note")
		Call Fn_ReadyStatusSync(2)
	End If
	
	'Confirm if its opened or not
	If objDialogNewNote.Exist(2) = False Then
		objDialogNewNote = Nothing
		Exit Function     
	End If
	
	Select Case sAction
		Case "Create"
			'Select Note Type from List
			Call Fn_List_Select( "Fn_MPP_NoteCreate" , objDialogNewNote , "NoteType" , StrNoteType )
			Call Fn_ReadyStatusSync(2)
			
			'Configuration Item
			If StrConfNote <> "" Then
				If cBool(StrConfNote) Then
				  Call Fn_CheckBox_Set("Fn_MPP_NoteCreate",objDialogNewNote,"ConfigurationItem", "ON" )
				Else
				 Call Fn_CheckBox_Set("Fn_MPP_NoteCreate",objDialogNewNote,"ConfigurationItem", "OFF" )
				End IF
				Call Fn_ReadyStatusSync(2)
			End If
			
			'Click on [ Next ] Button
			Call Fn_Button_Click("Fn_MPP_NoteCreate", objDialogNewNote ,"Next")
			Call Fn_ReadyStatusSync(2)
			
			'Enter ID
			If StrNoteID <> "" Then
				Call Fn_Edit_Box("Fn_MPP_NoteCreate",objDialogNewNote,"NoteID", StrNoteID )
				Call Fn_ReadyStatusSync(2)
			End If
				
			'Enter Revision
			If StrNoteRev <> "" Then
				Call Fn_Edit_Box("Fn_MPP_NoteCreate",objDialogNewNote,"NoteRevision", StrNoteRev )
				Call Fn_ReadyStatusSync(2)
			End If
				
			'If either StrNoteID or StrNoteRev are Blank, then click on Assign Button
			If StrNoteID = "" Or StrNoteRev = "" Then
				'Click on Assign Button
				Call Fn_Button_Click("Fn_MPP_NoteCreate", objDialogNewNote, "Assign")
				Call Fn_ReadyStatusSync(2)
			End If
			Wait(4)
			
			'Get ID and Revision
			StrNoteID = Fn_Edit_Box_GetValue("Fn_MPP_NoteCreate", objDialogNewNote,"NoteID")
			StrNoteRev = Fn_Edit_Box_GetValue("Fn_MPP_NoteCreate", objDialogNewNote,"NoteRevision")
			
			'Enter Name
			Call Fn_Edit_Box("Fn_MPP_NoteCreate",objDialogNewNote,"NoteName", StrNoteName )
			Call Fn_ReadyStatusSync(2)
			
			'Enter Description
			If StrNoteDesc <> "" Then
				Call Fn_Edit_Box("Fn_MPP_NoteCreate",objDialogNewNote,"NoteDescription", StrNoteDesc )
				Call Fn_ReadyStatusSync(2)
			End If
			
			'Enter UOM
			If StrNoteUOM <> "" Then
				Call Fn_Edit_Box("Fn_MPP_NoteCreate",objDialogNewNote,"UnitOfMeasure", StrNoteUOM )
				Call Fn_ReadyStatusSync(2)
			End If
			Wait 4
			
			'Wait property
			objDialogNewNote.JavaButton("Finish").WaitProperty "enabled", 1, 20000
			
			'Click on Finish Button 
			Call Fn_Button_Click("Fn_MPP_NoteCreate", objDialogNewNote,"Finish")
			Wait 4
			
			'Return Values
			Fn_MPP_NoteCreate = StrNoteID & "~" & StrNoteRev
			Call Fn_ReadyStatusSync(3)
			
			'If dialog Exists, click on Close Button
			If objDialogNewNote.Exist(2) Then
				'Click on Close button
				Call Fn_Button_Click("Fn_MPP_NoteCreate", objDialogNewNote, "Close") 
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Note - [" & StrNoteID & "-" & StrNoteRev & "]") 
	End Select
	Set objDialogNewNote = Nothing
End Function
'*******************************************************************************   End of  Fn_MPP_NoteCreate   **************************************************************************************************

'*********************************************************		Function to Verify IC Active   ***********************************************************************
'Function Name		:				Fn_MPP_VerifyActiveIC

'Description		:		 		 Verifies the Current Active IC

'Parameters			:	 			1.ICName -  Active IC Name 

'Return Value		 : 				True/False

'Pre-requisite		 :		 		Manufacturing Process Planner should be open with a IC Active

'Examples			:			   Call Fn_MPP_VerifyActiveIC(ICName)

'History			:		
'									Developer Name					Date				Rev. No.				
'----------------------------------------------------------------------------------------------------------------------------------------------------
' 								   Sachin Joshi 			 		09-Dec-2011 	         1.0					
'----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_VerifyActiveIC(ICName)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_VerifyActiveIC"
   Dim obj
   Fn_MPP_VerifyActiveIC = False

	'Creating object
   Set obj = JavaWindow("Manufacturing Process").JavaStaticText("ICStatic")
	'Check IC Static exists
   If obj.exist(3) = False Then
	   	Fn_MPP_VerifyActiveIC = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail no IC is Active.")
		Exit Function
   End If 

	If obj.GetROProperty("value") <> ICName Then
		Fn_MPP_VerifyActiveIC = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Verify IC ["&ICName&"] is Set to active.")
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified IC ["&ICName&"] is Set to active.")
		Fn_MPP_VerifyActiveIC = True
	End If

	Set obj = Nothing
End Function
'*******************************************************************************   End of  Fn_MPP_VerifyActiveIC   **************************************************************************************************
'*********************************************************		Function to create basic Process		***********************************************************************
'Function Name		:	Fn_SISW_MPP_SetIndexMPPAppletFromTab

'Description		:	Function to set index of Applet with the help of Tab Names

'Parameters			:	None

'Return Value		:	True / False

'Pre-requisite		:	should be Prespective to Manufacturing Process Planner

'Examples			:	Call Fn_SISW_MPP_SetIndexMPPAppletFromTab()

'History				 :		
'Developer Name						Date					Rev. No.		Reviewer					Changes Done					
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Koustubh Watwe 			 		19-Jun-2012 	         1.0			Koustubh					Created		
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_MPP_SetIndexMPPAppletFromTab()
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_MPP_SetIndexMPPAppletFromTab"
	Dim iTabIndex, sTabName, iAppletCnt, sTableTopNode, aTopNode, aTabText
	Dim objApplet
	Dim sTab_ID, sTab_RevID, sTab_Name
	Dim sTable_ID, sTable_RevID, sTable_Name
	sTab_ID = ""
	sTab_RevID = ""
	sTab_Name = ""
	Fn_SISW_MPP_SetIndexMPPAppletFromTab = False
	Set objApplet = Window("MPPWindow").JavaWindow("MPPApplet")
	
	If JavaWindow("Manufacturing Process").JavaTab("InnerViewAll").Exist(3) = True Then ' [TC11.4 - ]
		If lcase(JavaWindow("Manufacturing Process").JavaTab("InnerViewAll").GetROProperty("value")) <> "base view" Then
			sTabName = lcase(JavaWindow("Manufacturing Process").JavaTab("InnerViewAll").GetROProperty("value"))
			For iAppletCnt = 0 to 10
				objApplet.SetTOProperty  "Index",  iAppletCnt
				If Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").Exist(1) Then
					If lcase(sTabName) = lcase(objApplet.JavaTable("CMEBOMTreeTable").Object.getValueAt(0,0).toString()) Then
						Fn_SISW_MPP_SetIndexMPPAppletFromTab = True
						exit function
					End If
				End If
			Next
		End If
	End If	
	
	'iTabIndex = cInt(JavaWindow("Manufacturing Process").JavaObject("RACTabFolderWidget").Object.getSelectedTabIndex)
	'sTabName = JavaWindow("Manufacturing Process").JavaObject("RACTabFolderWidget").Object.getItem(iTabIndex ).text
	sTabName = JavaWindow("Manufacturing Process").JavaTab("ViewAll").GetROProperty("value")
	If instr(sTabName,";") > 0  Then
		sTabName = replace(sTabName,"/","-")
		sTabName = replace(sTabName,";1-","-")
		aTabText = split(sTabName, "-")
		sTab_ID = aTabText(0)
		sTab_RevID = aTabText(1)
		sTab_Name = aTabText(2)
	ElseIf instr(sTabName,"-") > 0  Then
		aTabText = split(sTabName, "-")
		sTab_ID = aTabText(0)
		sTab_Name = aTabText(1)
	Else
		' Name
		sTab_Name = sTabName
	End If

	For iAppletCnt = 0 to 10
		objApplet.SetTOProperty  "Index",  iAppletCnt

		If Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").Exist(2) Then
			sTableTopNode = Window("MPPWindow").JavaWindow("MPPApplet").JavaTable("CMEBOMTreeTable").Object.getValueAt(0,0).toString()
			sTable_ID = ""
			sTable_RevID = ""
			sTable_Name = ""
			If instr(sTableTopNode,";") > 0  Then
				' id rev name
				sTableTopNode = replace(sTableTopNode,"/","-")
				sTableTopNode = replace(sTableTopNode,";1-","-")
				sTableTopNode = trim(replace(sTableTopNode,"(View)",""))
				aTopNode = split(sTableTopNode, "-")
				sTable_ID = aTopNode(0)
				sTable_RevID = aTopNode(1)
				sTable_Name = aTopNode(2)
			ElseIf instr(sTableTopNode,"-") > 0  Then
				' id name
				aTopNode = split(sTableTopNode, "-")
				sTable_ID = aTopNode(0)
				sTable_Name = aTopNode(1)
			Else
				' Name
				sTable_Name = sTableTopNode
			End If

			If sTab_RevID  <> "" Then		' Tab Revision ID does not exist when we send a Struc Context, therefore below change in code.... Amit T - 04 - July - 2012
				If sTab_RevID = sTable_RevID Then
						If sTab_ID <> "" Then
								'match Name and ID
								If sTab_ID = sTable_ID AND sTab_Name = sTable_Name Then
									Fn_SISW_MPP_SetIndexMPPAppletFromTab = True
									Exit For
								End If
						ElseIf sTab_Name = sTable_Name Then
								Fn_SISW_MPP_SetIndexMPPAppletFromTab = True
								Exit for
						End If
				End If
			Else
				If sTab_ID <> "" Then
					'match Name and ID
					If sTab_ID = sTable_ID AND sTab_Name = sTable_Name Then
						Fn_SISW_MPP_SetIndexMPPAppletFromTab = True
						Exit For
					End If
				Else
					Fn_SISW_MPP_SetIndexMPPAppletFromTab = True
					Exit for
				End If
			End If ' End of sTab_RevID
		End If
	Next
End Function

'****************************************		Function to remove level of the BOMLine.  ***************************************

'Function Name		          :		   Fn_MPP_EndItemAssemblyStateOperation  

'Description			    :  		       Performs operation on End Item Assembly State dialog

'Parameters			   :	   	          1.  String - sAction ( Select / Verify )
'						          					2.  dicAssemblyState ("sBOMLine" /"Item" /  "Selected Occurrence"  / "sButton")
											
'Return Value		       :			True / False

'Pre-requisite			    :		 	 Structure Manager window should be displayed .

'Examples				    :			  	Set dicAssemblyState = CreateObject( "Scripting.Dictionary" )
'														dicAssemblyState("sBOMLine") = "003580/A;1-Product (View):003582/A;1-Assy2 (View)"
'														dicAssemblyState("Item") = "ON"
'														dicAssemblyState("Selected Occurrence") = "ON"
'														dicAssemblyState("sButton") = "OK"
'														Call Fn_MPP_EndItemAssemblyStateOperation("Select", dicAssemblyState)
'
'													Set dicAssemblyState = CreateObject( "Scripting.Dictionary" )
'														dicAssemblyState("Item") = "ON"
'														dicAssemblyState("Selected Occurrence") = "OFF"
'														dicAssemblyState("sButton") = "OK"
'														Call Fn_MPP_EndItemAssemblyStateOperation("Verify", dicAssemblyState)
'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					 	   Gaurav Singh			  17-Jan-2014		  1.0					Created
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							Poonam Chopade		 15-Feb-2018		  1.1					Added Case "Verify"				
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_EndItemAssemblyStateOperation(sAction, dicAssemblyState)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_EndItemAssemblyStateOperation"
	Dim objRemove,objEndItem
	Dim bReturn,bFlag
	Fn_MPP_EndItemAssemblyStateOperation = False
	Set objEndItem = Fn_SISW_MPP_GetObject("EndItemAssemblyState")
	If dicAssemblyState("sBOMLine") <> "" Then
		' selecting BOM Line from BOM Table
		bReturn = Fn_MPP_BOMTable_NodeOperation("Select",dicAssemblyState("sBOMLine") , "","","")
		If bReturn = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_EndItemAssemblyStateOperation ] BOM Line [ "+ dicAssemblyState("sBOMLine")  +" ] selected successfully") 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_EndItemAssemblyStateOperation ]  Failed to select BOMLine[ "+ dicAssemblyState("sBOMLine")  +" ]") 
			Fn_MPP_EndItemAssemblyStateOperation = False
			Exit function
		End If
	End If

	Select Case sAction

		Case "Select","SelectWithoutClose"
			If Fn_UI_ObjectExist("Fn_MPP_EndItemAssemblyStateOperation",objEndItem)=False Then
				bReturn = Fn_MenuOperation("Select", "Edit:End Item Assembly")
				If bReturn <> True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_EndItemAssemblyStateOperation ] Case [ Menu ] Failed to select Edit > End Item Assembly State For BOMLine[ "+ dicAssemblyState("sBOMLine")  +" ]") 
					Fn_MPP_EndItemAssemblyStateOperation = False
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_EndItemAssemblyStateOperation ] Case [ Menu ] Edit > End Item Assembly State selected successfully") 
				End If
			End If
			If Fn_UI_ObjectExist("Fn_MPP_EndItemAssemblyStateOperation",objEndItem)=False Then Exit Function
	
			If dicAssemblyState("Item") <> "" Then 		' Value of CheckBox to set  i.e. ON/OFF 
				Call Fn_CheckBox_Set("Fn_MPP_EndItemAssemblyStateOperation", objEndItem, "Item",dicAssemblyState("Item"))
			End If
	
			If dicAssemblyState("Selected Occurrence") <> "" Then 		' ' Value of CheckBox to set  i.e. ON/OFF
				Call Fn_CheckBox_Set("Fn_MPP_EndItemAssemblyStateOperation", objEndItem, "SelectedOccurrence",dicAssemblyState("Selected Occurrence"))
			End If
			Call Fn_ReadyStatusSync(1)
			
			If sAction <> "SelectWithoutClose"  Then
				If dicAssemblyState("sButton") <> "" Then 			' Button Name to be clicked on the  End Item Assembly State dialog
					Fn_MPP_EndItemAssemblyStateOperation = Fn_Button_Click("Fn_MPP_EndItemAssemblyStateOperation",objEndItem,dicAssemblyState("sButton") )
				End If
			Else
				Fn_MPP_EndItemAssemblyStateOperation = True				
			End If
			
		Case "Verify" '[ TC11.5_20180122.00_NewDevlopment_PoonamC_15Feb2018 : Added Case to verify values for fields ]
			If dicAssemblyState("Item") <> "" Then 		' Check Value of CheckBox to set  i.e. ON/OFF 
				If dicAssemblyState("Item") = "ON" Then
					bFlag = True
				ElseIf dicAssemblyState("Item") = "OFF" Then
					bFlag = False
				End If
				If cbool(Fn_UI_Object_GetROProperty("Fn_MPP_EndItemAssemblyStateOperation",objEndItem.JavaCheckBox("Item"),"value")) <> bFlag Then
					Set objEndItem = nothing
					Exit Function
				End if
			End If
	
			If dicAssemblyState("Selected Occurrence") <> "" Then 		' Check Value of CheckBox to set  i.e. ON/OFF
				If dicAssemblyState("Selected Occurrence") = "ON" Then
					bFlag = True
				ElseIf dicAssemblyState("Selected Occurrence") = "OFF" Then
					bFlag = False
				End If
				If cbool(Fn_UI_Object_GetROProperty("Fn_MPP_EndItemAssemblyStateOperation",objEndItem.JavaCheckBox("SelectedOccurrence"),"value")) <> bFlag Then
					Set objEndItem = nothing
					Exit Function
				End if
			End If
			
			If dicAssemblyState("For the current selection") <> "" Then 		' Check Value of radio to set  i.e. ON/OFF 
				If dicAssemblyState("For the current selection") = "ON" Then
					bFlag = True
				ElseIf dicAssemblyState("For the current selection") = "OFF" Then
					bFlag = False
				End If
				If cbool(Fn_UI_Object_GetROProperty("Fn_MPP_EndItemAssemblyStateOperation",objEndItem.JavaRadioButton("For the current selection"),"value")) <> bFlag Then
					Set objEndItem = nothing
					Exit Function
				End if
			End If
	
			If dicAssemblyState("In all structures where") <> "" Then 		' Check Value of radio to set  i.e. ON/OFF
				If dicAssemblyState("In all structures where") = "ON" Then
					bFlag = True
				ElseIf dicAssemblyState("In all structures where") = "OFF" Then
					bFlag = False
				End If
				If cbool(Fn_UI_Object_GetROProperty("Fn_MPP_EndItemAssemblyStateOperation",objEndItem.JavaRadioButton("In all structures where"),"value")) <> bFlag Then
					Set objEndItem = nothing
					Exit Function
				End if
			End If
			Fn_MPP_EndItemAssemblyStateOperation = True
			
			If dicAssemblyState("sButton") <> "" Then 			' Button Name to be clicked on the  End Item Assembly State dialog
				Call Fn_Button_Click("Fn_MPP_EndItemAssemblyStateOperation",objEndItem,dicAssemblyState("sButton") )
			End If	
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_EndItemAssemblyStateOperation ]  Failed to perform operation on End Item Assemebly State") 
	End Select

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_EndItemAssemblyStateOperation ] executed successfully.") 
	Set objEndItem = nothing
End Function

 '*********************************************************		Function to create basic Process		***********************************************************************
'Function Name		:	Fn_MPP_WorkareaCreate

'Description			:	Creats an Workarea with basic information

'Parameters			:	1.StrItemType: Type of the Process.
'						2.StrConfItem: True or False
'						3.StrItemID: ID of the item it should be unique.
'						4.StrItemRevID:Revision ID of the Process.
'						5.StrItemName:Name of Process.
'						6.StrItemDesc: Description of the Process.
'						7:StrItemUOM: Unit of measure of Process.

'Return Value		: 		Item Id  -  Revision Id

'Pre-requisite		:	Should be Prespective to Manufacturing Process Planner

'Examples			:	Call Fn_MPP_WorkareaCreate("MEPlant","","","","Name","Desc","")

'History				:		
'							Developer Name		     Date				Rev. No.		Changes Done		Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'							  Veena Patil 		01-March-2017		1.0					Created			Poonam
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_WorkareaCreate(StrProcessType,StrConfProcess,StrProcessID,StrProcessRevID,StrProcessName,StrProcessDesc,StrProcessUOM)
	Dim sItemId, sRevId,sNewWorkareaMenu
	Dim objDialogNewProcess
	
	'Select menu [File:New:Workarea...]
	'Check the existence of "New Process " window.Activate
	'tc 11.4 20171201 - Sandip C maintenance -  Modified function for changed object types and different UI's observed during creation of 'MEWorkarea'
	Set objDialogNewProcess = JavaWindow("Manufacturing Process").JavaWindow("New Process")
	'Set objDialogNewProcess=Window("MPPWindow").JavaDialog("New Process")
	sNewWorkareaMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"FileNewWorkarea")
	If objDialogNewProcess.Exist(3) = False Then
        Call Fn_MenuOperation("Select",sNewWorkareaMenu)
	End If
	
		If StrProcessType = "MEWorkarea" Then
				'Select Process Type
			   Call Fn_JavaTree_Select("Fn_MPP_WorkareaCreate", objDialogNewProcess,"Enter Additional Process","Complete List"+":"+StrProcessType)
			   'Click on "Next" button
	             Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess,"Next")
				'Verify Id is Empty
				If StrItemID <> "" Then
					'Set  Item Id
	                 Call Fn_Edit_Box("Fn_MPP_WorkareaCreate",objDialogNewProcess,"NewWorkAreaID", StrItemID)
				End If
				'Click on Assign Button
				If  StrItemID = "" Then
					'click on assign button
					JavaWindow("Manufacturing Process").JavaWindow("New Process").JavaButton("Assign").SetTOProperty "index", 0
	                  Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess, "Assign")
				End If
	
	            wait(3)
				'Verify RevId is Empty
				If StrItemRevID <> "" Then
					'Set Revision ID
	                Call Fn_Edit_Box("Fn_MPP_WorkareaCreate",objDialogNewProcess,"NewWorkAreaRev", StrItemRevID)
				End If
	
				'Click on Assign Button
				If StrItemRevID = "" Then
					'click on assign button
					  JavaWindow("Manufacturing Process").JavaWindow("New Process").JavaButton("Assign").SetTOProperty "index", 1
	                  Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess, "Assign")
				End If
	
	            wait(3)
	
	            'Extract Creation data
				sItemId = Fn_Edit_Box_GetValue("Fn_MPP_WorkareaCreate", objDialogNewProcess,"NewWorkAreaID")
	            sRevId = Fn_Edit_Box_GetValue("Fn_MPP_WorkareaCreate", objDialogNewProcess,"NewWorkAreaRev")
	
				'Set Process Name
				If StrProcessName <> "" Then
	                 Call Fn_Edit_Box("Fn_MPP_WorkareaCreate",objDialogNewProcess,"Name", StrProcessName)
				End If
	
				'Set Process Desc
				If StrProcessDesc <> "" Then
	                 Call Fn_Edit_Box("Fn_MPP_WorkareaCreate",objDialogNewProcess,"Description", StrProcessDesc)
				End If
				
				'Set Process UOM
				If StrProcessUOM <> "" Then
	              Call Fn_Edit_Box("Fn_MPP_WorkareaCreate", objDialogNewItem,"Unit of Measure",StrProcessUOM)
				End If
				
	
		Else
				wait 2
				'Select Process Type
			   Call Fn_JavaTree_Select("Fn_MPP_WorkareaCreate", objDialogNewProcess,"Enter Additional Process","Complete List"+":"+StrProcessType)
			   'Click on "Next" button
	             Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess,"Next")
				'Verify Id is Empty
				If StrItemID <> "" Then
					'Set  Item Id
	                 Call Fn_Edit_Box("Fn_MPP_WorkareaCreate",objDialogNewProcess,"ID", StrItemID)
				End If
	
				'Click on Assign Button
				If  StrItemID = "" or StrItemRevID = "" Then
					'click on assign button
					JavaWindow("Manufacturing Process").JavaWindow("New Process").JavaButton("Assign").SetTOProperty "index", 0
	                  Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess, "Assign")
				End If
	
	            wait(3)
	
	            If Fn_SISW_UI_Object_Operations("Fn_MPP_WorkareaCreate", "Exist", JavaWindow("Manufacturing Process").JavaWindow("New Process").JavaEdit("ID"),"") = False Then
	            'Extract Creation data
					sItemId = Fn_Edit_Box_GetValue("Fn_MPP_WorkareaCreate", objDialogNewProcess,"NewWorkAreaID")
				Else
					sItemId = Fn_Edit_Box_GetValue("Fn_MPP_WorkareaCreate", objDialogNewProcess,"ID")
				End If
	
				'Set Process Name
				If StrProcessName <> "" Then
	                 Call Fn_Edit_Box("Fn_MPP_WorkareaCreate",objDialogNewProcess,"Name", StrProcessName)
				End If
	
				'Set Process Desc
				If StrProcessDesc <> "" Then
	                 Call Fn_Edit_Box("Fn_MPP_WorkareaCreate",objDialogNewProcess,"Description", StrProcessDesc)
				End If
				
				'Set Process UOM
				If StrProcessUOM <> "" Then
	              Call Fn_Edit_Box("Fn_MPP_WorkareaCreate", objDialogNewItem,"Unit of Measure",StrProcessUOM)
				End If
			
							'Verify RevId is Empty
				If StrItemRevID <> "" Then
					'Set Revision ID
	                Call Fn_Edit_Box("Fn_MPP_WorkareaCreate",objDialogNewProcess,"RevID", StrItemRevID)
				End If
							'Click on Assign Button
				If StrItemRevID = "" Then
					'click on assign button
					JavaWindow("Manufacturing Process").JavaWindow("New Process").JavaButton("Assign").SetTOProperty "index", 1
	                  Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess, "Assign")
				End If
				If Fn_SISW_UI_Object_Operations("Fn_MPP_WorkareaCreate", "Exist", JavaWindow("Manufacturing Process").JavaWindow("New Process").JavaEdit("RevID"),"") = False Then
		
					sRevId = Fn_Edit_Box_GetValue("Fn_MPP_WorkareaCreate", objDialogNewProcess,"NewWorkAreaRev")
				Else
					sRevId = Fn_Edit_Box_GetValue("Fn_MPP_WorkareaCreate", objDialogNewProcess,"RevID")
				End  If
				
					'Click on "Next" button
				If Fn_SISW_UI_Object_Operations("Fn_MPP_WorkareaCreate", "Exist", JavaWindow("Manufacturing Process").JavaWindow("New Process").JavaEdit("NewWorkAreaRev"),"") = False Then
	            			 Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess,"Next")
	            		End  If
				wait(2)

		End If
	
			objDialogNewProcess.JavaButton("Finish").WaitProperty "enabled", 1, 20000
			'Click on Finish Button 
			Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess,"Finish")
			wait(1)
			Fn_MPP_WorkareaCreate = sItemId & "-" & sRevId
			Call Fn_ReadyStatusSync(1)

			 If objDialogNewProcess.Exist(3)Then
				'Click on Close button
				Call Fn_Button_Click("Fn_MPP_WorkareaCreate", objDialogNewProcess, "Close") 
			End If
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Workarea of ID [" + CStr(sItemId) + "]")

		Set objDialogNewProcess=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name			:	Fn_MPP_CreatePublishLink_Operation
'@@
'@@    Description				:	Function used to perform operation on Create Publish Link in MPP
'@@
'@@    Parameters			   	:	1. sAction			: Action [ActionName]
'@@								:	2. dicDetails		: Information for Create Publish Link
'@@								:   3. sButton 			: Button Name										 	         
'@@
'@@    Return Value		   	   	: 	True Or False
'@@
'@@    							:	Set dicDetails = CreateObject("Scripting.Dictionary")
'@@										dicDetails("SelectSourceTab") =  "000817-Design1"
'@@										dicDetails("SelectSourceObject") = "000817/A;1-Design1 (View):000818/A;1-Design2"
'@@										dicDetails("SelectTargetTab") = "000819-Part1"
'@@										dicDetails("SelectTargetObject") = "000819/A;1-Part1 (View):000820/A;1-Part2"
'@@    Examples					:	Call Fn_MPP_CreatePublishLink_Operation("Add",dicDetails,"OK")
'@@
'@@    History					:	
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Developer Name			Date			 Rev. No.	   				Changes Done								 Reviewer
'@@ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  Poonam Chopade	 	01-Feb-2018	 			1.0			Created - Added for PSM new TC's development			TC11.4(2017120100)_NewDevelopment_PoonamC_01Feb2018
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_MPP_CreatePublishLink_Operation(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_CreatePublishLink_Operation"
	Dim ObjCreatePubLink,sMenu
	Dim bReturn
	Fn_MPP_CreatePublishLink_Operation = False
	
	'Select BomLine in MPP
	If dicDetails("SelectSourceTab") <> "" Then
		bReturn = Fn_MPP_TabOperations("Select",dicDetails("SelectSourceTab"))
		Call Fn_ReadyStatusSync(1)
		If bReturn = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_CreatePublishLink_Operation ] Tab [ "+ dicDetails("SelectSourceTab") +" ] selected successfully") 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_CreatePublishLink_Operation ]  Failed to Tab [ "+ dicDetails("SelectSourceTab")  +" ]") 
			Fn_MPP_CreatePublishLink_Operation = False
			Exit function
		End If
	End If
	
	'Select BomLine in MPP
	If dicDetails("SelectSourceObject") <> "" Then
		If sAction="AddExt" Then
			bReturn = Fn_MPP_BOMTable_NodeOperation("SelectExt",dicDetails("SelectSourceObject") , "","","")
		Else
			bReturn = Fn_MPP_BOMTable_NodeOperation("Select",dicDetails("SelectSourceObject") , "","","")
		End If
		Call Fn_ReadyStatusSync(1)
		If bReturn = True Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_CreatePublishLink_Operation ] BOM Line [ "+ dicDetails("SelectSourceObject")  +" ] selected successfully") 
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_CreatePublishLink_Operation ]  Failed to select BOMLine [ "+ dicDetails("SelectSourceObject")  +" ]") 
			Fn_MPP_CreatePublishLink_Operation = False
			Exit function
		End If
	End If
	
	Set ObjCreatePubLink = Fn_SISW_MPP_GetObject("CreatePublishLink")
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_MPP_CreatePublishLink_Operation",ObjCreatePubLink)  = False  Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"CreatePublishLink")
		Call Fn_MenuOperation("WinMenuSelect","Tools")
		Call Fn_ReadyStatusSync(2)
		wait 1
		Call Fn_KeyBoardOperation("SendKeys", "{UP}")
		Wait 1
		Call Fn_MenuOperation("WinMenuSelect",sMenu)
		Call Fn_ReadyStatusSync(2)
		wait 1
		If Fn_UI_ObjectExist("Fn_MPP_CreatePublishLink_Operation",ObjCreatePubLink)  = False  Then
			Set ObjCreatePubLink = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_MPP_CreatePublishLink_Operation ] Create Publish Link dialog dose not exists ].")
			Exit Function
		End If
	End If
	
	Select Case sAction

		Case "Add","AddExt"
			If dicDetails("SelectTargetTab") <> "" Then
				 	'Select Target Tab
				 	bReturn = Fn_MPP_TabOperations("Select",dicDetails("SelectTargetTab"))
				 	Call Fn_ReadyStatusSync(1)
					If bReturn = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_CreatePublishLink_Operation ] Target Tab [ "+ dicDetails("SelectTargetTab")  +" ] selected successfully") 
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_CreatePublishLink_Operation ]  Failed to select Target Tab [ "+ dicDetails("SelectTargetTab")  +" ]") 
						Set ObjCreatePubLink = Nothing
						Fn_MPP_CreatePublishLink_Operation = False
						Exit Function
					End If
			 End If	
		
			 If dicDetails("SelectTargetObject") <> "" Then
			 	'Select Target Node
				If sAction="AddExt" Then
					bReturn = Fn_MPP_BOMTable_NodeOperation("SelectExt",dicDetails("SelectTargetObject"), "","","")
				Else
					bReturn = Fn_MPP_BOMTable_NodeOperation("Select",dicDetails("SelectTargetObject"), "","","")
				End If
				Call Fn_ReadyStatusSync(1)
				If bReturn = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_CreatePublishLink_Operation ] Target Object [ "+ dicDetails("SelectTargetObject")+" ] selected successfully") 
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_CreatePublishLink_Operation ]  Failed to select Target Object [ "+ dicDetails("SelectTargetObject") +" ]") 
					Set ObjCreatePubLink = Nothing
					Fn_MPP_CreatePublishLink_Operation = False
					Exit Function
				End If
				
				'Click on Add Target button
				Call Fn_Button_Click("Fn_MPP_CreatePublishLink_Operation",ObjCreatePubLink,"AddTarget")
				Wait 1
			 End If	
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_MPP_CreatePublishLink_Operation ] Invalid case.") 
	End Select
	
	If sButton <> "" Then
		Call Fn_Button_Click("Fn_MPP_CreatePublishLink_Operation",ObjCreatePubLink,sButton)
		Call Fn_ReadyStatusSync(1)
	End If
	
	Fn_MPP_CreatePublishLink_Operation = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_MPP_CreatePublishLink_Operation ] executed successfully.") 
	Set ObjCreatePubLink = nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_MPP_AdvancedAccountabilityCheck_Operation
'@@
'@@    Description		:	Function Used to perform operations on "Advanced Accountability Check" window
'@@
'@@    Parameters		:	1. sAction		: Action to be performed
'@@						:	2. dicDetails	: Dictionary object
'@@						:	3. sButton		: OK / Cancel button
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:	Set dicDetails = CreateObject("Scripting.Dictionary")
'@@								dicDetails("SelectMainTab1") = "Scope"
'@@								dicDetails("SelectSourceTab") = ""
'@@								dicDetails("SelectSourceObject") = ""
'@@								dicDetails("SelectTargetTab") = ""
'@@								dicDetails("SelectTargetObject") = ""
'@@								dicDetails("SelectMainTab2") = "Inclusion Rules"
'@@								dicDetails("Source filtering rule") = ""
'@@								dicDetails("Target filtering rule") = ""
'@@								dicDetails("SelectMainTab3") = "Reporting"
'@@								dicDetails("SetColorCheckBox") = "SetAllON"
'@@								dicDetails("SelectMainTab4") = "Equivalence"
'@@								dicDetails("PublishLink Connection") = "ON"
'@@								dicDetails("SelectMainTab5") = "Partial Match"
'@@								dicDetails("SelectInternalTab") = "BOM Properties"
'@@								dicDetails("AddProperties") = "Absolute Transformation Matrix"
'@@							bReturn = Fn_MPP_AdvancedAccountabilityCheck_Operation("Compare",dicDetails,"OK")
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History			:	Developer Name			Date	   		Rev. No.		Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	  						Poonam Chopade	 	01-Feb-2018			1.0			Created - Added for PSM new TC's development		[TC114-2017120100-01Feb2018-PoonamC-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_MPP_AdvancedAccountabilityCheck_Operation(sAction,dicDetails,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_AdvancedAccountabilityCheck_Operation"
	Dim objAACWindow, objDevReplay
	Dim dicCount, dicItems, dicKeys, aProperty, aChkBox
	Dim iCounter, iCount, bFlag, iBtnIndex,sMenu
	Dim sSubAction, sProperty, sTableName,iClickCounter
	
	Const VK_CONTROL = 29
	
	Fn_MPP_AdvancedAccountabilityCheck_Operation = False
	Set objAACWindow = Fn_SISW_MPP_GetObject("AdvancedAccountabilityCheck")
	
	'Check Dialog existence
	If Fn_UI_ObjectExist("Fn_MPP_AdvancedAccountabilityCheck_Operation",objAACWindow)  = False  Then
		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"ToolsAccountabilityCheckAdvancedAccountabilityCheck")
		Call Fn_MenuOperation("WinMenuSelect","Tools")
		Call Fn_ReadyStatusSync(2)
		Call Fn_KeyBoardOperation("SendKeys", "{DOWN}")
		Wait 1
		Call Fn_MenuOperation("WinMenuSelect",sMenu)
		Call Fn_ReadyStatusSync(2)
		If Fn_UI_ObjectExist("Fn_MPP_AdvancedAccountabilityCheck_Operation",objAACWindow)  = False  Then
			Set objAACWindow = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_MPP_AdvancedAccountabilityCheck_Operation ] Create Publish Link dialog dose not exists ].")
			Exit Function
		End If
	End If
	
	Select Case sAction
		Case "Compare","CompareExt"
				dicCount = dicDetails.Count
				dicItems = dicDetails.Items
				dicKeys = dicDetails.Keys
				
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"SelectMainTab")>0 Then
						sSubAction = "SelectMainTab"
					ElseIf Instr(dicKeys(iCounter),"SelectInternalTab")>0 Then
						sSubAction = "SelectInternalTab"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					sProperty = dicItems(iCounter)
					bFlag = False
					
					Select Case sSubAction
						'Select MainTab or InternalTab in Partial Match tab
						Case "SelectMainTab","SelectInternalTab"
							If sProperty<>"" Then
								If sSubAction = "SelectMainTab" Then
									objAACWindow.JavaTab("MainTab").Select sProperty
									bFlag = True	
								ElseIf sSubAction = "SelectInternalTab" Then
									objAACWindow.JavaTab("InternalTab").Select sProperty
									bFlag = True
								End If
								If bFlag = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+"] Tab in [AdvancedAccountabilityCheck] window.")
									Set objAACWindow = Nothing
									Exit Function
								End If
							End If
						Case "SelectSourceTab"
								objAACWindow.Minimize 
								wait 1
								bFlag = Fn_MPP_TabOperations("Select",sProperty)
								objAACWindow.Restore
								If bFlag = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+"] Tab in MPP.")
									Set objAACWindow = Nothing
									Exit Function
								End If  	
						Case "SelectSourceObject"
							objAACWindow.Minimize 
							wait 1
							If sAction="CompareExt" Then
								bFlag = Fn_MPP_BOMTable_NodeOperation("SelectExt",sProperty, "","","")
							Else
								bFlag = Fn_MPP_BOMTable_NodeOperation("Select",sProperty, "","","")							
							End If
							objAACWindow.Restore
							If bFlag = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+"] Source Object in MPP.")
									Set objAACWindow = Nothing
									Exit Function
							End If
							objAACWindow.Activate()
							Wait 1
							bFlag = Fn_Button_Click("Fn_MPP_AdvancedAccountabilityCheck_Operation", objAACWindow, "AddSource")
							
						Case "SelectTargetTab"
								objAACWindow.Minimize 
								wait 1
								bFlag = Fn_MPP_TabOperations("Select",sProperty)
								objAACWindow.Restore
								If bFlag = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+"] Tab in MPP.")
									Set objAACWindow = Nothing
									Exit Function
								End If
						Case "SelectTargetObject"
							objAACWindow.Minimize 
							wait 1
							If sAction="CompareExt" Then
								bFlag = Fn_MPP_BOMTable_NodeOperation("SelectExt",sProperty, "","","")
							Else
								bFlag = Fn_MPP_BOMTable_NodeOperation("Select",sProperty, "","","")							
							End If
							objAACWindow.Restore
							If bFlag = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+"] Target Object in MPP.")
									Set objAACWindow = Nothing
									Exit Function
							End If
							objAACWindow.Activate()
							Wait 1
							bFlag = Fn_Button_Click("Fn_MPP_AdvancedAccountabilityCheck_Operation", objAACWindow, "AddTarget")	
							
						Case "Source filtering rule","Target filtering rule"
							objAACWindow.JavaRadioButton("InclusionRules").SetTOProperty "attached text","Search lines per filtering rule"						
							Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_MPP_AdvancedAccountabilityCheck_Operation", "Set",objAACWindow,"InclusionRules", "ON")
							'Select filtering rule
							objAACWindow.JavaList("filteringrule").SetTOProperty "attached text",sSubAction&":"
							bFlag = Fn_SISW_UI_JavaList_Operations("Fn_MPP_AdvancedAccountabilityCheck_Operation","Select",objAACWindow,"filteringrule",sProperty, "", "")
							If bFlag = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+"] in [AdvancedAccountabilityCheck] window.")
									Set objAACWindow = Nothing
									Exit Function
							End If
						Case "PublishLink Connection"
							objAACWindow.JavaCheckBox("Equivalence").SetTOProperty "attached text",sSubAction
							bFlag = Fn_SISW_UI_JavaCheckBox_Operations("Fn_MPP_AdvancedAccountabilityCheck_Operation","Set",objAACWindow,"Equivalence",sProperty)
							If bFlag = False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+sProperty+"] in [AdvancedAccountabilityCheck] window.")
									Set objAACWindow = Nothing
									Exit Function
							End If	
						'Case to Set Color check boxes values as ON or OFF in Reporting tab
						Case "SetColorCheckBox"
							If sProperty<>"" Then
								Call Fn_UI_JavaTab_Select("Fn_MPP_AdvancedAccountabilityCheck_Operation",objAACWindow,"MainTab","Reporting")
								If sProperty = "SetAllON" Then
									sProperty = "Color the compared objects:ON~Full match:ON~Partial match:ON~Missing target:ON~Missing source:ON~Multiple match:ON~Multiple partial match:ON"
									aProperty = Split(sProperty,"~")
								Else
									aProperty = Split(sProperty,"~")
								End If
								'Set color check boxes values as ON or OFF
								For iCount = 0 To UBound(aProperty)
									aChkBox = Split(aProperty(iCount),":")
									objAACWindow.JavaCheckBox("ColorCheckBox").SetTOProperty "attached text",aChkBox(0)
									bFlag = Fn_CheckBox_Set("Fn_MPP_AdvancedAccountabilityCheck_Operation",objAACWindow,"ColorCheckBox",aChkBox(1))
									If bFlag=False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Set Checkbox ["+aChkBox(0)+"] value as ["+aChkBox(1)+"] in [AdvancedAccountabilityCheck] window.")
										Set objAACWindow = Nothing
										Exit Function
									End If
									Wait 0,500
								Next
							End If
						'Case to Add properties in 
						Case "AddProperties","RemoveProperties"
							If sProperty<>"" Then
								'Set Checkbox "Consider values of properties when searching for a partial match" value as ON
								objAACWindow.JavaCheckBox("ColorCheckBox").SetTOProperty "attached text","Consider values of properties when searching for a partial match"
								bFlag = Fn_CheckBox_Set("Fn_MPP_AdvancedAccountabilityCheck_Operation",objAACWindow,"ColorCheckBox","ON")
								If bFlag=False Then
									Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Set Checkbox [Consider values of properties when searching for a partial match] value as [ON] in [AdvancedAccountabilityCheck] window.")
									Set objAACWindow = Nothing
									Exit Function
								End If
								
								If sSubAction="AddProperties" Then
									sTableName = "Available Properties:"	'or it could be "Available Attribute Group properties:"
									iBtnIndex = 0
								ElseIf sSubAction="RemoveProperties" Then
									sTableName = "Selected Properties:"	'or it could be "Selected Attribute Group properties:"
									iBtnIndex = 1
								End If
								
								Set objDevReplay = CreateObject("Mercury.DeviceReplay")
								'Add Properties from Available Properties Table to Selected Properties Table
								objAACWindow.JavaTable("PropertiesTable").SetTOProperty "attached text",sTableName
								aProperty = Split(sProperty,"~")
								iClickCounter = 0
								For iCount = 0 To UBound(aProperty)
									bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_MPP_AdvancedAccountabilityCheck_Operation","Exist",objAACWindow,"PropertiesTable","",0,aProperty(iCount),"","","","")
									If bFlag=True Then
										bFlag = Fn_SISW_UI_JavaTable_Operations("Fn_MPP_AdvancedAccountabilityCheck_Operation","ClickCell",objAACWindow,"PropertiesTable","",0,aProperty(iCount),"","","","")
										If bFlag=False Then
											Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Select ["+aProperty(iCount)+"] in ["+sTableName+"] table in [AdvancedAccountabilityCheck] window.")
											Set objAACWindow = Nothing
											Set objDevReplay = Nothing
											Exit Function
										End If
										iClickCounter = iClickCounter+1
									End If
									If iClickCounter=1 Then
										objDevReplay.KeyDown VK_CONTROL
									End If
								Next
								If iClickCounter>0 Then
									objDevReplay.KeyUp VK_CONTROL
									Set objDevReplay = Nothing
									'Click on Add or Remove button as per SubCase
									objAACWindow.JavaButton("AddRemoveProperties").SetTOProperty "Index",iBtnIndex
									bFlag = Fn_Button_Click("Fn_MPP_AdvancedAccountabilityCheck_Operation",objAACWindow,"AddRemoveProperties")
									If bFlag=False Then
										Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Click on [AddRemoveProperties] button in [AdvancedAccountabilityCheck] window.")
										Set objAACWindow = Nothing
										Exit Function
									End If
								End If
								bFlag = True
							End If
					End Select
					
					If bFlag = False Then
						Fn_MPP_AdvancedAccountabilityCheck_Operation = False
						Set objAACWindow = Nothing
						Call Fn_WriteLogFile("","FAIL : Function [Fn_MPP_AdvancedAccountabilityCheck_Operation] Failed to Perform Case ["&sAction&"] SubCase ["+sSubAction+"].")
						Exit Function
					End If
				Next
				If sButton<>"" Then
					bFlag = Fn_Button_Click("Fn_MPP_AdvancedAccountabilityCheck_Operation",objAACWindow,sButton)
					If bFlag=False Then
						Call Fn_WriteLogFile("","FAIL: [Fn_MPP_AdvancedAccountabilityCheck_Operation]-Action-["+sAction+"]-SubAction-["+sSubAction+"] : Failed to Click on ["+sButton+"] button in [AdvancedAccountabilityCheck] window.")
						Set objAACWindow = Nothing
						Exit Function
					End If
				End If
				Fn_MPP_AdvancedAccountabilityCheck_Operation = True
		Case Else
				Set objAACWindow = Nothing
				Exit Function
		End Select
		Set objAACWindow = Nothing
End Function

'*********************************************************		Function to create basic Process		***********************************************************************
'Function Name			:				Fn_MPP_BasicProcessAreaCreate

'Description			:		 		Creats an Basic Process Area with basic information

'Parameters			    :	 			1.StrProcessAreaType: Type of the Process Area.
'										2.StrConfItem: True or False
'										2.StrItemID: ID of the Area it should be unique.
'										3.StrItemRevID:Revision ID of the Process.
'										4.StrProcessAreaName:Name of Area Process.
'										5.StrProcessAreaDesc: Description of the Process Area.
'										6:StrItemUOM: Unit of measure of Process Area.

'Return Value		   : 				Item Id  -  Revision Id

'Pre-requisite		   :		 		should be Prespective to Manufacturing Process Planner

'Examples			   :				 Call Fn_MPP_BasicProcessAreaCreate("Process Area","","","","Name","Desc","")

'History			   :		
'										Developer Name								Date						Rev. No.						Changes Done				Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Komal Khedkar							07-May-2021			       			  1.0						   Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MPP_BasicProcessAreaCreate(StrProcessAreaType,StrConfProcess,StrProcessID,StrProcessRevID,StrProcessAreaName,StrProcessAreaDesc,StrProcessUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_BasicProcessAreaCreate"
	Dim sItemId, sRevId
	Dim objDialogNewProcess,objDialogNewProcessNew

	Set objDialogNewProcessArea = JavaWindow("Manufacturing Process").JavaWindow("New Process Area")

	If objDialogNewProcessArea.Exist(3) = False Then
        Call Fn_MenuOperation("Select","File:New:Process Area...")
	End If
	Wait 3
	If objDialogNewProcessArea.Exist(3) Then
	 	 Call Fn_JavaTree_Select("Fn_MPP_BasicProcessAreaCreate", objDialogNewProcessArea, "Business Object Type","Complete List:"&StrProcessAreaType)
   		'Click on "Next" button
      Call Fn_Button_Click("Fn_MPP_BasicProcessAreaCreate", objDialogNewProcessArea,"Next")
	End If
	wait 5
	'Verify Id is Empty
	If StrProcessID <> "" Then
		'Set  Process Id
		  Call Fn_Edit_Box("Fn_MPP_BasicProcessAreaCreate",objDialogNewProcessArea,"NewWorkAreaID", StrProcessID)
	End If
	wait 5
	'Verify RevId is Empty
	If StrItemRevID <> "" Then
	'Set Revision ID
        Call Fn_Edit_Box("Fn_MPP_BasicProcessAreaCreate",objDialogNewProcessArea,"NewWorkAreaRev", StrItemRevID)
	End If
	wait 5
	'Click on Assign Button
	If  StrProcessID = "" or StrItemRevID = "" Then
		'click on assign button
          objDialogNewProcessArea.JavaButton("Assign").SetTOProperty "Index","0"
          Call Fn_Button_Click("Fn_MPP_BasicProcessAreaCreate", objDialogNewProcessArea, "Assign")
          objDialogNewProcessArea.JavaButton("Assign").SetTOProperty "Index","1"
          Call Fn_Button_Click("Fn_MPP_BasicProcessAreaCreate", objDialogNewProcessArea, "Assign")
	End If
       wait(3)

 	sItemId = Fn_Edit_Box_GetValue("Fn_MPP_BasicProcessAreaCreate", objDialogNewProcessArea,"NewWorkAreaID")
       sRevId = Fn_Edit_Box_GetValue("Fn_MPP_BasicProcessAreaCreate", objDialogNewProcessArea,"NewWorkAreaRev")
	
	'Set Process Area Name
	If StrProcessAreaName <> "" Then
         Call Fn_Edit_Box("Fn_MPP_BasicProcessAreaCreate",objDialogNewProcessArea,"Name", StrProcessAreaName)
	End If

	'Set Process Desc
	If StrProcessAreaDesc <> "" Then
         Call Fn_Edit_Box("Fn_MPP_BasicProcessAreaCreate",objDialogNewProcessArea,"Description", StrProcessDesc)
	End If
	
	'Set Process UOM
	If StrProcessUOM <> "" Then 
		If objDialogNewProcessArea.JavaEdit("Unit of Measure:").Exist(5) =False Then
			objDialogNewProcessArea.JavaEdit("Unit of Measure:").SetTOProperty"toolkit class","org.eclipse.swt.custom.StyledText"
		End If
      Call Fn_Edit_Box("Fn_ItemBasicCreate", objDialogNewItem,"Unit of Measure:",StrProcessUOM)
	End If

	wait(2)
	objDialogNewProcessArea.JavaButton("Finish").WaitProperty "enabled", 1, 20000
	'Click on Finish Button 
	Call Fn_Button_Click("Fn_MPP_BasicProcessAreaCreate", objDialogNewProcessArea,"Finish")
	wait(1)
	Fn_MPP_BasicProcessAreaCreate = sItemId & "-" & sRevId
	Call Fn_ReadyStatusSync(1)

	 If objDialogNewProcessArea.Exist(3)Then
		'Click on Close button
		Call Fn_Button_Click("Fn_ItemBasicCreate", objDialogNewProcessArea, "Close") 
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Process of ID [" + CStr(sItemId) + "]")

Set objDialogNewProcessArea=Nothing
End Function

'*********************************************	Function To perform operation on Variant Rule Dialog  ***************************************************************
'	Function Name			:		Fn_MPP_VariantRuleOperations
'
'	Description				:		This function is used to perform operation on Variant Rule Dialog
'
'	Parameters				:		1.	sAction - Action need to perform
'										"Modify"
'										
'									2.  dicVariantRule - Dictionary Object
'
'	Return Value			:		True / False

'	Pre-requisite			:	    Manufacturing process planner window should be displayed .
'									BOM Line should be selected

'	Examples				:		dicVariantRule("Option") = DataTable("LegVarName",dtGlobalSheet)
'                                   dicVariantRule("Value") ="1200cc"
'                                   bReturn=   Fn_PSE_VariantRuleOperations("Modify", dicVariantRule)

'History			   :		
'										Developer Name								Date						Rev. No.						Changes Done				Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								    	Neha Patil					         	26-July-2021			       1.0		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_MPP_VariantRuleOperations(sAction, dicVariantRule)
	GBL_FAILED_FUNCTION_NAME="Fn_MPP_VariantRuleOperations"
	Dim objVarConfig, objVarRule
	Dim bReturn, iCount, iRows, sData,bFlag,aOption,aValue,iCounter,iCnt
	Dim aRows, sProperties, sPropertyValues
	Dim sPerspective,objDefaultWindow,sMenuFile, sMenu
	Dim sItem, sValue, arrValue,ColumnNames,sTitle

	Set objDefaultWindow = Fn_SISW_GetObject("DefaultWindow")
	sTitle =Fn_UI_Object_GetROProperty("Fn_MPP_VariantRuleOperations",objDefaultWindow, "title") 
	
	If (instr (sTitle, "Manufacturing Process Planner") <= 0) Then
		bReturn = Fn_MPP_VariantRuleOperations(sAction, dicVariantRule)
		Fn_MPP_VariantRuleOperations =bReturn 
		Exit Function
	End If

	sProperties = "Class Name~path$RegularExpression"
	sPropertyValues =  "JavaTable~Table;Shell;Shell;.*"
	Fn_MPP_VariantRuleOperations = False
	' Configure variant window.
	
	Set objVarConfig = Fn_SISW_MPP_GetObject("Configure")
	Set objVarRule = Fn_SISW_MPP_GetObject("Variant Rule")
	

	If Fn_SISW_UI_Object_Operations("Fn_MPP_VariantRuleOperations","Exist", objVarRule, SISW_MICRO_TIMEOUT) = False Then
		If Fn_SISW_UI_Object_Operations("Fn_MPP_VariantRuleOperations","Exist", objVarConfig,SISW_MICRO_TIMEOUT) = False Then
			'Operate Tools>>Variants>>Configure Variants... menu to invoke required dialog
				sMenuFile = Fn_LogUtil_GetXMLPath("PSE_Menu")
				sMenu = Fn_GetXMLNodeValue(sMenuFile, "ToolsVariantsConfigureVariants")
				Call  Fn_MenuOperation("Select",sMenu) 
				Call Fn_ReadyStatusSync(1)	
	    End If

		If Fn_SISW_UI_Object_Operations("Fn_MPP_VariantRuleOperations","Exist", objVarRule, SISW_MICRO_TIMEOUT) = False Then
			' clicking on variant rule button toOpen Variant rule window
			   Call Fn_Button_Click("Fn_MPP_VariantRuleOperations",objVarConfig,"LegacySwitch")
	
			If Fn_SISW_UI_Object_Operations("Fn_MPP_VariantRuleOperations","Exist", objVarRule,SISW_DEFAULT_TIMEOUT) = False Then
			   If Fn_SISW_UI_Object_Operations("Fn_MPP_VariantRuleOperations","Exist", objVarRule2,SISW_DEFAULT_TIMEOUT) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_MPP_VariantRuleOperations ] Failed to open Variant Rule window.")
				Exit function
			   End If
			End If
	    End If
	 End If
	
	wait SISW_MIN_TIMEOUT
	If Trim(dicVariantRule("AllowMultipleValues")) <> ""  Then
		Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_MPP_VariantRuleOperations", "Set", objVarRule , "Allow multiple values", "ON")
		wait SISW_MIN_TIMEOUT
    End If
		
	Select Case sAction
	
	Case "Modify"
	 		Dim objWshell
	 		set objWshell = createobject("Wscript.Shell")
			If dicVariantRule("Option") <>  "" then
				aOption = Split(dicVariantRule("Option"),"~",-1,1)
			Else 
				Exit Function 
			End If 

			If dicVariantRule("Value") <>  "" then
				aValue = Split(dicVariantRule("Value"),"~",-1,1)
			Else 
				Exit Function 
			End If
			
			iRows = Fn_UI_Object_GetROProperty("Fn_MPP_VariantRuleOperations",objVarRule.JavaTable("JTable"), "rows") 
			For iCounter = 0 To Ubound(aOption)
				For iCount = 0 to iRows -1 
					sData = Fn_SISW_UI_JavaTable_Operations("Fn_MPP_VariantRuleOperations", "GetCellData", objVarRule.JavaTable("JTable") , "", "", "", iCount, "1", "", "", "")
					If Trim(sData) = Trim(aOption(iCounter))Then
						Call Fn_SISW_UI_JavaTable_Operations("Fn_MPP_VariantRuleOperations", "DeselectRow", objVarRule.JavaTable("JTable") , "", "", "",iCount , "", "", "", "")
						wait SISW_MICRO_TIMEOUT
						Call Fn_SISW_UI_JavaTable_Operations("Fn_MPP_VariantRuleOperations", "ClickCell", objVarRule.JavaTable("JTable") , "", "", "",iCount , "3", "", "", "")
						wait SISW_MIN_TIMEOUT
						objWshell.SendKeys aValue(iCounter)
						wait SISW_MIN_TIMEOUT				
						Exit For
					End If
				Next
				If iCount = iRows Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_MPP_VariantRuleOperations :Option value not found in table.")
					Exit Function 
				End If
			Next
			
			Fn_MPP_VariantRuleOperations = True
			' clicking ok
			If sAction = "Modify" Then
				Call Fn_Button_Click("Fn_MPP_VariantRuleOperations",objVarRule,"OK")
			Else
				Call Fn_Button_Click("Fn_MPP_VariantRuleOperations",objVarRule,"Apply")
			End If
			
	Case "Clear"
			Call Fn_Button_Click("Fn_MPP_VariantRuleOperations",objVarRule,"Clear")
			Call Fn_Button_Click("Fn_MPP_VariantRuleOperations",objVarRule,"OK")
			Fn_MPP_VariantRuleOperations = True		
			
	Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_MPP_VariantRuleOperations ] : Invalid Case [ " + sAction + "].")
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_MPP_VariantRuleOperations : With  Case [ " + sAction + "].")
	set objVarConfig = Nothing
	set objVarRule = Nothing
	set objWshell = Nothing
	Set objDefaultWindow =Nothing
	End Function

	
