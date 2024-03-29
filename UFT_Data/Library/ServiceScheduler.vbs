Option Explicit
Dim bResult
'---------------------------------------------------------	Function List ------------------------------------------------------------------------------------------------------------------------------------------------
'000. Fn_SISW_SrvScheduler_GetObject()
'001. Fn_SISW_SrvScheduler_ServiceEditorTabOperations
'002. Fn_SISW_SrvScheduler_TableRowIndex
'003. Fn_SISW_SrvScheduler_BOMTable_ColIndex
'004. Fn_SISW_SrvScheduler_BOMTable_NodeOperation
'005. Fn_SISW_SrvScheduler_CreateWorkOrder
'006. Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan
'007. Fn_SISW_SrvScheduler_CreatePhysicalLocation
'008. Fn_SISW_SrvScheduler_SearchOperations
'009. Fn_SISW_SrvScheduler_PartRequestTypeCreate
'010. Fn_SISW_SrvScheduler_GetTreeItemPath
'011. Fn_SISW_SrvScheduler_ViewTreeOperations
'012. Fn_SISW_SrvScheduler_TreeTableRowIndex
'013. Fn_SISW_SrvScheduler_SchTable_NodeOperation
'014. Fn_SISW_SrvScheduler_JobCardOperations
'015. Fn_SISW_SrvScheduler_JobTaskOperations
'016. Fn_SISW_SrvScheduler_DeleteJobCard
'017. Fn_SISW_SrvScheduler_SummaryTabOperations() 
'018. Fn_SISW_SrvScheduler_ReturnPartsOperations
'019. Fn_SISW_SrvScheduler_CreateDiscrepancyType
'020. Fn_SISW_SrvScheduler_UpdateConfigurationOperations()
'021  Fn_SISW_SrvScheduler_SearchNoticeOperations()
'022. Fn_SISW_SrvScheduler_ViewDefineCharToleOperation
'023. Fn_SISW_SrvScheduler_IssueParts
'024. Fn_SISW_SrvScheduler_PartRequestOperations
'025. Fn_SISW_SrvScheduler_SearchCharacteristicOperation
'026. Fn_SISW_SrvScheduler_FindProxyTask
'027. Fn_SISW_SrvScheduler_WorkflowRuleConfiguration
'028. Fn_SISW_GenAutoServiceSchedule_Opearations
'029. Fn_SrvScheduler_MaintenanceActions_Ops
'030. Fn_SISW_SrvScheduler_NewNoticeOperations
'****************************************    Function to return required Object ***************************************
'
''Function Name		 	:	Fn_SISW_SrvScheduler_GetObject
'
''Description		    :  	Function to get objects of Service Manager

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_SrvMgr_GetObject("GenerateAsMaintainedStructure")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 22-Oct-2012		 1.0						Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_SrvScheduler_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ServiceScheduler.xml"
	Set Fn_SISW_SrvScheduler_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function 

'*********************************************************		Function for Tabs in Service Schedule ***********************************************************************

'Function Name		:			Fn_SISW_SrvScheduler_ServiceEditorTabOperations

'Description		:			This function is used to Activate, Verify Activate, PopupMenuSelect(Close Panel, split panel) for Service Schduler Tabs.

'Parameters			:			   1.	strAction:
'												2.	strPanelName:
'												3.	strMenuName:"Close Panel" or "Split Panel"
											
'Return Value		:			True/False

'Pre-requisite		:			Requirement Manager window should be displayed .

'Examples			:			
		'Call Fn_SISW_SrvScheduler_ServiceEditorTabOperations("Activate","(000020-asd)","")
		'Call Fn_SISW_SrvScheduler_ServiceEditorTabOperations("VerifyActivate","(000020-asd)","")
		'Call Fn_SISW_SrvScheduler_ServiceEditorTabOperations("PopupMenuSelect","(000020-asd)","Split Panel")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ashwini Kumar				09-Oct-2013			1.0											
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_ServiceEditorTabOperations(sAction, sTabName, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_ServiceEditorTabOperations"
   Dim objTabName, objApplet
   Set objApplet = Fn_SISW_SrvScheduler_GetObject("SrvSchdulerApplet")
   Fn_SISW_SrvScheduler_ServiceEditorTabOperations=False
   objApplet.JavaStaticText("ServiceEditorTabName").SetTOProperty "label",sTabName

   Select Case sAction
		Case "Activate"
			objApplet.JavaStaticText("ServiceEditorTabName").Click 5,5,"LEFT"
            Fn_SISW_SrvScheduler_ServiceEditorTabOperations=True

		Case "VerifyActivate"
			Set objTabName = objApplet.JavaStaticText("ServiceEditorTabName").Object			
			Fn_SISW_SrvScheduler_ServiceEditorTabOperations=objTabName.selected()	
			If Fn_SISW_SrvScheduler_ServiceEditorTabOperations="true" then
				Fn_SISW_SrvScheduler_ServiceEditorTabOperations=True
			Else
				Fn_SISW_SrvScheduler_ServiceEditorTabOperations=False
			End If
			Set objTabName = Nothing

		Case "PopupMenuSelect"
			objApplet.JavaStaticText("ServiceEditorTabName").Click 5,5,"RIGHT"
			wait 2
			If Fn_UI_JavaMenu_Select("Fn_SISW_SrvScheduler_ServiceEditorTabOperations",objApplet,sPopupMenu) then
				Fn_SISW_SrvScheduler_ServiceEditorTabOperations=True
			 Else
				Fn_SISW_SrvScheduler_ServiceEditorTabOperations=False
			 End If
		Case Else
			Fn_SISW_SrvScheduler_ServiceEditorTabOperations=False
	End Select
End Function

'*********************************************************		Function to Get BOM Table Node Index in Service Scheduler***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_TableRowIndex

'Description			 :		 		  This function is used to get the BOM Table Node Index.

'Parameters			   :	 			1. objTable - Table Object 
'													 2. sNodeName:Name of the Node to retrieve Index for.
											
'Return Value		   : 				 Node index

'Pre-requisite			:		 		Service Scheduler window should be displayed .

'Examples				:				 Fn_SISW_SrvScheduler_TableRowIndex(objTable, "000020/A;1-asd")

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			9-Oct-2013				1.0											
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_TableRowIndex(objTable, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_TableRowIndex"
	Dim nodeArr, aRowNode, iColIndex, aPath
	Dim iRowCounter, sNode, iInstance, iNodeCounter, bFound 
	Dim iRows, sNodePath, sPath, StrNodePath, objComponent

	sPath = ""
    If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_TableRowIndex","Exist", objTable,"") = False Then
        Fn_SISW_SrvScheduler_TableRowIndex = -1
		Exit function
	End If
    iColIndex = 0
	bFound = False
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
    						set objComponent = ObjTable.object.getComponentForRow(iRowCounter)
							StrNodePath = ""
							Do while NOT (objComponent is Nothing)
								If StrNodePath = "" Then
									StrNodePath = objComponent.getProperty("bl_indented_title")
								Else
									StrNodePath =objComponent.getProperty("bl_indented_title") & ", " & StrNodePath
								End If
								set objComponent = objComponent.parent()
								If  objComponent is Nothing Then
									Exit do
								End If
							Loop
						If instr(StrNodePath, "@BOM::") > 0 Then
							StrNodePath = trim(replace(StrNodePath,"""",""))
							aPath = split(StrNodePath,",")
							StrNodePath = ""
							For icnt = 0 to uBound(aPath)
								aPath(iCnt) = Left(aPath(iCnt), instr(aPath(iCnt),"@")-1)
								If StrNodePath = "" Then
									StrNodePath = trim(aPath(iCnt))
								Else
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
								If  objComponent is Nothing Then
										Exit do
								End If
							Loop
                        StrNodePath = trim(replace(StrNodePath,", ",":"))
						If instr(sPath, StrNodePath ) > 0 Then
							If UBound(nodeArr) = iNodeCounter Then
								bFound = True
							End If
							Exit do
							'exit loop
						End if
					End if
				End If
				iRowCounter = iRowCounter + 1
				' increment counter
			loop
		Next
	End If
	If bFound Then
		Fn_SISW_SrvScheduler_TableRowIndex = iRowCounter
	Else
		Fn_SISW_SrvScheduler_TableRowIndex = -1
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvScheduler_TableRowIndex ] executed successfully.")
End Function

'*********************************************************		Function to Get BOM Table Column Index into Service Schedule***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_BOMTable_ColIndex

'Description			 :		 		  This function is used to get the BOM Table Node Index.

'Parameters			   :	 			1.  StrColName:Name of the Col to retrieve Index for.
											
'Return Value		   : 				 Col index

'Pre-requisite			:		 		Structure Manager window should be displayed .

'Examples				:				Fn_SISW_SrvScheduler_BOMTable_ColIndex("Item Type")

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			09-Oct-2013			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_BOMTable_ColIndex(StrColName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_BOMTable_ColIndex"
	Dim IntCols , IntCounter, ObjTable, StrColIndex, StrName,  iColCount
	Fn_SISW_SrvScheduler_BOMTable_ColIndex = -1
	'Verify that PSE BOM Table is displayed

		Set ObjTable = Fn_SISW_SrvScheduler_GetObject("SrvSchedulerBOMTable")
		'Get the No. of cols present in the BOM Table
		iColCount = CInt(ObjTable.GetROProperty("cols"))
		'Get the Col No. of required Column
		For IntCounter = 0 to iColCount -1
			StrName = ObjTable.GetColumnName(IntCounter)		  
			If Trim(StrName) = Trim(StrColName) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: The Column Index for Column [" + StrColName + "] is [" &IntCounter&"] in BOMTable")
				Fn_SISW_SrvScheduler_BOMTable_ColIndex = IntCounter
				Exit For
			End If
		Next
		If Cint(IntCounter) = Cint(iColCount) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"WARNING: The Column [" + StrColName + "] dose not exist in BOM table." )
		End If
		'Release the Table object
		set ObjTable = Nothing
End Function

'******************************************************************Function to perform BOM Table Node operations************************************************************************************************************

'Function Name:				Fn_SISW_SrvScheduler_BOMTable_NodeOperation

'Description: 				 1. This function is used to perform operations on the BOM Table Node.


'Parameters:				  1. sAction: Action to be performed (Eg : Select/Exist etc.)
'						  2. sNodeName: Fully qualified path of the BOM Table Node (Node delimiter as ':') (Multi-Nodes delimiter as ',')
'						  3. sColName: Name of the BOM Table Column
'						  4. sValue: BOM Table cell value for Edit or Verify actions
'						  5. sPopupMenu: BOM Table Node context menu to be selected
'											  

'Return Value:				TRUE \ FALSE

'Pre-requisite:				Structure Manager window should be displayed with BOM Table loaded.

'Examples:				Call Fn_SISW_SrvScheduler_BOMTable_NodeOperation("Select", "000020/A;1-asd@2", "", "", "")
'									Call Fn_SISW_SrvScheduler_BOMTable_NodeOperation("Deselect", "000020/A;1-asd", "", "", "")
'									Call Fn_SISW_SrvScheduler_BOMTable_NodeOperation("Exist", "000020/A;1-asd", "", "", "")
'									Call Fn_SISW_SrvScheduler_BOMTable_NodeOperation("PopupSelect", "000020/A;1-asd", "", "", "Send To:My Teamcenter")
'History:
'										Developer Name			Date				Rev. No.			Changes Done												Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ashwini Kumar 		   09-Oct-2013			 1.0  									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_BOMTable_NodeOperation(sAction, sNodeName, sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_BOMTable_NodeOperation"
   Dim objTable
	Fn_SISW_SrvScheduler_BOMTable_NodeOperation = False
   Set objTable=Fn_SISW_SrvScheduler_GetObject("SrvSchedulerBOMTable")

	Select Case sAction
			Case "Select"
				If sNodeName <> "" Then
					iRowCounter = Fn_SISW_SrvScheduler_TableRowIndex(objTable,sNodeName) 
					If iRowCounter <> -1 Then
						objTable.Object.clearSelection  
						objTable.SelectRow iRowCounter 
						Fn_SISW_SrvScheduler_BOMTable_NodeOperation = True
					End If
				End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			Case "Deselect"
				If sNodeName <> "" Then
					iRowCounter = Fn_SISW_SrvScheduler_TableRowIndex(objTable,sNodeName) 
					If iRowCounter <> -1 Then
						objTable.DeselectRow iRowCounter 
						Fn_SISW_SrvScheduler_BOMTable_NodeOperation = True
					End If
				End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			Case "Exist", "Exists"
				If sNodeName <> "" Then
					iRowCounter = Fn_SISW_SrvScheduler_TableRowIndex(objTable,sNodeName) 
					If iRowCounter <> -1 Then
						Fn_SISW_SrvScheduler_BOMTable_NodeOperation = True
					End If
				End If

	'---------------------------------------This case is used to get the Cell Value For BOM Table Node cell.----------------------------------------------
		Case "GetCellData"
			If sNodeName <> "" Then
					iRowCounter =Fn_SISW_SrvScheduler_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
					iBOMLineColIndex = Fn_SISW_SrvScheduler_BOMTable_ColIndex(sColName)
					iCount = cInt(objTable.GetROProperty("cols"))
                    objTable.DoubleClickCell iRowCounter,iBOMLineColIndex , "LEFT", "NONE" 
					Fn_SISW_SrvScheduler_BOMTable_NodeOperation = Trim(cstr(objTable.GetCellData( iRowCounter,sColName)))
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SrvScheduler_BOMTable_NodeOperation] Cell verified of PSE BOM Table Node [" + sNodeName + "]")
				End If
			End If

      '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "PopupSelect"
			objTable.Object.clearSelection
			If sNodeName <> "" Then
				iRowCounter = Fn_SISW_SrvScheduler_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
					objTable.ClickCell iRowCounter ,0, "RIGHT","NONE"
        		End If
				wait 5
				bResult= Fn_UI_JavaMenu_Select("",JavaWindow("ServiceScheduler"),sPopupMenu)
				If bResult <> False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SrvScheduler_BOMTable_NodeOperation] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
					Fn_SISW_SrvScheduler_BOMTable_NodeOperation = True
				End If
			End If
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvScheduler_BOMTable_NodeOperation ] Invalid Action [ " & sAction & " ].")
			Set objTable = nothing
			exit function
    End Select

	If Fn_SISW_SrvScheduler_BOMTable_NodeOperation <> FALSE then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvScheduler_BOMTable_NodeOperation ] executed successfully with Action [ " & sAction & " ].")	
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to execute Function [ Fn_SISW_SrvScheduler_BOMTable_NodeOperation ] with Action [ " & sAction & " ].")
	End if
	Set objTable = nothing
End Function

'*********************************************************		Function to Create Work Order On Any node ***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_CreateWorkOrder

'Description			 :		 		  This function is used to Create Work Order On Any node.

'Parameters			   :	 			1.  sAction
'													 2. sButtonName: Name of Node of BOMTable
'													 3. dicInputs: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		Create Work Order Window should be visible

'Examples				:               Set dic1= CreateObject( "Scripting.Dictionary" )
'													dic1("Synopsis") = "xyzesfd"
'																				
'													set dic2 = CreateObject( "Scripting.Dictionary" )
'													dic2("Name") = "temp"
'													dic2("StartDate") = "10-Aug-2013~6:00:00 PM"
'													dic2("FinishDate") = "10-Aug-2014~6:00:00 PM"
'													dic2("Is Schedule Public") = "False"
'													dic2("Is Percent Linked") = "True"
'													dic2("Published") = "True"
'													
'													set dic3 = CreateObject( "Scripting.Dictionary" )
'													dic3("sAction") = "Add..."
'													dic3("ID") = "0*"
'													dic3("SearchResults_Select") = "000017-Phy1"
'													
'													Set dic1("Plan") = dic2
'													Set dic1("Work Performed At") = dic3
'													
'													Call Fn_SISW_SrvScheduler_CreateWorkOrder("Create" , "Finish" , dic1)

'													Case "Remove"
'													Set dic1 = CreateObject( "Scripting.Dictionary" )
'													dic1("Asset") = "002352/--A"
'													dic1("Work Performed At") = "734265-PhysicalLocation734265"
'													dic1("Company Locations")="CompanyName734265"
'													dic1("Company Contacts")="Shital734265 Desai734265"
'													bReturn = Fn_SISW_SrvScheduler_CreateWorkOrder("Remove" ,"" , dic1)

'													Case "Verify"
'													Set dic1 = CreateObject( "Scripting.Dictionary" )
'													dic1("Asset") = "002352/--A"
'													dic1("Work Performed At") = "734265-PhysicalLocation734265"
'													dic1("Company Locations")="CompanyName734265"
'													dic1("Company Contacts")="Shital734265 Desai734265"
'													bReturn = Fn_SISW_SrvScheduler_CreateWorkOrder("Verify" ,"" , dic1)


'History:	
'	Developer Name			Date				Rev. No.					Changes Done														Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			09-Oct-2013			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle			2-Dec-2013			1.0							Modified Case "Create" To handle Properties 
'																							"Asset", Company Contacts
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle			3-Dec-2013			1.0							Added Cases "Remove", "Verify"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_CreateWorkOrder(sAction , sButtonName , dicInputs)
	
GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_CreateWorkOrder"
Dim objDialog, dicItem,dicValue,dicPlanInputs,dicAction,dicWorkPerformedInputs,dicCompanyLocations
Dim iRowCount, iCounter, arrValue

	Set objDialog=Fn_SISW_SrvScheduler_GetObject("CreateWorkOrder")

	Fn_SISW_SrvScheduler_CreateWorkOrder=False
	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_CreateWorkOrder","Exist", objDialog,"") = False Then
		For iCount = 0 To 10
			Set objDialog = Fn_SISW_SrvScheduler_GetObject("GenerateAutomatedService")
			objDialog.JavaWindow("Shell").SetTOProperty "index",iCount
			Set objDialog = objDialog.JavaWindow("Shell").JavaWindow("CreateWorkOrder")
			If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_CreateWorkOrder","Exist", objDialog,"") Then
				bFlag = True
				Exit For
			End if
		Next
		If bFlag = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Dialog [Create Work Order] does not Exist.")
			Exit Function
		End If
		Call Fn_ReadyStatusSync( 3 )
	End If

	Select Case sAction
		Case "Create"
				dicItem = dicInputs.Keys
				dicValue = dicInputs.Items
				For iCount = 0 to dicInputs.Count - 1
					Select Case trim(dicItem(iCount))
						Case "ID","Revision"
							'JavaList
							objDialog.JavaList(trim(dicItem(iCount))).Click 1,1,"LEFT"
							objDialog.JavaList(trim(dicItem(iCount))).Type trim(dicValue(iCount))
							If Err.number < 0 Then
								Exit Function
							End If
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Synopsis"
							'Edit Box
							bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_CreateWorkOrder", "Set",  objDialog, "Synopsis", trim(dicValue(iCount)) )
							If bResult = False Then Exit Function
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						 Case "Plan"
							Set dicPlanInputs=dicInputs("Plan")
							Call Fn_UI_JavaStaticText_Click(" Fn_TcObjectDelete", objDialog, "CommonDownArrow", 1, 1, "LEFT")
							Wait 2
							bResult=Fn_UI_JavaMenu_Select("",objDialog,"Create...")
							If bResult = False Then Exit Function
							Wait 2
							bResult = Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan("Create" , dicPlanInputs)
							If bResult = False Then Exit Function
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Work Performed At"
							objDialog.JavaStaticText("CommonDownArrowLabel").SetTOProperty "label",trim(dicItem(iCount))&":"
							Set dicWorkPerformedInputs=dicInputs("Work Performed At")
							dicAction=dicWorkPerformedInputs("sAction")
							Call Fn_UI_JavaStaticText_Click(" Fn_SISW_SrvScheduler_CreateWorkOrder", objDialog, "CommonDownArrow", 1, 1, "LEFT")
							Wait 2
							bResult=Fn_UI_JavaMenu_Select("",objDialog,dicAction)
							If bResult = False Then Exit Function
							Wait 2
							If dicAction="Create..."	 Then
									bResult = Fn_SISW_SrvScheduler_CreatePhysicalLocation("Create" , dicWorkPerformedInputs)
									If bResult = False Then Exit Function
							ElseIf dicAction="Add..." Then
								bResult=Fn_SISW_SrvScheduler_SearchOperations("SearchAndSelect", dicWorkPerformedInputs)	
								If bResult = False Then Exit Function
							End If
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Asset" , "Company Contacts","Company Locations"
							'JavaTable
							objDialog.JavaStaticText("CommonDownArrowLabel").SetTOProperty "label",trim(dicItem(iCount))&":"
							Set dicCompanyLocationsInputs = dicValue(iCount)
							dicAction=dicCompanyLocationsInputs("sAction")
							Call Fn_UI_JavaStaticText_Click(" Fn_SISW_SrvScheduler_CreateWorkOrder", objDialog, "CommonDownArrow", 1, 1, "LEFT")
							Wait 2
							bResult=Fn_UI_JavaMenu_Select("",objDialog,dicAction)
							If bResult = False Then Exit Function
							Wait 2
							If dicAction="Create..."	 Then
								If Instr(1, dicItem(iCount), "Asset") > 0 Then
			
								ElseIf Instr(1, dicItem(iCount), "Company Contacts") > 0 Then
			
								ElseIf Instr(1, dicItem(iCount), "Company Locations") > 0 Then
										bResult = Fn_SISW_SrvScheduler_CreateCompanyLocation("Create" , dicCompanyLocationsInputs)
								End If
			
								If bResult = False Then Exit Function
							ElseIf dicAction="Add..." Then
								bResult=Fn_SISW_SrvScheduler_SearchOperations("SearchAndSelect", dicCompanyLocationsInputs)	
								If bResult = False Then Exit Function
							End If
					   End Select
				Next

		Case "Remove"
				dicItem = dicInputs.Keys
				dicValue = dicInputs.Items
				For iCount = 0 to dicInputs.Count - 1
					Select Case trim(dicItem(iCount))
					
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Work Performed At"
							objDialog.JavaStaticText("CommonDownArrowLabel").SetTOProperty "label",trim(dicItem(iCount))&":"
							Call Fn_UI_JavaStaticText_Click(" Fn_SISW_SrvScheduler_CreateWorkOrder", objDialog, "CommonDownArrow", 1, 1, "LEFT")
							Wait 2
							bResult=Fn_UI_JavaMenu_Select("",objDialog,"Clear")
							If bResult = False Then Exit Function
							
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Company Contacts", "Asset", "Company Locations"
							iRowCount=objDialog.JavaTable(dicItem(iCount)).GetROProperty("rows")
							For iCounter = 0 To iRowCount-1
								sValue = objDialog.JavaTable(dicItem(iCount)).GetCellData(iCounter,0)
								If sValue= dicValue(iCount) Then
									objDialog.JavaTable(dicItem(iCount)).SelectRow iCounter
									Exit For
								End If
							Next
	
							'JavaTable
							objDialog.JavaStaticText("CommonDownArrowLabel").SetTOProperty "label",trim(dicItem(iCount))&":"
							Call Fn_UI_JavaStaticText_Click(" Fn_SISW_SrvScheduler_CreateWorkOrder", objDialog, "CommonDownArrow", 1, 1, "LEFT")
							Wait 2
							bResult=Fn_UI_JavaMenu_Select("",objDialog,"Remove")
							If bResult = False Then Exit Function
					
					   End Select
				Next

		Case "Verify"
				dicItem = dicInputs.Keys
				dicValue = dicInputs.Items
				For iCount = 0 to dicInputs.Count - 1
					Select Case trim(dicItem(iCount))
						Case "ID","Revision"
							' In Progress							
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Synopsis"
							' In Progress							
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Work Performed At","Plan"
							sValue = objDialog.JavaObject(dicItem(iCount)).Object.getText()
							If sValue <> dicValue(iCount) Then
								Exit Function
							End If

						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Company Contacts", "Asset", "Company Locations"
							arrValue=Split(dicValue(iCount), "~")
							For intCount = 0 To UBound(arrValue)
								bResult = False
								iRowCount=objDialog.JavaTable(dicItem(iCount)).GetROProperty("rows")
								For iCounter = 0 To iRowCount-1
									sValue = objDialog.JavaTable(dicItem(iCount)).GetCellData(iCounter,0)
									If sValue= arrValue(intCount) Then
										bResult = True
										Exit For
									End If
								Next
								If bResult = False Then
									Exit Function
								End If
							Next
					   End Select
				Next
		Case Else
			Exit Function
	End Select
    
	If sButtonName <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan", "Click", objDialog,sButtonName)
	End If
	Fn_SISW_SrvScheduler_CreateWorkOrder=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan:Successfully created Schedule for Work Order Plan")
	Set objDialog = Nothing
End Function

'*********************************************************		Function to create plan value while Creating Work Order On Any node ***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan

'Description			 :		 		  This function is used to create plan value while Creating Work Order On Any node

'Parameters			   :	 			1.  sAction
'													 2. dicInputs: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		BomTable Should be visible in Service Scheduler

'Examples				:              	set dic1 = CreateObject( "Scripting.Dictionary" )
'													dic1("Name") = "temp"
'													dic1("StartDate") = "10-Aug-2013~6:00:00 PM"
'													dic1("FinishDate") = "10-Aug-2014~6:00:00 PM"
'													dic1("Is Schedule Public") = "False"
'													dic1("Is Percent Linked") = "True"
'													dic1("Published") = "True"
													
'													Call Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan("Create" ,  dic1)

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			09-Oct-2013			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan(sAction , dicWopInputs)

GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan"
Dim objDialog, dicItem,dicValue,aRadioButtonInputs

Set objDialog=Fn_SISW_SrvScheduler_GetObject("CreateScheduleForWOP")

If Fn_UI_ObjectExist("Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan",objDialog) = false then
	objDialog.SetTOProperty "index",3
	If Fn_UI_ObjectExist("",objDialog) = false then
		 Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan = FALSE
		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan ] Failed to find [ Create Schedule For Work Order Plan ] window.")
        Exit Function
	End If	
End if	

   Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan=False
	Select Case sAction
		Case "Create"
			'Do Nothing
		Case Else
			Exit Function
	End Select

    dicItem = dicWopInputs.Keys
	dicValue = dicWopInputs.Items
	For iCount = 0 to dicWopInputs.Count - 1
		Select Case trim(dicItem(iCount))
        	Case "ID","Revision","TimeZone"
				'JavaList
				'Not Done yet
			 Case "Name"
				 'EditBox
				bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan", "Set",  objDialog, trim(dicItem(iCount)), trim(dicValue(iCount)) )
				If bResult = False Then
                    Exit Function
				End If
		    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			Case "Is Schedule Public" , "Is Percent Linked","Published","Are notifications enabled","Use Finish Date Scheduling"
				'Radio Button
				objDialog.JavaStaticText("CommonRadioButtonLabel").SetTOProperty "label",trim(dicItem(iCount))&":"
				If trim(dicValue(iCount))="False" Then
                    objDialog.JavaRadioButton("CommonRadioButton").SetTOProperty "attached text", "false"
				End If
				bResult = Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan", "Set", objDialog, "CommonRadioButton", "ON")
				If bResult = False Then
                  	Exit Function
				End If
				objDialog.JavaRadioButton("CommonRadioButton").SetTOProperty "attached text", "true"
		    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------              			
			Case "StartDate","FinishDate"
			
					If trim(dicItem(iCount)) = "StartDate"  Then
						objDialog.JavaStaticText("PropertyName").SetTOProperty "label", "Start Date:"
					ElseIf trim(dicItem(iCount)) = "FinishDate"  Then
						objDialog.JavaStaticText("PropertyName").SetTOProperty "label", "Finish Date:"
					End If

					aDate = Split( trim(dicValue(iCount)) , "~" )
					objDialog.JavaEdit("Date").Set aDate(0)
					wait 1
					call Fn_KeyBoardOperation("SendKeys", "{TAB}")
					objDialog.JavaList("Time").Type aDate(1)
				     If Err.Number <  0 Then
						 Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan = FALSE				 
						Exit Function
					End If
           End Select
	Next
	If objDialog.JavaButton("Finish").GetROProperty("enabled") = "1" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan", "Click", objDialog,"Finish")
	 Else
		Exit Function
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan:Successfully created Schedule for Work Order Plan")
	Fn_SISW_SrvScheduler_CreateScheduleForWorkOrderPlan= True
	Set objDialog = Nothing
End Function

'*********************************************************		Function to create Work Performed At value while Creating Work Order On Any node ***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_CreatePhysicalLocation

'Description			 :		 		  This function is used to 'Create Physical Location' for 'Work Performed At' value while Creating Work Order On Any node

'Parameters			   :	 			1.  sAction
'													 2. dicInputs: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		Create Work Order Window should be Visible

'Examples				:              	set dic1 = CreateObject( "Scripting.Dictionary" )
'													dic1("ID") = "123213"
'													dic1("LocationType") = "Contractor, Contractor"
'													dic1("LocationName") = "abc"
'													dic1("Description") = "desc"
													
'													Call Fn_SISW_SrvScheduler_CreatePhysicalLocation("Create" ,  dic1)

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			09-Oct-2013			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_CreatePhysicalLocation(sAction , dicCPLInputs)

GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_CreatePhysicalLocation"
Dim objDialog, dicItem,dicValue,aRadioButtonInputs

Set objDialog=Fn_SISW_SrvScheduler_GetObject("CreatePhysicalLocation")

   Fn_SISW_SrvScheduler_CreatePhysicalLocation=False
	Select Case sAction
		Case "Create"
			'Do Nothing
		Case Else
			Exit Function
	End Select

    dicItem = dicCPLInputs.Keys
	dicValue = dicCPLInputs.Items
	For iCount = 1 to dicCPLInputs.Count - 1
		Select Case trim(dicItem(iCount))
        	Case "ID"
				'JavaList
				'Not Done yet
			 Case "LocationName","Description"
				 'EditBox
				bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_CreatePhysicalLocation", "Set",  objDialog, trim(dicItem(iCount)), trim(dicValue(iCount)) )
				If bResult = False Then
                    Exit Function
				End If
			Case "LocationType"
				If objDialog.JavaList(trim(dicItem(iCount))).Exist(3) Then
					bResult = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_CreatePhysicalLocation", "Select", objDialog,trim(dicItem(iCount)), trim(dicValue(iCount)), "", "")
					If bResult = False Then
						Exit Function
					End If
				End If 		
        End Select
   	Next
	wait 2
	If objDialog.JavaButton("Finish").GetROProperty("enabled") = "1" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_CreatePhysicalLocation", "Click", objDialog,"Finish")
	 Else
		Exit Function
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_CreatePhysicalLocation:Successfully created Physical Location")
	Fn_SISW_SrvScheduler_CreatePhysicalLocation=True
	Set objDialog = Nothing
End Function
'******************************************************************************************************************************************************************************************
'********************** Function to perform operations on Search dialog in Service Manager***************************************
'
''Function Name		 	:	Fn_SISW_SrvScheduler_SearchOperations
'
''Description		    :	Function to perform operations on Search dialog in Service Schedule
'
''Parameters		    :	1. sAction : Action need to perform
'					  					2. dicSearch : Dictionary object to set Search criteria.
'								
'Return Value		    :  		True / False
'
'Pre-requisite		    :		Search dialog should be already opened in Service Schedule perspective.

''Examples  			:	Dim dicSearch
'					  		Set dicSearch = CreateObject("Scripting.Dictionary")
'					  		dicSearch("ID") = "0*"
'					  		dicSearch("SearchResults_Select") = "000051-P1"
'					  		msgbox Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dicSearch)

'History:
'	Developer Name			Date			Rev. No.			Changes Done																		Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		10-Oct-2013			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle		  03-Dec-2013			1.1				Modified function to handle all objects
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_SearchOperations(sAction, dicSearch)

	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_SearchOperations"
	Dim objDialog,dicItem, dicKey

	Set objDialog = Fn_SISW_SrvScheduler_GetObject("Search")
 	Fn_SISW_SrvScheduler_SearchOperations = False
	Select Case sAction
		Case "SearchAndSelect"
			If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_SearchOperations","Exist", objDialog,"") = False Then
				Set objDialog = Fn_SISW_SrvScheduler_GetObject("Search2")
				If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_SearchOperations","Exist", objDialog,"") = False Then
					Exit Function
				End If
			End If
			objDialog.JavaTab("SearchTab").Select "Search"
			dicItem = dicSearch.Keys
			dicValue = dicSearch.Items
			For iCount = 1 to dicSearch.Count - 2
				objDialog.JavaStaticText("PropertyName").SetTOProperty "label", dicItem(iCount)&":"
				bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_SearchOperations", "Set",  objDialog, "PropertyValue", trim(dicValue(iCount)))
				If bResult = False Then
					Fn_SISW_SrvScheduler_SearchOperations=False
					Exit Function
				End If
			Next
			Call Fn_Button_Click("Fn_SISW_SrvScheduler_SearchOperations", objDialog, "Find")
			Wait 5
			Call Fn_ReadyStatusSync(4)
            bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_SearchOperations", "ClickCell", objDialog , "SearchResultTable", "", "Object",trim(dicSearch("SearchResults_Select")), "", "", "", "")
			If bResult = False Then
				If Fn_UI_ObjectExist("Fn_SISW_SrvScheduler_SearchOperations",objDialog.JavaButton("LoadAll")) = True Then
					Call Fn_Button_Click("Fn_SISW_SrvScheduler_SearchOperations",objDialog,"LoadAll") 	
					Call Fn_ReadyStatusSync(1)
					If Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_SearchOperations", "ClickCell", objDialog , "SearchResultTable", "", "Object",trim(dicSearch("SearchResults_Select")), "", "", "", "") = False Then
						Exit Function
					End If
				Else
					Exit Function
				End If
			End If
	End Select

	Call Fn_Button_Click("Fn_SISW_SrvScheduler_SearchOperations", objDialog , "OK")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvScheduler_SearchOperations ] Successfully executed with case [ " & sAction & " ].")
	Fn_SISW_SrvScheduler_SearchOperations=True
	Set objDialog = Nothing
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'******************************************************************************************************************************************************************************************
'********************** Function to Create Part Request***************************************
'
''Function Name		 	:	Fn_SISW_SrvScheduler_PartRequestTypeCreate
'
''Description		    :	Function to Create Part Request
'
''Parameters		    :	1. dicPartRequest : Dictioary of parameters required for creating Part Request
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	

''Examples  			:	Dim dicItem
'					  		Set dicItem = CreateObject("Scripting.Dictionary")
'							dicItem("ID") = "000019"
'					  		dicItem("SearchResults_Select") = "000019-Item1"
'                           
'							Dim dicPartRequest
'					  		Set dicPartRequest = CreateObject("Scripting.Dictionary")
'					  		dicPartRequest("Name") = "TestPartReq"
'					  		dicPartRequest("Description") = "Test"
'					  		dicPartRequest("QuantityRequested") = 1
'					  		Set dicPartRequest("Item") = dicItem
'					  		msgbox Fn_SISW_SrvScheduler_PartRequestTypeCreate(dicPartRequest)

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pritam Shikare		 18-Oct-2013		1.0					Created                 Pallavi J.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_PartRequestTypeCreate(dicPartRequest)

	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_PartRequestTypeCreate"
   Dim bReturn, iCount, arrDicKeys, arrDicValues, sMenuXml, sMenu
   Dim objWindow, dicItem 

   Fn_SISW_SrvScheduler_PartRequestTypeCreate = False

   'Set the Required Object
   Set objWindow = Fn_SISW_SrvScheduler_GetObject("CreatePartRequestType")

   'Check the Existence of the Window Create Part Request Type
   If objWindow.Exist(SISW_MIN_TIMEOUT) = False Then
	   'If  Window Create Part Request Type do not exists then Call Menu File:NewPart Request...:
	    sMenuXml = Fn_LogUtil_GetXMLPath("ServiceScheduler")
		sMenu = Fn_GetXMLNodeValue(sMenuXml, "FileNewPartRequest")
	    bReturn = Fn_MenuOperation("Select", sMenu)
		If bReturn = False Then 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Failed to select menu [ "+sMenu+" ]" )
			Exit Function
		End If
		 'Check the Existence of the Window Create Part Request Type. If Not Exist then Exit the function
	    If objWindow.Exist(SISW_DEFAULT_TIMEOUT) = False Then
		   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: The Create Part Request Type window does not exists" )
		   Exit Function
	   End If
   End If

   'Take the Keys and Value of the dicPartRequest into arrays
   arrDicKeys = dicPartRequest.Keys
   arrDicValues = dicPartRequest.Items

   'Perform operations 
   For iCount = 0 to uBound(arrDicKeys)
	   Select Case arrDicKeys(iCount)
			'------------Enter Part Number if field is enabled---------------------------------
	 		Case "PartNumber"
				bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_PartRequestTypeCreate","Set",objWindow,"PartNumber",arrDicValues(iCount))
				If bReturn = False Then Exit Function

			'------------Select Item ----------NOTE : Pass the dictionary dicItem to dicPartRequest("Item"), Refer the example given above-----------------------
			Case "Item"
				Set dicItem=dicPartRequest("Item")
				Call Fn_UI_JavaStaticText_Click(" Fn_TcObjectDelete", objWindow, "ItemDropDown", 1, 1, "LEFT")
				Wait 2
				bReturn=Fn_UI_JavaMenu_Select("",objWindow,"Relate Item...")
				If bReturn = False Then Exit Function
				Wait 2
				bReturn = Fn_SISW_SrvScheduler_SearchOperations("SearchAndSelect" , dicItem)
				If bReturn = False Then Exit Function

			'------------Enter Name---------------------------------
			Case "Name"
				bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_PartRequestTypeCreate","Set",objWindow,"Name",arrDicValues(iCount))
				If bReturn = False Then Exit Function

			'------------Enter Description---------------------------------
			Case "Description"
				bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_PartRequestTypeCreate","Set",objWindow,"Description",arrDicValues(iCount))
				If bReturn = False Then Exit Function

			'------------Select Date of Initiation---------------------------------
			Case "InitiationDate"
				'to be implemented

			'------------Select Due Date---------------------------------
			Case "DueDate"
				'to be implemented

			'------------Select QuantityRequested---------------------------------
			Case "QuantityRequested"
				bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_PartRequestTypeCreate","Set",objWindow,"QuantityRequested",arrDicValues(iCount))
				If bReturn = False Then Exit Function

			Case Else
				Exit Function
	   End Select
	Next

	'Click Finish Button
	bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_PartRequestTypeCreate", "Click", objWindow, "Finish")
    If bReturn = False Then Exit Function

	'Return True if Operation performed Successfully
	Set objWindow = Nothing
	Fn_SISW_SrvScheduler_PartRequestTypeCreate = True
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_SrvScheduler_GetTreeItemPath(objTree, sNode, sDelimiter, sInstanceHandler)

'Description		:	Function used to Perform operations on Component Tab of Service Scheduler


'Parameters			:	1) objTree: Object of a tree.
'						2) sNode: Node Path
'						3) sDelimiter: for future use.
'						4) sInstanceHandler: for future use.
											
'Return Value		: 	item path \ FALSE

'Pre-requisite		:	Service Scheduler perspective should be present

'Examples			:	Call Fn_SISW_SrvScheduler_GetTreeItemPath(JavaWindow("ServiceScheduler").JavaTree("ViewTree"), "Work Order Home:My Open Works:PR-000003/A;1-test", "", "")

'History:
'	Developer Name			Date			Rev. No.		sChanges Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		22-Oct-2013			1.0				Created                            Pritam Shikare
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function  Fn_SISW_SrvScheduler_GetTreeItemPath(objTree, sNode, sDelimiter, sInstanceHandler)

	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_GetTreeItemPath"
	Dim iCnt, aNodeArr, iArrCnt, iItemCnt
	Dim objCurrTreeItm, sTreeNodeToStr2
	Dim sItmPath
	If sDelimiter = "" Then
		sDelimiter = ":"
	End If

	Fn_SISW_SrvScheduler_GetTreeItemPath = False

	aNodeArr = split(sNode, sDelimiter, -1, 1)
	set objCurrTreeItm = objTree.Object.getItem(0)

	If sNode <> "" Then
		sTreeNodeToStr2 = ""
		sTreeNodeToStr2 = trim(objTree.Object.getItem(0).getData().toString())
		If sTreeNodeToStr2 = "" Then
			sTreeNodeToStr2 = trim(objTree.Object.getItem(0).getData().getDisplayName())
		End If
		If  sTreeNodeToStr2 = trim(aNodeArr(0)) Then
			sItmPath = "#0"
		 Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Capture Root Node  " + aNodeArr(0) + "of Change Manager Tree." )	
			Fn_SISW_SrvScheduler_GetTreeItemPath = False
			Exit Function			
		End If
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to sNode parameter is empty." )	
			Exit Function 
	End If	

	For iArrCnt = 1 to UBOund(aNodeArr)
		iItemCnt = cInt(objCurrTreeItm.getItemCount())
		' Added By Ashok kakade
		aNodePath = split(trim(aNodeArr(iArrCnt)), "@")
		If uBound(aNodePath) > 0 Then
				aNodePath(0) = trim(aNodePath(0))
				iInstance = cint(aNodePath(1))
			Else
				iInstance = 1
			End If
		For iCnt = 0 to iItemCnt -1
			sTreeNodeToStr2 = ""
			sTreeNodeToStr2 = trim(objCurrTreeItm.getItem(iCnt).getData().toString()) 
			If sTreeNodeToStr2 = "" Then
				sTreeNodeToStr2 = trim(objCurrTreeItm.getItem(iCnt).getData().getDisplayName()) 
			End If
			If instr(sTreeNodeToStr2,"(") > 0 Then
				If instr(sTreeNodeToStr2, aNodePath(0)) > 0 Then
					If  iInstance = 1 Then
						sItmPath = sItmPath + ":#" +  cstr(iCnt)
						set objCurrTreeItm = objCurrTreeItm.getItem(iCnt)
						Exit For
					Else
						iInstance = iInstance - 1
					End If
				End If
			Else
				If sTreeNodeToStr2 = aNodePath(0) Then
					If  iInstance = 1 Then
						sItmPath = sItmPath + ":#" +  cstr(iCnt)
						set objCurrTreeItm = objCurrTreeItm.getItem(iCnt)
						Exit For
					Else
						iInstance = iInstance - 1
					End If
				End If
			End If
		Next
        If iCnt = iItemCnt Then
			set objCurrTreeItm = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Retrieve Node Index for Change Manager Tree Node [" + aNodeArr(iArrCnt) + "]" )
			Fn_SISW_SrvScheduler_GetTreeItemPath = False
			Exit Function
		End If
	Next
	set objCurrTreeItm = Nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Retrieved Node Index" + sItmPath + "of Change Manager Tree Node [" + sNode + "]" )	
	Fn_SISW_SrvScheduler_GetTreeItemPath = sItmPath
End Function

'-------------------------------------------------------------------Function Used to perform operatons on View Tree---------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_SISW_SrvScheduler_ViewTreeOperations

'Description			 :	Function Used to perform operatons on View Tree

'Parameters			   :	1.strAction: Action Name
										'2.strNodeName: Node Name		
										'3.strMenu: Pop Up Menu Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective

'Examples				:	Fn_SISW_SrvScheduler_ViewTreeOperations("Select","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
										'Fn_SISW_SrvScheduler_ViewTreeOperations("DoubleClick","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
										'Fn_SISW_SrvScheduler_ViewTreeOperations("Expand","Change Home:My Open Changes","")
										'Fn_SISW_SrvScheduler_ViewTreeOperations("VerifyNode","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
										'Fn_SISW_SrvScheduler_ViewTreeOperations("PopupMenuSelect","Change Home:Test1","Refresh")
										'Fn_SISW_SrvScheduler_ViewTreeOperations("MultiSelect","Change Home:My Open Changes:ECN-074691/A;1-CM76,Change Home:My Open Changes:ECN-459679/A;1-CM76,Change Home:My Open Changes:PR-581018/A;1-PR130","")   
										'Fn_SISW_SrvScheduler_ViewTreeOperations("MultiSelect","Change Home:Snd:ECR-954847/A;1-ECR1,Change Home:Snd:ECR-775411/A;1-ECR3","")
										'Fn_SISW_SrvScheduler_ViewTreeOperations("MultiSelectPopupMenu","Change Home:Snd:ECR-516067/A;1-ECR2","Derive Change...")
'										Fn_SISW_SrvScheduler_ViewTreeOperations("Collapse","Change Home:My Open Changes","")
'History					 :			
'	Developer Name				Date					Rev. No.			Changes Done																		Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 	Ashwini Kumar				22-Oct-2013				01																											Pritam S.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' 	Pranav Ingle				  28-Nov-2013			 1.1				Added Code click on Load All														Pritam S.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_ViewTreeOperations(strAction,StrNodePath,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_ViewTreeOperations"
    'Variable declaration
   Dim arrNodeName,ItemCount,iCount,sNode, aMenuList,intCount
   Dim intNodeCount,sTreeItem,NodeName,StrMultiNodePath,arrMultiNodeName,iCnt
   Dim ObjSrvSchdWnd,iPath
	Fn_SISW_SrvScheduler_ViewTreeOperations = False
	'Creating Object of ServiceScheduler window
	Set ObjSrvSchdWnd=Fn_SISW_SrvScheduler_GetObject("ServiceScheduler")
	
	If StrNodePath <> "" Then
		' expanding parent
		arrNodeName = Split(StrNodePath,":")
		If uBound(arrNodeName) <> 0 Then
			For iCount = 0 to uBound(arrNodeName) - 1
				If iCount = 0 Then 
					sNode = arrNodeName(iCount)
				Else
					sNode = sNode &":" & arrNodeName(iCount)
				End If
			Next
			' expanding parent
			iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),sNode,"","")
			If iPath <> False Then ObjSrvSchdWnd.JavaTree("ViewTree").Expand iPath
			wait 1
		End IF
	End IF

	iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),StrNodePath,"","")
	If iPath = False Then
		iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),"Work Order Home:My Open Works:Load All","","")
		 If iPath<>False Then
			ObjSrvSchdWnd.JavaTree("ViewTree").Select iPath
			Fn_SISW_SrvScheduler_ViewTreeOperations=True
		End If
	End If

	Select Case strAction
		'===================================================================================================
		Case "Select" 		'Fn_SISW_SrvScheduler_ViewTreeOperations("Select","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
				iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),StrNodePath,"","")
				If iPath<>False Then
					ObjSrvSchdWnd.JavaTree("ViewTree").Select iPath
					Fn_SISW_SrvScheduler_ViewTreeOperations=True
				End If
		'===================================================================================================
		Case "Expand" 'Fn_SISW_SrvScheduler_ViewTreeOperations("Expand","Change Home:My Open Changes","")
				iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),StrNodePath,"","")
				Call Fn_UI_JavaTree_Expand("Fn_SISW_SrvScheduler_ViewTreeOperations", ObjSrvSchdWnd, "ViewTree",iPath)
				Fn_SISW_SrvScheduler_ViewTreeOperations=True
		'===================================================================================================
		Case "Collapse"
				iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),StrNodePath,"","")
				Call Fn_UI_JavaTree_Collapse("Fn_SISW_SrvScheduler_ViewTreeOperations", ObjSrvSchdWnd, "ViewTree",iPath)
				Fn_SISW_SrvScheduler_ViewTreeOperations=True
		'===================================================================================================
		Case "PopupMenuSelect"		'Fn_SISW_SrvScheduler_ViewTreeOperations("PopupMenuSelect","Change Home:Test1","Refresh")
				iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),StrNodePath,"","")
				If iPath<>False Then
					ObjSrvSchdWnd.JavaTree("ViewTree").Select iPath
					'Open context menu
					wait 1
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_SrvScheduler_ViewTreeOperations",ObjSrvSchdWnd, "ViewTree",iPath)
					Wait 2
					'Select Menu action
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Select Case intCount
						Case "0"
							 StrMenu =ObjSrvSchdWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu =ObjSrvSchdWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu =ObjSrvSchdWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_SISW_SrvScheduler_ViewTreeOperations = False
							Set ObjSrvSchdWnd=Nothing
							Exit Function
					End Select
					ObjSrvSchdWnd.WinMenu("ContextMenu").Select StrMenu
					Fn_SISW_SrvScheduler_ViewTreeOperations=True
				End If				
		'===================================================================================================
		Case "VerifyNode"		'Fn_SISW_SrvScheduler_ViewTreeOperations("VerifyNode","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
				iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),StrNodePath,"","")
				If iPath=False Then
					Fn_SISW_SrvScheduler_ViewTreeOperations = FALSE
				Else
					Fn_SISW_SrvScheduler_ViewTreeOperations=True
				End If
		'===================================================================================================
		Case "DoubleClick"
						Call Fn_SISW_SrvScheduler_ViewTreeOperations("Select",StrNodePath,"") ' Modified By Vidya 22/2/2013
						wait 1
						Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						Fn_SISW_SrvScheduler_ViewTreeOperations=True

''				iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),StrNodePath,"","")  '
''				ObjSrvSchdWnd.JavaTree("ViewTree").Activate iPath
''				Fn_SISW_SrvScheduler_ViewTreeOperations=True
		'===================================================================================================
		 Case "MultiSelect"
				arrMultiNodeName=split(StrNodePath,",",-1,1)
				For iCount = 0 to uBound(arrMultiNodeName)
					iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),arrMultiNodeName(iCount),"","")
					If iCount = 0 Then
						Fn_SISW_SrvScheduler_ViewTreeOperations = Fn_UI_JavaTree_ExtendSelect("Fn_SISW_SrvScheduler_ViewTreeOperations", ObjSrvSchdWnd, "ViewTree",iPath)
					Else
						Fn_SISW_SrvScheduler_ViewTreeOperations = Fn_UI_JavaTree_ExtendSelect("Fn_SISW_SrvScheduler_ViewTreeOperations", ObjSrvSchdWnd, "ViewTree",iPath)
					End If
					If Fn_SISW_SrvScheduler_ViewTreeOperations = False Then exit for
				Next		
		'===================================================================================================
		Case "MultiSelectPopupMenu"		'Fn_SISW_SrvScheduler_ViewTreeOperations("MultiSelectPopupMenu","Change Home:Test1","Refresh")
					arrMultiNodeName=split(StrNodePath,",",-1,1)
					For iCount = 0 to uBound(arrMultiNodeName)
						iPath = Fn_SISW_SrvScheduler_GetTreeItemPath(ObjSrvSchdWnd.JavaTree("ViewTree"),arrMultiNodeName(iCount),"","")
						If iPath = False Then
							Fn_SISW_SrvScheduler_ViewTreeOperations = False 
							exit for
						End If
						If iCount = 0 Then
							ObjSrvSchdWnd.JavaTree("ViewTree").select iPath
							Fn_SISW_SrvScheduler_ViewTreeOperations = True
						Else
							Fn_SISW_SrvScheduler_ViewTreeOperations = Fn_UI_JavaTree_ExtendSelect("Fn_SISW_SrvScheduler_ViewTreeOperations", ObjSrvSchdWnd, "ViewTree",iPath)
						End If
						If Fn_SISW_SrvScheduler_ViewTreeOperations = False Then exit for
					Next
					If Fn_SISW_SrvScheduler_ViewTreeOperations <> False Then
						'Open context menu
						Fn_SISW_SrvScheduler_ViewTreeOperations = Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_SrvScheduler_ViewTreeOperations",ObjSrvSchdWnd, "ViewTree", iPath)
						wait 2
						'Select Menu action
						aMenuList = split(StrMenu, ":",-1,1)
							Select Case Ubound(aMenuList)
								Case 0
									StrMenu =ObjSrvSchdWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
								Case 1
									StrMenu =ObjSrvSchdWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
								Case 2
									StrMenu =ObjSrvSchdWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
								Case Else
									Fn_SISW_SrvScheduler_ViewTreeOperations = False
									Set ObjSrvSchdWnd=Nothing
									Exit Function
							End Select
						ObjSrvSchdWnd.WinMenu("ContextMenu").Select StrMenu				
						Fn_SISW_SrvScheduler_ViewTreeOperations=True
					End If
'		'====================================Not Yet implemented the new workaround for "GetNodeIndex" Case
'		Case "GetNodeIndex"		'Fn_SISW_SrvScheduler_ViewTreeOperations("GetNodeIndex","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
'				Fn_SISW_SrvScheduler_ViewTreeOperations = Fn_CM_getJavaTreeIndex(ObjSrvSchdWnd.JavaTree("ViewTree"), StrNodePath) 
		Case Else
				Fn_SISW_SrvScheduler_ViewTreeOperations=False
	End Select
	'Rleasing Change Manager Window Object
	Set ObjSrvSchdWnd=Nothing
End Function

'*********************************************************		Function to create Work Performed At value while Creating Work Order On Any node ***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_CreateCompanyLocation

'Description			 :		 		  This function is used to 'Create Physical Location' for 'Work Performed At' value while Creating Work Order On Any node

'Parameters			   :	 			1.  sAction
'													 2. dicInputs: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		Create Work Order Window should be Visible

'Examples				:              	set dic1 = CreateObject( "Scripting.Dictionary" )
'													dic1("Name") = "123213"
'													dic1("Street") = "xyz"
'													dic1("City") = "abc"
'													dic1("Description") = "desc"
													
'													Call Fn_SISW_SrvScheduler_CreateCompanyLocation("Create" ,  dic1)

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			12-Nov-2013			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_CreateCompanyLocation(sAction , dicCCLInputs)

GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_CreateCompanyLocation"
Dim objDialog, dicItem,dicValue,aRadioButtonInputs

Set objDialog=Fn_SISW_SrvScheduler_GetObject("CreateCompanyLocation")

   Fn_SISW_SrvScheduler_CreateCompanyLocation=False
	Select Case sAction
		Case "Create"
			'Do Nothing
		Case Else
			Exit Function
	End Select

    dicItem = dicCCLInputs.Keys
	dicValue = dicCCLInputs.Items
	For iCount = 1 to dicCCLInputs.Count - 1
		'EditBox
		bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_CreateCompanyLocation", "Set",  objDialog, trim(dicItem(iCount)), trim(dicValue(iCount)) )
		If bResult = False Then
			Exit Function
		End If 		
   	Next
	wait 2
	If objDialog.JavaButton("Finish").GetROProperty("enabled") = "1" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_CreateCompanyLocation", "Click", objDialog,"Finish")
	 Else
		Exit Function
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_CreateCompanyLocation:Successfully created Physical Location")
	Fn_SISW_SrvScheduler_CreateCompanyLocation=True
	Set objDialog = Nothing
End Function
'******************************************************************************************************************************************************************************************

'*********************************************************		Function to  get  Schedule table Row Index	***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_TreeTableRowIndex

'Description			 :		 		  This function is used to get Schedule table Row Index.

'Parameters			   :	 			1.  sNodeName:Name of the Node to retrieve Index for.
											
'Return Value		   : 				 Node index

'Pre-requisite			:		 		Service Schedular window should be displayed .

'Examples				:				 Fn_SISW_SrvScheduler_TreeTableRowIndex("Sch1:Job1")

'History:
'	Developer Name			Date			Rev. No.	Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			14-Nov-2013		1.0			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_TreeTableRowIndex(objTreeTable, sNodeName, sColname)

	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_TreeTableRowIndex"
	
	Dim IntRows ,sNodePath, IntCounter, StrIndex, ArrNode
	Dim iGlobalCnt, aNodePath, iInstance, iLoop

	On Error Resume Next

	iGlobalCnt = 0

	 If instr(sNodeName, "@") > 0 Then
		aNodePath = split(sNodeName, "@",-1, 1)
		sNodeName = aNodePath(0)
		iInstance = cint(aNodePath(1))
	Else
		iInstance = 1
	End If
	
	'Verify that PSE BOM Table is displayed
	If objTreeTable.Exist(iTimeOut) Then

		'Get the No. of rows present in the BOM Table
		IntRows = objTreeTable.GetROProperty("rows")
		ArrNode = split(sNodeName, ":",-1,1)

       'Get the Row No. of required Node
	   For iLoop = 1 to iInstance
				For IntCounter = iGlobalCnt to IntRows -1
					If Trim(objTreeTable.Object.getRow(IntCounter).tostring)  = Trim(ArrNode(Ubound(ArrNode))) Then
						objTreeTable.SelectRow IntCounter
						sNodePath = objTreeTable.GetCellData(IntCounter, sColname)
						If Trim(sNodePath) = Trim(sNodeName) Then
							StrIndex = "#" + Cstr(IntCounter)
							iGlobalCnt = Cint (IntCounter) + 1
							Exit For
						Else
							'Do Nothing
						End If
					End If
				Next
	   Next

		Fn_SISW_SrvScheduler_TreeTableRowIndex = StrIndex
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_TreeTableRowIndex: Row Index of [" + sNodeName +"] Node is [" + StrIndex + "]")	

		If  cint(IntCounter) = cint(IntRows) Then
			Fn_SISW_SrvScheduler_TreeTableRowIndex = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_TreeTableRowIndex:Failed to Get  Row Index of [" + sNodeName +"]")	
		End If

  End If
End Function

'******************************************************************Function to perform schedule table  operation************************************************************************************************************

'Function Name					:Fn_SISW_SrvScheduler_SchTable_NodeOperation

'Description						Actions performed in this function are:
'                                              1. Node Select
'                                               2. Node multi-select
'                                              3. Node Expand
'                                              4. Node Collapse
'                                              5. Node Popup menu select
'                                              6. Cell Verify
'                                             7. Cell edit   ( Pass column index.)
'                                            8. Cell double-click
'                                            9.Exists
'											10.MultiSelectPopup
'

'Parameters						:1. sAction: Action to be performed
'                                            2. sObject: Name of the object ( , is use as seperator)
'                                           3. sColName: Column name or index ( For cell edit case pass column index only)
'                                          4. sValue: Value to be set
'                                         5. sMenu: Context menu to be selected
'											  

'Return Value		   : 	    TRUE \ FALSE

'Pre-requisite					:Schedule manger panel need to open

'Examples						:	   call Fn_SISW_SrvScheduler_SchTable_NodeOperation ( "PopupMenu", "T3", "", "" , "New:Job Card..." )
'											call Fn_SISW_SrvScheduler_SchTable_NodeOperation("Select" , "test3" , "" , "" , "")									

'History:
'										Developer Name			Date			Rev. No.			Changes Done						Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ashwini Kumar   		14/Nov/13			1.0  									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function  Fn_SISW_SrvScheduler_SchTable_NodeOperation(sAction , sObject , sColName , sValue , sMenu)

	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_SchTable_NodeOperation"
	Dim objSchTable,sIndex,iCounter, bReturn,iInstance,aNodePath
	Dim objSelectType,iCnter,sPath

	Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
	iInstance = 1

	'Create object of  table
	Set objSchTable =Fn_SISW_SrvScheduler_GetObject("SchJobTable")
    objSchTable.Click 0,0
	wait SISW_MICRO_TIMEOUT
	Select Case sAction
		'.---------------------------------------This case is used to select the Schedule Table Node.----------------------------------------------
		Case "Select"
			If instr(sObject, "@") <> 0 Then
				aNodePath = split(sObject, "@",-1, 1)
				sObject = aNodePath(0)
				iInstance = cint(aNodePath(1))
				iCnter = 0
				sPath=split(sObject, ":",-1, 1)
				For iCounter=0 to objSchTable.Object.getRowCount-1			
					If  sPath(0)+":"+objSchTable.Object.getRow(iCounter).tostring()= sObject Then
						iCnter = iCnter+1
						If  iCnter = iInstance Then
							sIndex = iCounter
							objSchTable.SelectRow sIndex
							Exit For
						End If
					End If
				Next
				If iCnter <> iInstance Then
					 Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation : Row with Object "&sObject&" does not selected.")	
						Exit Function
				Else
						Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation: Row with Object "&sObject&" is selected.")	
				End If
			Else
				sIndex = Fn_SISW_SrvScheduler_TreeTableRowIndex(objSchTable, sObject, "Object")
				If Instr(sIndex, "#") > 0 Then
					'Select the Expected  scheduleTable Node
					 objSchTable.SelectRow sIndex
				     If Err.Number <  0 Then
						 Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation : Row with Object "&sObject&" does not selected.")	
						Exit Function
					Else
						Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation: Row with Object "&sObject&" is selected.")	
					End If
				Else
					Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation : Row with Object "&sObject&" does not exists in schedule table..")	
					Exit Function
				End If
			End If
			'.---------------------------------------This case is used to multiselect the Shedule Table Node.----------------------------------------------
			Case "MultiSelect"
			aNodePath = split(sObject, "~", -1, 1)
			objSchTable.Object.clearSelection
			For iCounter = 0 To UBound(aNodePath)
				sIndex = Fn_SISW_SrvScheduler_TreeTableRowIndex(objSchTable, aNodePath(iCounter), "Object")
				If Instr(sIndex, "#") > 0 Then
					aNodePath(iCounter) = sIndex
'					'Select the Expected  scheduleTable Node
'					 Err.Clear
'					 objSchTable.ExtendRow sIndex
'				     If Err.Number <  0 Then
'						 Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation : Row with Object "&aNodePath(iCounter)&" does not selected.")	
'						Exit Function
'					Else
'						Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE				 
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation: Row with Object "&aNodePath(iCounter)&" is selected.")	
'					End If
				Else
					Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation : Row with Object "&sObject&" does not exists in schedule table..")	
					Exit Function
				End If				
			Next
			objSchTable.Object.clearSelection
			For iCounter = 0 To UBound(aNodePath)
				 Err.Clear
				 objSchTable.ExtendRow aNodePath(iCounter)
			     If Err.Number <  0 Then
					 Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation : Row with Object "&aNodePath(iCounter)&" does not selected.")	
					Exit Function
				Else
					Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation: Row with Object "&aNodePath(iCounter)&" is selected.")	
				End If
			Next

		 '.---------------------------------------This case is used to expand all the Shedule Table Node.----------------------------------------------
		Case	 "Expand" 
            bReturn = Fn_MenuOperation("Select", "View:Expand All")
			If bReturn = TRUE Then
				Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE				 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation:Expand all the node.")
			Else
			   Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation : Fail to expand the node..")	
			  Exit Function
			End If

		'.---------------------------------------This case is used to  RMB  click  and  select  pop menu  of   Schedule Table Node cell.----------------------------------------------
		Case	 "PopupMenu"
			sIndex = Fn_SISW_SrvScheduler_TreeTableRowIndex(objSchTable, sObject, "Object")
 			If Instr(sIndex, "#") > 0 Then
				objSchTable.SelectRow sIndex
				objSchTable.ClickCell sIndex,"Object","RIGHT","NONE"
				bReturn= Fn_UI_JavaMenu_Select("",JavaWindow("ServiceScheduler"),sMenu)					
				If bReturn =False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation : Failed to perform RMB.")
					 Exit Function
				End If
				Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE				 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation passed with case "&sAction&" on Object "&sObject)
            Else
			    Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
				Exit Function
			End If

		'.---------------------------------------This case is used  to check exists of node in schedule table .  .----------------------------------------------
		Case "Exists"
				Dim sActualval
				sIndex = Fn_SISW_SrvScheduler_TreeTableRowIndex(objSchTable, sObject, "Object")
				

              If Instr(sIndex, "#") > 0 Then
				objSchTable.SelectRow sIndex
				sActualval = objSchTable.GetCellData(sIndex,"Object")

				If  sActualval = sObject  Then
					Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_SISW_SrvScheduler_SchTable_NodeOperation: Row with Object "&sObject&" is exists.")	
				Else 
					Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SISW_SrvScheduler_SchTable_NodeOperation : Row with Object "&sObject&" does not exists.")	
					Exit Function
			   End If
            Else 
				Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE				 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SISW_SrvScheduler_SchTable_NodeOperation : Row with Object "&sObject&" does not exists.")	
				Exit Function	
			End If
			
			'---------------------------------------This case is used to get the Cell Value For BOM Table Node cell.----------------------------------------------
		Case "GetCellData"
			If sObject <> "" Then
				iRowCount = objSchTable.GetROProperty("rows")
				For iCounter = 0 To (iRowCount - 1)
					objSchTable.SelectRow iCounter
					sPath = objSchTable.GetCellData(iCounter,"Object")
					If sPath = sObject Then
							Fn_SISW_SrvScheduler_SchTable_NodeOperation = objSchTable.GetCellData(iCounter,sColName)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SrvScheduler_SchTable_NodeOperation] Cell verified of PSE BOM Table Node [" + sObject + "]")
							Exit Function
					End If
				Next
			End If
		Case "CellVerify"
			If sObject <> "" Then
				iRowCount = objSchTable.GetROProperty("rows")
				For iCounter = 0 To (iRowCount - 1)
					objSchTable.SelectRow iCounter
					sPath = objSchTable.GetCellData(iCounter,"Object")
					If sPath = sObject Then
							bReturn = objSchTable.GetCellData(iCounter,sColName)
							If instr(trim(bReturn), trim(sValue)) > 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SrvScheduler_SchTable_NodeOperation] Cell verified of SchJobTable Table Node [" + sObject + "]")
								Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE	
								Exit Function
							End If
					End If
				Next
				If iCounter = iRowCount Then
					Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SrvScheduler_SchTable_NodeOperation] failed to verify Cell value of SchJobTable Table Node [" + sObject + "]")
				End If
			End If
		Case "CellEdit", "SetCellData"
			If sObject <> "" Then
				iRowCount = objSchTable.GetROProperty("rows")
				For iCounter = 0 To (iRowCount - 1)
					objSchTable.SelectRow iCounter
					sPath = objSchTable.GetCellData(iCounter,"Object")
					If sPath = sObject Then
						Err.Clear
						objSchTable.SetCellData iCounter, sColName, sValue
						If Err.Number < 0 Then
							Fn_SISW_SrvScheduler_SchTable_NodeOperation = False	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_SISW_SrvScheduler_SchTable_NodeOperation] Cell edit of  SchJobTable Node [" + sObject + "]")
						Else
							Fn_SISW_SrvScheduler_SchTable_NodeOperation = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_SISW_SrvScheduler_SchTable_NodeOperation] Cell edit of SchJobTable Node [" + sObject + "]")
						End If
						Exit Function
					End If
				Next
			End If
	'.---------------------------------------This case is used  to Double click node in schedule table .  .----------------------------------------------			
		Case "CellDoubleClick"
			sIndex = Fn_SISW_SrvScheduler_TreeTableRowIndex(objSchTable, sObject, "Object")
 			If Instr(sIndex, "#") > 0 Then
				objSchTable.SelectRow sIndex
				wait 2
				Err.Clear
				objSchTable.DoubleClickCell sIndex,"Object"
				If Err.Number <> 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation : Failed to perform double click.")
					Exit Function
				End If
				Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation passed with case "&sAction&" on Object "&sObject)
			Else
				Fn_SISW_SrvScheduler_SchTable_NodeOperation = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
				Exit Function
			End If
	End Select
	
	Fn_SISW_SrvScheduler_SchTable_NodeOperation = TRUE
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SISW_SrvScheduler_SchTable_NodeOperation passed with case "&sAction&" on Object "&sObject)
	Set objSchTable = nothing 
End Function

'*********************************************************		Function to Create Job Card On Any node ***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_JobCardOperations

'Description			 :		 		  This function is used to Create Work Order On Any node.

'Parameters			   :	 			1.  sAction
'													 2. sButtonName: Name of Node of BOMTable
'													 3. dicInputs: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		Create Job Card Window should be visible

'Examples				:              set dic1 = CreateObject( "Scripting.Dictionary" )
'											dic1("ID") = "12312"
'											dic1("Revision") = "A"
'											dic1("Name") = "tempName"
'											dic1("WorkEstimate") = "text"
'											dic1("ActivityExecutionType") = "Sign-Off"
'											dic1("FixedType") = "1, Fixed Duration "
'											dic1("Narrative") = "narrativeText "
'											
'											Call Fn_SISW_SrvScheduler_JobCardOperations("Create", dic1, "Finish")
'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			14-Nov-2013			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_SrvScheduler_JobCardOperations(sAction , dicInputs, sButtonName)

	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_JobCardOperations"
	Dim objDialog, dicItem,dicValue,bResult

	Set objDialog=Fn_SISW_SrvScheduler_GetObject("CreateJobCard")
    Fn_SISW_SrvScheduler_JobCardOperations=False

	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_JobCardOperations","Exist", objDialog,"") = False Then
    	Exit Function
		Call Fn_ReadyStatusSync( 3 )
	End If
	dicItem = dicInputs.Keys
	dicValue = dicInputs.Items
	Select Case sAction
		Case "Create"
			For iCount = 0 to dicInputs.Count - 1
				Select Case trim(dicItem(iCount))
				Case "ID"
					'JavaList
					objDialog.JavaList(trim(dicItem(iCount))).Click 1,1,"LEFT"
					objDialog.JavaList(trim(dicItem(iCount))).Type trim(dicValue(iCount))
					If Err.number < 0 Then
						Exit Function
					End If
				'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				Case "Name","WorkEstimate","Narrative"
					'Edit Box
					bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_JobCardOperations", "Set", objDialog, trim(dicItem(iCount)), trim(dicValue(iCount)) )
					If bResult = False Then Exit Function
				'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				 Case "ActivityExecutionType","FixedType"
					bResult=Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_JobCardOperations", "Select", objDialog,trim(dicItem(iCount)),trim(dicValue(iCount)), "", "")
					If bResult = False Then Exit Function
				'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				Case "Asset", "Impacted Part"
					'Not Done Yet
							objDialog.JavaStaticText("CommonDownArrowLabel").SetTOProperty "label",trim(dicItem(iCount))&":"
							Set dicAssetInputs = dicValue(iCount)
							dicAction=dicAssetInputs("sAction")
							Call Fn_UI_JavaStaticText_Click("Fn_SISW_SrvScheduler_JobCardOperations", objDialog, "CommonDownArrow", 1, 1, "LEFT")
							Wait 2
							bResult=Fn_UI_JavaMenu_Select("",objDialog,dicAction)
							If bResult = False Then Exit Function
							Wait 2
							If dicAction="Add..." Then
								If dicItem(iCount) = "Asset" Then
									bResult= Fn_SISW_SrvScheduler_SelectPhysicalAsset("Select", dicAssetInputs("Assets_Select"), "OK")
								ElseIf  dicItem(iCount) = "Impacted Part" Then
									bResult= Fn_SISW_SrvScheduler_MaintenanceTree("Select", dicAssetInputs("Impacted_Part_Select"), "OK")
								End If
								If bResult = False Then Exit Function
							End If
				'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
				End Select
			Next
		Case Else
			Exit Function
	End Select

	If sButtonName <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_JobCardOperations", "Click", objDialog,sButtonName)
	End If
	Fn_SISW_SrvScheduler_JobCardOperations=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_JobCardOperations:Successfully created Job Card.")
	Set objDialog = Nothing
End Function
'*********************************************************		Function to Create Job Task On Any node ***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_JobTaskOperations

'Description			 :		 		  This function is used to Create Work Order On Any node.

'Parameters			   :	 			1.  sAction
'													 2. sButtonName: Name of Node of BOMTable
'													 3. dicInputs: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		Create Job Task Window should be visible

'Examples				:              set dic1 = CreateObject( "Scripting.Dictionary" )
'											dic1("ID") = "12312"
'											dic1("Revision") = "A"
'											dic1("Name") = "tempName"
'											dic1("WorkEstimate") = "text"
'											dic1("ActivityExecutionType") = "Sign-Off"
'											dic1("FixedType") = "1, Fixed Duration "
'											dic1("Narrative") = "narrativeText "
'											
'											Call Fn_SISW_SrvScheduler_JobTaskOperations("Create", dic1, "Finish")
'History:
'	Developer Name			Date				Rev. No.				Changes Done																	Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar			14-Nov-2013			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle			  04-Dec-2013			1.1				Added Cases "StartDate","FinishDate"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle			  04-Dec-2013			1.2				Added Cases "VerifyErrorMessage"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_JobTaskOperations(sAction , dicInputs, sButtonName)

	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_JobTaskOperations"
	Dim objDialog, dicItem,dicValue,bResult

	Set objDialog=Fn_SISW_SrvScheduler_GetObject("CreateJobTask")
    Fn_SISW_SrvScheduler_JobTaskOperations=False

	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_JobTaskOperations","Exist", objDialog,"") = False Then
    	Exit Function
		Call Fn_ReadyStatusSync( 3 )
	End If


	Select Case sAction
		Case "Create"

				dicItem = dicInputs.Keys
				dicValue = dicInputs.Items
				For iCount = 0 to dicInputs.Count - 1
						Select Case trim(dicItem(iCount))
								Case "ID"						',"Revision" - Revision field deprecated
										'JavaList
										objDialog.JavaList(trim(dicItem(iCount))).Click 1,1,"LEFT"
										objDialog.JavaList(trim(dicItem(iCount))).Type trim(dicValue(iCount))
										If Err.number < 0 Then
											Exit Function
										End If
								'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
								Case "Name","WorkEstimate","Narrative"
										'Edit Box
										bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_JobTaskOperations", "Set", objDialog, trim(dicItem(iCount)), trim(dicValue(iCount)) )
										If bResult = False Then Exit Function
								'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
								 Case "ActivityExecutionType","FixedType"
										bResult=Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_JobTaskOperations", "Select", objDialog,trim(dicItem(iCount)),trim(dicValue(iCount)), "", "")
										If bResult = False Then Exit Function
								'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
								Case "StartDate","FinishDate"
											'Date Panel
											'bResult = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_JobTaskOperations", "Click", objDialog,trim(dicItem(iCount)))
											aDate = Split( trim(dicValue(iCount)) , "~" )
											bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_JobTaskOperations", "Set", objDialog, trim(dicItem(iCount)), aDate(0))
											If bResult = False Then
												Exit Function
											End If
											'To set time field
											If trim(dicItem(iCount)) = "StartDate" Then
												bResult = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_JobTaskOperations", "Select", objDialog, "StartTime", aDate(1), "", "")
											Else
												bResult = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_JobTaskOperations", "Select", objDialog, "FinishTime", aDate(1), "", "")
											End If
											'aDate = Split( trim(dicValue(iCount)) , "~" )
											'bResult=Fn_UI_SetDateAndTime("Fn_SISW_SrvScheduler_JobTaskOperations",aDate(0), aDate(1))
											If bResult = False Then
												Exit Function
											End If
								'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
								Case "Asset"
										'Not Done Yet
								'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						End Select
				Next

		Case "VerifyErrorMessage"
				sErrMessage = JavaWindow("ServiceScheduler").JavaWindow("CreateJobTask").JavaEdit("ErrMsg").GetROProperty("value")
				If Instr(1,sErrMessage, dicInputs) <= 0 Then
					Exit Function
				End If
				
		Case "VerifyConfirmation"
				sErrMessage = objDialog.JavaWindow("PleaseConfirm").JavaStaticText("ErrMsg").GetROProperty("label")
				If Instr(1,sErrMessage, dicInputs) <= 0 Then
					Exit Function
				End If
				Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_JobTaskOperations", "Click",  objDialog.JavaWindow("PleaseConfirm"),sButtonName)
				wait 1
				Fn_SISW_SrvScheduler_JobTaskOperations=True
				Exit Function

		Case Else
			Exit Function
	End Select

	If sButtonName <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_JobTaskOperations", "Click", objDialog,sButtonName)
	End If
	Fn_SISW_SrvScheduler_JobTaskOperations=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_JobTaskOperations:Successfully created Job Task.")
	Set objDialog = Nothing
End Function

''*********************************************************		Function to Delete the Object from Structure Manager		**************************************************************
'Function Name		:				Fn_SISW_SrvScheduler_DeleteJobCard.

'Description			 :		 		 This function is used to delete the Object Component from Schedule Manager.

'Parameters			   :	 			1.  sTitle - Title Of Dialog
'												   2. sMessage - message to Verify
'												   3. sAction - Specify the action(i.e.Menu,ShortKey,Toolbar)
'												   4. sButton - Button to Click On
											
'Return Value		   : 				PASS \ FAIL

'Examples				:				 call Fn_SISW_SrvScheduler_DeleteJobCard("Confirmation","Delete the selected task(s)?","Toolbar", "Yes") OR Call Fn_SISW_SrvScheduler_DeleteJobCard("Confirmation","Delete the selected task(s)?","Menu", "Yes")

' History:
'		Developer Name					Date						Rev. No.			Changes Done									Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle 						27-Nov-2013				1.0																						Sunny			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_DeleteJobCard(sTitle,sMessage,sAction, sButton)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_DeleteJobCard"
Dim objDelete, sMsg,sFilePath,sMenu
Set objDelete = JavaDialog("Confirmation")
Fn_SISW_SrvScheduler_DeleteJobCard = FALSE

	Select Case sAction
		' ************************************************************Case To delete Item from Menu option *****************************************************************
		Case "Menu"  
				call Fn_MenuOperation("Select","Edit:Delete")
		'************************************* Case To Delete Item from Toolbar option and click delete icon************************************************************
		Case "Toolbar" 
				sFilePath=Fn_LogUtil_GetXMLPath("RAC_Toolbar")
				sMenu=Fn_GetXMLNodeValue(sFilePath, "Delete")
				call Fn_ToolbatButtonClick(sMenu)
		Case Else
				Exit function
	End Select
	wait 1
	''***********************************************************'' To check whether Delete Dialog  is displayed*************************************************************
	If  sTitle <> "" Then
		objDelete.SetTOProperty "title", sTitle
	End If
	
	If objDelete.Exist(SISW_MICRO_TIMEOUT) Then
		If  sMessage <> "" Then
			sMsg = 	JavaDialog("Confirmation").JavaObject("MLabel").Object.getText()
			If Instr(1, sMsg, sMessage) <= 0   Then
				Call Fn_Button_Click("Fn_SISW_SrvScheduler_DeleteJobCard", objDelete,"Yes")
				Exit Function
			End If
		End If
		Call Fn_Button_Click("Fn_SISW_SrvScheduler_DeleteJobCard", objDelete,"Yes")
	Else
		Exit Function
	End If
	
	Set objDelete = nothing
	Fn_SISW_SrvScheduler_DeleteJobCard = True
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SrvScheduler_SummaryTabOperations

'Description			 :	Function Used to perform operations on Summary Tab

'Parameters			   :   1.StrAction: Action Name
'										2. sTabName : Tab Name select
'										3.dicSummaryInfo: Summary information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Summary tab should be activated

'Examples				:  	Dim dicSummaryInfo
'										Set dicSummaryInfo=CreateObject("Scripting.Dictionary")

'										Case "Verify"	
'										dicSummaryInfo("Physical Parts")="000449/--A~000449/--A:ASM051039"
'										bReturn=Fn_SISW_SrvScheduler_SummaryTabOperations("Parts","Verify",dicSummaryInfo, "")


'										Case "Expand"
'										dicSummaryInfo("Physical Parts")="000449/--A"
'										bReturn=Fn_SISW_SrvScheduler_SummaryTabOperations("Parts","Expand",dicSummaryInfo, "")

'										Case "Select"
'										dicSummaryInfo("Resulting Information")="MSWordX1122"
'										bReturn=Fn_SISW_SrvScheduler_SummaryTabOperations("References","Select",dicSummaryInfo, "")
'
'										Case "ClickButton"
'										dicSummaryInfo("Supporting Information")="Paste"
'										bReturn=Fn_SISW_SrvScheduler_SummaryTabOperations("References","ClickButton",dicSummaryInfo, "")
'
'										Case "ClickJavaObject"
'										dicSummaryInfo("Actions")="Open"
'										bReturn=Fn_SISW_SrvScheduler_SummaryTabOperations("Overview","ClickJavaObject",dicSummaryInfo, "")

'										Case "PopupMenuSelect"
'										dicSummaryInfo("Resulting Information")="MSWordX1122"
'										bReturn=Fn_SISW_SrvScheduler_SummaryTabOperations("References","PopupMenuSelect",dicSummaryInfo, "Copy")

'History					 :			
'					Developer Name					Date									Rev. No.						Changes Done														Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle					03-Dec-2013									1.0																				
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle					12-Dec-2013									1.1						Addde Cases "EditBox","Part Requests", 
'																																					"Part Movements", "Contact Information"																
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle					18-Dec-2013									1.2					Addde Cases "GetCellData" 
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle					2-Jan-2014									1.3					Addde Cases "Recorded Utilization"
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle					17-Jan-2014									1.4					Addde Cases "ClickJavaObject"
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Poonam Chopade					30-Jun-2016									1.0					Added Case "Zone" to perform various operations on tree object of zone panel		Shweta Rathod
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SrvScheduler_SummaryTabOperations(sTabName, StrAction, dicSummaryInfo, StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_SummaryTabOperations"
 	'Declaring variables
    Dim DictItems,DictKeys
    Dim iCounter,bFlag,iCount,aValue, sValue
	Dim aMenuList, intCount, arrTableValues
	Dim sColName, iRowCounter
	Dim objDic, objChilds

	If sTabName<> "" Then
		JavaWindow("ServiceScheduler").JavaTab("SummaryTab").Select sTabName
	End If
	
	'Taking Items & Keys from dictionary
	DictItems = dicSummaryInfo.Items
	DictKeys = dicSummaryInfo.Keys

	Fn_SISW_SrvScheduler_SummaryTabOperations=False
	Select Case StrAction

			Case "Select", "Expand", "PopupMenuSelect", "Verify", "GetCellData", "DoubleClick", "Deselect"
					For iCounter=0 to dicSummaryInfo.count-1
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							'Case to verify value from summary tab
							Select Case DictKeys(iCounter)
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									Case "Resulting Information", "Physical Parts", "Supporting Information"
											aValue=Split(DictItems(iCounter),"~")
											For iCount = 0 To UBound(aValue)
													bFlag=False
													bResult = Fn_UI_getJavaTreeIndex(JavaWindow("ServiceScheduler").JavaTree(DictKeys(iCounter)),aValue(iCount))
													' Modify path in returns 0 To "#0"
													If bResult = 0 Then bResult = "#0"
													 ' Check Condition if path returns false 
													If bResult = False Then Exit Function
	
													' Perform operation 
													If StrAction = "Expand"  Then
														JavaWindow("ServiceScheduler").JavaTree(DictKeys(iCounter)).Expand bResult
													ElseIf StrAction = "Select"  Then
														JavaWindow("ServiceScheduler").JavaTree(DictKeys(iCounter)).Select bResult
													ElseIf StrAction = "DoubleClick" Then
														JavaWindow("ServiceScheduler").JavaTree(DictKeys(iCounter)).Select bResult
														Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
													ElseIf StrAction = "PopupMenuSelect"  Then
														JavaWindow("ServiceScheduler").JavaTree(DictKeys(iCounter)).Select bResult
														wait 1
														Call Fn_UI_JavaTree_OpenContextMenu("Fn_SISW_SrvScheduler_SummaryTabOperations",JavaWindow("ServiceScheduler"), DictKeys(iCounter),bResult)
														Wait 2
														'Select Menu action
														aMenuList = split(StrMenu, ":",-1,1)
														intCount = Ubound(aMenuList)
														Select Case intCount
															Case "0"
																 StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
															Case "1"
																StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
															Case "2"
																StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
															Case Else
																Exit Function
														End Select
														JavaWindow("ServiceScheduler").WinMenu("ContextMenu").Select StrMenu
													End If
													bFlag=true
											Next
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									Case "Part Requests", "Part Movements", "Contact Information","Discovered Discrepancies","Corrected Discrepancies"
											If Instr(DictKeys(iCounter),"Part Movements") Then
												If LCase(JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").GetROProperty("maximized")) = Cstr(LCase(False)) Then
													Call Fn_TabFolder_Operation("DoubleClickTab", "Summary", "")
													Call Fn_ReadyStatusSync(1)
												End If 
											End If
											aValue=Split(DictItems(iCounter),"~")
											JavaWindow("ServiceScheduler").JavaStaticText("PropertyName").SetTOProperty "label", DictKeys(iCounter)
											For iCount = 0 To UBound(aValue)
													arrTableValues = Split(aValue(iCount), "$")
													If UBound(arrTableValues) = 2 Then
														aValue(iCount) = arrTableValues(0)
													End If
													bFlag=False
													bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_SummaryTabOperations", "GetRowIndex", JavaWindow("ServiceScheduler"), "PartRequests", "Object.GetItem", "Object", arrTableValues(0), "", "", "", ":")
													 ' Check Condition if path returns false 
													If bResult = -1 Then Exit Function

													If UBound(arrTableValues) = 2 Then
															arrTableValues(1) = Fn_GetXMLNodeValue(Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectRealNames\ServiceSchedular.xml" , Trim(arrTableValues(1)))
                                                            sValue = JavaWindow("ServiceScheduler").JavaTable("PartRequests").Object.getItem(bResult).getData().getComponent().getProperty(arrTableValues(1))
															If  StrAction = "GetCellData" Then
																Fn_SISW_SrvScheduler_SummaryTabOperations = sValue
															ElseIf StrAction = "Verify" Then
																If Instr(1, sValue,arrTableValues(2)) = 0  Then
																	Exit Function
																End If
															End If	
													End If

													' Perform operation 
													If StrAction = "Select" Then
														JavaWindow("ServiceScheduler").JavaTable("PartRequests").SelectCell bResult,"Object"
'														JavaWindow("ServiceScheduler").JavaTable("PartRequests").SelectRow bResult
													ElseIf StrAction = "PopupMenuSelect" Then
														JavaWindow("ServiceScheduler").JavaTable("PartRequests").SelectCell bResult,"Object"
														wait 1
														JavaWindow("ServiceScheduler").JavaTable("PartRequests").SelectColumnHeader "Object","RIGHT"
														wait 1

														aMenuList = split(StrMenu, ":",-1,1)
														intCount = Ubound(aMenuList)
														Select Case intCount
															Case "0"
																 StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
															Case "1"
																StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
															Case "2"
																StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
															Case Else
																Exit Function
														End Select
														JavaWindow("ServiceScheduler").WinMenu("ContextMenu").Select StrMenu
													ElseIf StrAction = "Deselect" Then
														JavaWindow("ServiceScheduler").JavaTable("PartRequests").DeselectRow bResult
													End If
											Next
											If Instr(DictKeys(iCounter),"Part Movements") Then
												If LCase(JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").GetROProperty("maximized")) = Cstr(LCase(True)) Then
													Call Fn_TabFolder_Operation("DoubleClickTab", "Summary", "")
													Call Fn_ReadyStatusSync(1)
												End If 
											End If
											bFlag=true
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									Case "Recorded Utilization","Resources"
											aValue=Split(DictItems(iCounter),"~")
											JavaWindow("ServiceScheduler").JavaStaticText("PropertyName").SetTOProperty "label", DictKeys(iCounter)
											For iCount = 0 To UBound(aValue)
													arrTableValues = Split(aValue(iCount), "$")
													If UBound(arrTableValues) = 2 Then
														aValue(iCount) = arrTableValues(0)
													End If
													bFlag=False
													If DictKeys(iCounter) = "Resources" Then
														sColName = Fn_GetXMLNodeValue(Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectRealNames\ServiceSchedular.xml" , "Resource")
													Else
														sColName = Fn_GetXMLNodeValue(Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectRealNames\ServiceSchedular.xml" , "Characteristics Name")
													End If

													bResult = -1
													For iRowCounter = 0 To JavaWindow("ServiceScheduler").JavaTable("PartRequests").GetROProperty("rows")-1
															If arrTableValues(0) = JavaWindow("ServiceScheduler").JavaTable("PartRequests").Object.getItem(iRowCounter).getData().getComponent().getProperty(sColName) Then
																bResult = iRowCounter
																Exit For
															End If
													Next
													 ' Check Condition if path returns false 
													If bResult = -1 Then Exit Function

													If UBound(arrTableValues) = 2 Then
															arrTableValues(1) = Fn_GetXMLNodeValue(Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectRealNames\ServiceSchedular.xml" , Trim(arrTableValues(1)))
                                                            sValue = JavaWindow("ServiceScheduler").JavaTable("PartRequests").Object.getItem(bResult).getData().getComponent().getProperty(arrTableValues(1))
															If  StrAction = "GetCellData" Then
																Fn_SISW_SrvScheduler_SummaryTabOperations = sValue
															ElseIf StrAction = "Verify" Then
																If Instr(1, sValue,arrTableValues(2)) = 0  Then
																		Exit Function
																End If
															End If	
													End If

													' Perform operation 
													If StrAction = "Select" Then
														JavaWindow("ServiceScheduler").JavaTable("PartRequests").SelectCell bResult,0
													ElseIf StrAction = "PopupMenuSelect" Then
														JavaWindow("ServiceScheduler").JavaTable("PartRequests").SelectCell bResult,0
														wait 1
														JavaWindow("ServiceScheduler").JavaTable("PartRequests").SelectColumnHeader 0,"RIGHT"
														wait 1

														aMenuList = split(StrMenu, ":",-1,1)
														intCount = Ubound(aMenuList)
														Select Case intCount
															Case "0"
																 StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
															Case "1"
																StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
															Case "2"
																StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
															Case Else
																Exit Function
														End Select
														JavaWindow("ServiceScheduler").WinMenu("ContextMenu").Select StrMenu
													End If
											Next
											bFlag=true
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
									Case "Notices"
											Call Fn_TabFolder_Operation("DoubleClickTab", "Summary", "")
											Call Fn_ReadyStatusSync(1)
											aValue=Split(DictItems(iCounter),"~")
											JavaWindow("ServiceScheduler").JavaStaticText("PropertyName").SetTOProperty "label", DictKeys(iCounter)
											For iCount = 0 To UBound(aValue)
													arrTableValues = Split(aValue(iCount), "$")
													bFlag=False
													sColName = Fn_GetXMLNodeValue(Environment.Value("sPath")+ "\TestData\AutomationXML\ObjectRealNames\ServiceSchedular.xml" , Trim(arrTableValues(1)))
													bResult = -1
													For iRowCounter = 0 To JavaWindow("ServiceScheduler").JavaTable("PartRequests").GetROProperty("rows")-1
															If arrTableValues(0) = JavaWindow("ServiceScheduler").JavaTable("PartRequests").Object.getItem(iRowCounter).getData().getComponent().getProperty(sColName) Then
																bResult = iRowCounter
																bFlag=True
																Exit For
															End If
													Next
													 ' Check Condition if path returns false 
													If bResult = -1 Then 
														Call Fn_TabFolder_Operation("DoubleClickTab", "Summary", "")
														Call Fn_ReadyStatusSync(1)	
														Exit Function
													End If
											Next	
											Call Fn_TabFolder_Operation("DoubleClickTab", "Summary", "")
											Call Fn_ReadyStatusSync(1)											
									'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  
									Case "Zone"
										Call Fn_TabFolder_Operation("DoubleClickTab", "Summary", "")
										Call Fn_ReadyStatusSync(1)
										aValue=Split(DictItems(iCounter),"~")
										JavaWindow("DefaultWindow").JavaTree("Notices").SetTOProperty "attached text","Zone"
										bResult = -1
										For iCount = 0 To UBound(aValue)
											For iRowCounter = 0 To cint(JavaWindow("DefaultWindow").JavaTree("Notices").Object.getitemcount()) - 1
												If cstr(JavaWindow("DefaultWindow").JavaTree("Notices").Object.getItem(iRowCounter).getData().toString()) = aValue(iCount) Then
													 bResult = iRowCounter
													 bFlag=True
												     Exit For
												End if	
											Next
										
										 ' Check Condition if path returns false 
											If bResult = -1 Then 
												Call Fn_TabFolder_Operation("DoubleClickTab", "Summary", "")
												Call Fn_ReadyStatusSync(1)	
												Exit Function
											End If
									Next	
									Call Fn_TabFolder_Operation("DoubleClickTab", "Summary", "")
									Call Fn_ReadyStatusSync(1)	
	
							   Case Else
										'Do Nothing
							End Select
					Next
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 					
             Case "EditBox"
					For iCounter=0 to dicSummaryInfo.count-1
							JavaWindow("ServiceScheduler").JavaStaticText("PropertyName").SetTOProperty "label", DictKeys(iCounter)&":"
							bFlag=False
							bResult = JavaWindow("ServiceScheduler").JavaEdit("SummaryPropValue").GetROProperty("value")
							 ' Check Condition if path returns false 
							If Trim(Cstr(bResult)) <> Trim(DictItems(iCounter)) Then Exit Function
							bFlag=true
					Next
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "ClickButton"
					JavaWindow("ServiceScheduler").JavaStaticText("PropertyName").SetTOProperty "label", DictKeys(0)
					'JavaWindow("ServiceScheduler").JavaToolbar("SummaryToolbar").Press DictItems(0)
					JavaWindow("ServiceScheduler").JavaButton("SummaryButton").SetTOProperty "label" ,DictItems(0)
					JavaWindow("ServiceScheduler").JavaButton("SummaryButton").Click
					If Err.Number < 0 Then
						Exit Function
					End If
					bFlag=true
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "ClickJavaObject"
					Set objDic=Description.Create()
					objDic("Class Name").value="JavaObject"
					objDic("toolkit class").value="org\.eclipse\.ui\.forms\.widgets\.ImageHyperlink"
					Set objChilds=JavaWindow("ServiceScheduler").JavaTab("SummaryTab").ChildObjects(objDic)

					For iCounter = 0 To objChilds.Count - 1
						JavaWindow("ServiceScheduler").JavaTab("SummaryTab").JavaObject("SummaryTabObject").SetTOProperty "index", iCounter
						sValue=JavaWindow("ServiceScheduler").JavaTab("SummaryTab").JavaObject("SummaryTabObject").Object.getText()
						If sValue = DictItems(0) Then
							Exit For
						End If
					Next

					JavaWindow("ServiceScheduler").JavaTab("SummaryTab").JavaObject("SummaryTabObject").Click 1,1,"LEFT"
					If Err.Number < 0 Then
						Exit Function
					End If
					bFlag=true
	End Select

	If bFlag=True Then
		Fn_SISW_SrvScheduler_SummaryTabOperations=True
	End If
End Function
''*********************************************************		Function to Perform Operations on Return Parts Dailog**************************************************************
'Function Name		:				Fn_SISW_SrvScheduler_ReturnPartsOperations

'Description			 :		 		 This function is used to perform operatons on Update configuration Dailog.

'Parameters			   :	 		 1. sAction - Specify the action
'												2. dicReturnParts - Dic Object
'											    3. sButton - Button to Click On

'Return Value		   : 				PASS \ FAIL

'Examples				:           	
'														Dim dicReturnParts
'														Set dicReturnParts=CreateObject("Scripting.Dictionary")
'
'														Case "Select"
'														dicReturnParts("Location")="PhysicalLocation193941"
'														dicReturnParts("Disposition")="In-Service"
'														dicReturnParts("ReturnDataTime")="23-Dec-2013 9:49"  ' Time here is converted to 24 hr if 9:49 pm them pass 21:49
'														dicReturnParts("SelectPhysicalPart")="001566/--A"
'														bReturn=Fn_SISW_SrvScheduler_ReturnPartsOperations("Select",dicReturnParts,"", "", "OK", "")

'														Case "PopupMenuSelect"
'														dicReturnParts("SelectPhysicalPart")="532361/--A"
'														bReturn=Fn_SISW_SrvScheduler_ReturnPartsOperations("PopupMenuSelect",dicReturnParts,"", "", "", "Properties")


'														Case "Verify"
'														dicReturnParts("SelectPhysicalPart")="001566/--A~002452/--A"
'														
'														For Verifying Objects only 
'														1. bReturn=Fn_SISW_SrvScheduler_ReturnPartsOperations("Verify",dicReturnParts,"", "", "", "")

'														For Verifying objects with other columns values as well
'														2. bReturn=Fn_SISW_SrvScheduler_ReturnPartsOperations("Verify",dicReturnParts,"Revision~Revision", "A~A", "", "")

'														Case "ManageColumn"
'														bReturn=Fn_SISW_SrvScheduler_ReturnPartsOperations("ManageColumn","","Revision~Lot", "", "", "")

'														Case "SortCoulmnAndVerify"  ' Click on column header to sort then verify sorted order
'														bReturn=Fn_SISW_SrvScheduler_ReturnPartsOperations("SortCoulmnAndVerify","","Object", "", "", "")

'History					 :			
'					Developer Name					Date									Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle					16-Jan-2014									1.0
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SrvScheduler_ReturnPartsOperations(sAction, dicReturnParts, sColName, sValues, sButtonName, StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_ReturnPartsOperations"
	Dim objReturnParts, bResult, aDate, aMenuList, intCount
	Dim iCounter, arrPhysicalParts, arrColumnName, arrValue, objColumnManage, iRowCounter
	Set objReturnParts = JavaWindow("ServiceScheduler").JavaWindow("ReturnParts")
	Fn_SISW_SrvScheduler_ReturnPartsOperations=False

	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations","Exist", objReturnParts,"") = False Then
		Exit Function
	End If

	Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
				If dicReturnParts("Location") <> "" Then
					bResult = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "Select", objReturnParts,"Location",dicReturnParts("Location"), "", "")
					If bResult = False Then
						Exit Function
					End If
				End If

				If dicReturnParts("Disposition") <> "" Then
					bResult = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "Select", objReturnParts,"Disposition",dicReturnParts("Disposition"), "", "")
					If bResult = False Then
						Exit Function
					End If
				End If

				If dicReturnParts("ReturnDataTime") <> "" Then
					'Date Panel
					aDate = Split( trim(dicReturnParts("ReturnDataTime")) , "~" )
					objReturnParts.JavaEdit("Return Date/Time:").Set aDate(0)
					wait 1
					call Fn_KeyBoardOperation("SendKeys", "{TAB}")
					objReturnParts.JavaList("Time").Type aDate(1)
					if objReturnParts.JavaEdit("Return Date/Time:").getROProperty ("value") = aDate(0) and instr(aDate(1),objReturnParts.JavaList("Time").getROProperty ("value")) then
						bResult = True
					else
						bResult = False
					End If
					If bResult = False Then Exit Function
				End If

				If dicReturnParts("SelectPhysicalPart") <> "" Then
					bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "GetRowIndex", objReturnParts, "SelectPhysicalPart", "GetCellData", "Object", dicReturnParts("SelectPhysicalPart"), "", "", "", ":")
					 ' Check Condition if path returns false 
					If bResult = -1 Then Exit Function
					objReturnParts.JavaTable("SelectPhysicalPart").SelectCell bResult,"Object"
				End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PopupMenuSelect"
				If dicReturnParts("SelectPhysicalPart") <> "" Then
					bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "GetRowIndex", objReturnParts, "SelectPhysicalPart", "GetCellData", "Object", dicReturnParts("SelectPhysicalPart"), "", "", "", ":")
					 ' Check Condition if path returns false 
					If bResult = -1 Then Exit Function
					objReturnParts.JavaTable("SelectPhysicalPart").SelectCell bResult,"Object"
					wait 1
					objReturnParts.JavaTable("SelectPhysicalPart").SelectColumnHeader "Object","RIGHT"
					wait 1
	
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Select Case intCount
						Case "0"
							 StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu =JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Exit Function
					End Select
					JavaWindow("ServiceScheduler").WinMenu("ContextMenu").Select StrMenu
				Else
					Exit Function
				End If
		'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Verify"
                arrPhysicalParts = Split(dicReturnParts("SelectPhysicalPart"), "~")
				arrColumnName = Split(sColName, "~")
				arrValue = Split(sValues, "~")
				For iCounter = 0 To UBound(arrPhysicalParts)
					bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "GetRowIndex", objReturnParts, "SelectPhysicalPart", "GetCellData", "Object", arrPhysicalParts(iCounter), "", "", "", ":")
					 ' Check Condition if path returns false 
					If bResult = -1 Then Exit Function
					If sColName <> "" Then
							bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "VerifyCellData", objReturnParts, "SelectPhysicalPart", "GetCellData", "Object", bResult, arrColumnName(iCounter), arrValue(iCounter), "", "")
							' Check Condition if path returns false 
							If bResult = False Then 
								Exit Function
							End If
					End If
				Next
		'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "ManageColumn"
				arrColumnName = Split(sColName, "~")
				objReturnParts.JavaToolbar("ManageColumn").Press "Manage Columns..."
                If Err.Number < 0 Then
					Exit Function
				End If
				wait 1
			         Set 	objColumnManage = objReturnParts.JavaWindow("ColumnManagement")
				If objColumnManage.Exist(1) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "ColumnManagement Dialog does not exist")
					Exit Function
				End If

				For iCounter = 0 To UBound(arrColumnName)
							bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "GetRowIndex", objColumnManage, "DisplayedColumns", "GetCellData", "Property", arrColumnName(iCounter), "", "", "", ":")
							' Check Column Exist in DisplayedColumns Table
							If bResult = -1 Then 
									bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "GetRowIndex", objColumnManage, "AvailableProperties", "GetCellData", "Property", arrColumnName(iCounter), "", "", "", ":")
									' Check Column path returns False from  AvailableProperties Table
									If bResult = -1 Then Exit Function
		
									objColumnManage.JavaTable("AvailableProperties").DoubleClickCell bResult,"Property"
		
									' Check Double Clicking on Column from AvailableProperties table work correctly
									If Err.Number < 0 Then
											Exit Function
									End If
							End If
						Next
				Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "Click", objColumnManage,"Apply")
				wait 2
				If objColumnManage.Exist(1) Then
					Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "Click", objColumnManage,"Close")
				End If
		'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "SortCoulmnAndVerify"
				objReturnParts.JavaTable("SelectPhysicalPart").SelectColumnHeader sColName
				' Check Clicking on Column Header worked correctly
				If Err.Number < 0 Then
						Exit Function
				End If

				iRowCounter = objReturnParts.JavaTable("SelectPhysicalPart").GetROProperty("rows")
				For iCounter = 1 To iRowCounter - 1
					sValue = objReturnParts.JavaTable("SelectPhysicalPart").GetCellData(iCounter-1, sColName)
					If sValue > objReturnParts.JavaTable("SelectPhysicalPart").GetCellData(iCounter, sColName) Then
						Exit Function
					End If
				Next

		Case Else
				Exit function
	End Select

	If sButtonName <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_ReturnPartsOperations", "Click", objReturnParts,sButtonName)
	End If
	Set objReturnParts = nothing
	Fn_SISW_SrvScheduler_ReturnPartsOperations = True
End Function

'*********************************************************		Function to Create Discrepancy Type in Service Scheduler  ***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_CreateDiscrepancyType

'Description			 :		 		  This function is used to Create Discrepancy Type in Service Scheduler.

'Parameters			   :	 			1.  sAction
'													 2. sButtonName: Name of Button to Click
'													 3. dicInputs: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		Create Discrepancy Type Window should be visible

'Examples				:              
'											Case "Create"

'											set dicInputs = CreateObject( "Scripting.Dictionary" )
'											dicInputs("Name")="Descripancy1234000"
'											dicInputs("ID")="12340000"
'											dicInputs("ActivityNumber")="00070000"
'											dicInputs("Description")="Descripancy"
'											dicInputs("Severity")="Major"
'											dicInputs("DiscoveryDate")="30-1-2014~12:54:55 PM"
'											bReturn=Fn_SISW_SrvScheduler_CreateDiscrepancyType("Create" , "Finish" , dicInputs)

'History:
'	Developer Name			Date				Rev. No.				Changes Done																	Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Tewari			30-Jan-2014			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_CreateDiscrepancyType(sAction , sButtonName , dicInputs, strReserved)

GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_CreateDiscrepancyType"
Dim objDialog, dicItem,dicValue,dicAction,dicObjInputs
Dim aDate
	Set objDialog=Fn_SISW_SrvScheduler_GetObject("CreateDiscrepancy")

	Fn_SISW_SrvScheduler_CreateDiscrepancyType=False
	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_CreateDiscrepancyType","Exist", objDialog,"") = False Then
		Exit Function
		Call Fn_ReadyStatusSync( 3 )
	End If

	Select Case sAction
		Case "Create"
				dicItem = dicInputs.Keys
				dicValue = dicInputs.Items
				For iCount = 0 to dicInputs.Count - 1
					Select Case trim(dicItem(iCount))
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Name","ID","ActivityNumber","Description","DiscoveredBy"
							'Edit Box
							bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_CreateDiscrepancyType", "Set",  objDialog,trim(dicItem(iCount)), trim(dicValue(iCount)) )
							If bResult = False Then Exit Function
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------						
						 Case "Initiation Date","Discovery Date","Due Date"
							objDialog.JavaStaticText("PropertyName").SetTOProperty "label", dicitem(iCount)&":"
							aDate = Split( trim(dicValue(iCount)) , "~" )
							objDialog.JavaEdit("DiscoveryDate").Set aDate(0)
							wait 1
							call Fn_KeyBoardOperation("SendKeys", "{TAB}")
							objDialog.JavaList("Time").Type aDate(1)
						     If Err.Number <  0 Then
								 Fn_SISW_SrvScheduler_CreateDiscrepancyType = FALSE				 
								Exit Function
							End If
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						 Case "Severity"
							bResult=Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_CreateDiscrepancyType", "Select", objDialog,trim(dicItem(iCount)),trim(dicValue(iCount)), "", "")
							If bResult = False Then Exit Function
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Physical Part In Progress"
							objDialog.JavaStaticText("PropertyName").SetTOProperty "label",trim(dicItem(iCount))&":"
							Set dicObjInputs=dicInputs("Physical Part In Progress")
							dicAction=dicObjInputs("sAction")
							Call Fn_UI_JavaStaticText_Click(" Fn_SISW_SrvScheduler_CreateDiscrepancyType", objDialog, "CommonDownArrow", 1, 1, "LEFT")
							Wait 2
							bResult=Fn_UI_JavaMenu_Select("",objDialog,dicAction)
							If bResult = False Then Exit Function
							Wait 2
							If dicAction="Add..." Then
								bResult=Fn_SISW_SrvScheduler_SearchOperations("SearchAndSelect", dicObjInputs)	
								If bResult = False Then Exit Function
							End If
							'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Fault Code"
							objDialog.JavaStaticText("PropertyName").SetTOProperty "label",trim(dicItem(iCount))&":"
							Set dicObjInputs=dicInputs("Fault Code")
							dicAction=dicObjInputs("sAction")
							Call Fn_UI_JavaStaticText_Click(" Fn_SISW_SrvScheduler_CreateDiscrepancyType", objDialog, "CommonDownArrow", 1, 1, "LEFT")
							Wait 2
							bResult=Fn_UI_JavaMenu_Select("",objDialog,dicAction)
							If bResult = False Then Exit Function
							Wait 2
							If dicAction="Add..." Then
								bResult=Fn_SISW_SrvScheduler_SearchOperations("SearchAndSelect", dicObjInputs)	
								If bResult = False Then Exit Function
							End If

					   End Select
				Next

		Case Else
			Exit Function
	End Select
    
	If sButtonName <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_CreateDiscrepancyType", "Click", objDialog,sButtonName)
	End If
	Fn_SISW_SrvScheduler_CreateDiscrepancyType=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_CreateDiscrepancyType:Successfully created Discrepancy Type")
	Set objDialog = Nothing
End Function

''*********************************************************		Function to Select Physical Asset Object from Structure Scheduler**************************************************************
'Function Name		:				Fn_SISW_SrvScheduler_SelectPhysicalAsset.

'Description			 :		 		 This function is used to delete the Object Component from Schedule Manager.

'Parameters			   :	 			1. sAction - Action To Perform
'												2. sAssets: Assets names
'												3. sButton - Button to Click On
											
'Return Value		   : 				PASS \ FAIL

'Examples				:				 Fn_SISW_SrvScheduler_SelectPhysicalAsset("Select", "000124/--A", "OK")

' History:
'		Developer Name					Date						Rev. No.			Changes Done									Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle 						9-Dec-2013				1.0																						Sunny			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_SelectPhysicalAsset(sAction, sAssets, sButton)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_SelectPhysicalAsset"
Dim objPhysicalAsset, arrAssets, iCounter, iRowCounter, sValue
Set objPhysicalAsset = JavaWindow("ServiceScheduler").JavaWindow("SelectPhysicalAsset")
Fn_SISW_SrvScheduler_SelectPhysicalAsset = FALSE
	
arrAssets= Split(sAssets, "~")

	Select Case sAction
		Case "Select", "Verify"
			For iCounter = 0 To UBound(arrAssets)
				bResult = False
				iRowCounter = objPhysicalAsset.JavaTable("Assets").GetROProperty("rows")
				For iCount = 0 To iRowCounter - 1 
					sValue= objPhysicalAsset.JavaTable("Assets").GetCellData(iCount,"Object")
					If  sValue= arrAssets(iCounter) Then
						bResult = True
						If sAction = "Select" Then
							objPhysicalAsset.JavaTable("Assets").ActivateRow iCount
						End If
						Exit For
					End If
				Next
				If bResult = False Then
					Exit Function
				End If
			Next
		Case Else
				Exit function
	End Select


	If sButton <> "" Then
		Call Fn_Button_Click("Fn_SISW_SrvScheduler_SelectPhysicalAsset", objPhysicalAsset, sButton)
	End If
	Set objPhysicalAsset = nothing
	Fn_SISW_SrvScheduler_SelectPhysicalAsset = True
End Function
''*********************************************************		Function to Select Object from Fn_SISW_SrvScheduler_Maintenance Tree **************************************************************
'Function Name		:				Fn_SISW_SrvScheduler_MaintenanceTree.

'Description			 :		 		 This function is used to delete the Object Component from Schedule Manager.

'Parameters			   :	 			1. sAction - Action To Perform
'												2. sAssets: Assets names
'												3. sButton - Button to Click On
											
'Return Value		   : 				PASS \ FAIL

'Examples				:				 Fn_SISW_SrvScheduler_MaintenanceTree("Select", "000123/--A (View):000124/--A", "OK")

' History:
'		Developer Name					Date						Rev. No.			Changes Done									Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle 						9-Dec-2013				1.0																						Sunny			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_MaintenanceTree(sAction, sAssets, sButton)

GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_MaintenanceTree"
Dim objPhysicalAsset, arrAssets, iCounter, iRowCounter, sValue
Set objPhysicalAsset = JavaWindow("ServiceScheduler").JavaWindow("MaintenanceTree")
Fn_SISW_SrvScheduler_MaintenanceTree = FALSE

arrAssets= Split(sAssets, "~")

	Select Case sAction
		' ************************************************************Case To delete Item from Menu option *****************************************************************
		Case "Select"
			JavaWindow("ServiceScheduler").JavaWindow("MaintenanceTree").JavaTree("Tree").Select sAssets
			If Err.Number < 0 Then
				Exit Function
			End If

		Case "Verify"
			iRowCounter = objPhysicalAsset.JavaTree("Tree").GetROProperty("items count")
            For iCounter = 0 To UBound(arrAssets)
				bResult = False
				For iCount = 0 To iRowCounter - 1 
					sValue= objPhysicalAsset.JavaTree("Tree").GetItem(iCount)
					If  sValue= arrAssets(iCounter) Then
						bResult = True
						Exit For
					End If
				Next
				If bResult = False Then
					Exit Function
				End If
			Next
		Case Else
				Exit function
	End Select

	If sButton <> "" Then
		Call Fn_Button_Click("Fn_SISW_SrvScheduler_MaintenanceTree", objPhysicalAsset, sButton)
	End If

	Set objPhysicalAsset = nothing
	Fn_SISW_SrvScheduler_MaintenanceTree = True
End Function

''*********************************************************		Function to handle the Part Request Operations from Service Scheduler	**************************************************************
'Function Name		:				Fn_SISW_SrvScheduler_PartRequestOperations.

'Description			 :		 		 This function is used to delete the Object Component from Schedule Manager.

'Parameters			   :	 			1. sAction - Specify the action(i.e.Menu,ShortKey,Toolbar)
'												   2. sMessage - message to Verify
'												   3. sButton - Button to Click On
											
'Return Value		   : 				PASS \ FAIL

'Examples				:				 Call Fn_SISW_SrvScheduler_PartRequestOperations("CancelPart","Delete the selected task(s)?","Yes")

' History:
'		Developer Name					Date						Rev. No.			Changes Done									Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle 						25-Dec-2013				1.0																						Sunny			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_PartRequestOperations(sAction,sMessage,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_PartRequestOperations"
Dim objPartReq, sMsg, sTitle
Set objPartReq = Dialog("CancelPartRequestWarning")

Fn_SISW_SrvScheduler_PartRequestOperations = FALSE

	Select Case sAction
		Case "CancelPart"  
				sTitle = "Cancel Part Request Warning Dialog"
		Case "ClosePart" 
				sTitle = "Close Part Request Warning Dialog"
		Case "ValidatePart"
				sTitle = "Part Request validation error dialog"
	End Select

	''***********************************************************'' To check whether part Request Dialog  is displayed*************************************************************
	If  sTitle <> "" Then
		objPartReq.SetTOProperty "text", sTitle
	End If
	
	If objPartReq.Exist(2) Then
		If  sMessage <> "" Then
			sMsg = objPartReq.Static("ErrMsg").GetROProperty("text")
			If sMsg <> sMessage Then
				Call Fn_Button_Click("Fn_SISW_SrvScheduler_PartRequestOperations", objPartReq,"No")
				Exit Function
			End If
		End If
	Else
		Exit Function
	End If

	If sButton <> "" Then
		Call Fn_UI_WinButton_Click("Fn_SISW_SrvScheduler_PartRequestOperations",objPartReq,sButton,5,5,micLeftBtn)
	End If

	Set objPartReq = nothing
	Fn_SISW_SrvScheduler_PartRequestOperations = True
End Function

''*********************************************************		Function to Perform Operations on Update Configuration Dailog**************************************************************
'Function Name		:				Fn_SISW_SrvScheduler_UpdateConfigurationOperations

'Description			 :		 		 This function is used to perform operatons on Update configuration Dailog.

'Parameters			   :	 		1. sAction - Specify the action
'												2. sObject - Object Column Value
'												3. dicUpdateConfig - Dic Object
'											  4. sButton - Button to Click On

'Return Value		   : 				PASS \ FAIL

'Examples				:           	
'														Dim dicUpdateConfig
'														Set dicUpdateConfig=CreateObject("Scripting.Dictionary")
'				
'														Case "Select"
'														bReturn=Fn_SISW_SrvScheduler_UpdateConfigurationOperations("Select","Part_Mov81967","", "OK")

'														Case "Verify"
'														dicUpdateConfig("Part Movement Type")="Replace"
'														dicUpdateConfig("Action Date")="23-Dec-2013 9:49"  ' Time here is converted to 24 hr if 9:49 pm them pass 21:49
'														dicUpdateConfig("Parent Physical Element")="000689"
'														bReturn=Fn_SISW_SrvScheduler_UpdateConfigurationOperations("Verify","Part_Mov81967",dicUpdateConfig, "")

'														Case "GetCellData"             -->  Return Column Values with ~ Separator
'														dicUpdateConfig("Part Movement Type")=""
'														dicUpdateConfig("Action Date")=""
'														dicUpdateConfig("Parent Physical Element")=""
'														bReturn=Fn_SISW_SrvScheduler_UpdateConfigurationOperations("Verify","Part_Mov81967",dicUpdateConfig, "")

'History					 :			
'					Developer Name					Date									Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Ankit Tewari					09-Dec-2013									1.0																				
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle					26-Dec-2013									1.1					Added Case "Verify"																				
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'					Pranav Ingle					26-Dec-2013									1.1					Added Case "GetCellData"																				
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SrvScheduler_UpdateConfigurationOperations(sAction,sObject, dicUpdateConfig, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_UpdateConfigurationOperations"
	Dim objUpdateConfig, arrAssets, iCounter, iRowCounter, sValue, iCount, sDate, aDate
	Dim DictKeys, DictItem
	Set objUpdateConfig = Fn_SISW_SrvScheduler_GetObject("UpdateConfiguration")
	Fn_SISW_SrvScheduler_UpdateConfigurationOperations=False
	bResult = False

	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_UpdateConfigurationOperations","Exist", objUpdateConfig,"") = False Then
		Exit Function
	End If

	If sAction <> "SetRebaseDateAndTime" Then
		iRowCounter = objUpdateConfig.JavaTable("PartMovements").GetROProperty("rows")
		For iCount = 0 To iRowCounter - 1 
			sValue= objUpdateConfig.JavaTable("PartMovements").GetCellData(iCount,"Object")
			If  sValue= sObject Then
				iRowIndex = iCount
				Exit For
			End If
		Next
	End If

	Select Case sAction
		Case "Select"
				objUpdateConfig.JavaTable("PartMovements").ActivateRow iRowIndex
				bResult = True
		'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Verify"		
				'Taking Items & Keys from dictionary
				DictKeys = dicUpdateConfig.Keys
				DictItem = dicUpdateConfig.Items
				For iCounter = 0 To dicUpdateConfig.Count -1
						bResult = False
						sValue = JavaWindow("ServiceScheduler").JavaWindow("UpdateConfiguration").JavaTable("PartMovements").GetCellData(iRowIndex, DictKeys(iCounter))
						If  IsNumeric(sValue) Then
							If cInt(sValue) = CInt(DictItem(iCounter)) Then
								bResult = True
							End If
						ElseIf IsDate(sValue) Then
							'sDate = split(DictItem(iCounter),"/")
							'aDate = Split(sDate(2)," ")
							'sDate(2) = aDate(0)
							'DictItem(iCounter) = sDate(1)+"-"+MonthName( sDate(0),True)+"-"+sDate(2)+" "+FormatDateTime(Cdate(DictItem(iCounter)),4)
							
							sValue = FormatDateTime(Cdate(sValue),0)
							DictItem(iCounter) = FormatDateTime(Cdate(DictItem(iCounter)),0)
							
							If cDate(DateValue(sValue)) = CDate(DateValue(DictItem(iCounter))) AND CDate(FormatDateTime(Cdate(sValue),4)) = CDate(FormatDateTime(Cdate(DictItem(iCounter)),4)) Then
								bResult = True
							End If
						Else
							If sValue = DictItem(iCounter) Then
								bResult = True
							End If
						End If

						' Exit Fucntion if Column value does not match with expected value for Part Movement
						If bResult = False Then
							Exit Function
						End If
				Next
		'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "GetCellData"		
				'Taking Items & Keys from dictionary
				DictKeys = dicUpdateConfig.Keys
				DictItem = dicUpdateConfig.Items
				sValue=""
				For iCounter = 0 To dicUpdateConfig.Count -1
						bResult = False
						sValue = sValue&"~"&JavaWindow("ServiceScheduler").JavaWindow("UpdateConfiguration").JavaTable("PartMovements").GetCellData(iRowIndex, DictKeys(iCounter))
				Next
				Fn_SISW_SrvScheduler_UpdateConfigurationOperations = sValue
		'-----------------------------------------------------------------------------------------------------------------------------------------------
		'	 Added Case by Poonam C  - to Set Rebase Date and Time in Update Configuration dialog
		Case "SetRebaseDateAndTime"	
				If dicUpdateConfig("RebaseDate") <> "" Then
					bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_UpdateConfigurationOperations","Set",objUpdateConfig,"RebaseDate",dicUpdateConfig("RebaseDate"))
				End If

				If dicUpdateConfig("Time") <> "" Then
					bResult = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_UpdateConfigurationOperations","Select",objUpdateConfig,"Time",dicUpdateConfig("Time"),"","")
				End If

		Case Else
				Exit function
	End Select

	If sButtonName <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_UpdateConfigurationOperations", "Click", objUpdateConfig,sButtonName)
	End If
	Set objUpdateConfig = nothing
	If bResult = True Then
		Fn_SISW_SrvScheduler_UpdateConfigurationOperations = True
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SrvScheduler_SearchNoticeOperations

'Description			 :	Function Used to Perform operation on Search Notice dialog

'Parameters			   :  1.sAction : Action name
'									2.sInvokeOption: Dialog Invoke option
'								 	3.dictSearchInfo: Search information
'								    4.sButton: Button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be login to Service Scheduler perspective

'Examples				:   Dim dictSearchInfo
'							Set dictSearchInfo=CreateObject("Scripting.Dictionary")
'							dictSearchInfo("Name")="Notice1"
'							dictSearchInfo("Object")="Notice1"
'							bReturn=Fn_SISW_SrvScheduler_SearchNoticeOperations("SearchAndRelate","toolbar",dictSearchInfo,"")
'
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Reema W								15-May-2014					1.0																				Paresh D
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_SrvScheduler_SearchNoticeOperations(sAction,sInvokeOption,dictSearchInfo,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_SearchNoticeOperations"
    'Declaring variables
	Dim objSearchNotice
	Dim iCounter
	Dim bFlag

	Fn_SISW_SrvScheduler_SearchNoticeOperations=False
	'Invoke Search Notice dialog
	Select Case Lcase(sInvokeOption)
		Case "toolbar"
			If Fn_ToolbarOperation("Click", "Assign Notice...","")=False Then
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_SearchNoticeOperations ] fail to click on toolbar button  [ Assign Notice... ]")
				Exit Function
			End If
            Call Fn_ReadyStatusSync(2)
        Case "button"
			bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_SearchNoticeOperations", "Object.click", sButton,"")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_SearchNoticeOperations ] fail to click on table button  [ Assign Notice... ]")
				Exit Function
			End If
			Call Fn_ReadyStatusSync(1)
		Case "nooption"
			'Use this option when user wants to invoke Search Notice dialog outside of function
	End Select
	'Checking existance of  [ Search Notice ] dialog
	If JavaWindow("ServiceScheduler").JavaWindow("SearchNotice").Exist(20) Then
		'creating object of  [ Search Notice ] dialog
		Set objSearchNotice=JavaWindow("ServiceScheduler").JavaWindow("SearchNotice")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_SearchNoticeOperations ] fail to invoke  [ Search Notice ] dialog")
		Exit Function
	End If

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Action to search notice and relate
		Case "SearchAndRelate"
			'Enter name
			If dictSearchInfo("Name")<>"" Then
				If Fn_Edit_Box("Fn_SISW_SrvScheduler_SearchNoticeOperations", objSearchNotice,"Name", dictSearchInfo("Name") )=False Then
					Set objSearchNotice = Nothing
					Exit Function
				End If
			End If
			'Enter Notice Type
			If dictSearchInfo("NoticeType")<>"" Then
				If Fn_Edit_Box("Fn_SISW_SrvScheduler_SearchNoticeOperations", objSearchNotice,"NoticeType", dictSearchInfo("NoticeType") )=False Then
					Set objSearchNotice = Nothing
					Exit Function
				End If
			End If
			'Enter Description
			If dictSearchInfo("Description")<>"" Then
				If Fn_Edit_Box("Fn_SISW_SrvScheduler_SearchNoticeOperations", objSearchNotice,"Description", dictSearchInfo("Description") )=False Then
					Set objSearchNotice = Nothing
					Exit Function
				End If
			End If
			'Enter Owning User
			If dictSearchInfo("OwningUser")<>"" Then
				If Fn_Edit_Box("Fn_SISW_SrvScheduler_SearchNoticeOperations", objSearchNotice,"OwningUser", dictSearchInfo("OwningUser") )=False Then
					Set objSearchNotice = Nothing
					Exit Function
				End If
			End If
			'Enter Owning Group
			If dictSearchInfo("OwningGroup")<>"" Then
				If Fn_Edit_Box("Fn_SISW_SrvScheduler_SearchNoticeOperations", objSearchNotice,"OwningGroup", dictSearchInfo("OwningGroup") )=False Then
					Set objSearchNotice = Nothing
					Exit Function
				End If
			End If
			'Clicking on Find button
			If Fn_Button_Click("Fn_SISW_SrvScheduler_SearchNoticeOperations", objSearchNotice, "Find")=False Then
				Set objSearchNotice = Nothing
				Exit Function
			End If
			Call Fn_ReadyStatusSync(5)
			'Checking existance of [ Search Results ] table
			If objSearchNotice.JavaTable("SearchResults").Exist(6) Then
				If dictSearchInfo("Object")="" Then
					dictSearchInfo("Object")=dictSearchInfo("Name")
				End If
				bFlag=False
				For iCounter=0 to objSearchNotice.JavaTable("SearchResults").GetROProperty("rows")-1
                    If Trim(dictSearchInfo("Object"))=Trim(objSearchNotice.JavaTable("SearchResults").GetCellData(iCounter,"Object")) Then
						objSearchNotice.JavaTable("SearchResults").SelectCell iCounter,"Object"
						wait 1
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_SearchNoticeOperations ] : Object [ " & Cstr(dictSearchInfo("Object")) & "] not found in Search Results")
					Set objSearchNotice = Nothing
					Exit Function
				End If
				'Clicking on Relate button
				If Fn_Button_Click("Fn_SISW_SrvScheduler_SearchNoticeOperations", objSearchNotice, "Relate")=False Then
					Set objSearchNotice = Nothing
					Exit Function
				End If
				Call Fn_ReadyStatusSync(5)
				'Function returns value
				If Err.Number<0 Then
					Fn_SISW_SrvScheduler_SearchNoticeOperations=False
				Else
					Fn_SISW_SrvScheduler_SearchNoticeOperations=True
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_SearchNoticeOperations ] : Search Results table doen not exist")
				Set objSearchNotice = Nothing
				Exit Function
			End If
	End Select
	'Releasing object of Search Notice dialog
	Set objSearchNotice = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SrvScheduler_ViewDefineCharToleOperation

'Description			 :	Function Used to Perform operation on Search  Characteristic dialog

'Parameters			   :  1.sAction : Action name
'									2.sInvokeOption: Dialog Invoke option
'								 	3.CharacteristicName: Characteristic Name to be  Verified
'								    4.MaxValue: Maximum Value to be set or Verified
'								    5.MinValue: Minimum Value to be set or Verified
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be login to Service Scheduler perspective

'Examples				:  									
'									msgbox  Fn_SISW_SrvScheduler_ViewDefineCharToleOperation("Set", "toolbar", "", "123", "10")
'									
'									msgbox  Fn_SISW_SrvScheduler_ViewDefineCharToleOperation("Verify", "toolbar", "char1", "123", "10")
'
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Ganesh B									03-Jun-2014					1.0																				
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function  Fn_SISW_SrvScheduler_ViewDefineCharToleOperation(sAction, sInvokeOption, CharacteristicName, MaxValue, MinValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_ViewDefineCharToleOperation"
   Dim sValue, bReturn
   Dim objCharacteristicTolerance
	Set objCharacteristicTolerance = JavaWindow("DefaultWindow").JavaWindow("ViewDefineCharacteristicTolerance")
	Fn_SISW_SrvScheduler_ViewDefineCharToleOperation =False
	If Not objCharacteristicTolerance.Exist(1) Then
		Select Case Lcase(sInvokeOption)
			Case "toolbar"
				If 	 Fn_SISW_UI_JavaToolbar_Operations("", "ClickExt", JavaWindow("DefaultWindow"), "", "View/Define Characteristic Tolerance...", "", "", 1) = False Then
					Fn_SISW_SrvScheduler_ViewDefineCharToleOperation =False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_SearchCharacteristicOperation ] fail to click on toolbar button  [ Assign Characteristic... ]")
					Set objSearchDialog = Nothing
					Exit Function
				End If
			Case "nooption"
				'Use this option when user wants to invoke Search Characteristic dialog outside of function
		End Select
	End If
	
	objCharacteristicTolerance.Resize 500, 300
	Select Case sAction
	 Case "Verify"
		 	If 	CharacteristicName <> "" Then
			sValue = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_ViewDefineCharToleOperation", "GetText", objCharacteristicTolerance, "CharacteristicsName", "")
				If strComp(sValue,CharacteristicName) = 0  Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified Characteristic Name Value")	
				Else					
					Fn_SISW_SrvScheduler_ViewDefineCharToleOperation =False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Verify Characteristic Name Value.")
						Set objSearchDialog = Nothing
						Exit Function 
				End If
			End If
			'Commented for verifying the Empty Value
			'If 	MaxValue <> "" Then
			sValue = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_ViewDefineCharToleOperation", "GetText", objCharacteristicTolerance, "MaximumValue", "")
				If strComp(sValue,MaxValue) = 0  Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified Maximum Value")	
				Else					
					Fn_SISW_SrvScheduler_ViewDefineCharToleOperation =False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Verify Maximum Value.")
						Set objSearchDialog = Nothing
						Exit Function 
				End If
			'End If
			'If 	MinValue <> "" Then
				sValue = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_ViewDefineCharToleOperation", "GetText", objCharacteristicTolerance, "MinimumValue", "")
				If strComp(sValue,MinValue) = 0  Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified Minimum Value")	
				Else					
					Fn_SISW_SrvScheduler_ViewDefineCharToleOperation =False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Verify Minimum Value.")
						Set objSearchDialog = Nothing
						Exit Function 
				End If
			'End If
			call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Click", objCharacteristicTolerance, "Cancel")
			If Err.number < 0 Then
				Fn_SISW_SrvScheduler_ViewDefineCharToleOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Cancel] Button")
				Set objSearchDialog = Nothing
				Exit Function
			Else
				Fn_SISW_SrvScheduler_ViewDefineCharToleOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on  [Cancel] Button")	
			End If
    	 Case "Set"
		 	If 	CharacteristicName <> "" Then
			bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_ViewDefineCharToleOperation", "Set", objCharacteristicTolerance, "CharacteristicsName", CharacteristicName)
				If bReturn = True  Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set  Characteristic Name as [" + 	CharacteristicName+ "]")	
				Else					
					Fn_SISW_SrvScheduler_ViewDefineCharToleOperation =False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Characteristic Name as [" + CharacteristicName + "]")
						Set objSearchDialog = Nothing
						Exit Function 
				End If
			End If
			'Commented for Set the Empty Value
			'If 	MaxValue <> "" Then
			bReturn  = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_ViewDefineCharToleOperation", "Set", objCharacteristicTolerance, "MaximumValue", MaxValue)
				If bReturn = True  Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set   Maximum Value as [" + 	MaxValue+ "]")	
				Else					
					Fn_SISW_SrvScheduler_ViewDefineCharToleOperation =False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Maximum Value as [" + MaxValue + "]")
						Set objSearchDialog = Nothing
						Exit Function 
				End If
			'End If
			'If 	MinValue <> "" Then
				bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_ViewDefineCharToleOperation", "Set", objCharacteristicTolerance, "MinimumValue", MinValue)
				If bReturn = True  Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set  Minimum Value as [" + 	MinValue+ "]")	
				Else					
					Fn_SISW_SrvScheduler_ViewDefineCharToleOperation =False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Minimum Value as [" + MinValue + "]")
						Set objSearchDialog = Nothing
						Exit Function 
				End If
			'End If
			call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Click", objCharacteristicTolerance, "OK")
			If Err.number < 0 Then
				Fn_SISW_SrvScheduler_ViewDefineCharToleOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [OK] Button")
				Set objSearchDialog = Nothing
				Exit Function
			Else
				Fn_SISW_SrvScheduler_ViewDefineCharToleOperation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on  [OK] Button")	
			End If
	End Select
End Function

'****************************************    Function to perform Find and Select Physical Part ***************************************
'
''Function Name		 	:	Fn_SISW_SrvScheduler_IssueParts
'
''Description		    :  	Function to perform  assign lot operation

''Parameters		    :	1. sAction : Action need to perform
			'					   		2. sSearchType : to open Lot dialog ( AssignLot_RMB / AssignLot_Menu )
			'					   		3. dicIssueParts
								
''Return Value		    :  	True \ False
'
''Pre-requisite		    :	Search Window should be present.

''Examples		     	:	
'							Dim dicIssueParts
'							Set dicIssueParts = CreateObject( "Scripting.Dictionary" )
							'dicIssueParts("bClear") = True
							'dicIssueParts("Serial Number") 
							'dicIssueParts("Serial Number After")
							'dicIssueParts("Serial Number Before")
							'dicIssueParts("Lot Number")
							'dicIssueParts("Manufacturer's ID")
							'dicIssueParts("Manufacturered After")
							'dicIssueParts("Manufacturered Before")
							'dicIssueParts("Location Name")
							'dicIssueParts("Disposition Value")

'Examples		     	:	Case "Preferred Parts"
'								dicIssueParts("SelectPhysicalPart")
'								Fn_SISW_SrvScheduler_IssueParts("FindAndSelect", "Preferred Parts", dicIssueParts)
'-----------------------------------------------------------------------------------------------------------------------------------
'								Call Fn_SISW_SrvScheduler_IssueParts("CloseDialog", "", "")
'-----------------------------------------------------------------------------------------------------------------------------------

'History:
'	Developer Name				Date					Rev. No.			Reviewer			Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle		 			16-Jan-2014				1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_IssueParts(sAction, sSearchType, dicIssueParts, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_IssueParts"
	Dim objIssueParts, sTitle, iCnt, bFlag, iRowCount, aSearchCri , aFieldSet
	Dim DictItems, DictKeys, iCounter
	Set objShell = JavaWindow("ServiceSchedulerShell")
	Set objIssueParts = objShell.JavaWindow("IssueParts")
	bFlag = False
	Fn_SISW_SrvScheduler_IssueParts = False

	'Check Existance of  Issue Part Dialog
	For iCnt = 0 To 50
		objShell.SetTOProperty "Index", iCnt
		If objIssueParts.Exist(1) Then
			bFlag = True
			Exit for
		End If
	Next

	If Not(bFlag) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_IssueParts ] Failed to find [ " & sTitle & " ] window.")
		Exit function
	End If

	Select Case sAction
		Case "CloseDialog"
				Call Fn_Button_Click("Fn_SISW_SrvScheduler_IssueParts", objIssueParts, "Cancel")
				Fn_SISW_SrvScheduler_IssueParts = True

		Case "FindAndSelect"
					' clearing fields
					If dicIssueParts("bClear") <> "" Then
						If cBool(dicIssueParts("bClear")) Then
							Call Fn_Button_Click("Fn_SISW_SrvScheduler_IssueParts", objIssueParts, "Clear")
						End If
					End If
		
					'Taking Items & Keys from dictionary
					DictItems = dicIssueParts.Items
					DictKeys = dicIssueParts.Keys

					' Enter Details for Search Operation
					For iCounter=0 to dicIssueParts.count-1
						Select Case DictKeys(iCounter)
							Case "Serial Number", "Serial Number After","Serial Number Before", "Lot Number", "Manufacturer's ID","Part Number"
									objIssueParts.JavaEdit("PropertyValue").SetTOProperty "Attached text", DictKeys(iCounter)&":"
									bResult= Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_IssueParts", "Set",  objIssueParts, "PropertyValue", DictItems(iCounter) )
									If bResult = False Then Exit Function
	
							Case "Location Name", "Disposition Value"
									objIssueParts.JavaList("PropertyList").SetTOProperty "Attached text", DictKeys(iCounter)&":"
									bResult= Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_IssueParts", "Select", objIssueParts,"PropertyList",DictItems(iCounter), "", "")
									If bResult = False Then Exit Function
	
						End Select
					Next
				
					' clicking on FInd button
					Call Fn_Button_Click("Fn_SISW_SrvScheduler_IssueParts", objIssueParts, "Find")
					wait(2)
					objIssueParts.Maximize
					wait(1)

					' Select Physical Part  from Search Result
					Select Case sSearchType
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "Preferred Parts"
									objIssueParts.JavaTable("SelectParts").SetTOProperty "attached text", sSearchType& " :"
									If dicIssueParts("SelectPhysicalPart") <> ""  Then
											bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_IssueParts", "GetRowIndex", objIssueParts, "SelectParts", "Object.GetItem", "Object", dicIssueParts("SelectPhysicalPart"), "", "", "", ":")
											 ' Check Condition if path returns false 
											If bResult = -1 Then Exit Function
											objIssueParts.JavaTable("SelectParts").ActivateRow bResult
									End If
							
									If objIssueParts.Exist(5) Then
										Call Fn_Button_Click("Fn_SISW_SrvScheduler_IssueParts", objIssueParts, "OK")
									End If


							Case "Installable Parts"
									objIssueParts.JavaTable("SelectParts").SetTOProperty "attached text", sSearchType
									If dicIssueParts("SelectPhysicalPart") <> ""  Then
											bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_IssueParts", "GetRowIndex", objIssueParts, "SelectParts", "Object.GetItem", "Object", dicIssueParts("SelectPhysicalPart"), "", "", "", ":")
											 ' Check Condition if path returns false 
											If bResult = -1 Then Exit Function
											objIssueParts.JavaTable("SelectParts").ActivateRow bResult
									End If
							
									If objIssueParts.Exist(5) Then
										Call Fn_Button_Click("Fn_SISW_SrvScheduler_IssueParts", objIssueParts, "OK")
									End If

					End Select
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_IssueParts ] invalied case [ " & sAction & " ].")
			Exit Function
	End Select

	Set objShell = Nothing
	Set objIssueParts = Nothing
	Fn_SISW_SrvScheduler_IssueParts = True
End Function


''*********************************************************		Function to handle the Part Request Operations from Service Scheduler	**************************************************************
'Function Name		:				Fn_SISW_SrvScheduler_PartRequestOperations.

'Description			 :		 		 This function is used to delete the Object Component from Schedule Manager.

'Parameters			   :	 			1. sAction - Specify the action(i.e.Menu,ShortKey,Toolbar)
'												   2. sMessage - message to Verify
'												   3. sButton - Button to Click On
											
'Return Value		   : 				PASS \ FAIL

'Examples				:				 Call Fn_SISW_SrvScheduler_PartRequestOperations("CancelPart","Delete the selected task(s)?","Yes")

' History:
'		Developer Name					Date						Rev. No.			Changes Done									Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'		Pranav Ingle 						25-Dec-2013				1.0																						Sunny			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_PartRequestOperations(sAction,sMessage,sButton)
GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_PartRequestOperations"
Dim objPartReq, sMsg, sTitle
Set objPartReq = Dialog("CancelPartRequestWarning")

Fn_SISW_SrvScheduler_PartRequestOperations = FALSE

	Select Case sAction
		Case "CancelPart"  
				sTitle = "Cancel Part Request Warning Dialog"
		Case "ClosePart" 
				sTitle = "Close Part Request Warning Dialog"
		Case "ValidatePart"
				sTitle = "Part Request validation error dialog"
	End Select

	''***********************************************************'' To check whether part Request Dialog  is displayed*************************************************************
	If  sTitle <> "" Then
		objPartReq.SetTOProperty "text", sTitle
	End If
	
	If objPartReq.Exist(2) Then
		If  sMessage <> "" Then
			sMsg = objPartReq.Static("ErrMsg").GetROProperty("text")
			If sMsg <> sMessage Then
				Call Fn_Button_Click("Fn_SISW_SrvScheduler_PartRequestOperations", objPartReq,"No")
				Exit Function
			End If
		End If
	Else
		Exit Function
	End If

	If sButton <> "" Then
		Call Fn_UI_WinButton_Click("Fn_SISW_SrvScheduler_PartRequestOperations",objPartReq,sButton,5,5,micLeftBtn)
	End If

	Set objPartReq = nothing
	Fn_SISW_SrvScheduler_PartRequestOperations = True
End Function





'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SrvScheduler_SearchCharacteristicOperation

'Description			 :	Function Used to Perform operation on Search  Characteristic dialog

'Parameters			   :  1.sAction : Action name
'									2.sInvokeOption: Dialog Invoke option
'								 	3.dicSearchCharacteristicInfo: Search information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be login to Service Scheduler perspective

'Examples				:   Dim dicSearchCharacteristicInfo
'										Set dicSearchCharacteristicInfo = CreateObject( "Scripting.Dictionary" )
'									dicSearchCharacteristicInfo("SearchName") = "ch*"			Note- Characteristic name to be serached
'									dicSearchCharacteristicInfo("Name") = "char1"						Note- the Characteristic name to be selected from Table 
'									dicSearchCharacteristicInfo("Unit") = ""
'									dicSearchCharacteristicInfo("Derived") = "ON"
'									dicSearchCharacteristicInfo("Type") = ""
'									msgbox  Fn_SISW_SrvScheduler_SearchCharacteristicOperation("AssignCharacteristic","toolbar" dicSearchCharacteristicInfo )
'
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Ganesh B									03-Jun-2014					1.0																				
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function  Fn_SISW_SrvScheduler_SearchCharacteristicOperation(sAction, sInvokeOption, dicSearchCharacteristicInfo )
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_SearchCharacteristicOperation"
	Dim objSearchDialog
	Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
	Set objSearchDialog = JavaWindow("DefaultWindow").JavaWindow("SearchCharacteristic")

		'Invoke Search Characteristic dialog
	If Not objSearchDialog.Exist(1) Then
		Select Case Lcase(sInvokeOption)
			Case "toolbar"
				If 	 Fn_SISW_UI_JavaToolbar_Operations("", "ClickExt", JavaWindow("DefaultWindow"), "", "Assign Characteristic...", "", "", 1) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_SearchCharacteristicOperation ] fail to click on toolbar button  [ Assign Characteristic... ]")
					Set objSearchDialog = Nothing
					Exit Function
				End If
				Call Fn_ReadyStatusSync(2)
			Case "nooption"
				'Use this option when user wants to invoke Search Characteristic dialog outside of function
		End Select
	End If
	Select Case  sAction
		Case "AssignCharacteristic"
			 call Fn_SISW_UI_JavaTab_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Select", objSearchDialog, "TabFolder", "Search")
			If Err.number < 0 Then
				Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select [Search Results] tab.")
				Set objSearchDialog = Nothing
				Exit Function
			End If
			If 	dicSearchCharacteristicInfo("SearchName") <> "" Then
				call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Set", objSearchDialog, "CharacteristicName", dicSearchCharacteristicInfo("SearchName"))
				If Err.Number < 0 Then
						Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Characteristic Name as [" + 	dicSearchCharacteristicInfo("SearchName")  + "]")
								Set objSearchDialog = Nothing
						Exit Function 
				Else					
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set  Characteristic Name as [" + 	dicSearchCharacteristicInfo("SearchName")  + "]")	
				End If
			End If
		
				If 	dicSearchCharacteristicInfo("Unit") <> "" Then
				call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Set", objSearchDialog, "Unit", dicSearchCharacteristicInfo("Unit"))
				If Err.Number < 0 Then
						Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Unit as [" + 	dicSearchCharacteristicInfo("Unit")  + "]")
								Set objSearchDialog = Nothing
						Exit Function 
				Else					
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set  Unit as [" + 	dicSearchCharacteristicInfo("Unit")  + "]")	
				End If
			End If
				If 	dicSearchCharacteristicInfo("Derived") <> "" Then
					If 	dicSearchCharacteristicInfo("Derived") = "ON"  Then
						call Fn_UI_Object_SetTOProperty("Fn_SISW_SrvScheduler_AssignCharacteristic",objSearchDialog.JavaRadioButton("Derived"), "attached text", "true")
						call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Set", objSearchDialog, "Derived", "ON")
						If Err.Number < 0 Then
								Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Derived as [" + 	dicSearchCharacteristicInfo("Derived")  + "]")
								Set objSearchDialog = Nothing
								Exit Function 
						Else					
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set  Derived as [" + 	dicSearchCharacteristicInfo("Derived")  + "]")	
						End If
					ElseIf 	dicSearchCharacteristicInfo("Derived") = "OFF"  Then
						call Fn_UI_Object_SetTOProperty("Fn_SISW_SrvScheduler_AssignCharacteristic",objSearchDialog.JavaRadioButton("Derived"), "attached text", "false")
						call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Set", objSearchDialog, "Derived", "ON")
						If Err.Number < 0 Then
								Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Derived as [" + 	dicSearchCharacteristicInfo("Derived")  + "]")
								Set objSearchDialog = Nothing
								Exit Function 
						Else					
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set  Derived as [" + 	dicSearchCharacteristicInfo("Derived")  + "]")	
						End If
					End If
			End If
				If 	dicSearchCharacteristicInfo("Type") <> "" Then
				call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Set", objSearchDialog, "Type", dicSearchCharacteristicInfo("Type"))
				If Err.Number < 0 Then
						Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Type as [" + 	dicSearchCharacteristicInfo("Type")  + "]")
						Set objSearchDialog = Nothing
						Exit Function 
				Else					
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set  Type as [" + 	dicSearchCharacteristicInfo("Type")  + "]")	
				End If
			End If
			call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Click", objSearchDialog, "Find")
			If Err.number < 0 Then
				Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Find] Button")
				Set objSearchDialog = Nothing
				Exit Function
			End If
			call Fn_SISW_UI_JavaTab_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Select", objSearchDialog, "TabFolder", "Search Results")
			If Err.number < 0 Then
				Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select [Search Results] tab.")
				Set objSearchDialog = Nothing
				Exit Function
			End If
			wait 5
			call Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "ClickCell", objSearchDialog , "MROSearchCharacteristic", "", "Object", dicSearchCharacteristicInfo("Name"), "", "", "", "")
			If Err.number < 0 Then
				Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select [" & dicSearchCharacteristicInfo("Name") & "].")
				Set objSearchDialog = Nothing
				Exit Function
			End If
			call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_AssignCharacteristic", "Click", objSearchDialog, "OK")
			If Err.number < 0 Then
				Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [OK] Button")
				Set objSearchDialog = Nothing
				Exit Function
			Else
				Fn_SISW_SrvScheduler_SearchCharacteristicOperation = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on  [OK] Button")	
			End If
		Case Else
				Fn_SISW_SrvScheduler_SearchCharacteristicOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "[Fn_SISW_SrvScheduler_AssignCharacteristic]FAIL: Invalid Case.")		
				Set objSearchDialog = Nothing
	End Select
End Function


'****************************************    Function to perform Find and Open Proxy Task***************************************
'
''Function Name		 	:	Fn_SISW_SrvScheduler_FindProxyTask
'
''Description		    :  	Function to perform  assign lot operation

''Parameters		    :	1. sAction : Action need to perform
			'					   		2. sScheduleName
			'					   		3. sMessage
			'							4. sButtonName
								
''Return Value		    :  	True \ False
'
''Pre-requisite		    :	Proxy Tasks Window should be present.

''Examples		     	:	

'Examples		     	:	Case "SelectAndGoTo"
'								Fn_SISW_SrvScheduler_FindProxyTask("SelectAndGoTo", "WPlan067903", "", "GoTo")
'History:
'	Developer Name				Date					Rev. No.			Reviewer			Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle		 			24-Jan-2014				1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvScheduler_FindProxyTask(sAction, sScheduleName, sMessage, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_FindProxyTask"
	Dim objProxyTask
	Set objProxyTask = JavaWindow("ServiceScheduler").JavaWindow("ProxyTasks")
	Fn_SISW_SrvScheduler_FindProxyTask = False

	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvScheduler_FindProxyTask","Exist", objProxyTask,"") = False Then
			Exit Function
	End If


	Select Case sAction
		Case "CloseDialog"
				Call Fn_Button_Click("Fn_SISW_SrvScheduler_FindProxyTask", objProxyTask, "Close")
				Fn_SISW_SrvScheduler_FindProxyTask = True

		Case "SelectAndGoTo"
					If sScheduleName <> "" Then
							bResult= Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvScheduler_FindProxyTask", "Select", objProxyTask,"SchedulesThatHaveProxy",sScheduleName, "", "")
							If bResult = False Then Exit Function
					End If

					If sButtonName <> "" Then
							Call Fn_Button_Click("Fn_SISW_SrvScheduler_FindProxyTask", objProxyTask, sButtonName)
					End If
			
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_FindProxyTask ] invalied case [ " & sAction & " ].")
			Exit Function
	End Select

	Set objProxyTask = Nothing
	Fn_SISW_SrvScheduler_FindProxyTask = True
End Function

'*********************************************************		Function to Apply Workflow Rule  ***********************************************************************

'Function Name		:					Fn_SISW_SrvScheduler_WorkflowRuleConfiguration

'Description			 :		 		  This function is used to Apply Workflow Rule in Service Scheduler.

'Parameters			   :	 			1.  sAction
'													 2. dicWorkFlowRuleInfo: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		Service Scheduler apllication should be opened

'Examples				:              
'											Case "Apply"
'													'Declaration of the object as structure.
'													
'													Set dicWorkFlowRuleInfo = CreateObject( "Scripting.Dictionary" )
'													
'													With dicWorkFlowRuleInfo  
'																 .Add "NodeName", "rd:card1~rd:card2"
'																.Add "WorkFlowTrigger", "No workFlow trigger"	                
'																.Add "WorkFlowTemplate", "TCM Release Process"
'																.Add "PrivilegedUser", ""
'																.Add "ProcessOwner", ""
'													End with

'													call  Fn_SISW_SrvScheduler_WorkflowRuleConfiguration("Apply",  dicWorkFlowRuleInfo)

'History:
'	Developer Name					Date				Rev. No.				Changes Done																	Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ganesh Bhosale			13-May-2014				1.0							Created New Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_SrvScheduler_WorkflowRuleConfiguration(sAction, dicWorkFlowRuleInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_WorkflowRuleConfiguration"
	Dim objWorkflowRuleConfiguration, objSchJobTable, sMenu
	Dim sTriggerType, sTemplateType, intNoOfObjects
	Set objWorkflowRuleConfiguration = JavaWindow("ServiceScheduler").JavaApplet("SrvSchdulerApplet").JavaDialog("WorkflowRuleConfiguration")
    Select Case sAction
		Case "Apply"
				If NOT objWorkflowRuleConfiguration.Exist(1) Then
					If dicWorkFlowRuleInfo("NodeName") <> "" Then

						''  Select Node to invoke  Schedule: WorkflowTask Menu
						If  inStr(dicWorkFlowRuleInfo("NodeName"), "~")> 0 Then
							call Fn_SISW_SrvScheduler_SchTable_NodeOperation("MultiSelect" ,dicWorkFlowRuleInfo("NodeName"), "" , "" , "")		
						Else
							call Fn_SISW_SrvScheduler_SchTable_NodeOperation("Select" ,dicWorkFlowRuleInfo("NodeName"), "" , "" , "")		
						End If
					End If
					sFilePath = Fn_LogUtil_GetXMLPath("ServiceScheduler")
					sMenu = Fn_GetXMLNodeValue(sFilePath, "ScheduleWorkflowTask")
					Call Fn_MenuOperation("Select", sMenu)
					If Err.Number <  0 Then
						 Fn_SISW_SrvScheduler_WorkflowRuleConfiguration = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_WorkflowRuleConfiguration : Failed To invoke WorkFlow Rule Configuration Window")	
						Exit Function
					End If
				End If
				''  select  Workflow Template
				If dicWorkFlowRuleInfo("WorkFlowTemplate") <> ""Then
						objWorkflowRuleConfiguration.JavaStaticText("PropertyLabel").SetTOProperty "Label", "Workflow Template"
						objWorkflowRuleConfiguration.JavaButton("TemplateButton").Click micLeftBtn
						wait(1)
						Set sTemplateType=Description.Create()
						sTemplateType("Class Name").value = "JavaStaticText"
						sTemplateType("label").value = dicWorkFlowRuleInfo("WorkFlowTemplate") 	
						Set  intNoOfObjects =objWorkflowRuleConfiguration.ChildObjects(sTemplateType)
					   If  intNoOfObjects.count > 0 Then
								intNoOfObjects(0).Click 5,5
					   End If
						  Set  intNoOfObjects = Nothing
						Set sTemplateType = Nothing
						If Err.Number <  0 Then
							 Fn_SISW_SrvScheduler_WorkflowRuleConfiguration = FALSE				 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_WorkflowRuleConfiguration : Failed To select  Workflow Template " & dicWorkFlowRuleInfo("WorkFlowTemplate") )	
							Exit Function
						End If
				End If

				''  select  Workflow Trigger
				If dicWorkFlowRuleInfo("WorkFlowTrigger") <> ""Then
						objWorkflowRuleConfiguration.JavaStaticText("PropertyLabel").SetTOProperty "Label", "Workflow Trigger"
						objWorkflowRuleConfiguration.JavaButton("TemplateButton").Click micLeftBtn
						wait(1)
						Set sTriggerType=Description.Create()
						sTriggerType("Class Name").value = "JavaStaticText"
						sTriggerType("label").value = dicWorkFlowRuleInfo("WorkFlowTrigger")	
						Set  intNoOfObjects =objWorkflowRuleConfiguration.ChildObjects(sTriggerType)
					   If  intNoOfObjects.count > 0 Then
								intNoOfObjects(0).Click 5,5
					   End If
						 Set  intNoOfObjects = Nothing
						Set sTriggerType = Nothing
						If Err.Number <  0 Then
							 Fn_SISW_SrvScheduler_WorkflowRuleConfiguration = FALSE				 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_WorkflowRuleConfiguration : Failed To select  WorkFlow Trigger " & dicWorkFlowRuleInfo("WorkFlowTrigger") )	
							Exit Function
						End If
				End If

				'' need to update this case if required
				If dicWorkFlowRuleInfo("PrivilegedUser") <> ""Then
				End If
				'' need to update this case if required
				If dicWorkFlowRuleInfo("ProcessOwner") <> ""Then
				End If

				'' Click on OK button
				Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_WorkflowRuleConfiguration", "Click", objWorkflowRuleConfiguration, "Apply")
				Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvScheduler_WorkflowRuleConfiguration", "Click", objWorkflowRuleConfiguration, "OK")
				If Err.Number <  0 Then
					Fn_SISW_SrvScheduler_WorkflowRuleConfiguration = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_WorkflowRuleConfiguration : Failed To Apply WorkFlow Rule Configuration.")	
					Exit Function
				Else 
					Fn_SISW_SrvScheduler_WorkflowRuleConfiguration = True	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvScheduler_WorkflowRuleConfiguration : Successfully Apllied  WorkFlow Rule Configuration")
				End If
		Case Else
			Fn_SISW_SrvScheduler_WorkflowRuleConfiguration = FALSE				 
	End Select
	Set objWorkflowRuleConfiguration = Nothing
End Function

'*********************************************************	Function to Generate Automated Service Schedule  ***********************************************************************

'Function Name		:				Fn_SISW_GenAutoServiceSchedule_Opearations

'Description	    :		 		This function is used to perform various operations on Generate Automated Service Schedule dialog.

'Parameters			:	 			1.  sAction - action to be performed
'									2. sRootStructureNode - Root Node name to invoke Generate Automated Service Schedule dialog. 
'								    3. dicGenAutoServiceScheule - Dictionary For Inputs to perform operations on Generate Automated Service Schedule dialog.
'									4. dicWorkOrderInputs - Dictionary For Inputs to create work order.
'									5. sButtonName - button name to click while exiting from Generate Automated Service Schedule dialog.

'Return Value		: 				 True Or False

'Pre-requisite	    :		 		Service Scheduler application should be opened

'Examples			:                  ----------------------------------------------------------------------------------
'										Dictionary to create generate Automated Service Schedule
'									   ----------------------------------------------------------------------------------
'										Set dicGenAutoServiceScheule = CreateObject( "Scripting.Dictionary" )												
'										With dicGenAutoServiceScheule  
'											.Add "TabName", "Service Plan"
'											.Add "ServicePlanName", "Sp10"	                
'											.Add "BOMNode", "Sp10:000196/A;1-Sr10 (View)"
'											.Add "NewMaintenanceAction","Due Date:6/21/2016|10:22:23 AM~Note:TestNote~Button:Finish"
'											.Add "ShowMaintenanceAction","SrcReq"
'											.Add "GenerateSchedule","True"
'										End with
'
'										----------------------------------------------------------------------------------
'										'Create Work Order
'										----------------------------------------------------------------------------------
'										Set dic2 = CreateObject( "Scripting.Dictionary" )
'										PlanID = Fn_Setup_RandNoGenerate(6)
'
'										dic2.RemoveAll
'										dic2("Name") = "WPlan10"
'										dic2("Is Schedule Public") = "True"
'										dic2("Is Percent Linked") = "True"
'										dic2("Published") = "True"
'										dic2("Are notifications enabled") = "True"
'										dic2("Use Finish Date Scheduling") = "False"
'
'										Set dic3 = CreateObject( "Scripting.Dictionary" )
'										dic3.RemoveAll
'										dic3("sAction") = "Add..."
'										dic3("ID") = "000193"
'										dic3("SearchResults_Select") ="000193"+"-"+"Part10_Loc10"
'
'										Set dicWorkOrderInputs = CreateObject( "Scripting.Dictionary" )
'										dicWorkOrderInputs.RemoveAll
'										dicWorkOrderInputs("Synopsis") = "WorkOrder"
'										dicWorkOrderInputs("ID") = PlanID
'										dicWorkOrderInputs("Revision") = "A"
'										Set dicWorkOrderInputs("Plan") = dic2
'										Set dicWorkOrderInputs("Work Performed At") = dic3
'
'										call Fn_SISW_GenAutoServiceSchedule_Opearations("CreateWorkOrder","000192/--A",dicGenAutoServiceScheule,dicWorkOrderInputs,"Close")
'							Example for "VerifyDisplayedMaintenanceTableRow", "GetDisplayedMaintenanceTableRowNo"
'										Set dic = CreateObject("Scripting.Dictionary")
'										dic("ColumnName") = "Due Date~Is Overdue~Name~Creation Date~Is Scheduled~Note~Auto-complete~Asset~In Progress~Corrective Action~Requirements~Disposition~Approval~Configuration"
'										dic("ColumnValue") = "~False~SR1~20-Jul-2016 12:57~True~~False~000071~000072~[WC1, WC2]~000079/A;1-SR1~Requested~~SP1"
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("VerifyDisplayedMaintenanceTableRow","",dic,"","")
'										bReturn1 = Fn_SISW_GenAutoServiceSchedule_Opearations("GetDisplayedMaintenanceTableRowNo","",dic,"","")
'							Example for "CreateWorkOrder", "GetCreateMaintActionTableRowNo"			
'										dic("TabName") = "Maintenance Action"
'										dic("GetRowToSet") = "Name:SrcReq1~Asset:000293/-~Impacted Part:000293/-~Part Position:-"
'										dic("NewMaintenanceAction") = "Due Date:6/21/2016|10:22:23 AM~AutoComplete:ON~Note:TestNote~Button:Finish"
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("CreateWorkOrder","",dic,"","")
'							Example for "SetCreateMaintActTableRow", "GetCreateMaintActionTableRowNo"
'										dicGenAutoServiceScheule("GetRowFromValues") = "Name:SrcReq1~Asset:000293/-~Impacted Part:000293/-~Part Position:-"
'										dicGenAutoServiceScheule("SetValuesToRow") = "Due Date:6/21/2016|10:22:23 AM~AutoComplete:ON~Note:TestNote~Button:Finish"
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("SetCreateMaintActTableRow","",dicGenAutoServiceScheule,"","")
'							Example for "PopupMenuSelectDMTRow" on Displyed Maintenance Table row
'										dic("Button") = "Show Maintenance Action"
'										dic("ColumnName") = "Name~Creation Date~Due Date~Is Overdue~Is Scheduled~Note~Auto-complete~Asset~In Progress"
'										dic("ColumnValue") = "RQ_48538_08907~01-Aug-2016 17:25~01-Aug-2016 17:25~True~False~TestNote~False~000045/1~000045/1"
										'PopupMenu "Cancel Maintenance Action" with Error verification
'										dic("PopupMenu") = "CancelMaintenanceActionWithError"
'										dic("Button1") = "Yes"
'										dic("ErrorMessage") = " role to cancel Maintenance Action."
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("PopupMenuSelectDMTRow","",dic,"","")
'							Example for Verify BOM node in tabs "VerifyBOMNode"
'										dic("TabName") = "ServicePlan"
'										'dic("ServicePlanName") = "SrcPlan5372"
'										dic("ServicePlanName") = "SrcPlan5372:000181/A;1-SrcReq1_5372 (View)"
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("VerifyBOMNode","",dic,"","")
'							Example for Verify or Get Row in "Service Discreoancies" table after clickin "Show Discrepancies" button
'							Example for check the row in [Service Discrepancies] or [Displayed Maintenance] Table
'										dic("ColumnName") = "Name~Discovered By~Discovery Date~Is Failure"
'										dic("ColumnValue") = "SD2~AutoTest2 (autotest2)~03-Aug-2016 11:23~False"
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("VerifyServiceDiscrepanciesTableRow","",dic,"","")
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("GetServiceDiscrepanciesTableRowNo","",dic,"","")
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("CheckRowInSDTable","",dic,"","")
'							Example for verify "Relate Service Discrepancy and Service Requirement" window
'										dic("Button") = "Relate Requirement"
'										dic("VerifyServiceDiscrepancyNode") = "SD1~SD2"
'										dic("VerifyServiceRequirementNode") = "000064-SR1"
'										dic("ConfirmORCancelButton") = "Cancel"
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("RelateSDandSROperations","",dic,"","")
'							Example for clicking any button on [Generate Automated Service Schedule] window
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("ButtonClick","","","","Relate Requirement")
'							Example for clicking button below Service Tree on [Generate Automated Service Schedule] window
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("ServiceTreeBtnClick","","","","Clear Selection")
'							Example for clicking button below Service Table on [Generate Automated Service Schedule] window
'										bReturn = Fn_SISW_GenAutoServiceSchedule_Opearations("ServiceTableBtnClick","","","","Clear Selection")
'History:
'	Developer Name		Date		Rev. No.	Changes Done													Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod	06-Jun-2016		1.0			Created New Function											Ankit N
'	Vivek Ahirrao	21-Jul-2016		1.0			Added cases "VerifyDisplayedMaintenanceTableRow", "GetDisplayedMaintenanceTableRowNo"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_GenAutoServiceSchedule_Opearations(sAction, sRootStructureNode, dicGenAutoServiceScheule, dicWorkOrderInputs, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_GenAutoServiceSchedule_Opearations"
	Dim sArr, bFlag, arrNodes, sPath, sTreePath, jCnt, iCnt, xCo, yCo,iRowCount,sNodeName
	Dim objGenAutoSerSchedule, objTree, oCurrentNode,objNewMaint,objTable,objProperties,objCheckOut
	Dim WshShell,objDateControl,sDate,sTime,ArrDateTime,iCounter,objCheckIn
	Dim sBounds, aBounds, sSubAction
	Dim sProperty,sAppValue,sValue,iCount,iTotalElements,aProperty,aValues,iCount1
	
	Fn_SISW_GenAutoServiceSchedule_Opearations = False
	
	Set objGenAutoSerSchedule = Fn_SISW_SrvScheduler_GetObject("GenerateAutomatedService")
	
	If objGenAutoSerSchedule.Exist(2) = False Then
		If sRootStructureNode <> "" then
		    bFlag = Fn_SISW_SrvScheduler_BOMTable_NodeOperation("PopupSelect", sRootStructureNode, "", "", "Generate Automated Service Schedule")
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to perform [ RMB : Generate Automated Service Schedule ] on Root Structure Node [ " & sRootStructureNode & " ].")
				Set objGenAutoSerSchedule = Nothing
				Exit function
			End If
		End If
	End If
		
	If objGenAutoSerSchedule.Exist(2) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to find [ Generate Automated Service Schedule ] window.")
		Set objGenAutoSerSchedule = Nothing
		Exit function
	End If
	
	If sAction <> "EditMAProperties" AND sAction <> "VerifyMAProperties" Then
		objGenAutoSerSchedule.Maximize
		wait 2
	End If
		
	Select Case sAction
		'Case to Select Tab in [Generate Automated Service Schedule] Window
		Case "SelectTab"
				If dicGenAutoServiceScheule<>"" Then
						sTabName = dicGenAutoServiceScheule
						If sTabName <> "" Then
							If sTabName = "Service Plan" Then
								objGenAutoSerSchedule.JavaTab("ScheduleTab").Select "Service Plan"
							ElseIf sTabName = "Service Discrepancy" Then
								objGenAutoSerSchedule.JavaTab("ScheduleTab").Select "Service Discrepancy"
							ElseIf sTabName = "Maintenance Action" Then
								objGenAutoSerSchedule.JavaTab("ScheduleTab").Select "Maintenance Action"
							End If
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select tab ["+sTabName+"] in Action ["+sAction+"].")
								Set objGenAutoSerSchedule = Nothing
								Exit Function
							End if
						End If
				End If
		'Case to verify BOM node in [Service Plan], [Service Descrepancy] or [Maintenance Action] tabs
		Case "VerifyBOMNode"
				'Select Tab
				If dicGenAutoServiceScheule("TabName") <> "" Then
						bFlag = Fn_SISW_GenAutoServiceSchedule_Opearations("SelectTab","",dicGenAutoServiceScheule("TabName"),"","")
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select tab ["+dicGenAutoServiceScheule("TabName")+"] in Action ["+sAction+"].")
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End if
				End If
				'Verify Node in Service Requirements tree
				If dicGenAutoServiceScheule("ServicePlanName") <> "" Then
						Set objTree = objGenAutoSerSchedule.JavaTree("ServiceRquireTree")
						'Expand Path upto Parent node
						aNodeName = Split(dicGenAutoServiceScheule("ServicePlanName"),":")
						sNodePath = ""
						For iCount = 0 To UBound(aNodeName)
							If iCount=0 Then
								sNodePath = aNodeName(iCount)
							Else
								sNodePath = sNodePath & ":" & aNodeName(iCount)
							End If
							'Get node path
							sTreePath = Fn_UI_JavaTreeGetItemPathExt("",objTree,sNodePath,"","")
							If sTreePath <> False Then
								Wait 0,500
								'Expand Path
								If iCount <> UBound(aNodeName) Then
									objTree.Expand sTreePath
									If Err.Number<0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Expand node ["+sNodePath+"].")
										Set objGenAutoSerSchedule = Nothing
										Set objTree = Nothing
										Exit function
									End If
									Wait 1
								End If
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : BOM Table Node ["+sNodePath+"] is not Present.")
								Set objGenAutoSerSchedule = Nothing
								Set objTree = Nothing
								Exit function
							End If
						Next
						Wait 1
						'Check existance of node
						sTreePath = Replace(sTreePath,"#","")
						aTreePath = Split(sTreePath,":")
						Set oCurrentNode = ObjTree.Object.getItem(aTreePath(0))
						sFullNode = oCurrentNode.getData.toString()
						For iCount = 1 To UBound(aTreePath)
							Set oCurrentNode = oCurrentNode.getItem(aTreePath(iCount))
							sFullNode = sFullNode + ":" + oCurrentNode.getData.toString()
						Next
						If dicGenAutoServiceScheule("ServicePlanName")<>sFullNode Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : BOM Table Node ["+dicGenAutoServiceScheule("ServicePlanName")+"] is not Present.")
							Set objGenAutoSerSchedule = Nothing
							Set oCurrentNode = Nothing
							Set objTree = Nothing
							Exit Function
						End If
						Set oCurrentNode = Nothing
						Set objTree = Nothing
				End If	
				
		Case "CreateWorkOrder","SelectBOMNode","SetCreateMaintActionRow"
				If sAction<>"SetCreateMaintActionRow" Then
						'Select Tab
						If dicGenAutoServiceScheule("TabName") <> "" Then
							bFlag = Fn_SISW_GenAutoServiceSchedule_Opearations("SelectTab","",dicGenAutoServiceScheule("TabName"),"","")
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Failed to Select tab ["+dicGenAutoServiceScheule("TabName")+"] in Action ["+sAction+"].")
								Set objGenAutoSerSchedule = Nothing
								Exit Function
							End if
						End If

						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","Show Service Plan(s)"
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
						If bFlag = false then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Show Service Plan(s) ] button in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit function
						End if
				End If
				wait 5
				Set objTree = objGenAutoSerSchedule.JavaTree("ServiceRquireTree")
				if dicGenAutoServiceScheule("ServicePlanName") <> "" then
						arrNodes = split(dicGenAutoServiceScheule("ServicePlanName"), ":")
						sPath = arrNodes(0)
						sTreePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_GenAutoServiceSchedule_Opearations", objTree, sPath, "", "")
						If sTreePath <> False Then
							Err.clear
							objTree.Expand sTreePath
							Wait 1
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] BOM TableNode is a root node.")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] BOM TableNode is not Present.")
							Set objGenAutoSerSchedule = Nothing
							Set objTree = Nothing
							Exit function
						End If
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to exapnd node [ "+sPath+" ] JavaTree [ ServiceRquireTree  ] in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Set objTree = Nothing
							Exit function
						End if
						if ubound(arrNodes) > 0 then
							For iCnt = 1 to ubound(arrNodes)
								sPath = sPath + ":"+arrNodes(iCnt)
								sTreePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_GenAutoServiceSchedule_Opearations", objTree, sPath, "", "")
								If sTreePath <> False Then
									Err.clear
									objTree.Expand sTreePath
									Wait 1
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to exapnd node [ "+sPath+" ] JavaTree [ ServiceRquireTree  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objTree = Nothing
										Exit function
									End if
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] expanded BOM TableNode .")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] fail to expanded BOM TableNode.")
									Set objGenAutoSerSchedule = Nothing
									Set objTree = Nothing
									Exit function
								End If
							next
						end if
				End if				
				If dicGenAutoServiceScheule("BOMNode") <> "" Then
						arrNodes = split(dicGenAutoServiceScheule("BOMNode"),"~")
						For iCnt = 0 to uBound(arrNodes)
							sTreePath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_GenAutoServiceSchedule_Opearations", objTree, arrNodes(iCnt), "", "")
							If sTreePath <> False Then
								If dicGenAutoServiceScheule("GenerateSchedule") = "True" Then								
									sTreePath = replace(sTreePath,"#","")
									sArr = split(sTreePath,":")
									Set oCurrentNode = ObjTree.Object.getItem(sArr(0))
									For jCnt = 1 to ubound(sArr)
										Set oCurrentNode = oCurrentNode.getItem(sArr(jCnt))
									next
									Set oCurrentNode = oCurrentNode.getBounds()
									sTreePath = oCurrentNode.tostring 
									sArr = split(trim(replace(replace(sTreePath,"Rectangle {",""),"}","")),",")
									xCo = sArr(0)+10
									yCo = sArr(1)+3
									Err.clear
									objTree.Click xCo,yCo
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to click check box of node [ "+arrNodes(iCnt)+" ] JavaTree [ ServiceRquireTree  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objTree = Nothing
										Set oCurrentNode = Nothing
										Exit function
									End if
									Set oCurrentNode = Nothing
								else
									Err.clear
									objTree.Select sTreePath
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to select node [ "+arrNodes(iCnt)+" ] JavaTree [ ServiceRquireTree  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objTree = Nothing
										Exit function
									End if
								End if
								wait 3
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] BOM TableNode ["+arrNodes(iCnt)+"] is checked.")
							else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] BOM TableNode ["+arrNodes(iCnt)+"] is failed to checked.")
								Set objTree = Nothing
								Exit function
							End if
						next
				End if
				Set objTree = Nothing
			'------------------------- Maintainance Action Tab code -----------------------------------------------------------------------
				'for setting row which we want
				'dicGenAutoServiceScheule("NewMaintenanceAction") = "Due Date:6/21/2016|10:22:23 AM~AutoComplete:ON~Note:TestNote~Button:Finish"
				'dicGenAutoServiceScheule("GetRowToSet") = "Name:SrcReq1~Asset:000293/-~Impacted Part:000293/-~Part Position:-"
							'OR
				'for setting row only 0th
				'dicGenAutoServiceScheule("NewMaintenanceAction") = "Due Date:6/21/2016|10:22:23 AM~AutoComplete:ON~Note:TestNote~Button:Finish"
				If dicGenAutoServiceScheule("NewMaintenanceAction") <> "" OR (dicGenAutoServiceScheule("NewMaintenanceAction") <> "" AND dicGenAutoServiceScheule("GetRowToSet") <> "") Then
						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","New Maintenance Action"
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
						wait 1
						If bFlag = false then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Show Service Plan(s) ] button in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit function
						End if
						Wait 1
						Set objNewMaint = objGenAutoSerSchedule.JavaWindow("CreateMaintenanceAction")
						objNewMaint.SetTOProperty "title","Create Maintenance Action"
						If Fn_UI_ObjectExist("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint) = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to find [ "+objNewMaint.tostring+" ] window.")
							Set objNewMaint = Nothing
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						Else
							Call Fn_Window_Maximize("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint)
							wait 1
						End If
						
						If dicGenAutoServiceScheule("GetRowToSet") <> "" Then
							Set dicCreateMainActionDic = CreateObject("Scripting.Dictionary")
							dicCreateMainActionDic.RemoveAll
							aRowToSet = Split(dicGenAutoServiceScheule("GetRowToSet"),"~")
							For iCount = 0 To UBound(aRowToSet)
								aColNameValue = Split(aRowToSet(iCount),":")
								If aColNameValue(0)="Due Date" Then
									aColNameValue(1) = sArr(1)
									If UBound(sArr)>1 Then
										For iTimeCount = 2 To UBound(sArr) Step + 1
											aColNameValue(1) = aColNameValue(1) &":"& sArr(iTimeCount)
										Next
									End If
								End If
								If iCount = 0 Then
									dicCreateMainActionDic("ColumnName") = aColNameValue(0)
									dicCreateMainActionDic("ColumnValue") = aColNameValue(1)
								Else
									dicCreateMainActionDic("ColumnName") = dicCreateMainActionDic("ColumnName") &"~"& aColNameValue(0)
									dicCreateMainActionDic("ColumnValue") = dicCreateMainActionDic("ColumnValue") &"~"& aColNameValue(1)
								End If
							Next
							
							iRowNumber = Fn_SISW_GenAutoServiceSchedule_Opearations("GetCreateMaintActionTableRowNo","",dicCreateMainActionDic,"","")
							If iRowNumber=-1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Find row in [ Create Maintenance Action ] Table.")
								Set objGenAutoSerSchedule = Nothing
								Set objNewMaint = Nothing
								Set dicCreateMainActionDic = Nothing
								Exit Function
							End If
							Set dicCreateMainActionDic = Nothing
						Else
							sNodeName = Fn_UI_JavaTable_GetCellData("Fn_SISW_GenAutoServiceSchedule_Opearations",objNewMaint,"CreateMaintenanceTable",0,"Name")
							If instr(1,dicGenAutoServiceScheule("BOMNode"),sNodeName) = 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to find row with [ "+sNodeName+" ] in [ Create Maintenance Action  ] dialog in [ Generate Automated Service Schedule ] window.")
								Set objGenAutoSerSchedule = Nothing
								Set objNewMaint = Nothing
								Exit function
							End If
							iRowNumber = 0
						End If
						
						arrNodes = split(dicGenAutoServiceScheule("NewMaintenanceAction"),"~")
						For iCnt = 0 to uBound(arrNodes)
							sArr = split(arrNodes(iCnt),":")
							Select Case sArr(0)
								Case "Due Date"
									'set Row Number as Index of button
									objNewMaint.JavaButton("DueDate").SetTOProperty "index",iRowNumber
									sDateTime = sArr(1)
									If UBound(sArr)>1 Then
										For iTimeCount = 2 To UBound(sArr) Step + 1
											sDateTime = sDateTime &":"& sArr(iTimeCount)
										Next
									End If								
									ArrDateTime = Split(sDateTime,"|")
									'Date Conversion from 6/21/2016 to 21-Jun-2016 format
									sDate = ArrDateTime(0)
									aDate = Split(sDate,"/")
									sDate = aDate(1) &"-"& MonthName(aDate(0),True) &"-"& aDate(2)
									'Time
									sTime = ArrDateTime(1)
									'Click on "Due Date" Control button
									bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint,"DueDate")
									If bFlag = false then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set [ Due Date ] in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
									Wait 1
									bFlag = Fn_UI_SetDateAndTime("Fn_SISW_GenAutoServiceSchedule_Opearations",sDate,sTime)
									If bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set on [ Due Date ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
									Wait 1
								Case "AutoComplete"
									'set Row Number as Index of CheckBox
									objNewMaint.JavaCheckBox("AutoComplete").SetTOProperty "index",iRowNumber
									bFlag = Fn_CheckBox_Set("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint, "AutoComplete", sArr(1))
									If bFlag = false then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to check [ AutoComplete ] in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
								Case "Note"
									'set Row Number as Index of Edit Box
									objNewMaint.JavaEdit("Note").SetTOProperty "index",iRowNumber
									bFlag = Fn_Edit_Box("Fn_SISW_GenAutoServiceSchedule_Opearations",objNewMaint,"Note",sArr(1))
									If bFlag = false then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set [ Note ] in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
								Case "Button"
									bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint,sArr(1))
									wait 3
									If bFlag = false then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+sArr(1)+" ] button in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
							End select
						Next
						Set objNewMaint = Nothing
				End if
				
				If dicGenAutoServiceScheule("ShowMaintenanceAction") <> "" Then
						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","Show Maintenance Action"
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
						If bFlag = false then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Show Maintenance Action ] button in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit function
						End if
						wait 2
	
						Set objTable = objGenAutoSerSchedule.JavaTable("DisplayedMaintenance")
						'iRowCount = objTable.GetROProperty("rows")
						For iCnt = 0 to cint(objTable.GetROProperty("rows")) - 1
							If instr(1,dicGenAutoServiceScheule("ShowMaintenanceAction"),objTable.Object.getItem(iCnt).getData().toString()) > 0 Then
								 objTable.Object.getItem(iCnt).setChecked("true")
								 wait 2
							Else
							    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Maintenance Action ] checkbox in [ Generate Automated Service Schedule ] window.")
								Set objGenAutoSerSchedule = Nothing
								Set objTable = Nothing
								Exit Function
							End if 
						Next
						Set objTable = Nothing
				ElseIf dicGenAutoServiceScheule("ColumnName")<>"" AND dicGenAutoServiceScheule("ColumnValue")<>"" Then
						Set objTable = objGenAutoSerSchedule.JavaTable("DisplayedMaintenance")
						Set dicDispMainDic = CreateObject("Scripting.Dictionary")
						dicDispMainDic.RemoveAll
						dicDispMainDic("Button") = "Show Maintenance Action"
						dicDispMainDic("ColumnName") = dicGenAutoServiceScheule("ColumnName")
						dicDispMainDic("ColumnValue") = dicGenAutoServiceScheule("ColumnValue")
						iRowNumber = Fn_SISW_GenAutoServiceSchedule_Opearations("GetDisplayedMaintenanceTableRowNo","",dicDispMainDic,"","")
						If iRowNumber<>-1 Then
							objTable.Object.getItem(iRowNumber).setChecked("true")
							Wait 2
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Find row in [ Display Maintenance ] Table in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Set dicDispMainDic = Nothing
							Exit Function
						End If
						Set dicDispMainDic = Nothing
				End if	
				'------------------------- Maintainance Action Tab code end -----------------------------------------------------------------------
				If sAction = "CreateWorkOrder" Then
						Call Fn_UI_JavaStaticText_Click(" Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "DownArrow", 1, 1, "LEFT")
						Wait 2
						
						bFlag=Fn_UI_JavaMenu_Select("Fn_SISW_GenAutoServiceSchedule_Opearations",objGenAutoSerSchedule,"Create...")
						If bFlag = False Then 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to perform menu operation [ Create... ] on [ GGenerate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit function
						End if
						Wait 2
						If varType(dicWorkOrderInputs)="9" Then
							bFlag = Fn_SISW_SrvScheduler_CreateWorkOrder("Create","Finish",dicWorkOrderInputs)
							If bFlag = False Then 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to create work order")
								Set objGenAutoSerSchedule = Nothing
								Exit function
							End if			
						End If
						If dicGenAutoServiceScheule("GenerateSchedule") = "True" then
							objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","Generate Schedule"
							bFlag=Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objGenAutoSerSchedule,"ShowServicePlan")
							If bFlag = False Then 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Generate Schedule ] button on [ Generate Automated Service Schedule ] window.")
								Set objGenAutoSerSchedule = Nothing
								Exit function
							End if
							Wait 5
						End if
				End If

				'Case to Add existing WorkOrder
		Case "AddExistingWorkOrder","SelectFrmMyWrkOrderPage","VerifyWorkOrders","VerifyDefaultHighlightedWorkOrder","AddExistingWorkOrderAndGenerate"

				if sAction <> "AddExistingWorkOrderAndGenerate" Then 
					Call Fn_UI_JavaStaticText_Click(" Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "DownArrow", 1, 1, "LEFT")
					Wait 2
					bFlag=Fn_UI_JavaMenu_Select("Fn_SISW_GenAutoServiceSchedule_Opearations",objGenAutoSerSchedule,"Add...")
					If bFlag = False Then 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to perform menu operation [ Add... ] on [ Generate Automated Service Schedule ] window.")
						Set objGenAutoSerSchedule = Nothing
						Exit function
					End if
					Wait 2
				End if

				Set objTable = JavaWindow("ServiceScheduler").JavaWindow("Search") 
				If sAction = "VerifyDefaultHighlightedWorkOrder" Then 
					If dicGenAutoServiceScheule("WorkOrder") <> "" Then 

						iRowNumber = objTable.JavaTable("SearchResultTable").Object.getSelectionIndex
							sNodeName = Fn_UI_JavaTable_GetCellData("Fn_SISW_GenAutoServiceSchedule_Opearations", objTable, "SearchResultTable",iRowNumber,0)
							If Instr(trim(sNodeName),dicGenAutoServiceScheule("WorkOrder")) = 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify [ "+dicGenAutoServiceScheule("WorkOrder")+" ] is highlighted.")
								Set objGenAutoSerSchedule = Nothing
								Set objTable = Nothing
								Exit Function
							End If

					End If
				End If

				If sAction = "VerifyWorkOrders" Then 
					If dicGenAutoServiceScheule("WorkOrder") <> "" Then 
						objTable.JavaTab("SearchTab").Select "My Work Order" 
						wait 1
						iRowNumber = objTable.JavaTable("SearchResultTable").GetROProperty("rows")
						arrNodes = Split(dicGenAutoServiceScheule("WorkOrder"),"~")
						For iCnt = 0 To UBound(arrNodes)
							bFlag = False
							For iCounter = 0 To iRowNumber - 1
								sNodeName = Fn_UI_JavaTable_GetCellData("Fn_SISW_GenAutoServiceSchedule_Opearations", objTable, "SearchResultTable",iCounter,0)
								If Instr(1,trim(dicGenAutoServiceScheule("WorkOrder")),trim(sNodeName)) > 0 Then
									bFlag = True
									Exit For
								End If 
							Next
							If bFlag = False Then 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify [ "+dicGenAutoServiceScheule("WorkOrder")+" ] work Orders in [ My Work Order ] tab.")
								Set objGenAutoSerSchedule = Nothing
								Set objTable = Nothing
								Exit function
							End if
						Next
					End If
				End If

				If sAction = "SelectFrmMyWrkOrderPage" Then 
					If dicGenAutoServiceScheule("WorkOrder") <> "" Then
						objTable.JavaTab("SearchTab").Select "My Work Order" 
						wait 1
						bFlag = False
						iRowNumber = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_GenAutoServiceSchedule_Opearations","GetRowIndex",objTable,"SearchResultTable", "", "",dicGenAutoServiceScheule("WorkOrder"),0,"", "", "")
						If iRowNumber <> - 1 Then
							bFlag = Fn_UI_JavaTable_ClickCell("Fn_SISW_GenAutoServiceSchedule_Opearations",objTable,"SearchResultTable",iRowNumber, 0)
						End If
						If bFlag = False Then 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to select [ Work Order ] in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Set objTable = Nothing
							Exit function
						End if
					End if
				End if 

				If sAction <> "AddExistingWorkOrder" and sAction <> "AddExistingWorkOrderAndGenerate" Then
					If dicGenAutoServiceScheule("Button") <> "" Then
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objTable , dicGenAutoServiceScheule("Button"))
						If bFlag = False Then 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to click on [ OK ] in [ Serch ] window.")
							Set objGenAutoSerSchedule = Nothing
							Set objTable = Nothing
							Exit function
						End if
					Else
						Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objTable ,"Cancel")
					End if	
				End If

				If sAction = "AddExistingWorkOrderAndGenerate" Then
					If dicGenAutoServiceScheule("GenerateSchedule") = "True" then
						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","Generate Schedule"
						bFlag=Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objGenAutoSerSchedule,"ShowServicePlan")
						If bFlag = False Then 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Generate Schedule ] button on [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit function
						End if
						Wait 5
					End if
				End If

				Set objTable = Nothing

		'Case to verify Column names in Create Maintenance Action table
		Case "VerifyColNamesInCMATable"
				objGenAutoSerSchedule.JavaWindow("CreateMaintenanceAction").SetTOProperty "title","Create Maintenance Action"
		        Set objNewMaint = objGenAutoSerSchedule.JavaWindow("CreateMaintenanceAction").JavaTable("CreateMaintenanceTable")
				If dicGenAutoServiceScheule("ColumnName") <> "" Then
						aColName = Split(dicGenAutoServiceScheule("ColumnName"),"~")
						iRowCount = objNewMaint.Object.getColumnCount() 
						For iCnt = 0 To UBound(aColName)
							bFlag = False
							For iCount1 = 0 To iRowCount - 1
								If Trim(objNewMaint.GetColumnName(iCount1)) = Trim(aColName(iCnt)) Then
							 		bFlag = True
							 	  	Exit For
							 	End If
							Next
							 
							If bFlag = False Then
							 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to verify existence of column name [ "+aColName(iCnt)+" ] in [ Create Maintenance Action ] window.")
								Set objGenAutoSerSchedule = Nothing
								Set objNewMaint = Nothing
							    Exit Function
							End If 
						Next	
				End If
				'------------- Click Finish/Cancel button -------------------------
				If dicGenAutoServiceScheule("Button") <> "" Then
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule.JavaWindow("CreateMaintenanceAction"),dicGenAutoServiceScheule("Button"))
						Wait 3
						If bFlag = False then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+dicGenAutoServiceScheule("Button")+" ] button in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Set objNewMaint = Nothing
							Exit Function
						End If
				End If
		'for setting (Checked) one or multiple rows which we want in "Create Maintenance Action" table
		'dicGenAutoServiceScheule("GetRowFromValues") = "Name:SrcReq1~Asset:000293/-~Impacted Part:000293/-~Part Position:-"
		'dicGenAutoServiceScheule("SetValuesToRow") = "Due Date:6/21/2016|10:22:23 AM~AutoComplete:ON~Note:TestNote~Button:Finish"
		Case "SetCreateMaintActTableRow"
				If dicGenAutoServiceScheule("GetRowFromValues") <> "" AND dicGenAutoServiceScheule("SetValuesToRow") <> "" Then
						If dicGenAutoServiceScheule("Button")<>"" Then
							objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","New Maintenance Action"
							bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
							wait 1
							If bFlag = false then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Show Service Plan(s) ] button in [ Generate Automated Service Schedule ] window.")
								Set objGenAutoSerSchedule = Nothing
								Exit function
							End if
						End If
						Set objNewMaint = objGenAutoSerSchedule.JavaWindow("CreateMaintenanceAction")
						objNewMaint.SetTOProperty "title","Create Maintenance Action"
						If Fn_UI_ObjectExist("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint) = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to find [ "+objNewMaint.tostring+" ] window.")
							Set objNewMaint = Nothing
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						Else
							Call Fn_Window_Maximize("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint)
							wait 1
						End If
						
						Set dicCreateMainActionDic = CreateObject("Scripting.Dictionary")
						dicCreateMainActionDic.RemoveAll
						aRowToSet = Split(dicGenAutoServiceScheule("GetRowToSet"),"~")
						For iCount = 0 To UBound(aRowToSet)
							aColNameValue = Split(aRowToSet(iCount),":")
							If aColNameValue(0)="Due Date" Then
								aColNameValue(1) = sArr(1)
								If UBound(sArr)>1 Then
									For iTimeCount = 2 To UBound(sArr) Step + 1
										aColNameValue(1) = aColNameValue(1) &":"& sArr(iTimeCount)
									Next
								End If
							End If
							If iCount = 0 Then
								dicCreateMainActionDic("ColumnName") = aColNameValue(0)
								dicCreateMainActionDic("ColumnValue") = aColNameValue(1)
							Else
								dicCreateMainActionDic("ColumnName") = dicCreateMainActionDic("ColumnName") &"~"& aColNameValue(0)
								dicCreateMainActionDic("ColumnValue") = dicCreateMainActionDic("ColumnValue") &"~"& aColNameValue(1)
							End If
						Next
						
						iRowNumber = Fn_SISW_GenAutoServiceSchedule_Opearations("GetCreateMaintActionTableRowNo","",dicCreateMainActionDic,"","")
						If iRowNumber=-1 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Find row in [ Create Maintenance Action ] Table.")
							Set objGenAutoSerSchedule = Nothing
							Set objNewMaint = Nothing
							Set dicCreateMainActionDic = Nothing
							Exit Function
						End If
						Set dicCreateMainActionDic = Nothing
						
						arrNodes = split(dicGenAutoServiceScheule("SetValuesToRow"),"~")
						For iCnt = 0 to uBound(arrNodes)
							sArr = split(arrNodes(iCnt),":")
							Select Case sArr(0)
								Case "Due Date"
									'set Row Number as Index of button
									objNewMaint.JavaButton("DueDate").SetTOProperty "index",iRowNumber
									sDateTime = sArr(1)
									If UBound(sArr)>1 Then
										For iTimeCount = 2 To UBound(sArr) Step + 1
											sDateTime = sDateTime &":"& sArr(iTimeCount)
										Next
									End If								
									ArrDateTime = Split(sDateTime,"|")
									'Date Conversion from 6/21/2016 to 21-Jun-2016 format
									sDate = ArrDateTime(0)
									aDate = Split(sDate,"/")
									sDate = aDate(1) &"-"& MonthName(aDate(0),True) &"-"& aDate(2)
									'Time
									sTime = ArrDateTime(1)
									'Click on "Due Date" Control button
									bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint,"DueDate")
									If bFlag = false then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set [ Due Date ] in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
									Wait 1
									bFlag = Fn_UI_SetDateAndTime("Fn_SISW_GenAutoServiceSchedule_Opearations",sDate,sTime)
									If bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set on [ Due Date ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
									Wait 1
								Case "AutoComplete"
									'set Row Number as Index of CheckBox
									objNewMaint.JavaCheckBox("AutoComplete").SetTOProperty "index",iRowNumber
									bFlag = Fn_CheckBox_Set("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint, "AutoComplete", sArr(1))
									If bFlag = false then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to check [ AutoComplete ] in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
								Case "Note"
									'set Row Number as Index of Edit Box
									objNewMaint.JavaEdit("Note").SetTOProperty "index",iRowNumber
									bFlag = Fn_Edit_Box("Fn_SISW_GenAutoServiceSchedule_Opearations",objNewMaint,"Note",sArr(1))
									If bFlag = false then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set [ Note ] in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
								Case "Button"
									bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint,sArr(1))
									wait 3
									If bFlag = false then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+sArr(1)+" ] button in [ Create Maintenance Action  ] in [ Generate Automated Service Schedule ] window.")
										Set objGenAutoSerSchedule = Nothing
										Set objNewMaint = Nothing
										Exit function
									End if
							End select
						Next
						Set objNewMaint = Nothing
				End If
		'case to Verify or Get Row in "Display Maintenance" table after clicking "Show Maintenance Action" button
		Case "VerifyCreateMaintActionTableRow","GetCreateMaintActionTableRowNo"
				If sAction = "GetCreateMaintActionTableRowNo" Then
						Fn_SISW_GenAutoServiceSchedule_Opearations = -1
				End If
				If dicGenAutoServiceScheule("Button")<>"" Then
						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","New Maintenance Action"
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
						wait 1
						If bFlag = false then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Show Service Plan(s) ] button in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit function
						End if
				End If
				Set objNewMaint = objGenAutoSerSchedule.JavaWindow("CreateMaintenanceAction")
				objNewMaint.SetTOProperty "title","Create Maintenance Action"
				If Fn_UI_ObjectExist("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to find [ "+objNewMaint.tostring+" ] window.")
						Set objNewMaint = Nothing
						Set objGenAutoSerSchedule = Nothing
						Exit Function
				Else
						Call Fn_Window_Maximize("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint)
						wait 1
				End If
				'dicGenAutoServiceScheule("ColumnName") = "Name~Due Date~Auto-complete~Note~Asset~Impacted Part~Part Position"
				'dicGenAutoServiceScheule("ColumnValue") = "SrcReq1~~OFF~~000294/-~000294/-~"
				If dicGenAutoServiceScheule("ColumnName")<>"" AND dicGenAutoServiceScheule("ColumnValue")<>"" Then
						Set objTable = objNewMaint.JavaTable("CreateMaintenanceTable")
						iRowCount = objTable.GetROProperty("rows")
						aColName = Split(dicGenAutoServiceScheule("ColumnName"),"~")
						aColValue = Split(dicGenAutoServiceScheule("ColumnValue"),"~")
						For iCount = 0 To CInt(iRowCount)-1
							For iCount1 = 0 To UBound(aColName)
								bFlag = False
								Select Case aColName(iCount1)
									Case "Due Date"
										sAppText = objTable.GetCellData(iCount,aColName(iCount1))
										If sAppText="No date set" Then
											sAppText = ""
										End If
									Case "Auto-complete"
	'									Set objDesc = Description.Create
	'									objDesc("to_class").value="JavaCheckBox"
	'									Set objChild = objTable.ChildObjects(objDesc)
	'									sAppText = objChild(iCount).getRoProperty("checked")
										objNewMaint.JavaCheckBox("AutoComplete").SetTOProperty "index",iCount
										If objNewMaint.JavaCheckBox("AutoComplete").Exist(1) Then
											sAppText = objNewMaint.JavaCheckBox("AutoComplete").getRoProperty("checked")
										End If
									Case Else
										sAppText = objTable.GetCellData(iCount,aColName(iCount1))
								End Select
								If sAppText<>aColValue(iCount1) Then
									bFlag = False
									Exit For
								End If
								bFlag = True
							Next
							'If Found Row with column values
							If bFlag = True Then
								iRowNumber = iCount
								Exit For
							End If
						Next
						
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Row not Found.")
							Set objTable = Nothing
							Set objNewMaint = Nothing
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						
						If sAction = "GetCreateMaintActionTableRowNo" Then
							If sButtonName <> "" Then
								objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label",sButtonName
								Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_GenAutoServiceSchedule_Opearations", "Click", objGenAutoSerSchedule,"ShowServicePlan")
							End If
							Fn_SISW_GenAutoServiceSchedule_Opearations = iRowNumber
							Set objTable = Nothing
							Set objNewMaint = Nothing
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						Set objTable = Nothing
						Set objNewMaint = Nothing
				End If
		'Case to Verify or Get Row in "Display Maintenance" table after clicking "Show Maintenance Action" button
		'Case to Verify or Get Row in "Service Discreoancies" table after clickin "Show Discrepancies" button
		Case "VerifyDisplayedMaintenanceTableRow","GetDisplayedMaintenanceTableRowNo","VerifyServiceDiscrepanciesTableRow","GetServiceDiscrepanciesTableRowNo"
				If sAction = "GetDisplayedMaintenanceTableRowNo" OR sAction = "GetServiceDiscrepanciesTableRowNo" Then
						Fn_SISW_GenAutoServiceSchedule_Opearations = -1
				End If
				If dicGenAutoServiceScheule("Button")<>"" Then
						If sAction="VerifyDisplayedMaintenanceTableRow" OR sAction="GetDisplayedMaintenanceTableRowNo" Then
							'Click on Show Maintenance Action button
							objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","Show Maintenance Action"
						ElseIf sAction="VerifyServiceDiscrepanciesTableRow" OR sAction="GetServiceDiscrepanciesTableRowNo" Then
							'Click on Show Discrepancies button
							objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","Show Discrepancies"
						End If
						
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Show Maintenance Action ] button in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						Wait 2
				End If
				'dicGenAutoServiceScheule("ColumnName") = "Name~Creation Date~Due Date~Is Overdue~Is Scheduled"
				'dicGenAutoServiceScheule("ColumnValue") = "SR1~21-Jul-2016 11:54~21-Jul-2016 11:54~True~False"
				If dicGenAutoServiceScheule("ColumnName")<>"" AND dicGenAutoServiceScheule("ColumnValue")<>"" Then
						Set objTable = objGenAutoSerSchedule.JavaTable("DisplayedMaintenance")
						iRowCount = objTable.GetROProperty("rows")
						aColName = Split(dicGenAutoServiceScheule("ColumnName"),"~")
						aColValue = Split(dicGenAutoServiceScheule("ColumnValue"),"~")
						bFlag = False
						For iCount = 0 To CInt(iRowCount)-1
							sColIntName = ""
							For iCount1 = 0 To UBound(aColName)
								bFlag = False
								'Use [ objTable.Object.getItem(iCount).getdata.getProperties.tostring() ]
								'to get all Internal names for column for this table
								Select Case aColName(iCount1)
									Case "Name"
										sColIntName = "object_string"
									Case "Creation Date"
										sColIntName = "creation_date"
									Case "Due Date"
										sColIntName = "due_date"
									Case "Is Overdue"
										sColIntName = "ssf0IsOverdue"
									Case "Is Scheduled"
										sColIntName = "ssf0IsScheduled"
									Case "Note"
										sColIntName = "transaction_note"
									Case "Auto-complete"
										sColIntName = "ssf0AutoComplete"
									Case "Asset"
										sColIntName = "ssf0Asset"
									Case "In Progress"
										sColIntName = "InProgress"
									Case "Corrective Action"
										sColIntName = "CorrectiveAction"
									Case "Requirements"
										sColIntName = "SSF0RequirementActions"
									Case "Configuration"
										sColIntName = "SSF0ConfiguresServicePlan"
	'								Case "Part Position"
	'									sColIntName = ""
									Case "Disposition"
										sColIntName = "disposition"
									Case "Approval"
										sColIntName = "approval"
									Case "Discovered By"
										sColIntName = "discovered_by"
									Case "Discovery Date"
										sColIntName = "discovery_date"
									Case "Fault Code"
										sColIntName = "fault_code"
									Case "Severity"
										sColIntName = "severity"
									Case "Is Failure"
										sColIntName = "is_failure"
									Case "Discrepancy Discovered"
										sColIntName = "SSS0DiscrepancyDiscovered"
									Case "Results In"
										sColIntName = "SPI0ResultsIn"
									Case Else
										sColIntName = ""
								End Select
								
								If sColIntName<>"" Then
									sAppText = objTable.Object.getItem(iCount).getdata.getProperty(sColIntName)
									If aColName(iCount1)="Corrective Action" Then
										aColValue(iCount1) = Replace(aColValue(iCount1),"[","")
										aColValue(iCount1) = Replace(aColValue(iCount1),"]","")
										aColValue(iCount1) = Replace(aColValue(iCount1)," ","")
										If sAppText<>aColValue(iCount1) Then
											bFlag = False
											Exit For
										End If
									Else
										If sAppText<>aColValue(iCount1) Then
											bFlag = False
											Exit For
										End If
									End If
									bFlag = True
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Column ["+aColName(iCount1)+"] not Found.")
									Set objTable = Nothing
									Set objGenAutoSerSchedule = Nothing
									Exit Function
								End If
							Next
							'If Found Row with column values
							If bFlag = True Then
								iRowNumber = iCount
								Exit For
							End If
						Next
						
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Row not Found.")
							Set objTable = Nothing
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						
						If sAction = "GetDisplayedMaintenanceTableRowNo" OR sAction = "GetServiceDiscrepanciesTableRowNo" Then
							If sButtonName <> "" Then
								objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label",sButtonName
								Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_GenAutoServiceSchedule_Opearations", "Click", objGenAutoSerSchedule,"ShowServicePlan")
							End If
							Fn_SISW_GenAutoServiceSchedule_Opearations = iRowNumber
							Set objTable = Nothing
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						Set objTable = Nothing
				End If
		'"CheckRowInSDTable" : Case to check the row in Service Discrepancies Table
		'"CheckRowInDMTable" : Case to check the row in Displayed Maintenance Table
		Case "CheckRowInSDTable","CheckRowInDMTable"
				If dicGenAutoServiceScheule("Button")<>"" Then
						'Click on Show Maintenance Action or Show Discrepancies button
						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label",dicGenAutoServiceScheule("Button")
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+dicGenAutoServiceScheule("Button")+" ] button in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						Wait 2
				End If
				If dicGenAutoServiceScheule("ColumnName")<>"" AND dicGenAutoServiceScheule("ColumnValue")<>"" Then
						Set objTable = objGenAutoSerSchedule.JavaTable("DisplayedMaintenance")
						Set dicDispMainDic = CreateObject("Scripting.Dictionary")
						dicDispMainDic.RemoveAll
						dicDispMainDic("ColumnName") = dicGenAutoServiceScheule("ColumnName")
						dicDispMainDic("ColumnValue") = dicGenAutoServiceScheule("ColumnValue")
						If sAction = "CheckRowInSDTable" Then
							sTableName = "Service Discrepancies"
							iRowNumber = Fn_SISW_GenAutoServiceSchedule_Opearations("GetServiceDiscrepanciesTableRowNo","",dicDispMainDic,"","")
						ElseIf sAction = "CheckRowInDMTable" Then
							sTableName = "Displayed Maintenance"
							iRowNumber = Fn_SISW_GenAutoServiceSchedule_Opearations("GetDisplayedMaintenanceTableRowNo","",dicDispMainDic,"","")
						End If
						If iRowNumber <> -1 Then
							If sAction = "CheckRowInSDTable" or sAction = "CheckRowInDMTable"  Then
								sBounds = objTable.Object.getItem(iRowNumber).getBounds().toString()
								sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
								aBounds = split(sBounds,",")
								xCo = cInt(trim(aBounds(0))) - 2
								yCo = cInt(trim(aBounds(1))) + (cInt(trim(aBounds(3)))/2)
								objTable.click xCo, yCo,"LEFT"
							Else
								objTable.Object.getItem(iRowNumber).setChecked("true")
							End If
							Wait 2
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Find row in [ "+sTableName+" ] Table in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Set dicDispMainDic = Nothing
							Set objTable = Nothing
							Exit Function
						End If
						Set dicDispMainDic = Nothing
						Set objTable = Nothing
				End If
		'Case to PopupMenuSelect on Row in [ Display Maintenance Table ]
		Case "PopupMenuSelectDMTRow"
				If dicGenAutoServiceScheule("Button")<>"" Then
						'Click on Show Maintenance Action button
						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label",dicGenAutoServiceScheule("Button")
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Show Maintenance Action ] button in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						Wait 2
				End If
				If dicGenAutoServiceScheule("ColumnName")<>"" AND dicGenAutoServiceScheule("ColumnValue")<>"" Then
						Set objTable = objGenAutoSerSchedule.JavaTable("DisplayedMaintenance")
						Set dicDispMainDic = CreateObject("Scripting.Dictionary")
						dicDispMainDic.RemoveAll
						dicDispMainDic("ColumnName") = dicGenAutoServiceScheule("ColumnName")
						dicDispMainDic("ColumnValue") = dicGenAutoServiceScheule("ColumnValue")
						iRowNumber = Fn_SISW_GenAutoServiceSchedule_Opearations("GetDisplayedMaintenanceTableRowNo","",dicDispMainDic,"","")
						If iRowNumber<>-1 Then
							objTable.DeselectRow iRowNumber
							objTable.DoubleClickCell iRowNumber,"Name","RIGHT"
							Wait 2
							'Select case for PopupMenu select
							Select Case dicGenAutoServiceScheule("PopupMenu")
								Case "CancelMaintenanceAction","CancelMaintenanceActionWithError"
									StrMenu = "Cancel Maintenance Action"
								Case "CompleteMaintenanceAction","CompleteMaintenanceActionWithError","CompleteMaintenanceActionWithConfirmation"
									StrMenu = "Complete Maintenance Action"
								Case else
									StrMenu = dicGenAutoServiceScheule("PopupMenu")
							End Select
							aMenuList = split(StrMenu, ":",-1,1)
							intCount = Ubound(aMenuList)
							Select Case intCount
								Case "0"
									StrMenu = JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
								Case "1"
									StrMenu = JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
								Case "2"
									StrMenu = JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
								Case Else
									Exit Function
							End Select
							JavaWindow("ServiceScheduler").WinMenu("ContextMenu").Select StrMenu
							If Err.Number<>0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Select PopupMenu ["+StrMenu+"] in [ Display Maintenance ] Table in [ Generate Automated Service Schedule ] window.")
								Set objGenAutoSerSchedule = Nothing
								Set dicDispMainDic = Nothing
								Set objTable = Nothing
								Exit Function
							End If
							
							Select Case dicGenAutoServiceScheule("PopupMenu")
								Case "CancelMaintenanceAction", "CancelMaintenanceActionWithError", "CompleteMaintenanceAction", "CompleteMaintenanceActionWithError", "CompleteMaintenanceActionWithConfirmation"
									Set objCancel = objGenAutoSerSchedule.JavaWindow("CreateMaintenanceAction")
									If dicGenAutoServiceScheule("PopupMenu") = "CancelMaintenanceAction" OR dicGenAutoServiceScheule("PopupMenu") = "CancelMaintenanceActionWithError" Then
										objCancel.SetTOProperty "title","Cancel Maintenance Action"	
									ElseIf dicGenAutoServiceScheule("PopupMenu") = "CompleteMaintenanceAction" OR dicGenAutoServiceScheule("PopupMenu") = "CompleteMaintenanceActionWithError" OR dicGenAutoServiceScheule("PopupMenu") = "CompleteMaintenanceActionWithConfirmation" Then
										objCancel.SetTOProperty "title","Complete Maintenance Action"
									End If
									If objCancel.Exist(2) = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Window [Cancel Maintenance Action] does not exist.")
										Set objGenAutoSerSchedule = Nothing
										Set objCancel = Nothing
										Exit Function
									End If
									'Click on Yes or No button on "Cancel Maintenance Action" window
									If dicGenAutoServiceScheule("Button1")<>"" Then
										Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objCancel,dicGenAutoServiceScheule("Button1"))
										Wait 1
									End If
									
									If dicGenAutoServiceScheule("PopupMenu") = "CancelMaintenanceAction" OR dicGenAutoServiceScheule("PopupMenu") = "CompleteMaintenanceAction" Then
										JavaWindow("ServiceScheduler").JavaWindow("MaintenanceActions").SetTOProperty "title","Maintenance Actions"
										If JavaWindow("ServiceScheduler").JavaWindow("MaintenanceActions").Exist(1) Then
											'Click on OK button
											Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",JavaWindow("ServiceScheduler").JavaWindow("MaintenanceActions"),"OK")
										End If
									ElseIf dicGenAutoServiceScheule("PopupMenu") = "CancelMaintenanceActionWithError" OR dicGenAutoServiceScheule("PopupMenu") = "CompleteMaintenanceActionWithError" Then
										If dicGenAutoServiceScheule("PopupMenu") = "CancelMaintenanceActionWithError" Then
											objCancel.SetTOProperty "title","Cancel Maintenance Action Error"
										ElseIf dicGenAutoServiceScheule("PopupMenu") = "CompleteMaintenanceActionWithError" Then
											objCancel.SetTOProperty "title","Complete Maintenance Action Error"
										End If
										If objCancel.Exist(2) = False Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Window [Cancel Maintenance Action Error] does not exist.")
											Set objGenAutoSerSchedule = Nothing
											Set objCancel = Nothing
											Exit Function
										End If
										If dicGenAutoServiceScheule("ErrorMessage") <> "" Then
											sAppText = objCancel.JavaEdit("Note").GetROProperty("value")
											If sAppText<>dicGenAutoServiceScheule("ErrorMessage") Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Verify Error message ["+dicGenAutoServiceScheule("ErrorMessage")+"].")
												Set objGenAutoSerSchedule = Nothing
												Set objCancel = Nothing
												Exit Function
											End If
											'Click on OK button
											Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objCancel,"OK")
										End If
									ElseIf dicGenAutoServiceScheule("PopupMenu") = "CompleteMaintenanceActionWithConfirmation" Then
										objCancel.SetTOProperty "title","Confirmation"
										If objCancel.Exist(2) = False Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Window [Cancel Maintenance Action Error] does not exist.")
											Set objGenAutoSerSchedule = Nothing
											Set objCancel = Nothing
											Exit Function
										End If
										If dicGenAutoServiceScheule("ErrorMessage")<>"" Then
											bResult = Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_GenAutoServiceSchedule_Opearations",objCancel.JavaStaticText("StaticText"),"Label",dicGenAutoServiceScheule("ErrorMessage"))
											If bResult <> False Then
												sAppText = objCancel.JavaStaticText("StaticText").GetROProperty("value")
												If sAppText<>dicGenAutoServiceScheule("ErrorMessage") Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Verify Error message ["+dicGenAutoServiceScheule("ErrorMessage")+"].")
													Set objGenAutoSerSchedule = Nothing
													Set objCancel = Nothing
													Exit Function
												End If
											End If	
											'Click on OK button
											If dicGenAutoServiceScheule("Button2")<>"" Then
												bResult = Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_GenAutoServiceSchedule_Opearations",objCancel.JavaButton("OK"),"Label",dicGenAutoServiceScheule("Button2"))
												If bResult <> False Then
													Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objCancel,"OK")
													Wait 1
												End If
											End If
										End If
									End If
							End Select
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Find row in [ Display Maintenance ] Table in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Set dicDispMainDic = Nothing
							Set objTable = Nothing
							Exit Function
						End If
						Set dicDispMainDic = Nothing
						Set objTable = Nothing
				End If
		'Case to verify Column names in Displayed Maintenance Action table
		Case "VerifyColNamesInDisplayedMATable"
				Set objTable = objGenAutoSerSchedule.JavaTable("DisplayedMaintenance")
				If dicGenAutoServiceScheule("ColumnName") <> "" Then
						aColName = Split(dicGenAutoServiceScheule("ColumnName"),"~")
						iRowCount = objTable.Object.getColumnCount() 
						For iCnt = 0 To UBound(aColName)
						     bFlag = False
							 For iCount1 = 0 To iRowCount - 1
							 	  If Trim(objTable.GetColumnName(iCount1)) = Trim(aColName(iCnt)) Then
							 	  	 bFlag = True
							 	  	 Exit For
							 	  End If
							 Next
							 
							 If bFlag = False Then
							 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to verify existence of column name [ "+aColName(iCnt)+" ] in [ Displayed Maintenance Action ] Table.")
								Set objGenAutoSerSchedule = Nothing
								Set objTable = Nothing
							    Exit Function
							 End If 
						Next	
				End If
		'case to verify "Relate Service Discrepancy and Service Requirement" window
		Case "RelateSDandSROperations"
				If dicGenAutoServiceScheule("Button")<>"" Then
						'Click on [Relate Requirement] button
						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label",dicGenAutoServiceScheule("Button")
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objGenAutoSerSchedule, "ShowServicePlan")
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+dicGenAutoServiceScheule("Button")+" ] button in [ Generate Automated Service Schedule ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						Wait 2
				End If
				If objGenAutoSerSchedule.JavaWindow("RelateSDandSRWindow").Exist(2) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] as [ Relate Service Discrepancy and Service Requirement ] Window does not exist.")
						Set objGenAutoSerSchedule = Nothing
						Exit Function
				End If
				Set objSDSRWindow = objGenAutoSerSchedule.JavaWindow("RelateSDandSRWindow")
				
				'Verify "Service Discrepancy" nodes
				If dicGenAutoServiceScheule("VerifyServiceDiscrepancyNode")<>"" Then
						objSDSRWindow.JavaObject("ServieDorRPanel").SetTOProperty "text","Service Discrepancy"
						aNodes = Split(dicGenAutoServiceScheule("VerifyServiceDiscrepancyNode"),"~")
						For iCount = 0 To UBound(aNodes)
							bFlag = False
							iRowCount = CInt(objSDSRWindow.JavaTable("ServiceDorRTable").GetROProperty("rows"))
							For iCount1 = 0 To iRowCount-1
								sAppText = objSDSRWindow.JavaTable("ServiceDorRTable").Object.getItem(iCount1).getData.tostring
								If sAppText=aNodes(iCount) Then
									bFlag = True
									Exit For
								End If
							Next
							If bFlag = False Then
								Exit For
							End If
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Node ["+aNodes(iCount)+"] does not Exist in [ Service Discrepancy ] table.")
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
				End If
				'Verify "Service Requirement" nodes
				If dicGenAutoServiceScheule("VerifyServiceRequirementNode")<>"" Then
						objSDSRWindow.JavaObject("ServieDorRPanel").SetTOProperty "text","Service Requirement"
						aNodes = Split(dicGenAutoServiceScheule("VerifyServiceRequirementNode"),"~")
						For iCount = 0 To UBound(aNodes)
							bFlag = False
							iRowCount = CInt(objSDSRWindow.JavaTable("ServiceDorRTable").GetROProperty("rows"))
							For iCount1 = 0 To iRowCount-1
								sAppText = objSDSRWindow.JavaTable("ServiceDorRTable").Object.getItem(iCount1).getData.tostring
								If sAppText=aNodes(iCount) Then
									bFlag = True
									Exit For
								End If
							Next
							If bFlag = False Then
								Exit For
							End If
						Next
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Node ["+aNodes(iCount)+"] does not Exist in [ Service Requirement ] table.")
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
				End If
				'Click on [Confirm or Cancel] button
				If dicGenAutoServiceScheule("ConfirmORCancelButton")<>"" Then
						objSDSRWindow.JavaButton("ConfirmORCancel").SetTOProperty "label",dicGenAutoServiceScheule("ConfirmORCancelButton")
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objSDSRWindow, "ConfirmORCancel")
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+dicGenAutoServiceScheule("ConfirmORCancelButton")+" ] button in [ Relate Service Discrepancy and Service Requirement ] window.")
							Set objGenAutoSerSchedule = Nothing
							Exit Function
						End If
						Wait 1
				End If
		'case to Verify active tab in Generate schedule dialog
		Case "VerifyTabActive"
				iRowNumber = objGenAutoSerSchedule.JavaTab("ScheduleTab").Object.getSelectionIndex()
		      	sNodeName = objGenAutoSerSchedule.JavaTab("ScheduleTab").Object.getItem(iRowNumber).tostring()
		      	If Instr(trim(sNodeName),dicGenAutoServiceScheule("TabName")) = 0 Then
			      		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify [ "+dicGenAutoServiceScheule("TabName")+" ] tab is active.")
			      		Set objGenAutoSerSchedule = Nothing
			      		Exit Function
		      	End If
		'case to click any button on [ Generate Automated Service Schedule ] Window
		'Service Tree is above and Service Table is below in window
		Case "ButtonClick","ServiceTreeBtnClick","ServiceTableBtnClick"
				If sButtonName <> "" Then
					If sAction = "ButtonClick" Then
						objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label",sButtonName
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objGenAutoSerSchedule,"ShowServicePlan")
					ElseIf sAction = "ServiceTreeBtnClick" Then
						objGenAutoSerSchedule.JavaButton("BtnServiceTree").SetTOProperty "label",sButtonName
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objGenAutoSerSchedule,"BtnServiceTree")
					ElseIf sAction = "ServiceTableBtnClick" Then
						objGenAutoSerSchedule.JavaButton("BtnServiceTable").SetTOProperty "label",sButtonName
						bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objGenAutoSerSchedule,"BtnServiceTable")
					End If
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+sButtonName+" ] button in [ Generate Automated Service Schedule ] window.")
						Set objGenAutoSerSchedule = Nothing
						Exit Function
					End If
					Wait 2
					Set objGenAutoSerSchedule = Nothing
					Fn_SISW_GenAutoServiceSchedule_Opearations=True
					Exit Function
				End If
				
		'Case to modify Properties on Properties window
		Case "EditMAProperties"
	
				Set objProperties = objGenAutoSerSchedule.JavaWindow("Properties")
				If objProperties.Exist(2) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] as [ Properties ] dialog Does not Exist.")
					Set objProperties = Nothing
					Set objGenAutoSerSchedule = Nothing
					Exit Function
				End If

				Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objProperties, "CheckOutandEdit")
				wait 1
				
				Set objCheckOut = Fn_SISW_GetObject("Check-Out@1")
				
				If dicGenAutoServiceScheule("ChangeID") <> "" Then
					Call Fn_Edit_Box("Fn_SISW_GenAutoServiceSchedule_Opearations", objCheckOut, "ChangeID", dicGenAutoServiceScheule("ChangeID"))
				End If
				
				If dicGenAutoServiceScheule("Comment") <> "" Then
					Call Fn_Edit_Box("Fn_SISW_GenAutoServiceSchedule_Opearations", objCheckOut, "ChangeID", dicGenAutoServiceScheule("Comment"))
				End If
				
				' To click on Yes/No button on Check-Out dialog
				Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objCheckOut, dicGenAutoServiceScheule("CheckOutButton"))
				Set objCheckOut = Nothing
				wait 2
				
				'Select tab in Properties dialog
				If dicGenAutoServiceScheule("TabName")<>"" Then
					objProperties.JavaTab("TabName").Select dicGenAutoServiceScheule("TabName")
					If Err.Number<>0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Select tab ["+dicGenAutoServiceScheule("TabName")+"] in [ Properties ] window.")
						Set objProperties = Nothing
						Set objGenAutoSerSchedule = Nothing
						Exit Function
					End If
				End If	
				
				'Remove dic key value "TabName"
				dicGenAutoServiceScheule.Remove("TabName")
				
				dicCount = dicGenAutoServiceScheule.Count
				dicItems = dicGenAutoServiceScheule.Items
				dicKeys = dicGenAutoServiceScheule.Keys
				
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter), "PropertyRadioButton") > 0 Then
						sSubAction = "PropertyRadioButton"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					
					sValue = dicItems(iCounter)
					
					Select Case sSubAction
						Case "PropertyRadioButton"
							If sValue <> "" Then
								aValue = Split(sValue, ":")
								If Fn_UI_Object_SetTOProperty_ExistCheck("Fn_SISW_GenAutoServiceSchedule_Opearations", objProperties.JavaStaticText("PropertyName"), "label", aValue(0)+":") = True Then
									objProperties.JavaRadioButton("PropertyRadio").SetTOProperty "attached text", aValue(1)
									Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_SISW_GenAutoServiceSchedule_Opearations", "Set", objProperties, "PropertyRadio", "ON")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] as Property field ["+aValue(0)+"] does not exist in [ Properties ] window.")
									Set objProperties = Nothing
									Set objGenAutoSerSchedule = Nothing
									Exit Function
								End If
								aValue = ""
							End If
					End Select
				Next
				
				If dicGenAutoServiceScheule("SaveandCheckIn") <> "" Then
					objProperties.JavaButton("CheckOutandEdit").SetTOProperty "label", "Save and Check-In"
					bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objProperties.JavaButton("CheckOutandEdit"), "")	
					Set objCheckIn = Fn_SISW_GetObject("Check-In@1")
					Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objCheckIn.JavaButton("Yes"), "")	
				End If
				
				Set objCheckIn = Nothing
				Set objProperties = Nothing
			
		'Case to verify Properties of row in [Displayed Maintenance] table in [Maintenance Actions] window
		'PopupMenuSelect on row and select "Properties" menu to open Properties dialog
		Case "VerifyMAProperties"
				If objGenAutoSerSchedule.JavaWindow("Properties").Exist(3) Then
					Set objProperties = objGenAutoSerSchedule.JavaWindow("Properties")
				ElseIf JavaWindow("ServiceScheduler").JavaWindow("Properties").Exist(3) Then   	' For Service Discrepancy tab 
					Set objProperties = JavaWindow("ServiceScheduler").JavaWindow("Properties")			
				End If
				
				If objProperties.Exist(2) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] as [ Properties ] dialog in [ Generate Automated Service Schedule ] window Does not Exist.")
					Set objProperties = Nothing
					Set objMaintActions = Nothing
					Exit Function
				End If
				'Select tab in Properties dialog
				If dicGenAutoServiceScheule("TabName")<>"" Then
					objProperties.JavaTab("TabName").Select dicGenAutoServiceScheule("TabName")
					If Err.Number<>0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to Select tab ["+dicGenAutoServiceScheule("TabName")+"] in [ Generate Automated Service Schedule ] window.")
						Set objProperties = Nothing
						Set objMaintActions = Nothing
						Exit Function
					End If
				End If
				'Remove dic key value "TabName"
				dicGenAutoServiceScheule.Remove("TabName")

				dicCount = dicGenAutoServiceScheule.Count
				dicItems = dicGenAutoServiceScheule.Items
				dicKeys = dicGenAutoServiceScheule.Keys
				
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"PropertyEdit")>0 Then
						sSubAction = "PropertyEdit"
					ElseIf Instr(dicKeys(iCounter),"PropertyImgHyperlink")>0 Then
						sSubAction = "PropertyImgHyperlink"
					ElseIf Instr(dicKeys(iCounter),"PropertyList")>0 Then
						sSubAction = "PropertyList"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					sProperty = dicItems(iCounter)
					bFlag = False
					Select Case sSubAction
						Case "PropertyEdit"
							If sProperty<>"" Then
								aProperty = Split(sProperty,":")
								objProperties.JavaStaticText("PropertyName").SetTOProperty "label",aProperty(0)+":"
								If objProperties.JavaEdit("PropertyEdit").Exist Then
									sAppValue = objProperties.JavaEdit("PropertyEdit").GetROProperty("value")
								End If
								sValue = ""
								If UBound(aProperty)>1 Then
									For iCount = 1 To UBound(aProperty)
										If iCount = 1 Then
											sValue = aProperty(iCount)
										Else
											sValue = sValue +":"+ aProperty(iCount)
										End If
									Next
								Else
									sValue = aProperty(1)
								End If
								If Trim(sAppValue)=Trim(sValue) Then
									bFlag = True
								End If
							End If
						Case "PropertyImgHyperlink"
							If sProperty<>"" Then
								aProperty = Split(sProperty,":")
								objProperties.JavaStaticText("PropertyName").SetTOProperty "label",aProperty(0)+":"
								If objProperties.JavaObject("PropertyImgHyperlink").Exist Then
									sAppValue = objProperties.JavaObject("PropertyImgHyperlink").Object.getText()
								End If
								If Trim(sAppValue)=Trim(aProperty(1)) Then
									bFlag = True
								End If
							End If
						Case "PropertyList"
							If sProperty<>"" Then
								aProperty = Split(sProperty,":")
								aValues = Split(aProperty(1),"~")
								objProperties.JavaStaticText("PropertyName").SetTOProperty "label",aProperty(0)+":"
								If objProperties.JavaList("PropertyList").Exist Then
									iTotalElements = objProperties.JavaList("PropertyList").GetROProperty("items count")
									Set objNewMaint = objProperties.JavaList("PropertyList")
								ElseIf objProperties.JavaApplet("JApplet").JavaList("PropertyList").Exist Then
									iTotalElements = objProperties.JavaApplet("JApplet").JavaList("PropertyList").GetROProperty("items count")
									Set objNewMaint = objProperties.JavaApplet("JApplet").JavaList("PropertyList")
								End If
								For iCount = 0 to Ubound(aValues)
									bFlag = False
									For iCount1 = 0 to iTotalElements-1
										If Trim(objNewMaint.GetItem(iCount1)) = Trim(aValues(iCount)) Then
											bFlag = True
											Exit For
										End If
									Next
									If bFlag <> True Then
										Exit For
									End If
								Next
							End If
						Case "Button"
							If sProperty<>"" Then
								'Click on [Check-Out and Edit] or [Close] button, [CheckOutandEdit] or [Close]
								bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objProperties,sProperty)
								If bFlag = False then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+sProperty+" ] button in [ Properties ] dilaog in [ Generate Automated Service Schedule ] window.")
								End If
							End If
					End Select
					
					If bFlag = False Then
						Fn_SISW_GenAutoServiceSchedule_Opearations = False
						Set objProperties = Nothing
						Set objGenAutoSerSchedule = Nothing
						Set objNewMaint = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Case [ "&sAction&" ] SubCase [ "+sSubAction+" ] - Property Value is not present.")
						Exit Function
					End If
				Next
				Set objProperties = Nothing
				Set objNewMaint = Nothing
		
	    'Case to Set & Verify Manage Due Date
		Case "SetAndVerifyManageDueDate"		
			  
			  'click on manage Due date button
			  objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label","Manage Due Date"
			  bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations",objGenAutoSerSchedule,"ShowServicePlan")
			  If bFlag = False then
				  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ Manage Due Date ] button in [ Generate Automated Service Schedule ] window.")
			  	  Set objGenAutoSerSchedule = Nothing
				  Exit function	
			  End If
			  
			  Set objNewMaint = objGenAutoSerSchedule.JavaWindow("CreateMaintenanceAction")
			  objNewMaint.SetTOProperty "title","Manage Due Date"
			  
			  'Set Date & Time
			  If dicGenAutoServiceScheule("SetDateTime") <> "" Then
			  
			       ArrDateTime = Split(dicGenAutoServiceScheule("SetDateTime"),"~")
				  'Date Conversion from 6/21/2016 to 21-Jun-2016 format
				   sDate = ArrDateTime(0)
				   aDate = Split(sDate,"/")
				   sDate = aDate(1) &"-"& MonthName(aDate(0),True) &"-"& aDate(2)
				  'Time
				  If UBound(ArrDateTime) > 0 Then
				  	 sTime = ArrDateTime(1) 
				  End If
				  
				  'Click on "Due Date" Control button
					bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint,"DueDate")
					If bFlag = false then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set [ Due Date ] in [ Manage Due Date  ] in [ Generate Automated Service Schedule ] window.")
						Set objGenAutoSerSchedule = Nothing
						Set objNewMaint = Nothing
						Exit function
					End if
					Wait 1					  
				    
				    bFlag = Fn_UI_SetDateAndTime("Fn_SISW_GenAutoServiceSchedule_Opearations",sDate,sTime)
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set on [ Due Date ] in [ Manage Due Date ] window.")
						Set objGenAutoSerSchedule = Nothing
						Set objNewMaint = Nothing
						Exit function
					End if
					Wait 2 
					
					 'Verify Date & Time is displayed
					  If dicGenAutoServiceScheule("VerifyDateTime") <> "" Then
					  	  objNewMaint.JavaEdit("Note").SetTOProperty "attached text" ,"Select Date"
					  	  If instr(objNewMaint.JavaEdit("Note").GetROProperty("text"),trim(dicGenAutoServiceScheule("VerifyDateTime"))) = 0 Then
					  	  	    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to verify [ Date : "+dicGenAutoServiceScheule("VerifyDateTime")+"] in [ Manage Due Date ] window.")
								Set objGenAutoSerSchedule = Nothing
								Set objNewMaint = Nothing
								Exit function
					  	  End If 
					  End If
			  End if  
			   
			  'Enter Note
			  If dicGenAutoServiceScheule("Note") <> "" Then
				    objNewMaint.JavaEdit("Note").SetTOProperty "attached text" ,"Note"
				    bFlag = Fn_Edit_Box("Fn_SISW_GenAutoServiceSchedule_Opearations",objNewMaint,"Note",dicGenAutoServiceScheule("Note"))
					wait 1
					If bFlag = false then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to set [ Note ] in [ Manage Due Date ] in [ Generate Automated Service Schedule ] window.")
						Set objGenAutoSerSchedule = Nothing
						Set objNewMaint = Nothing
						Exit function
					End if
			  End if
			  
			  'Click on OK/Cancel button
			  If dicGenAutoServiceScheule("Button") <> "" Then
			        objNewMaint.JavaButton("OK").SetTOProperty "label",dicGenAutoServiceScheule("Button")
				    bFlag = Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint,"OK")
					wait 3
					If bFlag = false then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Failed to clicked on [ "+dicGenAutoServiceScheule("Button")+" ] button in [ Manage Due Date  ] in [ Generate Automated Service Schedule ] window.")
						Set objGenAutoSerSchedule = Nothing
						Set objNewMaint = Nothing
						Exit function
					End if
			  Else
                    objNewMaint.JavaButton("OK").SetTOProperty "label","Cancel"
				    Call Fn_Button_Click("Fn_SISW_GenAutoServiceSchedule_Opearations", objNewMaint,"OK") 			  
			  End if
	          Set objNewMaint = Nothing
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_GenAutoServiceSchedule_Opearations ] Invalid case.")
				Set objGenAutoSerSchedule = Nothing
				Exit Function
	End Select
	
	If sButtonName <> "" Then
		objGenAutoSerSchedule.JavaButton("ShowServicePlan").SetTOProperty "label",sButtonName
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_GenAutoServiceSchedule_Opearations", "Click", objGenAutoSerSchedule,"ShowServicePlan")
	End If
	
	Fn_SISW_GenAutoServiceSchedule_Opearations=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_GenAutoServiceSchedule_Opearations:Successfully created Work Order for Service Schedule in [ Generate Automated Service Schedule ] dialog.")
	Set objGenAutoSerSchedule = Nothing
	Set objTree = Nothing
	Set oCurrentNode = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_SrvScheduler_MaintenanceActions_Ops
'@@
'@@    Description		:	Function Used to perform operations on "Maintenance Actions" window
'@@
'@@    Parameters		:	1. sAction		: Action to be performed
'@@						:	2. sNodeName	: BOM node name in Service Editor on which "PopupSelect" operation need to perform
'@@						:	3. sColNames 	: Column names
'@@						:	4. sColValues	: Column values
'@@						:	5. dicDetails	: Dictionary object
'@@						:	6. sButton		: Button name
'@@						:	7. sReserve		: Future use
'@@
'@@    Return Value		: 	True Or False or Row number in integer format
'@@
'@@    Examples			:   sColNames = "Name~Creation Date~Due Date~Is Overdue~Is Scheduled~Note~Auto-complete~Asset~In Progress~Corrective Action~Requirements~Configuration"
'@@    					:	sColValues = "SrcReq2_5372~01-Aug-2016 16:26~01-Aug-2016 16:25~True~False~Note_869~True~0000173/-~0000173/-~~000182/A;1-SrcReq2_5372~SrcPlan5372"
'@@    					:	bReturn = Fn_SrvScheduler_MaintenanceActions_Ops("VerifyRow","000173/--A (View)",sColNames,sColValues,"","Close","")
'@@    					:	bReturn = Fn_SrvScheduler_MaintenanceActions_Ops("GetRowNumber","000173/--A (View)",sColNames,sColValues,"","Close","")
'@@    					:	bReturn = Fn_SrvScheduler_MaintenanceActions_Ops("PopupMenuSelectOnRow","000061/--A (View)",sColNames,sColValues,"","","Properties")
'@@    					Example for Properties verify	
'@@    					:	Set dicDetails = CreateObject("Scripting.Dictionary")
'@@    					:		dicDetails("TabName") = "All"
'@@    					:		dicDetails("PropertyEdit1") = "Action Type:sem_repair_types"
'@@    					:		dicDetails("PropertyEdit2") = "Creation Date:03-Aug-2016 12:25"
'@@    					:		dicDetails("PropertyImgHyperlink1") = "Asset:000061/-"
'@@    					:		dicDetails("PropertyList1") = "Configuration:SP1"
'@@    					:		dicDetails("PropertyList2") = "In Progress:000061/-"
'@@    					:		dicDetails("Button") = "Close"
'@@    					:	bReturn = Fn_SrvScheduler_MaintenanceActions_Ops("VerifyMAProperties","","","",dicDetails,"Close","")
'@@    							
'@@	   History			:	
'@@		Developer Name		Date	  Rev. No.		Changes Done								Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@		Vivek Ahirrao	18-Dec-2015		1.0		  	Created										[TC1122-20151116d-15_12_2015-VivekA-NewDevelopment]
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_SrvScheduler_MaintenanceActions_Ops(sAction,sNodeName,sColNames,sColValues,dicDetails,sButton,sReserve)
	GBL_FAILED_FUNCTION_NAME="Fn_SrvScheduler_MaintenanceActions_Ops"
	Dim objMaintActions, objTable
	Dim bFlag, sColIntName, sAppText
	Dim iRowCount, iCount, iCount1, iRowNumber
	Dim aColName, aColValue
	
	Fn_SrvScheduler_MaintenanceActions_Ops = False
	
	Set objMaintActions = JavaWindow("ServiceScheduler").JavaWindow("MaintenanceActions")
	If objMaintActions.Exist(2) = False Then
		If sNodeName<>"" Then
			bFlag = Fn_SISW_SrvScheduler_BOMTable_NodeOperation("PopupSelect",sNodeName,"","","Show Maintenance Actions")
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_SrvScheduler_MaintenanceActions_Ops] Failed to perform [RMB : Show Maintenance Actions] on Root Structure Node ["&sNodeName&"].")
				Set objMaintActions = Nothing
				Exit Function
			End If
			Call Fn_ReadyStatusSync(1)
		End If
		If objMaintActions.Exist(2) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [Fn_SrvScheduler_MaintenanceActions_Ops] - [Maintenance Actions] window does not found.")
			Set objMaintActions = Nothing
			Exit Function
		End If
	End If
	
	'Maximize window
	If objMaintActions.GetROProperty("maximized") <> "1" Then
		objMaintActions.Maximize
		wait 2
	End If
	
	Select Case sAction
		'Case to verify Row in [Displayed Maintenance Actions] table
		Case "VerifyRow","GetRowNumber"
				If sColNames<>"" AND sColValues<>"" Then
						Set objTable = objMaintActions.JavaTable("DisplayedMaintenance")
						iRowNumber = -1
						iRowCount = objTable.GetROProperty("rows")
						aColName = Split(sColNames,"~")
						aColValue = Split(sColValues,"~")
						For iCount = 0 To CInt(iRowCount)-1
							sColIntName = ""
							For iCount1 = 0 To UBound(aColName)
								bFlag = False
								'Use [ objTable.Object.getItem(iCount).getdata.getProperties.tostring() ]
								'to get all Internal names for column for this table
								Select Case aColName(iCount1)
									Case "Name"
										sColIntName = "object_string"
									Case "Creation Date"
										sColIntName = "creation_date"
									Case "Due Date"
										sColIntName = "due_date"
									Case "Is Overdue"
										sColIntName = "ssf0IsOverdue"
									Case "Is Scheduled"
										sColIntName = "ssf0IsScheduled"
									Case "Note"
										sColIntName = "transaction_note"
									Case "Auto-complete"
										sColIntName = "ssf0AutoComplete"
									Case "Asset"
										sColIntName = "ssf0Asset"
									Case "In Progress"
										sColIntName = "InProgress"
									Case "Corrective Action"
										sColIntName = "CorrectiveAction"
									Case "Requirements"
										sColIntName = "SSF0RequirementActions"
									Case "Configuration"
										sColIntName = "SSF0ConfiguresServicePlan"
	'								Case "Part Position"
	'									sColIntName = ""
									Case "Disposition"
										sColIntName = "disposition"
									Case "Approval"
										sColIntName = "approval"
									Case Else
										sColIntName = ""
								End Select
								
								If sColIntName<>"" Then
									sAppText = objTable.Object.getItem(iCount).getdata.getProperty(sColIntName)
									If aColName(iCount1)="Corrective Action" Then
										aColValue(iCount1) = Replace(aColValue(iCount1),"[","")
										aColValue(iCount1) = Replace(aColValue(iCount1),"]","")
										aColValue(iCount1) = Replace(aColValue(iCount1)," ","")
										If sAppText<>aColValue(iCount1) Then
											bFlag = False
											Exit For
										End If
									ElseIf aColName(iCount1)="Asset" OR aColName(iCount1)="In Progress" Then
										If Instr(aColValue(iCount1),sAppText)<=0 Then
											bFlag = False
											Exit For
										End If
									Else
										If sAppText<>aColValue(iCount1) Then
											bFlag = False
											Exit For
										End If
									End If
									bFlag = True
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Column ["+aColName(iCount1)+"] not Found.")
									Set objTable = Nothing
									Set objMaintActions = Nothing
									Exit Function
								End If
							Next
							'If Found Row with column values
							If bFlag = True Then
								iRowNumber = iCount
								Exit For
							End If
						Next
						
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Row not Found.")
							Set objTable = Nothing
							Set objMaintActions = Nothing
							Exit Function
						End If
						If sAction = "GetRowNumber" Then
							Fn_SrvScheduler_MaintenanceActions_Ops = iRowNumber
							If sButton<>"" Then
								Call Fn_Button_Click("Fn_SrvScheduler_MaintenanceActions_Ops",objMaintActions,sButton)
							End If
							Set objTable = Nothing
							Set objMaintActions = Nothing
							Exit Function
						End If
						Set objTable = Nothing
						Fn_SrvScheduler_MaintenanceActions_Ops = True
						If sButton<>"" Then
							Call Fn_Button_Click("Fn_SrvScheduler_MaintenanceActions_Ops",objMaintActions,sButton)
						End If
				End If
		'case to Select PopupMenu on Row in [Displayed Maintenance] table in [Maintetanance Actions] window
		Case "PopupMenuSelectOnRow"
				If sColNames<>"" AND sColValues<>"" AND sReserve<>"" Then
						StrMenu = sReserve
						Set objTable = objMaintActions.JavaTable("DisplayedMaintenance")
						'Get row number on which PopupMenu Select operation wants to perform
						iRowNumber = Fn_SrvScheduler_MaintenanceActions_Ops("GetRowNumber","",sColNames,sColValues,"","","")
						If iRowNumber<>-1 Then
							objTable.DeselectRow iRowNumber
							objTable.DoubleClickCell iRowNumber,"Name","RIGHT"
							Wait 2
							aMenuList = split(StrMenu, ":",-1,1)
							intCount = Ubound(aMenuList)
							Select Case intCount
								Case "0"
									StrMenu = JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
								Case "1"
									StrMenu = JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
								Case "2"
									StrMenu = JavaWindow("ServiceScheduler").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
								Case Else
									Exit Function
							End Select
							JavaWindow("ServiceScheduler").WinMenu("ContextMenu").Select StrMenu
							If Err.Number<>0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvScheduler_MaintenanceActions_Ops ] Failed to Select PopupMenu ["+StrMenu+"] in [ Display Maintenance ] Table in [ Maintenance Actions ] window.")
								Set objTable = Nothing
								Set objMaintActions = Nothing
								Exit Function
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvScheduler_MaintenanceActions_Ops ] Failed to Find row in [ Display Maintenance ] Table in [ Maintenance Actions ] window.")
							Set objTable = Nothing							
							Set objMaintActions = Nothing
							Exit Function
						End If
						Set objTable = Nothing
						Fn_SrvScheduler_MaintenanceActions_Ops = True
						If sButton<>"" Then
							Call Fn_Button_Click("Fn_SrvScheduler_MaintenanceActions_Ops",objMaintActions,sButton)
						End If
				End If
		'Case to verify Properties of row in [Displayed Maintenance] table in [Maintenance Actions] window
		'PopupMenuSelect on row and select "Properties" menu to open Properties dialog
		Case "VerifyMAProperties"
				Set objProperties = objMaintActions.JavaWindow("Properties")
				If objProperties.Exist(2) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvScheduler_MaintenanceActions_Ops ] as [ Properties ] dialog in [ Maintenance Actions ] window Does not Exist.")
					Set objProperties = Nothing
					Set objMaintActions = Nothing
					Exit Function
				End If
				'Select tab in Properties dialog
				If dicDetails("TabName")<>"" Then
					objProperties.JavaTab("TabName").Select dicDetails("TabName")
					If Err.Number<>0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvScheduler_MaintenanceActions_Ops ] Failed to Select tab ["+dicDetails("TabName")+"] in [ Maintenance Actions ] window.")
						Set objProperties = Nothing
						Set objMaintActions = Nothing
						Exit Function
					End If
				End If
				'Remove dic key value "TabName"
				dicDetails.Remove("TabName")
				
				dicCount = dicDetails.Count
				dicItems = dicDetails.Items
				dicKeys = dicDetails.Keys
				
				For iCounter = 0 To dicCount - 1
					If Instr(dicKeys(iCounter),"PropertyEdit")>0 Then
						sSubAction = "PropertyEdit"
					ElseIf Instr(dicKeys(iCounter),"PropertyImgHyperlink")>0 Then
						sSubAction = "PropertyImgHyperlink"
					ElseIf Instr(dicKeys(iCounter),"PropertyList")>0 Then
						sSubAction = "PropertyList"
					Else
						sSubAction = dicKeys(iCounter)
					End If
					sProperty = dicItems(iCounter)
					bFlag = False
					Select Case sSubAction
						Case "PropertyEdit"
							If sProperty<>"" Then
								aProperty = Split(sProperty,":")
								objProperties.JavaStaticText("PropertyName").SetTOProperty "label",aProperty(0)+":"
								If objProperties.JavaEdit("PropertyEdit").Exist Then
									sAppValue = objProperties.JavaEdit("PropertyEdit").GetROProperty("value")
								End If
								sValue = ""
								If UBound(aProperty)>1 Then
									For iCount = 1 To UBound(aProperty)
										If iCount = 1 Then
											sValue = aProperty(iCount)
										Else
											sValue = sValue +":"+ aProperty(iCount)
										End If
									Next
								Else
									sValue = aProperty(1)
								End If
								If Trim(sAppValue)=Trim(sValue) Then
									bFlag = True
								End If
							End If
						Case "PropertyImgHyperlink"
							If sProperty<>"" Then
								aProperty = Split(sProperty,":")
								objProperties.JavaStaticText("PropertyName").SetTOProperty "label",aProperty(0)+":"
								If objProperties.JavaObject("PropertyImgHyperlink").Exist Then
									sAppValue = objProperties.JavaObject("PropertyImgHyperlink").Object.getText()
								End If
								If Trim(sAppValue)=Trim(aProperty(1)) Then
									bFlag = True
								End If
							End If
						Case "PropertyList"
							If sProperty<>"" Then
								aProperty = Split(sProperty,":")
								aValues = Split(aProperty(1),"~")
								objProperties.JavaStaticText("PropertyName").SetTOProperty "label",aProperty(0)+":"
								If objProperties.JavaList("PropertyList").Exist Then
									iTotalElements = objProperties.JavaList("PropertyList").GetROProperty("items count")
								End If
								For iCount = 0 to Ubound(aValues)
									bFlag = False
									For iCount1 = 0 to iTotalElements-1
										If Trim(objProperties.JavaList("PropertyList").GetItem(iCount1)) = Trim(aValues(iCount)) Then
											bFlag = True
											Exit For
										End If
									Next
									If bFlag <> True Then
										Exit For
									End If
								Next
							End If
						Case "Button"
							If sProperty<>"" Then
								'Click on [Check-Out and Edit] or [Close] button, [CheckOutandEdit] or [Close]
								bFlag = Fn_Button_Click("Fn_SrvScheduler_MaintenanceActions_Ops",objProperties,sProperty)
								If bFlag = False then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvScheduler_MaintenanceActions_Ops ] Failed to clicked on [ "+sProperty+" ] button in [ Properties ] dilaog in [ Maintenance Actions ] window.")
								End If
							End If
					End Select
					
					If bFlag = False Then
						Fn_SrvScheduler_MaintenanceActions_Ops = False
						Set objProperties = Nothing
						Set objMaintActions = Nothing
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvScheduler_MaintenanceActions_Ops ] Case [ "&sAction&" ] SubCase [ "+sSubAction+" ] - Property Value is not present.")
						Exit Function
					End If
				Next
				Set objProperties = Nothing
				Fn_SrvScheduler_MaintenanceActions_Ops = True
				If sButton<>"" Then
					Call Fn_Button_Click("Fn_SrvScheduler_MaintenanceActions_Ops",objMaintActions,sButton)
				End If
				
				'Case to verify Column names in [ Show Maintenance Action - Displayed Maintenance Action ] table
		Case "VerifyColNamesInMATable"
				If sColNames <> "" Then
					aColName = Split(sColNames,"~")
					sColIntName = Fn_SISW_UI_JavaTable_Operations("Fn_SrvScheduler_MaintenanceActions_Ops","GetAllColumnNames",objMaintActions,"DisplayedMaintenance","","","","","","","")
					For iCount = 0 To UBound(aColName)
						bFlag = False
						If instr(1,cstr(sColIntName),cstr(aColName(iCount))) > 0 Then
							bFlag = True
						End If
						If bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvScheduler_MaintenanceActions_Ops ] Failed to verify existence of column name [ "+aColName(iCount)+" ] in [ Maintenance Actions ] Table.")
							Set objMaintActions = Nothing
							Fn_SrvScheduler_MaintenanceActions_Ops = False
							Exit Function
						End If 
					Next
					Fn_SrvScheduler_MaintenanceActions_Ops = True
					If sButton<>"" Then
						Call Fn_Button_Click("Fn_SrvScheduler_MaintenanceActions_Ops",objMaintActions,sButton)
					End If
				End If
				
		Case Else
				Exit Function
	End Select
	Set objMaintActions = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_SrvScheduler_NewNoticeOperations

'Description			 :	Function Used to Perform operation New Notice Creation dialog

'Parameters			   :  1.sAction : Action name
'									2.sInvokeOption: Dialog Invoke option
'								 	3.sNoticeType: Notice Type
'								    4.sNoticeName: Notice Name
'								    5.sNoticeDescription: Notice Description
'								    6.sOptionNoticeType: Notice Sub Type
'								    7.sButton: Button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Job Card should be selected

'Examples				:   Fn_SISW_SrvScheduler_NewNoticeOperations("Create","Menu", "Notice", "Notice24", "Test","Note","")
'
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Reema W								15-May-2014					1.0																				Paresh D
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public function Fn_SISW_SrvScheduler_NewNoticeOperations(sAction,sInvokeOption, sNoticeType, sNoticeName, sNoticeDescription,sOptionNoticeType,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvScheduler_NewNoticeOperations"
	'Declaring variables
	Dim objNotice,sItemId,sRevId
	Dim arrType, aSPType, sMRUPath, sCmplitListPath
	'Creating Object of  [ New Notice ] dialog
	Set objNotice = JavaWindow("ServiceScheduler").JavaWindow("NewNotice")
	Fn_SISW_SrvScheduler_NewNoticeOperations = False
	
	Select Case lCase(sInvokeOption)
		Case "nooption"
			'Use this option when user wants to invoke New Notice dialog out of this function
		Case "menu",""
			'Performing Action [ File -> New -> Notice...  ] to invoke [ New Notice ]	dialog
			bReturn = Fn_MenuOperation("Select","File:New:Notice...")
			Call Fn_ReadyStatusSync(5)
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to Operate Menu [ File >New > Notice... ] of Function Fn_SISW_SrvScheduler_NewNoticeOperations.")
				Set objNotice = nothing
				Exit Function
			End If
	End Select

	'checking existance of 	[ New Notice ]	dialog
	If objNotice.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_NewNoticeOperations ] Failed to display Notice window.")
			Set objNotice = nothing
			Exit Function
	End If

     Select Case sAction
   		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'case to create new notice
		Case "Create"
			If Trim(sNoticeType) <> "" Then
				If objNotice.JavaTree("NoticeType").Exist(5) Then
					'Selecting Notice Type
					aSPType = Split(sNoticeType,":",-1,1)
					sMRUPath =  "Most Recently Used:" & aSPType(UBound(aSPType))
					sCmplitListPath = "Complete List:" & aSPType(UBound(aSPType))
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SrvScheduler_NewNoticeOperations",objNotice,"NoticeType", "Complete List" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SrvScheduler_NewNoticeOperations",objNotice,"NoticeType", "Complete List")
					End if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SrvScheduler_NewNoticeOperations",objNotice,"NoticeType", "Most Recently Used" , "", "") <> -1 then
						Call Fn_UI_JavaTree_Expand("Fn_SISW_SrvScheduler_NewNoticeOperations",objNotice,"NoticeType", "Most Recently Used")
					end if
					
					If Fn_JavaTree_NodeIndexExt("Fn_SISW_SrvScheduler_NewNoticeOperations",objNotice,"NoticeType", sMRUPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SrvScheduler_NewNoticeOperations", objNotice, "NoticeType",sMRUPath)
					Elseif Fn_JavaTree_NodeIndexExt("Fn_SISW_SrvScheduler_NewNoticeOperations",objNotice,"NoticeType", sCmplitListPath , "", "") <> -1 then
						Call Fn_JavaTree_Select("Fn_SISW_SrvScheduler_NewNoticeOperations", objNotice, "NoticeType",sCmplitListPath)
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_NewNoticeOperations ] Notice Type [ " & UBound(aSPType) & " ] is not present in the List tree.")
						Set objNotice = nothing
						Fn_SISW_SrvScheduler_NewNoticeOperations = False
						Exit function
					End if
					'Clicking on Next button
					If Fn_Button_Click("Fn_SISW_SrvScheduler_NewNoticeOperations", objNotice, "Next")=False Then
						Set objNotice = Nothing
						Exit Function
					End If
				End If
				Call Fn_ReadyStatusSync(5)
			End If
			'Setting Notice name
			If sNoticeName <> "" Then
				If Fn_Edit_Box("Fn_SISW_SrvScheduler_NewNoticeOperations", objNotice,"Name", sNoticeName )=False Then
					Set objNotice = Nothing
					Exit Function
				End If
			End If
			'Setting Notice Description
			If sNoticeDescription <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvScheduler_NewNoticeOperations", objNotice,"Description", sNoticeDescription )
			End If
			'Selecting Notice Type
			If sOptionNoticeType <> "" Then
				Call Fn_Button_Click("Fn_SISW_SrvScheduler_NewNoticeOperations", objNotice, "NoticeType")
				wait 2
				objNotice.JavaTree("Tree").Activate sOptionNoticeType
			End If
			'Clicking On Finish button
			If Fn_Button_Click("Fn_SISW_SrvScheduler_NewNoticeOperations", objNotice, "Finish")=False Then
				Set objNotice = Nothing
				Exit Function
			End IF

			If objNotice.exist(5) then
				'Clicking on Cancel
				Call Fn_Button_Click("Fn_SISW_SrvScheduler_NewNoticeOperations", objNotice, "Cancel")
			End if
			If Err.Number<0 Then
				Fn_SISW_SrvScheduler_NewNoticeOperations =  False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_NewNoticeOperations ] , fail to perform operation [ " & Cstr(Err.Description) & "] ")
			Else
				Fn_SISW_SrvScheduler_NewNoticeOperations =  True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvScheduler_NewNoticeOperations ] Invalid case [ " & sAction & " ].")
			Set objNotice = nothing
			Exit Function
	End Select
	If Fn_SISW_SrvScheduler_NewNoticeOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SrvScheduler_NewNoticeOperations ] Executed successfully with case [ " & sAction & " ].")
	End If
	'Releasing New Notice dialog object
	Set objNotice = Nothing
End Function
