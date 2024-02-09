Option Explicit
iTimeOut = 40

'---------------------------------------------------------	Function List ------------------------------------------------------------------------------------------------------------------------------------------------
'000 Fn_SISW_SrvMgr_GetObject()
'001 Fn_SrvMgr_NavTree_NodeOperation
'002 Fn_SrvMgr_NewPhysicalLocation
'003 Fn_SrvMgr_RootStructureTableColumnOperations
'004 Fn_SrvMgr_RootStructureTableRowIndex
'005 Fn_SrvMgr_RootStructureTableOperations
'006 Fn_SISW_SrvMgr_DispositionCreate
'007 Fn_SISW_SrvMgr_GenerateAsMaintainedStructure
'008 Fn_SISW_SrvMgr_NewAssetGroupCreate
'009 Fn_SISW_SrvMgr_NewCharacteristics
'010 Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations
'011 Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations
'012 Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName
'013 Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations
'014 Fn_SISW_SrvMgr_NewLogBook
'015 Fn_SISW_SrvMgr_RecordUtilization
'016 Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure
'017 Fn_SISW_SrvMgr_UtilizationTabOperations
'018 Fn_SISW_SrvMgr_ContainsPanelOperations
'019 Fn_SISW_SrvMgr_MoveToOperations
'020 Fn_SISW_SrvMgr_SearchOperations
'021 Fn_SISW_SrvMgr_CreateActivityEntryValueOperations
'022 Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations
'023 Fn_SISW_SrvMgr_CreateServiceEventTypeOperations
'024 Fn_SISW_SrvMgr_CreatePartMovementOperations
'025 Fn_SISW_SrvMgr_MaintenanceTreeOperations
'026 Fn_SISW_SrvMgr_SelectPhysicalLocationOperations
'027 Fn_SISW_SrvMgr_CustomerInformationOperations
'028 Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations
'029 Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations
'030 Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations
'031 Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations()
'032 Fn_SISW_SrvMgr_SelectNeutralPart()
'033 Fn_SISW_SrvMgr_LogEntriesOperations()
'034 Fn_SISW_SrvMgr_ServiceOfferingOperations()
'035 Fn_SISW_SrvMgr_CreateRequestedActivityOperations()
'036 Fn_SISW_SrvMgr_TimeAndCostTotalOperations()
'037 Fn_SISW_SrvMgr_DelegateRequestedActivitiesType()
'038 Fn_SISW_SrvMgr_AssignParticipantsOperations()
'039 Fn_SISW_SrvMgr_RevisionRuleSetDate()
'040 Fn_SISW_SrvMgr_SetupUpgrade()
'041.Fn_SISW_SrvMgr_SelectALot()
'****************************************    Function to return required Object ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_GetObject
'
''Description		    :  	Function to get objects of Service Manager

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_SrvMgr_GetObject("GenerateAsMaintainedStructure")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 25-June-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_SrvMgr_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ServiceManager.xml"
	Set Fn_SISW_SrvMgr_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'*********************************************************		Function to action perform on NavTree of  Service Manager ***********************************************************************
'Function Name		:				Fn_SrvMgr_NavTree_NodeOperation

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

'Pre-requisite			:		 		Service Manager module window should be displayed

'Examples				:				  Fn_SrvMgr_NavTree_NodeOperation("PopupMenuSelect","Home:Newstuff","Copy Ctrl+C")
'										            EXAMPLE for Case "Select" : Call Fn_SrvMgr_NavTree_NodeOperation( "Select" ,  "Home:Newstuff:000032-CarModel_VI_LS1:000032 @2" , "" ) 
'										          Call Fn_SrvMgr_NavTree_NodeOperation( "Select" ,  "Home:Newstuff:000032-CarModel_VI_LS1:000032" , "" ) 
'History					 :		
'	Developer Name				Date						Rev. No.			Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Rupali Palhade			26/09/2011			          1.0				      Created
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe			25/06/2012			        2.0				      Modified code to get NodePath
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_SrvMgr_NavTree_NodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SrvMgr_NavTree_NodeOperation"
	Dim intCount, aMenuList
	Dim objJavaWindowMyTc, objJavaTreeNav,ArrNodeName
	Dim sPath, sEle,arr
	Fn_SrvMgr_NavTree_NodeOperation = FALSE
	Set objJavaWindowMyTc = JavaWindow("ServiceManager")
	Set objJavaTreeNav = objJavaWindowMyTc.JavaTree("NavTree")

	Select Case StrAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
					sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SrvMgr_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					If sPath <> False Then
						objJavaTreeNav.Select sPath
						Fn_SrvMgr_NavTree_NodeOperation = True
					End If
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------

		Case "Deselect"
					sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SrvMgr_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					If sPath <> False Then
						objJavaTreeNav.Deselect sPath
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deselected Node [" + StrNodeName + "] of NavTree")
						Fn_SrvMgr_NavTree_NodeOperation = True
					End If
		'----------------------------------------------------------------------- For selecting multiple node at a time -------------------------------------------------------------------------
		Case "Multiselect"
					Set objJavaTreeNav = JavaWindow("ServiceManager").JavaTree("NavTree")
					sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SrvMgr_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					If sPath <> False Then
						Call Fn_UI_JavaTree_ExtendSelect("Fn_SrvMgr_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree", sPath)
						Fn_SrvMgr_NavTree_NodeOperation = TRUE
					End If

		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
					sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SrvMgr_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
					If sPath <> False Then
						objJavaTreeNav.Expand sPath
						Fn_SrvMgr_NavTree_NodeOperation = True
					End If
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
			sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SrvMgr_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
			If sPath <> False Then
				objJavaTreeNav.Collapse sPath
				Fn_SrvMgr_NavTree_NodeOperation = True
			End If
		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuSelect"
			sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SrvMgr_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
			If sPath <> False Then
					'Select node
                    Call Fn_JavaTree_Select("Fn_SrvMgr_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree",sPath )
					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_SrvMgr_NavTree_NodeOperation",objJavaWindowMyTc,"NavTree",sPath)
					wait 3
					Fn_SrvMgr_NavTree_NodeOperation = Fn_UI_JavaMenu_Select("Fn_SrvMgr_NavTree_NodeOperation",objJavaWindowMyTc,StrMenu)
			End If		
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "DoubleClick"
			sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SrvMgr_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
			If sPath <> False Then
				JavaWindow("ServiceManager").JavaTree("NavTree").Activate sPath 
				Fn_SrvMgr_NavTree_NodeOperation = TRUE
			End If
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "Exist"
				sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SrvMgr_NavTree_NodeOperation", objJavaTreeNav, StrNodeName, "", "")
				If sPath <> False Then
					Fn_SrvMgr_NavTree_NodeOperation = TRUE
				End If
         '----------------------------------------------------------------------- For  Select Range of  Nav tree -------------------------------------------------------------------------
		Case "SelectRange"
			ReDim ArrNodeName(2)
					ArrNodeName = Split(StrNodeName,"|")
					JavaWindow("ServiceManager").JavaTree("NavTree").SelectRange ArrNodeName(0),ArrNodeName(1)
					If err.number < 0 Then
						Fn_SrvMgr_NavTree_NodeOperation = False
					else
						Fn_SrvMgr_NavTree_NodeOperation = True
					End If

'- - - - - - - - - - - -  Retruns All Childs of any given Node in the tree in form of an array - - - - - - - - - - - - - - -
				Case "GetChildrenList"
						sReturn=""
						If Fn_SrvMgr_NavTree_NodeOperation("Expand",StrNodeName,"")=True Then
							arrStrNode = Split (StrNodeName, ":")
							If UBound(arrStrNode)=0 Then
								Set oCurrentNode = JavaWindow("ServiceManager").JavaTree("NavTree").Object.getItem(0)
								intNodeCount = oCurrentNode.getItemCount()
								For iCount=0 To intNodeCount-1
									If iCount=0 Then
										sReturn=oCurrentNode.getItem(iCount).getData().toString()
									Else
										sReturn=sReturn+","+oCurrentNode.getItem(iCount).getData().toString()
									End If
								Next
								arr = Split(sReturn,",")
								Fn_SrvMgr_NavTree_NodeOperation = arr
								Set oCurrentNode=Nothing
								Exit Function
							Else
								Set oCurrentNode = JavaWindow("ServiceManager").JavaTree("NavTree").Object.getItem(0)
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
								Fn_SrvMgr_NavTree_NodeOperation = arr
								Set oCurrentNode=Nothing
							End If
						Else
							Fn_SrvMgr_NavTree_NodeOperation = False
						End If
		'****************************************************************************************	
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail :[ Fn_SrvMgr_NavTree_NodeOperation ] Invalid case [ " & StrAction &" ].")
				Exit function
	End Select

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), StrAction &" Sucessfully completed on Node [" + StrNodeName + "] of JavaTree of function Fn_SrvMgr_NavTree_NodeOperation")
	Set objJavaWindowMyTc = nothing
	Set objJavaTreeNav = nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_SrvMgr_NewPhysicalLocation()
'###
'###    DESCRIPTION     :   Create a new Physical Location
'###
'###    PARAMETERS      :   1. sAction
'###						2.	sPhyLocType:
'###						3.	sID:
'###						4.	sRevision
'###						5.	sLocationName
'###						6.	sLocationType
'###						7.	sDescription
'###
'###    Function Calls  :   Fn_WriteLogFile() 
'###
'###	HISTORY         :   AUTHOR            			DATE                VERSION		CHANGES
'###
'###    CREATED BY     	:   Amit Talegaonkar       	26 - Sep - 2011         1.0
'###
'###    MODIFIED BY     :   Koustubh Watwe			29 - Aug - 2012			1.1			Modified code to select Location Type
'###
'###    EXAMPLE         : 	Call Fn_SrvMgr_NewPhysicalLocation( "New" , "Complete List:Physical Location" , "000555" , "A" , "TestLocation" , "Region, Region" , "Description")
'###
'#############################################################################################################
Public Function Fn_SrvMgr_NewPhysicalLocation( sAction , sPhyLocType , sID , sRevision , sLocationName , sLocationType , sDescription )
		GBL_FAILED_FUNCTION_NAME="Fn_SrvMgr_NewPhysicalLocation"
		Dim objNewBusinessObject , ArrProp , ArrVal , intg , aType , sExpand
		Dim objSelectType, iItemCnt, objDialog, iCnt, bFlag
		Dim i, sChar

       Set objNewBusinessObject = JavaWindow("ServiceManager").JavaWindow("NewBusinessObject")
	   Fn_SrvMgr_NewPhysicalLocation = False
	     
	   'If dialog does not exist, invoke Menu - [ File->New->Physical Location ]
	   	If Fn_UI_ObjectExist("Fn_SrvMgr_NewPhysicalLocation",objNewBusinessObject) = False Then
	   		Call Fn_MenuOperation("Select", "File:New:Physical Location...")
			Call Fn_ReadyStatusSync(2)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully invoked menu - [ File:New:Physical Location... ]")
		End If
		
		'Check if it is open now
		If Fn_UI_ObjectExist("Fn_SrvMgr_NewPhysicalLocation",objNewBusinessObject) = False Then
		   'Exit function
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_SrvMgr_NewPhysicalLocation ] Failed to Open [ NewBusinessObject ] dialog for new Physical Location")
			Set objEditQuan = nothing
			Exit Function
		End If

		Select Case sAction

			Case "New"
			
				 'If Location Type is Blank then Physical Location can not be created
				If sLocationType = "" Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_SrvMgr_NewPhysicalLocation ] Failed since Location Type is NOT specified")
					Exit Function
				End If

				If objNewBusinessObject.JavaTree("PhysicalLocationType").Exist(3) Then
						'Expand Nodes in tree
						aType = Split( sPhyLocType , ":")
			
						For intg = 0 to uBound( aType ) - 1
								If intg = 0 Then
									sExpand = aType(intg)
									Call Fn_UI_JavaTree_Expand( "Fn_SrvMgr_NewPhysicalLocation" , objNewBusinessObject, "PhysicalLocationType" , sExpand )
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully Expanded node - [ "+ sExpand +" ] in [ PhysicalLocationType ] tree")
								Else
									sExpand = sExpand + ":" + aType(intg)
									Call Fn_UI_JavaTree_Expand( "Fn_SrvMgr_NewPhysicalLocation" , objNewBusinessObject, "PhysicalLocationType" , sExpand )
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully Expanded node - [ "+ sExpand +" ] in [ PhysicalLocationType ] tree")
								End If
						  Next
					
						'Select [ Physical Location ]
						Call Fn_JavaTree_Select( "Fn_SrvMgr_NewPhysicalLocation", objNewBusinessObject, "PhysicalLocationType", sPhyLocType )
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully selected - [ "+ sPhyLocType +" ] in [ PhysicalLocationType ] tree") 
						Call Fn_ReadyStatusSync(2)
								
						'Click on NEXT button to navigate ahead
						Call Fn_Button_Click("Fn_SrvMgr_NewPhysicalLocation", objNewBusinessObject, "Next")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully Clicked on [ Next ] Button") 
						Call Fn_ReadyStatusSync(2)
				End If
			
				'Enter Details
                ArrProp = Array("ID" , "Revision" , "Location" ,"LocTypEdit" , "Description" )
				ArrVal = Array(sID , sRevision , sLocationName , sLocationType , sDescription )
			
					For intg  = 0 to 4
						If ArrVal(intg) <> "" Then
							Select Case ArrProp(intg)
								Case "LocTypEdit"
'										Call Fn_Button_Click("Fn_SrvMgr_NewPhysicalLocation", objNewBusinessObject, "btnLocationType")
'										wait 1
'										Set objSelectType = description.Create()
'										objSelectType("Class Name").value = "JavaTable"
'										objSelectType("Class Name").value = "JavaTable"
'										Set objDialog = objNewBusinessObject.ChildObjects(objSelectType)
'										Set objTable = objDialog(objDialog.Count -1)
'										iItemCnt = cInt(objTable.GetROProperty("rows"))
'										bFlag = False
'										For iCnt = 0 to iItemCnt - 1
'											If objTable.GetCellData(iCnt,0) =  ArrVal(intg) Then
'												objTable.ClickCell iCnt,0
'												bFlag = True
'												Exit For
'											End If
'										Next
'										If bFlag = False Then
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SrvMgr_NewPhysicalLocation ] Failed to select [ Viewed Partition Scheme = " & sViewedPartitionScheme & " ] in [ New Business Object ] window.")
'											Exit function
'										End If
'										Set objDialog = Nothing
'										Set objTable = Nothing

										ArrVal(intg) = Left(ArrVal(intg), len(ArrVal(intg)) - instr(ArrVal(intg), ",")-1)
										JavaWindow("ServiceManager").JavaWindow("NewBusinessObject").JavaButton("btnLocationType").Click micLeftBtn
										JavaWindow("ServiceManager").JavaWindow("NewBusinessObject").JavaWindow("Shell").JavaTree("LocationType").Activate ArrVal(intg)
'										Call Fn_UI_EditBox_Type("Fn_SrvMgr_NewPhysicalLocation",objNewBusinessObject, ArrProp(intg), ArrVal(intg) )
'										wait(1)
'										'objNewBusinessObject.JavaEdit(ArrProp(intg)).Activate
'										'Added by Vallari on 23-Jan-2013 : TO set focus to some other edit box
'										objNewBusinessObject.JavaEdit(ArrProp(intg)).PressKey micF1
								Case "Location"
										For i = 1 to Len(ArrVal(intg))
											sChar = mid(ArrVal(intg), i, 1)
											If Asc(sChar) = 95 Then
												objNewBusinessObject.JavaEdit(ArrProp(intg)).PressKey "_", micShift
											Else
												objNewBusinessObject.JavaEdit(ArrProp(intg)).Type Chr(Asc(sChar))
												objNewBusinessObject.JavaEdit(ArrProp(intg)).Set Chr(Asc(sChar))
											End If
										Next
										'Following piece of loop is written for typing "0" in the EditBox
										If trim(objNewBusinessObject.JavaEdit(ArrProp(intg)).GetROProperty("value")) <> TRIM(ArrVal(intg)) Then
											objNewBusinessObject.JavaEdit(ArrProp(intg)).Object.setText ArrVal(intg)
										End If
                                Case "ID"
										  Call Fn_Edit_Box("Fn_SrvMgr_NewPhysicalLocation",objNewBusinessObject, ArrProp(intg), ArrVal(intg) )
										  objNewBusinessObject.JavaEdit(ArrProp(intg)).Type "a"
										  Set WshShell = CreateObject("WScript.Shell")
										  WshShell.SendKeys "{BKSP}"
										  wait(1)
										  Set WshShell = nothing
								Case Else
										Call Fn_UI_EditBox_Type("Fn_SrvMgr_NewPhysicalLocation",objNewBusinessObject, ArrProp(intg), ArrVal(intg) )
							End Select
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully Entered " + ArrProp(intg) + " - [ "+ ArrVal(intg) +" ]")
							Call Fn_ReadyStatusSync(2)
						End If
					Next
					
				'Click on FINISH button
				objNewBusinessObject.JavaButton("Finish").WaitProperty "enabled", 1, 60000
				Call Fn_Button_Click("Fn_SrvMgr_NewPhysicalLocation", objNewBusinessObject, "Finish")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully Clicked on [ Finish ] Button") 
				Call Fn_ReadyStatusSync(3)
				
				'Click on CANCEL button
				Call Fn_Button_Click("Fn_SrvMgr_NewPhysicalLocation", objNewBusinessObject, "Cancel")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully Clicked on [ Cancel ] Button") 
				Call Fn_ReadyStatusSync(3)
					
				Fn_SrvMgr_NewPhysicalLocation = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] executed successfully.")

			Case Else
				Fn_SrvMgr_NewPhysicalLocation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] wrong case [ "+ sAction +" ] passed as argument")
			End Select

	Set objNewBusinessObject = nothing
	Set objSelectType = Nothing
	Set objDialog = Nothing
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''****************************************    Function to perform operatiosn for service Manager  Table ***************************************
'
''Function Name		 	:			  Fn_SrvMgr_RootStructureTableColumnOperations
'
''Description		    :  	      Function to perform operatiosn fon columns in service Manager 
'
''Parameters		    :	 	1. sAction : Action need to perform
'					   			            2. sColumnNames : Column's name
'								
''Return Value		    :  		True \ False | column nhumber \ -1
'
''Pre-requisite		    :		Service Manager perspective should be selected

''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableColumnOperations("GetIndex", "Lot")
''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableColumnOperations("Remove", "Lot")
''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableColumnOperations("Add", "Lot")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'					    Rupali Palhade 		27-Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SrvMgr_RootStructureTableColumnOperations(sAction, sColumnNames)
	GBL_FAILED_FUNCTION_NAME="Fn_SrvMgr_RootStructureTableColumnOperations"
	Dim iColIndex, objTable, objParentObject, strMenu, iCols, ArrCol
	Dim sColToAdd, iIndex, objChangeColumnDialog, objList,intCol
	Dim aColumns, bReturn, dicColumnMgnt
	Fn_SrvMgr_RootStructureTableColumnOperations = False
	Set objParentObject = JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet")
	Set objTable = JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet").JavaTable("RootStructuresTable")
'	JavaWindow("ServiceManager").JavaObject("ServiceEditor").Click 5,5,"LEFT"

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "GetIndex"
				iCols = Cint(objTable.GetROProperty("cols"))
				Fn_SrvMgr_RootStructureTableColumnOperations = -1
				For iColIndex =0 to iCols - 1
					If objTable.GetColumnName(iColIndex) = sColumnNames Then
						Fn_SrvMgr_RootStructureTableColumnOperations = iColIndex
						Exit for
					End If
				Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Add"
                 ArrCol = Split(sColumnNames,":",-1,1)
				sColToAdd = ""
				 For iIndex = 0 To Ubound(ArrCol)
						'Check that Column is present in the BOMTable.
						iColIndex =  Fn_SrvMgr_RootStructureTableColumnOperations("GetIndex", ArrCol(iIndex))		
						If iColIndex = -1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Warning: Column does not  exist in the Application.Need to Add Column ["& ArrCol(iIndex) &"]." )
								sColToAdd = sColToAdd +":"+ArrCol(iIndex)
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Column ["& ArrCol(iIndex) &"] exists in the Application" )
								Fn_SrvMgr_RootStructureTableColumnOperations =TRUE
						End if
				Next
				If sColToAdd <>""  Then
						sColToAdd = Mid(sColToAdd, 2,Len(sColToAdd))
						ArrCol = Split(sColToAdd,":",-1,1)
						'Invoke Choose Column Window if it is not present on the screen
						Set objChangeColumnDialog = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						If NOT objChangeColumnDialog.Exist( 1)  Then
								objTable.SelectColumnHeader "#1","RIGHT"       	
								objParentObject.JavaMenu("label:=Insert column\(s\) ...").Select 										       
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
						Call Fn_Button_Click("Fn_SrvMgr_RootStructureTableColumnOperations",objChangeColumnDialog, "Add")
						' Hit  Apply Button after selection
						Call Fn_Button_Click("Fn_SrvMgr_RootStructureTableColumnOperations",objChangeColumnDialog, "Apply")
						Call Fn_Button_Click("Fn_SrvMgr_RootStructureTableColumnOperations",objChangeColumnDialog, "Cancel")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass:Successfully Added  Column  ["& sColToAdd &"] in BOMTable")									
						Fn_SrvMgr_RootStructureTableColumnOperations = TRUE					
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "Remove"
					ArrCol = Split(sColumnNames,":",-1,1)
					For iIndex = 0 To Ubound(ArrCol)										
							'Check that Column is present in the BOMTable
							iColIndex =  Fn_SrvMgr_RootStructureTableColumnOperations("GetIndex", ArrCol(iIndex))						
							If iColIndex = -1 Then							
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"WARNING:Column dose not  exist in the Application.No Need to Remove Column ["& ArrCol(iIndex) &"]")
									Fn_SrvMgr_RootStructureTableColumnOperations  = FALSE
							Else
								'Remove the given Colum.													
								objTable.SelectColumnHeader iColIndex,"RIGHT"
								objParentObject.JavaMenu("label:=Remove this column").Select		
								Call Fn_Button_Click("Fn_SrvMgr_RootStructureTableColumnOperations",JavaWindow("DefaultWindow").JavaWindow("RemoveColumn"),"Yes")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Pass: Successfully removed Column  ["& ArrCol(iIndex) &"] from BOMTable.")          																
								Fn_SrvMgr_RootStructureTableColumnOperations  =TRUE										 						
							End if
					Next

			Case "AddAndMoveToIndex"

					aColumns = Split(sColumnNames,"~",-1,1)
					bReturn = Fn_SrvMgr_RootStructureTableColumnOperations("Add", aColumns(0))
					If bReturn = False Then
						Fn_SrvMgr_RootStructureTableColumnOperations  = FALSE
						Exit Function
					else
						Set objChangeColumnDialog = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Change Columns")
						If NOT objChangeColumnDialog.Exist( 1)  Then
								objTable.SelectColumnHeader "#1","RIGHT"       	
								objParentObject.JavaMenu("label:=Insert column\(s\) ...").Select 										       
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: RMB action Insert Column(s).... Executed successfully in the Application.")			
								Set objList = objChangeColumnDialog.JavaList("ListAvailableCols").Object
						End If
'						bReturn = Fn_SISW_LoadLibrary(Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")
'						If bReturn = False  Then
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to load Preference Library : " & Environment.Value("sPath") & "\Library\RAC_CommonFunctions.vbs")	
'							Fn_SrvMgr_RootStructureTableColumnOperations = False
'							Exit Function
'						End If
						Set dicColumnMgnt = CreateObject("Scripting.Dictionary")
						dicColumnMgnt("Columns") = aColumns(0)+":"+aColumns(1)
						dicColumnMgnt("CloseDialog") = True
						Fn_SrvMgr_RootStructureTableColumnOperations = Fn_SISW_RAC_Common_TableColumnManagement( "MoveColumnToIndex" , dicColumnMgnt )
					End If

		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select
	Set objParentObject = Nothing
	Set objTable = Nothing
	Set objChangeColumnDialog = Nothing
	Set objList = Nothing
End Function

''****************************************    Function to get Row INdex from service Manger Root Table **************************************
'
''Function Name		 	:			  Fn_SrvMgr_RootStructureTableRowIndex
'
''Description		    :  	      Function to to get Row INdex in Service Manager 
'
''Parameters		    :	 	1. objTable : Action need to perform
'					   			2. sNodeName : Root Structure Node Path
'								
''Return Value		    :  		row nhumber \ -1
'
''Pre-requisite		    :		MRO perspective should be selected

''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableRowIndex(objTable, "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Rupali Palhade	     27Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Ashwini		  		20-Mar-2014				2.1				Code modified to identify the QTP version
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Nitish S		  	26-Jun-2014				2.1				'added code to get RowIndex using getPathForRow method as objComponent.getProperty() method returns empty in some case(26-Jun-2014)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_SrvMgr_RootStructureTableRowIndex"
	Dim nodeArr, aRowNode, iColIndex, aPath
	Dim iRowCounter, sNode, iInstance, iNodeCounter, iPathCounter, bFound 
	Dim iRows, sNodePath, sPath, StrNodePath,objComponent
	Dim Iterator
	sPath = ""

	If Fn_UI_ObjectExist("Fn_SrvMgr_RootStructureTableRowIndex", objTable) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvMgr_RootStructureTableRowIndex ] Table does not exist.")	
		Fn_SrvMgr_RootStructureTableRowIndex = -1
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
							sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
							If trim(sNodePath) = trim(aRowNode(0)) then
			                        Set objComponent = objTable.object.getComponentForRow(iRowCounter)
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
											If objComponent Is Nothing Then
												Exit do
											End If
										End IF
									Loop
'									StrNodePath =objTable.Object.getPathForRow(iRowCounter).toString()
'									StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
'									StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
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
                    If objTable.object.getPathForRow(iRowCounter).getLastPathComponent().getClass().toString() <> "class com.teamcenter.rac.treetable.HiddenSiblingNode" Then
						sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
					Else
						sNodePath = ""
					End If
'					sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
					If trim(sNodePath) = trim(aRowNode(0)) then
						If iRows > 1 Then
							  If isObject(ObjTable.object.getComponentForRow(iRowCounter)) Then
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
												If objComponent Is Nothing Then
													Exit do
												End If
											End IF
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
									End if
							  Else
							 	 iRowCounter = -1
							 	 Exit for								
							  End If
						Else ''' added code to get row index of root node as getComponentForRow method not supported when there is single node in table
							sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
							If trim(sNodePath) = trim(aRowNode(0)) then	
								bFound = True
'								iRowCounter = iRowCounter
								Exit do								
							End If
						End If
					End if
				End If
				iRowCounter = iRowCounter + 1
				' increment counter
			loop
		Next
		If iRowCounter = -1 Then  ''added code to get RowIndex using getPathForRow method as objComponent.getProperty() method returns empty in some case(26-Jun-2014)
		   For Iterator = 0 To iRows-1
		      StrNodePath =objTable.Object.getPathForRow(Iterator).toString()
			  StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
			  StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
			  StrNodePath = replace(StrNodePath, ",", ":")
			  StrNodePath = replace(StrNodePath, " ", "")
			  If trim(sNodeName) = trim(StrNodePath) then	
				bFound = True
				iRowCounter =Iterator 
				Exit for
			  End If
			Next
		End If
	End If
	If bFound Then
				Fn_SrvMgr_RootStructureTableRowIndex = iRowCounter
	Else
				Fn_SrvMgr_RootStructureTableRowIndex = -1
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SrvMgr_RootStructureTableRowIndex ] executed successfully.")
	 set objComponent =Nothing
End Function

''****************************************    Function to perform operations on Root Structure Table in Service Manager***************************************
'
''Function Name		 	:			  Fn_SrvMgr_RootStructureTableOperations
'
''Description		    :  	      Function to perform operations on Root Structure Table in Service Manager 
'
''Parameters		    :	 	1. sAction : Action need to perform
'					   			2. sRootStructureHeader : Root Structure header tab label
'					   			3. sRootStructure : Root Structure item
'					   			4. sNodeName : Root Structure Node Path
'					   			5. sColName : column name
'					   			6. sValue : value 
'					   			7. sPopupMenu : Popup menu to select
'								
''Return Value		    :  		row nhumber \ -1
'
''Pre-requisite		    :		Service Manager perspective should be selected

''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableOperations(sAction, "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableOperations("Select", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableOperations("Exist", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableOperations("Expand", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")
''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableOperations("ExpandBelow", "", "", "000554/A;1-TopPart (View)", sColName, sValue, "")
''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableOperations("PopupSelect", "", "", "000554/A;1-TopPart (View)", sColName, sValue, "")
''Examples		     	:	Call  Fn_SrvMgr_RootStructureTableOperations("CellEdit", "", "", "000554/A;1-TopPart (View):000555/A;1-Ch1 (View)", "", "", "")

'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						 Rupali Palhade		 27-Sept-2011			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_SrvMgr_RootStructureTableOperations(sAction, sRootStructureHeader, sRootStructure, sNodeName, sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SrvMgr_RootStructureTableOperations"
	Dim iRowIndex, aMenu, objTable, objParentObject, strMenu, iColIndex,objExpandbelow

	Fn_SrvMgr_RootStructureTableOperations = False
	
	
	Set objParentObject = JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet")
	Set objTable = JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet").JavaTable("RootStructuresTable")
'		If objTable.Exist(1) Then
'			objTable.Click 0,0,"LEFT"
'		End If
	
	'JavaWindow("ServiceManager").JavaObject("ServiceEditor").Click 0,0, "LEFT" 
	' selcting Root Structure Header.
	If sRootStructureHeader <> "" then   ''replace "(" and ")" by "\(" and "\)" 
		If instr(sRootStructureHeader, "(")> 0 and instr(sRootStructureHeader, "(")> 0 and instr(sRootStructureHeader, "\(")= 0 and instr(sRootStructureHeader, "\)")= 0 Then
			sRootStructureHeader = replace(sRootStructureHeader, "(", "\(")
			sRootStructureHeader = replace(sRootStructureHeader, ")", "\)")
		End If
		objParentObject.JavaStaticText("RootStructureHeader").SetTOProperty "label", sRootStructureHeader
		objParentObject.JavaStaticText("RootStructureHeader").Click 1,1,"LEFT"
	End IF

	' selecting Root Structure from List.
	If sRootStructure <> "" Then
		Call Fn_List_Select("Fn_SrvMgr_RootStructureTableOperations",objParentObject,"RootStructures",sRootStructure)
	End If

	If Instr(sAction,"_OnBelowTable") > 0 Then
		objTable.SetTOProperty"Index", 1
	Else
		objTable.SetTOProperty"Index", 0
	End If

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "TabPopupMenuSelect"
				objParentObject.JavaStaticText("RootStructureHeader").SetTOProperty "label", sRootStructureHeader
				objParentObject.JavaStaticText("RootStructureHeader").Click 1,1,"RIGHT"
				wait 2
				Fn_SrvMgr_RootStructureTableOperations = Fn_UI_JavaMenu_Select("",JavaWindow("ServiceManager"),sPopupMenu)

		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Exist", "Exists","Exist_OnBelowTable"
				  iRowIndex = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
				  If iRowIndex <> -1 Then
					  Fn_SrvMgr_RootStructureTableOperations = True
					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SrvMgr_RootStructureTableOperations ] Successfully verified existence of Node [ " & sNodeName & " ].")
				  Else
    					  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SrvMgr_RootStructureTableOperations ] Node [ " & sNodeName & " ] is not exists.")
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select","Select_OnBelowTable"
				  iRowIndex = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
				  If iRowIndex <> -1 Then
					  Call Fn_UI_JavaTable_SelectRow("Fn_SrvMgr_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex)
					  Fn_SrvMgr_RootStructureTableOperations = True
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "MultiSelect" ,"MultiSelect_OnBelowTable"
			sNodeName=split(sNodeName,"~") 
				For iNodeNo=0 to Ubound(sNodeName)
					iRowNo = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName(iNodeNo))
					If isNumeric(iRowNo) Then
						If iNodeNo=0 Then
							objTable.SelectRow iRowNo
							Fn_SrvMgr_RootStructureTableOperations=True
						Else
							objTable.ExtendRow "#"&iRowNo
							Fn_SrvMgr_RootStructureTableOperations=True
						End If					
					End if
				Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Expand","Expand_OnBelowTable"
				  iRowIndex = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
				  If iRowIndex <> -1 Then
					  Call Fn_UI_JavaTable_SelectRow("Fn_SrvMgr_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex)
					  wait 2
					  Fn_SrvMgr_RootStructureTableOperations = Fn_MenuOperation("WinMenuSelect", "View:Expand")
				  Else 
				      Fn_SrvMgr_RootStructureTableOperations = False
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ExpandBelow","ExpandBelow_OnBelowTable"
				  iRowIndex = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
				  If iRowIndex <> -1 Then
					  Call Fn_UI_JavaTable_SelectRow("Fn_SrvMgr_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex)
					  wait 2
					  Fn_SrvMgr_RootStructureTableOperations = Fn_MenuOperation("WinMenuSelect", "View:Expand Below")   
					  'Set objExpandbelow =  JavaWindow("ServiceManager").JavaWindow("ServiceWindow").JavaDialog("ExpandBelow")
					  Set objExpandbelow = JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet").JavaDialog("Expand Below")
					  If Fn_UI_ObjectExist("Fn_SrvMgr_RootStructureTableOperations",objExpandbelow)  then
						  'Click Yes Button 
						  Call Fn_Button_Click("Fn_SrvMgr_RootStructureTableOperations", objExpandbelow, "Yes")
					  End If
					  Set objExpandbelow = nothing
                  Else 
				     Fn_SrvMgr_RootStructureTableOperations = False
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PopupSelect","PopupSelect_OnBelowTable"
				   iRowIndex = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
				   
				   If sColName = "" Then
	     			   iColIndex = 0
				   Else
						iColIndex = Fn_SrvMgr_RootStructureTableColumnOperations("GetIndex", sColName)						
				   End If
				   
				  If iRowIndex <> -1 AND iColIndex <> -1 Then
						Call Fn_UI_JavaTable_SelectRow("Fn_SrvMgr_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex)
						aMenu = split(sPopupMenu,":",-1,1)
						objTable.ClickCell iRowIndex, iColIndex ,"RIGHT"
						wait 1
						Select Case Ubound(aMenu)
							Case 0
								strMenu = JavaWindow("ServiceManager").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
								JavaWindow("ServiceManager").WinMenu("ContextMenu").Select strMenu
							Case 1
								strMenu = JavaWindow("ServiceManager").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
								JavaWindow("ServiceManager").WinMenu("ContextMenu").Select strMenu
						End Select
						Fn_SrvMgr_RootStructureTableOperations = True
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CellEdit", "CellEdit_OnBelowTable"
				   iRowIndex = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
				   
				   If sColName = "" Then
					   iColIndex = 0
				   Else
						iColIndex = Fn_SrvMgr_RootStructureTableColumnOperations("GetIndex", sColName)						
				   End If
				   
				  If iRowIndex <> -1 AND iColIndex <> -1 Then
					  Call Fn_UI_JavaTable_SelectRow("Fn_SrvMgr_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex)
					  JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet").JavaTable("RootStructuresTable").ClickCell iRowIndex, iColIndex,"LEFT"
					  If JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet").JavaEdit("RootStructTblCellEdit").Exist(2) Then
                          'Workaround fpr PR#6813387
						  JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet").JavaEdit("RootStructTblCellEdit").Set ""
						  
						  call Fn_Edit_Box("Fn_ABM_RootStructureTableOperations",JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet"),"RootStructTblCellEdit",sValue)

						  'JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet").JavaEdit("RootStructTblCellEdit").Set sValue                          
						  JavaWindow("ServiceManager").JavaWindow("TcDefaultApplet").JavaEdit("RootStructTblCellEdit").Activate
					  Else
						Call Fn_UI_JavaTable_SetCellData("Fn_SrvMgr_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex,iColIndex,sValue)
					  End If
					Fn_SrvMgr_RootStructureTableOperations = True
				  End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CellVerify", "CellVerify_OnBelowTable"
		            'Workaround fpr PR#6813387
					If Trim(LCase(sValue)) = "y" Then
						sValue = "True"
					ElseIf Trim(sValue) = "" Then
						sValue = "False"
					End If

				   iRowIndex = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
				   
				   If sColName = "" Then
					   iColIndex = 0
				   Else
						iColIndex = Fn_SrvMgr_RootStructureTableColumnOperations("GetIndex", sColName)						
				   End If
				   
				  If iRowIndex <> -1 AND iColIndex <> -1 Then
					Call Fn_UI_JavaTable_SelectRow("Fn_SrvMgr_RootStructureTableOperations",objParentObject ,"RootStructuresTable",iRowIndex)
					sActValue = Fn_UI_JavaTable_GetCellData("Fn_SrvMgr_RootStructureTableOperations", objParentObject, "RootStructuresTable",iRowIndex,iColIndex)
					If IsNumeric(sActValue) AND   IsNumeric(sValue) Then						
							If CLng(sActValue) = CLng(sValue) Then						
									Fn_SrvMgr_RootStructureTableOperations = True						
							Else							
							     Fn_SrvMgr_RootStructureTableOperations = False						
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvMgr_RootStructureTableOperations ] in case [ " & sAction & " ].")						
							End If
					ElseIf IsDate(sActValue) AND   IsDate(sValue) Then      
						  'If Instr(CDate(sActValue) ,CDate(sValue)) Then     
						  If (CDate(DateValue(sActValue))  = CDate(DateValue(sValue))) Then
							 Fn_SrvMgr_RootStructureTableOperations = True      
						   Else       
								Fn_SrvMgr_RootStructureTableOperations = False      
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvMgr_RootStructureTableOperations ] in case [ " & sAction & " ].")      
						   End If 
					ElseIf Trim(sActValue) = Trim(sValue) Then							
					      Fn_SrvMgr_RootStructureTableOperations = True						
					 Else							
					      Fn_SrvMgr_RootStructureTableOperations = False							
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvMgr_RootStructureTableOperations ] in case [ " & sAction & " ].")						
					 End If				
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "VerifyForegroundColour", "VerifyBackgroundColour", "VerifyForegroundColour_OnBelowTable", "VerifyBackgroundColour_OnBelowTable"
				iRowIndex = Fn_SrvMgr_RootStructureTableRowIndex(objTable, sNodeName)
				If cint(iRowIndex) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_ABM_RootStructureTableOperations] Couldnt find  Root Structure Table Node [" + sNodeName + "]")
					Exit function
				End If
				Set  objNodeForRow =  objTable.Object.getNodeForRow(cint(iRowIndex))
				' if background colour
				If sAction = "VerifyBackgroundColour" OR sAction = "VerifyBackgroundColour_OnBelowTable" Then
					sColour = objTable.Object.getBackground(objNodeForRow,False).toString()
				Else
				' if foreground colour
					sColour = objTable.Object.getForeground(objNodeForRow,False).toString()
				End If
		
				sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
				' comparing colour codes RGB
				Select Case lcase(sColour)
					Case "[r=0,g=0,b=0]"
						sColour = "BLACK"
					Case "[r=159,g=255,b=159]"
						sColour = "GREEN"
					Case "[r=255,g=255,b=128]", "[r=255,g=200,b=0]", "[r=255,g=255,b=0]"
						sColour = "YELLOW"
					Case "[r=0,g=255,b=255]"
						sColour = "CYAN"
					Case "[r=0,g=0,b=255]"
						sColour = "BLUE"
					Case "[r=255,g=121,b=121]", "[r=255,g=0,b=0]"
						sColour = "RED"
					Case Else
						sColour = ""
				End Select
				If sValue = sColour Then
					Fn_SrvMgr_RootStructureTableOperations = True
				Else
					Fn_SrvMgr_RootStructureTableOperations = False
				ENd If
          	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	  		 Case Else
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SrvMgr_RootStructureTableOperations ] Invalid case [ " & sAction & " ].")
			    Fn_SrvMgr_RootStructureTableOperations = False

	End Select

	If Fn_SrvMgr_RootStructureTableOperations = True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SrvMgr_RootStructureTableOperations ] executed successfully with case [ " & sAction & " ].")
	End If
	Set objParentObject = Nothing
	Set objTable = Nothing
End Function

''**********************    Function to create disposition in Service Manager***************************************
'
''Function Name		 	:			  Fn_SISW_SrvMgr_DispositionCreate
'
''Description		    :  	      Function to create disposition in Service Manager 
'
''Parameters		    :	 	1. sAction : Action need to perform
'					   			2. sDispositionValue : Root Structure header tab label
'					   			3. bOperational : Root Structure item
'					   			4. bIsActive
'								
''Return Value		    :  		True / False
'
''Pre-requisite		    :		Service Manager perspective should be selected

''Examples		     	:	Call  Fn_SISW_SrvMgr_DispositionCreate("Create", "B11", "False", "True")


'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						 Veena  Gurjar		 22-June-2012			1.0					 Created                            Kaustubh Watwe
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_DispositionCreate(sAction, sDispositionValue, bOperational, bIsActive)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_DispositionCreate"
	Dim objDispositionCreate
	Fn_SISW_SrvMgr_DispositionCreate = False
	Set objDispositionCreate = JavaWindow("ServiceManager").JavaWindow("NewDispositionType")
	
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_DispositionCreate", objDispositionCreate) = False Then
		Call Fn_MenuOperation("Select","File:New:Disposition Type...")
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_DispositionCreate", objDispositionCreate) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_DispositionCreate ] Failed to find [ New Disposition Type ] window.")
			Exit function
		End IF
	End IF
	
	Select Case sAction
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
		Case "Create"
			'set value
			If sDispositionValue <> "" Then
				Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_DispositionCreate", objDispositionCreate, "DispositionValue", sDispositionValue)
			End If

			If bOperational <> "" Then
				If cBool(bOperational) Then
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_DispositionCreate", objDispositionCreate, "Operational","ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_DispositionCreate", objDispositionCreate, "Operational","OFF")
				End If
			End If

			If bIsActive <> "" Then
				If cBool(bIsActive) Then
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_DispositionCreate", objDispositionCreate, "IsActive","ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_DispositionCreate", objDispositionCreate, "IsActive","OFF")
				End If
			End If
			
			' click on Ok
			Call Fn_Button_Click("Fn_SISW_SrvMgr_DispositionCreate",objDispositionCreate,"OK" )
			Fn_SISW_SrvMgr_DispositionCreate = True
			
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_DispositionCreate ] Invalid case [ " & sAction & " ].")
		' - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - -  - - - - - - 
	End Select
	If  Fn_SISW_SrvMgr_DispositionCreate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SrvMgr_DispositionCreate ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objDispositionCreate = Nothing
End Function
''**********************    Function to Generate As-Maintained Structure in Service Manager***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_GenerateAsMaintainedStructure
'
''Description		    :	Function to Generate As-Maintained Structure in Service Manager 
'
''Parameters		    :	1. sAction : Action need to perform
'							2. sOpenDialogBy : Root Structure header tab label
'					   		3. sRootStructureNode : Root Structure item
'					   		4. dicGenerateMaintainedStruct : dictionary object for Generate As-Maintained Structure parameters
'					   		5. dicProperties : for future use

							'Dim dicGenerateMaintainedStruct
							'Set dicGenerateMaintainedStruct = CreateObject( "Scripting.Dictionary" )

							'dicGenerateMaintainedStruct("Part") = ""
							'dicGenerateMaintainedStruct("SerialNumber") 
							'dicGenerateMaintainedStruct("bUseSerialNumberGenerators") = True
							'dicGenerateMaintainedStruct("Lot") = ""
							'dicGenerateMaintainedStruct("ManufacturersID") "1010101"
							'dicGenerateMaintainedStruct("StructureContextName") = "asd"
							'dicGenerateMaintainedStruct("ManufacturingDate") = ""
							'dicGenerateMaintainedStruct("InstallationTime") = ""
							'dicGenerateMaintainedStruct("LocationName") = ""
							'dicGenerateMaintainedStruct("DispositionValue") = ""
							'dicGenerateMaintainedStruct("NumberOfLevels") = ""
'								
'Return Value		    :  		True / False
'
'Pre-requisite		    :		Service Manager perspective should be selected

''Examples		     	:	Call  Fn_SISW_SrvMgr_GenerateAsMaintainedStructure("GenerateAsMaintainedStructure", "RMB", "000060/A;1-top (View)", dicGenerateMaintainedStruct, "")


'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Kaustubh Watwe 		25-June-2012			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Nikhil D	 		04-June-2014			1.0					 modified to set label property of Property_Label object
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_GenerateAsMaintainedStructure(sAction, sOpenDialogBy, sRootStructureNode, dicGenerateMaintainedStruct, dicProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_GenerateAsMaintainedStructure"
	Dim objMaintainedStruct, bReturn, arrDateTime
	Set objMaintainedStruct = Fn_SISW_SrvMgr_GetObject("GenerateAsMaintainedStructure")
	Fn_SISW_SrvMgr_GenerateAsMaintainedStructure = False
	
	'Added to handle 'Show Unconfigured is set' dialog popping up since Tc1122_1119 build 
'	If JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set").Exist(3) Then
'		JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set").JavaButton("OK").Click micLeftBtn
'	End If

	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure","Exist", JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set"),SISW_MINLESS_TIMEOUT) Then
		JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set").JavaButton("OK").Click micLeftBtn
	End If
	
	'If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct) = False Then
	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure","Exist", objMaintainedStruct,SISW_DEFAULT_TIMEOUT) = False Then
		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "RMB"
					If sRootStructureNode <> "" then
                        bReturn = Fn_SrvMgr_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", "Generate As-Maintained Structure...")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_GenerateAsMaintainedStructure ] Failed to perform [ RMB : Generate As-Maintained Structure... ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objMaintainedStruct = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Menu", ""
					If sRootStructureNode <> "" then
						bReturn = Fn_SrvMgr_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_GenerateAsMaintainedStructure ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objMaintainedStruct = Nothing
							Exit function
						End If
					End If
					Call Fn_MenuOperation("Select","Tools:Generate As-Maintained Structure...")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		End Select
		
		'Added to handle 'Show Unconfigured is set' dialog popping up since Tc1122_1119 build 
'		If JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set").Exist(3) Then
'			JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set").JavaButton("OK").Click micLeftBtn
'		End If
'		
'		If  Fn_UI_ObjectExist("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct) = False Then
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_GenerateAsMaintainedStructure ] Failed to find [ Generate As-Maintained Structure ] window.")
'			Set objMaintainedStruct = Nothing
'			Exit function
'		End If

		If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure","Exist", JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set"),SISW_MINLESS_TIMEOUT) Then
			JavaWindow("DefaultWindow").JavaWindow("ShowUnconfigured_is_set").JavaButton("OK").Click micLeftBtn
		End If
		
		If  Fn_SISW_UI_Object_Operations("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure","Exist", objMaintainedStruct,SISW_MICRO_TIMEOUT) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_GenerateAsMaintainedStructure ] Failed to find [ Generate As-Maintained Structure ] window.")
			Set objMaintainedStruct = Nothing
			Exit function
		End If
	End If
	Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
    Case "GenerateAsMaintainedStructure"
		' Part
		' not implemented yet

		'Serial Number
		If dicGenerateMaintainedStruct("SerialNumber") <> "" Then
		objMaintainedStruct.JavaStaticText("Property_Label").SetTOProperty "label", "Serial Number :"
		'wait 1
			Call Fn_Edit_Box("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct, "SerialNumber", dicGenerateMaintainedStruct("SerialNumber"))
		End If

		'Use Serial Number Generators
		If dicGenerateMaintainedStruct("bUseSerialNumberGenerators") <> "" Then
			If cBool(dicGenerateMaintainedStruct("bUseSerialNumberGenerators")) Then
				Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct, "UseSerialNumberGenerators","ON")
			Else
				Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct, "UseSerialNumberGenerators","OFF")
			End If
		End If

		'Lot
		If dicGenerateMaintainedStruct("Lot") <> "" Then
			Call Fn_List_Select("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct,"Lot", dicGenerateMaintainedStruct("Lot")) 
		End If

		' Manufacturer ID
		objMaintainedStruct.JavaStaticText("Property_Label").SetTOProperty "label", "Manufacturer's ID :"
		'wait 1
		If dicGenerateMaintainedStruct("ManufacturersID") <> "" Then
			Call Fn_Edit_Box("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct, "ManufacturersID", dicGenerateMaintainedStruct("ManufacturersID"))
		End If

		'Structure Context Name
		If dicGenerateMaintainedStruct("StructureContextName") <> "" Then
			Call Fn_Edit_Box("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct, "StructureContextName", dicGenerateMaintainedStruct("StructureContextName"))
		End If

	' set Manufacturing Date 
		If dicGenerateMaintainedStruct("ManufacturingDate")  <> "" Then
			objMaintainedStruct.JavaStaticText("Property_Label").SetTOProperty "label", "Manufacturing Date :"
			'wait 1
'			Call Fn_Button_Click("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct, "ManufacturingDateButton")
			objMaintainedStruct.JavaEdit("ManufacturingDate").Activate
			wait 1
			call Fn_KeyBoardOperation("SendKeys", "{TAB}~ ")
			If lcase(dicGenerateMaintainedStruct("ManufacturingDate")) = "today" then
				'Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure","Today","")
				arrDateTime = Split(Now," ")
			Else
				arrDateTime = Split(dicGenerateMaintainedStruct("ManufacturingDate")," ")
				'Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure",arrDateTime(0),arrDateTime(1))
			End If
			objMaintainedStruct.JavaEdit("ManufacturingDate").Set arrDateTime(0)
			'wait 1
			call Fn_KeyBoardOperation("SendKeys", "{TAB}")
			objMaintainedStruct.JavaList("ManufacturingDateList").Type arrDateTime(1)
		End If
			
		' set  Installation Time 
		If dicGenerateMaintainedStruct("InstallationTime")  <> "" Then
			objMaintainedStruct.JavaStaticText("Property_Label").SetTOProperty "label", "Installation Time :"
			'wait 1
'			Call Fn_Button_Click("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct, "InstallationTimeButton")
			objMaintainedStruct.JavaEdit("InstallationTime").Activate
			wait 1
			call Fn_KeyBoardOperation("SendKeys", "{TAB}~ ")
			wait 1
			If lcase(dicGenerateMaintainedStruct("InstallationTime")) = "today" then
				'Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure","Today","")
				arrDateTime = Split(Now," ")
			Else
				arrDateTime = Split(dicGenerateMaintainedStruct("InstallationTime")," ")
				'Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure",arrDateTime(0),arrDateTime(1))
			End If
		 	objMaintainedStruct.JavaEdit("InstallationTime").Set arrDateTime(0)
			'wait 1
			call Fn_KeyBoardOperation("SendKeys", "{TAB}")
			objMaintainedStruct.JavaList("InstallationTimeList").Type arrDateTime(1)
		End If

		If dicGenerateMaintainedStruct("LocationName") <> "" Then
			Call Fn_List_Select("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct,"LocationName", dicGenerateMaintainedStruct("LocationName")) 
		End If

		If dicGenerateMaintainedStruct("DispositionValue") <> "" Then
		objMaintainedStruct.JavaStaticText("Property_Label").SetTOProperty "label", "Disposition Value :"
		'wait 1
			Call Fn_List_Select("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct,"DispositionValue", dicGenerateMaintainedStruct("DispositionValue")) 
		End If

		'Number Of Levels
		If dicGenerateMaintainedStruct("NumberOfLevels") <> "" Then
		objMaintainedStruct.JavaStaticText("Property_Label").SetTOProperty "label", "Number of Levels:"
		'wait 1
			' Can not type 0 in NumberOfLevels field
			objMaintainedStruct.JavaEdit("NumberOfLevels").Type " "
			call Fn_KeyBoardOperation("SendKeys", "{END}~+{HOME}~{BKSP}")
			If dicGenerateMaintainedStruct("NumberOfLevels") = "0"  Then
				call Fn_KeyBoardOperation("SendKeys", "0")
			Else
				objMaintainedStruct.JavaEdit("NumberOfLevels").Type dicGenerateMaintainedStruct("NumberOfLevels")
			End If
		End If
		' Clicking on OK button
		Call Fn_Button_Click("Fn_SISW_SrvMgr_GenerateAsMaintainedStructure", objMaintainedStruct, "OK")
		Fn_SISW_SrvMgr_GenerateAsMaintainedStructure = True
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Case Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_GenerateAsMaintainedStructure ] Invalid case [ " & sAction & " ].")
	End Select

	If  Fn_SISW_SrvMgr_GenerateAsMaintainedStructure <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SrvMgr_GenerateAsMaintainedStructure ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objMaintainedStruct = Nothing
End Function
''**********************    Function to create New Asset in Service Manager***************************************
'
''Function Name		:	Fn_SISW_SrvMgr_NewAssetGroupCreate
'
''Description		    :	Function to create New Asset in Service Manager 
'
''Parameters		   :	1. sAction : Action need to perform
'					  2. sName : Asset Name
'					  3. sDescription : Description tesxt
'								
'Return Value		   :  		True / False
'
'Pre-requisite		    :		Service Manager perspective should be selected

''Examples		    :	Call Fn_SISW_SrvMgr_NewAssetGroupCreate("Create", "AssetName", "Desc")

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Kaustubh Watwe 		26-June-2012			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_NewAssetGroupCreate(sAction, sName, sDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_NewAssetGroupCreate"
	Dim objNewAsset
	Set objNewAsset = Fn_SISW_SrvMgr_GetObject("NewAssetGroup")
	Fn_SISW_SrvMgr_NewAssetGroupCreate = False
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_NewAssetGroupCreate", objNewAsset) = False Then
		' perform menu operation
		Call Fn_MenuOperation("Select","File:New:Asset Group...")
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_NewAssetGroupCreate", objNewAsset) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_NewAssetGroupCreate ] Failed to find window [ New Asset Group... ].")
			Exit Function
		End If
	End If
	Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Create"
			Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewAssetGroupCreate", objNewAsset,"AssetGroupName", sName)

			If sDescription <> "" Then
				Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewAssetGroupCreate", objNewAsset, "Description", sDescription)
			End If

			Call Fn_Button_Click("Fn_SISW_SrvMgr_NewAssetGroupCreate", objNewAsset, "OK")
			Fn_SISW_SrvMgr_NewAssetGroupCreate = True
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_NewAssetGroupCreate ] Invalid case [ " & sAction & " ].")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If  Fn_SISW_SrvMgr_NewAssetGroupCreate <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SrvMgr_NewAssetGroupCreate ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objNewAsset = Nothing
End Function
''**********************    Function to create New Characteristics in Service Manager***************************************
'
''Function Name		:	Fn_SISW_SrvMgr_NewCharacteristics
'
''Description		    :	Function to create New Life Characteristic in Service Manager 
''Description		    :	Function to create New Date Characteristic in Service Manager 
''Description		    :	Function to create New Observation Characteristic in Service Manager 
'
''Parameters		   :	1. sAction : Action need to perform
'					  2. sCharacteristicsName : Characteristics Name
'					  3. sUnit : Unit text
'					  4. sPrecision : Precision text
'					  5. sDerivedExpression : Derived expression
'					  6. sCharacteristics : ~ separated list of Characteristics
'					  7. sOperations : ~ separated list of operations
'					  8. sExpression : for future use ( Expression dialog )
'								
'Return Value		   :  True / False
'
'Pre-requisite		    :	Service Manager perspective should be selected

''Examples		    :	Call Fn_SISW_SrvMgr_NewCharacteristics("CreateNewLifeCharacteristic", "LifChar1", "ml", "", "", "", "", "")
''Examples		    :	Call Fn_SISW_SrvMgr_NewCharacteristics("CreateNewDateCharacteristic", "DateCharName", "", "", "", "", "", "")
''Examples		    :	Call Fn_SISW_SrvMgr_NewCharacteristics("CreateNewObservationCharacteristic", "ObservName", "ml", "", "", "", "", "")

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Kaustubh Watwe 		27-June-2012			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_NewCharacteristics(sAction, sCharacteristicsName, sUnit, sPrecision, sDerivedExpression, sCharacteristics, sOperations, sExpression)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_NewCharacteristics"
	Dim objNewChar, objExpressions, arrChar, arrOpr, iCnt
	Set objNewChar = Fn_SISW_SrvMgr_GetObject("NewCharacteristicsWindow")
	Set objExpressions = Fn_SISW_SrvMgr_GetObject("Expression")
	Fn_SISW_SrvMgr_NewCharacteristics = False
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_NewCharacteristics", objNewChar) = False Then
		' perform menu operation
		If instr(sAction, "Life") > 0 Then
			Call Fn_MenuOperation("Select","File:New:Life Characteristic...")
		ElseIf instr(sAction,"Date") > 0 Then
			Call Fn_MenuOperation("Select","File:New:Date Characteristic...")
		ElseIf instr(sAction,"Observation") > 0 Then
			Call Fn_MenuOperation("Select","File:New:Observation Characteristic...")
		End If
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_NewCharacteristics", objNewChar) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_NewCharacteristics ] failed to find New Characteristics window.")
			Exit Function
		End If
	End If
	Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CreateNewLifeCharacteristic", "CreateNewDateCharacteristic", "CreateNewObservationCharacteristic"

			Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewCharacteristics", objNewChar, "CharacteristicName", sCharacteristicsName)

			If sAction <> "CreateNewDateCharacteristic" Then
				If sUnit <> "" Then
					Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewCharacteristics", objNewChar, "Unit", sUnit)
				End If

				If sPrecision <> "" Then
					Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewCharacteristics", objNewChar, "Precision", sPrecision)
				End If

				If sDerivedExpression <> "" Then
					Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewCharacteristics", objNewChar, "DerivedExpression", sDerivedExpression)
				End If

				If sCharacteristics <> "" Then
					Call Fn_Button_Click("Fn_SISW_SrvMgr_NewCharacteristics", objNewChar, "DerivedExpButton")
                                        arrChar =split(sCharacteristics,"~")
					arrOpr = split(sOperations,"~")
					If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_NewCharacteristics",objExpressions) = False Then
						Exit Function
					End If
					For iCnt = 0 to UBound(arrChar)
						If trim(arrChar(iCnt)) <> "" Then
							Wait SISW_MIN_TIMEOUT
							Call Fn_List_Select("Fn_SISW_SrvMgr_NewCharacteristics",objExpressions,"Characteristics", trim(arrChar(iCnt)))
							Call Fn_Button_Click("Fn_SISW_SrvMgr_NewCharacteristics", objExpressions, "CharacteristicsAppend")
						End If
						wait 1
						If trim(arrOpr(iCnt)) <> "" Then
							Call Fn_List_Select("Fn_SISW_SrvMgr_NewCharacteristics",objExpressions,"Operations", trim(arrOpr(iCnt)))
							Call Fn_Button_Click("Fn_SISW_SrvMgr_NewCharacteristics", objExpressions, "OperationsAppend")
						End If
						wait 1
					Next
					If sExpression <> "" Then
						objExpressions.JavaEdit("DerivedExpression").Click 5, 5,micLeftBtn
						WAIT 1
						Call Fn_KeyBoardOperation("SendKeys", "{END}")
						Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewCharacteristics", objExpressions, "DerivedExpression", sExpression)
					End If
					Call Fn_Button_Click("Fn_SISW_SrvMgr_NewCharacteristics", objExpressions, "OK")
				End If
			End If
			
			Call Fn_Button_Click("Fn_SISW_SrvMgr_NewCharacteristics", objNewChar, "OK")
			Fn_SISW_SrvMgr_NewCharacteristics = True
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_NewCharacteristics ] Invalid case [ " & sAction & " ].")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If  Fn_SISW_SrvMgr_NewCharacteristics <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SrvMgr_NewCharacteristics ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objNewChar = Nothing
End Function
''**********************    Function to perform operations on Physical Part Usage History in Service Manager ***************************************
'
''Function Name		:	Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations
'
''Description		    :	Function to perform operations on Physical Part Usage History in Service Manager 
'
''Parameters		   :	1. sAction : Action need to perform
'					  2. sRow : Row
'					  3. sColumn : Column Name
'					  4. sValue : Value to be verified
'					  5. sPopupMenu : for future use
'								
'Return Value		   :  True / False
'
'Pre-requisite		    :	Service Manager perspective should be selected

''Examples		    :	Call Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations("Select", "00022-PhyLoc", "", "", "")
''Examples		    :	Call Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations("CellExist", "00022-PhyLoc", "Type", "Physical Location", "")

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Kaustubh Watwe 		28-June-2012			1.0				Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations(sAction, sRow, sColumn, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations"
	Dim objPhysicalPartUsageHistoryTable, iCnt, iRowCnt
    Dim arrValue,crrValue

	Set objPhysicalPartUsageHistoryTable = JavaWindow("ServiceManager").JavaTable("PhysicalPartUsageHistory")
	Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations = False

	' add code to select Physical Part Usage History panel if required.

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			iRowCnt = objPhysicalPartUsageHistoryTable.GetROProperty("rows")
			For iCnt = 0 to iRowCnt - 1
				If objPhysicalPartUsageHistoryTable.GetCellData(0,0) = sRow Then
					objPhysicalPartUsageHistoryTable.SelectCell iCnt, sColumn
					Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations = True
					Exit for
				End If
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CellExist"
			iRowCnt = objPhysicalPartUsageHistoryTable.GetROProperty("rows")
			For iCnt = 0 to iRowCnt - 1
				If objPhysicalPartUsageHistoryTable.GetCellData(iCnt,0) = sRow Then
					If sColumn="In Time" or  sColumn="Out Time" then
						arrValue=Split(objPhysicalPartUsageHistoryTable.GetCellData( iCnt, sColumn)," ")
						crrValue=arrValue(0)
					Else
						crrValue=objPhysicalPartUsageHistoryTable.GetCellData( iCnt, sColumn)
					End if

					If crrValue = sValue Then
						Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations = True
						Exit for
					End If
				End If
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations ] Invalid case [ " & sAction & " ].")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	End Select

	If  Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SrvMgr_PhysicalPartUsageHistoryTableOperations ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objPhysicalPartUsageHistoryTable = Nothing
End Function

''********************** Function to perform Unistall operation on Physical part  in Service Manager  ***************************************
'
''Function Name		:	Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations
'
''Description		    :	Function to perform Unistall operation on Physical part  in Service Manager 
'
''Parameters		   :	1. sOpenDialogBy :Open Dialog  Using Menu/RMB
'					   2. sRootStructureNode : Node to select  from BomTable
'					   3. dicUnInstallDateTime : Date and Time to uninstall physical part
'					   4. sDepositionValue :sDepositionValue
'					   5. sLocationName : LocationName to be selected 
'					   6.sButton=Buttons to be clicked on
'								
'Return Value		   :  True / False
'
'Pre-requisite		    :	Service Manager perspective should be selected
'
''Examples		     	:	
'							Dim dicUnInstallDateTime
'							Set dicUnInstallDateTime = CreateObject( "Scripting.Dictionary" )			
'			
'								Set dicUnInstallDateTime=CreateObject("Scripting.Dictionary")
'								dicUnInstallDateTime("UnInstallationDate") ="29-jan-2012"
'								dicUnInstallDateTime("UnInstallTime")="12:25:00 PM"
			:	
''Examples		    : Call	Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations("RMB","000853/--A (View):000855/--A",dicUnInstallDateTime,sDepositionValue,"NewPhyLoc_73475","OK:Cancel")
''Examples		    : Call	Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations("Menu","000853/--A (View):000855/--A",dicUnInstallDateTime,sDepositionValue,"NewPhyLoc_73475","OK:Cancel")

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Avinash Jagdale 		28-June-2012			1.0				Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations(sOpenDialogBy,sRootStructureNode,dicUnInstallDateTime,sDepositionValue,sLocationName,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations"
	Dim objUnInstallPhysicalPart,iCount
	Set objUnInstallPhysicalPart = Fn_SISW_SrvMgr_GetObject("UninstallPhysicalPart")
	bFlag = False
	Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations = False
	'If Fn_UI_ObjectExist("Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations", objUnInstallPhysicalPart) = False Then
	If Fn_SISW_UI_Object_Operations("Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations","Exist", objUnInstallPhysicalPart,SISW_MINLESS_TIMEOUT) = False Then
		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "RMB"
					If sRootStructureNode <> "" then
						bReturn =  Fn_SrvMgr_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode,"","" , "Uninstall Physical Part...")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations ] Failed to perform [ RMB : Uninstall Physical Part ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objUnInstallPhysicalPart = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Menu", ""
					If sRootStructureNode <> "" then
						bReturn = Fn_SrvMgr_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objUnInstallPhysicalPart = Nothing
							Exit function
						End If
					End If
					Call Fn_MenuOperation("Select","Tools:Uninstall Physical Part...")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		End Select
		End If

		If Fn_UI_ObjectExist("Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations", objUnInstallPhysicalPart) = True Then
			If TypeName(dicUnInstallDateTime) <> "String" Then
				' setting Installation date
				If dicUnInstallDateTime("UnInstallationDate") <> "" Then
					'Call Fn_Button_Click("Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations", objUnInstallPhysicalPart, "UnInstallTime")
					'wait 1
					'Call Fn_UI_SetDateAndTime("Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations",dicUnInstallDateTime("UnInstallationDate"),dicUnInstallDateTime("UnInstallTime"))
					objUnInstallPhysicalPart.JavaEdit("UninstallationTime").Set dicUnInstallDateTime("UnInstallationDate")
					wait 1
					call Fn_KeyBoardOperation("SendKeys", "{TAB}")					
				End If
				If dicUnInstallDateTime("UnInstallTime")<>"" Then
					objUnInstallPhysicalPart.JavaList("UninstallTimeList").Type dicUnInstallDateTime("UnInstallTime")
				End If
			End If
			' To select  Location Name
			If sLocationName <> "" Then
				Call Fn_List_Select("Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations",objUnInstallPhysicalPart,"LocationName",sLocationName)
			End If
			' clicking on OK/Cancel button
			If sButton <> "" Then
				sButton=split(sButton,":")
                		For iCount=0 to Ubound(sButton)
                    			Call Fn_Button_Click("Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations", objUnInstallPhysicalPart, sButton(iCount))
                		Next
			 End If
    
			Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations=true
		End If
	
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	If  Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_ServiceManager_UnInstallPhysicalPartOperations ] executed successfuly .")
	End If
	Set objInstallPhysicalParts = Nothing

End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName

'Description			 :	Function to retrive real column name of table Show Top Physical Part

'Parameters			   :  1.StrColumnName : Column Display name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Show Top Physical Part dialog should be Opened
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Sandeep N								03-Jul-2012					1.0																								Pranav Ingle
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName(StrColumnName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName"
   Select Case StrColumnName
		Case "Type"
			Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName="object_type"
		Case "Part Number"
			Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName="partNumber"
		Case "Lot Number"
			Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName="lotNumber"
		Case "Manufacturer's ID"
			Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName="manufacturerOrgId"
		Case "Serial Number"
			Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName="serialNumber"
		Case else
			Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName=false
   End Select
end function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations

'Description			 :	Function Used to Perform operation Show Top Physical Part dialog

'Parameters			   :  1.StrAction : Action name
'									2.StrObjectPath: Object node path
'								 	3.StrColName: Column Name
'								    4.StrValue: Cell value
'								    5.StrButtonName: Button Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Show Top Physical Part dialog should be Opened

'Examples				:   Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations("VerifyCellValue","001248/24-A:001250/--A","Part Number~Type","001250~Physical Part Revision","")
'								     Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations("VerifyCellValue","001248/24-A:001250/--A","Part Number","001250","Close")
'                       
'History					 :			
'	Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep N								03-Jul-2012					1.0																								Pranav Ingle
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Koustubh								30-Jul-2012					1.1						Modified case "VerifyCellValue"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations(StrAction,StrObjectPath,StrColName,StrValue,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations"
 	'Declaring variables
	Dim aStrNode,iCounter,i,iNodeItemsCount,bFlag,sProperty,aColName,aValue,iCount
	Dim ObjTree,oCurrentNode
	Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations=false
	'Checking existance of [ ShowTopPhysicalPart ] window
	If not JavaWindow("ServiceManager").JavaWindow("ShowTopPhysicalPart").Exist(3) Then
		Exit function
	End If
	'Creating object of [ ShowTopPhysicalPartTree ] tree
	Set ObjTree=JavaWindow("ServiceManager").JavaWindow("ShowTopPhysicalPart").JavaTree("ShowTopPhysicalPartTree")

	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to verify Cell data agiant specific object
		Case "VerifyCellValue"
			'Spliting node
			bFlag = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations", ObjTree, StrObjectPath,":","@")
			If bFlag <> False Then
					aStrNode = split(replace(bFlag,"#",""),":")
					Set oCurrentNode = ObjTree.Object
					For iCount = 0 to UBound(aStrNode)
						Set oCurrentNode = oCurrentNode.GetItem(aStrNode(iCount))
					Next
				'Spliting Column name and values
				aColName=Split(StrColName,"~")
				aValue=Split(StrValue,"~")
				'- - - - - - - - - - - - - - - - - - - - - - - - - -
				For iCount=0 to ubound(aColName)
					sProperty=Fn_SISW_ServiceManager_ShowTopPhysicalPartGetPropertyName(aColName(iCount))
					bFlag=False
					If oCurrentNode.getData().getProperty(sProperty)=aValue(iCount) then
						bFlag=True
					End If
					If bFlag = False Then
							Exit for
					End If
				Next
			End If
			Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations = bFlag
	End Select
	'Clicking on button
	If StrButtonName<>"" Then
		Call  Fn_Button_Click("Fn_SISW_ServiceManager_ShowTopPhysicalPartOperations",JavaWindow("ServiceManager").JavaWindow("ShowTopPhysicalPart"),StrButtonName)
	End If
	'Releasing object of [ ShowTopPhysicalPartTree ] tree
	Set ObjTree=nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_SISW_SrvMgr_NewLogBook()
'###
'###    DESCRIPTION        :   Create a new log book
'###
'###    PARAMETERS      :   1. sLogBookName
'###						2. sDescription
'###
'###	 HISTORY             :   AUTHOR                   DATE                          VERSION
'###
'###    CREATED BY     		:   Koustubh Watwe       13 - Aug - 2012                     1.0
'###
'###    REVIWED BY     :   
'###
'###    EXAMPLE          : 		Call Fn_SISW_SrvMgr_NewLogBook("LogBook1","Description")
'###
'#############################################################################################################
Public Function Fn_SISW_SrvMgr_NewLogBook(sLogBookName, sDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_NewLogBook"
	Dim objLogBook
	Fn_SISW_SrvMgr_NewLogBook = False
	Set objLogBook = JavaWindow("ServiceManager").JavaWindow("NewLogBook")
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_NewLogBook", objLogBook) = False  Then
		Call Fn_MenuOperation("Select","File:New:Log Book...")
		' perform menu operation
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_NewLogBook", objLogBook) = False  Then
			Exit Function
		end if
	End If

	Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewLogBook", objLogBook, "LogBookName", sLogBookName)

	Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_NewLogBook", objLogBook, "Description", sDescription)

	Fn_SISW_SrvMgr_NewLogBook = Fn_Button_Click("Fn_SISW_SrvMgr_NewLogBook", objLogBook, "OK")
	
	If  Fn_SISW_SrvMgr_NewLogBook <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SrvMgr_NewLogBook ] executed successfuly .")
	End If
	Set objLogBook = Nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_SISW_SrvMgr_RecordUtilization()
'###
'###    DESCRIPTION     :   Function to perform Record Utilization.
'###
'###    PARAMETERS      :   1. sAction
'###						2. dicRecordUtil
'###
'###	 HISTORY        :   AUTHOR                   DATE                          VERSION
'###
'###    CREATED BY     	:   Koustubh Watwe       13 - Aug - 2012                     1.0
'###
'###    REVIWED BY     	:   
'###
'###	Examples		:	Dim dicRecordUtil
'###							   Set dicRecordUtil = CreateObject( "Scripting.Dictionary" )
'###							   With dicRecordUtil  
'###								.Add "StructureNode",""
'###								.Add "Recording Time :","28_Feb_2008$5:00:00 PM"  
'###								.Add "Propagate", True
'###								.Add "Characteristic Name", "obs=123~lif=456"
'###								.Add "Date Characteristic", "date=28_Feb_2008$5:00:00 PM"
'###								.Add "Description :","desc"	
'###							  End with
'###							  Call Fn_SISW_SrvMgr_RecordUtilization("Set", dicRecordUtil)
'###
'#############################################################################################################
Public Function Fn_SISW_SrvMgr_RecordUtilization(sAction, dicRecordUtil)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_RecordUtilization"
	Dim objRecord, DictItems, DictKeys, arrDateTime, iCounter, iIndexPtr, iCnt
	Dim arrTableData, bFlag, iRowCnt, aFieldValue, iCount
	Set objRecord = JavaWindow("ServiceManager").JavaWindow("RecordUtilization")
	Fn_SISW_SrvMgr_RecordUtilization = False
	bFlag = False

	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_RecordUtilization", objRecord) = False  Then
		If dicRecordUtil("StructureNode") <> "" Then
			' select node
			Call Fn_SrvMgr_RootStructureTableOperations("PopupSelect", "", "", dicRecordUtil("StructureNode"), "", "", "Record Utilization...")
		Else
			' perform menu operations
			Call Fn_MenuOperation("Select","Tools:Record Utilization...")
		End If
		Call Fn_ReadyStatusSync(2)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_RecordUtilization", objRecord) = False  Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_RecordUtilization ] Failed to find [ Record Utilization ].")
			Exit Function
		end if
	End If
'	objRecord.Maximize
	Select Case sAction
		Case "Set", ""
			DictKeys = dicRecordUtil.Keys
			DictItems = dicRecordUtil.Items
			For iCounter = 0 to dicRecordUtil.Count - 1
				 If IsNull(DictKeys(iCounter)) = False Then
				  	If IsObject(DictItems(iCounter)) Then
									Set dicAddChar=dicRecordUtil("Add Characteristics")
									Fn_SISW_SrvMgr_RecordUtilization = Fn_Button_Click("Fn_SISW_SrvMgr_RecordUtilization", objRecord, dicAddChar("sAction"))
									Wait 2
									bFlag=Fn_SISW_SrvScheduler_SearchOperations("SearchAndSelect", dicAddChar)
									If bFlag = False Then Exit Function
					Else
				  
				  
				  
						IF  DictItems(iCounter) <> "" Then										
							' Set the value as per the data dictioanry key.
							Select case DictKeys(iCounter)
								Case "StructureNode"
									' Do Nothing
								Case "Characteristic Name"
									arrTableData =  Split(DictItems(iCounter),"~")
									For iIndexPtr = 0 to 5
										objRecord.JavaTable("RecordUtilTable").SetTOProperty "Index", iIndexPtr
										If objRecord.JavaTable("RecordUtilTable").GetColumnName(0) = "Characteristic Name" Then
											bFlag = True
											Exit for
										End If
									Next
									
									If bFlag = True Then
										iRowCnt = cInt(objRecord.JavaTable("RecordUtilTable").GetROProperty("rows")) 
										For iCnt = 0 to UBound(arrTableData)
											For iCount = 0  to iRowCnt-1
												aFieldValue = split(arrTableData(iCnt),"=") 
												If objRecord.JavaTable("RecordUtilTable").GetCellData(iCount,"Characteristic Name") = aFieldValue(0) Then
												    objRecord.JavaTable("RecordUtilTable").ClickCell iCount, "Value"
												    wait 1
													objRecord.JavaEdit("Text").Object.setText cInt(aFieldValue(1))
													Call Fn_KeyBoardOperation("SendKeys","{TAB}")
													Exit for
												End If
											Next
										Next
									End If
								Case "Date Characteristic"
									arrTableData =  Split(DictItems(iCounter),"~")
									For iIndexPtr = 0 to 5
										objRecord.JavaTable("RecordUtilTable").SetTOProperty "Index", iIndexPtr
										If objRecord.JavaTable("RecordUtilTable").GetColumnName(0) = "Date Characteristic" Then
											bFlag = True
											Exit for
										End If
									Next
	
									If bFlag = True Then
										iRowCnt = (cInt(objRecord.JavaTable("RecordUtilTable").GetROProperty("rows")) - 1)
										For iCnt = 0 to UBound(arrTableData)
											aFieldValue = split(arrTableData(iCnt),"=")
											For iCount = 0 To iRowCnt
												If objRecord.JavaTable("RecordUtilTable").GetCellData(iCount,"Date Characteristic") = aFieldValue(0) Then
													objRecord.JavaTable("RecordUtilTable").ClickCell iCount,"Value"
													wait 1
													Call Fn_KeyBoardOperation("SendKeys","{TAB}^ ")
													IF lcase(aFieldValue(1)) = "today" Then
														Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_RecordUtilization","Today","")
													Else
														arrDateTime = Split(aFieldValue(1),"$")
														Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_RecordUtilization",arrDateTime(0),arrDateTime(1))
													End If
												End If
											Next
										Next
									End If
								Case "Propagate"
									If cBool(DictItems(iCounter)) Then
										objRecord.JavaCheckBox("Propagate").Set  "ON"
									Else
										objRecord.JavaCheckBox("Propagate").Set  "OFF"
									End If
								Case Else
									objRecord.JavaStaticText("FieldLabel").SetTOProperty "label", DictKeys(iCounter)
									If objRecord.JavaButton("DateButton").Exist(2) Then
										Call Fn_Button_Click("Fn_SISW_SrvMgr_RecordUtilization", objRecord, "DateButton")
										' select date from dialog
										If lcase(DictItems(iCounter)) = "today" Then
											Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_RecordUtilization","Today","")
										Else
											arrDateTime = Split(DictItems(iCounter),"$")
											Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_RecordUtilization",arrDateTime(0),arrDateTime(1))
										End If
									ElseIf objRecord.JavaEdit("EditBox").Exist(2)  Then
										Call Fn_UI_EditBox_Type("Fn_SISW_SrvMgr_RecordUtilization", objRecord,"EditBox",DictItems(iCounter))
										ElseIf objRecord.JavaList("List").Exist(2)  Then
										Call Fn_List_Select("Fn_SISW_SrvMgr_RecordUtilization", objRecord,"List", DictItems(iCounter))
									Else
										' radio buttons, checkboxes are not yet implemented.
										Fn_SISW_SrvMgr_RecordUtilization = False
										Exit function
									End If
							End Select
						End If
					End If
				End If
			Next
			wait 10
			Fn_SISW_SrvMgr_RecordUtilization = Fn_Button_Click("Fn_SISW_SrvMgr_RecordUtilization", objRecord, "OK")
		Case Else
			' Invalid case
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_RecordUtilization ] Invalid case [ " & sAction & " ].")
	End Select
	If Fn_SISW_SrvMgr_RecordUtilization <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_RecordUtilization ] executed successfully with case [ " & sAction & " ].")
	End If
	Set objRecord = Nothing
End Function

'**********************    Function to Duplicate As-Maintained Structure in Service Manager***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure
'
''Description		    :	Function to Duplicate As-Maintained Structure in Service Manager 
'
''Parameters		    :	1. sAction : Action need to perform
'							2. sOpenDialogBy : Root Structure header tab label
'					   		3. sRootStructureNode : Root Structure item
'					   		4. dicGenerateMaintainedStruct : dictionary object for Generate As-Maintained Structure parameters
'					   		5. dicProperties : for future use

							'Dim dicGenerateMaintainedStruct
							'Set dicDuplicateMaintainedStruct = CreateObject( "Scripting.Dictionary" )

							'dicDuplicateMaintainedStruct("Part") = ""
							'dicDuplicateMaintainedStruct("SerialNumber") 
							'dicDuplicateMaintainedStruct("bUseSerialNumberGenerators") = True
							'dicDuplicateMaintainedStruct("Lot") = ""
							'dicDuplicateMaintainedStruct("ManufacturersID") "1010101"
							'dicDuplicateMaintainedStruct("ManufacturingDate") = ""
							'dicDuplicateMaintainedStruct("InstallationTime") = ""
							'dicDuplicateMaintainedStruct("LocationName") = ""
							'dicDuplicateMaintainedStruct("DispositionValue") = ""
							'dicGenerateMaintainedStruct("NumberOfLevels") = ""
'								
'Return Value		    :  		True / False
'
'Pre-requisite		    :		Service Manager perspective should be selected

''Examples		     	:	Call  Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure("DuplicateAsMaintainedStructure", "RMB", "000060/A;1-top (View)", dicDuplicateMaintainedStruct, "")


'History:
'						Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						Vrushali  Wani 		13-August-2012			1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure(sAction, sOpenDialogBy, sRootStructureNode, dicDuplicateMaintainedStruct, dicProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure"
	Dim objMaintainedStruct, bReturn, arrDateTime
	Set objMaintainedStruct = Fn_SISW_SrvMgr_GetObject("DuplicateAsMaintainedStructure")
	Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure = False
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct) = False Then
		Select Case sOpenDialogBy
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "RMB"
					If sRootStructureNode <> "" then
                        bReturn = Fn_SrvMgr_RootStructureTableOperations("PopupSelect", "", "", sRootStructureNode, "", "", "Duplicate As-Maintained Structure")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure ] Failed to perform [ RMB : Generate As-Maintained Structure... ] on Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objMaintainedStruct = Nothing
							Exit function
						End If
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "Menu", ""
					If sRootStructureNode <> "" then
						bReturn = Fn_SrvMgr_RootStructureTableOperations("Select", "", "", sRootStructureNode, "", "", "")
						If bReturn = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure ] Failed to select Root Strcture Node [ " & sRootStructureNode & " ].")
							Set objMaintainedStruct = Nothing
							Exit function
						End If
					End If
					Call Fn_MenuOperation("Select","File:Duplicate As-Maintained Structure")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		End Select
		If  Fn_UI_ObjectExist("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure ] Failed to find [ Duplicate As-Maintained Structure ] window.")
			Set objMaintainedStruct = Nothing
			Exit function
		End If
	End If
	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "DuplicateAsMaintainedStructure"
			' Part
			' not implemented yet

			'Serial Number
			If dicDuplicateMaintainedStruct("SerialNumber") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct, "SerialNumber", dicDuplicateMaintainedStruct("SerialNumber"))
			End If

			'Use Serial Number Generators
			If dicDuplicateMaintainedStruct("bUseSerialNumberGenerators") <> "" Then
				If cBool(dicDuplicateMaintainedStruct("bUseSerialNumberGenerators")) Then
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct, "UseSerialNumberGenerators","ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct, "UseSerialNumberGenerators","OFF")
				End If
			End If

			'Lot
			If dicDuplicateMaintainedStruct("Lot") <> "" Then
				Call Fn_List_Select("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct,"Lot", dicDuplicateMaintainedStruct("Lot")) 
			End If

			' Manufacturer ID
			If dicDuplicateMaintainedStruct("ManufacturersID") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct, "ManufacturersID", dicDuplicateMaintainedStruct("ManufacturersID"))
			End If

			' set Manufacturing Date 
			If dicDuplicateMaintainedStruct("ManufacturingDate")  <> "" Then
				Call Fn_Button_Click("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct, "ManufacturingDateButton")
				If lcase(dicDuplicateMaintainedStruct("ManufacturingDate")) = "today" Then
					Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure","Today","")
				Else
					arrDateTime = Split(dicDuplicateMaintainedStruct("ManufacturingDate")," ")
					Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure",arrDateTime(0),arrDateTime(1))
				End If
			End If
				
			' set  Installation Time 
			If dicDuplicateMaintainedStruct("InstallationTime")  <> "" Then
				Call Fn_Button_Click("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct, "InstallationTimeButton")
				If lcase(dicDuplicateMaintainedStruct("InstallationTime")) = "today" Then
					Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure","Today","")
				ElseIf lcase(dicDuplicateMaintainedStruct("InstallationTime")) = "ok" Then
					Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure","OK", "")					
				Else
					arrDateTime = Split(dicDuplicateMaintainedStruct("InstallationTime")," ")
					Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure",arrDateTime(0),arrDateTime(1))
				End If
			End If

			If dicDuplicateMaintainedStruct("LocationName") <> "" Then
				Call Fn_List_Select("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct,"LocationName", dicDuplicateMaintainedStruct("LocationName")) 
			End If

			If dicDuplicateMaintainedStruct("DispositionValue") <> "" Then
				Call Fn_List_Select("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct,"DispositionValue", dicDuplicateMaintainedStruct("DispositionValue")) 
			End If

			'Number Of Levels
			If dicDuplicateMaintainedStruct("NumberOfLevels") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct, "NumberOfLevels", dicDuplicateMaintainedStruct("NumberOfLevels"))
			End If
			' Clicking on OK button
			Call Fn_Button_Click("Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure", objMaintainedStruct, "OK")
			Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure ] Invalid case [ " & sAction & " ].")
	End Select

	If  Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: [ Fn_SISW_SrvMgr_DuplicateAsMaintainedStructure ] executed successfuly with case [ " & sAction & " ].")
	End If
	Set objMaintainedStruct = Nothing
End Function
'**********************    Function to to perform operations on Utilization tab in Service Manager***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_UtilizationTabOperations
'
''Description		    :	Function to to perform operations on Utilization tab in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. sCharacteristicName : Characteristic Name
'					  		3. sColumn : column name
'					  		4. sValue : value to be verified
'					  		5. sPopupMenu : for future use
'								
'Return Value		    :  		True / False
'
'Pre-requisite		    :		Service Manager perspective should be selected

''Examples		     	:	Call  Fn_SISW_SrvMgr_UtilizationTabOperations("Select", "life", "", "", "")
''Examples		     	:	Call  Fn_SISW_SrvMgr_UtilizationTabOperations("Exist", "life", "", "", "")
''Examples		     	:	Call  Fn_SISW_SrvMgr_UtilizationTabOperations("VerifyCell", "life", "Unit", "1", "")


'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		16-August-2012		1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Anumol P				27-March-2012		1.1				 Modified case : VerifyCell . Added code to check [ Last Utilization Date ] 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_UtilizationTabOperations(sAction, sCharacteristicName, sColumn, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_UtilizationTabOperations"
	Dim objUtilTable, iRowCnt, iCnt, sProeprtyName,crrValue,arrValue
	Set objUtilTable = JavaWindow("DefaultWindow").JavaTable("DetailsTable")

	JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").SetTOProperty "Index", 2
	Fn_SISW_SrvMgr_UtilizationTabOperations = False
	If Fn_TabFolder_Operation("VerifyActive", "Utilization", "") = False Then
		Call Fn_TabFolder_Operation("Select", "Utilization", "")
	End If

	' setting property's real name
	Select Case sColumn
		Case "Characteristics Name"
			sProeprtyName = "util_characteristic_name"
		Case "Unit"
			sProeprtyName = "util_unit"
		Case "Time Since New"
			sProeprtyName = "util_time_since_new"
		Case "Last Value"
			sProeprtyName = "util_last_value"
		Case "Last Utilization Date"
			sProeprtyName = "util_last_recorded_date"
		Case "Time on Parent"
			sProeprtyName = "util_time_on_parent"
		Case "Time Since Repair"
			sProeprtyName = "util_time_since_repair"
		Case "Time Since Overhaul"
			sProeprtyName = "util_time_since_overhaul"
		Case Else
			' Do nothing
			sProeprtyName = ""
	End Select

	Select Case sAction
		Case "Select", "Exist"
			iRowCnt = cInt(objUtilTable.GetROProperty("rows"))
			For iCnt = 0 to iRowCnt - 1
				If objUtilTable.Object.getItem(iCnt).getData().getStringProperty("util_characteristic_name") = sCharacteristicName Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_UtilizationTabOperations ] Successfully verified existence of node [ " & sCharacteristicName & " ].")
					If sAction = "Select" Then
						objUtilTable.SelectRow iCnt
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_UtilizationTabOperations ] Successfully selected node [ " & sCharacteristicName & " ].")
					End If
					Fn_SISW_SrvMgr_UtilizationTabOperations = True
					Exit for
				End If
			Next
		Case "VerifyCell"
			iRowCnt = cInt(objUtilTable.GetROProperty("rows"))
			For iCnt = 0 to iRowCnt - 1
				If objUtilTable.Object.getItem(iCnt).getData().getStringProperty("util_characteristic_name") = sCharacteristicName Then
						If sColumn="Last Utilization Date" Then
					    	 arrValue=Split(objUtilTable.Object.getItem(iCnt).getData().getStringProperty(sProeprtyName)," ")
                             crrValue=arrValue(0)
						Else
					         crrValue= objUtilTable.Object.getItem(iCnt).getData().getStringProperty(sProeprtyName) 
						End If
						 If crrValue = sValue Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_UtilizationTabOperations ] Successfully verified existence [ " & sColumn & " = " & sValue & "] for node [ " & sCharacteristicName & " ].")
							Fn_SISW_SrvMgr_UtilizationTabOperations = True
							Exit for
						End If
				End If
			Next
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_UtilizationTabOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_UtilizationTabOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_UtilizationTabOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").SetTOProperty "Index", 1
	Set objUtilTable = Nothing
End Function
'********************** Function to perform operations on Contains tab in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_ContainsPanelOperations
'
''Description		    :	Function to to perform operations on Contains tab in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. sRow : Object Name
'					  		3. sColumn : column name
'					  		4. sValue : value to be verified
'					  		5. sPopupMenu : Popup Menu select
'								
'Return Value		    :  		True / False
'
'Pre-requisite		    :	Contains Tab should be present in Service Manager perspective.

''Examples  Select	 						:	Call Fn_SISW_SrvMgr_ContainsPanelOperations("Select", "000026/-","","","")
''Examples  Multi Select = ~ Separator		:	Call Fn_SISW_SrvMgr_ContainsPanelOperations("Select", "000043/4c~000026/-","","","")
''Examples  CellVerify  					:	Call Fn_SISW_SrvMgr_ContainsPanelOperations("CellVerify", "000043/4c","Serial Number","4c","")
''Examples	PopupMenuSelect 				:	Call Fn_SISW_SrvMgr_ContainsPanelOperations("PopupMenuSelect", "000155/4c","","","Change Disposition...")
''Examples  MultiSelect and PopupMenuSelect	:	Call Fn_SISW_SrvMgr_ContainsPanelOperations("PopupMenuSelect", "000043/4c~000155/4c","","","Change Disposition...")


'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		28-August-2012		1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_ContainsPanelOperations(sAction, sRow, sColumn, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_ContainsPanelOperations"
	Dim myDeviceReplay
	Dim objGrid, iItemCount, iCnt, sCoordinates,aCoordinates, aPopupMenu
	Dim aRows, iArrCnt
	Fn_SISW_SrvMgr_ContainsPanelOperations = False
	JavaWindow("ServiceManager").Maximize
	Set objGrid = JavaWindow("ServiceManager").JavaObject("ContainsPanelGrid")

	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_ContainsPanelOperations", objGrid) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Failed to find Contains Panel Grid.")
		Exit function
	End If

	Select Case sAction
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetColumnIndex"
			Fn_SISW_SrvMgr_ContainsPanelOperations = -1
			iItemCount = cInt(objGrid.Object.getColumnCount())
			For iCnt = 0 to iItemCount -1
				If objGrid.Object.getColumn(iCnt).getText() = sColumn Then
					Fn_SISW_SrvMgr_ContainsPanelOperations = iCnt
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Successfully Found Column [ " & sColumn & " ] at position [ " & iCnt & " ].")
					Exit For
				End If
			Next

        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			aRows = split(sRow, "~")
			iItemCount = cInt(objGrid.Object.getItemCount())
			For iArrCnt = 0 to UBound(aRows)
				For iCnt = 0 to iItemCount - 1
					If objGrid.Object.getItem(iCnt).getText(0) = aRows(iArrCnt) Then
						sCoordinates = objGrid.Object.getItem(iCnt).getBounds(0).toString()
						sCoordinates = Right(sCoordinates,len(sCoordinates) - Instr(sCoordinates,"{"))
						sCoordinates = replace(sCoordinates, "}","")
						sCoordinates = replace(sCoordinates, " ","")
						aCoordinates = split(sCoordinates,",")
						If iArrCnt <> 0 Then
'							Pressing Ctrl Key for Multi select
							myDeviceReplay.KeyDown 29
						Else
							' creating Mercury.DeviceReplay object for Multiselect
							Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
						End If
						objGrid.Click (aCoordinates(0) + aCoordinates(2) /2),(aCoordinates(1) + aCoordinates(3) /2),"LEFT"
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Successfully selected row [ " & aRows(iArrCnt) & " ].")
						Fn_SISW_SrvMgr_ContainsPanelOperations = True
						Exit For
					End If
				Next
				If Fn_SISW_SrvMgr_ContainsPanelOperations <> True Then
					Fn_SISW_SrvMgr_ContainsPanelOperations = False
					Exit for
				End If
			Next
			' releasing Ctrl key
			myDeviceReplay.KeyUp 29
			Set myDeviceReplay = Nothing
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PopupMenuSelect"
			aRows = split(sRow, "~")
			iItemCount = cInt(objGrid.Object.getItemCount())
			For iArrCnt = 0 to UBound(aRows)
				For iCnt = 0 to iItemCount - 1
					If objGrid.Object.getItem(iCnt).getText(0) = aRows(iArrCnt) Then
						sCoordinates = objGrid.Object.getItem(iCnt).getBounds(0).toString()
						sCoordinates = Right(sCoordinates,len(sCoordinates) - Instr(sCoordinates,"{"))
						sCoordinates = replace(sCoordinates, "}","")
						sCoordinates = replace(sCoordinates, " ","")
						aCoordinates = split(sCoordinates,",")
						If iArrCnt <> 0 Then
'							Pressing Ctrl Key for Multi select
							myDeviceReplay.KeyDown 29
						Else
							' creating Mercury.DeviceReplay object for Multiselect
							Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
						End If
						If iArrCnt = UBound(aRows) Then
							objGrid.Click (aCoordinates(0) + aCoordinates(2) /2),(aCoordinates(1) + aCoordinates(3) /2),"LEFT"
							wait 1
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Successfully selected row [ " & aRows(iArrCnt) & " ].")
							objGrid.Click (aCoordinates(0) + aCoordinates(2) /2),(aCoordinates(1) + aCoordinates(3) /2),"RIGHT"
							wait 2
							Fn_SISW_SrvMgr_ContainsPanelOperations = Fn_UI_JavaMenu_Select("Fn_SISW_SrvMgr_ContainsPanelOperations",JavaWindow("ServiceManager"),sPopupMenu)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Successfully selected [ RMB > " & sPopupMenu & " ].")							
						Else
							objGrid.Click (aCoordinates(0) + aCoordinates(2) /2),(aCoordinates(1) + aCoordinates(3) /2),"LEFT"
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Successfully selected row [ " & aRows(iArrCnt) & " ].")
							Fn_SISW_SrvMgr_ContainsPanelOperations = True
						End If
						Exit For
					End If
				Next
				If Fn_SISW_SrvMgr_ContainsPanelOperations <> True Then
					Fn_SISW_SrvMgr_ContainsPanelOperations = False
					Exit for
				End If
			Next
			' releasing Ctrl key			
			myDeviceReplay.KeyUp 29
			Set myDeviceReplay = Nothing
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CellVerify"
			iItemCount = cInt(objGrid.Object.getItemCount())
			iColIndex = Fn_SISW_SrvMgr_ContainsPanelOperations("GetColumnIndex", "", sColumn, "", "")
			For iCnt = 0 to iItemCount - 1
				If objGrid.Object.getItem(iCnt).getText(0) = sRow Then
					If objGrid.Object.getItem(iCnt).getText(iColIndex) = sValue Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Successfully verified with case [ " & sColumn & " = " & sValue & " ] for row [ " & sRow & " ].")
						Fn_SISW_SrvMgr_ContainsPanelOperations = True
						Exit For
					End If
				End If
			Next

		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_ContainsPanelOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_ContainsPanelOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objGrid = Nothing
End Function
'********************** Function to perform operations on Move To dialog in Service Manager***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_MoveToOperations
'
''Description		    :	Function to perform operations on Move To dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. sPhysicalPart : Physical Part Name.
'					  		3. dicProprties : for future use
'					  		4. sMoveDate : Move Date, format "DD-MMM-YYY~HH:MM:SS PM"
'					  		5. sButtonName : Button name
'								
'Return Value		    :  		True / False
'
'Pre-requisite		    :		Move To dialog should be already opened in Service Manager perspective.

''Examples  			:	Call Fn_SISW_SrvMgr_MoveToOperations("Set", "", "", "28-Feb-2008~5:00:00 PM", "OK")
''Examples  			:	Call Fn_SISW_SrvMgr_MoveToOperations("Verify", "000168/-", "", "", "Cancel")


'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		28-August-2012		1.0					 Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_MoveToOperations(sAction, sPhysicalPart, dicProprties, sMoveDate, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_MoveToOperations"
	Dim objMoveTo, arrDate
	Dim objSelectType, intNoOfObjects, bFlag, iCounter

	Set objMoveTo = JavaWindow("ServiceManager").JavaWindow("MoveTo")
	Fn_SISW_SrvMgr_MoveToOperations = False
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_MoveToOperations", objMoveTo) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_MoveToOperations ] Failed to find [ Move To ].")
		Exit function
	End If
	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Set"
			If sMoveDate <> "" Then
				' call date button
				Call Fn_Button_Click("Fn_SISW_SrvMgr_MoveToOperations", objMoveTo, "DateButton")
				wait 2
				If lcase(sMoveDate) = "today" Then
					Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_MoveToOperations","Today", "")
				ElseIf lcase(sMoveDate) = "ok" Then
					Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_MoveToOperations","OK", "")					
				Else
					arrDate = split(sMoveDate,"~")
					Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_MoveToOperations",arrDate(0), arrDate(1))
				End IF
			End If
			Fn_SISW_SrvMgr_MoveToOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			bFlag = False
			If sPhysicalPart <> "" Then
				Set objSelectType = Description.Create()
				objSelectType("Class Name").value = "JavaObject"
				objSelectType("toolkit class").value = "org.eclipse.ui.forms.widgets.ImageHyperlink"
				Set intNoOfObjects = objMoveTo.ChildObjects(objSelectType)
				For iCounter = 0 to intNoOfObjects.count-1
					objMoveTo.JavaObject("CopiedPhysicalPartsHyperlink").SetTOProperty "Index", iCounter
					If objMoveTo.JavaObject("CopiedPhysicalPartsHyperlink").Object.getText() = sPhysicalPart Then
						bFlag = True
						Exit for
					End If
				Next
				If bFlag = False Then
					Exit function
				End If
			End If

			If sMoveDate <> "" Then
				bFlag = False
				If objMoveTo.JavaEdit("MoveDate").GetROProperty("value") = sMoveDate Then
					bFlag = True
				End iF
			End If
			Fn_SISW_SrvMgr_MoveToOperations = bFlag
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_MoveToOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If sButtonName <> "" Then
		Call Fn_Button_Click("Fn_SISW_SrvMgr_MoveToOperations", objMoveTo, sButtonName)
	End If

	If Fn_SISW_SrvMgr_MoveToOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_MoveToOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objGrid = Nothing
	Set intNoOfObjects = Nothing
	Set objSelectType = Nothing
End Function
'********************** Function to perform operations on Search dialog in Service Manager***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_SearchOperations
'
''Description		    :	Function to perform operations on Search dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicSearch : Dictionary object to set Search criteria.
'								
'Return Value		    :  		True / False
'
'Pre-requisite		    :		Search dialog should be already opened in Service Manager perspective.

''Examples  			:	Dim dicSearch
'					  		Set dicSearch = CreateObject("Scripting.Dictionary")
'					  		dicSearch("Name") = "to*"
'					  		dicSearch("SearchResults_Select") = "abc"
'					  		msgbox Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dicSearch)


'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		4-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_SearchOperations(sAction, dicSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_SearchOperations"
	Dim objSearch, arrDate,iRows,ArrSearchResultl
	Dim dicItems, dicKeys, iCounter,iCnt, bFlag
	Dim sRevVarRule
	
 	Fn_SISW_SrvMgr_SearchOperations = False

	Select Case sAction
		Case "SearchAndSelect"
			Set objSearch = JavaWindow("ServiceManager").JavaWindow("Shell").JavaWindow("Search")
			If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_SearchOperations",objSearch) = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SearchOperations ] Failed to find [ Search ] dialog.")
				Exit Function
			End If
			objSearch.JavaTab("SearchTab").Select "Search"
			dicItems = dicSearch.Items
			dicKeys = dicSearch.Keys
			For iCounter = 0 to dicSearch.Count - 1
				If IsNull(dicItems(iCounter)) = False Then
					If dicItems(iCounter) <> "" Then
						If instr(dicKeys(iCounter),"SearchResults") = 0 Then
							objSearch.JavaStaticText("Field_Label").SetTOProperty "label", trim(dicKeys(iCounter)) &":"
							If lcase(trim(dicItems(iCounter))) = "true" OR lcase(trim(dicItems(iCounter))) = "false" Then
								' Radio Buttons
								objSearch.JavaRadioButton("RadioButton").SetTOProperty "attached text", lcase(trim(dicItems(iCounter)))
								Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_SrvMgr_SearchOperations",objSearch, "RadioButton")
							Else

								' Date Buttons, Editbox
								If objSearch.JavaButton("DateButton").Exist(2) Then
									If lcase(trim(dicItems(iCounter))) = "today" Then
										Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_SearchOperations","Today", "")
									ElseIf lcase(trim(dicItems(iCounter))) = "ok" Then
										Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_SearchOperations","OK", "")											
									Else
										arrDate = split(trim(dicItems(iCounter)),"~")
										Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_SearchOperations",arrDate(0), arrDate(1))

									End If
								ElseIf objSearch.JavaEdit("EditBox").Exist(2) Then
									'Call Fn_Edit_Box("Fn_SISW_SrvMgr_SearchOperations", objSearch, "EditBox",trim(dicItems(iCounter)))
									objSearch.JavaEdit("EditBox").Object.setText trim(dicItems(iCounter))
									objSearch.JavaEdit("EditBox").Activate
								Else

									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SearchOperations ] Failed to find field [ " & trim(dicKeys(iCounter)) & " ].")
									Exit Function
								End If
							End IF


						End If
					End If
				End If
			Next

			Call Fn_Button_Click("Fn_SISW_SrvMgr_SearchOperations", objSearch, "Find")
			Wait 5



			'Click on  LoadAll button 
			If cInt(objSearch.JavaButton("DownButton").GetROProperty("enabled")) = 1 Then

				Call Fn_Button_Click("Fn_SISW_SrvMgr_SearchOperations",objSearch,"DownButton")
			End If
			Wait 2
			ArrSearchResultl = split(dicSearch("SearchResults_Select") ,"~")
			iRows = cInt(objSearch.JavaTable("SearchResultsTable").GetROProperty("rows"))

			For iCounter =0 to UBound(ArrSearchResultl)
				bFlag = False
				For iCnt = 0 to iRows - 1
					If objSearch.JavaTable("SearchResultsTable").GetCellData(iCnt,"Object") = ArrSearchResultl(iCounter) Then
						If  iCounter = 0 Then
							objSearch.JavaTable("SearchResultsTable").SelectCell iCnt,"Object"
						Else
							objSearch.JavaTable("SearchResultsTable").ExtendRow iCnt
						End If
						bFlag = True
						exit for
					End If
				Next
				If bFlag = False Then
					Exit Function
				End If
			Next
			Fn_SISW_SrvMgr_SearchOperations = bFlag
			If Fn_SISW_SrvMgr_SearchOperations = True Then
				Call Fn_Button_Click("Fn_SISW_SrvMgr_SearchOperations", objSearch , "OK")

			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SearchAndConfigure"
			Set objSearch = JavaWindow("ServiceManager").JavaWindow("Search_SP")
			If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_SearchOperations",objSearch) = False Then


				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SearchOperations ] Failed to find [ Search ] dialog.")
				Exit Function
			End If
			objSearch.JavaTab("SearchTab").Select "Search"
			dicItems = dicSearch.Items
			dicKeys = dicSearch.Keys
			For iCounter = 0 to dicSearch.Count - 1
				If IsNull(dicItems(iCounter)) = False Then
					If dicItems(iCounter) <> "" Then
						If instr(dicKeys(iCounter),"SearchResults") = 0 AND instr(dicKeys(iCounter),"Configure") = 0 Then
							objSearch.JavaStaticText("Field_Label").SetTOProperty "label", trim(dicKeys(iCounter)) &":"
							If lcase(trim(dicItems(iCounter))) = "true" OR lcase(trim(dicItems(iCounter))) = "false" Then
								' Radio Buttons
								objSearch.JavaRadioButton("RadioButton").SetTOProperty "attached text", lcase(trim(dicItems(iCounter)))
								Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_SrvMgr_SearchOperations",objSearch, "RadioButton")

							Else
								' Date Buttons, Editbox
								If objSearch.JavaButton("DateButton").Exist(2) Then
									If lcase(trim(dicItems(iCounter))) = "today" Then
										Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_SearchOperations","Today", "")
									ElseIf lcase(trim(dicItems(iCounter))) = "ok" Then
										Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_SearchOperations","OK", "")											
									Else


										arrDate = split(trim(dicItems(iCounter)),"~")
										Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_SearchOperations",arrDate(0), arrDate(1))
									End If


								ElseIf objSearch.JavaEdit("EditBox").Exist(2) Then
									Call Fn_Edit_Box("Fn_SISW_SrvMgr_SearchOperations", objSearch, "EditBox",trim(dicItems(iCounter)))
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SearchOperations ] Failed to find field [ " & trim(dicKeys(iCounter)) & " ].")
									Exit Function
								End If
							End IF
						End If
					End If
				End If
			Next

			Call Fn_Button_Click("Fn_SISW_SrvMgr_SearchOperations", objSearch, "Find")
			Wait 5

			'Click on  LoadAll button 
			If cInt(objSearch.JavaButton("DownButton").GetROProperty("enabled")) = 1 Then
				Call Fn_Button_Click("Fn_SISW_SrvMgr_SearchOperations",objSearch,"DownButton")
			End If
			Wait 2
			ArrSearchResultl = split(dicSearch("SearchResults_Select") ,"~")
			iRows = cInt(objSearch.JavaTable("SearchResultsTable").GetROProperty("rows"))

			For iCounter =0 to UBound(ArrSearchResultl)
				bFlag = False
				For iCnt = 0 to iRows - 1
					If objSearch.JavaTable("SearchResultsTable").GetCellData(iCnt, 0) = ArrSearchResultl(iCounter) Then
						If  iCounter = 0 Then
							objSearch.JavaTable("SearchResultsTable").SelectCell iCnt, 0
						Else
							objSearch.JavaTable("SearchResultsTable").ExtendRow iCnt
						End If
						bFlag = True
						exit for
					End If
				Next
				If bFlag = False Then
					Exit Function
				End If
			Next
			Fn_SISW_SrvMgr_SearchOperations = bFlag
			If Fn_SISW_SrvMgr_SearchOperations = True Then
				If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_SearchOperations",objSearch.JavaButton("Configure")) Then
					Call Fn_Button_Click("Fn_SISW_SrvMgr_SearchOperations", objSearch , "Configure")
				ElseIf Fn_UI_ObjectExist("Fn_SISW_SrvMgr_SearchOperations",objSearch.JavaButton("Relate")) Then
					Call Fn_Button_Click("Fn_SISW_SrvMgr_SearchOperations", objSearch , "Relate")
				End If
			Else
				exit function
			End If
			
			If dicSearch("Configure_BOMLine") <> "" Then
				Wait 2
				ArrSearchResultl = split(dicSearch("Configure_BOMLine") ,"~")
				iRows = cInt(objSearch.JavaTable("SearchResultsTable").GetROProperty("rows"))

				For iCounter =0 to UBound(ArrSearchResultl)
					bFlag = False
					For iCnt = 0 to iRows - 1
						If objSearch.JavaTable("SearchResultsTable").Object.getValueAt(0,0).toString() = ArrSearchResultl(iCounter) Then
							If  iCounter = 0 Then
								objSearch.JavaTable("SearchResultsTable").SelectCell iCnt,"BOM Line"



							Else
								objSearch.JavaTable("SearchResultsTable").ExtendRow iCnt
							End If
							bFlag = True
							exit for
						End If
					Next
					If bFlag = False Then
						Exit Function
					End If
				Next
			End If
			wait 1
			'-------------------- Set Revision Rule Code ----------------------
			If dicSearch("Configure_RevisionRule") <> "" Then
					JavaWindow("DefaultWindow").JavaStaticText("OpenItems_label").SetTOProperty "label","Latest Working"
					Call Fn_UI_JavaStaticText_Click("Fn_SISW_SrvMgr_SearchOperations",JavaWindow("DefaultWindow"), "OpenItems_label", 1, 1, "LEFT")
					wait 1
					sRevVarRule = dicSearch("Configure_RevisionRule")
					bFlag = Fn_RevisionRuleSet(sRevVarRule)
					If bFlag = False Then
						Exit Function
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to select Revision Rule [ " & dicSearch("Configure_RevisionRule") & " ].")
					End If
			End if
			'-------------------- Set Variant Rule Code ----------------------
			If dicSearch("Configure_VariantRule") <> "" Then
					JavaWindow("DefaultWindow").JavaStaticText("OpenItems_label").SetTOProperty "label","Click to add a variant rule."
					Call Fn_UI_JavaStaticText_Click("Fn_SISW_SrvMgr_SearchOperations",JavaWindow("DefaultWindow"), "OpenItems_label", 1, 1, "LEFT")
					wait 1
					ArrSearchResultl = Split(dicSearch("Configure_VariantRule"),"~")
					bFlag = Fn_PSE_VarConfigure("",ArrSearchResultl(0),ArrSearchResultl(1),ArrSearchResultl(2),ArrSearchResultl(3))
					If bFlag = False Then
						Exit Function
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to select Variant Rule [ " & dicSearch("Configure_VariantRule") & " ].")
					End If
			End if
			'-----------------------------------------------------------------
			
			If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_SearchOperations",objSearch.JavaButton("Finish")) Then
					Call Fn_Button_Click("Fn_SISW_SrvMgr_SearchOperations", objSearch , "Finish")
			End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SearchOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_SearchOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_SearchOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objSearch = Nothing
End Function
'********************** Function to perform operations on Create Activity Entry Value dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreateActivityEntryValueOperations
'
''Description		    :	Function to perform operations on Create Activity Entry Value dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicActiveEntry : Dictionary object to set  Activity Entry Value data.
'								
'Return Value		    :  		True / False
'
'Pre-requisite		    :		Service Manager perspective should be opened.

''Examples  			:	Dim dicSearch
'					  		Set dicSearch = CreateObject("Scripting.Dictionary")
'					  		dicSearch("Name") = "to*"
'					  		dicSearch("SearchResults_Select") = "abc"

'							Dim dicActiveEntry
'					  		Set dicActiveEntry = CreateObject("Scripting.Dictionary")
'					  		Set dicActiveEntry("ServiceEvent") = dicSearch
'					  		dicActiveEntry("Name") = "abc"
'					  		dicActiveEntry("Description") = "abc desc"
'					  		dicActiveEntry("CharacteristicDate") = "4-Sep-2012~11:30 PM"
'					  		dicActiveEntry("CharacteristicValue") = "123"
'					  		dicActiveEntry("bPropogate") = "true"
'					  		msgbox Fn_SISW_SrvMgr_CreateActivityEntryValueOperations("Create", dicActiveEntry)
'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		4-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateActivityEntryValueOperations(sAction, dicActiveEntry)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateActivityEntryValueOperations"
	Dim objActivityEntryValue, arrDate, bReturn
	Fn_SISW_SrvMgr_CreateActivityEntryValueOperations = False
	Set objActivityEntryValue = JavaWindow("ServiceManager").JavaWindow("CreateActivityEntryValue")
	
	'Select menu [File -> New -> Activity Entry Value...]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations",objActivityEntryValue.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Activity Entry Value...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations",objActivityEntryValue.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateActivityEntryValueOperations ] Failed to find [ Create Activity Entry Value ] dialog.")
			Exit Function
		End If
	End If
		
	Select Case sAction
		Case "Create"
			'Setting Service Event
			If TypeName(dicActiveEntry("ServiceEvent")) <> "String"  Then
				objActivityEntryValue.JavaStaticText("ServiceEventDropDown").Click 1, 1,"LEFT"
				wait 1
				objActivityEntryValue.JavaMenu("label:=Add...").Select
				wait 1
                If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dicActiveEntry("ServiceEvent")) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateActivityEntryValueOperations ] Failed to find [ Service Event ].")
					Exit Function
				End IF
			Else
				If dicActiveEntry("ServiceEvent") = "Clear" Then
					objActivityEntryValue.JavaStaticText("ServiceEventDropDown").Click 1, 1,"LEFT"
					wait 1
					objActivityEntryValue.JavaMenu("label:=Clear").Select
				End If
			End If

			' Name
			If dicActiveEntry("Name") <> ""  Then
				Call Fn_List_Select("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations", objActivityEntryValue,"Name",dicActiveEntry("Name"))
			End If

			'Description
			If dicActiveEntry("Description") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations",objActivityEntryValue,"Description", dicActiveEntry("Description"))
			End If
			
			If dicActiveEntry("CharacteristicValue") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations",objActivityEntryValue,"CharacteristicValue", dicActiveEntry("CharacteristicValue"))
			End If
			
			If dicActiveEntry("CharacteristicDate") <> ""  Then
				If objActivityEntryValue.JavaButton("CharacteristicDateButton").Exist(2) Then
					Call Fn_Button_Click("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations", objActivityEntryValue , "CharacteristicDateButton")
					If lcase(trim(dicActiveEntry("CharacteristicDate"))) = "today" Then
						Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations","Today", "")
					ElseIf lcase(trim(dicActiveEntry("CharacteristicDate"))) = "ok" Then
						Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations","OK", "")						
					Else
						arrDate = split(trim(dicActiveEntry("CharacteristicDate")),"~")
						Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations",arrDate(0), arrDate(1))
					End IF
				End If
			End If

			If dicActiveEntry("bPropogate") <> ""  Then
				objActivityEntryValue.JavaRadioButton("Propagate").SetTOProperty "attached text", lcase(dicActiveEntry("bPropogate"))
				Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations", objActivityEntryValue , "Propagate" )
			End If

			Fn_SISW_SrvMgr_CreateActivityEntryValueOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations", objActivityEntryValue , "Finish")
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			' Name
			If dicActiveEntry("Name") <> ""  Then
				If Fn_UI_ListItemExist("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations", objActivityEntryValue,"Name",dicActiveEntry("Name")) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateActivityEntryValueOperations ] Failed to verify [ " & dicActiveEntry("Name")  & " ] in Name list.")
					Exit function
				End If
			End If

			'Description
			If dicActiveEntry("Description") <> ""  Then
				If objActivityEntryValue.JavaEdit("Description").GetROProperty("value") <>dicActiveEntry("Description") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateActivityEntryValueOperations ] Failed to verify [ " & dicActiveEntry("Description")  & " ] in Description.")
					Exit Function
				End If
			End If
			
			If dicActiveEntry("CharacteristicValue") <> ""  Then
				If objActivityEntryValue.JavaEdit("CharacteristicValue").GetROProperty("value") <> dicActiveEntry("CharacteristicValue") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateActivityEntryValueOperations ] Failed to verify [ " & dicActiveEntry("CharacteristicValue")  & " ] in Characteristic Value.")
					Exit Function
				End If
			End If

			If dicActiveEntry("CharacteristicDate") <> ""  Then
				If objActivityEntryValue.JavaEdit("CharacteristicDate").GetROProperty("value") <> dicActiveEntry("CharacteristicDate") Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateActivityEntryValueOperations ] Failed to verify [ " & dicActiveEntry("CharacteristicDate")  & " ] in Characteristic Date.")
					Exit Function
				End If
			End If
			If dicActiveEntry("sButtonName") <> "" Then
				Call Fn_Button_Click("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations", objActivityEntryValue , dicActiveEntry("sButtonName"))
			End If
			Fn_SISW_SrvMgr_CreateActivityEntryValueOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateActivityEntryValueOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreateActivityEntryValueOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateActivityEntryValueOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objActivityEntryValue = Nothing
End Function
'********************** Function to perform operations on Create Service Group Type dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations
'
''Description		    :	Function to perform operations on Create Service Group Type dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicServiceGroup : Dictionary object to set Service Group Value data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Dim dicServiceGroup
'					  		Set dicServiceGroup = CreateObject("Scripting.Dictionary")

'							Dim dicContainedInSearch
'					  		Set dicContainedInSearch = CreateObject("Scripting.Dictionary")
'					  		dicContainedInSearch("Name") = "to*"
'					  		dicContainedInSearch("SearchResults_Select") = "abc"
'					  		Set dicServiceGroup("ContainedIn") = dicContainedInSearch
'										
'					  		dicServiceGroup("Name") = "abc"
'					  		dicServiceGroup("ID") = "123"
'					  		dicServiceGroup("Description") = "abc desc"
'					  		dicServiceGroup("Purpose") = "purpose"

'					  		Dim dicInProgressPhysicalPartsSearch
'					  		Set dicInProgressPhysicalPartsSearch = CreateObject("Scripting.Dictionary")
'					  		dicInProgressPhysicalPartsSearch("Name") = "part*"
'					  		dicInProgressPhysicalPartsSearch("SearchResults_Select") = "phyPart"
'					  		Set dicServiceGroup("InProgressPhysicalParts") = dicInProgressPhysicalPartsSearch
'										
'					  		msgbox Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations("Create", dicServiceGroup)


'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		5-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations(sAction, dicServiceGroup)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations"
	Dim objDialog, arrPhysicalParts, bReturn, iCount
	Dim iCnt, iRowCount
	Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations = False
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("CreateServiceGroupType")
	
	'Select menu [File -> New -> Service Group...]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Service Group...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations ] Failed to find [ Create Service Group Type ] dialog.")
			Exit Function
		End If
	End If
	objDialog.Maximize	
	Select Case sAction
		Case "Create"
			'Setting Contained In
			If TypeName(dicServiceGroup("ContainedIn")) <> "String"  Then
				objDialog.JavaStaticText("ContainedInDropDown").Click 1, 1,"LEFT"
				wait 1
				objDialog.JavaMenu("label:=Add...").Select
				wait 1
                If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dicServiceGroup("ContainedIn")) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations ] Failed to Add [ Contained In ].")
						Exit Function
				End IF
			Else
				If dicServiceGroup("ContainedIn") = "Clear" Then
					objDialog.JavaStaticText("ContainedInDropDown").Click 1, 1,"LEFT"
					wait 1
					objDialog.JavaMenu("label:=Clear").Select
				End If
			End If

			' Name
			If dicServiceGroup("Name") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations", objDialog,"Name",dicServiceGroup("Name"))
			End If

			'Description
			If dicServiceGroup("ID") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations",objDialog,"ID", dicServiceGroup("ID"))
			End If
			
			'Description
			If dicServiceGroup("Description") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations",objDialog,"Description", dicServiceGroup("Description"))
			End If
			'Setting Purpose
			If dicServiceGroup("Purpose") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations",objDialog,"Purpose", dicServiceGroup("Purpose"))
			End If
			If 	dicServiceGroup("InProgressPhysicalParts") <> "" Then
					'Setting In Progress Physical Parts
					If TypeName(dicServiceGroup("InProgressPhysicalParts")) <> "String"  Then
						objDialog.JavaStaticText("InProgressPhysicalPartDropDown").Click 1, 1,"LEFT"
						wait 1
						objDialog.JavaMenu("label:=Add...").Select
						wait 1
						If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dicServiceGroup("InProgressPhysicalParts")) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations ] Failed to Add [ In Progress Physical Parts ].")
								Exit Function
						End IF
						bReturn = True
					End If
			 End If

			If dicServiceGroup("RemoveInProgressPhysicalParts") <> "" Then
				arrPhysicalParts = split(dicServiceGroup("RemoveInProgressPhysicalParts") ,"~")
				iRowCount = cInt(objDialog.JavaTable("InProgressPhysicalParts").GetROProperty("rows")) 
				For iCount = 0 to UBound(arrPhysicalParts)
					bReturn = False
					For iCnt = 0 to iRowCount -1
						If objDialog.JavaTable("InProgressPhysicalParts").GetCellData(iCnt, 0) = arrPhysicalParts(iCount) Then
							objDialog.JavaTable("InProgressPhysicalParts").SelectCell iCnt, 0
							wait 1
							objDialog.JavaStaticText("InProgressPhysicalPartDropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Remove").Select
							bReturn = True
							Exit for
						End If
					Next
					If bReturn = False Then
						Exit For
					End If
				Next
			End If
			Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations = True
			If bReturn = False AND Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations ] Failed to Select [ " & arrPhysicalParts(iCount) & " ].")
				Exit Function
			End If

			Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations", objDialog , "Finish")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateServiceGroupTypeOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function
'********************** Function to perform operations on Create Service Group Type dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreateServiceEventTypeOperations
'
''Description		    :	Function to perform operations on Create Service Event Type dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicServiceEvent : Dictionary object to set Service Group Value data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Dim dicServiceEvent
'					  		Set dicServiceEvent = CreateObject("Scripting.Dictionary")

'							Dim dicContainedInSearch
'					  		Set dicContainedInSearch = CreateObject("Scripting.Dictionary")
'					  		dicContainedInSearch("Name") = "to*"
'					  		dicContainedInSearch("SearchResults_Select") = "abc"
'					  		Set dicServiceEvent("ContainedIn") = dicContainedInSearch
'										
'					  		dicServiceEvent("Name") = "abc"
'					  		dicServiceEvent("ID") = "123"
'					  		dicServiceEvent("Description") = "abc desc"
'					  		dicServiceEvent("Purpose") = "purpose"
'					  		dicServiceEvent("ActualLaborCost") = "1000"
'					  		dicServiceEvent("ActualLaborHours") = "8"

'					  		Dim dicInProgressPhysicalPartsSearch
'					  		Set dicInProgressPhysicalPartsSearch = CreateObject("Scripting.Dictionary")
'					  		dicInProgressPhysicalPartsSearch("Name") = "part*"
'					  		dicInProgressPhysicalPartsSearch("SearchResults_Select") = "phyPart"
'					  		Set dicServiceEvent("InProgressPhysicalParts") = dicInProgressPhysicalPartsSearch
'										
'					  		msgbox Fn_SISW_SrvMgr_CreateServiceEventTypeOperations("Create", dicServiceEvent)

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		5-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateServiceEventTypeOperations(sAction, dicServiceEvent)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateServiceEventTypeOperations"
	Dim objDialog, arrDate, bReturn, iCount
	Dim iCnt, iRowCount, iCounter
	Dim dictItems, dictKeys
	
	Fn_SISW_SrvMgr_CreateServiceEventTypeOperations = False
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("CreateServiceEventType")
	
	'Select menu [File -> New -> Service Event...]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceEventTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Service Event...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceEventTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceEventTypeOperations ] Failed to find [ Create Service Group Type ] dialog.")
			Exit Function
		End If
	End If

'	objDialog.Resize 440, 800
	objDialog.Maximize

	Select Case sAction
		Case "Create"
            'Get the keys & items count from data dictionary.	
			dictItems = dicServiceEvent.Items
			dictKeys = dicServiceEvent.Keys
			For iCounter = 0 to dicServiceEvent.Count - 1
				If IsNull(DictKeys(iCounter)) = False Then
					If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
						Select Case DictKeys(iCounter)
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "ActionType"
								Call Fn_List_Select("Fn_SISW_SrvMgr_CreateServiceEventTypeOperations", objDialog, DictKeys(iCounter),dictItems(iCounter))
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "WorkStartDate", "NeededByDate"
								If lcase(dictItems(iCounter)) = "today" Then
									arrDate = Split(Now," ")
								ElseIf lcase(dictItems(iCounter)) = "ok" Then	
									arrDate = Split(Now," ")
'									Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_CreateActivityEntryValueOperations","OK", "")									
								Else
									arrDate = split(dictItems(iCounter),"~")
								End IF
								If DictKeys(iCounter) = "NeededByDate" Then
							     	objDialog.JavaEdit("NeededByDate").Set arrDate(0)
									wait 1
									call Fn_KeyBoardOperation("SendKeys", "{TAB}")
									objDialog.JavaList("NeededByTime").Type arrDate(1)
								ElseIf DictKeys(iCounter) = "WorkStartDate" Then
								     objDialog.JavaEdit("WorkStartDate").Set arrDate(0)
									wait 1
									call Fn_KeyBoardOperation("SendKeys", "{TAB}")
									objDialog.JavaList("WorkStartTime").Type arrDate(1)
								End If

							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case "ContainedIn", "InProgressPhysicalParts"
								' ImageHyperlink  Clear Option
								objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
								wait 1
								objDialog.JavaMenu("label:=Clear").Select
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
							Case Else
								' Edit Box 
								bReturn = Fn_Edit_Box("Fn_SISW_SrvMgr_CreateServiceEventTypeOperations", objDialog , DictKeys(iCounter),dictItems(iCounter))
								If bReturn = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceEventTypeOperations ] Failed to set Editbox [ " & DictKeys(iCounter) & " = " & dictItems(iCounter) & " ].")
									Exit Function
								End If
							'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						End Select
					Else
						' For ImageHyperlink Add
						objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
						wait 1
						objDialog.JavaMenu("label:=Add...").Select
						wait 1
						If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dictItems(iCounter)) = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceEventTypeOperations ] Failed to find [ " & DictKeys(iCounter)  & " ].")
							Exit Function
						End IF
					End If
				End If
			Next

			Fn_SISW_SrvMgr_CreateServiceEventTypeOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateServiceEventTypeOperations", objDialog , "Finish")
			Wait 2
			Call Fn_ReadyStatusSync(2)
			If Fn_SISW_SrvMgr_CreateServiceEventTypeOperations=true and objDialog.Exist(2) Then
				objDialog.JavaButton("Finish").Object.click
				wait 1
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceEventTypeOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreateServiceEventTypeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateServiceEventTypeOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function
'********************** Function to perform operations on Create Part Movement dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreatePartMovementOperations
'
''Description		    :	Function to perform operations on Create Part Movement dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicPartMovement : Dictionary object to set Part Movement data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Dim dicPartMovement
'					  		Set dicPartMovement = CreateObject("Scripting.Dictionary")

'							Dim dicContainedInSearch
'					  		Set dicContainedInSearch = CreateObject("Scripting.Dictionary")
'					  		dicContainedInSearch("Name") = "to*"
'					  		dicContainedInSearch("SearchResults_Select") = "abc"
'					  		Set dicPartMovement("ContainedIn") = dicContainedInSearch
'										
'					  		dicPartMovement("Name") = "abc"
'					  		dicPartMovement("ID") = "123"
'					  		dicPartMovement("Description") = "abc desc"
'					  		dicPartMovement("UsageName") = "usageName"
'					  		dicPartMovement("PartMovementType") = "Install"
'					  		dicPartMovement("IsTraceable") = "true"
'					  		dicPartMovement("IsNew") = "true"
'					  		dicPartMovement("ShowAsMaintainedStructure") = "000034/--C (View):000042/--A"

'					  		Dim dicInPhysicalPartsSearch
'					  		Set dicInPhysicalPartsSearch = CreateObject("Scripting.Dictionary")
'					  		dicInPhysicalPartsSearch("AlternateParts")= "000045/--A" 
'					  		dicInPhysicalPartsSearch("SearchType") ="ReplacePhysicalParts"
'					  		Set dicPartMovement("PhysicalParts") = dicInPhysicalPartsSearch

'					  		dicPartMovement("PartNumber") = "123"
'					  		dicPartMovement("SerialNumber") = "20120509"		

'					  		Dim dicUnInstall 
'					  		Set dicUnInstall = CreateObject("Scripting.Dictionary")
'					  		dicUnInstall("sNode")  = "77349-PhyLoc177349"

'					  		Set dicPartMovement("UnInstallLocation") = dicUnInstall

'					  		msgbox Fn_SISW_SrvMgr_CreatePartMovementOperations("Create", dicPartMovement)
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Note : All parameters are same as case "Create"
'					  		dicPartMovement("sButtonName") = ""		
'					  		msgbox Fn_SISW_SrvMgr_CreatePartMovementOperations("Set", dicPartMovement)

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Note : All parameters are same as case "Create"
'					  		dicPartMovement("sButtonName") = "Cancel"		
'					  		msgbox Fn_SISW_SrvMgr_CreatePartMovementOperations("Verify", dicPartMovement)

'History:
'	Developer Name			Date			Rev. No.			Changes Done															Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		05-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		12-Sept-2012			1.0					Added case Set
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		13-Sept-2012			1.0					Added case Verify
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Shweta Rathod		21-Jul-2015				1.0					Modified Case : "Create", "Set" - Modified code to set date.		Vivek
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreatePartMovementOperations(sAction, dicPartMovement)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreatePartMovementOperations"
	Dim objDialog, arrDate, bReturn, iCount,WshShell
	Dim iCnt, iRowCount, iCounter
	Dim dictItems, dictKeys
	
	Fn_SISW_SrvMgr_CreatePartMovementOperations = False
	Set objDialog = Fn_SISW_SrvMgr_GetObject("CreatePartMovement")
	
	'Select menu [ File -> New -> Part Movement... ]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreatePartMovementOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Part Movement...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreatePartMovementOperations",objDialog.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreatePartMovementOperations ] Failed to find [ Create Part Movement ] dialog.")
			Exit Function
		End If
	End If

	objDialog.Resize 500, 850

	Select Case sAction
		Case "Create", "Set"
            'Get the keys & items count from data dictionary.	
			dictItems = dicPartMovement.Items
			dictKeys = dicPartMovement.Keys
			For iCounter = 0 to dicPartMovement.Count - 1
				If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" OR TypeName(dictItems(iCounter))="Date" Then
					Select Case DictKeys(iCounter)
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "PartMovementType"
							Call Fn_List_Select("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog, DictKeys(iCounter), dictItems(iCounter))
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "ActionDate"
							' Date Buttons
'							call Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog , DictKeys(iCounter) & "Button")
'							wait 1
							If lcase(dictItems(iCounter)) = "today" OR lcase(dictItems(iCounter)) = "ok" Then
								dictItems(iCounter)=now
								arrDate = split(dictItems(iCounter)," ")
								Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvMgr_CreatePartMovementOperations", "Set",  objDialog,"ActionDate",arrDate(0))
								wait 2
                                Set WshShell = CreateObject("WScript.Shell")
								WshShell.SendKeys "{TAB}"
								JavaWindow("ServiceManager").JavaWindow("CreatePartMovement").JavaList("ActionTime").Object.setText arrDate(1)+" "+arrDate(2)
								WshShell.SendKeys "{TAB}"
								wait(1)
								Set WshShell = nothing
							Else
								arrDate = split(dictItems(iCounter)," ")
								Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvMgr_CreatePartMovementOperations", "Set",  objDialog,"ActionDate",arrDate(0))
								wait 2
                                Set WshShell = CreateObject("WScript.Shell")
								WshShell.SendKeys "{TAB}"
								wait(1)
								If ubound(arrDate)= 2 then
									JavaWindow("ServiceManager").JavaWindow("CreatePartMovement").JavaList("ActionTime").Object.setText arrDate(1)+" "+arrDate(2)
								else
									JavaWindow("ServiceManager").JavaWindow("CreatePartMovement").JavaList("ActionTime").Object.setText arrDate(1)
								End if
								WshShell.SendKeys "{TAB}"
								Set WshShell = nothing
							End if
						'	bReturn = Fn_SISW_UI_JavaList_Operations("Fn_SISW_SrvMgr_CreatePartMovementOperations","Select",objDialog ,"ActionTime",arrDate(1)+" "+arrDate(2),"","")
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "IsTraceable", "IsApprovedDeviation", "IsExtraToDesign","IsNew"
							objDialog.JavaRadioButton( DictKeys(iCounter)).SetTOProperty "attached text", lcase(cstr(dictItems(iCounter)))
							Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog ,  DictKeys(iCounter))
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Parent", "NeutralPart", "PhysicalPart", "UnInstallLocation"
							' ImageHyperlink  Clear Option
							objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Clear").Select
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "ShowAsMaintainedStructure"
							If dictItems(iCounter) <> ""  Then
								Call Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog , "ShowAsMaintainedStructure")
								If Fn_SISW_SrvMgr_MaintenanceTreeOperations("Select", dictItems(iCounter), "", "") = False Then
									Exit Function
								End If
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "sButtonName"
							' Do Nothing
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case Else
							' Edit Box 
							bReturn = Fn_Edit_Box("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog , DictKeys(iCounter),dictItems(iCounter))
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreatePartMovementOperations ] Failed to set Editbox [ " & DictKeys(iCounter) & " = " & dictItems(iCounter) & " ].")
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					End Select
				Else
					' For ImageHyperlink Add
					objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
					wait 1
					objDialog.JavaMenu("label:=Add...").Select
					wait 1
					Select Case DictKeys(iCounter)
						Case "PhysicalPart"
							Dim PartData
							Set PartData = dictItems(iCounter)
							If Fn_SISW_ABM_SearchDialogOperations("FindAndSelect", PartData("SearchType"), dictItems(iCounter)) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreatePartMovementOperations ] Failed to find [ " & DictKeys(iCounter)  & " ].")
								Set PartData = Nothing
								Exit Function
							End IF
							Set PartData = Nothing
						Case "Parent"
							If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dictItems(iCounter)) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreatePartMovementOperations ] Failed to find [ " & DictKeys(iCounter)  & " ].")
								Exit Function
							End IF
						Case "NeutralPart"
							Dim dicNeutralPart
							Set dicNeutralPart = dictItems(iCounter)
							If Fn_SISW_SrvMgr_SelectNeutralPart("Select", dicNeutralPart("PreferredNeutralPart"), dicNeutralPart("AlternateNeutralParts"),dicNeutralPart("SubstituteNeutralParts"),"","OK") = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreatePartMovementOperations ] Failed to find [ " & DictKeys(iCounter)  & " ].")
								Set dicNeutralPart = Nothing
								Exit Function
							End IF
							Set dicNeutralPart = Nothing
						Case "UnInstallLocation"
							If Fn_SISW_SrvMgr_SelectPhysicalLocationOperations("Select", dictItems(iCounter)) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreatePartMovementOperations ] Failed to find [ " & DictKeys(iCounter)  & " ].")
								Exit Function
							End IF
					End Select
				End If
			Next
			If sAction = "Set" Then
				Fn_SISW_SrvMgr_CreatePartMovementOperations = True
				If dicPartMovement("sButtonName") <> "" Then
						Fn_SISW_SrvMgr_CreatePartMovementOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog ,  dicPartMovement("sButtonName") )
						If Fn_SISW_SrvMgr_CreatePartMovementOperations=true and objDialog.Exist(2) Then
						    If objDialog.JavaButton(dicPartMovement("sButtonName")).GetROProperty("enabled")="1" Then
							   objDialog.JavaButton(dicPartMovement("sButtonName")).Object.click
							End if
						End If
				End If
			Else
				Fn_SISW_SrvMgr_CreatePartMovementOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog , "Finish")
				If Fn_SISW_SrvMgr_CreatePartMovementOperations=true and objDialog.Exist(2) Then
							objDialog.JavaButton("Finish").Object.click
				End If
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			'Get the keys & items count from data dictionary.	
			dictItems = dicPartMovement.Items
			dictKeys = dicPartMovement.Keys
			For iCounter = 0 to dicPartMovement.Count - 1
				If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
					Select Case DictKeys(iCounter)
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "PartMovementType"
							If objDialog.JavaList("PartMovementType").GetROProperty("value") <> dictItems(iCounter) Then
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "PartMovementType_ExistsInList"
							If Fn_UI_ListItemExist("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog , DictKeys(iCounter), dictItems(iCounter)) = False Then
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "IsTraceable", "IsApprovedDeviation", "IsExtraToDesign","IsNew"
							objDialog.JavaRadioButton( DictKeys(iCounter)).SetTOProperty "attached text", lcase(cstr(dictItems(iCounter)))
							If cInt(objDialog.JavaRadioButton("IsTraceable").GetROProperty("value")) <> 1 Then
								Exit function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Parent", "NeutralPart", "PhysicalPart", "UnInstallLocation", "ParentPhysicalElement"
							If objDialog.JavaObject(DictKeys(iCounter) & "ImageHyperlink").Object.getText()  <> dictItems(iCounter) Then
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "ShowAsMaintainedStructure"
							If dictItems(iCounter) <> ""  Then
								Call Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog , "ShowAsMaintainedStructure")
								If Fn_SISW_SrvMgr_MaintenanceTreeOperations("Select", dictItems(iCounter), "", "") = False Then
									Exit Function
								End If
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "sButtonName"
							' Do Nothing
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case Else
							' Edit Box 
							If InStr(DictKeys(iCounter), "Date") > 0 Then
								If CDate(Fn_Edit_Box_GetValue("Fn_SISW_SrvMgr_CreatePartMovementOperations",objDialog, DictKeys(iCounter))) <> CDate(dictItems(iCounter)) Then
									Exit Function
								End If
							Else
								If Fn_Edit_Box_GetValue("Fn_SISW_SrvMgr_CreatePartMovementOperations",objDialog, DictKeys(iCounter)) <> dictItems(iCounter) Then
									Exit Function
								End If
							End If
							
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					End Select
				Else
' 					For ImageHyperlink Add
'					For future use
				End If
			Next
			Fn_SISW_SrvMgr_CreatePartMovementOperations = True
			If dicPartMovement("sButtonName") <> "" Then
					Fn_SISW_SrvMgr_CreatePartMovementOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog ,  dicPartMovement("sButtonName") )
					Call Fn_ReadyStatusSync(2)
					wait 3
					If Fn_SISW_SrvMgr_CreatePartMovementOperations=true and objDialog.Exist(2) Then
						objDialog.JavaButton(dicPartMovement("sButtonName")).Object.click
					End If
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreatePartMovementOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreatePartMovementOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreatePartMovementOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function
'********************** Function to perform operations on Maintenance Tree dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_MaintenanceTreeOperations
'
''Description		    :	Function to perform operations on Maintenance Tree dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. sNode   : Node to select
'					  		3. sColumn : Column Name - for future use
'					  		4. sValue  : Value to be verified  - for future use
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Maintenance Tree dialog should be opened.

''Examples  			:	msgbox Fn_SISW_SrvMgr_MaintenanceTreeOperations("Select", "000034/--C (View):000042/--A", "", "")

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		5-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_MaintenanceTreeOperations(sAction, sNode, sColumn, sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_MaintenanceTreeOperations"
	Dim arrRows, arrNode, iCount, sPath, iNodeCounter, iCnt
	Set objMaintenanceTree = JavaWindow("ServiceManager").JavaWindow("Shell").JavaWindow("MaintenanceTree")
	Fn_SISW_SrvMgr_MaintenanceTreeOperations = False
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_MaintenanceTreeOperations",objMaintenanceTree) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_MaintenanceTreeOperations ] Failed to find [ Maintenance Tree ] Dialog.")
		Exit Function
	End If
	Wait 5

	' Expanding Parent Nodes
	arrRows = split(sNode,"~")
    For iCount = 0 to UBound(arrRows)
		arrNode = split(arrRows(iCount),":")
		For iNodeCounter = 0 to UBound(arrNode)
			sPath = ""
			For iCnt = 0 to iNodeCounter
				If sPath = "" Then
					sPath = arrNode(iCnt)
				Else
					sPath = sPath & ":" & arrNode(iCnt)
				End If
			Next
			bReturn = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SrvMgr_MaintenanceTreeOperations", objMaintenanceTree.JavaTree("Tree"), sPath  ,":","@")
			If bReturn <> False Then
				objMaintenanceTree.JavaTree("Tree").Expand bReturn
				wait 5 ' - by Vrushali [17-Sept-2012 : Build TC100829]
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_MaintenanceTreeOperations ] Failed to find item [ " & sPath & " ].")
				Exit Function
			End If
		Next
	Next

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
				For iCount = 0 to UBound(arrRows)
					bReturn = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SrvMgr_MaintenanceTreeOperations", objMaintenanceTree.JavaTree("Tree"), arrRows(iCount)  ,":","@")
					If bReturn <> False Then
						If iCount = 0 Then
							objMaintenanceTree.JavaTree("Tree").Select  bReturn
						Else
							objMaintenanceTree.JavaTree("Tree").ExtendSelect bReturn
						End If
					Else
						Exit Function
					End If
				Next
				Fn_SISW_SrvMgr_MaintenanceTreeOperations = Fn_Button_Click("Fn_SISW_SrvMgr_MaintenanceTreeOperations", objMaintenanceTree, "OK")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_MaintenanceTreeOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_MaintenanceTreeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_MaintenanceTreeOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objMaintenanceTree = Nothing
End Function
'********************** Function to perform operations on Select Physical Location dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_SelectPhysicalLocationOperations
'
''Description		    :	Function to perform operations on Select Physical Location dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. sNode   : Node to select
'					  		3. sColumn : Column Name - for future use
'					  		4. AddColumn  : Value to be verified  - for future use
'					  		5. RemoveColumn  : Value to be verified  - for future use
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Select Physical Location dialog should be opened.

''Examples  			:	msgbox Fn_SISW_SrvMgr_SelectPhysicalLocationOperations("Select", "000034-Location", "", "","","")

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		5-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_SelectPhysicalLocationOperations(sAction, dicPhysicalLoc)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_SelectPhysicalLocationOperations"
	Dim iCount, iRowCnt
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("Shell").JavaWindow("SelectPhysicalLocation")

	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_SelectPhysicalLocationOperations",objDialog) = False Then
		Exit Function
	End If
	Wait 5

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			iRowCnt =  cInt(objDialog.JavaTable("PhysicalLocations").GetROProperty("rows"))
			For iCount = 0 to iRowCnt - 1
				If objDialog.JavaTable("PhysicalLocations").GetCellData(iCount,"Object") = dicPhysicalLoc("sNode") Then
					objDialog.JavaTable("PhysicalLocations").ClickCell iCount, "Object"
					wait 1
					Exit for
				End If
			Next
			Fn_SISW_SrvMgr_SelectPhysicalLocationOperations = Fn_Button_Click("Fn_SISW_SrvMgr_SelectPhysicalLocationOperations", objDialog, "OK")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SelectPhysicalLocationOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_SelectPhysicalLocationOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_SelectPhysicalLocationOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function
'********************** Function to perform operations on Create Service Discrepancy Type dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations
'
''Description		    :	Function to perform operations on Create Service Discrepancy Type dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicServDiscreType : Dictionary object to set Service Discrepancy Type data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Dim dicServDiscreType
'					  		Set dicServDiscreType = CreateObject("Scripting.Dictionary")

'							Dim dicContainedInSearch
'					  		Set dicContainedInSearch = CreateObject("Scripting.Dictionary")
'					  		dicContainedInSearch("Name") = "to*"
'					  		dicContainedInSearch("SearchResults_Select") = "abc"
'					  		Set dicServDiscreType("ContainedIn") = dicContainedInSearch
'										
'					  		dicServDiscreType("Name") = "abc"
'					  		dicServDiscreType("ID") = "123"
'					  		dicServDiscreType("ActivityNumber") = "1234"
'					  		dicServDiscreType("Description") = "abc desc"
'					  		dicServDiscreType("InitiationDate") = "5-Sep-2012~11:30 AM"
'					  		dicServDiscreType("IsFailure") = "true"
'					  		dicServDiscreType("IsNew") = "true"
'					  		dicServDiscreType("ShowAsMaintainedStructure") = "000034/--C (View):000042/--A"

'					  		Dim dicInPhysicalPartsSearch
'					  		Set dicInPhysicalPartsSearch = CreateObject("Scripting.Dictionary")
'					  		dicInPhysicalPartsSearch("Name") = "part*"
'					  		dicInPhysicalPartsSearch("SearchResults_Select") = "phyPart"
'					  		Set dicServDiscreType("PhysicalPartInProgress") = dicInPhysicalPartsSearch

'					  		Dim dicFaultCodeSearch
'					  		Set dicFaultCodeSearch = CreateObject("Scripting.Dictionary")
'					  		dicFaultCodeSearch("Name") = "fau*"
'					  		dicFaultCodeSearch("SearchResults_Select") = "faultCode"
'					  		Set dicServDiscreType("FaultCode") = dicFaultCodeSearch
'					  		
'					  		dicServDiscreType("RemoveFromFaultCode") = "" - 

'					  		msgbox Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations("Create", dicServDiscreType)

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		6-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations(sAction, dicServDiscreType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations"
	Dim objDialog, arrFaultCode, bReturn, iCount
	Dim iCnt, iRowCount, iCounter, arrDate
	Dim dictItems, dictKeys, sDate
	Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations = False
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("CreateServiceDiscrepancyType")
		
'Select menu [ File -> New -> Service Discrepancy Type... ]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Service Discrepancy...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations ] Failed to find [ Service Discrepancy Type ] dialog.")
			Exit Function
		End If
	End If

	Select Case sAction
		Case "Create"
            'Get the keys & items count from data dictionary.	
			dictItems = dicServDiscreType.Items
			dictKeys = dicServDiscreType.Keys
			For iCounter = 0 to dicServDiscreType.Count - 1
				If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
					Select Case DictKeys(iCounter)
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Severity"
							Call Fn_List_Select("Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations", objDialog, DictKeys(iCounter), dictItems(iCounter))
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "InitiationDate", "DiscoveryDate", "DueDate"

							'	Code added by Chandrakant Tyagi to add Date from JavaEdit as per design Change
							If DictKeys(iCounter) = "InitiationDate" Then
							   objDialog.JavaStaticText("Property_Label").SetTOProperty "label", "Initiation Date:"	
							ElseIf DictKeys(iCounter) = "DiscoveryDate" Then
							   objDialog.JavaStaticText("Property_Label").SetTOProperty "label", "Discovery Date:"	
							ElseIf DictKeys(iCounter) = "DueDate" Then
							   objDialog.JavaStaticText("Property_Label").SetTOProperty "label", "Due Date:"
							End If
							
							If lcase(dictItems(iCounter)) = "today" Then
								aDate = Split(Now," ")
							Else
								aDate = Split(dictItems(iCounter),"~")
							End If	
							objDialog.JavaEdit("DiscoveryDate").Set aDate(0)
							wait 1
							call Fn_KeyBoardOperation("SendKeys", "{TAB}")
							If Ubound(aDate) = 2 Then
								sDate = aDate(1)+" "+aDate(2)
							Else 
								sDate = aDate(1)
							End If
							objDialog.JavaList("Time").Type sDate
						    wait 1
							'call Fn_KeyBoardOperation("SendKeys", "{TAB}")
'						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "IsFailure"
							objDialog.JavaRadioButton( DictKeys(iCounter)).SetTOProperty "attached text", lcase(cstr(dictItems(iCounter)))
							Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations", objDialog ,  DictKeys(iCounter))
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "ContainedIn", "PhysicalPartInProgress"
							' ImageHyperlink  Clear Option
							objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Clear").Select
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "RemoveFromFaultCode"
							iRowCount = cInt(objDialog.JavaTable("FaultCode").GetROProperty("rows"))
							arrFaultCode = split(dictItems(iCounter),"~")
							For iCount = 0 to UBound(arrFaultCode)
								For iCnt = 0 to iRowCount - 1
									If objDialog.JavaTable("FaultCode").GetCellData(iCnt,0) = arrFaultCode(iCount) Then
										objDialog.JavaTable("FaultCode").SelectRow iCnt
										wait 1
										objDialog.JavaStaticText("FaultCodeDropDown").Click 1, 1,"LEFT"
										wait 1
										objDialog.JavaMenu("label:=Remove").Select
										Exit for
									End If
								Next
							Next
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case Else
							' Edit Box 
							bReturn = Fn_Edit_Box("Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations", objDialog , DictKeys(iCounter), dictItems(iCounter))
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations ] Failed to set Editbox [ " & DictKeys(iCounter) & " = " & dictItems(iCounter) & " ].")
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					End Select
				Else
					' For ImageHyperlink Add
					objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
					wait 1
					objDialog.JavaMenu("label:=Add...").Select
					wait 1
					If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dictItems(iCounter)) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations ] Failed to find [ " & DictKeys(iCounter)  & " ].")
						Exit Function
					End IF
				End If
			Next
			Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations", objDialog , "Finish")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateServiceDiscrepancyTypeOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function
'********************** Function to perform operations on Create Customer Contact / Location in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CustomerInformationOperations
'
''Description		    :	Function to perform operations on Create Customer Contact / Location in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. sCaption : Dictionary object to set  Activity Entry Value data.
'					  		3. dicCustomerDetails : Dictionary object to set Customer Data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Dim dicCustomerDetails
'					  		Set dicCustomerDetails = CreateObject("Scripting.Dictionary")
'					  		dicCustomerDetails("Title") = "Mr., Mr."
'					  		dicCustomerDetails("First Name") = "Kou"
'					  		dicCustomerDetails("Last Name") = "Kou"
'					  		dicCustomerDetails("Suffix") = "Kou"
'					  		msgbox Fn_SISW_SrvMgr_CustomerInformationOperations("Create", "Create Customer Contact", dicCustomerDetails)

''Examples  			:	Dim dicCustomerDetails
'					  		Set dicCustomerDetails = CreateObject("Scripting.Dictionary")
'					  		dicCustomerDetails("Title") = "Mr., Mr."
'					  		dicCustomerDetails("Name") = "Kou"
'					  		dicCustomerDetails("Street") = "Sadashiv Peth"
'					  		dicCustomerDetails("City") = "Pune"
'					  		msgbox Fn_SISW_SrvMgr_CustomerInformationOperations("Create", "Create Customer Location", dicCustomerDetails)


'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		6-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CustomerInformationOperations(sAction, sCaption, dicCustomerDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CustomerInformationOperations"
	Dim objDialog, iCounter
	Dim dictItems, dictKeys
	Fn_SISW_SrvMgr_CustomerInformationOperations = False
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("CustomerInfo")

	objDialog.JavaStaticText("Header_Label").SetTOProperty "label", sCaption

	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CustomerInformationOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CustomerInformationOperations ] Failed to find [ " & sCaption & " ] dialog.")
		Exit Function
	End If
	Select Case sAction
		Case "Create"
            dictItems = dicCustomerDetails.Items
			dictKeys = dicCustomerDetails.Keys
			For iCounter = 0 to dicCustomerDetails.Count - 1
                If IsNull(dictKeys(iCounter)) = False Then
					If  dictItems(iCounter) <> "" Then
                        objDialog.JavaStaticText("Field_Label").SetTOProperty  "label", DictKeys(iCounter) & ":"
						If objDialog.JavaEdit("EditBox").Exist(5) Then
							Call Fn_Edit_Box("Fn_SISW_SrvMgr_CustomerInformationOperations", objDialog, "EditBox", dictItems(iCounter))
						ElseIf objDialog.JavaList("List").Exist(5) Then
							If Fn_UI_ListItemExist("Fn_SISW_SrvMgr_CustomerInformationOperations", objDialog, "List", dictItems(iCounter)) Then
								Call Fn_List_Select("Fn_SISW_SrvMgr_CustomerInformationOperations", objDialog, "List", dictItems(iCounter))
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CustomerInformationOperations ] Failed to find[ " & dictItems(iCounter) & " ] in [ " & DictKeys(iCounter) & " ] List.")
								Exit function
							End If
						Else
'							 No Control Found
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CustomerInformationOperations ] Failed to set [ " & dictItems(iCounter) & " ] to [ " & DictKeys(iCounter) & " ].")
							Exit function
						End If
					End If
				End If
			Next
			Fn_SISW_SrvMgr_CustomerInformationOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CustomerInformationOperations", objDialog , "Finish")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CustomerInformationOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CustomerInformationOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CustomerInformationOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function
'********************** Function to perform operations on Create Service Request Type dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations
'
''Description		    :	Function to perform operations on Create Service Request Type dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicServRequestType : Dictionary object to set Service Request Type data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Dim dicServRequestType
'					  		Set dicServRequestType = CreateObject("Scripting.Dictionary")

'							Dim dicProductPhysicalPartsSearch
'					  		Set dicProductPhysicalPartsSearch = CreateObject("Scripting.Dictionary")
'					  		dicProductPhysicalPartsSearch("Name") = "to*"
'					  		dicProductPhysicalPartsSearch("SearchResults_Select") = "abc"
'					  		Set dicServRequestType("ProductPhysicalParts") = dicProductPhysicalPartsSearch
'										
'					  		dicServRequestType("Synopsis") = "abc"
'					  		dicServRequestType("RequestNumber") = "{auto generate}"

'					  		Dim dicCustomerContact
'					  		Set dicCustomerContact = CreateObject("Scripting.Dictionary")
'							dicCustomerContact("Title") = "Mr., Mr."
'					  		dicCustomerContact("First Name") = "Kou"
'					  		dicCustomerContact("Last Name") = "Kou"
'					  		dicCustomerContact("Suffix") = "Kou"
'					  		Set dicServRequestType("CreateCustomerContact") = dicCustomerContact

'					  		Dim dicCustomerLocation
'					  		Set dicCustomerLocation = CreateObject("Scripting.Dictionary")
'							dicCustomerLocation("Title") = "Mr., Mr."
'					  		dicCustomerLocation("Name") = "Kou"
'					  		dicCustomerLocation("Street") = "Sadashiv Peth"
'					  		dicCustomerLocation("City") = "Pune"
'					  		Set dicServRequestType("CreateCustomerLocation") = dicCustomerLocation

'					  		dicServRequestType("Purpose") = "abc desc"
'					  		msgbox Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations("Create", dicServRequestType)

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		6-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations(sAction, dicServRequestType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations"
	Dim objDialog, arrFaultCode, bReturn, iCount
	Dim iCnt, iRowCount, iCounter, arrDate
	Dim dictItems, dictKeys
	Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations = False
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("CreateServiceRequestType")
	
	'Select menu [ File -> New -> Service Request... ]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Service Request...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations ] Failed to find [ Service Discrepancy Type ] dialog.")
			Exit Function
		End If
	End If

	Select Case sAction
		Case "Create"
            'Get the keys & items count from data dictionary.	
			dictItems = dicServRequestType.Items
			dictKeys = dicServRequestType.Keys
			For iCounter = 0 to dicServRequestType.Count - 1
				If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
					Select Case DictKeys(iCounter)
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "RequestNumber"
							'Call Fn_List_Select("Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations", objDialog, DictKeys(iCounter), dictItems(iCounter))
							objDialog.JavaList("RequestNumber").Click 5,5,"LEFT"
							wait(1)
							objDialog.JavaList("RequestNumber").Type dictItems(iCounter)
							wait(1)
							objDialog.JavaList("RequestNumber").DblClick 5,5,"LEFT"
'							wait(1)
'							objDialog.JavaList("RequestNumber").Type " "
							wait(1)
							objDialog.JavaList("RequestNumber").Type dictItems(iCounter)
							wait(1)
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "RequestDate"
						
							If lcase(dictItems(iCounter)) = "today" Or lcase(dictItems(iCounter)) = "ok" then
								arrDate = Split(Now," ")
							Else
								arrDate = split(dictItems(iCounter),"~")
							End If
							objDialog.JavaEdit("RequestDate").Set arrDate(0)
							wait 1
							call Fn_KeyBoardOperation("SendKeys", "{TAB}")
							objDialog.JavaList("Time").Type arrDate(1)
										
					'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "sButtonName"
							'Do Nothing
						Case "CustomerContact", "CustomerLocation"
							' ImageHyperlink  Clear Option
							objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Clear").Select
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "RemoveProductPhysicalParts"
							iRowCount = cInt(objDialog.JavaTable("FaultCode").GetROProperty("rows"))
							arrFaultCode = split(dictItems(iCounter),"~")
							For iCount = 0 to UBound(arrFaultCode)
								For iCnt = 0 to iRowCount - 1
									If objDialog.JavaTable("FaultCode").GetCellData(iCnt,0) = arrFaultCode(iCount) Then
										objDialog.JavaTable("FaultCode").SelectRow iCnt
										wait 1
										objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
										wait 1
										objDialog.JavaMenu("label:=Remove").Select
										Exit for
									End If
								Next
							Next
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case Else
							' Edit Box 
							bReturn = Fn_Edit_Box("Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations", objDialog , DictKeys(iCounter), dictItems(iCounter))
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations ] Failed to set Editbox [ " & DictKeys(iCounter) & " = " & dictItems(iCounter) & " ].")
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					End Select
				Else
					' For ImageHyperlink Add
					Select Case DictKeys(iCounter)
						Case "CreateCustomerContact"
							objDialog.JavaStaticText("CustomerContactDropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Create...").Select
							wait 1

							If Fn_SISW_SrvMgr_CustomerInformationOperations("Create", "Create Customer Contact", dictItems(iCounter)) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations ] Failed to [ Create Customer Contact ].")
								Exit Function
							End IF

						Case "CreateCustomerLocation"
							objDialog.JavaStaticText("CustomerLocationDropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Create...").Select
							wait 1

							If Fn_SISW_SrvMgr_CustomerInformationOperations("Create", "Create Customer Location", dictItems(iCounter)) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations ] Failed to [ Create Customer Location ].")
								Exit Function
							End IF

						Case Else
							objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Add...").Select
							wait 1
							If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dictItems(iCounter)) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations ] Failed to find [ " & DictKeys(iCounter)  & " ].")
								Exit Function
							End IF
					End Select
				End If
			Next
			If dicServRequestType("sButtonName") <> "" Then
				Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog,dicServRequestType("sButtonName"))
			Else
				Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations =True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateServiceRequestTypeOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function
'********************** Function to perform operations on Create Service Catalog Type dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations
'
''Description		    :	Function to perform operations on Create Service Catalog Type dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicServCatalogType : Dictionary object to set Service Catalog Type data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Dim dicServCatalogType
'					  		Set dicServCatalogType = CreateObject("Scripting.Dictionary")
'					  		dicServCatalogType("Name") = "name1"

'							Dim dicNeutralPartSearch
'					  		Set dicNeutralPartSearch = CreateObject("Scripting.Dictionary")
'					  		dicNeutralPartSearch("Name") = "to*"
'					  		dicNeutralPartSearch("SearchResults_Select") = "abc"
'					  		Set dicServCatalogType("NeutralPart") = dicNeutralPartSearch
'					  		msgbox Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations("Create", dicServCatalogType)

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		6-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations(sAction, dicServCatalogType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations"
	Dim objDialog
	Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations = False
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("CreateServiceCatalogType")
	
	'Select menu [ File -> New -> Service Request... ]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Service Catalog...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations ] Failed to find [ Service Catalog Type ] dialog.")
			Exit Function
		End If
	End If

	Select Case sAction
		Case "Create"
			' Name
			If dicServCatalogType("Name") <> ""  Then
				'Call Fn_List_Select("Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations", objDialog,"Name",dicServCatalogType("Name"))
				Call Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations", "Set",  objDialog,"Name", dicServCatalogType("Name"))
			End If

			'Neutral Part  ImageHyperlink  Clear Option
			If TypeName(dicServCatalogType("NeutralPart")) = "String"  Then
				objDialog.JavaStaticText("NeutralPartDropDown").Click 1, 1,"LEFT"
				wait 1
				objDialog.JavaMenu("label:=Clear").Select
			Else
				objDialog.JavaStaticText("NeutralPartDropDown").Click 1, 1,"LEFT"
				wait 1
				objDialog.JavaMenu("label:=Add...").Select
				wait 1
				If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dicServCatalogType("NeutralPart")) = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations ] Failed to find [ Neutral Part ].")
					Exit Function
				End IF
			End If
            Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations", objDialog , "Finish")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateServiceCatalogTypeOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objDialog = Nothing
End Function
'********************** Function to perform operations on Create Fault Code Type dialog in Service Manager ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations
'
''Description		    :	Function to perform operations on Create Fault Code Type dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicFaultCodeType : Dictionary object to set Fault Code Type data.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Dim dicFaultCodeType
'					  		Set dicFaultCodeType = CreateObject("Scripting.Dictionary")
'					  		dicFaultCodeType("Name") = "name1"
'					  		dicFaultCodeType("Description") = "name desc"

'							Dim dicServiceDiscrepancySearch
'					  		Set dicServiceDiscrepancySearch = CreateObject("Scripting.Dictionary")
'					  		dicServiceDiscrepancySearch("Name") = "s*"
'					  		dicServiceDiscrepancySearch("SearchResults_Select") = "servicediscrepancy"
'					  		Set dicFaultCodeType("NeutralPart") = dicServiceDiscrepancySearch
'					  		msgbox Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations("Create", dicFaultCodeType)

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		6-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations(sAction, dicFaultCodeType)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations"
	Dim objDialog
	Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations = False
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("CreateFaultCodeType")
	
	'Select menu [ File -> New -> Service Request... ]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
		Call Fn_MenuOperation("Select","File:New:Fault Code...")
		Call  Fn_ReadyStatusSync(3)
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations",objDialog.JavaStaticText("Header_Label"))=False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations ] Failed to find [ Fault Code Type ] dialog.")
			Exit Function
		End If
	End If

	Select Case sAction
		Case "Create"
			'Service Discrepancy ImageHyperlink  Clear Option
			If TypeName(dicFaultCodeType("ServiceDiscrepancy")) <> "Empty"  Then
				If TypeName(dicFaultCodeType("ServiceDiscrepancy")) = "String"  Then
					If dicFaultCodeType("ServiceDiscrepancy") <> "" Then
						objDialog.JavaStaticText("ServiceDiscrepancyDropDown").Click 1, 1,"LEFT"
						wait 1
						objDialog.JavaMenu("label:=Clear").Select
					End If
				Else
					objDialog.JavaStaticText("ServiceDiscrepancyDropDown").Click 1, 1,"LEFT"
					wait 1
					objDialog.JavaMenu("label:=Add...").Select
					wait 1
					If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dicFaultCodeType("ServiceDiscrepancy")) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations ] Failed to find [ Service Discrepancy ].")
						Exit Function
					End IF
				End If
			End If

			' Name
			If dicFaultCodeType("Name") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations", objDialog,"Name",dicFaultCodeType("Name"))
			End If

			' Description
			If dicFaultCodeType("Description") <> ""  Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations", objDialog,"Description",dicFaultCodeType("Description"))
			End If

			Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations", objDialog , "Finish")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateFaultCodeTypeOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function
'********************** Fn_SISW_SrvMgr_SelectNeutralPart ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_SelectNeutralPart
'
''Description		    :	Function to perform operations on Select Neutral Part dialog in Service Manager
'
''Parameters		    :	1. sAction					: Action need to perform
'					  		2. sPreferredNeutralPart 	: Preferred Neutral Part to be selected
'					  		3. sAlternateNeutralParts	: Alternate Neutral Part to be selected
'					  		4. SSubstituteNeutralParts	: Substitute Neutral Part to be selected
'					  		5. dicManagecolumns			: For future use
'					  		6. sButtonName				: Button Name to be clicked.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Service Manager perspective should be opened.

''Examples  			:	Call Fn_SISW_SrvMgr_SelectNeutralPart("Select", "", "","","","OK")

'History:
'	Developer Name			Date			Rev. No.	Reviewer			Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vrushali Wani		27-Sept-2012			1.0			Koustubh Watwe		Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_SelectNeutralPart(sAction, sPreferredNeutralPart, sAlternateNeutralParts, SSubstituteNeutralParts, dicManagecolumns, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_SelectNeutralPart"
   	Dim objSelectNeutralPart, sTitle, iCnt, bFlag, iRowCount, aSearchCri , aFieldSet
	Set objSelectNeutralPart = JavaWindow("ServiceManager").JavaWindow("Shell").JavaWindow("SelectNeutralPart")
	Fn_SISW_SrvMgr_SelectNeutralPart = False
	
	If Fn_UI_ObjectExist("Fn_ABM_AssignLot",objSelectNeutralPart) = False then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Failed to find [ Select Neutral Part ] dialog.")
		Set objSelectNeutralPart = Nothing
		Exit function
	End If

	Select Case sAction
		Case "Select"
			If sPreferredNeutralPart <> "" Then
				iRowCount = cInt(objSelectNeutralPart.JavaTable("PreferredNeutralPart").GetROProperty("rows"))
				bFlag = False
				For iCnt = 0 to iRowCount -1
					If objSelectNeutralPart.JavaTable("PreferredNeutralPart").GetCellData(iCnt,0) = sPreferredNeutralPart  then
						objSelectNeutralPart.JavaTable("PreferredNeutralPart").ClickCell iCnt, 0
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Successfully selected [ Preferred Neutral Part = " & sPreferredNeutralPart & "].")
						bFlag = True
						Exit for
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Failed to select [ Preferred Neutral Part = " & sPreferredNeutralPart & "].")
					Set objSelectNeutralPart = Nothing
					Exit function
				End If
			End If

			If sAlternateNeutralParts <> ""	Then
				iRowCount = cInt(objSelectNeutralPart.JavaTable("AlternateNeutralParts").GetROProperty("rows"))
				bFlag = False
				For iCnt = 0 to iRowCount -1
					If objSelectNeutralPart.JavaTable("AlternateNeutralParts").GetCellData(iCnt,0) = sAlternateNeutralParts  then
						objSelectNeutralPart.JavaTable("AlternateNeutralParts").ClickCell iCnt, 0
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Successfully selected [ Alternate Neutral Part = " & sAlternateNeutralParts & "].")
						bFlag = True
						Exit for
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Failed to select [ Alternate Neutral Part = " & sAlternateNeutralParts & "].")
					Set objSelectNeutralPart = Nothing
					Exit function
				End If
			End If

			If SSubstituteNeutralParts <> "" Then
				bFlag = False
				iRowCount = cInt(objSelectNeutralPart.JavaTable("SubstituteNeutralParts").GetROProperty("rows"))
				For iCnt = 0 to iRowCount -1
					If objSelectNeutralPart.JavaTable("SubstituteNeutralParts").GetCellData(iCnt,0) = SSubstituteNeutralParts  then
						objSelectNeutralPart.JavaTable("SubstituteNeutralParts").ClickCell iCnt, 0
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Successfully selected [ Substitute Neutral Part = " & SSubstituteNeutralParts & "].")
						bFlag = True
						Exit for
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Failed to select [ Substitute Neutral Part = " & SSubstituteNeutralParts & "].")
					Set objSelectNeutralPart = Nothing
					Exit function
				End If
			End If

			If sButtonName <> "" Then
				Call Fn_Button_Click("Fn_SISW_SrvMgr_SelectNeutralPart", objSelectNeutralPart, sButtonName)
			End If
			Fn_SISW_SrvMgr_SelectNeutralPart = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_SelectNeutralPart <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_SelectNeutralPart ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objSelectNeutralPart = Nothing
End Function
'********************** Fn_SISW_SrvMgr_LogEntriesOperations ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_LogEntriesOperations
'
''Description		    :	Function to perform operations on Create Fault Code Type dialog in Service Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. sNode : ":" separated Node path
'					  		3. sColumn : Column Name
'					  		4. sValue : Value to be verified
'					  		5. sPopupMenu : for future use.
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Log Entries panel should be open in Service Manager perspective.

''Examples  			:	Fn_SISW_SrvMgr_LogEntriesOperations("Select", "d:", "", "", "")
''Examples  			:	Fn_SISW_SrvMgr_LogEntriesOperations("Expand", "d:", "", "", "")
''Examples  			:	Fn_SISW_SrvMgr_LogEntriesOperations("CellVerify", "d:", "Value", "1", "")

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		27-Sept-2012			1.0					Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_LogEntriesOperations(sAction, sNode, sColumn, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_LogEntriesOperations"
	Dim objSelectType, objTreeObjects, iCnt, objTree, bFlag, sActValue
	bFlag = False
	Fn_SISW_SrvMgr_LogEntriesOperations = False

	call Fn_SISW_UI_RACTabFolderWidget_Operation("Select","Log Entries","")

	' getting Log Entry Tree object
	Set objSelectType = description.Create()
	objSelectType("Class Name").value = "JavaTree"
	Set objTreeObjects= JavaWindow("DefaultWindow").ChildObjects(objSelectType)

	For iCnt = 0 to   objTreeObjects.count -1
		If cInt(objTreeObjects(iCnt).GetROProperty("columns_count")) > 0 Then
			If objTreeObjects(iCnt).GetColumnHeader(0) = "Log Entry" Then
				Set objTree = objTreeObjects(iCnt)
				bFlag = True
				Exit For
			End If
		End If
	Next

	If bFlag = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_LogEntriesOperations ] Failed to find Log Entries tree object.")
		Set objTree = Nothing
		Set objTreeObjects = Nothing
		Set objSelectType = Nothing
		Exit function
	End If

	Select Case sAction
		Case "Select"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SrvMgr_LogEntriesOperations", objTree, sNode, "", "")
			If iPath <> False Then
				objTree.select iPath
				Fn_SISW_SrvMgr_LogEntriesOperations = True
			End If
		Case "Expand"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SrvMgr_LogEntriesOperations", objTree, sNode, "", "")
			If iPath <> False Then
				objTree.Expand iPath
				Fn_SISW_SrvMgr_LogEntriesOperations = True
			End If
		Case "CellVerify"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_SISW_SrvMgr_LogEntriesOperations", objTree, sNode, "", "")
			If iPath <> False Then
				sActValue = objTree.GetColumnValue(iPath,sColumn)
				If isNumeric(sActValue) Then
					sActValue = cstr(cInt(sActValue ))
				End If
				If isNumeric(sValue) Then
					sValue = cstr(cInt(sValue ))
				End If
				If sActValue = sValue Then
					Fn_SISW_SrvMgr_LogEntriesOperations = True
				end If
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_LogEntriesOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_LogEntriesOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_LogEntriesOperations ] Successfully executed with case [ " & sAction & " ].")
	End If

	Set objTree = Nothing
	Set objTreeObjects = Nothing
	Set objSelectType = Nothing
End Function
'********************** Fn_SISW_SrvMgr_ServiceOfferingOperations ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_ServiceOfferingOperations
'
''Description		    :	Function to perform operations on Service Offering dialog in Service Request Manager
'
''Parameters		    :	1. sAction : Action need to perform
'					  		2. dicServiceOffering : dictionary object
'								
'Return Value		    :  	True / False
'
'Pre-requisite		    :	Log Entries panel should be open in Service Manager perspective.

''Examples  			:	Dim dicServiceOffering
'					  		Set dicServiceOffering = CreateObject("Scripting.Dictionary")
'					  		dicServiceOffering("SelectName") = "*"
'					  		dicServiceOffering("DeselectName") = ""
'					  		dicServiceOffering("Name") = "Name"
'					  		dicServiceOffering("ServiceOfferingNumber") = "123"
'					  		dicServiceOffering("ServiceCode") = "12"
'					  		Call Fn_SISW_SrvMgr_ServiceOfferingOperations("Select", dicServiceOffering)
'-------------------------------------------------------------------------------------------------------------------
''Examples  			:	Dim dicServiceOffering
'					  		Set dicServiceOffering = CreateObject("Scripting.Dictionary")
'					  		dicServiceOffering("Name") = "Name"
'					  		dicServiceOffering("ServiceCatalog") = "Clear" 
'							or
'					  		Dim dicSearch
'					  		Set dicSearch = CreateObject("Scripting.Dictionary")
'					  		dicSearch("Name") = "to*"
'					  		dicSearch("SearchResults_Select") = "abc"
'					  		Set dicServiceOffering("ServiceCatalog") = dicSearch

'					  		dicServiceOffering("ServiceOfferingNumber") = "123"
'					  		dicServiceOffering("ServiceCode") = "12"
'					  		Call Fn_SISW_SrvMgr_ServiceOfferingOperations("Create", dicServiceOffering)

'History:
'	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		08-Oct-2012			1.0					Created
'	Vrushali  Wani      	18-Oct-2012			2.0                  Modified            added code to work with Search Result  	
'-------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_ServiceOfferingOperations(sAction, dicServiceOffering)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_ServiceOfferingOperations"
	Dim objDialog, iCnt, iRowCnt
	Fn_SISW_SrvMgr_ServiceOfferingOperations = False
	
	Select Case sAction
		Case "Create"
			Set objDialog = JavaWindow("ServiceManager").JavaWindow("CreateServiceOfferingType")
			If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog.JavaStaticText("Header_Label")) = False Then
				Call Fn_MenuOperation("Select","File:New:Service Offering...")
				If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog.JavaStaticText("Header_Label")) = False Then
					Exit function
				End If
			End If
			If dicServiceOffering("Name") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "Name", dicServiceOffering("Name"))
			End If

			If Typename(dicServiceOffering("ServiceCatalog")) <> "String" Then
					objDialog.JavaStaticText("ServiceCatalogDropDown").Click 1, 1,"LEFT"
					wait 1
					objDialog.JavaMenu("label:=Add...").Select
					wait 1
					If Fn_SISW_SrvMgr_SearchOperations("SearchAndSelect", dicServiceOffering("ServiceCatalog")) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_ServiceOfferingOperations ] Failed to find [ Service Catalog ].")
						Exit Function
					End If
			Else
				If dicServiceOffering("ServiceCatalog") <> "" Then
					objDialog.JavaStaticText("ServiceCatalogDropDown").Click 1, 1,"LEFT"
					wait 1
					objDialog.JavaMenu("label:=Clear").Select
				End If
			End If

			If dicServiceOffering("ServiceOfferingNumber") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "ServiceOfferingNumber", dicServiceOffering("ServiceOfferingNumber"))
			End If

			If dicServiceOffering("ServiceCode") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "ServiceCode", dicServiceOffering("ServiceCode"))
			End If

			If dicServiceOffering("Narrative") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "Narrative", dicServiceOffering("Narrative"))
			End If
			Fn_SISW_SrvMgr_ServiceOfferingOperations = Fn_Button_Click("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "Finish")
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			Set objDialog = JavaWindow("ServiceManager").JavaWindow("Shell").JavaWindow("ServiceOfferings")
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			If cBool(objDialog.JavaObject("EnterSearchValuesTwistie").Object.isExpanded()) Then
				' closing twistie object
				objDialog.JavaObject("EnterSearchValuesTwistie").Click 1, 1,"LEFT"
			End If
			' expanding twistie object
			objDialog.JavaObject("EnterSearchValuesTwistie").Click 1, 1,"LEFT"
			' Closing twistie object
			objDialog.JavaObject("EnterSearchValuesTwistie").Click 1, 1,"LEFT"
			' expanding twistie object
			objDialog.JavaObject("EnterSearchValuesTwistie").Click 1, 1,"LEFT"



			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			If dicServiceOffering("SelectName") <> "" Then
				iRowCnt = cInt(objDialog.JavaTable("EnterSearchValues").GetROProperty("rows"))
				For iCnt = 0 to iRowCnt - 1
					If objDialog.JavaTable("EnterSearchValues").GetCellData (iCnt, "Name") = dicServiceOffering("SelectName") Then
						objDialog.JavaTable("EnterSearchValues").Object.getItem(iCnt).setChecked True, True
						Exit for
					end If
				Next
			End If

			If dicServiceOffering("DeselectName") <> "" Then
				iRowCnt = cInt(objDialog.JavaTable("EnterSearchValues").GetROProperty("rows"))
				For iCnt = 0 to iRowCnt - 1
					If objDialog.JavaTable("EnterSearchValues").GetCellData (iCnt, "Name") = dicServiceOffering("DeselectName") Then
						objDialog.JavaTable("EnterSearchValues").Object.getItem(iCnt).setChecked False, True
						Exit for
					end If
				Next
			End If

			If dicServiceOffering("Name") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "Name", dicServiceOffering("Name"))
			End If

			If dicServiceOffering("ServiceOfferingNumber") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "ServiceOfferingNumber", dicServiceOffering("ServiceOfferingNumber"))
			End If

			If dicServiceOffering("ServiceCode") <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "ServiceCode", dicServiceOffering("ServiceCode"))
			End If

			If cBool(objDialog.JavaObject("SearchResultsTwistie").Object.isExpanded()) Then
				' closing twistie object
				objDialog.JavaObject("SearchResultsTwistie").Click 1, 1,"LEFT"
			End If
			' expanding twistie object
			objDialog.JavaObject("SearchResultsTwistie").Click 1, 1,"LEFT"

				Call  Fn_UI_JavaToolbar_Press("",objDialog, "SearchResultsToolbar","Find")	

			' Closing twistie object
			objDialog.JavaObject("SearchResultsTwistie").Click 1, 1,"LEFT"
			' expanding twistie object
			objDialog.JavaObject("SearchResultsTwistie").Click 1, 1,"LEFT"

			 iCnt =objDialog.JavaTable("SearchResults").GetROProperty("rows")		
					For i= 0 to iCnt-1				
						If 	  cStr(objDialog.JavaTable("SearchResults").GetCellData(i,"Name")) = dicServiceOffering("Name") Then					
							 objDialog.JavaTable("SearchResults").Object.getItem(1).setChecked True, True
						End If					
					Next

			Fn_SISW_SrvMgr_ServiceOfferingOperations = Fn_Button_Click("Fn_SISW_SrvMgr_ServiceOfferingOperations", objDialog, "OK")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_ServiceOfferingOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_ServiceOfferingOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_ServiceOfferingOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function
'**************************************** Fn_SISW_SrvMgr_CreateRequestedActivityOperations ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_CreateRequestedActivityOperations
'
''Description		    :  	Function to create Requested Activity in Service Request Manager

''Parameters		    :	1. sAction : Object Handle name
'					  		2. sServiceRequest : dictionary object
'					  		3. dicProductphysicalPart : dictionary object
'					  		4. ServiceOffering : dictionary object
'					  		5. sSynopsis : dictionary object
'					  		6. sNeededByDate : dictionary object
'					  		7. sNarrative : dictionary object

''Return Value		    :  	True \ False
'
''Examples		     	:	
'							Dim dicServiceOffering
'					  		Set dicServiceOffering = CreateObject("Scripting.Dictionary")
'					  		dicServiceOffering("SelectName") = "*"
'					  		dicServiceOffering("DeselectName") = ""
'					  		dicServiceOffering("Name") = "Name"
'					  		dicServiceOffering("ServiceOfferingNumber") = "123"
'					  		dicServiceOffering("ServiceCode") = "12"
'					  		Call Fn_SISW_SrvMgr_CreateRequestedActivityOperations("Create", "", "",dicServiceOffering,"synop","08-10-2012~10:30PM","narat")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 08-Oct-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateRequestedActivityOperations(sAction, sServiceRequest, sProductphysicalPart, ServiceOffering, sSynopsis, sNeededByDate, sNarrative)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateRequestedActivityOperations"
	Dim objDialog, arrDate, objSearchResults, iRowCnt, iCnt, bFlag
	Dim arrServiceOffering, iArrCnt
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("CreateRequestedActivityType")
	Fn_SISW_SrvMgr_CreateRequestedActivityOperations = False
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objDialog.JavaStaticText("Header_Label")) = False Then
		Call Fn_MenuOperation("Select", "File:New:Requested Activity...")
		If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objDialog.JavaStaticText("Header_Label")) = False Then
			Exit Function
		End If
	End If
	Select Case sAction
		Case "Create", "SetData"

			If sProductphysicalPart <> "" Then
				' For ImageHyperlink Add
				objDialog.JavaStaticText("ProductPhysicalPartDropDown").Click 1, 1,"LEFT"
				wait 1
				objDialog.JavaMenu("label:=Add...").Select
				wait 1
				Set objSearchResults = JavaWindow("ServiceManager").JavaWindow("Shell").JavaWindow("SearchResults")
				If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objSearchResults) = False Then Exit function
				iRowCnt = cInt(objSearchResults.JavaTable("Table").GetROProperty("rows"))
				For iCnt = 0 to iRowCnt - 1
					If objSearchResults.JavaTable("Table").GetCellData(iCnt, 0) = sProductphysicalPart Then
						objSearchResults.JavaTable("Table").SelectCell iCnt, 0
						Exit for
					End If
				Next
				If iCnt = iRowCnt Then
					Exit function
				End If
				Call Fn_Button_Click("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objSearchResults, "OK")
			End If

			If TypeName(ServiceOffering) = "String" Then
				arrServiceOffering = split(ServiceOffering,"~")
				iRowCnt = cInt(objDialog.JavaTable("ServiceOfferings").GetROProperty("rows"))
				For iArrCnt = 0 to UBound(arrServiceOffering)
					bFlag = False
					For iCnt = 0 to iRowCnt - 1
						If objDialog.JavaTable("ServiceOfferings").GetCellData (iCnt, 0) = arrServiceOffering(iArrCnt) Then
							bFlag = True
							objDialog.JavaTable("ServiceOfferings").SelectCell iCnt,  0
							objDialog.JavaStaticText("ServiceOfferingsDropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Remove").Select
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateRequestedActivityOperations ] Successfully selected " &  arrServiceOffering(iArrCnt) & " and removed.")
							Exit for
						End If
					Next
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateRequestedActivityOperations ] Failed to select " &  arrServiceOffering(iArrCnt) & ".")
						Exit function
					End If
				Next
			Else
				bFlag = False
				objDialog.JavaStaticText("ServiceOfferingsDropDown").Click 1, 1,"LEFT"
				wait 1
				objDialog.JavaMenu("label:=Add...").Select
				wait 1
				bFlag = Fn_SISW_SrvMgr_ServiceOfferingOperations("Select", ServiceOffering)
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateRequestedActivityOperations ] Failed to select Service Offering.")
					Exit function
				End If
			End If

			If sSynopsis <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objDialog, "Synopsis", sSynopsis)
			End If
			' Needed by
			If sNeededByDate <> "" Then
                ' Date Buttons
					call Fn_Button_Click("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objDialog ,  "NeededByDateButton")
					wait 1
					If lcase(sNeededByDate) = "today" Then
						Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_CreateRequestedActivityOperations","Today", "")
					ElseIf lcase(sNeededByDate) = "ok" Then
						Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_CreateRequestedActivityOperations","OK", "")									
					Else
						arrDate = split(sNeededByDate,"~")
						Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_CreateRequestedActivityOperations",arrDate(0), arrDate(1))
					End IF
			End If

			'Narrative
			If sNarrative <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objDialog, "Narrative", sNarrative)
			End If
			If sAction <>  "SetData" Then
				Fn_SISW_SrvMgr_CreateRequestedActivityOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objDialog, "Finish")
			Else
				Fn_SISW_SrvMgr_CreateRequestedActivityOperations = True
			End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify", "VerifySet"
			If sProductphysicalPart <> "" Then
				' For ImageHyperlink Add
				If objDialog.JavaObject("ProductphysicalPartImageHyperlink").Object.getText()  <> sProductphysicalPart Then
					Exit Function
				End If
			End If

			If ServiceOffering <> "" Then
				arrServiceOffering = split(ServiceOffering,"~")
				iRowCnt = cInt(objDialog.JavaTable("ServiceOfferings").GetROProperty("rows"))
				For iArrCnt = 0 to UBound(arrServiceOffering)
					bFlag = False
					For iCnt = 0 to iRowCnt - 1
						If objDialog.JavaTable("ServiceOfferings").GetCellData (iCnt, 0) = arrServiceOffering(iArrCnt) Then
							bFlag = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateRequestedActivityOperations ] Successfully verified " &  arrServiceOffering(iArrCnt) & ".")
							Exit for
						End If
					Next
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateRequestedActivityOperations ] Failed to verify " &  arrServiceOffering(iArrCnt) & ".")
						Exit function
					End If
				Next
			End If

			If sSynopsis <> "" Then
				If Fn_Edit_Box_GetValue("Fn_SISW_SrvMgr_CreateRequestedActivityOperations",objDialog, "Synopsis") <> sSynopsis Then
					Exit Function
				End If
			End If
			' Needed by
			If sNeededByDate <> "" Then
                ' Date Buttons
				If Fn_Edit_Box_GetValue("Fn_SISW_SrvMgr_CreateRequestedActivityOperations",objDialog, "NeededBy") <> sNeededByDate Then
					Exit Function
				End If
			End If

			'Narrative
			If sNarrative <> "" Then
				If Fn_Edit_Box_GetValue("Fn_SISW_SrvMgr_CreateRequestedActivityOperations",objDialog, "Narrative") <> sNarrative Then
					Exit Function
				End If
			End If

			If 	sAction = "VerifySet" Then
				Fn_SISW_SrvMgr_CreateRequestedActivityOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objDialog, "Finish")
			Else
			    Fn_SISW_SrvMgr_CreateRequestedActivityOperations = Fn_Button_Click("Fn_SISW_SrvMgr_CreateRequestedActivityOperations", objDialog, "Cancel")
			End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_CreateRequestedActivityOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_CreateRequestedActivityOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_CreateRequestedActivityOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
	Set objSearchResults = Nothing
End Function
'**************************************** Fn_SISW_SrvMgr_TimeAndCostTotalOperations ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_TimeAndCostTotalOperations
'
''Description		    :  	Function to perform operations on Time and Cost dialog in Service Request Manager

''Parameters		    :	1. sAction : Object Handle name
'					  		2. sEstimatedLaborCost : Estimated Labor Cost
'					  		3. sEstimatedLaborHours : Estimated Labor Hours
'					  		4. sEstimatedMaterialCost : Estimated Material Cost
'					  		5. sEstimatedTotalCost : Estimated Total Cost
'					  		6. sActualLaborCost : Actual Labor Cost
'					  		7. sActualLaborHours : Actual Labor Hours
'					  		8. sActualMaterialCost : Actual Material Cost
'					  		9. sActualTotalCost : Actual Total Cost
'					  		10. sButtonName : Button Name ( "Finish" / "Cancel" / "" )

''Return Value		    :  	True \ False
'
''Examples		     	:	Call Fn_SISW_SrvMgr_TimeAndCostTotalOperations("Set", "500.0", "6", "500", "3500", "500.0", "6", "500", "3500", "Finish")
''Examples		     	:	Call Fn_SISW_SrvMgr_TimeAndCostTotalOperations("Verify", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "0.0", "Cancel")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 10-Oct-2012		1.0								Created	
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_TimeAndCostTotalOperations(sAction, sEstimatedLaborCost, sEstimatedLaborHours, sEstimatedMaterialCost, sEstimatedTotalCost, sActualLaborCost, sActualLaborHours, sActualMaterialCost, sActualTotalCost, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_TimeAndCostTotalOperations"
	Dim objDialog
	Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("TimeAndCostTotal")

	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog.JavaStaticText("Header_Label")) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_TimeAndCostTotalOperations ] Failed to find Time And Cost Total window.")
		Set objDialog = Nothing
		Exit Function
	End If

	Select Case sAction
		Case "Set"
			If sEstimatedLaborCost <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, "EstimatedLaborCost", sEstimatedLaborCost)
			End If
			If sEstimatedLaborHours <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, "EstimatedLaborHours", sEstimatedLaborHours)
			End If
			If sEstimatedMaterialCost <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, "EstimatedMaterialCost", sEstimatedMaterialCost)
			End If
			If sEstimatedTotalCost <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, "EstimatedTotalCost", sEstimatedTotalCost)
			End If
			If sActualLaborCost <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, "ActualLaborCost", sActualLaborCost)
			End If
			If sActualLaborHours <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, "ActualLaborHours", sActualLaborHours)
			End If
			If sActualMaterialCost <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, "ActualMaterialCost", sActualMaterialCost)
			End If
			If sActualTotalCost <> "" Then
				Call Fn_Edit_Box("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, "ActualTotalCost", sActualTotalCost)
			End If
			Fn_SISW_SrvMgr_TimeAndCostTotalOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			Fn_SISW_SrvMgr_TimeAndCostTotalOperations = True
			If sEstimatedLaborCost <> "" Then
				If cstr(cInt(objDialog.JavaEdit("EstimatedLaborCost").GetROProperty("value"))) <> cstr(cInt(sEstimatedLaborCost)) Then
					Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
				End If
			End If
			If sEstimatedLaborHours <> "" Then
				If cstr(cInt(objDialog.JavaEdit("EstimatedLaborHours").GetROProperty("value"))) <> cstr(cInt(sEstimatedLaborHours)) Then
					Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
				End If
			End If
			If sEstimatedMaterialCost <> "" Then
				If cstr(cInt(objDialog.JavaEdit("EstimatedMaterialCost").GetROProperty("value"))) <> cstr(cInt(sEstimatedMaterialCost)) Then
					Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
				End If
			End If
			If sEstimatedTotalCost <> "" Then
				If cstr(cInt(objDialog.JavaEdit("EstimatedTotalCost").GetROProperty("value"))) <> cstr(cInt(sEstimatedTotalCost)) Then
					Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
				End If
			End If
			If sActualLaborCost <> "" Then
				If cstr(cInt(objDialog.JavaEdit("ActualLaborCost").GetROProperty("value"))) <> cstr(cInt(sActualLaborCost)) Then
					Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
				End If
			End If
			If sActualLaborHours <> "" Then
				If cstr(cInt(objDialog.JavaEdit("ActualLaborHours").GetROProperty("value"))) <> cstr(cInt(sActualLaborHours)) Then
					Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
				End If
			End If
			If sActualMaterialCost <> "" Then
				If cstr(cInt(objDialog.JavaEdit("ActualMaterialCost").GetROProperty("value"))) <> cstr(cInt(sActualMaterialCost)) Then
					Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
				End If
			End If
			If sActualTotalCost <> "" Then
				If cstr(cInt(objDialog.JavaEdit("ActualTotalCost").GetROProperty("value"))) <> cstr(cInt(sActualTotalCost)) Then
					Fn_SISW_SrvMgr_TimeAndCostTotalOperations = False
				End If
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_TimeAndCostTotalOperations ] Invalid case [ " & sAction & " ].")
	End Select
	If sButtonName <> "" Then
		Call Fn_Button_Click("Fn_SISW_SrvMgr_TimeAndCostTotalOperations", objDialog, sButtonName)
	End If

	If Fn_SISW_SrvMgr_TimeAndCostTotalOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_TimeAndCostTotalOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function
'**************************************** Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_DelegateRequestedActivitiesType
'
''Description		    :  	Function to perform operations on Delegate Requested Activities Type dialog in Service Request Manager

''Parameters		    :	1. sAction : Object Handle name
'					  		2. dicDelegateReq

''Return Value		    :  	True \ False
'
''Examples		     	:	Dim dicDelegateReq
'					  		Set dicDelegateReq = CreateObject("Scripting.Dictionary")

'					  		dicDelegateReq("Synopsis") = "abc"
'					  		dicDelegateReq("RequestNumber") = "{auto generate}"
'					  		dicDelegateReq("PerformsServiceRequest") = "RA000007/A;1-A1"
'					  		dicDelegateReq("PerformRequestedActivities") = "New Delegate Service Request"

'					  		Dim dicCustomerContact
'					  		Set dicCustomerContact = CreateObject("Scripting.Dictionary")
'							dicCustomerContact("Title") = "Mr., Mr."
'					  		dicCustomerContact("First Name") = "Kou"
'					  		dicCustomerContact("Last Name") = "Kou"
'					  		dicCustomerContact("Suffix") = "Kou"
'					  		Set dicDelegateReq("CreateCustomerContact") = dicCustomerContact

'					  		Dim dicCustomerLocation
'					  		Set dicCustomerLocation = CreateObject("Scripting.Dictionary")
'							dicCustomerLocation("Title") = "Mr., Mr."
'					  		dicCustomerLocation("Name") = "Kou"
'					  		dicCustomerLocation("Street") = "Sadashiv Peth"
'					  		dicCustomerLocation("City") = "Pune"
'					  		Set dicDelegateReq("CreateCustomerLocation") = dicCustomerLocation

'					  		dicDelegateReq("Purpose") = "abc desc"
'					  		msgbox Fn_SISW_SrvMgr_DelegateRequestedActivitiesType("Set", dicDelegateReq)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 10-Oct-2012		1.0								Created	
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_DelegateRequestedActivitiesType(sAction, dicDelegateReq)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_DelegateRequestedActivitiesType"
	Dim objDialog, dictItems, dictKeys, iCounter, bFlag
	Dim iRowCnt, iCnt
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("DelegateRequestedActivitiesType")

	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType", objDialog.JavaStaticText("Header_Label")) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Failed to find [ Delegate Requested Activities Type ] window.")
		Set objDialog = Nothing
		Exit function
	End If

	Select Case sAction
		Case "Set"
            'Get the keys & items count from data dictionary.	
			dictItems = dicDelegateReq.Items
			dictKeys = dicDelegateReq.Keys
			For iCounter = 0 to dicDelegateReq.Count - 1
				If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
					Select Case DictKeys(iCounter)
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "PerformsServiceRequest", "PerformRequestedActivities"
							iRowCnt = cInt(objDialog.JavaTable(DictKeys(iCounter)).GetROProperty("rows"))
							bFlag = False
							For iCnt = 0 to iRowCnt -1
								If objDialog.JavaTable(DictKeys(iCounter)).GetCellData(iCnt,0) = dictItems(iCounter) Then
									objDialog.JavaTable(DictKeys(iCounter)).ClickCell iCnt ,0 ,"LEFT"
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Successfully selected [ " & dictItems(iCounter) & " ].")
									bFlag = True
									Exit for
								End If
							Next
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Failed to select [ " & dictItems(iCounter) & " ].")
								Set objDialog = Nothing
								Exit function
							End If
'						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "RequestNumber"
'							Call Fn_List_Select("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType", objDialog, DictKeys(iCounter), dictItems(iCounter))
							objDialog.JavaList("RequestNumber").Type dictItems(iCounter)
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "RequestDate"
							' Date Buttons
							call Fn_Button_Click("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType", objDialog , DictKeys(iCounter) & "Button")
							wait 1
							If lcase(dictItems(iCounter)) = "today" Then
								Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType","Today", "")
							ElseIf lcase(dictItems(iCounter)) = "ok" Then
								Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType","OK", "")
							Else
								arrDate = split(dictItems(iCounter),"~")
								Call Fn_UI_SetDateAndTime("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType",arrDate(0), arrDate(1))
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "CustomerContact", "CustomerLocation"
							' ImageHyperlink  Clear Option
							objDialog.JavaStaticText(DictKeys(iCounter) & "DropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Clear").Select
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "sButtonName"
							' Do Nothing
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case Else
							' Edit Box 
							bReturn = Fn_Edit_Box("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType", objDialog , DictKeys(iCounter), dictItems(iCounter))
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Failed to set Editbox [ " & DictKeys(iCounter) & " = " & dictItems(iCounter) & " ].")
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					End Select
				Else
					' For ImageHyperlink Add
					Select Case DictKeys(iCounter)
						Case "CreateCustomerContact"
							objDialog.JavaStaticText("CustomerContactDropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Create...").Select
							wait 1

							If Fn_SISW_SrvMgr_CustomerInformationOperations("Create", "Create Customer Contact", dictItems(iCounter)) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Failed to [ Create Customer Contact ].")
								Exit Function
							End IF

						Case "CreateCustomerLocation"
							objDialog.JavaStaticText("CustomerLocationDropDown").Click 1, 1,"LEFT"
							wait 1
							objDialog.JavaMenu("label:=Create...").Select
							wait 1

							If Fn_SISW_SrvMgr_CustomerInformationOperations("Create", "Create Customer Location", dictItems(iCounter)) = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Failed to [ Create Customer Location ].")
								Exit Function
							End IF
					End Select
				End If
			Next

			If dicDelegateReq("sButtonName") <> "" Then
				Fn_SISW_SrvMgr_DelegateRequestedActivitiesType = Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog,dicDelegateReq("sButtonName"))
			Else
				Fn_SISW_SrvMgr_DelegateRequestedActivitiesType = True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
			'Get the keys & items count from data dictionary.	
			dictItems = dicDelegateReq.Items
			dictKeys = dicDelegateReq.Keys
			For iCounter = 0 to dicDelegateReq.Count - 1
				If TypeName(dictItems(iCounter)) = "String" OR TypeName(dictItems(iCounter)) = "Boolean" Then
					Select Case DictKeys(iCounter)
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "PerformsServiceRequest", "PerformRequestedActivities"
							iRowCnt = cInt(objDialog.JavaTable(DictKeys(iCounter)).GetROProperty("rows"))
							bFlag = False
							For iCnt = 0 to iRowCnt -1
								If objDialog.JavaTable(DictKeys(iCounter)).GetCellData(iCnt,0) = dictItems(iCounter) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Successfully verified [ " & dictItems(iCounter) & " ].")
									bFlag = True
									Exit for
								End If
							Next
							If bFlag = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Failed to verify [ " & dictItems(iCounter) & " ].")
								Set objDialog = Nothing
								Exit function
							End If
'						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "RequestNumber"
							If objDialog.JavaList("RequestNumber").GetROProperty("value") <> dictItems(iCounter) Then
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "RequestNumber_ExistsInList"
							If Fn_UI_ListItemExist("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType", objDialog , DictKeys(iCounter), dictItems(iCounter)) = False Then
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "CustomerContact", "CustomerLocation","PrimaryServiceRequest"
							' ImageHyperlink
							If objDialog.JavaObject(DictKeys(iCounter) & "ImageHyperlink").Object.getText()  <> dictItems(iCounter) Then
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "sButtonName"
							' Do Nothing
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case Else
							' Edit Box 
							If Fn_Edit_Box_GetValue("Fn_SISW_SrvMgr_DelegateRequestedActivitiesType",objDialog, DictKeys(iCounter)) <> dictItems(iCounter) Then
								Exit Function
							End If
						'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					End Select
				Else
' 					For ImageHyperlink Add
'					For future use
				End If
			Next
			If dicDelegateReq("sButtonName") <> "" Then
				Fn_SISW_SrvMgr_DelegateRequestedActivitiesType = Fn_Button_Click("Fn_SISW_SrvMgr_CreatePartMovementOperations", objDialog,dicDelegateReq("sButtonName"))
			Else
				Fn_SISW_SrvMgr_DelegateRequestedActivitiesType = True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_DelegateRequestedActivitiesType <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_DelegateRequestedActivitiesType ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function
'**************************************** Fn_SISW_SrvMgr_AssignParticipantsOperations ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_AssignParticipantsOperations
'
''Description		    :  	Function to perform operations on Assign Participants dialog in Service Request Manager

''Parameters		    :	1. sAction : Object Handle name
'					  		2. dicAssignParticipants

''Return Value		    :  	True \ False
'
''Examples		     	:	Dim dicAssignParticipants
'					  		Set dicAssignParticipants = CreateObject("Scripting.Dictionary")

'					  		dicAssignParticipants("Participants") = "Participants"
'					  		dicAssignParticipants("Organization") = "Organization:Engineering:Designer:Amol Lanke (x_lanke)"
'					  		dicAssignParticipants("AnyMember") = True
'					  		dicAssignParticipants("AllMembers") = False
'					  		dicAssignParticipants("SpecificGroup") = False
'					  		dicAssignParticipants("AnyGroup") = True
'					  		dicAssignParticipants("ButtonName") = "OK"
'					  		msgbox Fn_SISW_SrvMgr_AssignParticipantsOperations("Assign", dicAssignParticipants)

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		 10-Oct-2012		1.0								Created	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_AssignParticipantsOperations(sAction, dicAssignParticipants)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_AssignParticipantsOperations"
	Dim objDialog, sPath, arrPath, iCnt, iArrCnt, iPath
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("Search").JavaDialog("AssignParticipants")
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_AssignParticipantsOperations", objDialog) = false Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_AssignParticipantsOperations ] Failed to find [ Assign Participants ] window.")
		Set objDialog = Nothing
		Exit Function
	End If
	Select Case sAction
		Case "Assign"
			If dicAssignParticipants("Participants") <> "" Then
				arrPath = split(dicAssignParticipants("Participants"),":")
				For iCnt = 0 to UBound(arrPath) - 1
					If iCnt > 0 Then
						sPath = ""
						For iArrCnt = 0 TO iCnt
							If iArrCnt = 0 Then
								sPath = arrPath(iArrCnt)
							Else
								sPath = sPath & ":" & arrPath(iArrCnt)
							End If
						Next
							objDialog.JavaTree("ParticipantsTree").Expand sPath
					End If
				Next

					objDialog.JavaTree("ParticipantsTree").Select dicAssignParticipants("Participants")
			End If

			objDialog.JavaTab("TabbedPane").Select "Organization"

			If dicAssignParticipants("Organization") <> ""  Then
				arrPath = split(dicAssignParticipants("Organization"),":")
				For iCnt = 0 to UBound(arrPath) - 1
					If iCnt > 0 Then
						sPath = ""
						For iArrCnt = 0 TO iCnt
							If iArrCnt = 0 Then
								sPath = arrPath(iArrCnt)
							Else
								sPath = sPath & ":" & arrPath(iArrCnt)
							End If
						Next
							objDialog.JavaTree("OrganizationTree").Expand sPath
							wait 2
					End If
				Next

					objDialog.JavaTree("OrganizationTree").Select  dicAssignParticipants("Organization")
			End If

			' Radio Buttons
			If dicAssignParticipants("AnyMember") <> ""  Then
				If cBool(dicAssignParticipants("AnyMember") ) Then
					objDialog.JavaRadioButton("AnyMember").Set "ON"
				Else
					objDialog.JavaRadioButton("AnyMember").Set "OFF"
				End If
			End If

			If dicAssignParticipants("AllMembers") <> ""  Then
				If cBool(dicAssignParticipants("AllMembers") ) Then
					objDialog.JavaRadioButton("AllMembers").Set "ON"
				Else
					objDialog.JavaRadioButton("AnyMember").Set "OFF"
				End If
			End If

			If dicAssignParticipants("SpecificGroup") <> ""  Then
				If cBool(dicAssignParticipants("SpecificGroup") ) Then
					objDialog.JavaRadioButton("SpecificGroup").Set "ON"
				Else
					objDialog.JavaRadioButton("AnyMember").Set "OFF"
				End If
			End If

			If dicAssignParticipants("AnyGroup") <> ""  Then
				If cBool(dicAssignParticipants("AnyGroup") ) Then
					objDialog.JavaRadioButton("AnyGroup").Set "ON"
				Else
					objDialog.JavaRadioButton("AnyGroup").Set "OFF"
				End If
			End If
			Call Fn_Button_Click("Fn_SISW_SrvMgr_AssignParticipantsOperations", objDialog, "Add")
			If dicAssignParticipants("ButtonName") <> ""  Then
				Call Fn_Button_Click("Fn_SISW_SrvMgr_AssignParticipantsOperations", objDialog, dicAssignParticipants("ButtonName"))
			End If
			Fn_SISW_SrvMgr_AssignParticipantsOperations = True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_AssignParticipantsOperations ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_AssignParticipantsOperations <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_AssignParticipantsOperations ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function

''****************************************    Function to set Date for Revision Rule ***************************************
'Function Name		      :			Fn_SISW_SrvMgr_RevisionRuleSetDate
'
'Description			     :  	     Function to set Date for Revision Rule 
'
''Parameters			   :	   	 	1. sAction : Action need to perform
''								  				  2. sDate, 
'								  				 3. sCurrent
'												4. sShowOnlyOperational
											
'Return Value		       : 			True \ False
'
'Pre-requisite			    :		 	 BOM Line should be selected
'
'Examples				    :			  
'								Call Fn_SISW_SrvMgr_RevisionRuleSetDate("Set", "29-Oct-2010 11:17:49", False, "", "OK")
'								Call Fn_SISW_SrvMgr_RevisionRuleSetDate("Set", "", True, "", "OK")
'
''History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle	 		29-Oct-2010		     	    1.0										    		Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_RevisionRuleSetDate(sAction, sDate, sCurrent, sShowOnlyOperational, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_RevisionRuleSetDate"
	Dim objRevRuleSetDate, arrDateTime
	
	Set objRevRuleSetDate = JavaWindow("ServiceManager").JavaWindow("SetDate")
	Fn_SISW_SrvMgr_RevisionRuleSetDate = True
    		
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_RevisionRuleSetDate", objRevRuleSetDate) = False Then
	   	Call Fn_MenuOperation("Select", "Tools:Revision Rule:Set Date..." )
	   If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_RevisionRuleSetDate", objRevRuleSetDate) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_RevisionRuleSetDate ] Failed to open [ Set Date ] window.")
			Fn_SISW_SrvMgr_RevisionRuleSetDate = False
			Set objRevRuleSetDate = nothing
	   End If
	End If
	Select Case sAction
		Case "Set"
            'set checkbox
			If sCurrent <> "" Then
				If cBool(sCurrent) Then
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_RevisionRuleSetDate", objRevRuleSetDate, "Current","ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_RevisionRuleSetDate", objRevRuleSetDate, "Current","OFF")
					arrDateTime = Split(sDate," ")
					objRevRuleSetDate.JavaEdit("Date").Set arrDateTime(0)		'' date control object changed. modified as per new design change BY Rajendra P on(17-Jau-2014), 
					wait 1
					call Fn_KeyBoardOperation("SendKeys", "{TAB}")
					objRevRuleSetDate.JavaList("SetDate").Type arrDateTime(1)
				End If
			End If

			 'set checkbox
			If sShowOnlyOperational <> "" Then
				If cBool(sShowOnlyOperational) Then
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_RevisionRuleSetDate", objRevRuleSetDate, "Current","ON")
				Else
					Call Fn_CheckBox_Set("Fn_SISW_SrvMgr_RevisionRuleSetDate", objRevRuleSetDate, "Current","OFF")
				End If
			End If

			If sButton = "" Then sButton = "OK"
            'clicking on OK
			Call Fn_Button_Click("Fn_SISW_SrvMgr_RevisionRuleSetDate", objRevRuleSetDate, sButton)
			Fn_SISW_SrvMgr_RevisionRuleSetDate = True
			
		Case Else 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_RevisionRuleSetDate ] Invalid Action [ " & sAction & " ].")
			Fn_SISW_SrvMgr_RevisionRuleSetDate = False
			Set objRevRuleSetDate = nothing
	End Select
	Set objRevRuleSetDate = nothing
End Function
''*******************************************************************************
'Function Name		:	Fn_SISW_SrvMgr_SetupUpgrade
'
'Description		:  	Function to perform operations on Setup Upgrade dialog 
'
''Parameters		:	1. sAction : Action need to perform
''						2. bReconfigureNeutralItem, 
'						3. bSearchNewNeutralItem
'						4. bSearchSavedStructure
'						5. dicSearch : to Search And Configure BOM Line
'						6. dicSetupUpgradeProperties : For future use.
											
'Return Value		: 	True \ False
'
'Pre-requisite		:	BOM Line should be selected
'
'Examples			: 
'						Dim dicSearch
'					  	Set dicSearch = CreateObject("Scripting.Dictionary")
'					  	dicSearch("Name") = "par*"
'					  	dicSearch("SearchResults_Select") = "abc"
'					  	dicSearch("Configure_BOMLine") = "abc"
'					  	Call Fn_SISW_SrvMgr_SetupUpgrade("SetupUpgrade", true, true, "", dicSearch, "")
'
''History:
'	Developer Name			Date		Rev. No.			Reviewer			Changes Done			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe	 	04-Dec-2012		1.0					Self			Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_SetupUpgrade(sAction, bReconfigureNeutralItem, bSearchNewNeutralItem, bSearchSavedStructure, dicSearch, dicSetupUpgradeProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_SetupUpgrade"
	Dim objDialog
	Set objDialog = JavaWindow("ServiceManager").JavaWindow("SetupUpgrade")
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_SetupUpgrade", objDialog) = False Then
		Exit function
	End If
	Select Case sAction
		Case "SetupUpgrade"
			If bReconfigureNeutralItem <> "" Then
				If cBool(bReconfigureNeutralItem) Then
					objDialog.JavaRadioButton("ReconfigureNeutralItem").Set "ON"
				End If
			End If

			If cBool(objDialog.JavaObject("Twistie").Object.isExpanded()) = False Then
				objDialog.JavaObject("Twistie").Click 1, 1,"LEFT"
			End if

			If bSearchNewNeutralItem <> "" Then
				If cBool(bSearchNewNeutralItem) Then
					objDialog.JavaRadioButton("SearchNewNeutralItem").Set "ON"
				End If
			End If

			If bSearchSavedStructure <> "" Then
				If cBool(bSearchSavedStructure) Then
					objDialog.JavaRadioButton("SearchSavedStructure").Set "ON"
				End If
			End If

			Call Fn_Button_Click("Fn_SISW_SrvMgr_SetupUpgrade", objDialog, "OK")

             Fn_SISW_SrvMgr_SetupUpgrade = Fn_SISW_SrvMgr_SearchOperations("SearchAndConfigure", dicSearch)
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_SISW_SrvMgr_SetupUpgrade ] Invalid case [ " & sAction & " ].")
	End Select

	If Fn_SISW_SrvMgr_SetupUpgrade <> False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_SISW_SrvMgr_SetupUpgrade ] Successfully executed with case [ " & sAction & " ].")
	End If
	Set objDialog = Nothing
End Function


'*********************************************************		Function to Create Company Location ***********************************************************************

'Function Name		:					Fn_SISW_SrvMgr_CreateCompanyLocation

'Description			 :		 		  This function is used to Create Company Location On Any node.

'Parameters			   :	 			1.  sAction
'													2. sCompanyLocType :Company Location Type
'													3. sButtonName: Name of Node of BOMTable
'													4. dicInputs: Dictionary For Inputs

'Return Value		   : 				 True Or False

'Pre-requisite			:		 		Create Company Location Window should be visible

'Examples				:               Set dic1= CreateObject( "Scripting.Dictionary" )
'													dic1("Name") = "xyzesfd"
'													dic1("Location Code") = "xyzesfd"
'													dic1("Location Type") = "xyzesfd"
'													dic1("Street") = "xyzesfd"
'													dic1("City") = "xyzesfd"
'													dic1("State / Province") = "xyzesfd"
'													dic1("Postal Code") = "xyzesfd"
'													dic1("Country") = "xyzesfd"
'													dic1("Region") = "xyzesfd"
'													dic1("URL") = "xyzesfd"
'													dic1("Description") = "xyzesfd"
'																				
'													
'													Call Fn_SISW_SrvMgr_CreateCompanyLocation("Create" , "Company Location", "Finish" , dic1)

'History:
'	Developer Name					Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle					29-Nov-2013			1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_CreateCompanyLocation(sAction , sCompanyLocType, sButtonName , dicInputs)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_CreateCompanyLocation"
	Dim dicItem,dicValue
	Dim objNewBusinessObject , intg, aType, sExpand
	Dim iCount, iItemCount, crrItem, bFlag

   Set objNewBusinessObject = JavaWindow("ServiceManager").JavaWindow("NewBusinessObject")
   Fn_SISW_SrvMgr_CreateCompanyLocation = False
	     
   'If dialog does not exist, invoke Menu - [ File->New->Physical Location ]
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateCompanyLocation",objNewBusinessObject) = False Then
		If Instr(1, sCompanyLocType,"Company Location")> 0  Then
			Call Fn_MenuOperation("Select", "File:New:Company Location...")
		Else
			Call Fn_MenuOperation("Select", "File:New:Company Contact...")
		End If
		Call Fn_ReadyStatusSync(2)
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SISW_SrvMgr_CreateCompanyLocation ] successfully invoked menu - [ File:New:Company Location... ]")
	End If
	
	'Check if it is open now
	If Fn_UI_ObjectExist("Fn_SISW_SrvMgr_CreateCompanyLocation",objNewBusinessObject) = False Then
	   'Exit function
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Function - [ Fn_SISW_SrvMgr_CreateCompanyLocation ] Failed to Open [ NewBusinessObject ] dialog for new Company Location")
		Set objEditQuan = nothing
		Exit Function
	End If

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	If objNewBusinessObject.JavaTree("PhysicalLocationType").Exist(3) Then
			iItemCount=Fn_UI_Object_GetROProperty("Fn_SISW_SrvMgr_CreateCompanyLocation",objNewBusinessObject.JavaTree("PhysicalLocationType"), "items count")
			For iCount=0 To iItemCount-1
				crrItem=objNewBusinessObject.JavaTree("PhysicalLocationType").GetItem(iCount)
				If Trim(crrItem)="Most Recently Used:"+Trim(StrItemType) Then
					bFlag=True
					Exit For
				ElseIf Trim(crrItem)="Complete List" Then
					Exit For
				End If
			Next
		
			If bFlag=True Then
				Call Fn_JavaTree_Select("Fn_SISW_SrvMgr_CreateCompanyLocation", objNewBusinessObject, "PhysicalLocationType","Most Recently Used:"+sCompanyLocType)
			Else
				Call Fn_UI_JavaTree_Expand("Fn_SISW_SrvMgr_CreateCompanyLocation", objNewBusinessObject, "PhysicalLocationType","Complete List")
				Call Fn_JavaTree_Select("Fn_SISW_SrvMgr_CreateCompanyLocation", objNewBusinessObject, "PhysicalLocationType","Complete List:"+sCompanyLocType)	
			End If
			wait 2
		
			'Click on NEXT button to navigate ahead
			Call Fn_Button_Click("Fn_SISW_SrvMgr_CreateCompanyLocation", objNewBusinessObject, "Next")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Function - [ Fn_SrvMgr_NewPhysicalLocation ] successfully Clicked on [ Next ] Button") 
			Call Fn_ReadyStatusSync(2)
	End If

	Select Case sAction
		Case "Create"
			    dicItem = dicInputs.Keys
				dicValue = dicInputs.Items
				For iCount = 0 to dicInputs.Count - 1
					Select Case trim(dicItem(iCount))
						Case "Location Type"
							'JavaList
'							objNewBusinessObject.JavaList(trim(dicItem(iCount))).Click 1,1,"LEFT"
'							objNewBusinessObject.JavaList(trim(dicItem(iCount))).Type trim(dicValue(iCount))
'							If Err.number < 0 Then
'								Exit Function
'							End If
						'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
						Case "Name", "Location Code", "Street", "City", "State / Province", "Postal Code", "Country", "Region", "URL", "Description",_
							"First Name","Last Name","Suffix","Phone (Business)","Phone (Home)","Phone (Mobile)","Fax","Pager","Email"
							'Edit Box
							objNewBusinessObject.JavaStaticText("PropertyLabel").SetTOProperty "label", trim(dicItem(iCount))&":"

							bResult = Fn_SISW_UI_JavaEdit_Operations("Fn_SISW_SrvMgr_CreateCompanyLocation", "Set",  objNewBusinessObject, "PropertyValue", trim(dicValue(iCount)) )
							If bResult = False Then Exit Function
					   End Select
				Next
		Case Else
			Exit Function
	End Select


	If sButtonName <> "" Then
		Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvMgr_CreateCompanyLocation", "Click", objNewBusinessObject,sButtonName)
		If  sButtonName  = "Finish" Then
			Call Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SrvMgr_CreateCompanyLocation", "Click", objNewBusinessObject,"Cancel")
		End If
	End If
	Fn_SISW_SrvMgr_CreateCompanyLocation=True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_SrvMgr_CreateCompanyLocation:Successfully created Company Location.")
	Set objNewBusinessObject = Nothing
End Function


'****************************************    Function to Select A Lot ***************************************
'
''Function Name		 	:	Fn_SISW_SrvMgr_SelectALot
'
''Description		    :  	Function to perform  assign lot operation

''Parameters		    :	1. sAction : Action need to perform
			'					   		2. sLotName
			'					   		3. sMessage
			'							4. sButtonName
								
''Return Value		    :  	True \ False
'
''Pre-requisite		    :	Select A Lot  Window should be present.

''Examples		     	:	

'Examples		     	:	Case "Select"
'										Fn_SISW_SrvMgr_SelectALot("Select", "000227/Lot4", "", "OK")
'History:
'	Developer Name				Date					Rev. No.			Reviewer			Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle		 			29-Jan-2014				1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_SrvMgr_SelectALot(sAction, sLotName, sMessage, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_SrvMgr_SelectALot"
	Dim objSelectALot, bResult,  sErrMsg

	Set objSelectALot = JavaWindow("ServiceManager").JavaWindow("CreatePartMovement").JavaWindow("SelectALot")

	Fn_SISW_SrvMgr_SelectALot = False

	If Fn_SISW_UI_Object_Operations("Fn_SISW_SrvMgr_SelectALot","Exist", objSelectALot,"") = False Then
			Exit Function
	End If

	Select Case sAction
		Case "CloseDialog"
				Call Fn_Button_Click("Fn_SISW_SrvMgr_SelectALot", objSelectALot, "Close")
				Fn_SISW_SrvMgr_SelectALot = True

		Case "Select", "Verify"

					If sMessage <> "" Then
							sErrMsg = objSelectALot.JavaStaticText("ErrMsg").GetROProperty("label")
							If sMessage <> sErrMsg Then
								Exit Function
							End If
					End If
					
					If sLotName <> "" Then
						bResult = Fn_SISW_UI_JavaTable_Operations("Fn_SISW_SrvMgr_SelectALot", "GetRowIndex", objSelectALot, "AvailableLots", "GetCellData", "Object", sLotName, "", "", "", ":")
						 ' Check Condition if path returns false 
						If bResult = -1 Then Exit Function
						objSelectALot.JavaTable("AvailableLots").SelectCell bResult,"Object"
					End If

					If sButtonName <> "" Then
							Call Fn_Button_Click("Fn_SISW_SrvMgr_SelectALot", objSelectALot, sButtonName)
					End If
			
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: [ Fn_SISW_SrvMgr_SelectALot ] invalied case [ " & sAction & " ].")
			Exit Function
	End Select

	Set objSelectALot = Nothing
	Fn_SISW_SrvMgr_SelectALot = True
End Function
