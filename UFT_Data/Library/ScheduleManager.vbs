Option Explicit

Dim sTskAssignId
Dim sTskAssignRev
Dim sErrorText

'----------------------'Global variables for Teamcenter Perspective Names----------------------------------------------------------------
Public GBL_PERSPECTIVE_SCHEDULE_MANAGER
GBL_PERSPECTIVE_SCHEDULE_MANAGER = "Schedule Manager"
'----------------------'Global variables for Teamcenter Perspective Names----------------------------------------------------------------
'*********************************************************	Function List		***********************************************************************
'0. Fn_SISW_PPM_GetObject(sObjectName)
'1. Fn_SchMgr_TreeTableRowIndex()
'2. Fn_SchMgr_TableColIndex()
'3. Fn_SchMgr_SchTable_NodeOperation()
'4. Fn_SchMgr_TaskCreate()
'5. Fn_SchMgr_MilestoneCreate()
'6. Fn_SchMgr_TaskDelete()
'7. Fn_SchMgr_TaskConstraint()
'8. Fn_SchMgr_ScheduleRecalculate()
'9. Fn_SchMgr_NodeRename()
'10. Fn_SchMgr_TableRowIndex()
'11. Fn_SchMgr_ScheduleMembership()
'12. Fn_SchMgr_PercentLinkedMsgVerify()
'13. Fn_SchMgr_ScheduleShift()
'14. Fn_SchMgr_iComboSet() - Moved to GeneralFunctions.vbs with name Fn_iComboSet()
'15. Fn_SchMgr_SchPropertyOperations()
'16. Fn_SchMgr_TaskPropertyOperations()
'17. Fn_SchMgr_TaskDetailCreate()
'18. Fn_SchMgr_SchedulingErrorVerify()
'19. Fn_SchMgr_TaskAssignment()
'20. Fn_SchMgr_TskIndentOutdent()
'21. Fn_SchMgr_CutNode()
'22. Fn_SchMgr_CopyNode()
'23. Fn_SchMgr_PasteNode()
'24. Fn_SchMgr_ColumnOperations()
'25. Fn_SchMgr_ReplaceTskAssignment()
'26. Fn_SchMgr_TskDesignateDisciplines()
'27. Fn_SchMgr_BaselineOperations()
'28. Fn_SchMgr_ManageBaselines()
'29. Fn_SchMgr_WarningMsgVerify()
'30. Fn_SchMgr_WinSchErrorVerify()	  - Eliminated. Not used anywhere . By Sushma Pagare [13-Jun-13]
'31. Fn_SchMgr_TaskDeliverable()
'32. Fn_SchMgr_ViewBaseline()
'33. Fn_SchMgr_SchDeliverableCreate()
'34. Fn_SchMgr_RevertAssignDiscpline()
'35. Fn_SchMgr_TaskDependency()
'36. Fn_SchMgr_BaselineTask()
'37. Fn_SchMgr_DialogMsgVerify()
'38. Fn_SchMgr_ApplyConstraintConfirm()
'39. Fn_SchMgr_InsertTemplates()
'40. Fn_SchMgr_InsertSchedule()
'41. Fn_SchMgr_RateModifiers()
'42. Fn_SchMgr_ChooseSchedules()
'43. Fn_SchMgr_FixedCostsAction()
'44. Fn_SchMgr_DateChooser()
'45. Fn_SchMgr_HoursChooser()
'46. Fn_SchMgr_FilterSettings()
'47. Fn_SchMgr_SummarySchPropertyOperations()
'48. Fn_SchMgr_FormatDate()
'49. Fn_SchMgr_ScheduleCalendarOperations()
'50. Fn_SchMgr_CostOperation()
'51. Fn_SchMgr_Load_Schedule()
'52. Fn_SchMgr_ProgViewColumnOperations
'53  Fn_SchMgr_WinListViewIndex
'54  Fn_SchMgr_ColumnChooserOperations()
'55  Fn_SchMgr_ChooseSchedulesOperations()
'56  Fn_SchMgr_SaveProgram()
'57. Fn_SchMgr_GanttChartOperations()
'58  Fn_SchMgr_CostBillValues_Operations()
'59  Fn_SchMgr_DefineWBSFormat()
'60. Fn_WBS_OptionsSettings()
'61. Fn_SchMgr_CriticalPathColourOperations()
'62  Fn_SchMgr_ConfirmLaunchWorkflow()
'63. Fn_SchMgr_ErrorDialogMessageVerify()
'64. Fn_SchMgr_InformationDialogHandle()
'65. Fn_SchMgr_PrintTablePropertyVerify()
'66 .Fn_SchMgr_ProgramViewCreate()
'67  Fn_SchMgr_CreateCrossScheduleDependency()
'68  Fn_SchMgr_CreateProxyTask()
'69  Fn_SchMgr_CustomizeGroupFilters()
'70  Fn_SchMgr_GetDay()
'71  Fn_SchMgr_ProgramViewSaveAs()
'72  Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day()
'73  Fn_SchMgr_GetDateByDay()
'74  Fn_SchMgr_DeleteAllExistingRates()
'75 Fn_SchMgr_WorkflowPrivilegedUser()
'76 Fn_SISW_SchMgr_ErrorVerify()
'77.Fn_ScMgr_DeliverableOpenByNameVerify()
'78. Fn_SchMgr_VerityTaskTimeAndDuration()
'*********************************************************	Function List		*********************************************************************************************************

'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_PPM_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_PPM_GetObject("Baseline Schedule")
'										Fn_SISW_PPM_GetObject("SchTaskTable")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sushma Pagare 		 18-June-2012		1.0				
'   Shreyas Waichal		  20 -June-2012		1.1
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_PPM_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ScheduleManager.xml"
	Set Fn_SISW_PPM_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'*********************************************************		Function to  get  Schedule table Row Index	***********************************************************************

'Function Name		:					Fn_SchMgr_TreeTableRowIndex

'Description			 :		 		  This function is used to get Schedule table Row Index.

'Parameters			   :	 			1.  sNodeName:Name of the Node to retrieve Index for.
											
'Return Value		   : 				 Node index

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:				 Fn_SchMgr_TreeTableRowIndex("Sch1:Task1:T2")

'History:
'	Developer Name			Date			Rev. No.	Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Rupali					19-May-2010		1.0			
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_TreeTableRowIndex(objTreeTable, sNodeName, sColname)
	
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TreeTableRowIndex"
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
							'Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TreeTableRowIndex: Row Index of [" + sNodeName +"] Node is [" + IntCounter + "]")	
							'objTreeTable.DeselectRow IntCounter
							Exit For
						Else
							'objTreeTable.DeselectRow IntCounter
						End If
					End If
				Next
	   Next

		Fn_SchMgr_TreeTableRowIndex = StrIndex
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TreeTableRowIndex: Row Index of [" + sNodeName +"] Node is [" + StrIndex + "]")	

		If  cint(IntCounter) = cint(IntRows) Then
			Fn_SchMgr_TreeTableRowIndex = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TreeTableRowIndex:Failed to Get  Row Index of [" + sNodeName +"]")	
		End If

  End If
End Function


'*********************************************************		Function to Get  Schedule Table Column Index 		***********************************************************************

'Function Name		:					Fn_SchMgr_TableColIndex

'Description			 :		 		  This function is used to get the schdule Table column Index.

'Parameters			   :	 			1.  StrColName:Name of the Col to retrieve Index for.
											
'Return Value		   : 				 Col index/False

'Pre-requisite			:		 		Stchdule Manager window should be displayed .

'Examples				:				Fn_SchMgr_TableColIndex("Task Duration")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							 21-May-2010   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_TableColIndex(ObjTable, sColName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TableColIndex"
	Dim iCols , iCounter, sColIndex, sName

	On Error Resume Next

	'Verify that  scheduleTable is displayed
	If ObjTable.Exist(5) Then

		'Get the No. of cols present in the schedule Table

		iCols = ObjTable.GetROProperty("cols")
		
		'Get the Col No. of required Column
		For iCounter = 0 to iCols -1
			sName =ObjTable.Object.getColumnName(iCounter)
		  
			If Trim(sName) = Trim(sColName) Then
				Fn_SchMgr_TableColIndex = iCounter
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TableColIndex:The Column Index for Column [" + sColName +"] is [" + iCounter + "]")	
				Exit For
			End If
		Next
		If Cint(iCounter) = Cint(iCols) Then
			Fn_SchMgr_TableColIndex = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TableColIndex:The Column [" + sColName + "] dose not exist in schedule  table")	
		End If

	End If
End Function

'******************************************************************Function to perform schedule table  operation************************************************************************************************************

'	Function Name		:	Fn_SchMgr_SchTable_NodeOperation
'
'	Description			:	Actions performed in this function are:
'							1.  Node Select
'							2.  Node multi-select
'							3.  Node Expand
'							4.  Node Collapse
'							5.  Node Popup menu select
'							6.  Cell Verify
'							7.  Cell edit   ( Pass column index.)
'							8.  Cell double-click
'							9.  Exists
'							10. MultiSelectPopup
'							11. Verify Cell Value When Node name is not given
'
'	Parameters			:	1. sAction: Action to be performed
'							2. sObject: Name of the object ( , is use as seprator)
'							3. sColName: Column name or index ( For cell edit case pass column index only)
'							4. sValue: Value to be set
'							5. sMenu: Context menu to be selected
'							
'	Return Value		:	TRUE \ FALSE
'
'	Pre-requisite		:	Schedule manger panel need to open
'
'	Examples			:	Call Fn_SchMgr_SchTable_NodeOperation ("PopupMenu", "sch:T3", "", "" , "New:Task..." )
'							Call Fn_SchMgr_SchTable_NodeOperation("CellVerify" , "Qwerty-Test:t3" , "Status Indicator" , "In_Progress" , "")
'							Call Fn_SchMgr_SchTable_NodeOperation("VerifyCellValueWithoutNodeName", "CDM_Quarterly_Template57230_001678_A:CDM_Quarterly_Template57230_001678_A" , "Start Date" , "08-Jan-2020 08:00~08-Jul-2020 08:00" , "")
'							Call Fn_SchMgr_SchTable_NodeOperation("VerifyBackGroundColor" ,"sch:T3" , "" , "RED" , "")
'	History:
'							Developer Name			  Date			Rev. No.				Changes Done												Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'								Rupali				18/05/10		  1.0  									
'								Shreyas				28/11/11		  1.1				Modified Case "Cell Verify"										Prasanna
'								Sushma				06/12/11		  1.2				Modified Case "CellEdit"										Vallari
'								Ashok kakade		02/01/13		  1.2				Modified Case "Select"			
''								Ashok kakade		02/01/13		  1.2			Added Case "CellVerifyWIthSameInstance"
'								Ankit Nigam			23/09/2016        1.3			Added Case "VerifyCellValueWithoutNodeName"				[TC1123-20160831b00-23_09_2019-VivekA-NewDevelopment]
'								Poonam Chopade		22/02/2017        1.4			Added Case "VerifyBackGroundColor"						[TC1123-20161205c00-23_02_2017-PoonamC-NewDevelopment]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function  Fn_SchMgr_SchTable_NodeOperation(sAction , sObject , sColName , sValue , sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_SchTable_NodeOperation"
	On Error Resume Next
	Fn_SchMgr_SchTable_NodeOperation = FALSE

	Dim objSchTable,sIndex,iCounter,ArrNodes, IntArrCounter,sCount,sContext,aMenuList,bReturn,iInstance,aNodePath,objTree,row,iIndex, aColArr
	Dim objSelectType,obj,WshShell,defaultColWidth
	Dim aDate, sActDate, aActDate, iMonDiff, iYrDiff, iDiff, iCnt, objShiftWin,iCnter,sPath
	Dim sNodecolor,arrcolor,rColor,gColor,bColor
	
	iInstance = 1
	Call Fn_LoadSchedule("No", "")
	'Create object of  table
	Set objSchTable = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable")

	If objSchTable.Exist(5) Then

		'Verify if the column is already present in the Table .
		' Reason for addition of this code segment is to handle the application change in the TC8_3_0_2 which will port to the Tc9.0
		'1. Renamed the Column "Workflow Task Template" to "Workflow Template" in Tc 8_3_0_2
		'2. Try to Add the Column if not present in the application. - Code Added by Archana (11-Nov10)

		' ******START
			If sColName = "Workflow Task Template" then 
				sColName = "Workflow Template"
			ElseIf sColName = "Duration" Then
				sColName = "Task Duration"
			ElseIf sColName = "Work Estimate" Then
				sColName = "Task Work Estimate"			
			ElseIf sColName = "ResourceAssignment" Then   ' Added to handle name change for Resource Assignment column
				sColName = "Resource Assignment"	
			End if
			
			If sColName <> "" Then
				ReDim aColArr(0)
				aColArr(0) = sColName
				bReturn = Fn_SchMgr_ColumnOperations("Add",aColArr )
			End If
		'******End

'	 If instr(sObject, "@") > 0 And Trim(Lcase(sAction)) <> "multi-select" Then
'		aNodePath = split(sObject, "@",-1, 1)
'		sObject = aNodePath(0)
'		iInstance = cint(aNodePath(1))
'	End If
	Call Fn_ReadyStatusSync(1)
	JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Click 0,0
	wait 2
	Select Case sAction
'.-------------------------------------------------------------------------------------	----------------------------------------------	----------------------------------------------	
		' Added New Case to Verify Value of Cell without Node Name - [TC1123-20160831b00-23_09_2019-AnkitN-NewDevelopment]
		Case "VerifyCellValueWithoutNodeName"
			If CInt(CInt(objSchTable.GetROProperty("rows")) - CInt(UBound(Split(sObject, ":")) + 1)) > CInt(UBound(Split(sValue, "~")) + 1) Then
				Set objSchTable = Nothing
				Fn_SchMgr_SchTable_NodeOperation = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation : Extra rows are coming in Table.")
				Exit Function
			Else
				For iCnt = 0 To CInt(objSchTable.GetROProperty("rows")) - 1
					objSchTable.SelectRow iCnt
					wait 1
					If Instr(1, sObject, objSchTable.GetCellData(iCnt, "Object")) = 0 Then
						If Cstr(objSchTable.GetCellData(iCnt, sColName)) <> Cstr(Split(sValue, "~")(iCnt - CInt(UBound(Split(sObject, ":")) + 1))) Then
							Set objSchTable = Nothing
							Fn_SchMgr_SchTable_NodeOperation = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation : Cell Value is not verified.")
							Exit Function
						End If
					End If
				Next
			End If

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
					 Fn_SchMgr_SchTable_NodeOperation = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not selected.")	
						Exit Function
				Else
						Fn_SchMgr_SchTable_NodeOperation = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation: Row with Object "&sObject&" is selected.")	
				End If
			Else
				sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
				If Instr(sIndex, "#") > 0 Then
					'Calculate the instance number
'					sIndex = cint(sIndex) + iInstance - 1
					'sIndex = "#" + cstr(sIndex)
					'Select the Expected  scheduleTable Node
					 objSchTable.SelectRow sIndex

				     If Err.Number <  0 Then
						 Fn_SchMgr_SchTable_NodeOperation = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not selected.")	
						Exit Function
					Else
						Fn_SchMgr_SchTable_NodeOperation = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation: Row with Object "&sObject&" is selected.")	
					End If
				Else
					Fn_SchMgr_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists in schedule table..")	
					Exit Function
				End If
			End If
		'.---------------------------------------This case is used to multi-select the ScheduleTable Nodes.----------------------------------------------
        Case  "MultiSelect"

			 Dim ArrNode(),IntRows
			 ArrNodes = split(sObject, ",",-1,1)
			 objSchTable.Object.clearSelection
			 ReDim ArrNode(Ubound(ArrNodes))
			 For IntArrCounter = 0 to Ubound(ArrNodes)
'				   'Verify node multiple occurrence
'					iInstance = 1
'					If instr(ArrNodes(IntArrCounter), "@") > 0 Then
'						aNodePath = split(ArrNodes(IntArrCounter), "@",-1, 1)
'						sObject = aNodePath(0)
'						iInstance = cint(aNodePath(1))
'					Else
'						sObject = ArrNodes(IntArrCounter)
'					End If
'					 sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
					sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, ArrNodes(IntArrCounter), "Object")

					 If Instr(sIndex, "#") > 0 Then
						'Calculate the instance number
'						sIndex = cint(sIndex) + iInstance - 1
						'sIndex = "#" + cstr(sIndex)
						ArrNode(IntArrCounter) = sIndex
					Else
						Fn_SchMgr_SchTable_NodeOperation = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists in schedule table..")	
						Exit Function
					End If
			 Next

			 'objSchTable.PressKey micCtrl
'			 Const VK_CONTROL = 29
'			 Set WshShell  = CreateObject("Mercury.DeviceReplay")
'			 WshShell.KeyDown VK_CONTROL
			  objSchTable.Object.clearSelection
			  wait(1)
			     
			  'added Code to get the default width of the column object on 14/4/2015 as multi select not working for lower nodes
			'-------------------------------------------------------------------------------------------------------------			  
			  defaultColWidth=objSchTable.object.getColumnModel().getColumn(0).getPreferredWidth()
			  wait(1)
			  
			  'Code added to expand object column to maximum size
			  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaObject("ColumnHeaders").DblClick "0","0"
			  wait (1)
			  
			  objSchTable.SelectRow ArrNode(0)
			  wait(1)
			  For iCounter = 1 to Ubound(ArrNode)
				  objSchTable.ExtendRow ArrNode(iCounter)
				

				  ' Code Added By PRiyanka [11-11-10]
				  ' Adding Hard Coded Wait for Synchronization as Ready Status is not changing in Multi Select Operation
				  ' Reviewd By Archana
				  Wait(3)
				  If Err.Number <  0 Then
							 Fn_SchMgr_SchTable_NodeOperation = FALSE				 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Row with Object " + ArrNode(iCounter) + " does not selected for mutiselect case.")	
							Exit Function
                   Else
							Fn_SchMgr_SchTable_NodeOperation = TRUE				 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation: Row with Object " + ArrNode(iCounter) + " is selected for multiselect case..")	
				End If
			  Next
			'Modified code to reset column width to default
			objSchTable.object.getColumnModel().getColumn(0).setPreferredWidth(defaultColWidth)
	 		wait (1)	
	 '.---------------------------------------This case is used to expand all the Shedule Table Node.----------------------------------------------
		Case	 "Expand" 
            bReturn = Fn_MenuOperation("Select", "View:Expand All")
			If bReturn = TRUE Then
				Fn_SchMgr_SchTable_NodeOperation = TRUE				 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Expand all the node.")
			Else
			   Fn_SchMgr_SchTable_NodeOperation = FALSE				 
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation : Fail to expand the node..")	
			  Exit Function
			End If

		'.---------------------------------------This case is used to collapse all the Schedule Table Node.----------------------------------------------
		Case 	 "Collapse"
			 bReturn =  Fn_MenuOperation("Select", "View:Collapse All")
			If bReturn = TRUE Then
				Fn_SchMgr_SchTable_NodeOperation = TRUE				 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Collapse all the node.")
			Else
			   Fn_SchMgr_SchTable_NodeOperation = FALSE				 
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Fail to collapse the node..")	
			  Exit Function
			End If

		'.---------------------------------------This case is used to  RMB  click  and  select  pop menu  of   Schedule Table Node cell.----------------------------------------------
		Case	 "PopupMenu"
			sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
 
			   If Instr(sIndex, "#") > 0 Then
				   'Calculate the instance number
'					sIndex = cint(sIndex) + iInstance - 1
					'sIndex = "#" + cstr(sIndex)
		
					objSchTable.SelectRow  sIndex
					objSchTable.ClickCell sIndex,"Object","RIGHT","NONE"
					wait 0,200
						aMenuList = split(sMenu, ":",-1,1)			
						sCount = Ubound(aMenuList)
								Select Case sCount
										Case "0"
												 sContext = JavaWindow("ScheduleManagerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
												 'objSchTable.ClickCell sIndex,"Object","RIGHT","NONE"						'Commented By Ketan as per discussion with vallari on 12/10/2010
												 JavaWindow("ScheduleManagerWindow").WinMenu("ContextMenu").Select sContext
'												 msgbox Err.number
										Case "1"
												sContext = JavaWindow("ScheduleManagerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
												'objSchTable.ClickCell sIndex,"Object","RIGHT","NONE" 						'Commented By Ketan as per discussion with vallari on 12/10/2010
												JavaWindow("ScheduleManagerWindow").WinMenu("ContextMenu").Select sContext
										Case Else
												'Context Menu Case NOT Exists for Supplied Menu
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Context Menu Case NOT Exists for Supplied Menu")
												Fn_SchMgr_SchTable_NodeOperation = FALSE
								End Select
										
					Fn_SchMgr_SchTable_NodeOperation = TRUE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation passed with case "&sAction&" on Object "&sObject)
            Else
			    Fn_SchMgr_SchTable_NodeOperation = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
				Exit Function
			End If

		'.---------------------------------------This case is used to  verify cell of Schedule Table  .----------------------------------------------
		Case "CellVerify"
			Dim sActValue,aActValue

'$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$ Note  $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'1 For Status [In Progress]  pass the value to be verified as  [In_Progress]
'2 For Status [Needs Attention]  pass the value to be verified as  [Needs_Attention]
'3 For Status [Complete,Abandoned & Late]  pass the value to be verified as  [Complete,Abandoned & Late Respectively]

'$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

    If   sColName = "Status Indicator" Then
			sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
				If instr(sIndex, "#") > 0 Then
				'sIndex = "#" + cstr(sIndex)
						objSchTable.SelectRow sIndex
						sActValue = objSchTable.GetCellData(sIndex,sColName)
						If instr(1,lcase(sActValue),lcase(sValue))>0 Then
							 Fn_SchMgr_SchTable_NodeOperation = TRUE
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation:Successfully Verified Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
						Else
							Fn_SchMgr_SchTable_NodeOperation = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation: Actual Value  ["+ sActValue +"] dose not matches with given Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
							Exit Function
						End If
			Else
					Fn_SchMgr_SchTable_NodeOperation = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
					Exit Function
		  End If
	Else	

			 sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")

			If instr(sIndex, "#") > 0 Then

				'Calculate the instance number
'				sIndex = cint(sIndex) + iInstance - 1
				'sIndex = "#" + cstr(sIndex)

				objSchTable.SelectRow sIndex
				sActValue = objSchTable.GetCellData(sIndex,sColName)
				If sColName = "Start Date"  Or sColName = "Finish Date" Or sColName = "Actual Start Date" Or sColName = "Actual Finish Date" Then
					aActValue =  Split(sActValue," " ,-1,1)
					sActValue =  aActValue(0)
				ElseIf  sColName = "Baseline Start" Or sColName = "Baseline Finish"Then 
					aActValue =  Split(sActValue," " ,-1,1)
					sActValue =  aActValue(0)
				End If

				 If IsNumeric(sActValue) And Trim(sColName) <> "Successors" And Trim(sColName) <> "Predecessors" Then
					objSchTable.ClickCell sIndex,sColName
					sActValue = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaList("SchTableCellList").GetItem(sActValue)
				 End If

					If   sColName = "Status Indicator" Then
						
					End If

				If IsArray(sValue )Then
						For i=0 to Ubound(sValue)
							If   sColName = "Resource Assignment" Then
									If  Instr(sValue(i), "Members:") >0 Then    ''For build prior to Tc10, value included 'Members:', Tc10 onwards 'Members:" should be removed.
										sValue(i) = Replace(sValue(i),"Members:","")
									End If
									If Right(sValue(i),1) = "%" Then                '' Tc10 onwards, resource Level % is included within parenthesis ()  e.g. AutoTest2 (autotest2)(50%)
											iPos = Instr(sValue(i),")")
                                            sValue(i)  = Left(sValue(i), iPos) & "(" & Right(sValue(i), Len(sValue(i))-iPos) & ")"
									End If
							End If
							If Instr(sActValue, sValue(i)) =0 Then
									Fn_SchMgr_SchTable_NodeOperation = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation: Actual Value  ["+ sActValue +"] dose not matches with given Value [" + sValue(i) + "] for  [" + sObject + "] and Column [" + sColName + "]")
									Exit Function
							End If
						Next
						 Fn_SchMgr_SchTable_NodeOperation = TRUE
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation:Successfully Verified Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
				Else
					 If Trim(Cstr(sActValue)) = Trim(Cstr(sValue)) Then
						 Fn_SchMgr_SchTable_NodeOperation = TRUE
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation:Successfully Verified Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
					 Else
						Fn_SchMgr_SchTable_NodeOperation = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation: Actual Value  ["+ sActValue +"] dose not matches with given Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
						Exit Function
					 End If
				End If
				objSchTable.ClickCell sIndex,"#0","LEFT","NONE"
		  Else
			Fn_SchMgr_SchTable_NodeOperation = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
			Exit Function
		  End If
	End if
	''-----------------------------------------------------------------Added By Ashok kakade--------------------------------------------------------------------------------
	''-----------------------------------------------------------------This case used to Cell verify of Same Instance--------------------------------------------------------------------------------
	Case "CellVerifyWithSameInstance"
			Dim sActVal,aActVal
			If instr(sObject, "@") <> 0 Then
				aNodePath = split(sObject, "@",-1, 1)
				sObject = aNodePath(0)
				iInstance = cint(aNodePath(1))
				iCnter = 1
				sPath=split(sObject, ":",-1, 1)
				For iCounter=1 to objSchTable.Object.getRowCount-1			
					If  sPath(0)+":"+objSchTable.Object.getRow(iCounter).tostring()= sObject Then
						If  iCnter = iInstance Then
							sIndex = iCounter
							objSchTable.SelectRow sIndex
						End If
						iCnter=iCnter+1
					End If
				Next
			End If
			sIndex = "#"+cstr(sIndex)
			If sColName = "Status Indicator" Then
				If instr(sIndex, "#") > 0 Then
					sActVal = objSchTable.GetCellData(sIndex,sColName)
					If instr(1,lcase(sActVal),lcase(sValue))>0 Then
						Fn_SchMgr_SchTable_NodeOperation = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation:Successfully Verified Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
					Else
						Fn_SchMgr_SchTable_NodeOperation = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation: Actual Value  ["+ sActVal +"] dose not matches with given Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
						Exit Function
					End If
				Else
					Fn_SchMgr_SchTable_NodeOperation = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
					Exit Function
				End If
			Else	
'			 sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
				If instr(sIndex, "#") > 0 Then
					sActVal = objSchTable.GetCellData(sIndex,sColName)
					If sColName = "Start Date"  Or sColName = "Finish Date" Or sColName = "Actual Start Date" Or sColName = "Actual Finish Date" Then
						aActVal =  Split(sActVal," " ,-1,1)
						sActVal =  aActVal(0)
					ElseIf  sColName = "Baseline Start" Or sColName = "Baseline Finish"Then 
						aActVal =  Split(sActVal," " ,-1,1)
						sActVal =  aActVal(0)
					End If
					If IsNumeric(sActVal) And Trim(sColName) <> "Successors" And Trim(sColName) <> "Predecessors" Then
						objSchTable.ClickCell sIndex,sColName
						sActVal = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaList("SchTableCellList").GetItem(sActVal)
					End If

					If sColName = "Status Indicator" Then						
					End If

					If IsArray(sValue )Then
						For i=0 to Ubound(sValue)
							If   sColName = "Resource Assignments" Then
									If  Instr(sValue(i), "Members:") >0 Then    ''For build prior to Tc10, value included 'Members:', Tc10 onwards 'Members:" should be removed.
										sValue(i) = Replace(sValue(i),"Members:","")
									End If
									If Right(sValue(i),1) = "%" Then                '' Tc10 onwards, resource Level % is included within parenthesis ()  e.g. AutoTest2 (autotest2)(50%)
											iPos = Instr(sValue(i),")")
                                            sValue(i)  = Left(sValue(i), iPos) & "(" & Right(sValue(i), Len(sValue(i))-iPos) & ")"
									End If
							End If
							If Instr(sActVal, sValue(i)) =0 Then
									Fn_SchMgr_SchTable_NodeOperation = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation: Actual Value  ["+ sActVal +"] dose not matches with given Value [" + sValue(i) + "] for  [" + sObject + "] and Column [" + sColName + "]")
									Exit Function
							End If
						Next
						 Fn_SchMgr_SchTable_NodeOperation = TRUE
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation:Successfully Verified Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
					Else
						If Trim(Cstr(sActVal)) = Trim(Cstr(sValue)) Then
							Fn_SchMgr_SchTable_NodeOperation = TRUE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation:Successfully Verified Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
						Else
							Fn_SchMgr_SchTable_NodeOperation = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation: Actual Value  ["+ sActVal +"] dose not matches with given Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
							Exit Function
						End If
					End If
					objSchTable.ClickCell sIndex,"#0","LEFT","NONE"
				Else
					Fn_SchMgr_SchTable_NodeOperation = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
					Exit Function
				End If
			End if
'					.---------------------------------------This case is used to  Verify cell with TIME & Date value of Schedule Table  .----------------------------------------------
		Case "CellTimeVerify"
					sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")

					If instr(sIndex, "#") > 0 Then

								'Calculate the instance number
				'				sIndex = cint(sIndex) + iInstance - 1
								'sIndex = "#" + cstr(sIndex)
			
								objSchTable.SelectRow sIndex
								sActValue = objSchTable.GetCellData(sIndex,sColName)
								'[TC1123(20161205c00)_PoonamC_NewDevelopment_22Feb2017 : Added column conditions for "Baseline Start" , "Baseline Finish"]
								If sColName = "Start Date"  Or sColName = "Finish Date" Or sColName = "Actual Start Date" Or sColName = "Actual Finish Date" Or sColName = "Baseline Start" Or sColName = "Baseline Finish"  Then
										If Cstr(sActValue) = Cstr(sValue) Then
													 Fn_SchMgr_SchTable_NodeOperation = TRUE
													 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation:Successfully Verified Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
										Else
													Fn_SchMgr_SchTable_NodeOperation = FALSE
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation: Actual Value  ["+ sActValue +"] dose not matches with given Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
													Exit Function
										End If
										objSchTable.ClickCell sIndex,"#0","LEFT","NONE"
								Else
												Fn_SchMgr_SchTable_NodeOperation = FALSE
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
												Exit Function
								 End If
					Else
								Fn_SchMgr_SchTable_NodeOperation = FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
								Exit Function
				  End If
       	'.---------------------------------------This case is used to  edit cell of Schedule Table  .----------------------------------------------
		Case "CellEdit","CellEditWithMsgVerify","CellEditWithoutMsgVerify"

		Dim scolIndex
		 sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")

		If instr(sIndex, "#") > 0 Then
				 'Calculate the instance number
'				 sIndex = cint(sIndex) + iInstance - 1
				sIndex = Right(sIndex, Len(sIndex) -1)

        	  'Get column index
				 scolIndex =  Fn_SchMgr_TableColIndex(objSchTable, sColName)

				 If scolIndex <> False Then
						scolIndex = cint(scolIndex)

						If  instr(objSchTable.Object.getCellEditor(cint(sIndex),scolIndex).tostring, "ComboBoxCellEditor") > 0  Or sColName = "Workflow Task Template" Or sColName = "Workflow Template" Then
								objSchTable.ClickCell sIndex,scolIndex, "LEFT","NONE"
								If sColName = "Workflow Task Template" Or sColName = "Workflow Template" Then
									If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaList("SchTableCellList").Exist = false then
												JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaList("SchTableCellList").SetToProperty "attached text","Work"
									end if 
									iIndex = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaList("SchTableCellList").GetItemIndex (sValue)
									JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaList("SchTableCellList").Object.setSelectedIndex Cint(iIndex)	
								Else 
									JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaList("SchTableCellList").Select sValue
								End If
						 elseif instr(objSchTable.Object.getCellEditor(cint(sIndex),scolIndex).tostring, "Button") > 0  then
								objSchTable.ClickCell sIndex,scolIndex, "LEFT","NONE"
								If sColName = "Resource Assignments" Or "Task Deliverables" Then
										Set objSelectType = description.Create()
										objSelectType("attached text").value = "edit_16"
										objSelectType("Class Name").value = "JavaButton"
										Set obj = Window("SchMgrWin").JavaWindow("JApplet").ChildObjects(objSelectType)    
										obj(0).Click  
										Fn_SchMgr_SchTable_NodeOperation = true
										Exit function
								end if
						ElseIf instr(objSchTable.Object.getCellEditor(cint(sIndex),scolIndex).tostring, "ScheduleTreeTableDateEditorRenderer") > 0  then

									JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").DoubleClickCell cint(sIndex),sColName,"LEFT","NONE"
									Call Fn_ReadyStatusSync(2)' Added by Nilesh for Build change 2012020800

									If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").Exist (10) Then
											JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").WaitProperty "enabled",1,20000
											JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click
									End If

									Call Fn_ReadyStatusSync(1)' Added by Nilesh for Build change 2012020800
									Set objShiftWin = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet")
									sDate  = sValue
									aDate = Split(sValue, " ",-1,1)
									'Extract Actual Date and Split it
									sActDate = objShiftWin.JavaEdit("ShiftDate").GetROProperty("value")
									aActDate = Split(sActDate, "-", -1,1)
									'Call Fn_SISW_UI_JavaEdit_Operations("Fn_SchMgr_SchTable_NodeOperation", "Set", objShiftWin, "ShiftDate", Trim(aDate(0)))  ' Added by Vivek Ahirrao
									If Trim(sActDate) <> Trim(aDate(0)) Then
										objShiftWin.JavaEdit("ShiftDate").RefreshObject	
										wait 3
									    objShiftWin.JavaEdit("ShiftDate").Click 1,1
									    wait 10
									    objShiftWin.JavaEdit("ShiftDate").Set Trim(aDate(0))
									    wait 5
									End If							
									Wait 1
									Call Fn_ReadyStatusSync(1)
									'Need to check the calender is popup or not
									If JavaWindow("ScheduleManagerWindow").JavaWindow("CalenderShell").Exist(2) Then
										Wait 3
									End If
									If Ubound(aDate) = 1 Then
										objShiftWin.JavaList("ShiftDate").Click 1,1,"LEFT"
										Wait 1
										Call Fn_ReadyStatusSync(1)
										objShiftWin.JavaList("ShiftDate").Select Trim(aDate(1))
										wait 1
										objShiftWin.JavaList("ShiftDate").RefreshObject	
										'set WshShell = CreateObject("WScript.Shell")
										'WshShell.SendKeys "{ESC}"
										JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Click 0,0
										wait 3
									End If
'                                    aDate = Split(sValue, "-", -1,1)
'									'Extract Actual Date and Split it
'									sActDate = objShiftWin.JavaCheckBox("ShiftDate").GetROProperty("label")
'									aActDate = Split(sActDate, "-", -1,1)
'									'Calculate Date Differences
'									iMonDiff = DateDiff("m", sActDate, sDate)
'									iYrDiff = DateDiff("yyyy", sActDate, sDate)
'									iDiff = iMonDiff - (iYrDiff * 12)
'									'Decide Scroll Direction
'									If iDiff > 0 Then
'										set objButton = objShiftWin.JavaButton("ScrollRight")
'									Elseif iDiff < 0 Then
'										set objButton = objShiftWin.JavaButton("ScrollLeft")
'									End If
'									'Set Year
'									objShiftWin.JavaEdit("Year").Set aDate(2)
'									objShiftWin.JavaEdit("Year").Activate
'									'Scroll to Get Proper Month
'									For iCnt = 1 to abs(iDiff)
'										objButton.Click micLeftBtn
'									Next
'									
'									'Set Required Date Digit
'									objShiftWin.JavaCheckBox("DateDigit").SetTOProperty "attached text", cstr(cint(aDate(0)))
'									objShiftWin.JavaCheckBox("DateDigit").Click 2,2,"LEFT"
'
'									wait(1)' Added by Nilesh for Build change 2012020800

'									objShiftWin.JavaButton("DateOK").Click micLeftBtn
									wait 1
'									JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Activate
									JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").ActivateRow "#3"
									If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").Exist (10) Then
										JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").WaitProperty "enabled",1,20000
										JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click
									End If
'									objShiftWin.JavaCheckBox("ShiftDate").Click 5,5,"LEFT"
									JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").ClickCell 0,0

						 Else
'							objSchTable.DoubleClickCell sIndex,scolIndex,"LEFT","NONE"
'							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Set sValue
'							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Activate
							'*objSchTable.SetCellData sIndex, scolIndex,sValue

							 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
                             JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Set ""
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setText sValue
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable False
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").ClickCell 0,0
						End If
						wait 4
							'[TC1123(20161205c00)_PoonamC_NewDevelopment_PPM_01Feb2017:Added Case "CellEditWithoutMsgVerify" after modifying cell not to verify any scheduling message]
							If sAction = "CellEditWithoutMsgVerify"  Then
								Fn_SchMgr_SchTable_NodeOperation = True
								Exit Function
							End If
						
							If sAction = "CellEditWithMsgVerify" Then
							   bReturn = Fn_SchMgr_DialogMsgVerify("Scheduling Error",sMenu,"OK") 
							   Fn_SchMgr_SchTable_NodeOperation = bReturn
							   Exit Function
							Else
							   bReturn = Fn_SchMgr_DialogMsgVerify("Scheduling Error","","OK") 
							End if
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Handle Scheduling Error Error dialog.")
								Fn_SchMgr_SchTable_NodeOperation = FALSE
								sErrorText = ""
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow sIndex
								'[TC1121-09_11-2015-2015102600-VivekA-Maintenance] - Added by Kaveri P
								If Fn_SISW_UI_Object_Operations("Fn_SchMgr_SchTable_NodeOperation","Exist", JavaWindow("ScheduleManagerWindow").Dialog("ErrorDialog"),"") = True Then
									bReturn = Fn_SchMgr_DialogMsgVerify("Scheduling Error","","OK")	
									If bReturn Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Handle Scheduling Error Error dialog.")
										Fn_SchMgr_SchTable_NodeOperation = FALSE
									End If	
								End If
								'--------------------------------------------------
								Exit Function
							End If
							
							'Handled Update Dependancy Error dialog appears while editing cell from schedule table
							bReturn = Fn_SchMgr_DialogMsgVerify("UpdateDependenciesRunner","","OK") 
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Handle Update Dependencies Error dialog.")
								Fn_SchMgr_SchTable_NodeOperation = FALSE
								sErrorText = ""
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow sIndex
								Exit Function
							End If

							'Handle Scheduling Error If Pops Up
							bReturn = Fn_SchMgr_SchedulingErrorVerify("", "OK")
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error on ScheduleTable Node Editing")
								Fn_SchMgr_SchTable_NodeOperation = FALSE
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow sIndex
								Exit Function
							End If

							'Handle Warning on Cell Editing
							bReturn = Fn_SchMgr_DialogMsgVerify("Warning", sErrorText,"OK")
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning on ScheduleTable Node Editing")
								Fn_SchMgr_SchTable_NodeOperation = FALSE
								sErrorText = ""
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow sIndex
								Exit Function
							End If

'							Handle Error on Cell Editing
							bReturn = Fn_SchMgr_DialogMsgVerify("Error", sErrorText,"OK")
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error on ScheduleTable Node Editing")
								Fn_SchMgr_SchTable_NodeOperation = FALSE
								sErrorText = ""
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow sIndex
								Exit Function
							End If

							'Handle error 
							bReturn = Fn_SchMgr_DialogMsgVerify("Validate Inline Editing","","OK") 
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Handle error Validate Inline Editing")
								Fn_SchMgr_SchTable_NodeOperation = FALSE
								sErrorText = ""
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow sIndex
								Exit Function
							End If


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$----By Shreyas----$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''$$$$
'''$$$$					Commenting this segmnt of the code as the focus does shift after editing the field in "SchTaskTable" on the build 0330
'''$$$$					This was an issue on earlier builds
'''$$$$
''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


'							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
'							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
'							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow "#0"
'							'objSchTable.ActivateCell sIndex, scolIndex
'							'objSchTable.ClickCell sIndex,"#0","LEFT","NONE"

				 End If
	
				 If Err.Number <  0 Then
					 Fn_SchMgr_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation: Fail to edit Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "]")	
					 Exit Function
				  Else
						Fn_SchMgr_SchTable_NodeOperation = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Successfully edit Cell Value [" + sValue + "] for  [" + sObject + "] and Column [" + sColName + "].")   					
				End If
		 Else 
				Fn_SchMgr_SchTable_NodeOperation = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
				Exit Function
		 End If

        '.---------------------------------------This case is used to cell  double click of Schedule Table  .----------------------------------------------
		Case "CellDoubleClick"
			sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
			
            If Instr(sIndex, "#") > 0 Then

'				sIndex = cint(sIndex) + iInstance - 1
				'sIndex = "#" + cstr(sIndex)

				objSchTable.DoubleClickCell sIndex,sColName
				
				If Err.Number >= 0 Then
					Fn_SchMgr_SchTable_NodeOperation = TRUE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Successfully Double-Clicked Cell at  [" + sObject + "] and Column [" + sColName + "]")
				Else
					Fn_SchMgr_SchTable_NodeOperation = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Failed to Double-Clicked Cell at  [" + sObject + "] and Column [" + sColName + "]")
					Exit Function
				End If

			 Else
			    Fn_SchMgr_SchTable_NodeOperation = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
				Exit Function 
			End If

	'.---------------------------------------This case is used  to check exists of node in schedule table .  .----------------------------------------------
		Case "Exists"
				Dim sActualval
				sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
				

              If Instr(sIndex, "#") > 0 Then
						'Calculate the instance number
'						sIndex = cint(sIndex) + iInstance - 1
						'sIndex = "#" + cstr(sIndex)

		 			    objSchTable.SelectRow sIndex
						sActualval = objSchTable.GetCellData(sIndex,"Object")

						If  sActualval = sObject  Then
							Fn_SchMgr_SchTable_NodeOperation = TRUE				 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_SchMgr_SchTable_NodeOperation: Row with Object "&sObject&" is exists.")	
						Else 
							Fn_SchMgr_SchTable_NodeOperation = FALSE				 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists.")	
							Exit Function
					   End If
						
				Else 
							Fn_SchMgr_SchTable_NodeOperation = FALSE				 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists.")	
							Exit Function	
				End If

		'.---------------------------------------This case is used  expand particular node in schedule table .  .----------------------------------------------

			Case "ExpandNode"

				sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
'				sIndex = cint(sIndex) + iInstance - 1
				If  Instr(sIndex, "#") > 0 Then
					sIndex = Right(sIndex, Len(sIndex) -1)
					Set objTree = objSchTable.Object.getTree
					Set row = objTree.getPathForRow(Cint(sIndex))
					objSchTable.Object.expandPath row

					If Err.Number < 0 Then
						Fn_SchMgr_SchTable_NodeOperation = FALSE		
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation :Fail to expand node " +  sObject )	
						Exit Function 
					Else
						Fn_SchMgr_SchTable_NodeOperation = TRUE	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Successfully expand node " + sObject )				 
					End If
				Else 
					Fn_SchMgr_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists.")	
					Exit Function	
				End If

			'.---------------------------------------This case is used  collapse particular node in schedule table .  .----------------------------------------------

			Case "CollapseNode"

				sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
'				sIndex = cint(sIndex) + iInstance - 1
				If  Instr(sIndex, "#") > 0 Then
					sIndex = Right(sIndex, Len(sIndex) -1)
					Set objTree = objSchTable.Object.getTree
					Set row = objTree.getPathForRow(Cint(sIndex))
					objSchTable.Object.collapsePath row

					If Err.Number < 0 Then
						Fn_SchMgr_SchTable_NodeOperation = FALSE		
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation :Fail to collapse node " +  sObject )	
						Exit Function 
					Else
						Fn_SchMgr_SchTable_NodeOperation = TRUE	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Successfully collapse node " + sObject )				 
					End If
				Else 
					Fn_SchMgr_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists.")	
					Exit Function	
				End If
				'---------------------------------DeSelect the specified task/schedule--------------------------------------------------------'
			Case 	"DeSelect"

				sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
				If Instr(sIndex, "#") > 0 Then
					'Calculate the instance number
					'sIndex = "#" + cstr(sIndex)
					'Select the Expected  scheduleTable Node
					 objSchTable.DeselectRow sIndex

					 If Err.Number <  0 Then
						 Fn_SchMgr_SchTable_NodeOperation = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation : Failed to Deselect Row with Object "&sObject& " from Schedule Table .")	
						Exit Function
					Else
						Fn_SchMgr_SchTable_NodeOperation = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation: Successfully selected Row with Object "&sObject&" from Schedule Table.")	
					End If
				Else
					Fn_SchMgr_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists in schedule table..")	
					Exit Function
				End If
	
			'.---------------------------------------This case is used  to verify Index in schedule table .  .----------------------------------------------------------------

			Case "VerifyIndex"
				sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
				
'				sIndex = cint(sIndex) + iInstance - 1
				If  Instr(sIndex, "#") > 0 Then
					sIndex = Right(sIndex, Len(sIndex) -1)
					If Cint(sValue) = Cint(sIndex) Then
						Fn_SchMgr_SchTable_NodeOperation = TRUE	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Succesfully verify index of node " & sObject)				 
					Else
						Fn_SchMgr_SchTable_NodeOperation = FALSE	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation :Failed to verify index of node " & sObject)				 
						Exit Function	
					End If
				Else 
					Fn_SchMgr_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists.")	
					Exit Function	
				End If

			Case  "MultiSelectPopup"

			 'Dim ArrNode(),IntRows
			 ArrNodes = split(sObject, ",",-1,1)
			 objSchTable.Object.clearSelection
			 ReDim ArrNode(Ubound(ArrNodes))
			 For IntArrCounter = 0 to Ubound(ArrNodes)
					sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, ArrNodes(IntArrCounter), "Object")

					 If Instr(sIndex, "#") > 0 Then						
						'sIndex = "#" + cstr(sIndex)
						ArrNode(IntArrCounter) = sIndex
					Else
						Fn_SchMgr_SchTable_NodeOperation = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists in schedule table..")	
						Exit Function
					End If
			 Next

			For iCounter = 0 to Ubound(ArrNode)
				objSchTable.ExtendRow ArrNode(iCounter)
				If Err.Number <  0 Then
						 Fn_SchMgr_SchTable_NodeOperation = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Row with Object "&ArrNodes(IntArrCounter)&" does not selected for mutiselect case.")	
						Exit Function
				Else
						Fn_SchMgr_SchTable_NodeOperation = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation: Row with Object "&ArrNodes(IntArrCounter)&" is selected for multiselect case..")	
				End If
			Next
			
			objSchTable.ClickCell ArrNode(iCounter-1),"Object","RIGHT","CONTROL"
			aMenuList = split(sMenu, ":",-1,1)			
			sCount = Ubound(aMenuList)
				Select Case sCount
						Case "0"
								 sContext = JavaWindow("ScheduleManagerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
								 Wait 1
								 JavaWindow("ScheduleManagerWindow").WinMenu("ContextMenu").Select sContext
						Case "1"
								sContext = JavaWindow("ScheduleManagerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
								JavaWindow("ScheduleManagerWindow").WinMenu("ContextMenu").Select sContext
						Case Else
								'Context Menu Case NOT Exists for Supplied Menu
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Context Menu Case NOT Exists for Supplied Menu")
								Fn_SchMgr_SchTable_NodeOperation = FALSE
								Exit Function
				End Select  

		Case "ClearAll"
			objSchTable.Object.clearSelection
			If Err.Number < 0 Then
				Fn_SchMgr_SchTable_NodeOperation = FALSE		
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation :Fail to Clear Selection")	
				Exit Function 
			Else
				Fn_SchMgr_SchTable_NodeOperation = TRUE	
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_SchMgr_SchTable_NodeOperation : Successfully Cleared Selectiopn" )				 
			End If
			
		Case "GetCellValue"
			sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")

			If Instr(sIndex, "#") > 0 Then
			
				'sIndex = "#" + cstr(sIndex)

				objSchTable.SelectRow sIndex
				sActValue = objSchTable.GetCellData(sIndex,sColName)

				If sActValue<>"" Then
					Fn_SchMgr_SchTable_NodeOperation=sActValue
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fn_SchMgr_SchTable_NodeOperation:Successfully Retrieved Cell Value [" + sActValue + "] for  [" + sObject + "] and Column [" + sColName + "]")
					Exit Function
				Else
					Fn_SchMgr_SchTable_NodeOperation = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not selected.")
					Exit Function	
				End If
			Else
					Fn_SchMgr_SchTable_NodeOperation = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Row with Object "&sObject&" does not exists in schedule table..")	
					Exit Function
			End If
	'------------------------------------------------------------------------------------------------------------------------------------------		
		'[TC1123_20161205c00_PoonamC_NewDevelopment_17Feb2017_Added New Case : VerifyBackGroundColor to verify back ground color of node]	
	    Case "VerifyBackGroundColor"
            sIndex = Fn_SchMgr_TreeTableRowIndex(objSchTable, sObject, "Object")
			If sIndex = False Then
				Fn_SchMgr_SchTable_NodeOperation = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_SchTable_NodeOperation:Row with Object "&sObject&" does not exists.")
				Exit Function
			Else	
				sIndex = Replace(sIndex,"#","")
				sNodecolor = objSchTable.Object.getBackgroundColorForRow(sIndex,False).toString()
				sNodecolor = Mid(sNodecolor,instr(1,sNodecolor,"["),instr(1,sNodecolor,"]"))
				arrcolor = Split(Replace(Replace(sNodecolor,"[",""),"]",""),",")
				sNodecolor = Split(arrcolor(0),"=")
				rColor =  sNodecolor(1)
				sNodecolor = Split(arrcolor(1),"=")
				gColor = sNodecolor(1)
				sNodecolor = Split(arrcolor(2),"=")
				bColor = sNodecolor(1)
				
				Select Case sValue
					  Case "RED"
					      If rColor = 255 and gColor = 0 and bColor = 0 Then
					      		Fn_SchMgr_SchTable_NodeOperation = TRUE
					      Else
					      		Fn_SchMgr_SchTable_NodeOperation = FALSE
					      End If
					  Case "WHITE" 
					  	   If rColor = 255 and gColor = 255 and bColor = 255 Then
					      		Fn_SchMgr_SchTable_NodeOperation = TRUE
					      Else
					      		Fn_SchMgr_SchTable_NodeOperation = FALSE
					      End If
					  Case else
					     Fn_SchMgr_SchTable_NodeOperation = FALSE				 
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation : Invalid Case "&sValue&".")	
						 Exit Function
				End Select				
			End If	
	'------------------------------------------------------------------------------------------------------------------------------------------	
	
	End Select
	
	
	Fn_SchMgr_SchTable_NodeOperation = TRUE				 
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_SchTable_NodeOperation passed with case "&sAction&" on Object "&sObject)

	Set objSchTable = nothing 
	Set objTree = nothing
Else
	Fn_SchMgr_SchTable_NodeOperation = FALSE
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Fn_SchMgr_SchTable_NodeOperation :Schedule Manager table does not exist.")
End If

End Function


'*************************************************Function to Create a new Task****************************************************************************
'Function Name		         :			   Fn_SchMgr_TaskCreate  

'Description			         :		 	   Create basic task 


'Parameters			         :	 		     1. SsAction: Basic or quick task creation
'									                      2.sAction: Basic or quick task creation
'									                     3.sTaskId: Task ID
'					 				                    4.sTaskrev: Task Rev ID
'									                   5. sTaskname: Name of the task
'                                                     6.sTaskDesc:Description of the task.
'                                                    7.sTaskDuration:Time assign for the task.


'Return Value		       : 		   TRUE \ FALSE

'Pre-requisite			   :		 	Shedular Manager need to be open
 '                                                 Select the Schedule under which task is create.

'Examples				  :			   Fn_SchMgr_TaskCreate("BasicCreation","ScheduleTask","111101","A","Task4","Enter","12h") .Call Fn_Fn_SchMgr_TaskCreate"BasicCreation","ScheduleTask","111101","A","Task4","Enter","12h")

'History				     :		
'									                Developer Name					        Date					         Rev. No.			            Changes Done						Reviewer
'								              ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									               Manish Verma				                18/05/2010			              1.0										Created
'													Vallari S.										 21/05/2010							1.0										Modified for Error Handling and Win Ref releasing
'								              --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ganesh B.								18-Jul-2014								1.1									Modified case"BasicCreate" to handle "New Task" Dialog asper desgin changes on TC11.1(20140709) 
'								              --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_TaskCreate(sAction,sTaskType,sTaskId,sTaskrev,sTaskName,sTaskDesc,sTaskDuration) 
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TaskCreate"
	Dim objTask, bReturn
	Dim aType
	Dim  iItemCount, iCount, crrItem, bFlag
	Dim sNewItemMenu
	On Error Resume Next 
	

		Select Case sAction
			Case "BasicCreate"
					sNewItemMenu = Fn_GetXMLNodeValue( Fn_LogUtil_GetXMLPath("RAC_Menu"),"FileNewTask")
					Set objTask=Fn_SISW_PPM_GetObject("New Task")
					If objTask.Exist (SISW_MIN_TIMEOUT) = False  Then
			             'Select menu  [File -> New -> Task...]       
						bReturn = Fn_MenuOperation("Select",sNewItemMenu)
						Call Fn_ReadyStatusSync(1)
						If bReturn = False Then
								Fn_SchMgr_TaskCreate = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [File:New:Task...]")
								Set objTask = Nothing
								Exit Function
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [File:New:Task...]")
						End If
					End If

					'Check the existence of the "NewTask" Window
					If objTask.Exist (SISW_MIN_TIMEOUT)  Then
						'Added by Vallari, to remove New line character i.e chr(10) that is appended to value from DataTable 
						aType = split(sTaskType, chr(10))
						sTaskType = aType(0)
'						Added by Vallari, as Tc91_1121 onwards type changed for localization testing
						If trim(sTaskType) = "ScheduleTask" Then
							sTaskType = "Schedule Task"
						End If
							'Select  "Task Type"
	 							' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	 					If objTask.JavaTree("TaskType").Exist(SISW_MIN_TIMEOUT) Then
							iItemCount=Fn_UI_Object_GetROProperty("Fn_SchMgr_TaskCreate",objTask.JavaTree("TaskType"), "items count")
							For iCount=0 To iItemCount-1
								crrItem=objTask.JavaTree("TaskType").GetItem(iCount)
								If Trim(crrItem)="Most Recently Used:"+Trim(sTaskType) Then
									bFlag=True
									Exit For
								ElseIf Trim(crrItem)="Complete List" Then
									Exit For
								End If
							Next
						
							If bFlag=True Then
								Call Fn_JavaTree_Select("Fn_SchMgr_TaskCreate", objTask, "TaskType","Most Recently Used")
								Call Fn_JavaTree_Select("Fn_SchMgr_TaskCreate", objTask, "TaskType","Most Recently Used:"+sTaskType)
							Else
								Call Fn_UI_JavaTree_Expand("Fn_SchMgr_TaskCreate", objTask, "TaskType","Complete List")
								Call Fn_JavaTree_Select("Fn_SchMgr_TaskCreate", objTask, "TaskType","Complete List")
								Call Fn_JavaTree_Select("Fn_SchMgr_TaskCreate", objTask, "TaskType","Complete List:"+sTaskType)	
							End If
							If Err.number < 0 Then
								Fn_SchMgr_TaskCreate = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Task Type [" + sTaskType + "]")
								Set objTask = Nothing
								Exit Function
							End If
							'Clicking On Next button
							objTask.JavaButton("Next").WaitProperty "enabled", 1, 60000
							call Fn_Button_Click("Fn_SchMgr_TaskCreate", objTask, "Next")
							If Err.number < 0 Then
								Fn_SchMgr_TaskCreate = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
								Set objTask = Nothing
								Exit Function
							End If
						End If
						   'Check  "Item Id and Revision ID"
							If sTaskId <> "" Then
								objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Task ID:"
								Call Fn_Edit_Box("Fn_SchMgr_TaskCreate", objTask,"TaskEdit",sTaskId)
								If Err.number < 0 Then
									Fn_SchMgr_TaskCreate = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Task Id as [" + sTaskId + "]")
									Set objTask = Nothing
									Exit Function
								End If
							End If
							'Check  "Item Id and Revision ID"
							If sTaskId = "" or sTaskrev = "" Then
								'Click on "Assign" button
								objTask.JavaButton("Assign").WaitProperty "enabled", 1, 20000
								call Fn_Button_Click("Fn_SchMgr_TaskCreate", objTask, "Assign")
								If Err.number < 0 Then
									Fn_SchMgr_TaskCreate = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Assign] Button")
									Set objTask = Nothing
									Exit Function
								End If
							End If
							sTskAssignId = objTask.JavaEdit("TaskEdit").GetROProperty("value")

							'Set the Task  Name
								objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Name:"
							Call Fn_Edit_Box("Fn_SchMgr_TaskCreate", objTask,"TaskEdit",sTaskName)
							 If Err.number < 0 Then
								Fn_SchMgr_TaskCreate = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Task Name as [" + sTaskName + "]")
								Set objTask = Nothing
								Exit Function
							End If
		
							'Set the Description
							objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Description:"
							Call Fn_Edit_Box("Fn_SchMgr_TaskCreate", objTask,"TaskEdit",sTaskDesc)
		
							'Click on "Finish" button
							objTask.JavaButton("Finish").WaitProperty "enabled", 1, 20000
							call Fn_Button_Click("Fn_SchMgr_TaskCreate", objTask, "Finish")
							If Err.Number < 0 Then
								Fn_SchMgr_TaskCreate = False																			      
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Finish] Button")
								Set objTask = Nothing
								Exit Function
							End If
						   Fn_SchMgr_TaskCreate =True																	                                                       
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task [" + sTaskName + "] Created Successfully") 
							Set objTask = Nothing
				Else 
						Fn_SchMgr_TaskCreate = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[New Task] Dialog Not Found") 
						Set objTask = Nothing
				End If

			Case "QuickCreate"

					Set objTask=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet")
					If objTask.Exist (SISW_MIN_TIMEOUT) = True Then
						Call Fn_SISW_UI_JavaEdit_Operations("Fn_SchMgr_TaskCreate", "Set",  objTask, "QuickTaskName", sTaskName )
						If sTaskDuration <>"" Then
							objTask.JavaEdit("QuickTaskWork").Set  sTaskDuration
						End If
						If Err.Number < 0 Then
								Fn_SchMgr_TaskCreate = False																			      
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set TaskName [" + sTaskName + "] OR Task Duration [" + sTaskDuration + "]")
								Set objTask = Nothing
								Exit Function
						 End If

						objTask.JavaButton("Create").WaitProperty  "enabled", 1, 20000
						objTask.JavaButton("Create").Click
							If Err.Number < 0 Then
								Fn_SchMgr_TaskCreate = False																			      
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Create] Button")
								Set objTask = Nothing
								Set objWin = nothing
								Exit Function
'						   Else
'								Fn_SchMgr_TaskCreate = True
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task [" + sTaskName + "] Created Successfully")
						   End If
						   Call Fn_ReadyStatusSync(2)

						'Handle Scheduling Error if Pops up
							bReturn = Fn_SchMgr_SchedulingErrorVerify("", "OK")
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error on Task Creation")
								Fn_SchMgr_TaskCreate = False
								Set objTask = Nothing
								Set objWin = nothing
								Exit Function
							End If

							Fn_SchMgr_TaskCreate = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task [" + sTaskName + "] Created Successfully")

					Else
							Fn_SchMgr_TaskCreate = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Schedule Manager] Window Not Found") 
							Set objTask = Nothing
					End If
						Set objTask = Nothing
						Set objWin = nothing
		End Select
End Function


'*************************************************Function to Create a new Milestone****************************************************************************
'Function Name		         :			   Fn_SchMgr_MilestoneCreate  

'Description			         :		 	   Create MilsStone


'Parameters			         :	 		     1. SAction: Basic or quick task creation
'									                      2.sTasktype: Basic or quick task creation
'									                     3.sMilestoneId: MilestoneIdID
'					 				                    4.sMileRevID: Milestone Rev ID
'									                   5. sMileName: Name of the Milestone
'                                                     6.sMileDesc:Description of the Milestone.
'                                                    7.bAutoComplete:True/False tag for auto complete
'													8.sDate:Milestonr date

'Return Value		       : 		   TRUE \ FALSE

'Pre-requisite			   :		 	Shedular Manager need to be open
 '                                                 Select the Schedule under which task is create.

'Examples				  :			   Fn_SchMgr_MilestoneCreate("BasicCreation","ScheduleTask","111101","A","Task4","Enter","on","") .Call Fn_Fn_SchMgr_TaskCreate"BasicCreation","ScheduleTask","111101","A","Task4","Enter","on","")

'History				     :		
'									                Developer Name					        Date					         Rev. No.			            Changes Done						Reviewer
'								              ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									               Manish Verma				                 20/05/2010			              1.0										Created
'													Vallari S.										 21/05/2010							1.0										Modified for Error Handling and Win Ref releasing
'								              --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      
Public Function Fn_SchMgr_MilestoneCreate(sAction,sTaskType,sMilestoneId,sMileRevID,sMileName,sMileDesc ,bAutoComplete,sDate) 
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_MilestoneCreate"
	Dim objMilestone, bReturn

	On Error Resume Next 
	
		Select Case sAction
			Case "BasicCreate"

'					Set objMilestone=Window("SchMgrWin").JavaDialog("New Milestone")
'					'If objTask.Exist (5) = False  Then
'					If objMilestone.Exist(5)=False Then
			           'Select menu  [File -> New -> Task...]               
						bReturn = Fn_MenuOperation("Select","File:New:Milestone...")
						If bReturn = False Then
								Fn_SchMgr_MilestoneCreate = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Create [Milestone]")
								Set objTask = Nothing
								Exit Function
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created [Milestone]")
								Fn_SchMgr_MilestoneCreate = True
						End If
'					End If

'					'Check the existence of the "NewTask" Window
'					If objMilestone.Exist (20)  Then
'						'Select  Task Type
'					    objMilestone.JavaList("MilestoneType").Select sTaskType
'						If Err.number < 0 Then
'							Fn_SchMgr_MilestoneCreate = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Task Type [" + sTaskType + "]")
'							Set objMilestone = Nothing
'							Exit Function
'						End If 											  
'						'Click the next Button
'						 objMilestone.JavaButton("Next").Click
'						 If Err.number < 0 Then
'							Fn_SchMgr_MilestoneCreate = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
'							Set objMilestone = Nothing
'							Exit Function
'						End If
'
'					   'Check  "Item Id and Revision ID"
'						If sMilestoneId <> "" Then
'							 'objTask.JavaEdit("TaskId").Set sTaskId 
'                             objMilestone.JavaEdit("MilestoneId").Set sMilestoneId
'							 If Err.number < 0 Then
'								Fn_SchMgr_MilestoneCreate = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Milestone Id as [" + sMilestoneId + "]")
'								Set objMilestone = Nothing
'								Exit Function
'							End If
'						End If
'						If sMileRevID <> "" Then
'							'objTask.JavaEdit("TaskRev").Set sTaskrev 
'							objMilestone.JavaEdit("MilestoneRev").Set sMileRevID
'							 If Err.number < 0 Then
'								Fn_SchMgr_MilestoneCreate = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Milestone RevId as [" + sMileRevID + "]")
'								Set objMilestone = Nothing
'								Exit Function
'							End If 
'						End If
'
'						If sMilestoneId = "" or sMileRevID = "" Then
'							objMilestone.JavaButton("Assign").Click
'							If Err.Number < 0 Then
'									Fn_SchMgr_MilestoneCreate = False
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Assign] Button")
'									Set objMilestone = Nothing
'									Exit Function
'							End If
'					    End If
'
'						sTskAssignId = objMilestone.JavaEdit("sMilestoneId").GetROProperty("value")
'					    sTskAssignRev = objMilestone.JavaEdit("sMileRevID").GetROProperty("value")
'
'                        'Set the Task  Name
'                        objMilestone.JavaEdit("MilestoneName").Set sMileName 
'						If Err.Number < 0 Then
'							Fn_SchMgr_MilestoneCreate = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Milestone Name as [" + sMileName + "]")
'							Set objMilestone = Nothing
'							Exit Function
'						End If
'						'Set the Description
'                         objMilestone.JavaEdit("Description").Set sMileDesc
'
'                         objMilestone.JavaButton("Finish").WaitProperty "enabled",1,20000
'                         objMilestone.JavaButton("Finish").Click
'						 If Err.Number < 0 Then
'							   Fn_SchMgr_MilestoneCreate = False																			      
'							   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Finish] Button")
'							  Exit Function
'					    End If					          
'                        Fn_SchMgr_MilestoneCreate =True																	                                                       
'					    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Milestone [" + sMileName + "] Created Successfully") 
'    	           Else 
'						Fn_SchMgr_MilestoneCreate = False
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[New Milestone] Dialog Not Found") 
'                  End If
				  Set objMilestone = Nothing

			Case "QuickCreate"
					Set objMilestone = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet")
						If objMilestone.Exist(SISW_DEFAULT_TIMEOUT) = True Then
							objMilestone.JavaEdit("QuickTaskName").Set  sMileName
							objMilestone.JavaEdit("QuickTaskWork").Set  "0h"
							objMilestone.JavaButton("Create").WaitProperty  "enabled", 1, 20000
						   objMilestone.JavaButton("Create").Click
							  If Err.Number < 0 Then
									 Fn_SchMgr_MilestoneCreate = False																			      
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Create] Button")
									 Exit Function
							  Else
									  Fn_SchMgr_MilestoneCreate = True
							 End If	
						Else
							Fn_SchMgr_MilestoneCreate = False																			      
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Schedule Manager] Window Not Found")
							Exit Function
						End If
						Set objMilestone = Nothing

'			 Case "ExtendedCreate"	
'				
'					 Set objMilestone=Window("SchMgrWin").JavaDialog("New Milestone")
'					'If objTask.Exist (5) = False  Then
'					If objMilestone.Exist(5)=False Then
'			           'Select menu  [File -> New -> Task...]               
'						bReturn = Fn_MenuOperation("Select","File:New:Milestone...")
'						If bReturn = False Then
'								Fn_SchMgr_MilestoneCreate = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [File:New:Milestone...]")
'								Set objTask = Nothing
'								Exit Function
'						Else
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [File:New:Milestone...]")
'						End If
'					End If
'
'					'Check the existence of the "NewTask" Window
'					If objMilestone.Exist (20)  Then
'						'Select  Task Type
'					    objMilestone.JavaList("MilestoneType").Select sTaskType
'						If Err.number < 0 Then
'							Fn_SchMgr_MilestoneCreate = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Task Type [" + sTaskType + "]")
'							Set objMilestone = Nothing
'							Exit Function
'						End If 											  
'						'Click the next Button
'						 objMilestone.JavaButton("Next").Click
'						 If Err.number < 0 Then
'							Fn_SchMgr_MilestoneCreate = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
'							Set objMilestone = Nothing
'							Exit Function
'						End If
'
'					   'Check  "Item Id and Revision ID"
'						If sMilestoneId <> "" Then
'							 'objTask.JavaEdit("TaskId").Set sTaskId 
'                             objMilestone.JavaEdit("MilestoneId").Set sMilestoneId
'							 If Err.number < 0 Then
'								Fn_SchMgr_MilestoneCreate = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Milestone Id as [" + sMilestoneId + "]")
'								Set objMilestone = Nothing
'								Exit Function
'							End If
'						End If
'						If sMileRevID <> "" Then
'							'objTask.JavaEdit("TaskRev").Set sTaskrev 
'							objMilestone.JavaEdit("MilestoneRev").Set sMileRevID
'							 If Err.number < 0 Then
'								Fn_SchMgr_MilestoneCreate = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Milestone RevId as [" + sMileRevID + "]")
'								Set objMilestone = Nothing
'								Exit Function
'							End If 
'						End If
'
'						If sMilestoneId = "" or sMileRevID = "" Then
'							objMilestone.JavaButton("Assign").Click
'							If Err.Number < 0 Then
'									Fn_SchMgr_MilestoneCreate = False
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Assign] Button")
'									Set objMilestone = Nothing
'									Exit Function
'							End If
'					    End If
'
'						sMilestoneId = objMilestone.JavaEdit("sMilestoneId").GetROProperty("value")
'					    sMileRevID = objMilestone.JavaEdit("sMileRevID").GetROProperty("value")
'
'                        'Set the Task  Name
'                        objMilestone.JavaEdit("MilestoneName").Set sMileName 
'						If Err.Number < 0 Then
'							Fn_SchMgr_MilestoneCreate = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Milestone Name as [" + sMileName + "]")
'							Set objMilestone = Nothing
'							Exit Function
'						End If
'						'Set the Description
'                         objMilestone.JavaEdit("Description").Set sMileDesc
'
'						 objMilestone.JavaButton("Next").WaitProperty "enabled",1,20000
'	        			  objMilestone.JavaButton("Next").Click
'							If Err.Number < 0 Then
'								   Fn_SchMgr_MilestoneCreate = False																			      
'								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Next] Button")
'								   Set objMilestone = Nothing
'								  Exit Function
'						   End If
'
'						'Set Auto Complete Check Box as per the input
'						If bAutoComplete <> "" Then
'							If CBool(bAutoComplete) = True Then
'									objMilestone.JavaCheckBox("AutoComplete").Set "ON"
'							Else
'									objMilestone.JavaCheckBox("AutoComplete").Set "OFF"
'							End If
'							If Err.Number < 0 Then
'								   Fn_SchMgr_MilestoneCreate = False																			      
'								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Operate Upon [Auto Complete] CheckBox")
'								   Set objMilestone = Nothing
'								  Exit Function
'						   End If
'						End If
'
'						'Set Milestone Date if passed in
'						If sDate <> "" Then
'							Set objDate = objMilestone.JavaCheckBox("MilestoneDate").Object
'							objDate.setDate sDate
'							Set objDate = Nothing
'							If Err.Number < 0 Then
'								   Fn_SchMgr_MilestoneCreate = False																			      
'								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Milestone Date to [" + sDate + "]")
'								   Set objMilestone = Nothing
'								  Exit Function
'						   End If
'						End If
'							   		          
'								 Fn_SchMgr_MilestoneCreate =True																	                                                       
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Milestone [" + sMileName + "] Created Successfully") 
'					   Else 
'							Fn_SchMgr_MilestoneCreate = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[New Milestone] Dialog Not Found") 
'					  End If
'					  Set objMilestone = Nothing
							   
     End Select
End Function


''*********************************************************		Function to Delete the Object from Schedule Manager		**************************************************************
'Function Name		:				Fn_SchMgr_TaskDelete.

'Description			 :		 		 This function is used to delete the Task  from Schedule Manager.

'Parameters			   :	 			1.  sAction

'Return Value		   : 				True/False

'Pre-requisite			:		 		Select the Task to be deleted.

'Examples				:				 call Fn_SchMgr_TaskDelete("Menu")
'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	Reviewed date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rucha                         20/05/10																							26-Mar-10
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_TaskDelete(sAction)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TaskDelete"
   Dim bReturn,strMenuPath,strMenu

	Select Case sAction
		Case "Menu"  

			bReturn = Fn_MenuOperation("Select","Edit:Delete")
			If bReturn = True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Edit;Delete]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Edit;Delete]")
				Fn_SchMgr_TaskDelete = False
				Exit Function
			End If

		Case "Toolbar" 
			strMenuPath= Fn_LogUtil_GetXMLPath("RAC_Toolbar")  '' used xml to get toolbar button name
            strMenu = Fn_GetXMLNodeValue(strMenuPath ,"Delete" )
			bReturn= Fn_ToolbatButtonClick(strMenu)
			If bReturn=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Toolbar Button [Delete]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Toolbar Button [Delete]")
				Fn_SchMgr_TaskDelete = False
				Exit Function
			End If

		Case "KeyBoard"
            bReturn= Fn_KeyBoardOperation("SendKey","{DEL}")
			If bReturn= True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Pressed KeyBoard Key [Delete]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Press KeyBoard Key [Delete]")
				Fn_SchMgr_TaskDelete = False
				Exit Function
			End If

	End Select

	JavaDialog("Confirmation").SetTOProperty "title", "Confirmation"
	If JavaDialog("Confirmation").Exist(20)   Then 
		JavaDialog("Confirmation").JavaButton("Yes").Click
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Yes] Button on [Confirmation] Dialog")
			Fn_SchMgr_TaskDelete = FALSE
		End If
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Deleted Task Successfully")	
		Fn_SchMgr_TaskDelete = True
	Else
		Fn_SchMgr_TaskDelete = FALSE
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Confirmation] Dialog Not Found")
	End If
End Function


'*********************************************************		Function to add task constraint.	***********************************************************************
'Function Name		:				Fn_SchMgr_TaskConstraint(sConstraint )

'Description			 :		 		 This Function is to add task constraint.

'Parameters			   :	 			1. sConstraint:Select the constraint for Task
											
'Return Value		  : 			

'Pre-requisite		  :		 		    Select Task from schedule table.

'Examples			:					Fn_SchMgr_TaskConstraint("As Soon As Possible")

'History:
'
'											Developer Name			Date						Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Rucha							21-May-2010				1.0
'												Vallari S.						24-May-2010				1.0						Error Handling and dialog checking
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_TaskConstraint(sConstraint)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TaskConstraint"
	Dim objDialogTaskConstraints, bReturn

	Set objDialogTaskConstraints=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Task Constraints")

		 'Select menu [Schedule:Task Constraints...]
		 If Not objDialogTaskConstraints.Exist(SISW_MIN_TIMEOUT) Then
					bReturn = Fn_MenuOperation("Select","Schedule:Task constraints")
					If bReturn = False Then
							Fn_SchMgr_TaskConstraint = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:Task constraints]")
							Set objDialogTaskConstraints = Nothing
							Exit Function
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:Task constraints]")
					End If
		 End If

        'Check the existence of the "Task Constraints" Window
				If objDialogTaskConstraints.Exist(SISW_MAX_TIMEOUT)  Then
					objDialogTaskConstraints.JavaRadioButton("ConstraintType").SetTOProperty "attached text",sConstraint
					objDialogTaskConstraints.JavaRadioButton("ConstraintType").Set "ON"
					If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TaskConstraint: Failed to Set Task Constraint as [" + sConstraint + "]")
							Fn_SchMgr_TaskConstraint = FALSE
							Set objDialogTaskConstraints = Nothing
							Exit Function
					End If
					objDialogTaskConstraints.JavaButton("OK").Click

						If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TaskConstraint: Failed to click on OK button ")
								Fn_SchMgr_TaskConstraint = FALSE
								Set objDialogTaskConstraints = Nothing
								Exit Function
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TaskConstraint: Successfully set Task Constraints [" + sConstraint + "]")	
						Fn_SchMgr_TaskConstraint = True
				Else
							Fn_SchMgr_TaskConstraint = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_TaskConstraint:Task Constraints Dialog  does Not Exist ")
				End If

				Set objDialogTaskConstraints = Nothing

End Function


'************************************************* ******* **************************Function to  Recalculate Schedule*************************************************************************************************************************************************
'Function Name		         :			       Fn_SchMgr_ScheduleRecalculate

'Description			         :		 	       Function recalculate the schedule.

'Parameters			         :	 		       NA

'Return Value		       : 		        TRUE \ FALSE

'Pre-requisite			   :		 	   NA

'Examples				  :		         Call  Fn_SchMgr_ScheduleRecalculate()

'History				                     Developer Name					               Date					               Rev. No.			                  Changes Done						           Reviewer
'								              ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									               Manish Verma				                       20/05/2010			              1.0										      Created
'													Vallari S.												24-May-2010						1.0										Report comments, Error Handling, Release reference
'								            --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ScheduleRecalculate()	
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ScheduleRecalculate"
	Dim objRecalSchedule, bReturn

	On Error Resume Next 

		Set objRecalSchedule=JavaDialog("Recalculate Schedule")	
		If objRecalSchedule.Exist(SISW_MIN_TIMEOUT)=False Then
				Set objRecalSchedule=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Recalculate Schedule")							
                 If objRecalSchedule.Exist(SISW_MICRO_TIMEOUT)=False Then
			           'Select menu  [Schedule:Recalculate schedule]               
						bReturn = Fn_MenuOperation("Select","Schedule:Recalculate schedule")
						If bReturn = False Then
							Fn_SchMgr_ScheduleRecalculate = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:Recalculate schedule]")
							Set objRecalSchedule = Nothing
							Exit Function
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:Recalculate schedule]")
					End If
				End If
			End If

				If objRecalSchedule.Exist (SISW_DEFAULT_TIMEOUT *2 )  Then
				    'Click the yes button
				    objRecalSchedule.JavaButton("Yes").WaitProperty "enabled",1,20000
				    objRecalSchedule.JavaButton("Yes").Click							
					If Err.Number < 0 Then
							Fn_SchMgr_ScheduleRecalculate = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Yes] Button")
							objRecalSchedule = Nothing
							Exit Function
					End If
							Fn_SchMgr_ScheduleRecalculate =True																	                                                       
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Recalculate Schedule") 
    	       Else 
							Fn_SchMgr_ScheduleRecalculate = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Recalculate Schedule] Dialog Does Not Exist") 
              End If
			  objRecalSchedule = Nothing
End Function	


'*********************************************************	**************************	Function to Delete the Object from Schedule Manager	*	****************************************************************************
'Function Name		:				Fn_SchMgr_NodeRename.

'Description			  :		 		This function is used to delete the Task  from Schedule Manager.

'Parameters			    :	 		1.  sAction :Action need to perform
'												 2.sNodeName: Fully qulified : separated path of SchTable
'                                               3.sNewName :New name

'Return Value		    : 		 True/False

'Pre-requisite			 :		 	Select the task/milestone from schedule table.

'Examples				 :		   call Fn_SchMgr_TaskDelete("Menu","sch:task1",newtask")
':
'History					 :		   Developer Name			             Date			                 Rev. No.			          Changes Done			  Reviewer	                   Reviewed date
'                                      -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										       Manish Verma                          21/05/10							1.0															
'                                     ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_NodeRename(sAction,sNodeName,sNewName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_NodeRename"
   Dim bReturn
   Dim objRename

   bReturn= Fn_SchMgr_SchTable_NodeOperation ("Select", sNodeName,"","","")
	If bReturn=True Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected SchTable Node [" + sNodeName + "]")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select SchTable Node [" + sNodeName + "]")
		Fn_SchMgr_NodeRename   = False
		Exit Function
	End If	

   'Set the object name
	'Set objRename=JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Rename")
		Set objRename=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Rename")

	Select Case sAction
		Case "Menu"  
			bReturn = Fn_MenuOperation("Select","Edit:Rename")
			If bReturn = True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Edit;Rename]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke Menu [Edit;Rename]")
				Fn_SchMgr_NodeRename = False
				Set objRename = Nothing
				Exit Function
			End If

	 Case "PopupMenu"		
			bReturn= Fn_SchMgr_SchTable_NodeOperation( "PopupMenu", sNodeName,"","","Edit:Rename" )
			If bReturn=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Popupmenu [Edit:Rename] on SchTable Node[" + sNodeName + "]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Popupmenu [Edit:Rename] on SchTable Node[" + sNodeName + "]")
				Fn_SchMgr_NodeRename   = False
				Set objRename = Nothing
				Exit Function
			End If		
	End Select	

          If objRename.Exist(SISW_DEFAULT_TIMEOUT *2 )   Then 
				objRename.JavaEdit("NewName").Set sNewName
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set New Name as [" + sNewName + "]")
					 Fn_SchMgr_NodeRename = FALSE
					 Set objRename = Nothing
					Exit Function
				End If 
				objRename.JavaButton("OK").Click
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [OK] Button")
					 Fn_SchMgr_NodeRename = FALSE
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Renamed Node to [" + sNewName + "]")	
				 Fn_SchMgr_NodeRename = True
	    Else
				Fn_SchMgr_NodeRename = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Rename] Dialog Not Found")
		End If
		Set objRename = Nothing
End Function


'*********************************************************		Function to  get  Resource table Row Index	***********************************************************************

'Function Name		:					Fn_SchMgr_TableRowIndex

'Description			 :		 		  This function is used to get  Resource table Row Index.

'Parameters			   :	 			1.  sResource:Name of the Resource to retrieve Index for.
'													2.sColName :column name
											
'Return Value		   : 				 Resource index

'Pre-requisite			:		 		Schedule membership dialog should be displayed .

'Examples				:				 Fn_SchMgr_TableRowIndex("Organization:User:Rupali Palhade (x_palhad)")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							24-May-2010	   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_TableRowIndex(objSchTable, sResource,sColName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TableRowIndex"
	Dim IntRows ,sNodePath, IntCounter, ObjTable, StrIndex, ArrNode

	On Error Resume Next

	'Verify that Resource Table is displayed
	If objSchTable.Exist(SISW_MICRO_TIMEOUT) Then

		'Get the No. of rows present in the ResourceTable
		IntRows = objSchTable.GetROProperty("rows")
		
		'Get the Row No. of required Resource
		For IntCounter = 0 to IntRows -1
                sNodePath = objSchTable.GetCellData(IntCounter,sColName)
				If Trim(sNodePath) = Trim(sResource) Then
				    StrIndex = Cstr(IntCounter)
					Fn_SchMgr_TableRowIndex = StrIndex
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ": Row Index of [" + sResource +"] resource is [" + IntCounter + "]")	
					Exit For
				End If
		 Next

		If  cstr(IntCounter) = IntRows Then
			Fn_SchMgr_TableRowIndex = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ":Failed to Get  Row Index of [" + sResource +"]")	
		End If
  End If
End Function

' *********************************************************		Function to  assign , modify,verify,Remove  member of  Schedule	***********************************************************************

'Function Name		:					Fn_SchMgr_ScheduleMembership 

'Description			 :		 		  This function is used  to add ,remove  memeber  to schedule and modify , verify the property of schedule.

'Parameters			   :	 			1.sAction : Action need to perform.
'                                                    2.aUserName : Array of User name which need to add as member of schedule ( This will be : seprated  e.g. Organization:Group:Engineering)
'                                                  3.aPrivilageval : Array of Member privilage (maintain the sequence in which users are passed in)
'                                                 4.aRateval :  Array of Rate (maintain the sequence in which users are passed in)
'                                                5.sCurrencyval : Currency (maintain the sequence in which users are passed in)

											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Select the schedule from schedule table 

'Examples				:				 sPartUser1=Split("AutoTest2:AutoTest2:Engineering:Designer::autotest2",":",-1,1)
'										aUserName =Array("Organization:User:"+sPartUser1(0)+" ("+sPartUser1(5)+")")
'										bReturn=Fn_SchMgr_ScheduleMembership("SetCalendar",aUserName ,"" ,"" ,"" )

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							21-May-2010	   1.0
'										SHREYAS							30-MAY-2011	   1.1				Added Case "SetCalendar"  SHREYAS
'										Pritam S.						27-Dec-2011	   1.2				Added Case "CalendarExist"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ScheduleMembership(sAction,aUserName ,aPrivilageval ,aRateval ,aCurrencyval )
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ScheduleMembership"
	On Error Resume Next
	Dim objTree,iCounter,sRow, iItemCount,sNodename,iIndex,ArrNodes,objTable,bReturn, sTitle
	
	If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").Exist(SISW_MIN_TIMEOUT) = False Then
		bReturn =  Fn_MenuOperation("Select","Schedule:Schedule Membership")
		wait 7
		If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").Exist(SISW_MICRO_TIMEOUT) = False Then
				Fn_SchMgr_ScheduleMembership = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule Membership dialog is not displayed.")
				Exit Function
		 End If
	End If

	If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").Exist(SISW_MICRO_TIMEOUT) Then
		'Create object of  Tree
		set objTree =  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaTree("ScheduleMemberTree")
		Set objTable =  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaTable("ResourceTable")

		Select Case sAction
			Case   "Assign"

				For iCounter =0 to Ubound(aUserName)
					ArrNodes = split(aUserName(iCounter), ":",-1,1)
					sRow =  Fn_SchMgr_TableRowIndex(objTable, ArrNodes(Ubound(ArrNodes)),"#0")
					'Get item count  of java tree.
					If sRow = FALSE Then
						'Expand Organization:User / Organization:Group/ Organization:Discipline Node
'						iItemCount = objTree.GetROProperty("items count")
						iItemCount= Fn_UI_Object_GetROProperty("Fn_SchMgr_ScheduleMembership",objTree, "items count")
						If iItemCount > 0 Then
							For iIndex = 0 to iItemCount -1
								sNodename = objTree.Object.getPathForRow(iIndex).tostring
								sNodename = Replace(sNodename,", ",":",1,-1,1)   
								sNodename = Mid (sNodename,2,Len(sNodename)-2)
								If Trim(sNodename) =  ArrNodes(0) & ":" & ArrNodes(1)  Then 
									objTree.expand objTree.GetItem(iIndex) 
									Fn_SchMgr_ScheduleMembership = TRUE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded  [Organization:User] node" )
									Exit For
								End If
							Next
						End If
						iItemCount = objTree.GetROProperty("items count")
						If iItemCount > 0 Then
							For iIndex = 0 to iItemCount -1
								sNodename = objTree.Object.getPathForRow(iIndex).tostring
								sNodename = Replace(sNodename,", ",":",1,-1,1)   
								sNodename = Mid (sNodename,2,Len(sNodename)-2)
								If Trim(sNodename) = Trim(aUserName(iCounter)) Then
									objTree.Activate objTree.GetItem(iIndex) 
									Fn_SchMgr_ScheduleMembership = TRUE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully assign  schedule memebership to  "  & aUserName(iCounter) )
									Exit For
								End If
							Next
						End If
	
						If  cint(iIndex) = cint(iItemCount)Then
							Fn_SchMgr_ScheduleMembership = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to assign  schedule memebership to  "  & aUserName(iCounter) )
							 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
							Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "Cancel")
							set objTree = Nothing
							set objTable = Nothing
							Exit Function
						End If
					Else
						Fn_SchMgr_ScheduleMembership = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  aUserName(iCounter) & "  is already assign to schedule membership." )
					End If
				Next
				If aPrivilageval <> Empty Then
						bReturn = Fn_SchMgr_ScheduleMembership("Modify",aUserName ,aPrivilageval  ,"" ,"" )
						If bReturn = False Then
								Fn_SchMgr_ScheduleMembership = FALSE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Roles could not be assigned")
								Exit Function
						Else
								Fn_SchMgr_ScheduleMembership = TRUE
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully assigned roles through Schedule Membership")
						End If
				Else
						JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("OK").WaitProperty "enabled","1",20000
						wait 2 
						 Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "OK")
				End If				 
				
				If Err.Number < 0 Then
					Fn_SchMgr_ScheduleMembership = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to assign  schedule memebership to  "  & aUserName(iCounter) )
					 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
					wait 2 
					Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "Cancel")
					set objTree = Nothing
					set objTable = Nothing
					Exit Function
				End If

			Case  "Modify"
				For iCounter =0 to Ubound(aUserName)
					ArrNodes = split(aUserName(iCounter), ":",-1,1)
					sRow =  Fn_SchMgr_TableRowIndex(objTable, ArrNodes(Ubound(ArrNodes)),"#0")
					If sRow <> FALSE Then

                        'Modify the value of  Member Privileges column
						If IsArray(aPrivilageval) = True Then
							If aPrivilageval(iCounter) <> " " Then
								sNodename = objTable.GetCellData(sRow,1)
								If  sNodename <> "NA"Then
									    objTable.ClickCell sRow,1
										' JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaList("ResourceTableCombo").Select aPrivilageval(iCounter)
										bReturn = Fn_List_Select("", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"),"ResourceTableCombo",aPrivilageval(iCounter))
									    If Err.Number <  0 Then
											Fn_SchMgr_ScheduleMembership = FALSE				 
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  : Fail to edit Cell Value [" + aPrivilageval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Member Privileges")	
											 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
											Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "Cancel")
											set objTree = Nothing
											set objTable = Nothing
											Exit Function
									    Else
											Fn_SchMgr_ScheduleMembership = TRUE				 
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ":Successfully edit Cell Value [" + aPrivilageval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Member Privileges")	
										End If
								Else 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Can not modify the value Member Privileges for  " & aUserName(iCounter) )
								End If
							End If
						End If  

						 'Modify the value of  Rate column
						If IsArray(aRateval) = True Then
							If aRateval(iCounter) <> " " Then
								sNodename = objTable.GetCellData(sRow, "Rate")
								If  sNodename <> "NA"Then
									objTable.SetCellData sRow, 2,aRateval(iCounter)
									objTable.Click "0","0"
									If Err.Number  <  0 Then
										Fn_SchMgr_ScheduleMembership = FALSE				 
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  : Fail to edit Cell Value [" + aRateval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Rate.")	
										 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
										Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "Cancel")
										set objTree = Nothing
										set objTable = Nothing
										Exit Function
									 Else
										Fn_SchMgr_ScheduleMembership = TRUE				 
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ": Successfully edit Cell Value [" + aRateval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Rate.")	
									End If
								Else 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Can not modify the value Rate  for  " & aUserName(iCounter) )
								End If
							End If
						End If

						'Modify the value of  Currency column
						If IsArray(aCurrencyval) = True Then
							If aCurrencyval(iCounter) <> " " Then
								sNodename = objTable.GetCellData(sRow, "Currency")
								If  sNodename <> "NA"Then
									    objTable.ClickCell sRow,3
										 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaList("ResourceTableCombo").Select aCurrencyval(iCounter)
									    If Err.Number <  0 Then
											Fn_SchMgr_ScheduleMembership = FALSE				 
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  : Fail to edit Cell Value [" + aCurrencyval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and ColumnCurrency")	
											 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
											Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "Cancel")
											set objTree = Nothing
											set objTable = Nothing
											Exit Function
									    Else
											Fn_SchMgr_ScheduleMembership = TRUE				 
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ": Successfully edit Cell Value [" + aCurrencyval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Currency")	
										End If
								Else 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Can not modify the value Currency  for  " & aUserName(iCounter) )
								End If
							End If
						End If  

					Else 
						Fn_SchMgr_ScheduleMembership = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Resource is not exist in resource table." )
						 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
						Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "Cancel")
						set objTree = Nothing
						set objTable = Nothing
						Exit Function
					End If
				Next 
				 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("OK").WaitProperty "enabled","1",20000
				 wait 5
				 Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "OK")
				If Err.Number < 0 Then
					Fn_SchMgr_ScheduleMembership = FALSE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  : Fail to edit Cell Value for user " +  aUserName(iCounter))	
					 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
					Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_ScheduleMembership", "DeviceReplay.Click", JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership"), "Cancel")
					set objTree = Nothing
					set objTable = Nothing
					Exit Function
				End If

			Case "Verify"
				For iCounter =0 to Ubound(aUserName)
					ArrNodes = split(aUserName(iCounter), ":",-1,1)
					sRow =  Fn_SchMgr_TableRowIndex(objTable, ArrNodes(Ubound(ArrNodes)),"#0")
					If sRow <> FALSE Then

                        'Verify the value of  Member Privileges column
						If IsArray(aPrivilageval) = True Then
							If aPrivilageval(iCounter) <> " " Then
								sNodename = objTable.GetCellData(sRow,1)
								If  Trim(sNodename) = Trim(aPrivilageval(iCounter))Then
									Fn_SchMgr_ScheduleMembership = TRUE				 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ": Successfully verify Cell Value [" + aPrivilageval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Member Privileges")		
								Else
									Fn_SchMgr_ScheduleMembership = FALSE				 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  : Fail to verify Cell Value [" + aPrivilageval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Member Privileges")	
									 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
							         JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
									set objTree = Nothing
									set objTable = Nothing
									Exit Function
								End If
							End If
						End If  

						 'Verify the value of  Rate column
						If IsArray(aRateval) = True Then
							If aRateval(iCounter) <> " " Then
								sNodename = objTable.GetCellData(sRow, "Rate")
									If Trim(sNodename) = Trim(aRateval(iCounter))Then
										Fn_SchMgr_ScheduleMembership = TRUE				 
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ": Successfully verify Cell Value [" + aRateval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Rate.")	
									 Else
										Fn_SchMgr_ScheduleMembership = FALSE				 
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  : Fail to verify Cell Value [" + aRateval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Rate.")	
										 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
										 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
										set objTree = Nothing
										set objTable = Nothing
										Exit Function
									End If
							End If
						End If

						'Verify the value of  Currency column
						If IsArray(aCurrencyval) = True Then
							If aCurrencyval(iCounter) <> " " Then
								sNodename = objTable.GetCellData(sRow, "Currency")
								If Trim( sNodename) = Trim(aCurrencyval(iCounter))Then
									Fn_SchMgr_ScheduleMembership = TRUE				 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "  : Successfully verify Cell Value [" + aCurrencyval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and ColumnCurrency")	
								 Else
									Fn_SchMgr_ScheduleMembership = FALSE				 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ": Fail to verify Cell Value [" + aCurrencyval(iCounter) + "] for  [" + aUserName(iCounter)  + "] and Column Currency")	
									 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
									 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
									set objTree = Nothing
									set objTable = Nothing
									Exit Function
								End If
							End If
						End If  

					Else 
						Fn_SchMgr_ScheduleMembership = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Resource is not exist in resource table." )
'						- Added by Vandana [ 16-Feb-2012 Build : 2012020800 ]
						 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
						 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
'						 -------------------------
						Exit Function
					End If
				Next 
				 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
				
			Case "Remove"
				'Dim WshShell,ArrNodes
               ' set WshShell = CreateObject("WScript.Shell")
				For iCounter =0 to Ubound(aUserName)
					ArrNodes = split(aUserName(iCounter), ":",-1,1)
					sRow =  Fn_SchMgr_TableRowIndex(objTable, ArrNodes(Ubound(ArrNodes)),"#0")
					If sRow <> FALSE Then
						'Expand Organization:User / Organization:Group/ Organization:Discipline Node
						'iItemCount = objTree.GetROProperty("items count")
						iItemCount=Fn_UI_Object_GetROProperty("", objTree,"items count")
						If iItemCount > 0 Then
							For iIndex = 0 to iItemCount -1
								sNodename = objTree.Object.getPathForRow(iIndex).tostring
								sNodename = Replace(sNodename,", ",":",1,-1,1)   
								sNodename = Mid (sNodename,2,Len(sNodename)-2)
								If Trim(sNodename) =  ArrNodes(0) & ":" & ArrNodes(1)  Then 
									objTree.expand objTree.GetItem(iIndex) 
									Fn_SchMgr_ScheduleMembership = TRUE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded  [Organization:User] node" )
									Exit For
								End If
							Next
						End If
						'Get item count  of java tree.
						'iItemCount = objTree.GetROProperty("items count")
						iItemCount=Fn_UI_Object_GetROProperty("", objTree,"items count")
						If iItemCount > 0 Then
							For iIndex = 0 to iItemCount -1
								sNodename = objTree.Object.getPathForRow(iIndex).tostring
								sNodename = Replace(sNodename,", ",":",1,-1,1)   
								sNodename = Mid (sNodename,2,Len(sNodename)-2)
								If Trim(sNodename) = Trim(aUserName(iCounter)) Then
									objTree.select objTree.GetItem(iIndex)   

									'Handle Error if Exists
									sTitle = "Error"
									sErrorText = "Schedule owner can not be unassigned."
									bReturn = Fn_SchMgr_DialogMsgVerify(sTitle,sErrorText,"OK")
									If bReturn Then
										Fn_SchMgr_ScheduleMembership = FALSE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to unassign  schedule memebership to  "  & aUserName(iCounter) )
										 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
										 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
										set objTree = Nothing
										set objTable = Nothing
										sErrorText = ""
										Exit Function
									End If

									If Instr(1,aUserName(iCounter),"Group",1) > 0 Then
										 If JavaDialog("Remove Schedule Group.").Exist(5) Then
											JavaDialog("Remove Schedule Group.").JavaButton("Yes").Click micLeftBtn
										 End If
									 End If

									Fn_SchMgr_ScheduleMembership = TRUE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully unassign  schedule memebership to  "  & aUserName(iCounter) )
									Exit For
								End If
							Next
						End If
	
						If  cint(iIndex) = cint(iItemCount)Then
							Fn_SchMgr_ScheduleMembership = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to unassign  schedule memebership to  "  & aUserName(iCounter) )
							 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
							 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
							set objTree = Nothing
							set objTable = Nothing
							Exit Function
						End If
					Else 
						Fn_SchMgr_ScheduleMembership = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to unassign  " & aUserName(iCounter) & "because this user is not assign schedule membership." )
						 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
						 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
						set objTree = Nothing
						set objTable = Nothing
						Exit Function
					End If
				Next
				 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("OK").WaitProperty "enabled","1",20000
				 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("OK").Click micLeftBtn

				If Err.Number < 0Then
					Fn_SchMgr_ScheduleMembership = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to unassign user  " +  aUserName(iCounter) )
					 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").WaitProperty "enabled",1,5000
					 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaButton("Cancel").Click micLeftBtn
					set objTree = Nothing
					set objTable = Nothing
					Exit Function
				End If

		Case "SetCalendar"

				For iCounter =0 to Ubound(aUserName)
					ArrNodes = split(aUserName(iCounter), ":",-1,1)
					sRow =  Fn_SchMgr_TableRowIndex(objTable, ArrNodes(Ubound(ArrNodes)),"#0")
					If sRow <> FALSE Then
						'Get item count  of java tree.
'						iItemCount = objTree.GetROProperty("items count")
'						If iItemCount > 0 Then
'							For iIndex = 0 to iItemCount -1
'								sNodename = objTree.Object.getPathForRow(iIndex).tostring
'								sNodename = Replace(sNodename,", ",":",1,-1,1)   
'								sNodename = Mid (sNodename,2,Len(sNodename)-2)
'								If Trim(sNodename) = Trim(aUserName(iCounter)) Then
									 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaTable("ResourceTable").ClickCell sRow,4
									'click on yes button of  Create Schedule Member dialog
									JavaDialog("Create Schedule Member").JavaButton("Yes").Click micLeftBtn
									Fn_SchMgr_ScheduleMembership = true
'									Exit For
'								End if
'							Next
'						End if
					Exit For
				End if
			Next
			
		Case "CalendarExist"
				''''Added By Pritam S. to verify wether the Cslendar is created for Schedule member or Not
						For iCounter =0 to Ubound(aUserName)
							ArrNodes = split(aUserName(iCounter), ":",-1,1)
							sRow =  Fn_SchMgr_TableRowIndex(objTable, ArrNodes(Ubound(ArrNodes)),"#0")
							If sRow <> FALSE Then
									'Get item count  of java tree.
'									iItemCount = objTree.GetROProperty("items count")
'									If iItemCount > 0 Then
'											For iIndex = 0 to iItemCount -1
'												sNodename = objTree.Object.getPathForRow(iIndex).tostring
'												sNodename = Replace(sNodename,", ",":",1,-1,1)   
'												sNodename = Mid (sNodename,2,Len(sNodename)-2)
'												If Trim(sNodename) = Trim(aUserName(iCounter)) Then
													     JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Membership").JavaTable("ResourceTable").ClickCell sRow,4
														If  JavaDialog("Create Schedule Member").Exist Then
															JavaDialog("Create Schedule Member").JavaButton("No").Click micLeftBtn
															 Fn_SchMgr_ScheduleMembership = False
															 Exit Function 
														 Else
															Fn_SchMgr_ScheduleMembership = True
															Exit Function
														End If
'												End if
'											Next
'								End if
								Exit For
						End if
					Next

		End Select
		set objTree = Nothing
		set objTable = Nothing
	Else
		Fn_SchMgr_ScheduleMembership = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule Membership dialog does not exist.")
		Exit Function
	End If 
End Function

' *********************************************************		Function to  assign , modify,verify,Remove  member of  Schedule	***********************************************************************

'Function Name		:					Fn_SchMgr_PercentLinkedMsgVerify 

'Description			 :		 		  This function verifies the message if SchProperty Percent Linked is modified.

'Parameters			   :	 			1.sMessage:  This is an Optional parameter. Pass in only if you want to very the message
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Select the schedule from schedule table 

'Examples				:				 Fn_SchMgr_PercentLinkedMsgVerify("")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari							27-May-2010	   1.0
'										Sushma						24-Jun-2013	   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_PercentLinkedMsgVerify(sMessage)

   		Dim dicErrorInfo, bReturn 
		Set dicErrorInfo = CreateObject("Scripting.Dictionary")
		With dicErrorInfo 		 
		 .Add "Message", sMessage		 
		 .Add "Action", "PercentLinkedMsgVerify"		 
		End with
		Fn_SchMgr_PercentLinkedMsgVerify = Fn_SISW_SchMgr_ErrorVerify(dicErrorInfo)

End Function

' *********************************************************		Function to  assign , modify,verify,Remove  member of  Schedule	***********************************************************************

'Function Name		:					Fn_SchMgr_ScheduleShift 

'Description			 :		 		  Sets Schedule Shit date

'Parameters			   :	 			1.sDate:  Date in d-mmm-yyyy format
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Select the schedule from schedule table 

'Examples				:				 Fn_SchMgr_ScheduleShift("6-Jul-2012")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari							27-May-2010	   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ScheduleShift(sDate)
		GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ScheduleShift"
		Dim WshShell
		Dim objButton, objShiftWin
		Dim sActDate, aDate, aActDate,sCurrentValue,sNewItemMenu
		Dim iMonDiff, iYrDiff, iDiff,bReturn,sDiff
		Dim aBtn, sBtnName, iCnt

		If instr(sDate, "~") > 0 Then
			aBtn = split(sDate, "~", -1, 1)
			sBtnName = aBtn(1)
		Else
			sBtnName = "Yes"
		End If

		On Error Resume Next

		Set objShiftWin = Fn_SISW_PPM_GetObject("Shift Schedule")
		sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"ScheduleShiftSchedule")
		If objShiftWin.Exist (SISW_MIN_TIMEOUT) = False  Then
		   ' Select menu  [File -> New -> Task...]               
			bReturn = Fn_MenuOperation("Select",sNewItemMenu)	  	
			If bReturn = False Then
					Fn_SchMgr_ScheduleShift = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:Shift Schedule]")
					Set objShiftWin = Nothing
					Exit Function
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:Shift Schedule]")
			End If
		End If
		
		Wait 1	
		JavaDialog("Confirmation").SetTOProperty "title", "Confirmation"
		If JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
			JavaDialog("Confirmation").JavaButton(sBtnName).Click micLeftBtn
		End If
		
		If objShiftWin.Exist (SISW_DEFAULT_TIMEOUT) Then

				'Set the value of  Shift Date
				If sDate <> "" Then
				
				
				
					sCurrentValue = objShiftWin.JavaEdit("ShiftDate").GetROProperty("value")
					sDiff = DateDiff("d",Cdate(sCurrentValue),Cdate(sDate))
					iMonDiff = DateDiff("m",Cdate(sCurrentValue),Cdate(sDate))
					Set WshShell = CreateObject("WScript.Shell")	
					'[TC1121-2015102600-20_11_2015-VivekA-Maintenance] - As discussed with Vallari added condition sDiff > 0
					If sDiff < 28 and sDiff > 0 and iMonDiff = 0 Then	' Modified by Vallari [TC112-2015071500-27_07_2015-VivekA-Porting]
						aDate = Split(sDate,"-")
						objShiftWin.JavaEdit("ShiftDate").DblClick 0, 0, "LEFT"
						wait 5
						objShiftWin.JavaEdit("ShiftDate").Type aDate(0)
						wait 1
						objShiftWin.JavaEdit("ShiftDate").Activate
						wait 1
						objShiftWin.JavaEdit("ShiftDate").RefreshObject
						'WshShell.SendKeys "{ESC}"
						WshShell.SendKeys "{TAB}"
					Else
						
						wait 1
				     	objShiftWin.JavaEdit("ShiftDate").Click 0, 0, "LEFT"
						wait 5
						objShiftWin.JavaEdit("ShiftDate").Set sDate
						wait 5
						objShiftWin.JavaEdit("ShiftDate").Activate
						wait 1
						objShiftWin.JavaEdit("ShiftDate").RefreshObject
						'WshShell.SendKeys "{ESC}"
					     WshShell.SendKeys "{TAB}"
					End If
					Set WshShell = Nothing
					wait 2	
				
				End If
'				'Split Date to be Set
'				aDate = Split(sDate, "-", -1,1)
'				'Extract Actual Date and Split it
'				sActDate = objShiftWin.JavaCheckBox("ShiftDate").GetROProperty("label")
'				aActDate = Split(sActDate, "-", -1,1)
'				'Calculate Date Differences
'				iMonDiff = DateDiff("m", sActDate, sDate)
'				iYrDiff = DateDiff("yyyy", sActDate, sDate)
'				iDiff = iMonDiff - (iYrDiff * 12)
'				'Decide Scroll Direction
'				If iDiff > 0 Then
'					set objButton = objShiftWin.JavaButton("ScrollRight")
'				Elseif iDiff < 0 Then
'					set objButton = objShiftWin.JavaButton("ScrollLeft")
'				End If
'				'Set Year
'				objShiftWin.JavaCheckBox("ShiftDate").Click 5,5,"LEFT"
'				objShiftWin.JavaEdit("Year").Set aDate(2)
'				objShiftWin.JavaEdit("Year").Activate
'				'Scroll to Get Proper Month
'				For iCnt = 1 to abs(iDiff)
'					objButton.Click micLeftBtn
'				Next
'				
'				'Set Required Date Digit
'				objShiftWin.JavaCheckBox("DateDigit").SetTOProperty "attached text", cstr(aDate(0))
'				objShiftWin.JavaCheckBox("DateDigit").Click 2,2,"LEFT"
'
'				objShiftWin.JavaButton("DateOK").Click micLeftBtn
'				wait(1)
'				JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Shift Schedule").Activate
'				JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Shift Schedule").JavaCheckBox("ShiftDate").Click 5,5,"LEFT"
'				JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Shift Schedule").JavaCheckBox("ShiftDate").DblClick 5,5,"LEFT"
				objShiftWin.JavaButton("OK").WaitProperty "enabled", 1, 20000
				objShiftWin.JavaButton("OK").Click micLeftBtn
				wait 2
				bReturn =  Fn_SchMgr_DialogMsgVerify("Scheduling Error","","OK") 
				If  bReturn Then
					Fn_SchMgr_ScheduleShift = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully handle scheduling error.")
					Exit Function
				End If

				If err.number < 0 Then
						Fn_SchMgr_ScheduleShift = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Shift Schedule to Date [" + sDate + "]")
				Else
						Fn_SchMgr_ScheduleShift = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Shifted Schedule to Date [" + sDate + "]")
				End If
		Else
				Fn_SchMgr_ScheduleShift = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Shift Schedule] Dialog Not Found")
		End If

		Set objShiftWin = Nothing
		Set objButton = Nothing
End Function


'*********************************************************		Function to  modify , Verify ,  to check IsEditable  Schedule  property.***********************************************************************

'Function Name		:					Fn_SchMgr_SchPropertyOperations

'Description			 :		 		  This function is used to  modify , Verify ,  to check IsEditable  Schedule  properties.

'Parameters			   :	 			sAction: Action need to perform.
'                                                   dicProperties: Dictionary object to hold property names and the values
'                                                   sButtonName: Button to be clicked after modifying properties
'													sConfirmationBtn : Name of button of confirmation dialog.
											
'Return Value		   : 			  True/False  

'Pre-requisite			:		 		Schedule manger panel need to be open.
'                                                   Scedule need to be selected.
'                                                   Properties dialog need to be displayed.

'Examples				:			Call Fn_SchMgr_SchPropertyOperations("Modify",dicScheduleProperty,"OK","YES")
'												Call Fn_SchMgr_SchPropertyOperations("IsEditable",dicScheduleProperty,"","")
'									dicScheduleProperty("ScheduleDeliverable") =True
'									bReturn=Fn_SchMgr_SchPropertyOperations("SchDeliverable",dicScheduleProperty,"","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali					25-May-2010	   			1.0
'										SHREYAS					11-05-2011				1.1				Added Case SchDeliverable SHREYAS
'										Sachin Joshi			28-June-2011            1.1             Modified Case "StatusDropDown"
'										Pritam					16-Jan-2012             1.2             Added case "State" in Case "Verify"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_SchPropertyOperations(sAction,dicProperties,sButtonName,sConfirmationBtn)
GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_SchPropertyOperations"
   On Error Resume Next
   Dim dicCount , dicKeys , dicItems , iCounter,objWin ,bReturn, sDate, aDate,sGetText
   Dim objCheckBox,sObjects,sValue,WshShell,arrDate, sMsg
   Dim objError, intNoOfObjects, iCount, sText

	sMsg = ""
   Set  objWin =  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties")

	If  not objWin.Exist(SISW_MIN_TIMEOUT) Then
		bReturn = Fn_MenuOperation("Select","View:View Properties")
		If bReturn = False Then
				Fn_SchMgr_SchPropertyOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [View:Properties]")
				objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
				objWin.JavaButton("Cancel").Click
				Set objWin = Nothing
				Exit Function
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [View:Properties]")
		End If
	End If
     
	 objWin.WaitProperty "displayed",1,20000
	 If  objWin.Exist(SISW_MIN_TIMEOUT) Then
		 If objWin.JavaRadioButton("ViewScheduleProperties").Exist(SISW_MICRO_TIMEOUT) Then
				objWin.JavaRadioButton("ViewScheduleProperties").Set "ON"
				objWin.JavaButton("OK").WaitProperty "enabled",1,20000
				objWin.JavaButton("OK").Click
		 End If

		Call Fn_ReadyStatusSync(2)  

         dicCount  = dicProperties.Count
		 dicItems = dicProperties.Items
		 dicKeys = dicProperties.Keys

		Select Case  sAction
			Case "Modify","ModifyWithoutMsgVerification"
				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "CustomerName" ,"CustomerNumber"
								objWin.JavaEdit(dicKeys(iCounter)).Set dicItems(iCounter)
								
							Case "StartDate" ,"FinishDate" 
                                   	arrDate = Split(dicItems(iCounter) ," ")
                                   	objWin.JavaEdit(dicKeys(iCounter)).click 1,1
                                   	wait 7
									objWin.JavaEdit(dicKeys(iCounter)).Set arrDate(0)
									wait 1
									objWin.JavaEdit(dicKeys(iCounter)).Activate
									wait 1
									objWin.JavaEdit(dicKeys(iCounter)).RefreshObject
									'[TC1121-2015102600-10_11_2015-VivekA-Maintenance] - Added for Error verification while putting incorrect Date in Date field
									Set objError = Description.Create
									objError("Class Name").value = "JavaStaticText"
									objError("text").value = "Invalid input.*"
									Set  intNoOfObjects =objWin.ChildObjects(objError)
									For iCount = 0 To intNoOfObjects.count-1
										sText = intNoOfObjects(iCount).GetROProperty("text")
										If Instr(sText,"Invalid input.")>0 Then
											Fn_SchMgr_SchPropertyOperations = FALSE
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set Correct Date in "+dicKeys(iCounter))
											objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
											objWin.JavaButton("Cancel").Click
											Exit Function
										End If
									Next
									Wait 2
									'-------------------------------------------------------------------------------------------------------
                                    Set WshShell = CreateObject("WScript.Shell")
									WshShell.SendKeys "{TAB}"
									Set WshShell = Nothing
									wait 2
									if Ubound(arrDate) = 1 Then
										objWin.JavaList(dicKeys(iCounter)).Select arrDate(1)									
									End If


							Case "StatusDropDown" ,"TskStatusDrpDwn","PriorityDropDown"
								'Code Commented as drop down not works here
								'objWin.JavaButton(dicKeys(iCounter)).Click
								'Wait(2)
								'bReturn =  Fn_iComboSet(objWin,dicItems(iCounter))
								'Added By Sachin
								If dicKeys(iCounter) ="StatusDropDown" OR dicKeys(iCounter) ="TskStatusDrpDwn"  Then
									objWin.JavaEdit("SchStatus").Set dicItems(iCounter)
									if err.number < 0 then
										bReturn = false
									else
										bReturn = true	
									end if
									objWin.JavaEdit("SchStatus").Activate
								Else
									objWin.JavaEdit("SchPriority").Set dicItems(iCounter)
									if err.number < 0 then
										bReturn = false
									else
										bReturn = true	
									end if
									objWin.JavaEdit("SchPriority").Activate
							End If
							'Modified by Omkar K... to Handle Change in Label of the Radio button - "Published"		
							Case  "IsScheduleTemplate" ,"IsPublic","FinishDateSchedul", "IsPercentLinked","Published","NotificationsEnabled","AreDatesLinked","ExcnOverrideEnabled","IsTemplate"
	'									If dicKeys(iCounter) = "Published" Then
	'											sGetText =Cstr((dicItems(iCounter)))
	'											If  sGetText="True"	Then
	'												objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label","true"
	'												If objWin.JavaRadioButton(dicKeys(iCounter)).Exist(5)=False Then
	'														objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label","True"
	'												End If
	'											ElseIf  sGetText="False"	Then
	'													objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label","false"
	'														If objWin.JavaRadioButton(dicKeys(iCounter)).Exist(5)=False Then
	'															objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label","False"
	'														End If
	'											End If									
	'									Else
	'											objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label",Cstr((dicItems(iCounter)))	
	'									End If
	
										'Added by Vallari - 01-Apr-11 - All radio button labels changed to lower case
										objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label",lcase((dicItems(iCounter)))
										If objWin.JavaRadioButton(dicKeys(iCounter)).Exist(SISW_MIN_TIMEOUT)=False Then
												objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label",(dicItems(iCounter))
										End If
	
										If objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty("value") = "0" Then
											objWin.JavaRadioButton(dicKeys(iCounter)).Set "ON"
										End If	
								
							 Case "ErrorMsg"
										sMsg = dicItems(iCounter)

						End Select
						
						 If Err.Number < 0 Then	
							Fn_SchMgr_SchPropertyOperations = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to modify " & dicKeys(iCounter)  & " schedule property.")
							objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
							objWin.JavaButton("Cancel").Click
							Set objWin = Nothing
							Exit Function
						 Else 
							Fn_SchMgr_SchPropertyOperations = TRUE
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully modify " & dicKeys(iCounter)  & " schedule property.")
						End If
				   End If
				Next
				JavaDialog("Confirmation").SetTOProperty "title", "Confirmation"
				wait 3
				If Ucase(sButtonName) = "OK"  And Ucase(sConfirmationBtn) = "YES"Then
					objWin.JavaButton("OK").Click
					JavaDialog("Confirmation").WaitProperty "displayed",1,20000
					If JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
						JavaDialog("Confirmation").JavaButton("Yes").Click
					End If
				ElseIf Ucase(sButtonName) = "APPLY" And Ucase(sConfirmationBtn) = "YES"Then
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Apply").Click
					JavaDialog("Confirmation").WaitProperty "displayed",1,20000
					If JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
						JavaDialog("Confirmation").JavaButton("Yes").Click
					End If
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click
				ElseIf Ucase(sButtonName) = "OK" And Ucase(sConfirmationBtn) = "NO"Then
					objWin.JavaButton("OK").Click
					JavaDialog("Confirmation").WaitProperty "displayed",1,20000
					If JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
						JavaDialog("Confirmation").JavaButton("No").Click
					End If
				ElseIf Ucase(sButtonName) = "APPLY" And Ucase(sConfirmationBtn) = "NO"Then
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Apply").Click
					JavaDialog("Confirmation").WaitProperty "displayed",1,20000
					If JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
						JavaDialog("Confirmation").JavaButton("No").Click
					End If
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click
				End If

				'''''''''''''''''''''''''sButtonName click '''''''''''  Added By  : Harshal Tanpure. Date : 1-June-2011'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").Exist(SISW_DEFAULT_TIMEOUT) Then
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton(sButtonName).Click
					 If Err.Number < 0 Then	
							Fn_SchMgr_SchPropertyOperations = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Click " & sButtonName  & " Button")
							objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
							objWin.JavaButton("Cancel").Click
							Set objWin = Nothing
							Exit Function
					Else
						Fn_SchMgr_SchPropertyOperations = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully Clicked " & sButtonName  & " Button")
					End If					
				End If
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				If sAction = "ModifyWithoutMsgVerification" Then
				    Set objWin = Nothing
					Exit Function
				End If
				
				bReturn = Fn_SchMgr_DialogMsgVerify("Error", sMsg,"OK")
				If bReturn Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Handel error dialog.")
					Fn_SchMgr_SchPropertyOperations = FALSE
					Set objWin = Nothing
					Exit Function
				End If
				
				bReturn = Fn_SchMgr_DialogMsgVerify("Scheduling Error", sMsg,"OK")
				If bReturn Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Handel error dialog.")
					Fn_SchMgr_SchPropertyOperations = FALSE
					Set objWin = Nothing
					Exit Function
				End If

			Case "Verify"
				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "CustomerName" ,"CustomerNumber","InitialWBSValue","WBSFormat"
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "value") = dicItems(iCounter) Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.Close
									Set objWin = Nothing
									Exit Function
								End If
								
							Case "StartDate" ,"FinishDate" 
								'Added by Nilesh
'								If dicKeys(iCounter)="FinishDate" Then
'										Set objCheckBox=Description.Create()
'										objCheckBox("Class Name").value="JavaCheckBox"
'										set sObjects=JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").ChildObjects(objCheckBox)
'										sValue= sObjects.Count
'										If  sValue=4 Then
'												JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaCheckBox("FinishDate").SetToProperty "Index",2
'										End If
'								End If


								'commented above code by Koustubh
								
								sDate = objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "value")
'								sDate = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaCheckBox(dicKeys(iCounter)).GetROProperty( "label")
'								aDate = split(sDate, " ", -1,1)

								If Trim(sDate) =  Trim(dicItems(iCounter)) Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.Close
									Set objWin = Nothing
									Exit Function
								End If

							Case "SchPriority" ,"SchStatus","TaskStatus" 
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "value") = dicItems(iCounter) Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case "State"
								
								If objWin.JavaStaticText("State").GetROProperty("label") = dicItems(iCounter) Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
							
							Case  "IsScheduleTemplate" ,"IsPublic","FinishDateSchedul", "IsPercentLinked","Published","NotificationsEnabled","AreDatesLinked","ExcnOverrideEnabled","IsTemplate"
'								If Lcase(objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty("label")) = Lcase(Cstr(dicItems(iCounter))) Then
'									Fn_SchMgr_SchPropertyOperations = TRUE
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
'								Else
'									Fn_SchMgr_SchPropertyOperations = FALSE
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
'									Exit Function
'								End If

								'Added Lcase to the following line by Omkar - 10 Mr 2011
								If dicKeys(iCounter) = "IsScheduleTemplate" Then
									 dicKeys(iCounter) = "IsTemplate"
								End If
								objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label",Cstr(Lcase(dicItems(iCounter)))	

								sValue = objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty( "value")
								If  sValue = "1" And  Lcase(Cstr(dicItems(iCounter))) = "true" Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								ElseIf sValue = "1" And  Lcase(Cstr(dicItems(iCounter))) = "false" Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

							Case "ScheduleMembers"
								Dim ItemCount,arrItem,iIndex,iIndexItem
								arrItem = Split(dicItems(iCounter),",",-1,1)
								objWin.JavaStaticText("BootomLink").Click  8,6,"LEFT"
								Wait(2)
								ItemCount = objWin.JavaList(dicKeys(iCounter)).GetROProperty("items count")

								If  Err.Number < 0 Then
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click All bottom link")
									Set objWin = Nothing
									objWin.Close
									Exit Function
								End If
							
								For iIndexItem = 0 to Ubound(arrItem) 
									For iIndex = 0 to ItemCount - 1
										 If  objWin.JavaList(dicKeys(iCounter)).GetItem(iIndex) = arrItem(iIndexItem) Then
											 Exit For
										End If
									Next

									If  Cstr(iIndex) = Cstr(ItemCount) Then
										Fn_SchMgr_SchPropertyOperations = FALSE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter) & " with value " &arrItem(iIndexItem))
										objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
										objWin.JavaButton("Cancel").Click
										Set objWin = Nothing
										Exit Function
									Else
										Fn_SchMgr_SchPropertyOperations = TRUE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property  " & dicKeys(iCounter) & " with value " &arrItem(iIndexItem))
									End If
								Next
						End Select
				   End If
				Next
			  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click

			Case "IsEditable"

				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "CustomerName" ,"CustomerNumber"
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "editable") = "1" Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case "StartDate" ,"FinishDate" 
								If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaEdit(dicKeys(iCounter)).GetROProperty( "enabled") =  "1" Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

							Case "StatusDropDown" ,"TskStatusDrpDwn","PriorityDropDown","ScheduleDeliverable"
								If objWin.JavaButton(dicKeys(iCounter)).GetROProperty( "enabled") = "1" Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case  "IsTemplate" ,"IsPublic","FinishDateSchedul", "IsPercentLinked","Published","NotificationsEnabled","AreDatesLinked"
								If objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty( "enabled") = "1" Then
									Fn_SchMgr_SchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_SchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

						End Select
				   End If
				Next
				 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click

		Case "SchDeliverable"

			For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "ScheduleDeliverable"

									If dicScheduleProperty("ScheduleDeliverable") =True Then
										If objWin.JavaButton("ScheduleDeliverable").GetROProperty( "enabled") = "1" Then
										objWin.JavaButton("ScheduleDeliverable").Click
											If Err.number < 0 Then
												Fn_SchMgr_SchPropertyOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [ScheduleDeliverable] Button")
												Set objWin = Nothing
												Exit Function
											Else
													Fn_SchMgr_SchPropertyOperations = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Button ScheduleDeliverable is successfully clicked")
													Exit Function
												End If 
										Else
												Fn_SchMgr_SchPropertyOperations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Button ScheduleDeliverable is not enabled")
												Set objWin = Nothing
												Exit Function
										End if
								End If 
					End Select
			End if
		Next
		End Select
	Else
		Fn_SchMgr_SchPropertyOperations = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to displayed Properties dialog.")
	End If

	Set objWin = Nothing

End  Function

'*********************************************************		Function to  modify , Verify ,  to check IsEditable  Task  property.***********************************************************************

'Function Name		:					Fn_SchMgr_TaskPropertyOperations 

'Description			 :		 		  This function is used to  modify , Verify ,  to check IsEditable Task  properties.

'Parameters			   :	 			sAction: Action need to perform.
'                                                   dicProperties: Dictionary object to hold property names and the values
'                                                   sButtonName: Button to be clicked after modifying properties
											
'Return Value		   : 			  True/False  

'Pre-requisite			:		 		Schedule manger panel need to be open.
'                                                   Task need to be selected.
'                                                   Properties dialog need to be displayed.

'Examples				:			Call Fn_SchMgr_TaskPropertyOperations ("Modify",dicTaskProperty,"OK")
'												Call Fn_SchMgr_TaskPropertyOperations ("IsEditable",dicTaskProperty,"")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							30-May-2010	   1.0
'										Pritam					        16-Jan-2012    1.1            Added case "State" in Case "Verify"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_TaskPropertyOperations(sAction,dicProperties,sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TaskPropertyOperations"
   On Error Resume Next
   Dim dicCount , dicKeys , dicItems , iCounter,objWin ,bReturn,sDate,aDate,arrDate
   Dim sTemplateType,intNoOfObjects1,i
   Set  objWin =   JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("TaskProperties")

   If  not objWin.Exist(SISW_MIN_TIMEOUT) Then
		bReturn = Fn_MenuOperation("Select","View:View Properties")
		If bReturn = False Then
				Fn_SchMgr_TaskPropertyOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [View:Properties]")
				objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
				objWin.JavaButton("Cancel").Click
				Set objWin = Nothing
				Exit Function
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [View:Properties]")
		End If
	End If

	Call Fn_ReadyStatusSync(2) 

	 If  objWin.Exist(SISW_DEFAULT_TIMEOUT) Then
		dicCount  = dicProperties.Count
		 dicItems = dicProperties.Items
		 dicKeys = dicProperties.Keys

		Select Case  sAction
			Case "Modify"
				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "Duration" ,"WorkEstimate" ,"WorkComplete" , "WorkCompletePercent" ,"WorkEstimate","TaskType"
								'[TC1123-20161010-24_10_2016-VivekA-Maintenance] - 'changes done by shweta Rathod as per design change as discussed with sandeep chavan
							    If dicKeys(iCounter) = "Duration" Then
         							objWin.JavaEdit(dicKeys(iCounter)).SetTOProperty "attached text","Task Duration:"
        						End If
        						'--------------------------------------------------
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty("enabled") = "1" Then
									objWin.JavaEdit(dicKeys(iCounter)).Set dicItems(iCounter)
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to modify " & dicKeys(iCounter)  & " Task property.")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case "StartDate" ,"ActualStartDate" , "FinishDate" ,"ActualFinishDate" 
'								If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("TaskProperties").JavaCheckBox(dicKeys(iCounter)).GetROProperty("enabled") = "1" Then
'									 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("TaskProperties").JavaCheckBox(dicKeys(iCounter)).Object.setDate(dicItems(iCounter))
								If inStr(dicItems(iCounter), "~")>0 Then	''  added if condtion to check  date seperator and split Accordigly
										arrDate = Split (dicItems(iCounter),"~")
								ElseIf inStr(dicItems(iCounter), " ")>0 Then
										arrDate = Split (dicItems(iCounter)," ")
								End If
								If  objWin.JavaEdit(dicKeys(iCounter)).GetROProperty("enabled") = "1" Then
									objWin.JavaEdit(dicKeys(iCounter)).RefreshObject
									wait 1
								     objWin.JavaEdit(dicKeys(iCounter)).Click 1,1
								     wait 7
									 objWin.JavaEdit(dicKeys(iCounter)).Set arrDate(0)
									 wait 1
									 objWin.JavaEdit(dicKeys(iCounter)).Activate
									 wait 1
									 objWin.JavaEdit(dicKeys(iCounter)).RefreshObject
									 
									Set WshShell = CreateObject("WScript.Shell")
									'WshShell.SendKeys "{ESC}"
									WshShell.SendKeys "{TAB}"
									Set WshShell = Nothing
									wait 2
									If Ubound(arrDate) = 1 Then
                                    	objWin.JavaList(dicKeys(iCounter)).Select arrDate(1)
                                    	Wait 1
									End If
                                Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to modify " & dicKeys(iCounter)  & " Task property.")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

							Case "Constraint" ,"Priority","Status","FixedType","WorkflowTrigger","WorkflowTaskTemplate"

								If dicKeys(iCounter) =   "Constraint" Then
									objWin.JavaButton("ConstraintDrpDwn").Click
									Wait(3)
									bReturn =  Fn_iComboSet(objWin,dicItems(iCounter))
								ElseIf dicKeys(iCounter) =   "Priority" Then
									objWin.JavaButton("PriorityDrpDwn").Click
									Wait(2)
									bReturn =  Fn_iComboSet(objWin,dicItems(iCounter))
								ElseIf dicKeys(iCounter) =   "Status" Then
									'objWin.JavaButton("StatusDrpDwn").Click
									Wait(2)
									'bReturn =  Fn_iComboSet(objWin,dicItems(iCounter))
									objWin.JavaEdit("Status").Activate	
									objWin.JavaEdit("Status").Type dicItems(iCounter)
									if err.number < 0 then
										bReturn = false
									else
										bReturn = true	
									end if
									objWin.JavaEdit("Status").Activate
									
								ElseIf dicKeys(iCounter) =   "FixedType" Then
									objWin.JavaButton("FixedTypeDrpDwn").Click
									Wait(2)
									bReturn =  Fn_iComboSet(objWin,dicItems(iCounter))
								ElseIf dicKeys(iCounter) =   "WorkflowTrigger" Then
'									objWin.JavaButton("WrkFlwTriggerDrpDwn").Click
'									Wait(2)
'									bReturn =  Fn_iComboSet(objWin,dicItems(iCounter))
									objWin.JavaEdit("WorkflowTrigger").Set dicItems(iCounter)
									if err.number < 0 then
										bReturn = false
									else
										bReturn = true	
									end if
									objWin.JavaEdit("WorkflowTrigger").Activate
								ElseIf dicKeys(iCounter) =   "WorkflowTaskTemplate" Then
									bReturn = false
									objWin.JavaButton("WrkFlwTemDrpDwn").Click
									Wait(2)
                                    Set sTemplateType=Description.Create()
									sTemplateType("Class Name").value = "JavaStaticText"
							
									Set  intNoOfObjects1 = objWin.ChildObjects(sTemplateType)
									  For i = 0 to intNoOfObjects1.count-1
										   If  intNoOfObjects1(i).getROProperty("label") = dicItems(iCounter) Then
													intNoOfObjects1(i).Click 1,1
													bReturn = True
													Exit for
										   End If
									  Next
'									objWin.JavaButton("WrkFlwTemDrpDwn").Click
'									Wait(2)
'									bReturn =  Fn_iComboSet(objWin,dicItems(iCounter))
'									objWin.JavaEdit("WorkflowTaskTemplate").Set dicItems(iCounter)
'									if err.number < 0 then
'										bReturn = false
'									else
'										bReturn = true	
'									end if
'									objWin.JavaEdit("WorkflowTaskTemplate").Activate
								End If

								If bReturn = False Then
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to modify " & dicKeys(iCounter)  & " Task property.")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
								Case  "AutoComplete"
								objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "attached text" ,Lcase(Cstr (dicItems(iCounter)))
								If objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty("enabled") = "1" Then
									objWin.JavaRadioButton(dicKeys(iCounter)).Set "ON"
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to modify " & dicKeys(iCounter)  & " Task property.")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
						End Select
						
						 If Err.Number < 0 Then	
							Fn_SchMgr_TaskPropertyOperations = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to modify " & dicKeys(iCounter)  & " Task property.")
							objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
							objWin.JavaButton("Cancel").Click
							Set objWin = Nothing
							Exit Function
						 Else 
							Fn_SchMgr_TaskPropertyOperations = TRUE
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully modify " & dicKeys(iCounter)  & " Task property.")
						End If
				   End If
				Next
				JavaDialog("Confirmation").SetTOProperty "title", "Confirmation"
				wait 5
				If Ucase(sButtonName) = "OK"  Then
					objWin.JavaButton("OK").Click
					If JavaDialog("Confirmation").Exist(SISW_DEFAULT_TIMEOUT) Then
						JavaDialog("Confirmation").JavaButton("Yes").Click
					End If
				ElseIf Ucase(sButtonName) = "APPLY" Then
					 objWin.JavaButton("Apply").Click
					If JavaDialog("Confirmation").Exist(SISW_DEFAULT_TIMEOUT) Then
						JavaDialog("Confirmation").JavaButton("Yes").Click
					End If
				ElseIf Ucase(sButtonName) = "OK_NO" Then
					objWin.JavaButton("OK").Click
					wait 1
					If JavaDialog("Confirmation").Exist(SISW_DEFAULT_TIMEOUT) Then
						JavaDialog("Confirmation").JavaButton("No").Click
					End If
				Else
					 objWin.JavaButton("Cancel").Click
				End If
				

			Case "Verify"
				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "Duration" ,"WorkEstimate" ,"WorkComplete" , "WorkCompletePercent" ,"Percent Complete","WorkEstimate","TaskType","TaskDuration","TaskWorkEstimate"
								If Fn_SISW_UI_Object_Operations("Fn_SchMgr_TaskPropertyOperations","Exist", objWin.JavaEdit(dicKeys(iCounter)),"") = True Then
									If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "value") = dicItems(iCounter) Then
										Fn_SchMgr_TaskPropertyOperations = TRUE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the task property " & dicKeys(iCounter))
									Else
										Fn_SchMgr_TaskPropertyOperations = FALSE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the task  property " & dicKeys(iCounter))
										objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
										objWin.JavaButton("Cancel").Click
										Set objWin = Nothing
										Exit Function
									End If
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the task  property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

							Case "StartDate" ,"ActualStartDate" , "FinishDate" ,"ActualFinishDate" 
								sDate =  objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "value")

								  'Changed By Sushma to consider 'No date Set ' 
								If Instr("No date set.", sDate)>0   Then
									aDate = split(sDate, ".", -1,1)     '' If Not date set   , remove appending dot.
								Else
									aDate = split(sDate, " ", -1,1)      ''If date is set the remove time part.
								End If

								If Trim(aDate(0)) =  Trim(dicItems(iCounter)) Then 
									  Fn_SchMgr_TaskPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the task  property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the task  property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

							Case "Constraint" ,"Priority","Status","FixedType","WorkflowTrigger","WorkflowTaskTemplate"
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "value") = dicItems(iCounter) Then
									Fn_SchMgr_TaskPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the task property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the task  property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
							Case "State"
								
								If objWin.JavaStaticText("State").GetROProperty("label") = dicItems(iCounter) Then
									Fn_SchMgr_TaskPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the task property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the task  property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

								
							Case  "AutoComplete"
								Dim sValue
									objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "attached text" ,Lcase(Cstr (dicItems(iCounter)))
									sValue = objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty( "value")
									If  sValue = "1" And   Lcase(Cstr(dicItems(iCounter))) = "true" Then
										Fn_SchMgr_TaskPropertyOperations = TRUE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the task property " & dicKeys(iCounter))
									ElseIf sValue = "1" And   Lcase(Cstr(dicItems(iCounter))) = "false" Then
										Fn_SchMgr_TaskPropertyOperations = TRUE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the task property " & dicKeys(iCounter))
									Else
										Fn_SchMgr_TaskPropertyOperations = FALSE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the task property " & dicKeys(iCounter))
										objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
										objWin.JavaButton("Cancel").Click
										Set objWin = Nothing
										Exit Function
									End If

						Case "ResourceAssignments" , "TaskDeliverable"
								Dim ItemCount,arrItem,iIndex,iIndexItem,aResourse
								arrItem = Split(dicItems(iCounter),",",-1,1)
								'objWin.JavaStaticText("BottomLink").SetTOProperty "label","All"
								'objWin.JavaStaticText("BootomLink").Click 8,6,"LEFT"
								 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("TaskProperties").JavaStaticText("BottomLink").SetTOProperty "label","All"
								 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("TaskProperties").JavaStaticText("BottomLink").Click 8,6,"LEFT" 

								Wait(2)
								ItemCount = objWin.JavaList(dicKeys(iCounter)).GetROProperty("items count")

								If Err.Number < 0 Then
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to click All bottom link")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

								For iIndexItem = 0 to Ubound(arrItem) 
									For iIndex = 0 to ItemCount - 1
										aResourse = Split (objWin.JavaList(dicKeys(iCounter)).GetItem(iIndex),"(",-1,1)
										 If  Trim(aResourse(0)) = Trim(arrItem(iIndexItem)) Then
											 Exit For
										End If
									Next

									If  Cstr(iIndex) = Cstr(ItemCount) Then
										Fn_SchMgr_TaskPropertyOperations = FALSE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the task  property " & dicKeys(iCounter) & " with value " &arrItem(iIndexItem))
										objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
										objWin.JavaButton("Cancel").Click
										Set objWin = Nothing
										Exit Function
									Else
										Fn_SchMgr_TaskPropertyOperations = TRUE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the task  property  " & dicKeys(iCounter) & " with value " &arrItem(iIndexItem))
									End If
								Next
						End Select
				   End If
				Next
			  objWin.JavaButton("Cancel").Click

			Case "IsEditable"

				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "Duration" ,"WorkEstimate" ,"WorkComplete" , "WorkCompletePercent" ,"WorkEstimate","TaskType"
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "editable") = "1" Then
									Fn_SchMgr_TaskPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case "StartDate" ,"ActualStartDate" , "FinishDate" ,"ActualFinishDate"
								If  objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "enabled") =  "1" Then
									Fn_SchMgr_TaskPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

							Case "Constraint" ,"Priority","Status","FixedType","WorkflowTrigger","WorkflowTaskTemplate","TaskDeliverable"
								If objWin.JavaButton(dicKeys(iCounter)).GetROProperty( "enabled") = "1" Then
									Fn_SchMgr_TaskPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case  "AutoComplete"
								If objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty( "enabled") = "1" Then
									Fn_SchMgr_TaskPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_TaskPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.Close
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,5000
									objWin.JavaButton("Cancel").Click
									Exit Function
								End If

						End Select
				   End If
				Next
			  objWin.JavaButton("Cancel").Click

		  Case "TSKDeliverable"    ' Added by Shreyas 13 th May 2011

					For iCounter = 0 to dicCount - 1
							If  dicItems(iCounter) <> ""Then		
								Select Case dicKeys(iCounter) 		
												Case "TaskDeliverable"
					
														If dicTaskProperty("TaskDeliverable") =True Then
																If objWin.JavaButton("TaskDeliverable").GetROProperty( "enabled") = "1" Then
																objWin.JavaButton("TaskDeliverable").Click
																		If Err.number < 0 Then
																			Fn_SchMgr_TaskPropertyOperations = False
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [TaskDeliverable] Button")
																			Set objWin = Nothing
																			Exit Function
																		Else
																				Fn_SchMgr_TaskPropertyOperations = True
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Button TaskDeliverable is successfully clicked")
																				Exit Function
																		End If 
																Else
																		Fn_SchMgr_TaskPropertyOperations = False
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Button TaskDeliverable is not enabled")
																		Set objWin = Nothing
																		Exit Function
																End if
														End If 
									End Select
							End if
					Next
			End Select
		
	Else
		Fn_SchMgr_TaskPropertyOperations = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to displayed Task Properties dialog.")
	End If

	Set objWin = Nothing

End  Function
 '*********************************************************		Function to create detail  Task		***********************************************************************
'Function Name		:        Fn_SchMgr_TaskDetailCreate   

'Description	    	:        Creates an task with detail information

'Parameters		     :    		sTaskType: Task type to be selected
'			                         		 sTaskID: Unique ID for the Task [if non-empty, then enter]
'							          		sTaskRevID: Revision of the Task [if non-empty, then enter] - if any one of the fields (id/rev) are blank then click Assign button
'									 		sTaskName: Name of the Task
'									  		sdesc: Description of the Task
' 											dicTaskParam : Dictionary paramter  for detail creation

'Return Value		: 			TaskId-RevId  

'Pre-requisite	    :		     Shedule Manager  pane should be open.

'Examples		    :			Call  Fn_SchMgr_TaskDetailCreate ("Schedule Task","","","TestTask","sfdwtrtyjghjjg",dicTaskInfo)

'History		    :		
'													Developer Name				Date						Rev. No.						Changes Done						Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Rupali 							     31/05/2010			              1.0								
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'
'													Ganesh B 					    21-Jul-2014			              1.1								modified function as per design changes for "New Task" Dialog on TC11.1(20140709)
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_TaskDetailCreate(sTaskType ,sTaskID,sTaskRevID,sTaskName ,sdesc,dicTaskInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TaskDetailCreate"
	Dim WshShell, arrDate
	Dim objTask, objSelectType, intNoOfObjects, bFlag
	Dim scellRec,x,y,comma,sWidth,sHeight, objTreeTable,sIndex,iCounter
	Dim objTaskAssign, objHierarchyTree     
	Dim   MyTime1, MyTime2
	Dim sTreeDescription, itr, sTreeItem, iItemsCount
	Dim sNewItemMenu

	Set objTask = Fn_SISW_PPM_GetObject("New Task")
	sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"FileNewTask")
	If objTask.Exist(SISW_MIN_TIMEOUT) = False  Then
		'Select menu  [File -> New -> Task...]               
		 bReturn = Fn_MenuOperation("Select",sNewItemMenu)	  	   
		 If bReturn = False Then
			Fn_SchMgr_TaskDetailCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [File:New:Task...]")
			Set objTask = Nothing
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [File:New:Task...]")
		End If
	End If

		'Check the existence of the "NewTask" Window
		If objTask.Exist (SISW_DEFAULT_TIMEOUT)  Then
			'Select  "Task Type"
			If trim(sTaskType) = "ScheduleTask" Then
				sTaskType = "Schedule Task"
			End If
			'Select  "Task Type"
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			wait(1)	
			If Fn_SISW_UI_Object_Operations("Fn_SchMgr_TaskDetailCreate", "Exist", objTask.JavaTree("TaskType") , "") Then
				iItemCount=Fn_UI_Object_GetROProperty("Fn_SchMgr_TaskDetailCreate",objTask.JavaTree("TaskType"), "items count")
				For iCount=0 To iItemCount-1
					crrItem=objTask.JavaTree("TaskType").GetItem(iCount)
					If Trim(crrItem)="Most Recently Used:"+Trim(sTaskType) Then
						bFlag=True
						Exit For
					ElseIf Trim(crrItem)="Complete List" Then
						Exit For
					End If
				Next
			
				If bFlag=True Then
					Call Fn_JavaTree_Select("Fn_SchMgr_TaskDetailCreate", objTask, "TaskType","Most Recently Used")
					Call Fn_JavaTree_Select("Fn_SchMgr_TaskDetailCreate", objTask, "TaskType","Most Recently Used:"+sTaskType)
				Else
					Call Fn_UI_JavaTree_Expand("Fn_SchMgr_TaskDetailCreate", objTask, "TaskType","Complete List")
					Call Fn_JavaTree_Select("Fn_SchMgr_TaskDetailCreate", objTask, "TaskType","Complete List")
					Call Fn_JavaTree_Select("Fn_SchMgr_TaskDetailCreate", objTask, "TaskType","Complete List:"+sTaskType)	
				End If
				If Err.number < 0 Then
					Fn_SchMgr_TaskDetailCreate = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Task Type [" + sTaskType + "]")
					objTask.JavaButton("Close").Click micLeftBtn
					Set objTask = Nothing
					Exit Function
				End If
				'Clicking On Next button
				objTask.JavaButton("Next").WaitProperty "enabled", 1, 60000
				call Fn_Button_Click("Fn_SchMgr_TaskDetailCreate", objTask, "Next")
				If Err.number < 0 Then
					Fn_SchMgr_TaskDetailCreate = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Task Id as [" + sTaskID + "]")
					objTask.JavaButton("Close").Click micLeftBtn
					Set objTask = Nothing
					Exit Function
				End If
			End If
				
		'Check  "Item Id and Revision ID"
		If sTaskID <> "" Then
			objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Task ID:"
			Call Fn_Edit_Box("Fn_SchMgr_TaskDetailCreate", objTask,"TaskEdit",sTaskID)
			If Err.number < 0 Then
				Fn_SchMgr_TaskDetailCreate = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Task Id as [" + sTaskID + "]")
				objTask.JavaButton("Close").Click micLeftBtn
				Set objTask = Nothing
				Exit Function
			End If
		End If
'		 If sTaskRevID <> "" Then
'			objTask.JavaEdit("TaskRev").Set sTaskRevID
'			If Err.number < 0 Then
'				Fn_SchMgr_TaskDetailCreate = False
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Task RevId as [" + sTaskRevID + "]")
'				objTask.JavaButton("Close").Click micLeftBtn
'				Set objTask = Nothing
'				Exit Function
'			End If
'		End If
		Wait 1
		If sTaskID = "" Or  sTaskRevID = "" Then
			'Click on "Assign" button
			objTask.JavaButton("Assign").WaitProperty "enabled", 1, 20000
			call Fn_Button_Click("Fn_SchMgr_TaskDetailCreate", objTask, "Assign")
			If Err.number < 0 Then
					Fn_SchMgr_TaskDetailCreate = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Assign] Button")
					Set objTask = Nothing
					Exit Function
			End If
		End If
		Wait 1
			objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Task ID:"
		sTskAssignId = objTask.JavaEdit("TaskEdit").GetROProperty("value")
'		sTskAssignRev = objTask.JavaEdit("TaskRev").GetROProperty("value")

		'Set the Task  Name
		objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Name:"
		Call Fn_Edit_Box("Fn_SchMgr_TaskDetailCreate", objTask,"TaskEdit",sTaskName)
		 If Err.number < 0 Then
			Fn_SchMgr_TaskDetailCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Set Task Name as [" + sTaskName + "]")
			objTask.JavaButton("Close").Click micLeftBtn
			Set objTask = Nothing
			Exit Function
		End If

		'Set the Description
		objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Description:"
		Call Fn_Edit_Box("Fn_SchMgr_TaskDetailCreate", objTask,"TaskEdit",sdesc)

		'Set Fix Type 
		If dicTaskInfo("FixedType") <> "" Then
			objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Fixed Type:"
			call Fn_Button_Click("Fn_SchMgr_TaskDetailCreate", objTask, "FixedType")
			 iItemsCount = objTask.JavaTree("FixedType").GetROProperty("items count")
			For itr = 0 to iItemsCount-1
				sTreeItem = objTask.JavaTree("FixedType").GetItem (itr)
				sTreeDescription = objTask.JavaTree("FixedType").GetColumnValue(sTreeItem, "Description")
				If sTreeDescription = dicTaskInfo("FixedType") Then
					objTask.JavaTree("FixedType").Activate sTreeItem 
					Exit For
				End If
			Next
            If Err.number < 0 Then
				Fn_SchMgr_TaskDetailCreate = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
				objTask.JavaButton("Close").Click micLeftBtn
				Set objTask = Nothing
				Exit Function
			End If
		End If

		Wait 1

		'Set the value of Admin Task.
		 If dicTaskInfo("AdminTask") <> "" Then
			objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Administrative Task?:"
			If Lcase(cstr(dicTaskInfo("AdminTask"))) = "true" Then
				objTask.JavaRadioButton("True").Set "ON"
			ElseIf Lcase(cstr(dicTaskInfo("AdminTask"))) = "false" Then
				objTask.JavaRadioButton("False").Set "ON"
			Else
				Fn_SchMgr_TaskDetailCreate = False																			      
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAILED: Wrong Argument Value[" & dicTaskInfo("AdminTask") & "].")
				Set objTask = Nothing
				Exit Function
			End If
		 End If 

		'Set the value of Impact Assessment.
		 If dicTaskInfo("ImpactAssReq") <> "" Then
			objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Impact Assessment Required?:"
			If Lcase(cstr(dicTaskInfo("ImpactAssReq"))) = "true" Then
				objTask.JavaRadioButton("True").Set "ON"
			ElseIf Lcase(cstr(dicTaskInfo("ImpactAssReq"))) = "false" Then
				objTask.JavaRadioButton("False").Set "ON"
			Else
				Fn_SchMgr_TaskDetailCreate = False																			      
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAILED: Wrong Argument Value[" & dicTaskInfo("ImpactAssReq") & "].")
				Set objTask = Nothing
				Exit Function
			End If
		 End If 
		'Set the value of Proposed Task.
		 If dicTaskInfo("ProposedTask") <> "" Then
			objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Proposed Task?:"
			If Lcase(cstr(dicTaskInfo("ProposedTask"))) = "true" Then
				objTask.JavaRadioButton("True").Set "ON"
			ElseIf Lcase(cstr(dicTaskInfo("ProposedTask"))) = "false" Then
				objTask.JavaRadioButton("False").Set "ON"
			Else
				Fn_SchMgr_TaskDetailCreate = False																			      
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAILED: Wrong Argument Value[" & dicTaskInfo("ProposedTask") & "].")
				Set objTask = Nothing
				Exit Function
			End If
		 End If 
		
'		'Set the value of Category 
		 If dicTaskInfo("Category") <> "" Then
		 	 objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Category:"
			 If objTask.JavaButton("FixedType").Exist(SISW_MICRO_TIMEOUT) Then
				objTask.JavaButton("FixedType").Click micLeftBtn
				Wait 15
				bFlag= False
				If objTask.JavaWindow("Shell").JavaTree("Tree").Exist(4) Then
					iItemsCount = objTask.JavaWindow("Shell").JavaTree("Tree").GetROProperty("items count")
				 	For itr = 0 to iItemsCount-1
						sTreeItem = objTask.JavaWindow("Shell").JavaTree("Tree").GetItem (itr)
						If sTreeItem = dicTaskInfo("Category") Then
							objTask.JavaWindow("Shell").JavaTree("Tree").Activate sTreeItem 
							bFlag = True   '' Value found among LOVs and is set 
							Exit For
						End If
					Next
					If Err.number < 0 Then
						Fn_SchMgr_TaskDetailCreate = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select"  & dicTaskInfo("Category")& " from Category list")
						objTask.JavaButton("Close").Click micLeftBtn
						Set objTask = Nothing
						Exit Function
					End If
				End If
			  Else				 
				 objTask.JavaEdit("Category").Set dicTaskInfo("Category")
				 bFlag = True
			  End If		 
			  If bFlag= False Then    '' Value is not found among LOVs and is not set , So set it explicitly
					objTask.JavaEdit("Category").Set dicTaskInfo("Category")
			  End If
		 End If
		'Set the value of Complexity 
		 If dicTaskInfo("Complexity") <> "" Then
			objTask.JavaStaticText("Property_Label").SetTOProperty "label", "Complexity:"
			Call Fn_Edit_Box("Fn_SchMgr_TaskDetailCreate", objTask,"TaskEdit",dicTaskInfo("Complexity"))
		 End If 
		 
		 objTask.JavaButton("Next").WaitProperty "enabled", 1, 20000
		call Fn_Button_Click("Fn_SchMgr_TaskDetailCreate", objTask, "Next")
		If Err.number < 0 Then
			Fn_SchMgr_TaskDetailCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
			objTask.JavaButton("Close").Click micLeftBtn
			Set objTask = Nothing
			Exit Function
		End If
		 	'Set  the value of  Create Phase Gate Structure
		If dicTaskInfo("CreatePhaseGate") <> "" Then
			objTask.JavaCheckBox("PhaseGateTask").WaitProperty "displayed", 1, 20000
			If Lcase(cstr(dicTaskInfo("CreatePhaseGate"))) = "true" Then
				objTask.JavaCheckBox("PhaseGateTask").Set "ON"
			ElseIf Lcase(cstr(dicTaskInfo("CreatePhaseGate"))) = "false" Then
				objTask.JavaCheckBox("PhaseGateTask").Set "OFF"
			Else
				Fn_SchMgr_TaskDetailCreate = False																			      
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAILED: Wrong Argument Value[" & dicTaskInfo("CreatePhaseGate") & "].")
				Set objTask = Nothing
				Exit Function
			End If
		End If

		'Set the value of  Start date 
		If dicTaskInfo("StartDate") <> "" Then
				arrDate = Split(dicTaskInfo("StartDate")," ")
				objTask.JavaCalendar("StartDate").SetDate arrDate(0)
				MyTime1 = Split(arrDate(1), ":")
				If UBound(MyTime1) = 1 Then
					MyTime2 = arrDate(1) + ":00"
					objTask.JavaCalendar("StartTime").SetTime MyTime2
				Else
					objTask.JavaCalendar("StartTime").SetTime arrDate(1)
				End If	
			wait 2
		End If

		'Set the value  of  Finish date
		If dicTaskInfo("FinishDate") <> "" Then
			arrDate = Split(dicTaskInfo("FinishDate")," ")
			objTask.JavaCalendar("FinishDate").SetDate arrDate(0)
			MyTime1 = Split(arrDate(1), ":")
			If UBound(MyTime1) = 1 Then
					MyTime2 = arrDate(1) + ":00"
					objTask.JavaCalendar("FinishTime").SetTime MyTime2
				Else
					objTask.JavaCalendar("FinishTime").SetTime arrDate(1)
			End If	
			wait 2
		End If
		If Err.number < 0 Then
			Fn_SchMgr_TaskDetailCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Start date and End Date While creating New task.")
			objTask.JavaButton("Close").Click micLeftBtn
			Set objTask = Nothing
			Exit Function
		End If

			'Code added by Omkar to Handle the Schedule Error Dialog if Exists ...Date 13 April 2011
		If  objTask.JavaDialog("scheduling Error").Exist(SISW_MIN_TIMEOUT)=True Then
        			Fn_SchMgr_TaskDetailCreate = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Start date and End Date While creating New task.")
				objTask.JavaDialog("scheduling Error").JavaButton("OK").Click micLeftBtn
				objTask.JavaButton("Close").Click micLeftBtn
				Set objTask = Nothing
				Exit Function   			
		End If


		 'Click Next Button 
		call Fn_Button_Click("Fn_SchMgr_TaskDetailCreate", objTask, "Next")
		If Err.number < 0 Then
			Fn_SchMgr_TaskDetailCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
			objTask.JavaButton("Close").Click micLeftBtn
			Set objTask = Nothing
			Exit Function
		End If
		
		'Code added by Amisha to handle New Schedule Task error appeared when out of boundary date
		If  objTask.JavaWindow("NewScheduleTask").Exist(SISW_MIN_TIMEOUT)=True Then
    		Fn_SchMgr_TaskDetailCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select the Start date and End Date While creating New task.")
			objTask.JavaWindow("NewScheduleTask").JavaButton("OK").Click micLeftBtn
			objTask.JavaButton("Close").Click micLeftBtn
			Set objTask = Nothing
			Exit Function   			
		End If
		
		'Click Next Button 
		call Fn_Button_Click("Fn_SchMgr_TaskDetailCreate", objTask, "Next")
		If Err.number < 0 Then
			Fn_SchMgr_TaskDetailCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
			objTask.JavaButton("Close").Click micLeftBtn
			Set objTask = Nothing
			Exit Function
		End If
		 'Click Next Button 
		call Fn_Button_Click("Fn_SchMgr_TaskDetailCreate", objTask, "Next")
		If Err.number < 0 Then
			Fn_SchMgr_TaskDetailCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Next] Button")
			objTask.JavaButton("Close").Click micLeftBtn
			Set objTask = Nothing
			Exit Function
		End If

	'Assign member to task.
	If dicTaskInfo("AssignMember") <> "" Then
            objTask.Javabutton("Resource").Object.Click
			Set descJavaWindow = Description.Create()
			descJavaWindow("Class Name").value = "JavaWindow"
			Set objChildObjects = JavaWindow("ScheduleManagerWindow").ChildObjects(descJavaWindow)
			For iCounter =0 to objChildObjects.Count-1
				If objChildObjects(iCounter).getROProperty("label") = "Assign To Task" Then
						Set objTaskAssign =  objChildObjects(iCounter)
						Exit For
				End If
			Next
			Set objTreeTable = objTaskAssign.JavaTree("index:=0")
			
			aAssignMem = Split(dicTaskInfo("AssignMember"),",",-1,1)
			For iCounter = 0 to Ubound(aAssignMem) 
				
					aUserInfo  = Split( aAssignMem(iCounter), ":", -1,1)                                                         '' split  Organization:Engg:Designer:AutoTest1 (autotest1)
					sSearchText = trim(aUserInfo(UBound(aUserInfo)))                   ''Type  AutoTest1 (autotest1)"  OR GrpName Or RoleName or Disc Name
					objTaskAssign.JavaTab("to_class:=JavaTab").Select aUserInfo(0)                  
					objTaskAssign.JavaEdit("index:=1").Set  sSearchText     
					objTreeTable.Select  sTaskName
					wait 2
					Set objHierarchyTree = objTaskAssign.JavaTree("index:=1", "displayed:=1")
					If objHierarchyTree.GetROProperty("items count") <> 0 Then  											
								objHierarchyTree.Select aAssignMem(iCounter)
								objTaskAssign.JavaButton("attached text:=Add").WaitProperty  "enabled", 1, 20000
								objTaskAssign.JavaButton("attached text:=Add").Click    								
					Else
								Fn_SchMgr_TaskDetailCreate = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to assign member to task")
								'objTaskAssign.Close
								Set objTaskAssign = Nothing
								Set objTreeTable = Nothing
								Set objHierarchyTree = Nothing
								Exit Function
					End If  
			Next
			objTaskAssign.JavaButton("attached text:=OK").Click	
			Set objTaskAssign = Nothing
			Set objTreeTable = Nothing
			Set objHierarchyTree = Nothing
		End If
		If Err.Number < 0 Then
			Fn_SchMgr_TaskDetailCreate = False																			      
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [OK] Button")
			objTask.JavaButton("Close").Click micLeftBtn
			Set objTask = Nothing
			Exit Function
		End If

		'		'Set Create WorkFlow Task
		If dicTaskInfo("CreateWorkflowTask") <> "" Then
				If Lcase(cstr(dicTaskInfo("CreateWorkflowTask"))) = "true" Then
					objTask.JavaCheckBox("Create Workflow Task").Set "ON"
					wait(1)
					objTask.JavaButton("WorkflowButton").Click
					If Err.number < 0 Then
							Fn_SchMgr_TaskDetailCreate = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [ Workflow ...] Button")
							objTask.JavaButton("Close").Click micLeftBtn
							Set objTask = Nothing
					Else
							Fn_SchMgr_TaskDetailCreate = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked On [ Workflow....] Button")
						Exit Function
					End If
				ElseIf Lcase(cstr(dicTaskInfo("CreateWorkflowTask"))) = "false" Then
					objTask.JavaCheckBox("Create Workflow Task").Set "OFF"
				Else
					Fn_SchMgr_TaskDetailCreate = False																			      
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [ Workflow ...] Button")
					Set objTask = Nothing
					Exit Function
			End If
	End If 

		'' Add  Deliverable
	If dicTaskInfo("Deliverable") <> "" Then
		If Lcase(cstr(dicTaskInfo("Deliverable"))) = "true" Then
			objTask.JavaButton("Deliverable").Click
		End If
		If Err.number < 0 Then
					Fn_SchMgr_TaskDetailCreate = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [ Deliverable.] Button")
					objTask.JavaButton("Close").Click micLeftBtn
					Set objTask = Nothing
					Exit Function
		Else
					Fn_SchMgr_TaskDetailCreate = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked On [ Deliverable.] Button")
				Exit Function
			End If
	End If

		'Click on "Finish" button
		objTask.JavaButton("Finish").WaitProperty "enabled", 1, 20000
        objTask.JavaButton("Finish").Click
		If Err.Number < 0 Then
			Fn_SchMgr_TaskDetailCreate = False																			      
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Finish] Button")
			Set objTask = Nothing
			Exit Function
		End If   
		objTask.WaitProperty "displayed",0,25000
		If objTask.GetROProperty("displayed") = "1" Then
			Fn_SchMgr_TaskDetailCreate = False																			      
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "After click Finish button New Task dialog is displayed.")
			Set objTask = Nothing
			Exit Function
		End If
		Fn_SchMgr_TaskDetailCreate =True																	                                                       
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task [" + sTaskName + "] Created Successfully") 
		Set objTask = Nothing
		Set objTreeTable = Nothing
	Else 
		Fn_SchMgr_TaskDetailCreate = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[New Task] Dialog Not Found") 
		Set objTask = Nothing
	End If
End Function
 '*********************************************************		Function to create detail  Task		***********************************************************************
'Function Name		:        Fn_SchMgr_SchedulingErrorVerify   

'Description	    	:        Verifies teh message on Scheduling Error dialog

'Parameters		     :    		sMesssage: Message to be Verified [Optional]
'			                         		 sButton: Button to be clicked on the doalig

'Return Value		: 			True/False

'Pre-requisite	    :		     Scheduling Error dialog is diaplyed

'Examples		    :			Call  Fn_SchMgr_SchedulingErrorVerify ("", "OK")

'History		    :		
'	Developer Name		Date			Rev. No.	Changes Done						Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vallari 			03/06/2010		1.0			
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sushma Pagare		4-Jan-2013 		1.0			Added code : Error Dialog title changes to 'Validate Inline editing'  for certain specific errors.
'	Sushma						18-Jun-2013	   1.0
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_SchedulingErrorVerify(sMessage, sButton)

		Dim dicErrorInfo, bReturn 
		Set dicErrorInfo = CreateObject("Scripting.Dictionary")
		With dicErrorInfo 
		 .Add "Title", "Scheduling Error"
		 .Add "Message", sMessage		 
		 .Add "Button", sButton
		End with
		Fn_SchMgr_SchedulingErrorVerify = Fn_SISW_SchMgr_ErrorVerify(dicErrorInfo)
		
End Function

'*********************************************************		Function will assign task to user	***********************************************************************

'Function Name		:					Fn_SchMgr_TaskAssignment 

'Description			 :		 		  This function is used assign task to user.

'Parameters			   :	 			1. sAction : Action need to perform. (Assign , Remove , Verify, Modify)
'                                                    2. sTaskName :   Task Name which need to select      (If multiple task need to select pass it , separated)
'													3.aUser : array of users to be assigned to Task. (This is : seprated value. Memebrs:Engineer:Designer:User)
'													4.aResourceLevel: Array of resource level (maintaint the sequence of users)
'												   5.sAssignBehaviour : Assign behaviour to task.  (Pass as parameter  "Add " or "Overwrite"  existing Assignments) 
 
'Return Value		   : 		     True/False

'Pre-requisite			:		 	Schedule Manager window should be displayed .

'Examples				:			
'												Call Fn_SchMgr_TaskAssignment("Assign", "Schedule1:Task1", aUser, "","") 
'												 Use Array aUser = Array("Organization:Engg:Designer:AutoTest1 (autotest1)", "Disciplines:AutoDisp1:AutoTest1 (autotest1)", "Disciplines:AutoDisp1")												

'History:
'	Developer Name		Date		 Rev. No.	Changes Done																					Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Rupali				02-Jun-2010	   1.0
'	Sushma				22-May-2012	   1.0		Tc10 UI changes
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W			07-Dec-2012	   1.0      Modifeid case Modify
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vivek A				05-Jan-2016	   1.1      Added new Case : "QualificationTabOperation"								[TC1122-20151116d-31_12_2015-VivekA-NewDevelopment]
'												Added "AssignQualification","QualificationErrorVerify",
'												"RemoveQualification","VerifyInExistingQualificationsTable",
'												"VerifyInQualificationList","VerifyInQualificationLevelList"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam			12-Jan-2016	   1.1      Added Case "AssignVerifySearch"												[TC1122-20151116d00-12_01_2016-VivekA-NewDevelopment]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_TaskAssignment(sAction, sTaskName, aUser,aResourceLevel,sAssignBehaviour)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TaskAssignment"
	Dim objTaskAssign,bReturn,aTaskName,IntRows,aDbClick,iCount,iOuterCount
	Dim scellRec,x,y,comma,sWidth,sHeight,objTreeTable,sIndex,iCounter,sWidth2
	Dim iCnt,sBuildPath,bExpanded,aUserInfo,aResourceValue,aUserInfo1,sSearchText,jCount
	Dim sButton,aTaskNames,aQualifications,sQualification,sQualiLevel,sName,sSelectValue,iRowCount,bFlag
	Dim objChild,objErr
	Dim arrTask,sTask
	
	On Error Resume Next

	Set objTaskAssign = JavaWindow("ScheduleManagerWindow").JavaWindow("Task Assignment")
    
	'select Task from Schedule table
	aTaskName = Split(sTaskName,",",-1,1)

	If objTaskAssign.Exist (SISW_MIN_TIMEOUT) = False  Then
		If Ubound(aTaskName) > 0 Then
			bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect" ,sTaskName , " " , " " , " ")
		Else
			bReturn = Fn_SchMgr_SchTable_NodeOperation("Select" ,sTaskName , "" , " " , " ")
		End If	
		If bReturn = False Then
			Fn_SchMgr_TaskAssignment = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select  Task  [" + sTaskName + "]" )
			Set objTaskAssign = Nothing
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Task  [" + sTaskName + "]")
		End If
	
		'Select menu  [Schedule -> Assignments -> Assign to Task...]               
		bReturn = Fn_MenuOperation("WinMenuSelect","Schedule:Assignments:Assign to Task...")	  	   
		If bReturn = False Then
			Fn_SchMgr_TaskAssignment = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule -> Assignments -> Assign to Task....]")
			Set objTaskAssign = Nothing
			Exit Function
		Else
			Call Fn_ReadyStatusSync(2)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule -> Assignments -> Assign to Task...]")
		End If
	End If

	If objTaskAssign.Exist (SISW_MIN_TIMEOUT) Then
		Set objTreeTable = objTaskAssign.JavaTree("AssignmentTree")
		Set objHierarchyTree = objTaskAssign.JavaTree("HierarchyTree")
		
		Select Case sAction
			Case "Assign", "AssignVerifySearch"
				If IsArray(aUser) Then
					For iCounter = 0 to Ubound(aUser)
						aUserInfo  = Split( aUser(iCounter), ":", -1,1)                                                         '' split  Organization:Engg:Designer:AutoTest1 (autotest1)
						sSearchText = trim(aUserInfo(UBound(aUserInfo)))                   ''Type  AutoTest1 (autotest1)"  OR GrpName Or RoleName or Disc Name
						objTaskAssign.JavaTab("HierarchyTab").Select aUserInfo(0)                  
						objTaskAssign.JavaEdit("Filter").Set  sSearchText     

						If instr(sTaskName,",")>0 Then
							aTaskNames=split(sTaskName,",",-1,1)
							For jCount=0 to Ubound(aTaskNames)
								'[TC1123(20161205c00)_PoonamC_NewDevelopment_08Mar2017 : Added Code as per discussion with Koustubh to select Task from Task Assignment Dialog]
								If Fn_UI_JavaTree_NodeExist("Fn_SchMgr_TaskAssignment",objTreeTable,aTaskNames(jCount)) Then
									objTreeTable.ExtendSelect aTaskNames(jCount)
								Else
									arrTask = split(aTaskNames(jCount),":")
									sTask = arrTask(ubound(arrTask)-1) & ":" & arrTask(ubound(arrTask))
									objTreeTable.ExtendSelect sTask
									'objTreeTable.ExtendSelect aTaskNames(jCount)
								End If	
							Next
						Else
							'to Select to Item/ Reset focus to Tree (Added By Ikhlaque - 10 Aug)
							'Modified by Pritam Shikare 
							Dim  jCnt,sFirstNode 
							If sTaskName<>""  Then
								If InStr(1, sTaskName,":") Then
									aTaskName = split(sTaskName,":",-1,1)
									sFirstNode = objTreeTable.GetItem(0)
									sTreeHierarchy =  ""
									For iCnt = 0 to UBound(aTaskName)
										If  sFirstNode = aTaskName(iCnt) Then
											sTreeHierarchy = aTaskName(iCnt)
											Exit For
										End If
									Next
									If  sTreeHierarchy="" Then
										Fn_SchMgr_TaskAssignment = False
										Exit Function
									End If
									For jCnt = iCnt+1 to Ubound(aTaskName)
										sTreeHierarchy = sTreeHierarchy+":"+aTaskName(jCnt)
									Next
									objTreeTable.Select sTreeHierarchy
								End If
							End If
						End if
						wait 2	
'							Added to point to tree on selected Tab
					    objHierarchyTree.SetTOProperty "index", "1"
						objHierarchyTree.SetTOProperty "displayed", "1"

						If objHierarchyTree.GetROProperty("items count") <> 0 Then  											
							objHierarchyTree.Select aUser(iCounter)
                            If IsObject(sAssignBehaviour) <> False Then
								If sAssignBehaviour("sAssignBehaviour") <> "" Then
									wait 2
									objTaskAssign.JavaCheckBox("PlaceholderAssignment").Set sAssignBehaviour("sAssignBehaviour")
								End If
							End If
							objTaskAssign.JavaButton("Add").WaitProperty  "enabled", 1, 20000
							objTaskAssign.JavaButton("Add").Click	
							'Handle Warning while trying to assign already assigned user 
							If objTaskAssign.JavaWindow("Warning").Exist(SISW_MIN_TIMEOUT) Then
								objTaskAssign.JavaWindow("Warning").JavaButton("OK").Click
								Fn_SchMgr_TaskAssignment = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning Message while assigning resources")
                            	objTaskAssign.JavaButton("Cancel").WaitProperty  "enabled", 1, 5000
								objTaskAssign.JavaButton("Cancel").Click
								Set objTaskAssign = Nothing
								Set objTreeTable = Nothing
								Set objHierarchyTree = Nothing
								Exit Function
							End If
						Else
							Fn_SchMgr_TaskAssignment = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to assign member to task")
							objTaskAssign.JavaButton("Cancel").WaitProperty  "enabled", 1, 5000
							objTaskAssign.JavaButton("Cancel").Click
							Set objTaskAssign = Nothing
							Set objTreeTable = Nothing
							Set objHierarchyTree = Nothing
							Exit Function
						End If  
					Next
				End If
				If sAction = "AssignVerifySearch" Then						' [TC1122-20151116d00-12_01_2016-VivekA-NewDevelopment]- Case to Assign Verify User under Search Tab 
					If instr(sTaskName,",")>0 Then
						aTaskNames=split(sTaskName,",",-1,1)
					 	For jCount=0 to Ubound(aTaskNames)
							objTreeTable.ExtendSelect  aTaskNames(jCount)
					 	Next
					Else
						If sTaskName<>""  Then
							If InStr(1, sTaskName,":") Then
								aTaskName = split(sTaskName,":",-1,1)
								sFirstNode = objTreeTable.GetItem(0)
								sTreeHierarchy =  ""
								For iCnt = 0 to UBound(aTaskName)
									If  sFirstNode = aTaskName(iCnt) Then
										sTreeHierarchy = aTaskName(iCnt)
										Exit For
									End If
								Next
								If  sTreeHierarchy="" Then
									Fn_SchMgr_TaskAssignment = False
									Exit Function
								End If
								For jCnt = iCnt+1 to Ubound(aTaskName)
									sTreeHierarchy = sTreeHierarchy+":"+aTaskName(jCnt)
								Next
								sUser = Join(aUser)
								aUserInfo = Split(sUser,":")
								If UBound(aUserInfo) = 2 Then
									sUser = aUserInfo(UBound(aUserInfo) - 1) &"/"& aUserInfo(UBound(aUserInfo)) 
								Else
									sUser = aUserInfo(UBound(aUserInfo) - 1) &"."& aUserInfo(UBound(aUserInfo) - 2) &"/"& aUserInfo(UBound(aUserInfo)) 									
								End If
								objTreeTable.Select sTreeHierarchy & ":" & sUser
							End If
						End If
					End if
					wait 2				
					If IsObject(sAssignBehaviour) <> False Then
						If sAssignBehaviour("sTab") <> "" Then
							JavaWindow("ScheduleManagerWindow").JavaWindow("Task Assignment").JavaTab("HierarchyTab").Click 1,1,"LEFT"
							sTabCount = objTaskAssign.JavaTab("HierarchyTab").object.getItemCount()
							For iCount = 0 To sTabCount-1
								sTab = objTaskAssign.JavaTab("HierarchyTab").object.getItem(iCount).tostring()
								Set WshShell = CreateObject("WScript.Shell")
								WshShell.SendKeys "{RIGHT}"
								wait(1)
								Set WshShell = nothing
								If Instr(sTab,sAssignBehaviour("sTab")) Then
									Exit For
								End If
							Next
							Wait 1
						End If
						wait 1
						Call Fn_SISW_UI_JavaButton_Operations("Fn_SchMgr_TaskAssignment", "Click", objTaskAssign.JavaButton("FindResources"),"")
'						objTaskAssign.JavaButton("FindResources").Click 
						Wait 1
						If sAssignBehaviour("sUser") <> "" Then
							For iCounter = 0 To objTaskAssign.JavaTable("UsersTable").GetROProperty("rows") - 1
								If objTaskAssign.JavaTable("UsersTable").GetCellData(iCounter , "User") = sAssignBehaviour("sUser") Then
									objTaskAssign.JavaTable("UsersTable").SelectCell iCounter,"User"
									Wait 1
									Exit For 
								End If
							Next
						End If
					End If	
					Wait 1
					objTaskAssign.JavaButton("Add").Object.click
					If sAssignBehaviour("sUserVerify") <> "" Then
						jCount = objTreeTable.GetROProperty( "items count")
						For iCnt=0 To (jCount-1)
							aUserInfo = objTreeTable.GetItem(iCnt)
							If Trim(Lcase(aUserInfo)) = Trim(Lcase(sAssignBehaviour("sUserVerify"))) Then
								Fn_SchMgr_TaskAssignment = True
								Exit For
							End If
						Next
						If  Cint(iCnt) = Cint(jCount) Then
							objTaskAssign.JavaButton("OK").Click
							Fn_SchMgr_TaskAssignment = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify user is assign to " + sAssignBehaviour("sUserVerify") + " Assignment Tree." )	
							Set objTreeTable = Nothing
							Exit Function 
						End If
					End If					
				End If	
				wait 2
				objTaskAssign.JavaButton("OK").Click				
				If Err.Number < 0 Then
					Fn_SchMgr_TaskAssignment = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to assign member to task")
					objTaskAssign.JavaButton("Cancel").WaitProperty  "enabled", 1, 5000
					objTaskAssign.JavaButton("Cancel").Click
					Set objTaskAssign = Nothing
					Set objTreeTable = Nothing
					Set objHierarchyTree = Nothing
					Exit Function
				Else
					Fn_SchMgr_TaskAssignment = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully assigned member to task")
				End If

				If objTaskAssign.JavaDialog("Information").exist(SISW_MIN_TIMEOUT) then  'Added on 16 June 11
					objTaskAssign.JavaDialog("Information").JavaButton("OK").Click
					Fn_SchMgr_TaskAssignment = False
				End if 
					
			Case "Remove"				
				If IsArray(aUser) Then
					For iCounter = 0 to Ubound(aUser)		
                        aUserInfo  = Split( aUser(iCounter), ":", -1,1)                                                         '' split  Organization:Engg:Designer:AutoTest1 (autotest1)
						sBuildPath =  aTaskName(iOuterCount) &  ":" & aUserInfo(UBound(aUserInfo))      '"PPMSchedule_81619:T1:AutoTest2 (autotest2)"
						objTreeTable.Select sBuildPath                                                                                  
						objTaskAssign.JavaButton("Remove").WaitProperty  "enabled", 1, 20000
						objTaskAssign.JavaButton("Remove").Click
					Next
				End If		
				objTaskAssign.JavaButton("OK").Click	
				If Err.Number < 0 Then
					Fn_SchMgr_TaskAssignment = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Remove member From task")
					objTaskAssign.JavaButton("Cancel").WaitProperty  "enabled", 1, 5000
					objTaskAssign.JavaButton("Cancel").Click
					Set objTaskAssign = Nothing
					Set objTreeTable = Nothing
					Set objHierarchyTree = Nothing
					Exit Function
				Else
					Fn_SchMgr_TaskAssignment = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Removed member from task")
                End If

			Case "Modify"
				For iCounter = 0 to Ubound(aUser)
					If IsArray(aResourceLevel) Then
						aUserInfo  = Split( aUser(iCounter), ":", -1,1)                                                         '' split  Organization:Engg:Designer:AutoTest1 (autotest1)
						sBuildPath =  sTaskName &  ":" & aUserInfo(UBound(aUserInfo))                                  '"PPMSchedule_81619:T1:AutoTest2 (autotest2)"				
									
						''Get Exact coordinates in Resource Level cell to click
						sHeight=  objTreeTable.Object.getItemHeight()
						x = 0
						For iCnt = 0 to cInt(objTreeTable.GetROProperty ("columns_count")) - 1
							 If objTreeTable.GetColumnHeader(iCnt) = "Load" Then
								x = x + (cInt(objTreeTable.Object.getColumn(iCnt).getWidth()) / 4)
								Exit for
							Else
								x = x + cInt(objTreeTable.Object.getColumn(iCnt).getWidth())
							End If
						Next
						IntRows = objTreeTable.GetROProperty("items count")
						For iCnt = 0  to IntRows-1
								If  trim(objTreeTable.GetItem(iCnt)) = sBuildPath Then
									Exit For
								End If
						Next
						y  =  (iCnt +0.5) * CInt(sHeight)
						objTreeTable.Click x, y,"LEFT"
						''Enter Resouce Value after removing '%'  if any
						 aResourceValue = Split(CStr(aResourceLevel(iCounter)), "%",-1,1)
						objTreeTable.Type aResourceValue(0)
						objTreeTable.Click 1,1, "LEFT"
					End If
				Next   '' For each User
				objTaskAssign.JavaButton("OK").Click

				If Err.Number < 0 Then
					Fn_SchMgr_TaskAssignment = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modified the value of Resource Level")
					objTaskAssign.JavaButton("Cancel").WaitProperty  "enabled", 1, 5000
					objTaskAssign.JavaButton("Cancel").Click
					Set objTaskAssign = Nothing
					Set objTreeTable = Nothing
					Set objHierarchyTree = Nothing
					Exit Function
				Else
					Fn_SchMgr_TaskAssignment = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully modified the value of Resource Level ")					
				End If

			Case  "Verify"
				For iCounter = 0 to Ubound(aUser)
					aUserInfo  = Split( aUser(iCounter), ":", -1,1)                                                         '' split  Organization:Engg:Designer:AutoTest1 (autotest1)
					If aUserInfo(0) = "Organization" Then
						If Ubound(aUserInfo) = 1 Then                                                                              ''   "Organization:system" or "Organization:Engineering"
							sBuildPath =  sTaskName &  ":" & aUserInfo(UBound(aUserInfo)) & "/*"              '"PPMSchedule_81619:T1:Engineering/*"   OR    '"PPMSchedule_81619:T1:system/*" 
						ElseIf  Ubound(aUserInfo) = 2 Then                                                                              ''   "Organization:Engineering:Designer"
							sBuildPath =  sTaskName &  ":" & aUserInfo(UBound(aUserInfo)-1) & "/" & aUserInfo(UBound(aUserInfo))      '"PPMSchedule_81619:T1:Engineering/Designer" 
						Else
							sBuildPath =  sTaskName &  ":" & aUserInfo(UBound(aUserInfo))                                  '"PPMSchedule_81619:T1:AutoTest2 (autotest2)"				
							'[TC1123(20161205c00)_PoonamC_NewDevelopment_08Mar2017 : Added Code as per discussion with Koustubh to select Task from Task Assignment Dialog]
							 arrTask = split(sBuildPath,":")
							 sBuildPath = arrTask(ubound(arrTask)-2) & ":" & arrTask(ubound(arrTask)-1) & ":" & arrTask(ubound(arrTask)) 
						End If
					Else
						sBuildPath =  sTaskName &  ":" & aUserInfo(UBound(aUserInfo))                                  '"PPMSchedule_81619:T1:AutoDisp1"				
						'[TC1123(20161205c00)_PoonamC_NewDevelopment_08Mar2017 : Added Code as per discussion with Koustubh to select Task from Task Assignment Dialog]
						arrTask = split(sBuildPath,":")
						sBuildPath = arrTask(ubound(arrTask)-2) & ":" & arrTask(ubound(arrTask)-1) & ":" & arrTask(ubound(arrTask))
					End If

					If IsArray(aResourceLevel) Then	   '' If Resource Level is given verify cell value 
						If Instr(aResourceLevel(iCounter),".") = 0 Then
							aResourceLevel(iCounter) = Left(aResourceLevel(iCounter), Len(aResourceLevel(iCounter))-1) & ".0%"
						End If
						If objTreeTable.GetColumnValue(sBuildPath, "Resource Level") = aResourceLevel(iCounter) Then
							Fn_SchMgr_TaskAssignment = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully verified the member with resouce level.")
						Else
							Fn_SchMgr_TaskAssignment = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the member with resource level.")
							objTaskAssign.JavaButton("Cancel").WaitProperty  "enabled", 1, 5000
							objTaskAssign.JavaButton("Cancel").Click
							Set objTaskAssign = Nothing
							Set objTreeTable = Nothing
							Set objHierarchyTree = Nothing
							Exit Function
						End If
					Else                                                       '' If Resource Level not given, Verify just the User assigned or not
						objTreeTable.Select sBuildPath
						If Err.Number < 0 Then
							Fn_SchMgr_TaskAssignment = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the member with resource level.")
							objTaskAssign.JavaButton("Cancel").WaitProperty  "enabled", 1, 5000
							objTaskAssign.JavaButton("Cancel").Click
							Set objTaskAssign = Nothing
							Set objTreeTable = Nothing
							Set objHierarchyTree = Nothing
							Exit Function
						Else
							Fn_SchMgr_TaskAssignment = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully verified the member.")
						End If
					End If
				Next   '' For each User
				objTaskAssign.JavaButton("Cancel").Click
				
			'[TC1122-20151116d-31_12_2015-VivekA-NewDevelopment]- Case to work on Qualification Tab 
			Case "QualificationTabOperation"
				If sAssignBehaviour<>"" Then
					sButton = sAssignBehaviour
				Else
					sButton = ""
				End If
				'Select Task
				If Instr(sTaskName,",")>0 Then
					aTaskNames = Split(sTaskName,",",-1,1)
				 	For jCount=0 To UBound(aTaskNames)
						objTreeTable.ExtendSelect aTaskNames(jCount)
				 	Next
				 	Wait 1
				End If
				'Select Qualifications tab
				objTaskAssign.JavaTab("HierarchyTab").Select "Qualifications"
				Wait 1
				
				If IsArray(aUser) Then
					Select Case aUser(0)
						Case "AssignQualification"
							For iCounter = 1 To UBound(aUser)
								aQualifications = Split(aUser(iCounter),":")
								sQualification = aQualifications(0)
								sQualiLevel = aQualifications(1)
								For iCnt = 0 To 1
									If iCnt = 0 Then
										sName = "Qualification:"
										sSelectValue = sQualification
									Else
										sName = "Qualification Level:"
										sSelectValue = sQualiLevel
									End If
									
									objTaskAssign.JavaStaticText("QualificationText").SetTOProperty "label",sName
									'Click on dropdown button
									Call Fn_Button_Click("Fn_SchMgr_TaskAssignment", objTaskAssign, "QualificationButton")
									Wait 2
									Set objChild = objTaskAssign.JavaWindow("QualificationShell").JavaTable("QualificationTableList")
									iRowCount = objChild.GetROProperty("rows")
									bFlag = False
									For iCount = 0 To iRowCount- 1
										'Check existance of Qualification from list
										If Trim(objChild.getCellData(iCount,0)) = Trim(sSelectValue) Then
											'Select Qualification/Qualification Level from list
											objChild.SelectCell iCount,0
											Wait 1
											bFlag = True
											Exit for
										End If
									Next
									If bFlag = False Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SchMgr_TaskAssignment ] Failed to select "&sName&" [ "+sSelectValue+" ].")
										Fn_SchMgr_TaskAssignment = False
										Set objTaskAssign = Nothing
										Set objChild = Nothing
										Exit Function
									End If
									'Call Fn_Button_Click("Fn_SchMgr_TaskAssignment", objTaskAssign, "QualificationButton")
									Set objChild = Nothing
									Wait 2
								Next
								'Click on QualificationAssign button
								bFlag = Fn_Button_Click("Fn_SchMgr_TaskAssignment",objTaskAssign,"QualificationAssign")
								If bFlag=False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SchMgr_TaskAssignment ] Failed to Click on [ QualificationAssign ] button.")
									Fn_SchMgr_TaskAssignment = False
									Set objTaskAssign = Nothing
									Exit Function
								End If
								Wait 1
							Next
						'[TC1122-20151116d-08_01_2016-VivekA-NewDevelopment] - Added by Shweta Rathod to handle Qualification Error Dialog
						Case "QualificationErrorVerify"
							Set objErr = objTaskAssign.JavaWindow("Warning")
							objErr.SetTOProperty "title","Qualifications"
							sButton = "OK"
							If objErr.Exist(SISW_MICRO_Timeout) Then
								objErr.JavaStaticText("ErrorMsg").SetTOProperty "label",aUser(1)
								If objErr.JavaStaticText("ErrorMsg").Exist(SISW_MICRO_Timeout) Then
									Fn_SchMgr_TaskAssignment = True
									objErr.JavaButton("OK").Click micLeftBtn
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Failed to Verify Message : "+ sErrorMsg)
									Fn_SchMgr_TaskAssignment = False
									objErr.JavaButton("OK").Click micLeftBtn
									Set objTaskAssign = Nothing
									Set objErr = Nothing
									Exit Function
								End If
							Else																												
								Fn_SchMgr_TaskAssignment = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Error dialog not found")
							End If
							Set objErr = Nothing
						Case "RemoveQualification"
							For iCounter = 1 To UBound(aUser)
								aQualifications = Split(aUser(iCounter),":")
								sQualification = aQualifications(0)
								sQualiLevel = aQualifications(1)
								bFlag = False
								iRowCount = objTaskAssign.JavaTable("ExistingQualificationsTable").GetROProperty("rows")
								For iCount = 0 To iRowCount-1
									If LCase(Trim(objTaskAssign.JavaTable("ExistingQualificationsTable").GetCellData(iCount,"Qualification"))) = LCase(Trim(sQualification)) Then
										If LCase(Trim(objTaskAssign.JavaTable("ExistingQualificationsTable").GetCellData(iCount,"Level"))) = LCase(Trim(sQualiLevel)) Then
											objTaskAssign.JavaTable("ExistingQualificationsTable").SelectCell iCount,"Qualification"
											Wait 1
											'Click on QualificationRemove button
											bFlag = Fn_Button_Click("Fn_SchMgr_TaskAssignment",objTaskAssign,"QualificationRemove")
											If bFlag=False Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SchMgr_TaskAssignment ] Failed to Click on [ QualificationRemove ] button.")
												Fn_SchMgr_TaskAssignment = False
												Set objTaskAssign = Nothing
												Exit Function
											End If
											Wait 1
											bFlag = True
											Exit For
										End If
									End If
								Next
								If bFlag=False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SchMgr_TaskAssignment ] Fail to Remove Qualification [ "+sQualification+" ] in [ ExistingQualificationsTable ].")
									Fn_SchMgr_TaskAssignment = False
									Set objTaskAssign = Nothing
									Exit Function
								End If	
							Next

						Case "VerifyInExistingQualificationsTable"
							For iCounter = 1 To UBound(aUser)
								aQualifications = Split(aUser(iCounter),":")
								sQualification = aQualifications(0)
								sQualiLevel = aQualifications(1)
								bFlag = False
								iRowCount = objTaskAssign.JavaTable("ExistingQualificationsTable").GetROProperty("rows")
								For iCount = 0 To iRowCount-1
									If LCase(Trim(objTaskAssign.JavaTable("ExistingQualificationsTable").GetCellData(iCount,"Qualification"))) = LCase(Trim(sQualification)) Then
										If LCase(Trim(objTaskAssign.JavaTable("ExistingQualificationsTable").GetCellData(iCount,"Level"))) = LCase(Trim(sQualiLevel)) Then
											bFlag = True
											Exit For
										End If
									End If
								Next
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SchMgr_TaskAssignment ] Fail to Verify Qualification [ "+sQualification+" ] in [ ExistingQualificationsTable ].")
									Fn_SchMgr_TaskAssignment = False
									Set objTaskAssign = Nothing
									Exit Function
								End If
							Next
							
						Case "VerifyInQualificationList"
							objTaskAssign.JavaStaticText("QualificationText").SetTOProperty "label","Qualification:"
							'Click on dropdown button
							Call Fn_Button_Click("Fn_SchMgr_TaskAssignment", objTaskAssign, "QualificationButton")
							Wait 2
							For iCounter = 1 To UBound(aUser)
								sQualification = aUser(iCounter)
								bFlag = False
								Set objChild = objTaskAssign.JavaWindow("QualificationShell").JavaTable("QualificationTableList")
								iRowCount = objChild.GetROProperty("rows")
								For iCount = 0 To iRowCount- 1
									'Check existance of Qualification from list
									If Trim(objChild.getCellData(iCount,0)) = Trim(sQualification) Then
										bFlag = True
										Exit for
									End If
								Next
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SchMgr_TaskAssignment ] Failed to Verify Qualification: [ "+sQualification+" ] in QualificationTableList.")
									Fn_SchMgr_TaskAssignment = False
									Set objTaskAssign = Nothing
									Set objChild = Nothing
									Exit Function
								End If
								Set objChild = Nothing
								Wait 1
							Next
							Call Fn_Button_Click("Fn_SchMgr_TaskAssignment", objTaskAssign, "QualificationButton")
							Wait 1
							If bFlag = False Then
								Fn_SchMgr_TaskAssignment = False
								Set objTaskAssign = Nothing
								Exit Function
							End If
							
						Case "VerifyInQualificationLevelList"
							'Call for Select Qualification list is needed
							objTaskAssign.JavaStaticText("QualificationText").SetTOProperty "label","Qualification Level:"
							'Click on dropdown button
							Call Fn_Button_Click("Fn_SchMgr_TaskAssignment", objTaskAssign, "QualificationButton")
							Wait 2
							For iCounter = 1 To UBound(aUser)
								sQualification = aUser(iCounter)
								bFlag = False
								Set objChild = objTaskAssign.JavaWindow("QualificationShell").JavaTable("QualificationTableList")
								iRowCount = objChild.GetROProperty("rows")
								For iCount = 0 To iRowCount- 1
									'Check existance of Qualification from list
									If Trim(objChild.getCellData(iCount,0)) = Trim(sQualification) Then
										bFlag = True
										Exit for
									End If
								Next
								If bFlag = False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SchMgr_TaskAssignment ] Failed to Verify Qualification Level: [ "+sQualification+" ] in QualificationTableList.")
									Fn_SchMgr_TaskAssignment = False
									Set objTaskAssign = Nothing
									Set objChild = Nothing
									Exit Function
								End If
								Set objChild = Nothing
								Wait 1
							Next
							Call Fn_Button_Click("Fn_SchMgr_TaskAssignment", objTaskAssign, "QualificationButton")
							Wait 1
							If bFlag = False Then
								Fn_SchMgr_TaskAssignment = False
								Set objTaskAssign = Nothing
								Exit Function
							End If
					End Select
				Else
					Fn_SchMgr_TaskAssignment = False
					Set objTaskAssign = Nothing
					Exit Function
				End If
				
				'Click on OK or Cancel if button is passed as parameter
				If sButton<>"" Then
					bFlag = Fn_Button_Click("Fn_SchMgr_TaskAssignment",objTaskAssign,sButton)
					If bFlag=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : [ Fn_SchMgr_TaskAssignment ] Failed to Click on [ "+sButton+" ] button.")
						Fn_SchMgr_TaskAssignment = False
						Set objTaskAssign = Nothing
						Exit Function
					End If
				End If
				Fn_SchMgr_TaskAssignment = True
	    End Select

		Set objTaskAssign = Nothing
		Set objTreeTable = Nothing
		Set objHierarchyTree = Nothing
	Else
		Fn_SchMgr_TaskAssignment = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task Assignment dialog does not exist.")
		Set objTaskAssign = Nothing
		Set objTreeTable = Nothing
		Set objHierarchyTree = Nothing
		Exit Function
	End If
End Function

'*********************************************  Function selects a task and indents it under the task above it**************************************************************

'Function Name		:					Fn_SchMgr_TskIndentOutdent  

'Description			 :		 		  The Function selects a task and indents it under the task above it..

'Parameters			   :	 			1.  sAction :Action need to perform Indent/Outdent
'													 2.sTaskName  :Name of the Task
'													The multiple ", (comma)" seperated values of the task names  For eg:- Sch1:Tsk1,Sch1:Tsk2
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:				Fn_SchMgr_TskIndentOutdent("Indent","Testch:t2")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali							03-Jun-2010	   		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_TskIndentOutdent(sAction,sTaskName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TskIndentOutdent"
	On Error Resume Next

	Dim bReturn, sCaption, objWin

	bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")

	If  bReturn <> False Then
		Fn_SchMgr_TskIndentOutdent = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
	ELse
		Fn_SchMgr_TskIndentOutdent = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
		Exit Function
	End If

	Select Case sAction
		Case "Indent"
			bReturn = Fn_MenuOperation("WinMenuSelect","Schedule:Indent task")	 
'			JavaWindow("ScheduleManagerWindow").JavaMenu("label:=Schedule","index:=0").JavaMenu("label:=Indent task","index:=0").Select
			'Handle Error if Exists
			bReturn = Fn_SchMgr_SchedulingErrorVerify("", "OK")
			If bReturn Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error on Indenting")
				Fn_SchMgr_TskIndentOutdent = False
			Else
				'Handle Warnign
				bReturn = Fn_SchMgr_WarningMsgVerify ("", "OK")
				If bReturn Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning on Indenting")
					Fn_SchMgr_TskIndentOutdent = False
				Else
					Fn_SchMgr_TskIndentOutdent = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected menu Schedule->Indent task")
				End If

				bReturn = Fn_SchMgr_DialogMsgVerify("Warning", sErrorText,"OK") ' Added by Omkar 5-May-2011
				If bReturn Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning on Indenting the Node of Schedule Table")
					Fn_SchMgr_TskIndentOutdent = FALSE
					Exit Function
				End If
			End If

'			Fn_SchMgr_TskIndentOutdent = True
'			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected menu Schedule->Indent task")


'			Call Fn_ReadyStatusSync(1) 
'				
'			If bReturn <> False Then
'				Fn_SchMgr_TskIndentOutdent = True
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected menu Schedule->Indent task")
'			Else
'				Fn_SchMgr_TskIndentOutdent = False
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select menu Schedule->Indent task")
'				Exit Function
'			End If

		Case "Outdent"

			JavaWindow("ScheduleManagerWindow").JavaMenu("label:=Schedule","index:=0").JavaMenu("label:=Outdent task","index:=0").Select
			'Handle Warnign
			bReturn = Fn_SchMgr_WarningMsgVerify ("", "OK")
			If bReturn Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Warning on Indenting")
				Fn_SchMgr_TskIndentOutdent = False
			Else
				Fn_SchMgr_TskIndentOutdent = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected menu Schedule->Indent task")
			End If
'			bReturn = Fn_MenuOperation("Select","Schedule:Outdent task")
'				
'			Call Fn_ReadyStatusSync(1) 
'			 	
'			If bReturn <> False Then
'				Fn_SchMgr_TskIndentOutdent = True
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected menu Schedule->Outdent task")
'			Else
'				Fn_SchMgr_TskIndentOutdent = False
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select menu Schedule->Outdent task")
'				Exit Function
'			End If
	End Select

End Function

'*********************************************  Functions performs a cut action on the specified node(s).**************************************************************

'Function Name		:					Fn_SchMgr_CutNode  

'Description			 :		 		  The Functions performs a cut action on the specified node(s).

'Parameters			   :	 			1.  sAction :Action need to perform Menu/RMB/ToolBar/Shortcut.
'													 2.sTaskName  :The name of the Task(s) to be cut.
'													The multiple ", (comma)" seperated values of the task names  For eg:- Sch1:Tsk1,Sch1:Tsk2
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .Task/Schedule should be created.

'Examples				:				Fn_SchMgr_CutNode("Menu","Testch:t2")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali							03-Jun-2010	   		1.0
'											Prasanna					23-Jul-2010	   		  1.0             Handled Add Resources Dialog 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_CutNode(sAction,sTaskName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_CutNode"
   On Error Resume Next

	Dim bReturn

	Select Case sAction

		Case "Menu"
			If Instr(sTaskName, ",") > 0 Then
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
			Else
				bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
			End If

			If  bReturn <> False Then
				Fn_SchMgr_CutNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
			ELse
				Fn_SchMgr_CutNode = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
				Exit Function
			End If
			
			Wait(1)
			bReturn = Fn_MenuOperation("Select","Edit:Cut")
			If bReturn = True Then
				Fn_SchMgr_CutNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Edit;Cut]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Edit;Cut]")
				Fn_SchMgr_CutNode = False
				Exit Function
			End If

		Case "Toolbar" 
			If Instr(sTaskName, ",") > 0 Then
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
			Else
				bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
			End If

			If  bReturn <> False Then
				Fn_SchMgr_CutNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
			ELse
				Fn_SchMgr_CutNode = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
				Exit Function
			End If

			bReturn= Fn_ToolbatButtonClick("Cut the selection and put it on the clipboard (Ctrl+X)")
			If bReturn=True Then
				Fn_SchMgr_CutNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Toolbar Button [Cut (Ctrl + X)]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Toolbar Button [Cut (Ctrl + X)]")
				Fn_SchMgr_CutNode = False
				Exit Function
			End If

			Case "RMB"
				bReturn =  Fn_SchMgr_SchTable_NodeOperation("PopupMenu", sTaskName, "", "", "Cut	Ctrl+X")
				If bReturn = True Then
					Fn_SchMgr_CutNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked RMB Menu [Cut]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked RMB Menu [Cut]")
					Fn_SchMgr_CutNode = False
					Exit Function
				End If

			Case "Shortcut" 
				If Instr(sTaskName, ",") > 0 Then
					bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
				Else
					bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
				End If

				If  bReturn <> False Then
					Fn_SchMgr_CutNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
				ELse
					Fn_SchMgr_CutNode = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
					Exit Function
				End If

				JavaWindow("ScheduleManagerWindow").Activate
                bReturn= Fn_KeyBoardOperation("SendKeys", "^{x 10}")
                bReturn= Fn_KeyBoardOperation("SendKeys", "^(x)")

				If bReturn=True Then
					Fn_SchMgr_CutNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully use shortcut key [ (Ctrl + X)]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to use shortcut key [ (Ctrl + X)]")
					Fn_SchMgr_CutNode = False
					Exit Function
				End If

	End Select
wait(3)
Call Fn_ReadyStatusSync(3)
End  Function


'*********************************************  Functions performs a copy action on the specified node(s).**************************************************************

'Function Name		:					Fn_SchMgr_CopyNode

'Description			 :		 		  The Functions performs a copy action on the specified node(s).

'Parameters			   :	 			1.  sAction :Action need to perform Menu/RMB/ToolBar/Shortcut.
'													 2.sTaskName  :The name of the Task(s) to be copy.
'													The multiple ", (comma)" seperated values of the task names  For eg:- Sch1:Tsk1,Sch1:Tsk2
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .Task/Schedule should be created.

'Examples				:				Fn_SchMgr_CopyNode("Menu","Testch:t2")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali							04-Jun-2010	   		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_CopyNode(sAction,sTaskName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_CopyNode"
   On Error Resume Next

	Dim bReturn

	Select Case sAction

		Case "Menu"
			If Instr(sTaskName, ",") > 0 Then
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
			Else
				bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
			End If

			If  bReturn <> False Then
				Fn_SchMgr_CopyNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
			ELse
				Fn_SchMgr_CopyNode = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
				Exit Function
			End If

			bReturn = Fn_MenuOperation("Select","Edit:Copy")
			If bReturn = True Then
                Fn_SchMgr_CopyNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Edit;Copy]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Edit;Copy]")
				Fn_SchMgr_CopyNode = False
				Exit Function
			End If

		Case "Toolbar" 
			If Instr(sTaskName, ",") > 0 Then
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
			Else
				bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
			End If

			If  bReturn <> False Then
				Fn_SchMgr_CopyNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
			ELse
				Fn_SchMgr_CopyNode = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
				Exit Function
			End If

			bReturn= Fn_ToolbatButtonClick("Copy (Ctrl+C)")
			If bReturn=True Then
                Fn_SchMgr_CopyNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Toolbar Button [Copy (Ctrl+C)]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Toolbar Button [Copy (Ctrl+C)]")
				Fn_SchMgr_CopyNode = False
				Exit Function
			End If

			Case "RMB"
			bReturn =  Fn_SchMgr_SchTable_NodeOperation("PopupMenu", sTaskName, "", "", "Copy	Ctrl+C")
			If bReturn = True Then
                    Fn_SchMgr_CopyNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked RMB Menu [Copy]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked RMB Menu [Copy]")
					Fn_SchMgr_CopyNode = False
					Exit Function
			End If 

			Case "Shortcut" 
				If Instr(sTaskName, ",") > 0 Then
					bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
				Else
					bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
				End If

				If  bReturn <> False Then
					Fn_SchMgr_CopyNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
				ELse
					Fn_SchMgr_CopyNode = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
					Exit Function
				End If

				bReturn= Fn_KeyBoardOperation("SendKeys", "^(c)")
				bReturn= Fn_KeyBoardOperation("SendKeys", "^(c)")
'				bReturn= Fn_KeyBoardOperation("PressKey", "c:micCtrl")
				If bReturn=True Then
					Fn_SchMgr_CopyNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully use shortcut key [ (Ctrl + C)]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to use shortcut key [ (Ctrl + C)]")
					Fn_SchMgr_CopyNode = False
					Exit Function
				End If

	End Select

	'Handle Copy Task Error
	JavaWindow("ScheduleManagerWindow").JavaWindow("Error").SetTOProperty "title", "Copy Task"
	Set objWin = JavaWindow("ScheduleManagerWindow").JavaWindow("Error")
	if objWin.Exist(SISW_MIN_TIMEOUT) Then
		objWin.JavaButton("OK").Click micLeftBtn	
		Set objWin = Nothing
		If Err.Number < 0 Then
				Fn_SchMgr_CopyNode = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Button [OK] on [Copy Task] Dialog")
		End If
	End If

End  Function  


'*********************************************  Functions performs a paste action on the specified node(s).**************************************************************

'Function Name		:					Fn_SchMgr_PasteNode

'Description			 :		 		  The Functions performs a paste action on the specified node(s).

'Parameters			   :	 			1.  sAction :Action need to perform Menu/RMB/ToolBar/Shortcut.
'													 2.sTaskName  :The name of the Task(s) to be paste.
'													The multiple ", (comma)" seperated values of the task names  For eg:- Sch1:Tsk1,Sch1:Tsk2
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .Task/Schedule should be created.

'Examples				:				Fn_SchMgr_PasteNode("Menu","Testch:t2")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali							04-Jun-2010	   		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_PasteNode(sAction,sTaskName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_PasteNode"
   On Error Resume Next

	Dim bReturn, aTskName, sBtnName, sCaption, sBtnName1
	
	if instr(sTaskName, "~") > 0 Then
		aTskName = split(sTaskName, "~", -1, 1)
		sTaskName = aTskName(0)
		sBtnName = aTskName(1)
	Else
		sBtnName = "Yes"
	End If

		if instr(sTaskName, "|") > 0 Then
			aTskName = split(sTaskName, "|", -1, 1)
			sTaskName = aTskName(0)
			sBtnName1 = aTskName(1)
		Else
			sBtnName1 = "Yes"
		End If

	Select Case sAction

		Case "Menu"
			If Instr(sTaskName, ",") > 0 Then
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
			Else
				bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
			End If

			If  bReturn <> False Then
				Fn_SchMgr_PasteNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
			ELse
				Fn_SchMgr_PasteNode = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
				Exit Function
			End If

			bReturn = Fn_MenuOperation("Select","Edit:Paste")
			If bReturn = True Then
                Fn_SchMgr_PasteNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Edit;Paste]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Edit;Paste]")
				Fn_SchMgr_PasteNode = False
				Exit Function
			End If

		Case "Toolbar" 
			If Instr(sTaskName, ",") > 0 Then
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
			Else
				bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
			End If

			If  bReturn <> False Then
				Fn_SchMgr_PasteNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
			ELse
				Fn_SchMgr_PasteNode = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
				Exit Function
			End If

			bReturn= Fn_ToolbatButtonClick("Paste a reference to the object contained in the clipboard (Ctrl+V)")
			If bReturn=True Then
                Fn_SchMgr_PasteNode = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Toolbar Button [Paste (Ctrl+V)]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Toolbar Button [Paste (Ctrl+V)]")
				Fn_SchMgr_PasteNode = False
				Exit Function
			End If

			Case "RMB"
			bReturn =  Fn_SchMgr_SchTable_NodeOperation("PopupMenu", sTaskName, "", "", "Paste	Ctrl+V")
			If bReturn = True Then
                    Fn_SchMgr_PasteNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked RMB Menu [Paste]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked RMB Menu [Paste]")
					Fn_SchMgr_PasteNode = False
					Exit Function
			End If

			Case "Shortcut" 
				If Instr(sTaskName, ",") > 0 Then
					bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
				Else
					bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTaskName,"","","")
				End If

				If  bReturn <> False Then
					Fn_SchMgr_PasteNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
				ELse
					Fn_SchMgr_PasteNode = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
					Exit Function
				End If

				bReturn= Fn_KeyBoardOperation("SendKeys", "^(v)")
'				bReturn= Fn_KeyBoardOperation("PressKey", "v:micCtrl")
				If bReturn=True Then
					Fn_SchMgr_PasteNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully use shortcut key [ (Ctrl + V)]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to use shortcut key [ (Ctrl + V)]")
					Fn_SchMgr_PasteNode = False
					Exit Function
				End If

	End Select
   
	sCaption = "Warning"
	JavaDialog("Confirmation").SetTOProperty "title", sCaption
	If JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
		JavaDialog("Confirmation").JavaButton(sBtnName).click micLeftBtn
	End If
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Handle Paste Warning")
		Fn_SchMgr_PasteNode = False
		Exit Function
	End If

	If JavaDialog("Add resources to schedule?").Exist(5) Then
		JavaDialog("Add resources to schedule?").JavaButton(sBtnName1).Click micLeftBtn		
		If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Button [" + sBtnName1 + "] on [Add Resources] Dialog")
					Fn_SchMgr_PasteNode = False
					Exit Function
				Else					
					Fn_SchMgr_PasteNode = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked Button [" + sBtnName1 + "] on [Add Resources] Dialog")
				End If
  	 End If
	
	sCaption = "Scheduling Error"
	bReturn = Fn_SchMgr_DialogMsgVerify(sCaption, sErrorText,"OK")
	If bReturn Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error on Pasting.....")
		Fn_SchMgr_PasteNode = False
		sErrorText = ""
		Exit Function
	End If
Call Fn_ReadyStatusSync(5)
End  Function     

'**********************************************  Function perform all the actions related to column management 		********************************************************

'Function Name		:					Fn_SchMgr_ColumnOperations 

'Description			 :		 		  The Function perform all the actions related to column management.

'Parameters			   :	 			1.  sAction: Action need to perform. (Add/Remove/Verify)
'													2.aColName : The name of the column(s) to be added/removed. 
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		 Schedule table should be displayed.

'Examples				:				Fn_SchMgr_ColumnOperations("Remove",aColName)
'											  bReturn=Fn_SchMgr_ColumnOperations("MoveUp&Verify",aColName)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							06-Jun-2010   			1.0
'										SHREYAS						07-04-2011				1.1				added 2 cases			Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ColumnOperations(sAction,aColName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ColumnOperations"
   On Error Resume Next 
'	Dim sIndex,iCounter,bReturn,objTable,objColChooser,sIndex2,sArrayAction,sCloseStatus

	Set objTable = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable")
	Set objColChooser = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Column Chooser")
	
	'Select Menu Column Chooser
	If Not objColChooser.Exist(SISW_MIN_TIMEOUT) Then
		objTable.SelectColumnHeader "Object", "RIGHT"
		JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaMenu("label:=Column Chooser","index:=0").Select
	End If

	If objColChooser.Exist(SISW_MIN_TIMEOUT) Then

		If instr(1,sAction,":")>0 Then
		sArrayAction=Split(sAction,":")
		sAction=sArrayAction(0)
		sCloseStatus=sArrayAction(1)
		End If

		Select Case sAction
	
			Case "Add"
				If IsArray(aColName) Then
					For iCounter = 0 to Ubound(aColName)
		
						bReturn = Fn_SchMgr_TableColIndex(objTable,aColName(iCounter))
		
						If cBool(bReturn) = False Then
							'Select column from available columns.
							objColChooser.JavaList("AvailableColumns").ExtendSelect aColName(iCounter)
							If Err.Number < 0 Then
								Fn_SchMgr_ColumnOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to select  " &aColName(iCounter)& " from available columns.")
								objColChooser.JavaButton("Close").Click
								Set objTable = Nothing
								Set objColChooser = Nothing
								Exit Function 
							End If
						Else
							Fn_SchMgr_ColumnOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Already  "  & aColName(iCounter) & "column exists in schedule table.")
							objColChooser.JavaButton("Close").click
							Exit Function
						End If
		
					Next
					'Click Add button
					objColChooser.JavaButton("AddCol").WaitProperty "enabled",1,20000
					objColChooser.JavaButton("AddCol").Click
					'Click Apply button
					objColChooser.JavaButton("OK").WaitProperty "enabled",1,20000
					objColChooser.JavaButton("OK").Click
		
					Fn_SchMgr_ColumnOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully added columns to schedule table .")
		
					If Err.Number < 0 Then
						Fn_SchMgr_ColumnOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to add columns to schedule table..")
						objColChooser.JavaButton("Close").click
						Set objTable = Nothing
						Set objColChooser = Nothing
						Exit Function 
					End If
				End If
	
			Case "Remove"
				If IsArray(aColName) Then
					For iCounter = 0 to Ubound(aColName)
						bReturn = Fn_SchMgr_TableColIndex(objTable,aColName(iCounter))

						If  bReturn <> False Then
							'Select column from displayed columns.
							objColChooser.JavaList("DisplayedColumns").ExtendSelect aColName(iCounter)
							If Err.Number < 0 Then
								Fn_SchMgr_ColumnOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to select  " &aColName(iCounter)& " from displayed columns.")
								objColChooser.JavaButton("Close").click
								Set objTable = Nothing
								Set objColChooser = Nothing
								Exit Function 
							End If
						Else 
							Fn_SchMgr_ColumnOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),aColName(iCounter) & " column does not exist in schedule table. ")
							objColChooser.JavaButton("Close").click
							Set objTable = Nothing
							Set objColChooser = Nothing
							Exit Function 
						End If
					Next
					'Click Remove button
					objColChooser.JavaButton("RemoveCol").WaitProperty "enabled",1,20000
					objColChooser.JavaButton("RemoveCol").Click
					'Click Apply button
					objColChooser.JavaButton("OK").WaitProperty "enabled",1,20000
					objColChooser.JavaButton("OK").Click
		
					Fn_SchMgr_ColumnOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove columns from schedule table .")
		
					If Err.Number < 0 Then
						Fn_SchMgr_ColumnOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to remove column from schedule table.")
						Set objTable = Nothing
						Set objColChooser = Nothing
						Exit Function 
					End If
				End If

			Case "Verify"
				If IsArray(aColName) Then
					For iCounter = 0 to Ubound(aColName)
						bReturn = Fn_SchMgr_TableColIndex(objTable,aColName(iCounter))
						If bReturn <> False Then
							Fn_SchMgr_ColumnOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColName(iCounter) &" column exist in schedule table.")
						Else  
							Fn_SchMgr_ColumnOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColName(iCounter) &" column does not exist in schedule table.")
							objColChooser.JavaButton("Close").click
							Set objTable = Nothing
							Set objColChooser = Nothing
							Exit Function 
						End If
					Next
					objColChooser.JavaButton("Close").click
				End If 

		Case "MoveUp&Verify"

					If IsArray(aColName) Then
					For iCounter = 0 to Ubound(aColName)
						
						sIndex= objColChooser.JavaList("DisplayedColumns").GetItemIndex(aColName(iCounter))

						'select the node
						objColChooser.JavaList("DisplayedColumns").Select aColName(iCounter)

						If objColChooser.JavaButton("MoveUp").GetROProperty("enabled")="1" Then
								objColChooser.JavaButton("MoveUp").Click micLeftBtn
								sIndex2=objColChooser.JavaList("DisplayedColumns").GetItemIndex(aColName(iCounter))
								If sIndex2=sIndex-1 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColName(iCounter) &" is successfully verified as moved up")
									Fn_SchMgr_ColumnOperations = True
								Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColName(iCounter) &" is not verified as moved up")
									Fn_SchMgr_ColumnOperations = False
									Exit function
								
								End If
								End if
					Next
								If sCloseStatus="No" Then
								 'do not close the dialog
							Else
								objColChooser.close
							 End If
						End If

		Case "MoveDown&Verify"

					If IsArray(aColName) Then
					For iCounter = 0 to Ubound(aColName)
						
						sIndex= objColChooser.JavaList("DisplayedColumns").GetItemIndex(aColName(iCounter))

						'select the node
						objColChooser.JavaList("DisplayedColumns").Select aColName(iCounter)

						If  objColChooser.JavaButton("MoveDown").GetROProperty("enabled")="1" Then
								objColChooser.JavaButton("MoveDown").Click micLeftBtn
								sIndex2=objColChooser.JavaList("DisplayedColumns").GetItemIndex(aColName(iCounter))
								If sIndex2=sIndex+1 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColName(iCounter) &" is successfully verified as moved down")
									Fn_SchMgr_ColumnOperations = True
								Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColName(iCounter) &" is not verified as moved down")
									Fn_SchMgr_ColumnOperations = False
									Exit function
								
								End If
								End if
					Next
							If sCloseStatus="No" Then
							 'do not close the dialog
						Else
							objColChooser.close
						 End If
						End If

		End Select
	End If
	Set objTable = Nothing
	Set objColChooser = Nothing
End Function

'*********************************************************  Function replaces user(s) with the specified user(s).************************************************************************************

'Function Name		:				   Fn_SchMgr_ReplaceTskAssignment 

'Description			 :		 		  The Function replaces user(s) with the specified user(s)

'Parameters			   :	 			sName: The name of the user(s) that has to be replaced. Full  path of name with : separated.Starts with Assignment Member:
'                                                   sReplace : The user(s) that replaces the selected user(s).Full  path of name with : separated.Starts with Schedule Member: 
'													sTaskName :Name of the task.

'Return Value		   : 			  True/False  

'Pre-requisite			:		 		Schedule manger panel need to be open.

'Examples				:				Fn_SchMgr_ReplaceTskAssignment ("Testch:t1","Assignment Member:Rupali Palhade (x_palhad)","Schedule Member:Amol Lanke (x_lanke)")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							05-Jun-2010	           1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ReplaceTskAssignment(sTaskName,sName,sReplace)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ReplaceTskAssignment"
   On Error Resume Next

    Dim bReturn,objReTaskAssign

	Set objReTaskAssign = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Replace Task Assignments")
	'Select  task from schedule table
    bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")

	If  bReturn <> False Then
		Fn_SchMgr_ReplaceTskAssignment = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
	ELse
		Fn_SchMgr_ReplaceTskAssignment = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
		Exit Function
	End If
	'Select menu Schedule->Assignments->Repalce Assignment....
	If Not objReTaskAssign.Exist(SISW_MIN_TIMEOUT)Then
		bReturn = Fn_MenuOperation("Select","Schedule:Assignments:Replace Assignment...")
		If bReturn = True Then
			Fn_SchMgr_ReplaceTskAssignment = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule->Assignments->Repalce Assignment....]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule->Assignments->Repalce Assignment....]")
			Fn_SchMgr_ReplaceTskAssignment = False
			Exit Function
		End If
		Wait 2
	End If

	
	If objReTaskAssign.Exist(SISW_MIN_TIMEOUT)Then

		'Select Assign Member
		objReTaskAssign.JavaTree("AssignmentMemberTree").Expand "Assignment Member:" & sName
		objReTaskAssign.JavaTree("AssignmentMemberTree").Select  "Assignment Member:" & sName
		If Err.Number < 0 Then
			Fn_SchMgr_ReplaceTskAssignment = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Assignment memebr " & sName )
			objReTaskAssign.JavaButton("Cancel").Click micLeftBtn
			Set objReTaskAssign = Nothing
			Exit Function
		End If

		'Select Replace member
		objReTaskAssign.JavaTree("ReplaceWithMemberTree").Expand "Schedule Member:" & sReplace
		objReTaskAssign.JavaTree("ReplaceWithMemberTree").Select "Schedule Member:" &  sReplace
		If Err.Number < 0 Then
			Fn_SchMgr_ReplaceTskAssignment = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Assignment memebr " & sReplace )
			objReTaskAssign.JavaButton("Cancel").Click micLeftBtn
			Set objReTaskAssign = Nothing
			Exit Function
		End If

		'Click Done button
		objReTaskAssign.JavaButton("Done").Click
		Fn_SchMgr_ReplaceTskAssignment = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully replace task assignment" )
		If Err.Number < 0 Then
			Fn_SchMgr_ReplaceTskAssignment = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Assignment memebr " & sReplace )
			objReTaskAssign.JavaButton("Cancel").Click micLeftBtn
			Set objReTaskAssign = Nothing
			Exit Function
		End If

	End If
	Set objReTaskAssign = Nothing
End Function

'*********************************************  Function designates Disciplines to the specified tasks..**************************************************************

'Function Name		:					Fn_SchMgr_TskDesignateDisciplines  

'Description			 :		 		  The Function designates Disciplines to the specified tasks.

'Parameters			   :	 			1.  sName : The name of the Discipline that has to be designated.
'													 2.sUser : The name of the user(s) under the discipline to which  the task is to be designated 														
'													3.sOption : All/Common
'													4.sTaskName : Name of the task for multi task seprated it using , (Comma)
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:				 Fn_SchMgr_TskDesignateDisciplines("Testch:t1","TestDis","Rupali Palhade (x_palhad)","All")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali						05-Jun-2010	   		1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_TskDesignateDisciplines(sTaskName,sName,sUser,sOption)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TskDesignateDisciplines"
	On Error Resume Next

	Dim bReturn,sIndex,objSchd,objAssignTable,objOutAssignTable,objDispMem,objOutAssign

	'Set objSchd = JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow")

   'Select Task from schedule table.
    bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
    If  bReturn <> False Then
		Fn_SchMgr_TskDesignateDisciplines = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sTaskName)
	 ELse
		Fn_SchMgr_TskDesignateDisciplines = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sTaskName)
		'Set objSchd = Nothing
		Exit Function
	 End If

	'Select Scheduel->Assignments->Designate discipline...
	 If ((Not JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Designate Disciplines").Exist(SISW_MIN_TIMEOUT))AND (NOt JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Designate Disciplines").Exist(SISW_MIN_TIMEOUT))) Or JavaDialog("Discipline Members").Exist(SISW_MIN_TIMEOUT) Then
		bReturn = Fn_MenuOperation("Select","Schedule:Assignments:Designate Discipline...")
			If bReturn = True Then
                Fn_SchMgr_TskDesignateDisciplines = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Assignments:Designate Discipline...]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Assignments:Designate Discipline...]")
				Fn_SchMgr_TskDesignateDisciplines = False
				Set objSchd = Nothing
				Exit Function
			End If
	 End If 

	If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Designate Disciplines").Exist(SISW_MIN_TIMEOUT) Then
		Set objSchd = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet")
	ElseIf JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Designate Disciplines").Exist(SISW_MIN_TIMEOUT) Then	
		Set objSchd = JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow")	
	End If	
	
	 'Choose a discipline to designate
	 If objSchd.JavaDialog("Designate Disciplines").Exist(SISW_MIN_TIMEOUT) Then

		 'Select show option
		If sOption <> "" Then
			If objSchd.JavaDialog("Designate Disciplines").JavaRadioButton("ShowOption").Exist(SISW_MICRO_TIMEOUT) Then
				objSchd.JavaDialog("Designate Disciplines").JavaRadioButton("ShowOption").SetTOProperty "attached text",sOption
				objSchd.JavaDialog("Designate Disciplines").JavaRadioButton("ShowOption").Set "ON"
			Else
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select show discipline option")
				Fn_SchMgr_TskDesignateDisciplines = False
				Set objSchd = Nothing
				Exit Function
			End If
		End If

		'Select Disciplines.
		 If sName <>  ""  Then
			 objSchd.JavaDialog("Designate Disciplines").JavaList("DisciplinesList").Select sName
			 If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to choose discipline " &sName& "to designate")
				Fn_SchMgr_TskDesignateDisciplines = False
				Set objSchd = Nothing
				Exit Function
			End If
		 End If

		'Click Next button.
		objSchd.JavaDialog("Designate Disciplines").JavaButton("Next").Click

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to choose discipline " &sName& "to designate")
			Fn_SchMgr_TskDesignateDisciplines = False
			Set objSchd = Nothing
			Exit Function
		End If
	End If    

	'Assign discipline members
	If  sName <> "" Then
		Set objDispMem = objSchd.JavaDialog("Discipline Members")
		Set objOutAssign =objSchd.JavaDialog("Outside Discipline Members")
	Else 
		Set objDispMem = JavaDialog("Discipline Members")
		Set objOutAssign = JavaDialog("Discipline Members").JavaDialog("Outside Discipline Members")
	End If
	
	If objDispMem.Exist(SISW_MIN_TIMEOUT) Then
		Set objAssignTable = objDispMem.JavaTable("AssignedDisciplines")
		sIndex = Fn_SchMgr_TableRowIndex(objAssignTable,sUser,"#1")
		
		If sIndex <> False Then
			objAssignTable.ClickCell sIndex,0,"LEFT" 
			objDispMem.JavaButton("OK").Click

			If  sName <> "" Then
				objSchd.JavaDialog("Designate Disciplines").Close
			End IF

			Fn_SchMgr_TskDesignateDisciplines = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully assign discipline member.")

			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to assign discipline member.")
				Fn_SchMgr_TskDesignateDisciplines = False
				objDispMem.JavaButton("Cancel").Click
				Set objSchd = Nothing
				Set objAssignTable = Nothing
				Set objDispMem = Nothing
				Set objOutAssign =Nothing 
				Exit Function
			End If

		Else 
			objDispMem.JavaButton("Expand").Click
			If  sName <> "" Then
			Else
			End If 
			If objOutAssign.Exist(SISW_MIN_TIMEOUT) Then

				Set objOutAssignTable = objOutAssign.JavaTable("AssignedDisciplines")
				sIndex = Fn_SchMgr_TableRowIndex(objOutAssignTable,sUser,"#1")

				If sIndex <> False Then
					objOutAssignTable.ClickCell sIndex,0,"LEFT" 
					objOutAssign.JavaButton("OK").Click
					If  sName <> "" Then
						objSchd.JavaDialog("Designate Disciplines").Close
					End If
					Fn_SchMgr_TskDesignateDisciplines = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully assign discipline memeber outside scheduel.")

					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to assign discipline memeber outside scheduel.")
						Fn_SchMgr_TskDesignateDisciplines = False
						objOutAssign.JavaButton("Cancel").Click
						objDispMem.JavaButton("Cancel").Click
						Set objSchd = Nothing
						Set objAssignTable = Nothing
						Set objOutAssignTable = Nothing
						Set objOutAssign = Nothing
						Set objDispMem = Nothing
						Exit Function
					End If

				Else

					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to assign discipline outside scheduel.")
					Fn_SchMgr_TskDesignateDisciplines = False
					objOutAssign.JavaButton("Cancel").Click
					objDispMem.JavaButton("Cancel").Click
					Set objSchd = Nothing
					Set objAssignTable = Nothing
					Set objOutAssignTable = Nothing
					Set objOutAssign = Nothing
					Set objDispMem = Nothing
					Exit Function
				End If
			End If
		End If
	End If

	Set objSchd = Nothing
	Set objAssignTable = Nothing
	Set objOutAssignTable = Nothing
	Set objDispMem = Nothing
	Set objOutAssign = Nothing 

End Function


'*********************************************  Function perform functionality related to Basline.**************************************************************

'Function Name		:					Fn_SchMgr_BaselineOperations  

'Description			 :		 		  The function perform functionality related to Basline.

'Parameters			   :	 			 1.  sAction : Action need to perform. (1.add 2.Modify)
'													  2.sName : The name to be specified for the Baseline. In Case of Modify and delete
'													 3.bActive : The Checkbox that is to be turned OFF if the Baseline  is to be done on active one. 
'													4.bCopyBaseline: The checkbox to be selected if the baseline to 
'                                                     be created is to be copied from a particular baseline.
'													5.sBName : The input to be provided is the name of  the baseline to be copied.
'                                                  6.bBNewTsks : The Check box to be turned on;if while copying  the baseline the new tasks have to be baselined.
'												 7.sTskBaselineUpdates  Values are (Baselines/Not Started tasks/Incomplete tasks)
'												8.sSchName :Name of the schedule need to select from schedule table.

'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed and schedule should be selected..

'Examples				:				 Fn_SchMgr_BaselineOperations("Testch","Modify","","","True","TestBase1","","Not Started tasks")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali						08-Jun-2010	   		1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_BaselineOperations(sSchName,sAction,sName,bActive,bCopyBaseline,sBName,bBNewTsks,sTskBaselineUpdates)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_BaselineOperations"
   On Error Resume Next 
   Dim objBaseSch,bReturn
   Set objBaseSch = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Baseline Schedule")

	If sSchName <> "" Then
		 bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sSchName,"","","")
		If  bReturn <> False Then
			Fn_SchMgr_BaselineOperations = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected schedule" + sSchName)
		ELse
			Fn_SchMgr_BaselineOperations = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  schedule" + sSchName)
			Set objBaseSch = Nothing
			Exit Function
		End If
	End If

	If Not objBaseSch.Exist(SISW_MIN_TIMEOUT) Then
		bReturn = Fn_MenuOperation("Select","Schedule:Baseline Schedule")
		If bReturn = True Then
			Fn_SchMgr_BaselineOperations = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Baseline Schedule]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Baseline Schedule]")
			Fn_SchMgr_BaselineOperations = False
			Set objBaseSch = Nothing
			Exit Function
		End If
	End If 

	If objBaseSch.Exist(SISW_MIN_TIMEOUT) Then
		Select Case sAction

			Case "Add"
				'Set the name of the baseline
				If sName <> "" Then
					objBaseSch.JavaEdit("BaselineName").Set  sName
				End If
				'Click OK button
				 objBaseSch.JavaButton("OK").WaitProperty "enabled", 1, 20000		
				 objBaseSch.JavaButton("OK").Click
				 
				 If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add new schedule baseline" & sName)
					Fn_SchMgr_BaselineOperations = False
					Set objBaseSch = Nothing
					Exit Function
				Else 
					Fn_SchMgr_BaselineOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully add new schedule baseline" & sName )
				 End If
		
			Case "Modify" 
				'Set name of the baseline
				If sName <> "" Then
					objBaseSch.JavaEdit("BaselineName").Set  sName
				End If
				'Set the value of  Active check box.
				If bActive <> "" Then
					If Cbool(bActive) = true Then
						objBaseSch.JavaCheckBox("Active").Set "ON"
					ElseIf Cbool(bActive) = false Then
						objBaseSch.JavaCheckBox("Active").Set "OFF"
					End If
				End If
				'Set the value of Copy Baseline checkbox.
				If bCopyBaseline <> "" Then
					If Cbool(bCopyBaseline) = true Then
						objBaseSch.JavaCheckBox("Copy Baseline").Set "ON" 
						
						'Select the value of choose baseline.
						If sBName <> "" Then
							objBaseSch.JavaList("ChooseBaseline").Select sBName
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select choose baseline value " & sBName)
								Fn_SchMgr_BaselineOperations = False
								Set objBaseSch = Nothing
								Exit Function
							End If
						End If

						'Set the value of  Baseline new tasks checkbox.
						If bBNewTsks <> "" Then
							If Cbool(bBNewTsks) = true Then
								objBaseSch.JavaCheckBox("Baseline new tasks").Set "ON"
							ElseIf Cbool(bBNewTsks) = false Then
								objBaseSch.JavaCheckBox("Baseline new tasks").Set "OFF"
							End If
						End If

						'Set the value of  Task Baseline updates
						If sTskBaselineUpdates <> ""  Then
							Select Case sTskBaselineUpdates
							Case "Baselines"
								objBaseSch.JavaRadioButton("Do not update baselines").Set "ON"
							Case "Not Started tasks"
								objBaseSch.JavaRadioButton("Update baselines for Not").Set "ON"
							Case "Incomplete tasks"
								objBaseSch.JavaRadioButton("Update baselines for Incomplet").Set "ON"
							Case Else
								Fn_SchMgr_BaselineOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set Task Baseline updates value " & sTskBaselineUpdates)
								Set objBaseSch = Nothing
								Exit Function
							End Select
						End If
					ElseIf Cbool(bCopyBaseline) = false Then
						objBaseSch.JavaCheckBox("Copy Baseline").Set "OFF"
					End If
				End If

				'Click OK button
				objBaseSch.JavaButton("OK").Click micLeftBtn

				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modify baseline")
					Fn_SchMgr_BaselineOperations = False
					Set objBaseSch = Nothing
					Exit Function
				Else
					Fn_SchMgr_BaselineOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully modified baseline")
				End If
				Wait 1
				JavaWindow("ScheduleManagerWindow").Dialog("ErrorDialog").SetTOProperty "title","Scheduling Error"
				If JavaWindow("ScheduleManagerWindow").Dialog("ErrorDialog").Exist(2) Then
					Call Fn_SchMgr_DialogMsgVerify("Scheduling Error", "","OK")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modify baseline")
					Fn_SchMgr_BaselineOperations = False
					Set objBaseSch = Nothing
					Exit Function
				End If
		End Select
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Baseline schedule dialog does not exist")
		Fn_SchMgr_BaselineOperations = False
		Set objBaseSch = Nothing
		Exit Function
	End If
   Set objBaseSch = Nothing
End Function

'*********************************************   Function is used to manage baselines.**************************************************************

'Function Name		:					Fn_SchMgr_ManageBaselines   

'Description			 :		 		  	The Function is used to manage baselines

'Parameters			   :	 			 1.  sAction : Action need to perform. 1. Set Active 2. Add Baseline 3. Delete  4.Verify Baseline
'													  2.sName : The name of the baseline to be Created/Set Active/Deleted.
'													 3.bActive : The Checkbox that is to be turned OFF if the Baseline  is to be done on active one. 
'													4.bCopyBaseline: The checkbox to be selected if the baseline to 
'                                                     be created is to be copied from a particular baseline.
'													5.sBName : The input to be provided is the name of  the baseline to be copied.
'                                                  6.bBNewTsks : The Check box to be turned on;if while copying  the baseline the new tasks have to be baselined.
'												  7.sTskBaselineUpdates  Values are (Baselines/Not Started tasks/Incomplete tasks)
'												 8.sSchName :Name of the schedule need to select from schedule table.

'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed 

'Examples				:				Fn_SchMgr_ManageBaselines("Testch","Verify","sdfdf","","","","","Baselines")
'											  bReturn=Fn_SchMgr_ManageBaselines("qwerty","Edit","Shreyas:Baseline","","","","","Baselines")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali						09-Jun-2010	   		1.0
'											Shreyas					  03-05-2011           1.1				Added case "Edit"
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_ManageBaselines(sSchName,sAction,sName,bActive,bCopyBaseline,sBName,bBNewTsks,sTskBaselineUpdates)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ManageBaselines"
   On Error Resume Next 

   Dim objBaseSch,objTable,bReturn,sIndex,aProperties,sRows,sValue

   Set objBaseSch = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Manage Baselines")

	If sSchName <> "" Then
		bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sSchName,"","","")
		If  bReturn <> False Then
			Fn_SchMgr_ManageBaselines = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected schedule " + sSchName)
		ELse
			Fn_SchMgr_ManageBaselines = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  schedule " + sSchName)
			Set objBaseSch = Nothing
			Exit Function
		End If
	End If

	If Not objBaseSch.Exist(SISW_MIN_TIMEOUT) Then
		bReturn = Fn_MenuOperation("Select","Schedule:Manage baselines")
		If bReturn = True Then
			Fn_SchMgr_ManageBaselines = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Manage baselines]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Manage baselines]")
			Fn_SchMgr_ManageBaselines = False
			Set objBaseSch = Nothing
			Exit Function
		End If
	End If 

	If objBaseSch.Exist(SISW_MIN_TIMEOUT) Then
		Set objTable  = objBaseSch.JavaTable("ManageBaselinesTable")

		Select Case sAction
			Case "Set Active"
				sIndex = Fn_SchMgr_TableRowIndex(objTable,sName,"#0")
				If sIndex <> False Then
					objTable.SelectRow sIndex
					objBaseSch.JavaButton("Set Active").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Set Active").Click
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,5000
					objBaseSch.JavaButton("Close").Click
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set active baseline " + sName )
						Fn_SchMgr_ManageBaselines = False
						Set objBaseSch = Nothing
						Set objTable = Nothing
						Exit Function
					Else
						Fn_SchMgr_ManageBaselines = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set active baseline " + sName )
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get row index of  " + sName )
					Fn_SchMgr_ManageBaselines = False
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,5000
					objBaseSch.JavaButton("Close").Click
					Set objBaseSch = Nothing
					Set objTable = Nothing
					Exit Function
				End If

			Case "Add Baseline"
				objBaseSch.JavaButton("Add Baseline").WaitProperty  "enabled",1,20000
				objBaseSch.JavaButton("Add Baseline").Click

				bReturn = Fn_SchMgr_BaselineOperations("","Modify",sName,bActive,bCopyBaseline,sBName,bBNewTsks,sTskBaselineUpdates)

				If bReturn = True Then
					Fn_SchMgr_ManageBaselines = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully added  " + sName + " baseline." )
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Close").Click
				ElseIf  bReturn = False Then
					Fn_SchMgr_ManageBaselines = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add  " + sName + " baseline." )
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Close").Click
					Set objBaseSch = Nothing
					Set objTable = Nothing
					Exit Function
				End If

				 If objBaseSch.JavaDialog("BaselineNameError").Exist (SISW_MICRO_TIMEOUT) Then
							bReturn= Fn_Button_Click("Fn_SchMgr_ManageBaselines", objBaseSch.JavaDialog("BaselineNameError"),"OK")
									If bReturn=False Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Close the [Baseline Name Error] Dialog" )
											Set objBaseSch = Nothing
											Exit Function
								End If

								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Baseline Schedule").Close
								If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Close the [Baseline Schedule] Dialog" )
											Set objBaseSch = Nothing
											Exit Function
								End If
							 	objBaseSch.Close
							 	If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Close the [Manage Baselines] Dialog" )
											Set objBaseSch = Nothing
											Exit Function
								End If
								Fn_SchMgr_ManageBaselines=False
								Set objBaseSch = Nothing
								Exit Function								
					End If	

			Case "Delete"
				sIndex = Fn_SchMgr_TableRowIndex(objTable,sName,"#0")

				If sIndex <> False Then
					objTable.SelectRow sIndex
					objBaseSch.JavaButton("Delete Baseline").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Delete Baseline").Click
					If JavaDialog("Delete Baseline").Exist(SISW_MIN_TIMEOUT) Then
						JavaDialog("Delete Baseline").JavaButton("Yes").Click
					End If
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Close").Click
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to delete baseline " + sName )
						Fn_SchMgr_ManageBaselines = False
						objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
						objBaseSch.JavaButton("Close").Click
						Set objBaseSch = Nothing
						Set objTable = Nothing
						Exit Function
					Else
						Fn_SchMgr_ManageBaselines = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully delete baseline " + sName )
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get row index of " + sName )
					Fn_SchMgr_ManageBaselines = False
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Close").Click
					Set objBaseSch = Nothing
					Set objTable = Nothing
					Exit Function
				End If

			Case "Verify"
				sIndex = Fn_SchMgr_TableRowIndex(objTable,sName,"#0")
				If sIndex <> False Then
					Fn_SchMgr_ManageBaselines = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify  baseline " + sName + " is active." )
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  verify  baseline " + sName + " is active." )
					Fn_SchMgr_ManageBaselines = False
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Close").Click
					Set objBaseSch = Nothing
					Set objTable = Nothing
					Exit Function
				End If 
				objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
				objBaseSch.JavaButton("Close").Click

		Case "Edit"
			If sName<>"" Then
				aProperties=split(sName,":",-1,1)
				'Search for the baseline of intrest to modify
				sRows=objBaseSch.JavaTable("ManageBaselinesTable").GetROProperty ("rows")
				For iCount=0 to sRows-1
					sValue=objBaseSch.JavaTable("ManageBaselinesTable").GetCellData (iCount,0)
					If lCase(aProperties(0))=lcase(sValue) Then
						objBaseSch.JavaTable("ManageBaselinesTable").DoubleClickCell iCount,0
						objBaseSch.JavaEdit("EditCell").set aProperties(1)
						Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ENTER}"
						Set WshShell = nothing
						'check the existence of the Error dialog and handle it
						If objBaseSch.JavaDialog("Error").Exist(3)=true Then
							objBaseSch.JavaDialog("Error").JavaButton("OK").Click micLeftBtn
							objBaseSch.close
							Fn_SchMgr_ManageBaselines = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modify the name from  " + aProperties(0) +" to "+aProperties(1)+" as it already exists with the name "+aProperties(1) )
							Exit function
						End If
						Fn_SchMgr_ManageBaselines = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully modified the name from  " + aProperties(0) +" to "+aProperties(1) )
						Exit For
					End If
	
				Next
			End If
			objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
			objBaseSch.JavaButton("Close").Click

			Case "VerifyInDateTaken"
					Fn_SchMgr_ManageBaselines = False
					iRowCnt = ObjTable.getROProperty("rows")
					For iCnt = 0 to iRowCnt - 1
						If trim(objtable.getCellData(iCnt, "Baseline Name")) = trim(sName) then
								If trim(sBName) = trim(objtable.getCellData(iCnt, "Date Taken"))  Then
										Fn_SchMgr_ManageBaselines = True
										Exit for
								End If
						End IF
					Next
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Close").Click

			Case "VerifyInPerson"
				Fn_SchMgr_ManageBaselines = False
					iRowCnt = ObjTable.getROProperty("rows")
					For iCnt = 0 to iRowCnt - 1
						If trim(objtable.getCellData(iCnt, "Baseline Name")) = trim(sName) then
								If trim(sBName) = trim(objtable.getCellData(iCnt, "Person"))  Then
										Fn_SchMgr_ManageBaselines = True
										Exit for
								End If
						End IF
					Next
					objBaseSch.JavaButton("Close").WaitProperty "enabled",1,20000
					objBaseSch.JavaButton("Close").Click

		End Select
	Else 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Manage baseline dialog does not exists.")
		Fn_SchMgr_ManageBaselines = False
		Set objBaseSch = Nothing
		Exit Function
	End If

	Set objBaseSch = Nothing
	Set objTable = Nothing
End Function



'*********************************************************		Function to Verify Warning Message		***********************************************************************
'Function Name		:        Fn_SchMgr_WarningMsgVerify  

'Description	    	:        Verifies the message on Task Indent Warning dialog

'Parameters		     :    		sMesssage: Message to be Verified [Optional]
'			                         		 sButton: Button to be clicked on the doalig

'Return Value		: 			True/False

'Pre-requisite	    :		     Task Indent Warning Dialog should be Present

'Examples		    :			Call  Fn_SchMgr_WarningMsgVerify ("", "OK")

'History		    :		
'													Developer Name				Date						Rev. No.						Changes Done						Reviewer
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Prasanna 							     11/06/2010			              1.0								
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_WarningMsgVerify(sMesssage, sButton)

		Dim dicErrorInfo 
		Set dicErrorInfo = CreateObject("Scripting.Dictionary")
		With dicErrorInfo 		 
		  .Add "Title", "Warning"		 
		 .Add "Message", sMessage		 
		 .Add "Button", sButton		 
		End with
		Fn_SchMgr_WarningMsgVerify = Fn_SISW_SchMgr_ErrorVerify(dicErrorInfo)

End Function

'*********************************************  The function creates a Task Deliverable. **************************************************************

'Function Name		:					Fn_SchMgr_TaskDeliverable 

'Description			 :		 		  	The function creates a Task Deliverable.

'Parameters			   :	 			 1.  sInvoke:- The input to this parameter can be either of the following; Menu/RMB.
'													  2. sAction:- The input to this parameter can be one of the following; Add/Remove/Verify. 
'													 3.sTskName:- The name of the task for which a deliverable is to be created. 		
'													4.sSchDeliverable :Schedule need to deliver.(Need to pass in case Add/Remove/Verify. )				
'													5.sSchSubmit  : Schedule submit .

'Return Value		   : 				True/False

'Pre-requisite			:		 		1.The Dataset that is to be attached as the deliverable is created in Teamcenter.
'                                                   2. The above created deliverable is listed as one of the deliverable for Schedules.

'Examples				:				Fn_SchMgr_TaskDeliverable("Testch:t1","Menu","Add","QQQ","Target")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali						10-Jun-2010	   		 1.0
'											Sachin Joshi				21-June-2011                  Added Case "ClickDelivarableCell"  Prasanna
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_TaskDeliverable(sTskName,sInvoke,sAction,sSchDeliverable,sSchSubmit)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TaskDeliverable"
   On Error Resume Next 

   Dim bReturn,sIndex,objTaskDel,objTable,sActualval,sIndexInt
   Set objTaskDel = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Task Deliverables")

   If Not objTaskDel.Exist(SISW_MIN_TIMEOUT) Then
			Select Case sInvoke
			Case "Menu"
				bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sTskName,"","","")
				If  bReturn <> False Then
					Fn_SchMgr_TaskDeliverable = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected task " + sTskName)
				ELse
					Fn_SchMgr_TaskDeliverable = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task " + sTskName)
					Set objTaskDel = Nothing
					Exit Function
				End If
				Call Fn_ReadyStatusSync(1)
				bReturn = Fn_MenuOperation("Select","Schedule:Task Deliverables")
				If bReturn = True Then
					Fn_SchMgr_TaskDeliverable = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Task Deliverables]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Task Deliverables]")
					Fn_SchMgr_TaskDeliverable = False
					Set objTaskDel = Nothing
					Exit Function
				End If
	
			Case "RMB"
				bReturn =  Fn_SchMgr_SchTable_NodeOperation("PopupMenu", sTskName, "", "", "Task Deliverables")
				If bReturn = True Then
					Fn_SchMgr_TaskDeliverable = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked RMB Menu [Task Deliverables]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked RMB Menu [Task Deliverables]")
					Fn_SchMgr_TaskDeliverable = False
					Set objTaskDel = Nothing
					Exit Function
				End If
		End Select
	End If
	If JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Information").Exist(SISW_MIN_TIMEOUT) Then
		JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Information").JavaButton("OK").Click
	ElseIf JavaWindow("ScheduleManagerWindow").JavaWindow("Information").Exist(SISW_MICRO_TIMEOUT) Then
		JavaWindow("ScheduleManagerWindow").JavaWindow("Information").JavaButton("OK").Click
	End If
	wait(2)
	If objTaskDel.Exist(SISW_MIN_TIMEOUT) Then
		Set objTable =  objTaskDel.JavaTable("TaskDeliverablesTable")

		Select Case sAction
			Case "Add"
				objTaskDel.JavaButton("Add").WaitProperty "enabled",1,10000
				objTaskDel.JavaButton("Add").Click micLeftBtn
				If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Add Task deliverable  " + sSchDeliverable)
						Fn_SchMgr_TaskDeliverable = False
						objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objTaskDel.JavaButton("Cancel").Click micLeftBtn
						Set objTaskDel = Nothing
						Set objTable = Nothing
						Exit Function
					End If

				sIndex =Cint( objTable.GetROProperty("rows"))
				If sIndex > 0 Then

					If sSchDeliverable <> "" Then
						objTable.ClickCell (sIndex-1),"#0", "LEFT", "NONE"
                        objTaskDel.JavaList("TskDelList").Select sSchDeliverable
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select schedule deliverable  " + sSchDeliverable)
							Fn_SchMgr_TaskDeliverable = False
							objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
							objTaskDel.JavaButton("Cancel").Click micLeftBtn
							Set objTaskDel = Nothing
							Set objTable = Nothing
							Exit Function
						End If
					End If 

					If sSchSubmit <> "" Then
						objTable.ClickCell (sIndex-1),"#1", "LEFT", "NONE"
                        objTaskDel.JavaList("TskDelList").Select sSchSubmit
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select schedule submit  " + sSchSubmit)
							Fn_SchMgr_TaskDeliverable = False
							objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
							objTaskDel.JavaButton("Cancel").Click micLeftBtn
							Set objTaskDel = Nothing
							Set objTable = Nothing
							Exit Function
						End If
					End If
					objTaskDel.JavaButton("OK").WaitProperty "enabled",1,20000
					objTaskDel.JavaButton("OK").Click micLeftBtn

					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select schedule deliverable  " + sSchDeliverable)
						Fn_SchMgr_TaskDeliverable = False
						objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objTaskDel.JavaButton("Cancel").Click micLeftBtn
						Set objTaskDel = Nothing
						Set objTable = Nothing
						Exit Function
					Else
						Fn_SchMgr_TaskDeliverable = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully added task deliverable")
					End If
				End If

			Case "Remove"
				If sSchDeliverable <> "" Then
					sIndex = Fn_SchMgr_TableRowIndex(objTable,sSchDeliverable,"#0")
					If sIndex <> False Then
						'objTable.SelectRow sIndex
						objTable.SelectRowsRange sIndex,sIndex
						objTaskDel.JavaButton("Remove").WaitProperty "enabled",1,20000
						objTaskDel.JavaButton("Remove").Click micLeftBtn 
						objTaskDel.JavaButton("OK").WaitProperty "enabled",1,20000
						objTaskDel.JavaButton("OK").Click micLeftBtn

						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  remove " + sSchDeliverable + " from task deliverable table.")
							Fn_SchMgr_TaskDeliverable = False
							objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
							objTaskDel.JavaButton("Cancel").Click micLeftBtn
							Set objTaskDel = Nothing
							Set objTable = Nothing
							Exit Function
						Else
							Fn_SchMgr_TaskDeliverable = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove " + sSchDeliverable + " from task deliverable table.")
						End If

					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sSchDeliverable + " row does not exist in task deliverable table.")
						Fn_SchMgr_TaskDeliverable = False
						objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objTaskDel.JavaButton("Cancel").Click micLeftBtn
						Set objTaskDel = Nothing
						Set objTable = Nothing
						Exit Function
					End If
				End If
				
			
			Case "Verify"

				If sSchDeliverable <> "" Then
					sIndex = Fn_SchMgr_TableRowIndex(objTable,sSchDeliverable,"#0")
					If sIndex <> False Then
						sActualval = objTable.GetCellData(sIndex,"#0")
						If  sActualval = sSchDeliverable Then
							Fn_SchMgr_TaskDeliverable = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify  " + sSchDeliverable + " value of  Schedule Deliverable.")
						Else 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  verify  " + sSchDeliverable + " value of  Schedule Deliverable.")
							Fn_SchMgr_TaskDeliverable = False
							objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
							objTaskDel.JavaButton("Cancel").Click micLeftBtn
							Set objTaskDel = Nothing
							Set objTable = Nothing
							Exit Function
						End If

						If sSchSubmit <> "" Then
							sActualval = objTable.GetCellData(sIndex,"#1")
							If  sActualval = sSchSubmit Then
								Fn_SchMgr_TaskDeliverable = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify  " + sSchSubmit + " value of  Submit Method.")
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  verify  " + sSchSubmit + " value of  Submit Method.")
								Fn_SchMgr_TaskDeliverable = False
								objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
								objTaskDel.JavaButton("Cancel").Click micLeftBtn
								Set objTaskDel = Nothing
								Set objTable = Nothing
								Exit Function
							End If
						End If
						objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objTaskDel.JavaButton("Cancel").Click micLeftBtn
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sSchDeliverable + " row does not exist in task deliverable table.")
						Fn_SchMgr_TaskDeliverable = False
						objTaskDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objTaskDel.JavaButton("Cancel").Click micLeftBtn
						Set objTaskDel = Nothing
						Set objTable = Nothing
						Exit Function
					End If
				End If
			Case "ClickDelivarableCell"
				If sSchDeliverable <> "" Then
					sIndex = Fn_SchMgr_TableRowIndex(objTable,sSchDeliverable,"#0")
					If sIndex <> False Then
						Fn_SchMgr_TaskDeliverable = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Link for Schedule Delivarable "+sSchDeliverable+" in Schedule Deliverable Table.")
						objTable.ClickCell cInt(sIndex),2
						wait(2)
					Else
						Fn_SchMgr_TaskDeliverable = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Click on Link for Schedule Delivarable "+sSchDeliverable+" in Schedule Deliverable Table.")
					End If
				End If
		End Select
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task Deliverable dialog does not exist.")
		Fn_SchMgr_TaskDeliverable = False
		Set objTaskDel = Nothing
		Exit Function
	End If

	Set objTaskDel = Nothing
	Set objTable = Nothing

End Function

'*********************************************     Function selects a baseline that has to be viewed..   *********************************************************************

'Function Name		:					Fn_SchMgr_ViewBaseline    

'Description			 :		 		  	The function selects a baseline that has to be viewed.

'Parameters			   :	 			 1.  sName: Name of the baseline to be viewed

'Return Value		   : 				True/False

'Pre-requisite			:		 		 Schedule Manager window should be displayed 

'Examples				:				 Fn_SchMgr_ViewBaseline ("Baseline1")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali						10-Jun-2010	   		1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ViewBaseline(sName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ViewBaseline"
   On Error Resume Next
   Dim objViewBase,bReturn
   Set objViewBase = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("View Baseline")

	If Not objViewBase.Exist(SISW_MIN_TIMEOUT) Then
		bReturn = Fn_MenuOperation("Select","View:View Baseline")

		If bReturn = True Then
			Fn_SchMgr_ViewBaseline = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [View:View Baseline]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [View:View Baseline]")
			Fn_SchMgr_ViewBaseline = False
			Set objViewBase = Nothing
			Exit Function
		End If
	End If

	If objViewBase.Exist(SISW_MIN_TIMEOUT) Then
		objViewBase.JavaList("BaselineList").Select sName

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select baseline " + sName)
			Fn_SchMgr_ViewBaseline = False
			objViewBase.JavaButton("Cancel").WaitProperty "enabled",1,20000 
			objViewBase.JavaButton("Cancel").Click
			Set objViewBase = Nothing
			Exit Function 
		End If

		objViewBase.JavaButton("OK").WaitProperty "enabled",1,20000 
		objViewBase.JavaButton("OK").Click

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select baseline " + sName)
			Fn_SchMgr_ViewBaseline = False
			objViewBase.JavaButton("Cancel").WaitProperty "enabled",1,20000 
			objViewBase.JavaButton("Cancel").Click
			Set objViewBase = Nothing
			Exit Function 
		Else 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected baseline " + sName)
			Fn_SchMgr_ViewBaseline = True
		End If
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "View Baseline dialog does not exist.]")
		Fn_SchMgr_ViewBaseline = False
		objViewBase.JavaButton("Cancel").WaitProperty "enabled",1,20000 
		objViewBase.JavaButton("Cancel").Click
		Set objViewBase = Nothing
		Exit Function 
	End If 

	Set objViewBase = Nothing

End Function

'*********************************************  The function creates a schedule deliverable. **************************************************************

'Function Name		:					Fn_SchMgr_SchDeliverableCreate  

'Description			 :		 		  	The function creates a schedule deliverable.

'Parameters			   :	 			1. sAction :- The input to this parameter can be one of the following; Add/Remove/Verify/Modify. 
'													 2.sSchName:- The name of the schedule for which a deliverable is to be created. 		
'													3.sName :The name to be given to the deliverable that needs to be created			
'													4.sDelName   : The name of the deliverable that needs to be attached to the schedule
'												   5.sDelType: The type of the deliverable to be attached.

'Return Value		   : 				True/False

'Pre-requisite			:		 		1.The Dataset that is to be attached as the deliverable is created in Teamcenter.

'Examples				:				Fn_SchMgr_SchDeliverableCreate("Verify","Testch","","AAA","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Rupali				15-Jun-2010	   		 1.0
'											Priyanka			11-11-2010			  1.0				Modified Column Name From 
'																									 From "Deliverable Name" "#0". 
'																									 changes Done As per Application Change in Tc8_3_0_2	
'											Vandana Patel	08-Aug-2012	   		 1.0					Modified object hierarchy of Open by Name dialog
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_SchDeliverableCreate(sAction,sSchName,sName,sDelName,sDelType)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_SchDeliverableCreate"
   On Error Resume Next 

   Dim bReturn,objScheduleDel,sIndex,objDelTable,objTable,iTypeIndex,sActualVal
   Dim sTitle,bFlag,aName, objOpenByName
   bFlag = False
   sTitle = "Update Schedule Deliverable Error"
   sErrorText = "Schedule deliverables should have unique names."
   Set objScheduleDel = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Deliverables")
   Set objOpenByName = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Open by Name")
   'Set the value of java table
   Set objTable = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Deliverables").JavaTable("SchDeliverablesTable")
   
   If Instr(1,JavaWindow("DefaultWindow").GetROProperty("title"),"My Teamcenter") > 0 Then
   		 Set objScheduleDel = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Schedule Deliverables")
   		 Set objTable = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Schedule Deliverables").JavaTable("SchDeliverablesTable")
   End If


	If Not objScheduleDel.Exist(SISW_MIN_TIMEOUT) Then
		bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sSchName,"","","")
		If  bReturn <> False Then
			Fn_SchMgr_SchDeliverableCreate = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected schedule " + sSchName)
		ELse
			Fn_SchMgr_SchDeliverableCreate = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  schedule " + sSchName)
			Set objScheduleDel = Nothing
			sErrorText = ""
			Exit Function
		End If

		bReturn = Fn_MenuOperation("Select","Schedule:Schedule Deliverables")
		If bReturn = True Then
			Fn_SchMgr_SchDeliverableCreate = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Schedule Deliverables]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Schedule Deliverables]")
			Fn_SchMgr_SchDeliverableCreate = False
			Set objScheduleDel = Nothing
			sErrorText = ""
			Exit Function
	    End If
	End If

    If objScheduleDel.Exist(SISW_MIN_TIMEOUT) Then

		Select Case sAction
			Case "Add"
				objScheduleDel.JavaButton("Add").WaitProperty "enabled",1,20000
				objScheduleDel.JavaButton("Add").Click micLeftBtn
				sIndex =Cint( objTable.GetROProperty("rows"))
				If sIndex > 0 Then

					If sDelName <> "" Then
						objTable.ClickCell (sIndex-1),"#0","LEFT","NONE"
                        objScheduleDel.JavaEdit("DelTableEdit").Object.setText sDelName
						objTable.ClickCell (sIndex-1),"#2","LEFT","NONE"
					End If

					bReturn = Fn_SchMgr_DialogMsgVerify(sTitle,sErrorText,"OK")
					If bReturn Then
						Fn_SchMgr_SchDeliverableCreate = False
						objScheduleDel.JavaButton("Cancel").Click
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add deliverable.[" + sDelName + "]")
						Set objScheduleDel  = Nothing 
						Set objTable = Nothing
						Set objDelTable = Nothing
						sErrorText = ""
						Exit Function
					End If

					If  sDelType <> "" Then
						objTable.ClickCell (sIndex-1), "#1","LEFT", "NONE"
						iTypeIndex = objScheduleDel.JavaList("SchDelList").GetItemIndex(sDelType)
						If Err.Number < 0 Then
							Fn_SchMgr_SchDeliverableCreate = False 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sDelType + " deliverable type not exist.")
							objScheduleDel.JavaButton("Cancel").Click
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							Set objDelTable = Nothing
							sErrorText = ""
							Exit Function
						End If
						objScheduleDel.JavaList("SchDelList").Object.setSelectedIndex Cint(iTypeIndex)
					End If
					
					If  sName <> "" Then
						If Instr(1,sName,"-")<>0 Then
							aName = split(sName, "-",-1, 1)
							bFlag = True
						End If
						objTable.ClickCell (sIndex-1), "#3","LEFT", "NONE"
						If objOpenByName.Exist(SISW_MIN_TIMEOUT) Then
							If bFlag = False Then
								objOpenByName.JavaEdit("Name").Object.setText sName
							ElseIf bFlag = True Then
								objOpenByName.JavaEdit("Name").Object.setText aName(1)
							End If
							objOpenByName.JavaButton("Find").Object.doClick(1)
							'Handle Nothing object dialog .
							If JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("Nothing found!").Exist(SISW_MIN_TIMEOUT) Then
								JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("Nothing found!").JavaButton("OK").Click, micLeftBtn
							End If

							Set objDelTable = objOpenByName.JavaTable("DelTable")
							'Modified By Ketan Raje on 25/11/2010
							If objOpenByName.JavaButton("LoadAll").Exist(SISW_MICRO_TIMEOUT) Then
								If objOpenByName.JavaButton("LoadAll").GetROProperty("enabled") = 1 Then
									objOpenByName.JavaButton("LoadAll").Object.doClick 1
								End If
							End If
							' Commented Due to it will not select proper value as it results shows two value as,
							' Datasetname and Datasetname;1
							'sIndex = Fn_SchMgr_TableRowIndex(objDelTable,sName,"#0") 
							sIndex = Fn_SchMgr_TableRowIndex(objDelTable,sName,"Object")
							If sIndex <> False Then
								'Commented below code. By Ketan Raje on 25/11/2010
'								If JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Open by Name").JavaButton("LoadAll").Exist(3) Then
'									JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Open by Name").JavaButton("LoadAll").WaitProperty "enabled",1,20000
'									JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Open by Name").JavaButton("LoadAll").Click ,micLeftBtn
'								End If
								objDelTable.DoubleClickCell sIndex,0
							Else 
								Fn_SchMgr_SchDeliverableCreate = False 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sName + " dataset does not exist.")
								objOpenByName.Close
								objScheduleDel.JavaButton("Cancel").Click
								Set objScheduleDel  = Nothing 
								Set objTable = Nothing
								Set objDelTable = Nothing
								sErrorText = ""
								Exit Function
							End If 
						Else 
							Fn_SchMgr_SchDeliverableCreate = False 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Open By Name dialog does not exist.")
							objScheduleDel.JavaButton("Cancel").Click
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							Set objDelTable = Nothing
							sErrorText = ""
							Exit Function
						End If 
					End If
					'Click OK button
					objScheduleDel.JavaButton("OK").WaitProperty "enabled",1,20000
					objScheduleDel.JavaButton("OK").Click micLeftBtn

					If Err.Number < 0 Then
						Fn_SchMgr_SchDeliverableCreate = False 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add Schedule deliverable.")
						objScheduleDel.JavaButton("Cancel").Click
						Set objScheduleDel  = Nothing 
						Set objTable = Nothing
						Set objDelTable = Nothing
						sErrorText = ""
						Exit Function
					Else
						Fn_SchMgr_SchDeliverableCreate = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully added Schedule deliverable..")
					End If
				End If

			Case "Remove"
				If sDelName <> "" Then
					sIndex = Fn_SchMgr_TableRowIndex(objTable,sDelName,"#0")
					If sIndex <> False Then
						objTable.SelectRow sIndex
						objScheduleDel.JavaButton("Remove").WaitProperty "enabled",1,20000 
						objScheduleDel.JavaButton("Remove").Click micLeftBtn
						objScheduleDel.JavaButton("OK").WaitProperty "enabled",1,20000
						objScheduleDel.JavaButton("OK").Click micLeftBtn
						JavaDialog("Confirmation").SetTOProperty "title", "Confirmation"
						If  JavaDialog("Confirmation").Exist(5) Then
							JavaDialog("Confirmation").JavaButton("Yes").Click micLeftBtn
						End If

							'Aded By Vidya
							If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Deliverables").JavaDialog("Error").Exist(SISW_MIN_TIMEOUT) Then
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Deliverables").JavaDialog("Error").JavaButton("OK").Click micLeftBtn
								Fn_SchMgr_SchDeliverableCreate = False 
								objScheduleDel.JavaButton("Cancel").Click
								Set objScheduleDel  = Nothing 
								Set objTable = Nothing
								sErrorText = ""
								Exit function
						End If

						If Err.Number < 0 Then
							Fn_SchMgr_SchDeliverableCreate = False 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to remove deliverable name  " + sDelName + " from table" )
							objScheduleDel.JavaButton("Cancel").Click
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							sErrorText = ""
							Exit Function
						Else
							Fn_SchMgr_SchDeliverableCreate = True 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully remove deliverable name  " + sDelName + " from table" ) 
						End If

					Else 
						Fn_SchMgr_SchDeliverableCreate = False 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sDelName +  " deliverable name does not exist in table.")
						objScheduleDel.JavaButton("Cancel").Click
						Set objScheduleDel  = Nothing 
						Set objTable = Nothing
						sErrorText = ""
						Exit Function
					End If
				End If

			Case "NoRemove"
				If sDelName <> "" Then
					sIndex = Fn_SchMgr_TableRowIndex(objTable,sDelName,"#0")
					If sIndex <> False Then
						objTable.SelectRow sIndex
						objScheduleDel.JavaButton("Remove").WaitProperty "enabled",1,20000 
						objScheduleDel.JavaButton("Remove").Click micLeftBtn
						objScheduleDel.JavaButton("OK").WaitProperty "enabled",1,20000
						objScheduleDel.JavaButton("OK").Click micLeftBtn
						JavaDialog("Confirmation").SetTOProperty "title", "Confirmation"
						If  JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
							JavaDialog("Confirmation").JavaButton("No").Click micLeftBtn
						End If

						If Err.Number < 0 Then
							Fn_SchMgr_SchDeliverableCreate = False 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to remove deliverable name  " + sDelName + " from table" )
							objScheduleDel.JavaButton("Cancel").Click
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							sErrorText = ""
							Exit Function
						Else
							Fn_SchMgr_SchDeliverableCreate = True 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully remove deliverable name  " + sDelName + " from table" ) 
						End If
							'Added By Vidya
							 JavaDialog("Error").SetTOProperty "title", "Error"
							If  JavaDialog("Error").Exist(5) Then
								JavaDialog("Error").JavaButton("No").Click micLeftBtn       
								objScheduleDel.JavaButton("Cancel").Click
								Fn_SchMgr_SchDeliverableCreate = False 
								Set objScheduleDel  = Nothing 
								Set objTable = Nothing
								sErrorText = ""
								Exit Function			
							End If
					Else 
						Fn_SchMgr_SchDeliverableCreate = False 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sDelName +  " deliverable name does not exist in table.")
						objScheduleDel.JavaButton("Cancel").Click
						Set objScheduleDel  = Nothing 
						Set objTable = Nothing
						sErrorText = ""
						Exit Function
					End If
				End If

			Case "CancelRemove"
				If sDelName <> "" Then
					sIndex = Fn_SchMgr_TableRowIndex(objTable,sDelName,"#0")
					If sIndex <> False Then
						objTable.SelectRow sIndex
						objScheduleDel.JavaButton("Remove").WaitProperty "enabled",1,20000 
						objScheduleDel.JavaButton("Remove").Click micLeftBtn
						objScheduleDel.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objScheduleDel.JavaButton("Cancel").Click micLeftBtn

						If Err.Number < 0 Then
							Fn_SchMgr_SchDeliverableCreate = False 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to remove deliverable name  " + sDelName + " from table" )
							objScheduleDel.JavaButton("Cancel").Click
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							sErrorText = ""
							Exit Function
						Else
							Fn_SchMgr_SchDeliverableCreate = True 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully remove deliverable name  " + sDelName + " from table" ) 
						End If
					Else 
						Fn_SchMgr_SchDeliverableCreate = False 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sDelName +  " deliverable name does not exist in table.")
						objScheduleDel.JavaButton("Cancel").Click
						Set objScheduleDel  = Nothing 
						Set objTable = Nothing
						sErrorText = ""
						Exit Function
					End If
				End If
				
			Case "Verify"
				If sDelName <> "" Then
						sIndex = Fn_SchMgr_TableRowIndex(objTable,sDelName,"#0")
						If sIndex <> False Then
	
							If sDelType <> "" Then
								sActualVal =  objTable.GetCellData(sIndex,"#1")
								If  sActualVal = sDelType Then
									Fn_SchMgr_SchDeliverableCreate = True 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully verify value "+ sDelType + " of deliverable type.")
								Else 
									Fn_SchMgr_SchDeliverableCreate = False 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to verify value "+ sDelType + " of deliverable type.")
									objScheduleDel.JavaButton("Cancel").Click
									Set objScheduleDel  = Nothing 
									Set objTable = Nothing
									sErrorText = ""
									Exit Function
								End If
							End If
	
							If sName <> "" Then
								sActualVal =  objTable.GetCellData(sIndex,"#2")
								If  sActualVal = sName Then
									Fn_SchMgr_SchDeliverableCreate = True 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully verify value "+ sName + " of deliverable.")
								Else 
									Fn_SchMgr_SchDeliverableCreate = False 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to verify value "+ sName + " of deliverable.")
									objScheduleDel.JavaButton("Cancel").Click
									Set objScheduleDel  = Nothing 
									Set objTable = Nothing
									sErrorText = ""
									Exit Function
								End If
							End If
	
							If sIndex <> False Then
								Fn_SchMgr_SchDeliverableCreate = True 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully verify value "+ sDelName + " of deliverable name.")
							Else 
								Fn_SchMgr_SchDeliverableCreate = False 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to verify value "+ sDelName + " of deliverable name.")
								objScheduleDel.JavaButton("Cancel").Click
								Set objScheduleDel  = Nothing 
								Set objTable = Nothing
								sErrorText = ""
								Exit Function
							End If
							objScheduleDel.JavaButton("Cancel").Click
				     Else
							Fn_SchMgr_SchDeliverableCreate = False 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sDelName +  " deliverable name does not exist in table.")
							objScheduleDel.JavaButton("Cancel").Click
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							sErrorText = ""
							Exit Function
					End If
				End If

			Case "Modify"
				If sDelName <> "" Then
					sIndex = Fn_SchMgr_TableRowIndex(objTable,sDelName,"#0")
					If sIndex <> False Then
						If sName <> "" Then
							objTable.ClickCell sIndex,"#0","LEFT","NONE"
							objScheduleDel.JavaEdit("DelTableEdit").Object.setText sName
							objTable.ClickCell sIndex,"#2","LEFT","NONE"
						End If

						bReturn = Fn_SchMgr_DialogMsgVerify(sTitle,sErrorText,"OK")
						If bReturn Then
							Fn_SchMgr_SchDeliverableCreate = False
							objScheduleDel.JavaButton("Cancel").Click
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modify deliverable.[" + sDelName + "]")
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							Set objDelTable = Nothing
							sErrorText = ""
							Exit Function
						End If

						If  sDelType <> "" Then
							objTable.ClickCell sIndex, "#1","LEFT", "NONE"
							iTypeIndex = objScheduleDel.JavaList("SchDelList").GetItemIndex(sDelType)
							If Err.Number < 0 Then
								Fn_SchMgr_SchDeliverableCreate = False 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sDelType + " deliverable type not exist.")
								objScheduleDel.JavaButton("Cancel").Click
								Set objScheduleDel  = Nothing 
								Set objTable = Nothing
								Set objDelTable = Nothing
								sErrorText = ""
								Exit Function
							End If
							objScheduleDel.JavaList("SchDelList").Object.setSelectedIndex Cint(iTypeIndex)
						End If

						If sIndex <> False Then
							Fn_SchMgr_SchDeliverableCreate = True 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Modified value "+ sDelName + " to " + sName)
						Else 
							Fn_SchMgr_SchDeliverableCreate = False 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to Modify value "+ sDelName + " to " + sName)
							objScheduleDel.JavaButton("Cancel").Click
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							sErrorText = ""
							Exit Function
						End If
						objScheduleDel.JavaButton("OK").Click
					End If
				Else
						Fn_SchMgr_SchDeliverableCreate = False 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sDelName +  " deliverable name does not exist in table.")
						objScheduleDel.JavaButton("Cancel").Click
						Set objScheduleDel  = Nothing 
						Set objTable = Nothing
						sErrorText = ""
						Exit Function
				End If

			Case "ClickSubmittal"
					Set objTable = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Deliverables").JavaTable("SchDeliverablesTable")
					sIndex = Fn_SchMgr_TableRowIndex(objTable,sDelName,"#0")
					If Len(sIndex)=5 Then
						Fn_SchMgr_SchDeliverableCreate = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sDelName +  " deliverable name does not exist in table.")
					Else
						objScheduleDel.JavaTable("SchDeliverablesTable").ClickCell sIndex,"Deliverable"
						'Click OK button
						objScheduleDel.JavaButton("OK").WaitProperty "enabled",1,20000
						objScheduleDel.JavaButton("OK").Click micLeftBtn
						Fn_SchMgr_SchDeliverableCreate = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sDelName +  " is Clicked Successfully.")
					End If

        Case "AddSpecial"  'this is a testcase specific case

				sIndex =Cint( objTable.GetROProperty("rows"))
				objScheduleDel.JavaButton("Add").WaitProperty "enabled",1,20000
				objScheduleDel.JavaButton("Add").Click micLeftBtn

					If sDelName <> "" Then
						objTable.ClickCell (sIndex),"#0","LEFT","NONE"
                        objScheduleDel.JavaEdit("DelTableEdit").Object.setText sDelName
						objTable.ClickCell (sIndex),"#2","LEFT","NONE"
					End If

					bReturn = Fn_SchMgr_DialogMsgVerify(sTitle,sErrorText,"OK")
					If bReturn Then
						Fn_SchMgr_SchDeliverableCreate = False
						objScheduleDel.JavaButton("Cancel").Click
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add deliverable.[" + sDelName + "]")
						Set objScheduleDel  = Nothing 
						Set objTable = Nothing
						Set objDelTable = Nothing
						sErrorText = ""
						Exit Function
					End If

					If  sDelType <> "" Then
						objTable.ClickCell (sIndex), "#1","LEFT", "NONE"
						iTypeIndex = objScheduleDel.JavaList("SchDelList").GetItemIndex(sDelType)
						If Err.Number < 0 Then
							Fn_SchMgr_SchDeliverableCreate = False 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sDelType + " deliverable type not exist.")
							objScheduleDel.JavaButton("Cancel").Click
							Set objScheduleDel  = Nothing 
							Set objTable = Nothing
							Set objDelTable = Nothing
							sErrorText = ""
							Exit Function
						End If
						objScheduleDel.JavaList("SchDelList").Object.setSelectedIndex Cint(iTypeIndex)
					End If

				'open & close Open By Name dialog
				objScheduleDel.JavaTable("SchDeliverablesTable").ClickCell 0,"#3","LEFT","NONE"
				If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Open by Name").Exist(2)=true Then
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Open by Name").Close
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully closed Open By Name dialog")
				else
					Fn_SchMgr_SchDeliverableCreate = false
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Open By Name dialog does not exist")
					Exit function
				End If

			'click on Ok button of schedule deliverables  Dialog
			 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Deliverables").JavaButton("OK").Click micLeftBtn
			 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the OK button.")
			 Fn_SchMgr_SchDeliverableCreate = True
				
		End Select
	End If

	Set objScheduleDel  = Nothing 
	Set objTable = Nothing
	Set objDelTable = Nothing 
	Set objOpenByName = Nothing 
	sErrorText = ""
End Function
'********************************************************* Function do Revert  Assignment to Discpline *********************************************************************

'Function Name		:					Fn_SchMgr_RevertAssignDiscpline

'Description			 :		 		  This function is used to Revert  Assignment to Discpline .

'Parameters			   :	 			1.  sTaskName:Name of the task need to select.
'													2.asDiscplineMem :  Name of the member of Displine
											
'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Schedule membership pane  should be displayed .

'Examples				:				Fn_SchMgr_RevertAssignDiscpline("S1:T1","Rupali Palhade(AutoDisp1)")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							17-Jun-2010	          1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_RevertAssignDiscpline(sTaskName,sDiscplineMem)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_RevertAssignDiscpline"
   On Error Resume Next

   Dim bReturn,objRevAss
	Set objRevAss = Fn_SISW_PPM_GetObject("Revert Assignments")

	If Not objRevAss.Exist(SISW_MIN_TIMEOUT) Then
   
		 bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")
		If  bReturn <> False Then
			Fn_SchMgr_RevertAssignDiscpline = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected task " + sTaskName)
		ELse
			Fn_SchMgr_RevertAssignDiscpline = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task " + sTaskName)
			Set objRevAss = Nothing
			Exit Function
		End If
	
	
		bReturn = Fn_MenuOperation("Select","Schedule:Assignments:Revert Assignments to Discipline")
		If bReturn = True Then
			Fn_SchMgr_RevertAssignDiscpline = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Assignments:Revert Assignments to Discipline]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Assignments:Revert Assignments to Discipline]")
			Fn_SchMgr_RevertAssignDiscpline = False
			Set objRevAss = Nothing
			Exit Function
		End If

	End If 

	If objRevAss.Exist(SISW_MIN_TIMEOUT) Then
		objRevAss.JavaList("AssignmentList").Select sDiscplineMem

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select value " + sDiscplineMem + " of Revert assignments")
			Fn_SchMgr_RevertAssignDiscpline = False
			objRevAss.JavaButton("Cancel").Click micLeftBtn
			Set objRevAss = Nothing
			Exit Function
		End If

		objRevAss.JavaButton("OK").WaitProperty  "enabled",1,2000
		objRevAss.JavaButton("OK").Click micLeftBtn

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Revert assignments of " + sDiscplineMem)
			Fn_SchMgr_RevertAssignDiscpline = False
			objRevAss.JavaButton("Cancel").Click micLeftBtn
			Set objRevAss = Nothing
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully  Revert assignments of " + sDiscplineMem)
			Fn_SchMgr_RevertAssignDiscpline = True
		End If

	End  If 

	Set objRevAss = Nothing

End Function

'*********************************************************  Function modifies a dependency *********************************************************************

'Function Name		:					Fn_SchMgr_TaskDependency  

'Description			 :		 		  This Function modifies a dependency.

'Parameters			   :	 			1.  sMode:- The input to be passed to the argument can be one of  the following; Menu/RMB
'													2.sAction:- The action to be performed on the task;Modify/Delete
'												   3.sTskName:- The of the task for which the dependency is to be edited.
'												  4.sPreDecessorTask :- Task name of PreDecessor
'												 5.sSuccessorTask : - Task name of Successor
'												6. sDepType:- The argument contains the value to which the current Dependency is edited to .
'											   7.sLag : - The argument contains the value of the current Dependency lag.
											
'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Schedule membership pane  should be displayed .

'Examples				:			 Fn_SchMgr_TaskDependency("RMB","Modify","Testch:t2","","t1","","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							18-Jun-2010	          1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_TaskDependency(sMode,sAction,sTskName,sPreDecessorTask,sSuccessorTask,sDepType,sLag)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_TaskDependency"
   On Error Resume Next 
   Dim bReturn,objDepend,iItemIndex,sValue
   Set objDepend = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Dependencies")

   If Not objDepend.Exist(SISW_MIN_TIMEOUT) Then
	   Select Case sMode
			Case "Menu"
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTskName,"","","")
				If  bReturn <> False Then
					Fn_SchMgr_TaskDependency = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected task " + sTskName)
				ELse
					Fn_SchMgr_TaskDependency = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task " + sTskName)
					Set objDepend = Nothing
					Exit Function
				End If
'Modified by Omkar... date 05 May 2011
				bReturn = Fn_MenuOperation("Select","Schedule:Link:Dependencies")
				If bReturn = True Then
					Fn_SchMgr_TaskDependency = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Dependencies.]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Dependencies.]")
					Fn_SchMgr_TaskDependency = False
					Set objDepend = Nothing 
					Exit Function
				End If
	
			Case "RMB"
				bReturn =  Fn_SchMgr_SchTable_NodeOperation("PopupMenu", sTskName, "", "", "Dependencies")
				If bReturn = True Then
					Fn_SchMgr_TaskDependency = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked RMB Menu [Dependencies]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked RMB Menu [Dependencies]")
					Fn_SchMgr_TaskDependency = False
					Set objDepend = Nothing
					Exit Function
				End If
		End Select

   End If

   If objDepend.Exist(SISW_MIN_TIMEOUT) Then
		If sPreDecessorTask <> "" Then
			objDepend.JavaList("PredecessorList").Select sPreDecessorTask
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select value of predecessor  " + sPreDecessorTask)
				Fn_SchMgr_TaskDependency = False
				objDepend.JavaButton("Close").Click micLeftBtn
				Set objDepend = Nothing
				Exit Function
			End If

		End If

		If sSuccessorTask <> "" Then
			objDepend.JavaList("SuccessorList").Select sSuccessorTask
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select value of successor  " + sSuccessorTask)
				Fn_SchMgr_TaskDependency = False
				objDepend.JavaButton("Close").Click micLeftBtn
				Set objDepend = Nothing
				Exit Function
			End If 
		End If   

		Select Case sAction
			Case "Modify"
				objDepend.JavaButton("Edit").WaitProperty "enabled",1,20000
				objDepend.JavaButton("Edit").Click micLeftBtn
				If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").Exist(SISW_MIN_TIMEOUT) Then

					If sDepType <> "" Then
						iItemIndex = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaList("DepType").GetItemIndex(sDepType)
						JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaList("DepType").Object.setSelectedIndex Cint(iItemInd)
						'Vallri [14-Jul-2010] - Added to work on Tc8.3_0616 Build
						JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaList("DepType").Select sDepType
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Task Dependency Type " + sDepType)
							Fn_SchMgr_TaskDependency = False
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("Cancel").Click
							objDepend.JavaButton("Close").Click micLeftBtn
							Set objDepend = Nothing
							Exit Function
						End If 
					End If
'  Added by Nilesh to handle Cancel button click operation
					If instr(1,sLag,"~")>0 Then

							aLag=Split(sLag,"~",-1,1)

							If aLag(0) <> "" Then
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaEdit("Lag").Object.setText aLag(0)
							End If

							If aLag(1)="Cancel" Then
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("Cancel").WaitProperty "enabled",1,20000
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("Cancel").Click micLeftBtn
							Else
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("OK").WaitProperty "enabled",1,20000
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("OK").Click micLeftBtn
							End If
				Else

							If sLag<> "" Then
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaEdit("Lag").Object.setText sLag
								wait 3
							End If
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("OK").WaitProperty "enabled",1,20000
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("OK").Click micLeftBtn
							
				End If
							wait(3)
							bReturn = Fn_SchMgr_DialogMsgVerify("Scheduling Error","","OK") 
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Handle Scheduling Error Error dialog.")
								Fn_SchMgr_TaskDependency = FALSE
								sErrorText = ""
							'	JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
							'	JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
							'	JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow sIndex
								Exit Function
							End If

							bReturn = Fn_SchMgr_DialogMsgVerify("UpdateDependenciesRunner","","OK") 
							If bReturn Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Handle Update Dependencies Error dialog.")
								Fn_SchMgr_TaskDependency = FALSE
								objDepend.JavaButton("Close").Click micLeftBtn
								Set objDepend = Nothing

'								Fn_SchMgr_SchTable_NodeOperation = FALSE
'								sErrorText = ""
'								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").Object.editCellAt Cint(sIndex),Cint(scolIndex)
'								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaEdit("SchTableCellEdit").Object.setFocusable True
'								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow sIndex
								Exit Function
							End If
									
					'Handle out of boundries error.
					bReturn = Fn_SchMgr_SchedulingErrorVerify("", "OK")
					If bReturn Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error - Attempting to schedule out of boundries.")
						Fn_SchMgr_TaskDependency = FALSE
						objDepend.JavaButton("Close").Click micLeftBtn
						Set objDepend = Nothing
						Exit Function
					End If

					objDepend.JavaButton("Close").Click micLeftBtn
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Edit Task  " + sTskName)
						Fn_SchMgr_TaskDependency = False
					   JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("Cancel").Click
					   objDepend.JavaButton("Close").Click micLeftBtn
						Set objDepend = Nothing
						Exit Function
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Edit Task  " + sTskName)
						 Fn_SchMgr_TaskDependency = True
					End If 
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Edit Dependency dialog does not exist")
					Fn_SchMgr_TaskDependency = False
					objDepend.JavaButton("Close").Click micLeftBtn
					Set objDepend = Nothing
					Exit Function
				End If

			Case "Delete"
				objDepend.JavaButton("Delete").WaitProperty "enabled",1,20000
				objDepend.JavaButton("Delete").Click micLeftBtn
				JavaDialog("Confirmation").SetTOProperty "title","Confirm"
				If JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
					JavaDialog("Confirmation").JavaButton("Yes").Click micLeftBtn
				End If   

				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to remove dependency between tasks." )
					Fn_SchMgr_TaskDependency = False
					objDepend.JavaButton("Close").Click micLeftBtn
					Set objDepend = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully  remove dependency between tasks.")
					Fn_SchMgr_TaskDependency = True
					objDepend.JavaButton("Close").Click micLeftBtn
				End If 

				Case "Verify"
				objDepend.JavaButton("Edit").WaitProperty "enabled",1,20000
				objDepend.JavaButton("Edit").Click micLeftBtn
				If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").Exist(SISW_MIN_TIMEOUT) Then

					If sDepType <> "" Then
						sValue = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaList("DepType").GetROProperty("value")

						If Trim(sDepType) = Trim(sValue) Then
							Fn_SchMgr_TaskDependency = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verify Task Dependency Type " + sDepType)
						Else
						    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify Task Dependency Type " + sDepType)
							Fn_SchMgr_TaskDependency = False
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("Cancel").Click
							objDepend.JavaButton("Close").Click micLeftBtn
							Set objDepend = Nothing
							Exit Function
						End If 
					End If

					If sLag <> "" Then
						sValue = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaEdit("Lag").GetROProperty("value")
						If Trim(sLag) = Trim(sValue) Then
							Fn_SchMgr_TaskDependency = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verify Task Dependency lag " + sLag)
						Else
						    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify Task Dependency lag " + sLag)
							Fn_SchMgr_TaskDependency = False
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("Cancel").Click
							objDepend.JavaButton("Close").Click micLeftBtn
							Set objDepend = Nothing
							Exit Function
						End If 
					End If

					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("Cancel").Click
					objDepend.JavaButton("Close").Click micLeftBtn

					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Edit Task  " + sTskName)
						Fn_SchMgr_TaskDependency = False
					   JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Dependency").JavaButton("Cancel").Click
					   objDepend.JavaButton("Close").Click micLeftBtn
						Set objDepend = Nothing
						Exit Function
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verify Task dependency  " + sTskName)
						 Fn_SchMgr_TaskDependency = True
					End If 
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Edit Dependency dialog does not exist")
					Fn_SchMgr_TaskDependency = False
					objDepend.JavaButton("Close").Click micLeftBtn
					Set objDepend = Nothing
					Exit Function
				End If

		End Select

   End If

	Set objDepend = Nothing
End Function


'*********************************************************  Function Baselines a task *********************************************************************

'Function Name		:					Fn_SchMgr_BaselineTask  

'Description			 :		 		  The Function Baselines a task,to a particular Schedule Baseline.

'Parameters			   :	 			1. sTaskName:- The name of the task that has to be baselined.
'													2.sBaselineName:- The name of the bvaseline against which the  task has to be baselined.

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Schedule membership pane  should be displayed .

'Examples				:			 Fn_SchMgr_BaselineTask("TestCh:T1","Baseline1")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							18-Jun-2010	          1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_BaselineTask(sTaskName,sBaselineName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_BaselineTask"
   On Error Resume Next
   Dim bReturn,objTaskBaseline,iIndex
   Set objTaskBaseline = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Task Baseline")

   If Not objTaskBaseline.Exist(SISW_MIN_TIMEOUT)Then

	   bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sTaskName,"","","")

	   If  bReturn <> False Then
			Fn_SchMgr_BaselineTask = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected task " + sTaskName)
		Else
			Fn_SchMgr_BaselineTask = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task " + sTaskName)
			Set objTaskBaseline = Nothing
			Exit Function
		End If

		bReturn = Fn_MenuOperation("Select","Schedule:Baseline Task")
		If bReturn = True Then
			Fn_SchMgr_BaselineTask = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule-->Baseline Task.]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule-->Baseline Task.]")
			Fn_SchMgr_BaselineTask = False
			Set objTaskBaseline = Nothing 
			Exit Function
		End If
   End If

   If objTaskBaseline.Exist(SISW_MIN_TIMEOUT)Then
	    iIndex = objTaskBaseline.JavaList("BaselineList").GetItemIndex(sBaselineName)
	   objTaskBaseline.JavaList("BaselineList").Object.setSelectedIndex Cint(iIndex) 

	   If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select value of choose Baseline  " + sBaselineName)
			Fn_SchMgr_BaselineTask = False
			objTaskBaseline.JavaButton("Cancel").Click micLeftBtn 
			Set objTaskBaseline = Nothing
			Exit Function
		End If 

		objTaskBaseline.JavaButton("OK").WaitProperty "enabled",1,20000
		objTaskBaseline.JavaButton("OK").Click micLeftBtn

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Task Baseline  " + sBaselineName)
			Fn_SchMgr_BaselineTask = False
			objTaskBaseline.JavaButton("Cancel").Click micLeftBtn 
			Set objTaskBaseline = Nothing
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected Task Baseline  " + sBaselineName)
			Fn_SchMgr_BaselineTask = True
		End If 
   End If

   Set objTaskBaseline = Nothing
End Function

'*********************************************************		Function to verify  dialog error message.	***********************************************************************

'Function Name		:					Fn_SchMgr_DialogMsgVerify

'Description			 :		 		  This function is used to get Schedule table Row Index.

'Parameters			   :	 			1.  sTitle:Title of dialog.
'													2. sMsg : Message to verify. (Optional)
'													3. sButton : Button Name.
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:				Fn_SchMgr_DialogMsgVerify("Update Schedule Deliverable Error","Schedule deliverables should have unique names.","OK")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							      23-Jun-2010	   1.0
'										Sushma						20-Jun-2013	   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_DialogMsgVerify(sTitle,sMsg,sButton) 

		Dim dicErrorInfo, bReturn 
		Set dicErrorInfo = CreateObject("Scripting.Dictionary")
		With dicErrorInfo 
		 .Add "Title", sTitle
		 .Add "Message", sMsg		 
		 .Add "Button", sButton		 
		End with
		Fn_SchMgr_DialogMsgVerify = Fn_SISW_SchMgr_ErrorVerify(dicErrorInfo)


End Function    

'*********************************************   Function validates the message in the "Apply Constraint" dialog .	***************************************************

'Function Name		:					Fn_SchMgr_ApplyConstraintConfirm  

'Description			 :		 		  The function validates the message in the "Apply Constraint" dialog .

'Parameters			   :	 			1.  sMessage:- The confirmation message that has to be verified.
'													2.sButton:- The button to be clicked to dismiss the dialog.
											
'Return Value		   : 			 True/False

'Pre-requisite			:		 	 The "Apply Constraint?" dialog is present.

'Examples				:			 Fn_SchMgr_ApplyConstraintConfirm("Message need to verify","Yes") 

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Rupali				 24-Jun-2010             1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vandana Patel		 21-Jun-2012             1.0			Modified object hierarchy
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ApplyConstraintConfirm(sMessage,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ApplyConstraintConfirm"
   On Error Resume Next
   Dim sActMsg,objConstrnt
    
   Fn_SchMgr_ApplyConstraintConfirm = True
	Set objConstrnt=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ApplyConstraintConfirm")
	If objConstrnt.Exist(SISW_MIN_TIMEOUT)=False Then
		Set objConstrnt=JavaDialog("ApplyConstraintConfirm")
		If objConstrnt.Exist(SISW_MIN_TIMEOUT)=False Then
			Fn_SchMgr_ApplyConstraintConfirm = False 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Apply Constraint dialog does not exist.")
			Exit Function 
		End If
	End If
		If sMessage  <> "" Then
			sActMsg = JavaDialog("ApplyConstraintConfirm").JavaObject("MsgObject").Object.getText()
			If instr(sActMsg, sMessage) > 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Validated Message [" + sMessage + "] on [ Apply Constraint ] Dialog")
			Else
				Fn_SchMgr_ApplyConstraintConfirm = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Validate Message [" + sMessage + "] on [ Apply Constraint ] Dialog")
			End If
		End If 

		If Ucase(sButton) = "YES" Then
			JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ApplyConstraintConfirm").JavaButton("Yes").WaitProperty "enabled",1,20000
			JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ApplyConstraintConfirm").JavaButton("Yes").Click micLeftBtn
		ElseIf  Ucase(sButton) = "NO" Then 
			JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ApplyConstraintConfirm").JavaButton("No").WaitProperty "enabled",1,20000
			JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ApplyConstraintConfirm").JavaButton("No").Click micLeftBtn
		End If 

		If Err.Number < 0 Then
			Fn_SchMgr_ApplyConstraintConfirm = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Button [" + sButton + "] on [ Apply Constraint ] Dialog")
			Exit Function 
	Else 
		Fn_SchMgr_ApplyConstraintConfirm = True 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked On Button [" + sButton + "] on [ Apply Constraint ] Dialog.")
		 End If
End Function

'*********************************************    Function performs the insertion of one Schedule template into another		**********************************************

'Function Name		:					Fn_SchMgr_InsertTemplates

'Description			 :		 		  The function performs the insertion of one Schedule template into another.

'Parameters			   :	 			1.  sAction:- The input to this parameter can be one of the following; Add/Remove/Seaarch
'													2.sNodeName:- The Schedule under which the Template has to be inserted.
'												   3.aTemplateName:- The name of the Scheduel Template that has   to be inserted.
'												 4 sOption : Template Copy/Reference.
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Stchdule Manager window should be displayed .

'Examples				:				Fn_SchMgr_InsertTemplates("Add","S1",arr,"Copy")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										    Rupali					   24-Jun-2010            1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_InsertTemplates(sAction,sNodeName,aTemplateName,sOption)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_InsertTemplates"
   On Error Resume Next
   Dim bReturn,objSchTemplate,iCounter

   'Set objSchTemplate = JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Insert Templates")

      Set objSchTemplate =JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Insert Templates")  '' Modified By Vidya 05/07/2012

   If Not objSchTemplate.Exist(SISW_MIN_TIMEOUT) Then
	   bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sNodeName,"","","")

	   If  bReturn <> False Then
			Fn_SchMgr_InsertTemplates = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected schedule  " + sNodeName)
		Else
			Fn_SchMgr_InsertTemplates = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  schedule  " + sNodeName)
			Set objSchTemplate = Nothing
			Exit Function
		End If

		bReturn = Fn_MenuOperation("Select","Schedule:Insert Schedule")
		If bReturn = True Then
			Fn_SchMgr_InsertTemplates = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule-->Insert Schedule.]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule-->Insert Schedule.]")
			Fn_SchMgr_InsertTemplates = False
			Set objSchTemplate = Nothing 
			Exit Function
		End If
   End If

'Added by Nilesh to handle Warning dialog on 10-Jan-2012

	If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Warning").Exist Then
			JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Warning").JavaButton("OK").Click
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Handle ' Warning' Dialog")
				Fn_SchMgr_InsertTemplates = False
				Exit Function
			End If
	End If


   If objSchTemplate.Exist(SISW_MIN_TIMEOUT) Then		
	   Call Fn_ReadyStatusSync(2)
		If sOption <> "" Then
			Select Case sOption
				Case "Reference"
					objSchTemplate.JavaRadioButton("InsetOption").SetTOProperty "attached text","Reference"
					objSchTemplate.JavaRadioButton("InsetOption").Set "ON"
				Case "Copy"
					objSchTemplate.JavaRadioButton("InsetOption").SetTOProperty "attached text","Copy"
					objSchTemplate.JavaRadioButton("InsetOption").Set "ON"
				Case Else 
					Fn_SchMgr_InsertTemplates = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Wrong value pass to sOption parameter.")
					Set objSchTemplate = Nothing
					Exit Function 
			End Select
		End If  

		'Search Dialog handle.
		If JavaWindow("ScheduleManagerWindow").JavaWindow("Search").Exist(SISW_MIN_TIMEOUT) Then
			JavaWindow("ScheduleManagerWindow").JavaWindow("Search").JavaButton("OK").Click  micLeftBtn
		End If

		Call Fn_ReadyStatusSync(2)

		If objSchTemplate.JavaButton("LoadAll").Exist(SISW_MIN_TIMEOUT)  Then
			If objSchTemplate.JavaButton("LoadAll").GetROProperty("enabled") = "1" Then
				objSchTemplate.JavaButton("LoadAll").Click  micLeftBtn
				Call Fn_ReadyStatusSync(2)
		   End If 
		End If 

	   Select Case sAction
			Case "Add"  

				If IsArray(aTemplateName)Then
					For iCounter = 0 to Ubound(aTemplateName)
						objSchTemplate.JavaTree("Available Templates").ExtendSelect  "#0:" + aTemplateName(iCounter) 
						If Err.Number < 0 Then
							Fn_SchMgr_InsertTemplates = False 
							objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aTemplateName(iCounter) + " from available templates.")
							Set objSchTemplate = Nothing
							Exit Function 
						End If
					Next   
					objSchTemplate.JavaButton("Add").WaitProperty "enabled",1,20000
					objSchTemplate.JavaButton("Add").Click micLeftBtn
					objSchTemplate.JavaButton("OK").WaitProperty "enabled",1,30000
					objSchTemplate.JavaButton("OK").Click micLeftBtn  

					bReturn = Fn_SchMgr_DialogMsgVerify("Error","","OK")

					If bReturn Then
						Fn_SchMgr_InsertTemplates = False 
						objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule template does not fall between the dates of the master schedule template  " + sNodeName )
						Set objSchTemplate = Nothing
						Exit Function
					End If     

					If Err.Number < 0 Then
						Fn_SchMgr_InsertTemplates = False 
						objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to added schedule template which need to inset into " + sNodeName)
						Set objSchTemplate = Nothing
						Exit Function
					Else
						Fn_SchMgr_InsertTemplates = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully  added schedule template which need to inset into " + sNodeName)
					End If
				Else 
					Fn_SchMgr_InsertTemplates = False 
					objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule template need to be add is not vallid.")
					 Set objSchTemplate = Nothing
					Exit Function 
				End If

			Case "Remove"
				If IsArray(aTemplateName)Then
					For iCounter = 0 to Ubound(aTemplateName)
					    objSchTemplate.JavaList("Selected Templates").ExtendSelect aTemplateName(iCounter)

						If Err.Number < 0 Then
							Fn_SchMgr_InsertTemplates = False 
							objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aTemplateName(iCounter) + " from selected Templates.")
							Set objSchTemplate = Nothing
							Exit Function 
						End If
					Next   
					objSchTemplate.JavaButton("Remove").WaitProperty "enabled",1,20000
					objSchTemplate.JavaButton("Remove").Click micLeftBtn

					 If JavaDialog("Delete").Exist(SISW_MIN_TIMEOUT) Then
						JavaDialog("Delete").JavaButton("Yes").Click micLeftBtn
					End If 

					objSchTemplate.JavaButton("OK").WaitProperty "enabled",1,20000
					objSchTemplate.JavaButton("OK").Click micLeftBtn

					If Err.Number < 0 Then
						Fn_SchMgr_InsertTemplates = False 
						objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to remove schedule from selected Templates." )
						Set objSchTemplate = Nothing
						Exit Function
					Else
						Fn_SchMgr_InsertTemplates = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove schedule from selected schedules")
					End If
				Else 
					Fn_SchMgr_InsertTemplates = False 
					objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule Template need to add is not vallid formate.")
					 Set objSchTemplate = Nothing
					Exit Function 
				End If

			Case "Search"
				If IsArray(aTemplateName)Then 
					For iCounter = 0 to Ubound(aTemplateName)
						objSchTemplate.JavaEdit("SearchText").Set aTemplateName(iCounter)
						objSchTemplate.JavaButton("Find").WaitProperty "enabled",1,20000
						objSchTemplate.JavaButton("Find").Click micLeftBtn

						If JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("Object Not Found").Exist(SISW_MIN_TIMEOUT) Then 
							JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("Object Not Found").JavaButton("OK").Click micLeftBtn
							Fn_SchMgr_InsertTemplates = False 
							objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No object found based upon search name." )
							Set objSchTemplate = Nothing
							Exit Function
						End If 

						objSchTemplate.JavaButton("Add").WaitProperty "enabled",1,20000
						objSchTemplate.JavaButton("Add").Click micLeftBtn

						If Err.Number < 0 Then
							Fn_SchMgr_InsertTemplates = False 
							objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aTemplateName(iCounter) + " from selected schedules templates.")
							Set objSchTemplate = Nothing
							Exit Function 
						End If
					Next

					objSchTemplate.JavaButton("OK").WaitProperty "enabled",1,20000
					objSchTemplate.JavaButton("OK").Click micLeftBtn

					If Err.Number < 0 Then
						Fn_SchMgr_InsertTemplates = False 
						objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to remove schedule from selected template." )
						Set objSchTemplate = Nothing
						Exit Function
					Else
						Fn_SchMgr_InsertTemplates = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove schedule template from selected templates")
					End If

				Else 
					Fn_SchMgr_InsertTemplates = False 
					objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule template need to add is not vallid.")
					Set objSchTemplate = Nothing
					Exit Function 
				End If

			Case "Cancel"
				If IsArray(aTemplateName)Then
					For iCounter = 0 to Ubound(aTemplateName)
						objSchTemplate.JavaTree("Available Templates").ExtendSelect  "#0:" + aTemplateName(iCounter) 
						If Err.Number < 0 Then
							Fn_SchMgr_InsertTemplates = False 
							objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aTemplateName(iCounter) + " from available templates.")
							Set objSchTemplate = Nothing
							Exit Function 
						End If
					Next   
					objSchTemplate.JavaButton("Add").WaitProperty "enabled",1,20000
					objSchTemplate.JavaButton("Add").Click micLeftBtn
					objSchTemplate.JavaButton("Cancel").WaitProperty "enabled",1,30000
					objSchTemplate.JavaButton("Cancel").Click micLeftBtn 

					If Err.Number < 0 Then
						Fn_SchMgr_InsertTemplates = False 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [" + Cancel + "] button After Addition of Templates")
						Set objSchTemplate = Nothing
						Exit Function
					Else
						Fn_SchMgr_InsertTemplates = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked [" + Cancel + "] button After Addition of Templates")
					End If
				Else 
					Fn_SchMgr_InsertTemplates = False 
					objSchTemplate.JavaButton("Cancel").Click micLeftBtn 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule template need to be add is not vallid.")
					 Set objSchTemplate = Nothing
					Exit Function 
				End If

	   End Select
   End If 

   Set objSchTemplate = Nothing
End Function

'***************************************** Functions inserts a specified schedule inside another schedule **************************************************************

'Function Name		:					Fn_SchMgr_InsertSchedule

'Description			 :		 		  The functions inserts a specified schedule inside another schedule.

'Parameters			   :	 			1. sSchName :- The name of the schedule that has to be inserted.
'													2.sAction:- The Action to be performed;   i.e Add/Remove the schedules.
'												  3.aSchSearch :- Schedule to be Add/Remove from list.

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Schedule membership pane  should be displayed .

'Examples				:			  Fn_SchMgr_InsertSchedule("Remove","Testch",arr)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Rupali					 23-Jun-2010	          1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_InsertSchedule(sAction,sSchName,aSchSearch)
   GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_InsertSchedule"
   On Error Resume Next 
   Dim bReturn,objInsetSch,iCounter, aTskName, sBtnName,WshShell
   Dim sRoot

	if instr(sSchName, "~") > 0 Then
		aTskName = split(sSchName, "~", -1, 1)
		sSchName = aTskName(0)
		sBtnName = aTskName(1)
	Else
		sBtnName = "Yes"
	End If


   Set objInsetSch = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Insert Schedules")

   If Not objInsetSch.Exist(SISW_MIN_TIMEOUT) Then
	   bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sSchName,"","","")

	   If  bReturn <> False Then
			Fn_SchMgr_InsertSchedule = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected schedule  " + sSchName)
		Else
			Fn_SchMgr_InsertSchedule = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  schedule  " + sSchName)
			Set objInsetSch = Nothing
			Exit Function
		End If

		bReturn = Fn_MenuOperation("Select","Schedule:Insert Schedule")
		If bReturn = True Then
			Fn_SchMgr_InsertSchedule = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule-->Insert Schedule.]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule-->Insert Schedule.]")
			Fn_SchMgr_InsertSchedule = False
			Set objInsetSch = Nothing 
			Exit Function
		End If
   End If

'	sCaption = "Warning"
'	JavaDialog("Confirmation").SetTOProperty "title", sCaption
'	If JavaDialog("Confirmation").Exist(3) Then
'		JavaDialog("Confirmation").JavaButton(sBtnName).click micLeftBtn
'	End If

If JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Warning").Exist(SISW_MIN_TIMEOUT) Then

	JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Warning").JavaButton("OK").Click
	If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Handle ' Warning' Dialog")
		Fn_SchMgr_InsertSchedule = False
		Exit Function
	End If
End If

   If objInsetSch.Exist(5) Then
	   Call Fn_ReadyStatusSync(2)

	   Select Case sAction
		 	Case "Add"
				If IsArray(aSchSearch)Then

					If objInsetSch.JavaButton("LoadAll").Exist(5) Then
						If objInsetSch.JavaButton("LoadAll").GetROProperty("enabled") = 1 then objInsetSch.JavaButton("LoadAll").Click
					End If
				   Call Fn_ReadyStatusSync(10)
				   sRoot = objInsetSch.JavaTree("AvailableSchedules").GetItem(0)
				   wait(2)					   
					objInsetSch.JavaTree("AvailableSchedules").Activate
					For iCounter = 0 to Ubound(aSchSearch)
						objInsetSch.JavaTree("AvailableSchedules").ExtendSelect  sRoot + ":" + aSchSearch(iCounter)
						If Err.Number < 0 Then
							Exit For 
						End If
					Next  
					If Err.Number < 0 Then
						Fn_SchMgr_InsertSchedule = False 
						objInsetSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aSchSearch(iCounter) + " from available schedules.")
						Set objInsetSch = Nothing
						Exit Function 
					End If

					objInsetSch.JavaButton("Add").WaitProperty "enabled",1,20000
					objInsetSch.JavaButton("Add").Object.doClick(1)
					objInsetSch.JavaButton("OK").WaitProperty "enabled",1,30000
					'objInsetSch.JavaButton("OK").Object.doClick(1)
					objInsetSch.JavaButton("OK").Click micLeftBtn

'					For iCounter=0 to 4
'													Set WshShell = CreateObject("WScript.Shell")
'											        WshShell.SendKeys "{TAB}"  ' Pressing the Tab button 5 times brings the focus on the "Add"  button and Pressing "Enter" after that adds the schedule to the Selected Scheduled list
	'				Next
'					wait 2
					
'					WshShell.SendKeys "{ENTER}" 
'					Set WshShell = nothing

					bReturn = Fn_SchMgr_DialogMsgVerify("Error","","OK")

					If bReturn Then
						Fn_SchMgr_InsertSchedule = False 
						objInsetSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule does not fall between the dates of the master schedule  " + sSchName )
						Set objInsetSch = Nothing
						Exit Function
					End If

					bReturn = Fn_SchMgr_DialogMsgVerify("Scheduling Error","","OK")

					If bReturn Then
						Fn_SchMgr_InsertSchedule = False 
						objInsetSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule is not able to add  " + sSchName )
						Set objInsetSch = Nothing
						Exit Function
					End If

					If Err.Number < 0 Then
						Fn_SchMgr_InsertSchedule = False 
						objInsetSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to added schedule which need to inset into " + sSchName)
						Set objInsetSch = Nothing
						Exit Function
					Else
						Fn_SchMgr_InsertSchedule = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully  added schedule which need to inset into " + sSchName)
					End If
				Else 
					Fn_SchMgr_InsertSchedule = False 
					objInsetSch.JavaButton("Cancel").Click micLeftBtn 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule need to add is not vallid.")
					 Set objInsetSch = Nothing
					Exit Function 
				End If
		
			Case "Remove"

				If IsArray(aSchSearch)Then
					For iCounter = 0 to Ubound(aSchSearch)
					    objInsetSch.JavaList("SelectedSchedules").ExtendSelect aSchSearch(iCounter)
						If Err.Number < 0 Then
							Fn_SchMgr_InsertSchedule = False 
							objInsetSch.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aSchSearch(iCounter) + " from selected schedules.")
							Set objInsetSch = Nothing
							Exit Function 
						End If
					Next   
					objInsetSch.JavaButton("Remove").WaitProperty "enabled",1,20000
					objInsetSch.JavaButton("Remove").Click micLeftBtn

					 If JavaDialog("Delete").Exist(5) Then
						JavaDialog("Delete").JavaButton("Yes").Click micLeftBtn
					End If 

					objInsetSch.JavaButton("OK").WaitProperty "enabled",1,20000
					objInsetSch.JavaButton("OK").Click micLeftBtn

					If Err.Number < 0 Then
						Fn_SchMgr_InsertSchedule = False 
						objInsetSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to remove schedule from selected schedules" )
						Set objInsetSch = Nothing
						Exit Function
					Else
						Fn_SchMgr_InsertSchedule = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove schedule from selected schedules")
					End If
				Else 
					Fn_SchMgr_InsertSchedule = False 
					objInsetSch.JavaButton("Cancel").Click micLeftBtn 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule need to add is not vallid.")
					 Set objInsetSch = Nothing
					Exit Function 
				End If

				Case "Search"   
					If IsArray(aSchSearch)Then 
						For iCounter = 0 to Ubound(aSchSearch)
							objInsetSch.JavaEdit("SearchText").Set aSchSearch(iCounter)
							objInsetSch.JavaButton("Find").WaitProperty "enabled",1,20000
							objInsetSch.JavaButton("Find").Click micLeftBtn

							If JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("Object Not Found").Exist(5) Then 
								JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("Object Not Found").JavaButton("OK").Click micLeftBtn
								Fn_SchMgr_InsertSchedule = False 
								objInsetSch.JavaButton("Cancel").Click micLeftBtn 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No object found based upon search name." )
								Set objInsetSch = Nothing
								Exit Function
							End If 

							objInsetSch.JavaButton("Add").WaitProperty "enabled",1,20000
							objInsetSch.JavaButton("Add").Click micLeftBtn

							If Err.Number < 0 Then
								Fn_SchMgr_InsertSchedule = False 
								objInsetSch.JavaButton("Cancel").Click micLeftBtn 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aSchSearch(iCounter) + "from selected schedules.")
								Set objInsetSch = Nothing
								Exit Function 
							End If
						Next

						objInsetSch.JavaButton("OK").WaitProperty "enabled",1,20000
						objInsetSch.JavaButton("OK").Click micLeftBtn

						If Err.Number < 0 Then
							Fn_SchMgr_InsertSchedule = False 
							objInsetSch.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to remove schedule from selected schedules" )
							Set objInsetSch = Nothing
							Exit Function
						Else
							Fn_SchMgr_InsertSchedule = True 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove schedule from selected schedules")
						End If

					Else 
						Fn_SchMgr_InsertSchedule = False 
						objInsetSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule need to add is not vallid.")
						 Set objInsetSch = Nothing
						Exit Function 
					End If
	   End Select
	 End If

    Set objInsetSch = Nothing
End Function

'***************************************** Function creates an adds rates,edits the added rates,deletes the added rates **********************************************

'Function Name		:					Fn_SchMgr_RateModifiers

'Description			 :		 		  The function creates an adds rates,edits the added rates,deletes the added rates.

'Parameters			   :	 			1. sAction :- The Action to be performed from one of the following;  Add/Modify/Delete/Verify.
'													 2.sName:- The name to be specified.
'												    3.sModifierRate:- One of the values to be selected form the Drop- down (Multiplier/Rate).
'												  4.sRate:- The rate to be applied.
'											     5.sCurrency:- The currency in which the rate is specified.
'												6. sSchName : Name of schedule need to select.

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Login as CostDBA 

'Examples				:			 Fn_SchMgr_RateModifiers("Modify","Testch","OPq3","Multiplier","4","","OPQ")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Rupali					   25-Jun-2010	          1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_RateModifiers(sAction,sSchName,sName,sModifierRate,sRate,sCurrency,sNewName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_RateModifiers"
    On Error Resume Next 
    Dim bReturn,objRateModifier,objTable,sRow,sIndex,sActualVal,aAction,sButton
	Dim iLen, sLetter, i
	Dim objEdit, objEditChild

    Set objRateModifier = JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers")
	Set objTable = JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers").JavaTable("RateModTable")

	If Not objRateModifier.Exist(5) Then
	   bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sSchName,"","","")

	   If  bReturn <> False Then
			Fn_SchMgr_RateModifiers = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected schedule  " + sSchName)
		Else
			Fn_SchMgr_RateModifiers = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  schedule  " + sSchName)
			Set objRateModifier = Nothing
			Exit Function
		End If

		bReturn = Fn_MenuOperation("Select","Schedule:Rate Modifiers")
		If bReturn = True Then
			Fn_SchMgr_RateModifiers = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule-->Rate Modifiers.]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule-->Rate Modifiers.]")
			Fn_SchMgr_RateModifiers = False
			Set objRateModifier = Nothing 
			Exit Function
		End If
   End If

   If objRateModifier.Exist(5) Then
				If instr(1,sAction,":") > 0 Then
					aAction = split(sAction,":",-1,1)    
					sAction=aAction(0)
					sButton = aAction(1)
				End If
	   Select Case sAction

	 		Case "Add"
				objRateModifier.JavaButton("Add").WaitProperty "enabled",1,20000
				objRateModifier.JavaButton("Add").Click micLeftBtn
				Wait 1
				If Err.Number < 0 Then
					Fn_SchMgr_RateModifiers = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add Rate modifiers " + sName)
					objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
					objRateModifier.JavaButton("Cancel").Click micLeftBtn
					Set objRateModifier = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If

				sRow = objTable.GetROProperty( "rows")
				sIndex = Cint(sRow) - 1

				If sName <> "" Then
					'by shreyas 29-12-2011
					'SetCellData method was not working on build 20111207  in Manage Rate Modifier Dialog

					'objTable.ActivateCell sIndex,"Name"
				    'objRateModifier.JavaEdit("RateTblNameEdit").set sName
				    '[TC1122-2016011300-28_01_2016-VivekA-Maintenance] - Added to set name in Name Column edit box
				    objTable.ActivateCell sIndex,"Name"
					If objRateModifier.JavaEdit("RateTblNameEdit").Exist(1) Then
						objRateModifier.JavaEdit("RateTblNameEdit").Type sName				
					Else
						Set objEdit = Description.Create()
						objEdit("Class Name").value = "JavaEdit"
						Set objEditChild = objRateModifier.JavaTable("RateModTable").ChildObjects(objEdit)
						'sValue = objEditChild.Count
						objEditChild(0).Set sName
					End If
				    
					'objRateModifier.JavaEdit("RateTblNameEdit").Activate
'''''''''''''
'					objTable.ActivateCell sIndex,"Name" 
'					objRateModifier.JavaEdit("RateTblNameEdit").SetTOProperty "Index","1" 
'					objRateModifier.JavaEdit("RateTblNameEdit").SetTOProperty "attached text","" 
'					objRateModifier.JavaEdit("RateTblNameEdit").Set sName
'					objTable.ClickCell 0,0
'					objRateModifier.JavaEdit("RateTblNameEdit").Object.setText sName
'					objTable.ClickCell 0,0
'
'					By Vallari - 27-Sept-2011
'					objTable.Click 5,5,"LEFT"
'					objTable.SetCellData sIndex,"Name", sName

					'By Vallari - 30-Sept-2011
'					iLen = Len(sName)
'					objTable.SelectCell sIndex,"Name"
'					For i = 1 to iLen
'						sLetter = Mid(sName, i, 1)
'						If Asc(sLetter) >= 33 AND Asc(sLetter) <= 43 Then
'							objTable.SendKey Lcase(sLetter), micShift
'						ElseIf Asc(sLetter) = 58 OR Asc(sLetter) = 60 Then
'							objTable.SendKey Lcase(sLetter), micShif
'						ElseIf Asc(sLetter) >= 62 AND Asc(sLetter) <= 90 Then
'							objTable.SendKey Lcase(sLetter), micShift
'						ElseIf Asc(sLetter) >= 94 AND Asc(sLetter) <= 96 Then
'							objTable.SendKey Lcase(sLetter), micShift
'						ElseIf Asc(sLetter) >= 123 AND Asc(sLetter) <= 127 Then
'							objTable.SendKey Lcase(sLetter), micShift
'						Else
'							objTable.SendKey sLetter
'						End If
'					Next

				End If

				If sModifierRate <> "" Then
						'by shreyas 29-12-2011
						'SetCellData method was not working on build 20111207  in Manage Rate Modifier Dialog
					objTable.SelectCell sIndex,"Modifier Type"
					wait 1
					objRateModifier.JavaList("RateTblCombo").Select sModifierRate

					'By Vallari - 27-Sept-2011
'					objTable.Click 5,5,"LEFT"
'					objTable.SetCellData sIndex,"Modifier Type", sModifierRate
				End If

				If sRate <> "" Then
						'by shreyas 29-12-2011
						'SetCellData method was not working on build 20111207  in Manage Rate Modifier Dialog
'						Added code by JotibaT - UFT14.52 -10-May-19
						objTable.ActivateCell sIndex,"Rate"
						wait 1
						If objRateModifier.JavaEdit("Text").Exist(1) Then
							objRateModifier.JavaEdit("Text").Type sRate				
						Else
							Set objEdit = Description.Create()
							objEdit("Class Name").value = "JavaEdit"
							Set objEditChild = objRateModifier.JavaTable("RateModTable").ChildObjects(objEdit)
							'sValue = objEditChild.Count
							objEditChild(0).Set sRate
						End If
		    
'					JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers").JavaTable("RateModTable").ActivateCell sIndex,2
'					wait 1
'					JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers").JavaEdit("Text").set sRate
					'JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers").JavaEdit("Text").Activate

					'By Vallari - 27-Sept-2011
'					objTable.Click 5,5,"LEFT"
'					objTable.SetCellData sIndex,"Rate", sRate

					'By Vallari - 30-Sept-2011
'					iLen = Len(sRate)
'					objTable.SelectCell sIndex,"Rate"
'					For i = 1 to iLen
'						sLetter = Mid(sRate, i, 1)
'						If Asc(sLetter) >= 33 AND Asc(sLetter) <= 43 Then
'							objTable.PressKey Lcase(sLetter), micShift
'						ElseIf Asc(sLetter) = 58 OR Asc(sLetter) = 60 Then
'							objTable.PressKey Lcase(sLetter), micShif
'						ElseIf Asc(sLetter) >= 62 AND Asc(sLetter) <= 90 Then
'							objTable.PressKey Lcase(sLetter), micShift
'						ElseIf Asc(sLetter) >= 94 AND Asc(sLetter) <= 96 Then
'							objTable.PressKey Lcase(sLetter), micShift
'						ElseIf Asc(sLetter) >= 123 AND Asc(sLetter) <= 127 Then
'							objTable.PressKey Lcase(sLetter), micShift
'						Else
'							objTable.PressKey sLetter
'						End If
'					Next
				End If

				If sCurrency <> "" Then
						'by shreyas 29-12-2011
						'SetCellData method was not working on build 20111207  in Manage Rate Modifier Dialog
					objTable.SelectCell sIndex,"Currency"
					wait 1
					objRateModifier.JavaList("RateTblCombo").Select sCurrency

					'By Vallari - 27-Sept-2011
'					objTable.Click 5,5,"LEFT"
'					objTable.SetCellData sIndex,"Currency", sCurrency
				End If

				If sButton="" Then
					objRateModifier.JavaButton("Finish").WaitProperty "enabled",1,20000
					objRateModifier.JavaButton("Finish").Click micLeftBtn	
				Else
					objRateModifier.JavaButton(sButton).WaitProperty "enabled",1,20000
					objRateModifier.JavaButton(sButton).Click micLeftBtn	
				End If

				If JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Error").Exist(5) Then
					JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Error").JavaButton("OK").Click micLeftBtn
					Fn_SchMgr_RateModifiers = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Manage rate error exist")
					objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
					objRateModifier.JavaButton("Cancel").Click micLeftBtn
					Set objRateModifier = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If

				If Err.Number < 0 Then
					Fn_SchMgr_RateModifiers = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add Rate modifiers " + sName)
					objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
					objRateModifier.JavaButton("Cancel").Click micLeftBtn
					Set objRateModifier = Nothing 
					Set objTable = Nothing 
					Exit Function
				Else
					Fn_SchMgr_RateModifiers = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Successfully added Rate modifiers " + sName)
				End If

			Case "Modify"

				sIndex = Fn_SchMgr_TableRowIndex(objTable,sName,"Name")

				If sIndex <> False Then

					If sName <> "" Then

						'By shreyas
					'the code was not working hence commented it
					'added the below code and is working

						'by shreyas 29-12-2011
						'SetCellData method was not working on build 20111207  in Manage Rate Modifier Dialog

						objTable.ActivateCell sIndex,"Name"
						objRateModifier.JavaEdit("RateTblNameEdit").Set ""
						Wait 0.500
						objRateModifier.JavaEdit("RateTblNameEdit").Type sNewName
	'					objRateModifier.JavaEdit("RateTblNameEdit").Activate

						'By Vallari - 27-Sept-2011
'						objTable.Click 5,5,"LEFT"
'						objTable.SetCellData sIndex,"Name", sName

						'objTable.ActivateCell sIndex,"Name" 
						'objRateModifier.JavaEdit("RateTblNameEdit").SetTOProperty "Index","1" 
						'objRateModifier.JavaEdit("RateTblNameEdit").SetTOProperty "attached text","" 
						'objRateModifier.JavaEdit("RateTblNameEdit").Object.setText sNewName
						'objTable.ClickCell 0,0

						'By Vallari - 30-Sept-2011
'						iLen = Len(sName)
'						objTable.SelectCell sIndex,"Name"
'						For i = 1 to iLen
'							sLetter = Mid(sName, i, 1)
'							If Asc(sLetter) >= 33 AND Asc(sLetter) <= 43 Then
'								objTable.PressKey Lcase(sLetter), micShift
'							ElseIf Asc(sLetter) = 58 OR Asc(sLetter) = 60 Then
'								objTable.PressKey Lcase(sLetter), micShif
'							ElseIf Asc(sLetter) >= 62 AND Asc(sLetter) <= 90 Then
'								objTable.PressKey Lcase(sLetter), micShift
'							ElseIf Asc(sLetter) >= 94 AND Asc(sLetter) <= 96 Then
'								objTable.PressKey Lcase(sLetter), micShift
'							ElseIf Asc(sLetter) >= 123 AND Asc(sLetter) <= 127 Then
'								objTable.PressKey Lcase(sLetter), micShift
'							Else
'								objTable.PressKey sLetter
'							End If
'						Next
					End If
	
					If sModifierRate <> "" Then


						'By Nilesh - 2-Jan-2012
'						objTable.SelectCell sIndex,"Modifier Type"
'						wait 1
						objTable.ActivateCell sIndex,"Modifier Type"
						objRateModifier.JavaList("RateTblCombo").Select sModifierRate

						'By Vallari - 27-Sept-2011
'						objTable.Click 5,5,"LEFT"
'						objTable.SetCellData sIndex,"Modifier Type", sModifierRate
					End If
	
					If sRate <> "" Then
						'by shreyas 29-12-2011
						'SetCellData method was not working on build 20111207  in Manage Rate Modifier Dialog
'						Added code by JotibaT - UFT14.52 -10-May-19
						objTable.ActivateCell sIndex,"Rate"
						wait 1
						If objRateModifier.JavaEdit("Text").Exist(1) Then
						objRateModifier.JavaEdit("Text").Type sRate				
						Else
							Set objEdit = Description.Create()
							objEdit("Class Name").value = "JavaEdit"
							Set objEditChild = objRateModifier.JavaTable("RateModTable").ChildObjects(objEdit)
							'sValue = objEditChild.Count
							objEditChild(0).Set sRate
						End If
'					JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers").JavaTable("RateModTable").ActivateCell sIndex,2
'					wait 1
'					JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers").JavaEdit("Text").set sRate
'					JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers").JavaEdit("Text").Activate

						'By Vallari - 27-Sept-2011
'						objTable.Click 5,5,"LEFT"
'						objTable.SetCellData sIndex,"Rate", 

						'By Vallari - 30-Sept-2011
'						iLen = Len(sRate)
'						objTable.SelectCell sIndex,"Rate"
'						For i = 1 to iLen
'							sLetter = Mid(sRate, i, 1)
'							If Asc(sLetter) >= 33 AND Asc(sLetter) <= 43 Then
'								objTable.PressKey Lcase(sLetter), micShift
'							ElseIf Asc(sLetter) = 58 OR Asc(sLetter) = 60 Then
'								objTable.PressKey Lcase(sLetter), micShif
'							ElseIf Asc(sLetter) >= 62 AND Asc(sLetter) <= 90 Then
'								objTable.PressKey Lcase(sLetter), micShift
'							ElseIf Asc(sLetter) >= 94 AND Asc(sLetter) <= 96 Then
'								objTable.PressKey Lcase(sLetter), micShift
'							ElseIf Asc(sLetter) >= 123 AND Asc(sLetter) <= 127 Then
'								objTable.PressKey Lcase(sLetter), micShift
'							Else
'								objTable.PressKey sLetter
'							End If
'						Next
					End If
	
					If sCurrency <> "" Then
						objTable.SelectCell sIndex,"Currency"
						wait 1
						objRateModifier.JavaList("RateTblCombo").Select sCurrency

						'By Vallari - 27-Sept-2011
'						objTable.Click 5,5,"LEFT"
'						objTable.SetCellData sIndex,"Currency", sCurrency
					End If
	
					objRateModifier.JavaButton("Finish").WaitProperty "enabled",1,20000
					objRateModifier.JavaButton("Finish").Click micLeftBtn
	
					If JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Error").Exist(5) Then
						JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Error").JavaButton("OK").Click micLeftBtn
						Fn_SchMgr_RateModifiers = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Manage rate error exist")
						objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
						objRateModifier.JavaButton("Cancel").Click micLeftBtn
						Set objRateModifier = Nothing 
						Set objTable = Nothing 
						Exit Function
					End If
	
					If Err.Number < 0 Then
						Fn_SchMgr_RateModifiers = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modify Rate modifiers " + sName)
						objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
						objRateModifier.JavaButton("Cancel").Click micLeftBtn
						Set objRateModifier = Nothing 
						Set objTable = Nothing 
						Exit Function
					Else
						Fn_SchMgr_RateModifiers = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Successfully modified Rate modifiers " + sName)
					End If
				Else 
					Fn_SchMgr_RateModifiers = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Row does not exist in table with name " + sName)
					objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
					objRateModifier.JavaButton("Cancel").Click micLeftBtn
					Set objRateModifier = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If

			Case "Delete"
				sIndex = Fn_SchMgr_TableRowIndex(objTable,sName,"Name")
				If sIndex <> False Then
					objTable.SelectRow sIndex
					objRateModifier.JavaButton("Delete").WaitProperty "enabled",1,20000
					objRateModifier.JavaButton("Delete").Click micLeftBtn
					objRateModifier.JavaButton("Finish").WaitProperty "enabled",1,20000
					objRateModifier.JavaButton("Finish").Click micLeftBtn

					If Err.Number < 0 Then
						Fn_SchMgr_RateModifiers = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to delete Rate modifiers " + sName)
						objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
						objRateModifier.JavaButton("Cancel").Click micLeftBtn
						Set objRateModifier = Nothing 
						Set objTable = Nothing 
						Exit Function
					Else
						Fn_SchMgr_RateModifiers = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Successfully delete Rate modifiers " + sName)
					End If
				Else 
					Fn_SchMgr_RateModifiers = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Row does not exist in table with name " + sName)
					objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
					objRateModifier.JavaButton("Cancel").Click micLeftBtn
					Set objRateModifier = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If

		Case "MultipleDelete"

			 If instr(1,sName,":") > 0 Then
			   aNames = split(sName,":",-1,1)    
			   For iCounter = 0 to Ubound(aNames)
					sIndex = Fn_SchMgr_TableRowIndex(objTable,aNames(iCounter),"Name")
					If sIndex <> False Then
						objTable.ExtendRow sIndex
					End if
			   Next
			 End If
			 objRateModifier.JavaButton("Delete").WaitProperty "enabled",1,20000
			 objRateModifier.JavaButton("Delete").Click micLeftBtn
			 objRateModifier.JavaButton("Finish").WaitProperty "enabled",1,20000
			 objRateModifier.JavaButton("Finish").Click micLeftBtn

			 If Err.Number < 0 Then
				Fn_SchMgr_RateModifiers = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to delete Rate modifiers " + sName)
				objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
				objRateModifier.JavaButton("Cancel").Click micLeftBtn
				Set objRateModifier = Nothing 
				Set objTable = Nothing 
				Exit Function
			 Else
				Fn_SchMgr_RateModifiers = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Successfully delete Rate modifiers " + sName)
			 End If

			Case "Verify"
				sIndex = Fn_SchMgr_TableRowIndex(objTable,sName,"Name")

				If sIndex <> False Then
					If sModifierRate<> "" Then
						sActualVal = objTable.GetCellData(sIndex, "Modifier Type")
						If sActualVal = sModifierRate Then
							Fn_SchMgr_RateModifiers = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify the modifier type  " + sModifierRate)
						Else 
							Fn_SchMgr_RateModifiers = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the modifier type  " + sModifierRate)
							objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
							objRateModifier.JavaButton("Cancel").Click micLeftBtn
							Set objRateModifier = Nothing 
							Set objTable = Nothing 
							Exit Function
						End If
					End If

					If sRate <> "" Then
						sActualVal = objTable.GetCellData(sIndex,"Rate")
						If sActualVal = sRate Then
							Fn_SchMgr_RateModifiers = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify the rate value  " + sRate)
						Else 
							Fn_SchMgr_RateModifiers = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the rate value   " + sRate)
							objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
							objRateModifier.JavaButton("Cancel").Click micLeftBtn
							Set objRateModifier = Nothing 
							Set objTable = Nothing 
							Exit Function
						End If
					End If  

					If sCurrency <> "" Then
						sActualVal = objTable.GetCellData(sIndex,"Currency")
						If sActualVal = sCurrency Then
							Fn_SchMgr_RateModifiers = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify the currency  " + sCurrency)
						Else 
							Fn_SchMgr_RateModifiers = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the currency   " + sCurrency)
							objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
							objRateModifier.JavaButton("Cancel").Click micLeftBtn
							Set objRateModifier = Nothing 
							Set objTable = Nothing 
							Exit Function
						End If
					End If

					If sIndex <> False Then
						Fn_SchMgr_RateModifiers = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify the name of Rate  " + sName)
					Else 
						Fn_SchMgr_RateModifiers = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the name of Rate    " + sName)
						objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
						objRateModifier.JavaButton("Cancel").Click micLeftBtn
						Set objRateModifier = Nothing 
						Set objTable = Nothing 
						Exit Function
					End If

					objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
					objRateModifier.JavaButton("Cancel").Click micLeftBtn
				Else 
					Fn_SchMgr_RateModifiers = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Row does not exist in table with name " + sName)
					objRateModifier.JavaButton("Cancel").WaitProperty "enabled",1,20000 
					objRateModifier.JavaButton("Cancel").Click micLeftBtn
					Set objRateModifier = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If 
	   End Select
   End If

   Set objRateModifier = Nothing 
   Set objTable = Nothing 
End Function

'***************************************** Function chooses the schedules that are to be inserted inside the Program view **********************************************

'Function Name		:					Fn_SchMgr_ChooseSchedules 

'Description			 :		 		  The Function chooses the schedules that are to be inserted inside the Program view.

'Parameters			   :	 			1. sAction :- The Action to be performed from one of the following;  Add/Remove
'													2.aSchName:- Array of  name of the Schedule that has to be searched  to be inserted.:- The name to be specified. 
'																				Full tree path need to pass.

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Peogramm view Dataset  should be created. 

'Examples				:			Fn_SchMgr_ChooseSchedules("Add",aSchName)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										      Rupali				  29-Jun-2010	         1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ChooseSchedules(sAction,aSchName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ChooseSchedules"
   On Error Resume Next
   Dim objChooseSch,bReturn,iCounter,aNodeName,sNodeName
	Set objChooseSch = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Choose Schedules...")	

	If Not objChooseSch.Exist(5) Then
		bReturn = Fn_MenuOperation("Select","Program:Choose Schedules")
		If bReturn = True Then
			Fn_SchMgr_ChooseSchedules = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Program->Choose Schedules.]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Program->Choose Schedules.]")
			Fn_SchMgr_ChooseSchedules = False
			Set objChooseSch = Nothing 
			Exit Function
		End If
	End If

	If objChooseSch.Exist(5) Then
'	   Wait(3)
	   Call Fn_ReadyStatusSync(30)
		If objChooseSch.JavaButton("LoadAll").Exist(5) Then
		   objChooseSch.JavaButton("LoadAll").WaitProperty "enabled",1,20000
		   If objChooseSch.JavaButton("LoadAll").GetROProperty("enabled") = 1 then objChooseSch.JavaButton("LoadAll").Click micLeftBtn
		End If

	   Call Fn_ReadyStatusSync(2)
        
	   Select Case sAction
			Case "Add"
				If IsArray(aSchName)Then
					For iCounter = 0 to Ubound(aSchName)
						aNodeName = Split(aSchName(iCounter),":")
						aNodeName(0) = objChooseSch.JavaTree("AvailableSchedules").GetItem(0)
						sNodeName = Join(aNodeName,":")
						objChooseSch.JavaTree("AvailableSchedules").ExtendSelect  sNodeName
						If Err.Number < 0 Then
							Fn_SchMgr_ChooseSchedules = False 
							objChooseSch.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " +sNodeName + " from available schedules.")
							Set objChooseSch = Nothing
							Exit Function 
						End If
					Next   
					objChooseSch.JavaButton("Add").WaitProperty "enabled",1,20000
					objChooseSch.JavaButton("Add").Click micLeftBtn
					objChooseSch.JavaButton("OK").WaitProperty "enabled",1,30000
					objChooseSch.JavaButton("OK").Click micLeftBtn  

					If Err.Number < 0 Then
						Fn_SchMgr_ChooseSchedules = False 
						objChooseSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to added current programm view  " + sNodeName)
						Set objChooseSch = Nothing
						Exit Function
					Else
						Fn_SchMgr_ChooseSchedules = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully added current programm view " + sNodeName)
					End If
				Else 
					Fn_SchMgr_ChooseSchedules = False 
					objChooseSch.JavaButton("Cancel").Click micLeftBtn 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule need to add is not vallid.")
					 Set objChooseSch = Nothing
					Exit Function 
				End If

			Case "Remove"

				If IsArray(aSchName)Then
					For iCounter = 0 to Ubound(aSchName)
					   objChooseSch.JavaList("SelectedSchedules").ExtendSelect aSchName(iCounter)

						If Err.Number < 0 Then
							Fn_SchMgr_ChooseSchedules = False 
							objChooseSch.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aSchName(iCounter) + " from selected schedules.")
							Set objChooseSch = Nothing
							Exit Function 
						End If
					Next   
					objChooseSch.JavaButton("Remove").WaitProperty "enabled",1,20000
					objChooseSch.JavaButton("Remove").Click micLeftBtn

					objChooseSch.JavaButton("OK").WaitProperty "enabled",1,20000
					objChooseSch.JavaButton("OK").Click micLeftBtn

					If Err.Number < 0 Then
						Fn_SchMgr_ChooseSchedules = False 
						objChooseSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to remove schedule from selected schedules" )
						Set objChooseSch = Nothing
						Exit Function
					Else
						Fn_SchMgr_ChooseSchedules = True 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove schedule from selected schedules")
					End If
				Else 
					Fn_SchMgr_ChooseSchedules = False 
					objChooseSch.JavaButton("Cancel").Click micLeftBtn 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule need to add is not vallid.")
					 Set objChooseSch = Nothing
					Exit Function 
				End If

				Case "Search"   
					If IsArray(aSchName)Then 
						For iCounter = 0 to Ubound(aSchName)
							aNodeName = Split(aSchName(iCounter),":")
							objChooseSch.JavaEdit("SearchText").Object.setText aSchName(iCounter)

							objChooseSch.JavaButton("Find").WaitProperty "enabled",1,20000
							objChooseSch.JavaButton("Find").Click micLeftBtn

							If JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("Object Not Found").Exist(5) Then 
								JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("Object Not Found").JavaButton("OK").Click micLeftBtn
								Fn_SchMgr_ChooseSchedules = False 
								objChooseSch.JavaButton("Cancel").Click micLeftBtn 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No object found based upon search name." )
								Set objChooseSch = Nothing
								Exit Function
							End If 

							objChooseSch.JavaButton("Add").WaitProperty "enabled",1,20000
							objChooseSch.JavaButton("Add").Click micLeftBtn

							If Err.Number < 0 Then
								Fn_SchMgr_ChooseSchedules = False 
								objChooseSch.JavaButton("Cancel").Click micLeftBtn 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aSchName(iCounter) + "from available schedules.")
								Set objChooseSch = Nothing
								Exit Function 
							End If
						Next

						objChooseSch.JavaButton("OK").WaitProperty "enabled",1,20000
						objChooseSch.JavaButton("OK").Click micLeftBtn

						If Err.Number < 0 Then
							Fn_SchMgr_ChooseSchedules = False 
							objChooseSch.JavaButton("Cancel").Click micLeftBtn 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select " + aSchName(iCounter) + " from available schedules" )
							Set objChooseSch = Nothing
							Exit Function
						Else
							Fn_SchMgr_ChooseSchedules = True 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully search and add schedules from available schedules.")
						End If

					Else 
						Fn_SchMgr_ChooseSchedules = False 
						objChooseSch.JavaButton("Cancel").Click micLeftBtn 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Schedule need to add is not vallid.")
						 Set objChooseSch = Nothing
						Exit Function 
					End If
	   End Select 
	End If 

	Set objChooseSch = Nothing 
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'The function is commented because of changes in the OR  and controls of Teamcenter 9.0  a function with the same name is inserted in the VBS
'Can be used again if the changes are reverted.
'By Shreyas
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'**********************************************    Function specifies the fixed costs for task/schedule   **********************************************************************************

'Function Name		:					Fn_SchMgr_FixedCostsAction 

'Description			 :		 		  The function specifies the fixed costs for task/schedule.

'Parameters			   :	 			1. sAction:- The action to be performed i.e;- Add/Modify/Verify.
'													2.sCostName: The name to be specified to the cost..
'												   3.sAccuralType: The accural Type to be selected from the Drop down.
'												  4.sEstimatedCost: The estimated cost for the task.
'												 5.sActualCost:- The Actual cost to be specified for the task.
'												6.sCurrency:- The currency in which the cost is to be specified.
'											   7.bUseActualCost:- The checkbox to be selected if the actual cost  is to be used.
'											  8.sBillCode:- The bill Code to be specified.
'											9.sBillSubCode:- The Bill Sub Code is to be specified.
'										   10.sBillType:- The Bill Type to be specified.
'										  11.sNodeName :- Name of the node to be selected.
'										 12.sMode :- Mode need to select (Menu/RMB)
'										13.sNewCostName : Modified name of Cost.
																			
'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Schedule Manager pane should be open.

'Examples				:			Fn_SchMgr_FixedCostsAction("Testch","Menu","Verify","NewTestCost","Start","$5.00","$4.00","USD","False","Training","Billable","Billed","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										      Rupali				   30-Jun-2010	         1.0
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_FixedCostsAction(sNodeName,sMode,sAction,sCostName,sAccuralType,sEstimatedCost,sActualCost,sCurrency,bUseActualCost,sBillCode,sBillSubCode,sBillType,sNewCostName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_FixedCostsAction"
   On Error Resume Next 
   Dim bReturn,objCost,objNewFixed,objTable,sIndex,sValue
	Set objCost = JavaWindow("ScheduleManagerWindow").JavaWindow("Costs")
	Set objNewFixed = JavaWindow("ScheduleManagerWindow").JavaWindow("Fixed Cost")
	Set objTable = JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaTable("FixedCostsTable")
	sValue = ""
    If Not objCost.Exist(5) Then

	   Select Case sMode
			Case "Menu"
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sNodeName,"","","")
				If  bReturn <> False Then
					Fn_SchMgr_FixedCostsAction = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected node " + sNodeName)
				ELse
					Fn_SchMgr_FixedCostsAction = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  node " + sNodeName)
					Set objCost = Nothing
					Set objNewFixed = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If

				bReturn = Fn_MenuOperation("Select","Schedule:Costs")
				If bReturn = True Then
					Fn_SchMgr_FixedCostsAction = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Costs.]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Costs.]")
					Fn_SchMgr_FixedCostsAction = False
					Set objCost = Nothing 
					Set objNewFixed = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If
	
			Case "RMB"
				bReturn =  Fn_SchMgr_SchTable_NodeOperation("PopupMenu", sNodeName, "", "", "Costs")
				If bReturn = True Then
					Fn_SchMgr_FixedCostsAction = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked RMB Menu [Costs]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked RMB Menu [Costs]")
					Fn_SchMgr_FixedCostsAction = False
					Set objCost = Nothing
					Set objNewFixed = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If
		End Select

   End If

    If objCost.Exist(5) Then
		Select Case sAction

			Case "Add"
				objCost.JavaButton("New").WaitProperty "enabled",1,20000
				objCost.JavaButton("New").Click micLeftBtn

				If objNewFixed.Exist(5) Then 
					If sCostName <> "" Then
						objNewFixed.JavaEdit("CostName").Set sCostName
					End If

					If sAccuralType <> "" Then
						objNewFixed.JavaList("AccrualType").Select sAccuralType
					End If
	
					If sEstimatedCost <> ""  Then
						objNewFixed.JavaEdit("EstimatedCost").Set sEstimatedCost
					End If
	
					If sActualCost <> "" Then
						objNewFixed.JavaEdit("ActualCost").Set sActualCost
					End If
	
					If  sCurrency <> "" Then
						objNewFixed.JavaList("Currency").Select sCurrency
						If Err.Number < 0 Then			'' some time Select Method  not working so added code to select currency value
								Err.clear
            					objNewFixed.JavaList("Currency").Type sCurrency
						End If
					End If 
	
					If bUseActualCost <> "" Then
						If Cbool(bUseActualCost) = True Then
							objNewFixed.JavaCheckBox("UseActualCost").Set "ON"
						ElseIf Cbool(bUseActualCost) = False Then
							objNewFixed.JavaCheckBox("UseActualCost").Set  "OFF"
						End If
					End If
	
					If sBillCode <> "" Then
						objNewFixed.JavaList("BillCode").Select sBillCode
					End If
	
					If sBillSubCode <> "" Then
						objNewFixed.JavaList("BillSub-code").Select sBillSubCode
					End If
	
					If sBillType <> "" Then
						objNewFixed.JavaList("BillType").Select sBillType
					End If
	
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add [Costs] " + sCostName)
						Fn_SchMgr_FixedCostsAction = False
						objNewFixed.JavaButton("Cancel").Click micLeftBtn
						objCost.JavaButton("Cancel").Click micLeftBtn
						Set objCost = Nothing
						Set objNewFixed = Nothing 
						Set objTable = Nothing 
						Exit Function
					End If
	
					objNewFixed.JavaButton("Finish").WaitProperty "enabled",1,20000
					objNewFixed.JavaButton("Finish").Click micLeftBtn
					objCost.JavaButton("Finish").WaitProperty "enabled",1,20000
					objCost.JavaButton("Finish").Click micLeftBtn
	
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add [Cost] " + sCostName)
						Fn_SchMgr_FixedCostsAction = False
						objNewFixed.JavaButton("Cancel").Click micLeftBtn
						objCost.JavaButton("Cancel").Click micLeftBtn
						Set objCost = Nothing
						Set objNewFixed = Nothing 
						Set objTable = Nothing 
						Exit Function
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully add [Cost] " + sCostName)
						Fn_SchMgr_FixedCostsAction = True
					End If
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fixed Cost  dialog does not exist.")
					Fn_SchMgr_FixedCostsAction = False
					Set objCost = Nothing
					Set objNewFixed = Nothing 
					Set objTable = Nothing 
					Exit Function
			    End If 

			Case "Modify"
				
				sIndex = Fn_SchMgr_TableRowIndex(objTable,sCostName, "Cost Name")
				If sIndex <> False Then
					objTable.ClickCell sIndex,"Cost Name" 
					objCost.JavaButton("Details").WaitProperty "enabled",1,20000
					objCost.JavaButton("Details").Click micLeftBtn
				
					If objNewFixed.Exist(5) Then 
						If sCurrency <> "" Then					
							objNewFixed.JavaList("Currency").Select sCurrency
							If Err.Number < 0 Then			'' some time Select Method  not working so added code to select currency value
								Err.clear
            					objNewFixed.JavaList("Currency").Type sCurrency
							End If			
						End If 
						
						If sNewCostName <> "" Then
							objNewFixed.JavaEdit("CostName").Set sNewCostName
						End If
	
						If sAccuralType <> "" Then
							objNewFixed.JavaList("AccrualType").Select sAccuralType
						End If
		
						If sEstimatedCost <> ""  Then
							objNewFixed.JavaEdit("EstimatedCost").Set sEstimatedCost
						End If
		
						If sActualCost <> "" Then
							objNewFixed.JavaEdit("ActualCost").Set sActualCost
						End If

						If bUseActualCost <> "" Then
							If Cbool(bUseActualCost) = True Then
								objNewFixed.JavaCheckBox("UseActualCost").Set "ON"
							ElseIf Cbool(bUseActualCost) = False Then
								objNewFixed.JavaCheckBox("UseActualCost").Set  "OFF"
							End If
						End If
		
						If sBillCode <> "" Then
							objNewFixed.JavaList("BillCode").Select sBillCode
						End If
		
						If sBillSubCode <> "" Then
							objNewFixed.JavaList("BillSub-code").Select sBillSubCode
						End If
		
						If sBillType <> "" Then
							objNewFixed.JavaList("BillType").Select sBillType
						End If
		
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modified [Costs] " + sCostName)
							Fn_SchMgr_FixedCostsAction = False
							objNewFixed.JavaButton("Cancel").Click micLeftBtn
							objCost.JavaButton("Cancel").Click micLeftBtn
							Set objCost = Nothing
							Set objNewFixed = Nothing 
							Set objTable = Nothing 
							Exit Function
						End If
		
						objNewFixed.JavaButton("Finish").WaitProperty "enabled",1,20000
						objNewFixed.JavaButton("Finish").Click micLeftBtn
						objCost.JavaButton("Finish").WaitProperty "enabled",1,20000
						objCost.JavaButton("Finish").Click micLeftBtn
		
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modified [Cost] " + sCostName)
							Fn_SchMgr_FixedCostsAction = False
							objNewFixed.JavaButton("Cancel").Click micLeftBtn
							objCost.JavaButton("Cancel").Click micLeftBtn
							Set objCost = Nothing
							Set objNewFixed = Nothing 
							Set objTable = Nothing 
							Exit Function
						Else 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully modified [Cost] " + sCostName)
							Fn_SchMgr_FixedCostsAction = True
						End If
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fixed Cost  dialog does not exist.")
						Fn_SchMgr_FixedCostsAction = False
						Set objCost = Nothing
						Set objNewFixed = Nothing 
						Set objTable = Nothing 
						Exit Function
					End If 
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sCostName +  " cost name does not exist in fixed coast table.")
					Fn_SchMgr_FixedCostsAction = False
					Set objCost = Nothing
					Set objNewFixed = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If

			Case "Verify"
				sIndex = Fn_SchMgr_TableRowIndex(objTable,sCostName, "Cost Name")
				If sIndex <> False Then
					objTable.ClickCell sIndex,"Cost Name" 
					objCost.JavaButton("Details").WaitProperty "enabled",1,20000
					objCost.JavaButton("Details").Click micLeftBtn
				
					If objNewFixed.Exist(5) Then 
						If sCostName <> "" Then
							If  sCostName = objNewFixed.JavaEdit("CostName").GetROProperty( "value") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Cost Name] value  " + sCostName)
								Fn_SchMgr_FixedCostsAction = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Cost Name] value  " + sCostName)
								Fn_SchMgr_FixedCostsAction = False
								objNewFixed.JavaButton("Cancel").Click micLeftBtn
								objCost.JavaButton("Cancel").Click micLeftBtn
								Set objCost = Nothing
								Set objNewFixed = Nothing 
								Set objTable = Nothing 
								Exit Function
							End If
						End If
	
						If sAccuralType <> "" Then
							If sAccuralType = objNewFixed.JavaList("AccrualType").GetROProperty( "value") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Accural Type] value  " + sAccuralType)
								Fn_SchMgr_FixedCostsAction = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Accural Type] value  " + sAccuralType)
								Fn_SchMgr_FixedCostsAction = False
								objNewFixed.JavaButton("Cancel").Click micLeftBtn
								objCost.JavaButton("Cancel").Click micLeftBtn
								Set objCost = Nothing
								Set objNewFixed = Nothing 
								Set objTable = Nothing 
								Exit Function
							End If
						End If
		
						If sEstimatedCost <> ""  Then
							If  sEstimatedCost = objNewFixed.JavaEdit("EstimatedCost").GetROProperty( "value") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Estimated Cost] value  " + sEstimatedCost)
								Fn_SchMgr_FixedCostsAction = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Estimated Cost] value  " + sEstimatedCost)
								Fn_SchMgr_FixedCostsAction = False
								objNewFixed.JavaButton("Cancel").Click micLeftBtn
								objCost.JavaButton("Cancel").Click micLeftBtn
								Set objCost = Nothing
								Set objNewFixed = Nothing 
								Set objTable = Nothing 
								Exit Function
							End If
						End If
		
						If sActualCost <> "" Then
							If  sActualCost = objNewFixed.JavaEdit("ActualCost").GetROProperty( "value") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Actual Cost ] value  " + sActualCost)
								Fn_SchMgr_FixedCostsAction = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Actual Cost] value  " + sActualCost)
								Fn_SchMgr_FixedCostsAction = False
								objNewFixed.JavaButton("Cancel").Click micLeftBtn
								objCost.JavaButton("Cancel").Click micLeftBtn
								Set objCost = Nothing
								Set objNewFixed = Nothing 
								Set objTable = Nothing 
								Exit Function
							End If
						End If
		
						If  sCurrency <> "" Then
							If sCurrency = objNewFixed.JavaList("Currency").GetROProperty( "value") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Currency] value  " + sCurrency)
								Fn_SchMgr_FixedCostsAction = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Currency] value  " + sCurrency)
								Fn_SchMgr_FixedCostsAction = False
								objNewFixed.JavaButton("Cancel").Click micLeftBtn
								objCost.JavaButton("Cancel").Click micLeftBtn
								Set objCost = Nothing
								Set objNewFixed = Nothing 
								Set objTable = Nothing 
								Exit Function
							End If
						End If 
		
						If bUseActualCost <> "" Then
							If Cbool(bUseActualCost) = True Then
								If objNewFixed.JavaCheckBox("UseActualCost").GetROProperty( "value") = "1" Then 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Use Actual Cost] value  " + UseActualCost)
								    Fn_SchMgr_FixedCostsAction = True
								Else 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Use Actual Cost] value  " + UseActualCost)
									Fn_SchMgr_FixedCostsAction = False
									objNewFixed.JavaButton("Cancel").Click micLeftBtn
									objCost.JavaButton("Cancel").Click micLeftBtn
									Set objCost = Nothing
									Set objNewFixed = Nothing 
									Set objTable = Nothing 
									Exit Function
								End If 
							ElseIf Cbool(bUseActualCost) = False Then
								If objNewFixed.JavaCheckBox("UseActualCost").GetROProperty( "value") = "0" Then 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Use Actual Cost] value  " + UseActualCost)
								    Fn_SchMgr_FixedCostsAction = True
								Else 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Use Actual Cost] value  " + UseActualCost)
									Fn_SchMgr_FixedCostsAction = False
									objNewFixed.JavaButton("Cancel").Click micLeftBtn
									objCost.JavaButton("Cancel").Click micLeftBtn
									Set objCost = Nothing
									Set objNewFixed = Nothing 
									Set objTable = Nothing 
									Exit Function
								End If
							End If
						End If
		
						If sBillCode <> "" Then
							If sBillCode = objNewFixed.JavaList("BillCode").GetROProperty( "value") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Bill Code] value  " + sBillCode)
								Fn_SchMgr_FixedCostsAction = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Bill Code] value  " + sBillCode)
								Fn_SchMgr_FixedCostsAction = False
								objNewFixed.JavaButton("Cancel").Click micLeftBtn
								objCost.JavaButton("Cancel").Click micLeftBtn
								Set objCost = Nothing
								Set objNewFixed = Nothing 
								Set objTable = Nothing 
								Exit Function
							End If
						End If
		
						If sBillSubCode <> "" Then
							If sBillSubCode = objNewFixed.JavaList("BillSub-code").GetROProperty( "value") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Bill SubCode] value  " + sBillSubCode)
								Fn_SchMgr_FixedCostsAction = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Bill SubCode] value  " + sBillSubCode)
								Fn_SchMgr_FixedCostsAction = False
								objNewFixed.JavaButton("Cancel").Click micLeftBtn
								objCost.JavaButton("Cancel").Click micLeftBtn
								Set objCost = Nothing
								Set objNewFixed = Nothing 
								Set objTable = Nothing 
								Exit Function
							End If
						End If
		
						If sBillType <> "" Then
							If sBillType = objNewFixed.JavaList("BillType").GetROProperty( "value") Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Bill Type] value  " + sBillType)
								Fn_SchMgr_FixedCostsAction = True
							Else 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Bill Type] value  " + sBillType)
								Fn_SchMgr_FixedCostsAction = False
								objNewFixed.JavaButton("Cancel").Click micLeftBtn
								objCost.JavaButton("Cancel").Click micLeftBtn
								Set objCost = Nothing
								Set objNewFixed = Nothing 
								Set objTable = Nothing 
								Exit Function
							End If
						End If
		
					    objNewFixed.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objNewFixed.JavaButton("Cancel").Click micLeftBtn
						objCost.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objCost.JavaButton("Cancel").Click micLeftBtn
		
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify")
							Fn_SchMgr_FixedCostsAction = False
							objNewFixed.JavaButton("Cancel").Click micLeftBtn
							objCost.JavaButton("Cancel").Click micLeftBtn
							Set objCost = Nothing
							Set objNewFixed = Nothing 
							Set objTable = Nothing 
							Exit Function
						End If
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fixed Cost  dialog does not exist.")
						Fn_SchMgr_FixedCostsAction = False
						Set objCost = Nothing
						Set objNewFixed = Nothing 
						Set objTable = Nothing 
						Exit Function
					End If 
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sCostName +  " cost name does not exist in fixed coast table.")
					Fn_SchMgr_FixedCostsAction = False
					Set objCost = Nothing
					Set objNewFixed = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If
			
			Case "GetValue"
				sIndex = Fn_SchMgr_TableRowIndex(objTable,sCostName, "Cost Name")
				If sIndex <> False Then
					objTable.ClickCell sIndex,"Cost Name" 
					objCost.JavaButton("Details").WaitProperty "enabled",1,20000
					objCost.JavaButton("Details").Click micLeftBtn
				
					If objNewFixed.Exist(5) Then 
						
						bReturn = objNewFixed.JavaEdit("CostName").GetROProperty("value")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Get [Cost Name] value  ")
						sValue = bReturn

	
						bReturn = objNewFixed.JavaList("AccrualType").GetROProperty( "value")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get [Accural Type] value")
						sValue = sValue+"~"+bReturn

		
						bReturn = objNewFixed.JavaEdit("EstimatedCost").GetROProperty( "value")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get [Estimated Cost] value")
						sValue = sValue+"~"+bReturn


						bReturn = objNewFixed.JavaEdit("ActualCost").GetROProperty( "value") 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get [Actual Cost ] value  " )
						sValue = sValue+"~"+bReturn


						bReturn = objNewFixed.JavaList("Currency").GetROProperty( "value")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get [Currency] value  " )
						sValue = sValue+"~"+bReturn


						bReturn = objNewFixed.JavaCheckBox("UseActualCost").GetROProperty( "value")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get [Use Actual Cost] value  " )
						sValue = sValue+"~"+bReturn


						bReturn = objNewFixed.JavaList("BillCode").GetROProperty( "value")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get [Bill Code] value  ")
						sValue = sValue+"~"+bReturn


						bReturn = objNewFixed.JavaList("BillSub-code").GetROProperty( "value")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get [Bill SubCode] value  " )
						sValue = sValue+"~"+bReturn


						bReturn = objNewFixed.JavaList("BillType").GetROProperty( "value")
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get [Bill Type] value  " )
						sValue = sValue+"~"+bReturn

		
					    objNewFixed.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objNewFixed.JavaButton("Cancel").Click micLeftBtn
						objCost.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objCost.JavaButton("Cancel").Click micLeftBtn
						Fn_SchMgr_FixedCostsAction = sValue
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fixed Cost dialog does not exist.")
						Fn_SchMgr_FixedCostsAction = False
						Set objCost = Nothing
						Set objNewFixed = Nothing 
						Set objTable = Nothing 
						Exit Function
					End If 
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sCostName +  " cost name does not exist in fixed cost table.")
					Fn_SchMgr_FixedCostsAction = False
					Set objCost = Nothing
					Set objNewFixed = Nothing 
					Set objTable = Nothing 
					Exit Function
				End If	
		End Select
  Else 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Cost dialog does not exist.")
		Fn_SchMgr_FixedCostsAction = False
		Set objCost = Nothing
		Set objNewFixed = Nothing 
		Set objTable = Nothing 
		Exit Function
	End If

	Set objCost = Nothing
	Set objNewFixed = Nothing 
	Set objTable = Nothing 
End Function


'**************************************************************    Filter setting to programm view  **************************************************************************

'Function Name		:					Fn_SchMgr_FilterSettings

'Description			 :		 		  The functions specifies the filter settings to be applied in the Program View in the Schedule Manager application.

'Parameters			   :	 			1. sAction :- sAction :- The action to be performed; Create/Edit
'													2.sIndex:-  Inedx at which modification need to do. (Starts with zero)
'												   3.sAnd/or :And or Or value need to select.
'												  4.sFieldName : Name of the field
'												 5.sCondition : Condition need to apply.
'												6.sValue : Value need  to put..(Multiple value should be , (Comma) seprated)

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Schedule Manager Pane should be displayed.

'Examples				:			 Fn_SchMgr_FilterSettings("1","And","Name","Equal To","TestName")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										      Rupali				  29-Jun-2010	         1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_FilterSettings(sIndex,sAndOr,sFieldName,sCondition,sValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_FilterSettings"
   On Error Resume Next 

	Dim bReturn,objFilterSet,sRow,objSch,aActDate,iCounter,bStatusFlag,iValCounter

	Set objFilterSet  = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Filter Settings")
									 
	bStatusFlag=True
	If Not objFilterSet.Exist(5) Then
		bReturn = Fn_MenuOperation("Select","Program:Filter")
		If bReturn = True Then
			Fn_SchMgr_FilterSettings = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Program->Filter.]")
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Program->Filter.]")
			Fn_SchMgr_FilterSettings = False
			objFilterSet.JavaButton("Cancel").Click micLeftBtn
			Set objFilterSet = Nothing 
			Exit Function
		End If
	End If

	If objFilterSet.Exist(5) Then
		sRow = sIndex
		If sIndex = "0" Then
			If sFieldName <> "" Then
				objFilterSet.JavaButton("DropDownBtn").SetTOProperty "Index","1"
				objFilterSet.JavaButton("DropDownBtn").Click 
				If sFieldName = "BLANK" Then
					sFieldName=" "
					bReturn = Fn_iComboSet(objFilterSet,sFieldName)
'					bReturn = Fn_UI_JavaObject_Click("Fn_SchMgr_FilterSettings", objFilterSet.JavaWindow("JavaWindow"), "BlankListItem", 4, 4, "LEFT")
				Else
					bReturn = Fn_iComboSet(objFilterSet,sFieldName)
				End If
				If Not bReturn Then	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set value of Field Name " + sFieldName)
					Fn_SchMgr_FilterSettings = False
					objFilterSet.JavaButton("Cancel").Click micLeftBtn
					Set objFilterSet = Nothing 
					Exit Function
				End If
			End If

			If sCondition <> "" Then
				objFilterSet.JavaButton("DropDownBtn").SetTOProperty "Index","2"
				objFilterSet.JavaButton("DropDownBtn").Click
				bReturn = Fn_iComboSet(objFilterSet,sCondition)
				If Not bReturn Then	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set value of Condition  " + sCondition)
					Fn_SchMgr_FilterSettings = False
					objFilterSet.JavaButton("Cancel").Click micLeftBtn
					Set objFilterSet = Nothing 
					Exit Function
				End If
			End If
		Else
			sIndex = (Cint(sIndex) * 3)
			If sAndOr <> "" Then
				objFilterSet.JavaButton("DropDownBtn").SetTOProperty "Index",sIndex
				objFilterSet.JavaButton("DropDownBtn").Click 
				bReturn = Fn_iComboSet(objFilterSet,sAndOr)
				If Not bReturn Then	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set value of Field Name " + sAndOr)
					Fn_SchMgr_FilterSettings = False
					objFilterSet.JavaButton("Cancel").Click micLeftBtn
					Set objFilterSet = Nothing 
					Exit Function
				End If
			End If
			
			If sFieldName <> "" Then
				objFilterSet.JavaButton("DropDownBtn").SetTOProperty "Index",(sIndex + 1)
				objFilterSet.JavaButton("DropDownBtn").Click 
				bReturn = Fn_iComboSet(objFilterSet,sFieldName)
				If Not bReturn Then	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set value of Field Name " + sFieldName)
					Fn_SchMgr_FilterSettings = False
					objFilterSet.JavaButton("Cancel").Click micLeftBtn
					Set objFilterSet = Nothing 
					Exit Function
				End If
			End If

			If sCondition <> "" Then
				objFilterSet.JavaButton("DropDownBtn").SetTOProperty "Index",(sIndex + 2)
				objFilterSet.JavaButton("DropDownBtn").Click
				bReturn = Fn_iComboSet(objFilterSet,sCondition)
				If Not bReturn Then	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set value of Condition  " + sCondition)
					Fn_SchMgr_FilterSettings = False
					objFilterSet.JavaButton("Cancel").Click micLeftBtn
					Set objFilterSet = Nothing 
					Exit Function
				End If
			End If
		End If   

		If sValue <> ""Then
			objFilterSet.JavaEdit("Value").SetTOProperty "Index",sRow 
			If objFilterSet.JavaEdit("Value").GetROProperty( "enabled") = "1" Then
				objFilterSet.JavaEdit("Value").Object.setText  sValue
			Else 
				objFilterSet.JavaButton("Browse").SetTOProperty "Index",sRow 
				objFilterSet.JavaButton("Browse").WaitProperty "enabled",1,20000
				objFilterSet.JavaButton("Browse").Click micLeftBtn 

				If sFieldName="Work Complete" Or sFieldName="Duration" Or sFieldName= "Work Estimate" or sFieldName="Task Duration" Then
					bReturn = Fn_SchMgr_HoursChooser(sValue)
					If bReturn = True Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set value of Total Hours " + sValue)
						Fn_SchMgr_FilterSettings = True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  set value of Total Hours " + sValue)
						Fn_SchMgr_FilterSettings = False
						objFilterSet.JavaButton("Cancel").Click micLeftBtn
						Set objFilterSet = Nothing 
						Exit Function
					End If
				ElseIf sFieldName="Actual Start Date" Or sFieldName="Actual Finish Date" Or sFieldName= "Start Date" Or sFieldName = "Finish Date" Then

					If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Date Chooser").Exist(5)=False Then
					     objFilterSet.JavaButton("Browse").Click		
                          If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on browse button.")
									Fn_SchMgr_FilterSettings = False
									Set objFilterSet = Nothing 
									Exit Function
					    	Else 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on browse button.")
									Fn_SchMgr_FilterSettings = True
							End If
					End If
					'Added code by Omkar to handle 'Between' Condition for Start date and End date field
					If instr(1,sValue,"~") Then
							sValue =  split(sValue,"~",-1,1)
					End If
					
					bStatusFlag = TRUE
					If IsArray(sValue) Then
								For iValCounter = 0  to UBound(sValue)
											 bReturn = Fn_SchMgr_DateChooser(sValue(iValCounter))
											If bReturn = false Then
												bStatusFlag = false
												Exit for
											End If
								Next
					 else
								bReturn = Fn_SchMgr_DateChooser(sValue)
								If bReturn = false Then
												bStatusFlag = false
								End If
					End If

					If bStatusFlag = True Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set value of Date " + sValue)
									Fn_SchMgr_FilterSettings = True
					Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  set value of Date  " + sValue)
									Fn_SchMgr_FilterSettings = False
									objFilterSet.JavaButton("Cancel").Click micLeftBtn
									Set objFilterSet = Nothing 
									Exit Function
					End If
                    				
				ElseIf sFieldName= "Priority" Or sFieldName = "Status" or "State" Then
					Set objSch = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("List of Values Chooser")
					If objSch.Exist(5) Then
						objSch.JavaButton("ListDrpDwnBtn").Click micLeftBtn
						bReturn = Fn_iComboSet(objSch,sValue)
						If Not bReturn Then	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select New set value  " + sValue)
							Fn_SchMgr_FilterSettings = False
							objSch.JavaButton("Cancel").Click micLeftBtn
							objFilterSet.JavaButton("Cancel").Click micLeftBtn
							Set objFilterSet = Nothing 
							Set objSch = Nothing
							Exit Function
						End If
						objSch.JavaButton("OK").Click micLeftBtn
						Set objSch = Nothing
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "List of Values Chooser dialog does not exist.")
						Fn_SchMgr_FilterSettings = False
						objFilterSet.JavaButton("Cancel").Click micLeftBtn
						Set objFilterSet = Nothing 
						Set objSch = Nothing
						Exit Function
					End If
				ElseIf sFieldName = "Work Complete Percent" Then
					Set objSch = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Percentage Chooser")
					aActDate = Split(sValue,",",-1,1)

					For iCounter = 0 to Ubound(aActDate)
						If objSch.Exist(5) Then
							objSch.JavaSpin("Percentage").Set aActDate(iCounter)
							objSch.JavaButton("OK").WaitProperty "enabled",1,20000
							objSch.JavaButton("OK").Click  micLeftBtn
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set value of Work Complete Percen " + aActDate(iCounter))
								Fn_SchMgr_FilterSettings = False
								Exit For
							End If
						Else 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Percentage Chooser dialog does not exist.")
							Fn_SchMgr_FilterSettings = False
							objFilterSet.JavaButton("Cancel").Click micLeftBtn
							Set objFilterSet = Nothing 
							Set objSch = Nothing
							Exit Function
						End If
					Next

					If Err.Number < 0 Then
						For iCounter = 0 to Ubound(aActDate)
							If objSch.Exist(5) Then
								objSch.JavaButton("Cancel").Click micLeftBtn
							End If 
						Next
						objFilterSet.JavaButton("Cancel").Click micLeftBtn
						Set objFilterSet = Nothing 
						Set objSch = Nothing
						Exit Function
					End If
					Set objSch = Nothing
				End If
			End If
		End If

		objFilterSet.JavaButton("OK").WaitProperty  "enabled",1,20000
		objFilterSet.JavaButton("OK").Click micLeftBtn

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set Filter settings.")
			Fn_SchMgr_FilterSettings = False
			objFilterSet.JavaButton("Cancel").Click micLeftBtn
			Set objFilterSet = Nothing 
			Exit Function
		Else 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set Filter settings.")
			Fn_SchMgr_FilterSettings = True
		End If

	 Else 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Filter Setting dialog does not exist.")
		Fn_SchMgr_FilterSettings = False
		Set objFilterSet = Nothing 
		Exit Function
	 End If

		'Handle Error Dialog for Filter Settings If Exists . . .
		If objFilterSet.JavaDialog("FilterSettingsError").Exist(5)=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:The Error Dialog Exist")
				Fn_SchMgr_FilterSettings = False
				objFilterSet.JavaDialog("FilterSettingsError").JavaButton("OK").Click
				Set objFilterSet = Nothing 
				Exit Function
		Else 
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Error Dialog does not Exist")
				Fn_SchMgr_FilterSettings = True
		End If	
	Set objFilterSet = Nothing 
End Function

'**************************************************************    Functions set Total Hours. **************************************************************************

'Function Name		:					Fn_SchMgr_HoursChooser

'Description			 :		 		  The functions set Total Hours.

'Parameters			   :	 			1. sTotalHours :-  Total hours need to set .

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Hour Chooser dialog should be displayed.

'Examples				:			 Fn_SchMgr_HoursChooser("2.02h")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										      Rupali				   01-July-2010	         1.0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_HoursChooser(sTotalHours)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_HoursChooser"
   On Error Resume Next

	Dim objHrChoose,aActDate,iCounter

	Set objHrChoose = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Hours Chooser")
	Set objHrChoosePlusBtn = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Hours Chooser").JavaButton("HoursChooserPlusButton")
	aActDate = Split(sTotalHours,",",-1,1)
	
'TC 12 (2017110800) - (27/12/17) - Sandip C - Modified code to choose hours and mins using buttons insted of directly setting value into Edit box, As per discussion with vallari S.
	For iCounter = 0 to Ubound(aActDate)
									
		aActDate1 = split(aActDate(iCounter),".")
		If uBound(aActDate1)>0 Then
			iTotalHrs = cint(aActDate1(0))
			iTotalMins = (cInt(Replace(aActDate1(1),"h","")*0.6))
		Else 
			iTotalHrs = (cInt(Replace(aActDate(iCounter),"h","")))
		End If
		
		
		If objHrChoose.Exist(5)Then
		'objHrChoose.JavaEdit("TotalHours").Set aActDate(iCounter)
		'Setting Hours
		If iTotalHrs > 0 Then
			objHrChoosePlusBtn.SetTOProperty "Index",1
			For iCnt = 1 To iTotalHrs 
				objHrChoosePlusBtn.Click
			Next
		End If
		wait 1
		'Setting Minutes
		If iTotalMins > 0 Then
			objHrChoosePlusBtn.SetTOProperty "Index",0
			For iCnt = 1 To iTotalMins 
				objHrChoosePlusBtn.Click
			Next
		End If
		objHrChoose.JavaButton("OK").WaitProperty "enabled",1,20000
		objHrChoose.JavaButton("OK").Click micLeftBtn

		If Err.number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set Total hours value " + aActDate(iCounter))
			Fn_SchMgr_HoursChooser = False
			Exit For
		Else
			Fn_SchMgr_HoursChooser = True  
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set Total hours value " + aActDate(iCounter))
		End If
	Else 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Hours Chooser Dialog does not exist.")
		Fn_SchMgr_HoursChooser = False 
		Set objHrChoose = Nothing 
		Exit Function 
	End If

	Next

	For iCounter = 0 to Ubound(aActDate)
		If objHrChoose.Exist(5)Then
			objHrChoose.JavaButton("Cancel").Click micLeftBtn
		End If
	Next

	Set objHrChoose = Nothing 
End Function 

'**************************************************************    Functions set Date. **********************************************************************************

'Function Name		:					Fn_SchMgr_DateChooser

'Description			 :		 		  The functions set Date .

'Parameters			   :	 			1. sDate: Date need to set .
'													Formate of date is dd-mm-yy (e.g.  06-September-2010")

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Date Chooser dialog should be displayed.

'Examples				:			 Fn_SchMgr_DateChooser("30-August-2010")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										      Rupali				   01-July-2010	         1.0
'										      Sushma			   16-Aug-2012	         1.0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_DateChooser(sDate)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_DateChooser"
   On Error Resume Next
	Dim aDate,aActDate,iCounter, sDateValue
	Dim objDateChooser
	aActDate = Split(sDate,",",-1,1)

	For iCounter = 0 To Ubound(aActDate)
			aDate = Split(aActDate(iCounter),"-",-1,1)
			
			If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Date Chooser").Exist(5)Then
				Set objDateChooser  = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Date Chooser")
			ElseIf JavaDialog("Date Chooser").Exist(5)Then
				Set objDateChooser  = JavaDialog("Date Chooser")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date Chooser Dialog does not exist.")
				Fn_SchMgr_DateChooser = False 
				Exit Function 
			End If
				
			sDateValue  =  aDate(1)& " " & aDate(0) & ", " & aDate(2)    'August 16, 2012
			objDateChooser.JavaEdit("Date").Set  sDateValue
			wait 1
			objDateChooser.JavaButton("OK").Click micLeftBtn

			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the date " + sDateValue)
				Fn_SchMgr_DateChooser = False 
				Exit For
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Date " +sDate)
				Fn_SchMgr_DateChooser = True  
			End If
		

	Next  

'Code commented by Omkar as the it fails while checking the between condition 

'	For iCounter = 0 To Ubound(aActDate)
'		If JavaDialog("Date Chooser").Exist(5)Then
'			JavaDialog("Date Chooser").JavaButton("Cancel").Click
'		End If
'	Next

End Function 


'*********************************************************		Function to  modify , Verify ,  to check IsEditable  Schedule  property.***********************************************************************

'Function Name		:					Fn_SchMgr_SummarySchPropertyOperations

'Description			 :		 		  This function is used to  modify , Verify ,  to check IsEditable  Summary Schedule  properties.

'Parameters			   :	 			sAction: Action need to perform.
'                                                   dicProperties: Dictionary object to hold property names and the values
'                                                   sButtonName: Button to be clicked after modifying properties
'													sConfirmationBtn : Name of button of confirmation dialog.
											
'Return Value		   : 			  True/False  

'Pre-requisite			:		 		Schedule manger panel need to be open.
'                                                   Scedule need to be selected.
'                                                   Properties dialog need to be displayed.

'Examples				:			Call Fn_SchMgr_SummarySchPropertyOperations("Modify",dicScheduleProperty,"OK","YES")
'												Call Fn_SchMgr_SummarySchPropertyOperations("IsEditable",dicScheduleProperty,"","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari							07-Jul-2010	   1.0

'										Pritam							16-Jan-2012    2.0             Added case "State" in Case "Verify"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_SummarySchPropertyOperations(sAction,dicProperties,sButtonName,sConfirmationBtn)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_SummarySchPropertyOperations"
   On Error Resume Next
   Dim dicCount , dicKeys , dicItems , iCounter,objWin ,bReturn, sDate, aDate
   Set  objWin =  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties")

	If  not objWin.Exist(3) Then
		bReturn = Fn_MenuOperation("Select","View:View Properties")
		If bReturn = False Then
				Fn_SchMgr_SummarySchPropertyOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [View:Properties]")
				objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
				objWin.JavaButton("Cancel").Click
				Set objWin = Nothing
				Exit Function
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [View:Properties]")
		End If
	End If
     
	 objWin.WaitProperty "displayed",1,20000
	 If  objWin.Exist(5) Then
		 If objWin.JavaRadioButton("ViewSummarySchProps").Exist(5) Then
				objWin.JavaRadioButton("ViewSummarySchProps").Set "ON"
				objWin.JavaButton("OK").WaitProperty "enabled",1,20000
				objWin.JavaButton("OK").Click
		 End If

		Call Fn_ReadyStatusSync(2)  

         dicCount  = dicProperties.Count
		 dicItems = dicProperties.Items
		 dicKeys = dicProperties.Keys

		Select Case  sAction
			Case "Modify"
				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "CustomerName" ,"CustomerNumber"
								objWin.JavaEdit(dicKeys(iCounter)).Set dicItems(iCounter)
								
							Case "StartDate" ,"FinishDate" 
								 JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaCheckBox(dicKeys(iCounter)).Object.setDate(dicItems(iCounter))

							Case "StatusDropDown" ,"TskStatusDrpDwn","PriorityDropDown"
								objWin.JavaButton(dicKeys(iCounter)).Click
								Wait(2)
								bReturn =  Fn_iComboSet(objWin,dicItems(iCounter))
								
							Case  "IsTemplate" ,"IsPublic","FinishDateSchedul", "IsPercentLinked","Published","NotificationsEnabled" 
								objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label",Cstr (dicItems(iCounter))
								If objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty("value") = "0" Then
									objWin.JavaRadioButton(dicKeys(iCounter)).Set "ON"
								End If

						End Select
						
						 If Err.Number < 0 Then	
							Fn_SchMgr_SummarySchPropertyOperations = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to modify " & dicKeys(iCounter)  & " schedule property.")
							objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
							objWin.JavaButton("Cancel").Click
							Set objWin = Nothing
							Exit Function
						 Else 
							Fn_SchMgr_SummarySchPropertyOperations = TRUE
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully modify " & dicKeys(iCounter)  & " schedule property.")
						End If
				   End If
				Next
				JavaDialog("Confirmation").SetTOProperty "title", "Confirmation"
				If Ucase(sButtonName) = "OK"  And Ucase(sConfirmationBtn) = "YES"Then
					objWin.JavaButton("OK").Click
					JavaDialog("Confirmation").WaitProperty "displayed",1,20000
					If JavaDialog("Confirmation").Exist(5) Then
						JavaDialog("Confirmation").JavaButton("Yes").Click
					End If
				ElseIf Ucase(sButtonName) = "APPLY" And Ucase(sConfirmationBtn) = "YES"Then
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Apply").Click
					JavaDialog("Confirmation").WaitProperty "displayed",1,20000
					If JavaDialog("Confirmation").Exist(5) Then
						JavaDialog("Confirmation").JavaButton("Yes").Click
					End If
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click
				ElseIf Ucase(sButtonName) = "OK" And Ucase(sConfirmationBtn) = "NO"Then
					objWin.JavaButton("OK").Click
					JavaDialog("Confirmation").WaitProperty "displayed",1,20000
					If JavaDialog("Confirmation").Exist(5) Then
						JavaDialog("Confirmation").JavaButton("No").Click
					End If
				ElseIf Ucase(sButtonName) = "APPLY" And Ucase(sConfirmationBtn) = "NO"Then
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Apply").Click
					JavaDialog("Confirmation").WaitProperty "displayed",1,20000
					If JavaDialog("Confirmation").Exist(5) Then
						JavaDialog("Confirmation").JavaButton("No").Click
					End If
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click
				End If
				

			Case "Verify"
				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "CustomerName" ,"CustomerNumber"
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "value") = dicItems(iCounter) Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.Close
									Set objWin = Nothing
									Exit Function
								End If
								
							Case "StartDate" ,"FinishDate" 
								sDate =  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaCheckBox(dicKeys(iCounter)).GetROProperty( "label")
								aDate = split(sDate, " ", -1,1)

								If Trim(aDate(0)) =  Trim(dicItems(iCounter)) Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.Close
									Set objWin = Nothing
									Exit Function
								End If

							Case "SchPriority" ,"SchStatus","TaskStatus" 
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "value") = dicItems(iCounter) Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
							Case "State"
								
								If objWin.JavaStaticText("State").GetROProperty("label") = dicItems(iCounter) Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case  "IsTemplate" ,"IsPublic","FinishDateSchedul", "IsPercentLinked","Published","NotificationsEnabled" 
'								If Lcase(objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty("label")) = Lcase(Cstr(dicItems(iCounter))) Then
'									Fn_SchMgr_SummarySchPropertyOperations = TRUE
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
'								Else
'									Fn_SchMgr_SummarySchPropertyOperations = FALSE
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
'									Exit Function
'								End If
								objWin.JavaRadioButton(dicKeys(iCounter)).SetTOProperty "label",Cstr(dicItems(iCounter))
								sValue = objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty( "value")
								If  sValue = "1" And  Lcase(Cstr(dicItems(iCounter))) = "true" Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								ElseIf sValue = "1" And  Lcase(Cstr(dicItems(iCounter))) = "false" Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property " & dicKeys(iCounter))
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter))
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

							Case "ScheduleMembers"
								Dim ItemCount,arrItem,iIndex,iIndexItem
								arrItem = Split(dicItems(iCounter),",",-1,1)
								objWin.JavaStaticText("BootomLink").Click  8,6,"LEFT"
								Wait(2)
								ItemCount = objWin.JavaList(dicKeys(iCounter)).GetROProperty("items count")

								If  Err.Number < 0 Then
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to click All bottom link")
									Set objWin = Nothing
									objWin.Close
									Exit Function
								End If
							
								For iIndexItem = 0 to Ubound(arrItem) 
									For iIndex = 0 to ItemCount - 1
										 If  objWin.JavaList(dicKeys(iCounter)).GetItem(iIndex) = arrItem(iIndexItem) Then
											 Exit For
										End If
									Next

									If  Cstr(iIndex) = Cstr(ItemCount) Then
										Fn_SchMgr_SummarySchPropertyOperations = FALSE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed  to verify the property " & dicKeys(iCounter) & " with value " &arrItem(iIndexItem))
										objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
										objWin.JavaButton("Cancel").Click
										Set objWin = Nothing
										Exit Function
									Else
										Fn_SchMgr_SummarySchPropertyOperations = TRUE
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Sucessfully verify the property  " & dicKeys(iCounter) & " with value " &arrItem(iIndexItem))
									End If
								Next
						End Select
				   End If
				Next
			  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click

			Case "IsEditable"

				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter) 

							Case "Name" ,"Description" , "CustomerName" ,"CustomerNumber"
								If objWin.JavaEdit(dicKeys(iCounter)).GetROProperty( "editable") = "1" Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case "StartDate" ,"FinishDate" 
								If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaCheckBox(dicKeys(iCounter)).GetROProperty( "enabled") =  "1" Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

							Case "StatusDropDown" ,"TskStatusDrpDwn","PriorityDropDown" 
								If objWin.JavaButton(dicKeys(iCounter)).GetROProperty( "enabled") = "1" Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If
								
							Case  "IsTemplate" ,"IsPublic","FinishDateSchedul", "IsPercentLinked","Published","NotificationsEnabled" 
								If objWin.JavaRadioButton(dicKeys(iCounter)).GetROProperty( "enabled") = "1" Then
									Fn_SchMgr_SummarySchPropertyOperations = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is editable")
								Else
									Fn_SchMgr_SummarySchPropertyOperations = FALSE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),dicKeys(iCounter) & " property is not editable")
									objWin.JavaButton("Cancel").WaitProperty "enabled",1,20000
									objWin.JavaButton("Cancel").Click
									Set objWin = Nothing
									Exit Function
								End If

						End Select
				   End If
				Next
			  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("ScheduleProperties").JavaButton("Cancel").Click

		End Select
	Else
		Fn_SchMgr_SummarySchPropertyOperations = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to displayed Properties dialog.")
	End If

	Set objWin = Nothing

End  Function

'*********************************************************		Function to  date in Schedule Manager Required Format	***********************************************************************

'Function Name		:					Fn_SchMgr_FormatDate

'Description			 :		 		  This function is used to get date in Schedule Manager Required Format.

'Parameters			   :	 			1.  sDate:date value to provide . Format required is (MM/dd/yyyy)  7/9/2010 or 17/9/2010
											
'Return Value		   : 				 Date in the format DD-MMM-yyyy (07-Jul-2010)

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:				 Fn_SchMgr_FormatDate(date)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari S.						09-June-2010	   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_FormatDate(sDate)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_FormatDate"
   Dim sDay, sMonth, sYear

   On Error Resume Next

   sDay = Day(sDate)
   sMonth = Month(sDate) 
   sYear = Year(sDate)

	If len(sDay) = 1 Then
		sDay = "0" + cstr(sDay)
	End If

   Select Case cStr(sMonth)
		Case "1"
			sMonth = "Jan"
		Case "2"
			sMonth = "Feb"
		Case "3"
			sMonth = "Mar"
		Case "4"
			sMonth = "Apr"
		Case "5"
			sMonth = "May"
		Case "6"
			sMonth = "Jun"
		Case "7"
			sMonth = "Jul"
		Case "8"
			sMonth = "Aug"
		Case "9"
			sMonth = "Sep"			
		Case "10"
		sMonth = "Oct"
		Case "11"
			sMonth = "Nov"
		Case "12"
			sMonth = "Dec"
   End Select

	 Fn_SchMgr_FormatDate = cstr(sDay) & "-" & cstr(sMonth) & "-" & cstr(sYear)
End Function

'*********************************************************		Function to  Set Schedule Calendar	***********************************************************************

'Function Name		:					Fn_SchMgr_ScheduleCalendarOperations

'Description			 :		 		  This function is used to set Schedule Calendar

'Parameters			   :	 			1.  sAction: Action to Execute
'												2. sSchName: Name of Schedule to be Selected
'												3. dicSchCalendar: Schedule Calendar Dictionary
											
'Return Value		   : 				 True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:				Call Fn_SchMgr_ScheduleCalendarOperations("SchHrCalendar", "hhh", dicSchCalendar)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari S.						23-June-2010	   1.0
'
'                                      Nilesh G.                      16-Dec-2011        2.0                    Added Verify Case

'										Pritam S.						27-Dec-2011		   3.0         Added Delete Case        
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_ScheduleCalendarOperations(sAction, sSchName, dicSchCalendar)
			GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ScheduleCalendarOperations"
			Dim bReturn, objCalWin,objWrkingDyExcep
			Dim aDays, objDailyWin,aValues
			Dim iCounter, iRow, iCol
			Dim aRowData, aCellData
			Dim sCheck,sHour,aHour


			On Error Resume Next
			
			Set objCalWin = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Schedule Calendar")
'			Set objWrkingDyExcep = JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Working Day Exception")
			Set objWrkingDyExcep = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Working Day Exception")
			
			
			If objCalWin.Exist(3) = False Then
				bReturn = Fn_SchMgr_SchTable_NodeOperation("Select" , sSchName , "" , "" , "")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Schedule [" + sSchName + "]")
					Fn_SchMgr_ScheduleCalendarOperations = False
					Set objCalWin = Nothing
					Exit Function
				End If

				bReturn = Fn_MenuOperation("Select", "Schedule:Schedule Calendar")
				If bReturn = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Operate Menu [Schedule:Schedule Calendar]")
					Fn_SchMgr_ScheduleCalendarOperations = False
					Set objCalWin = Nothing
					Exit Function
				End If
			End If

			If objCalWin.Exist(10) Then
				'Press {Esc} key to dismiss Sch Cal Info dialog
				objCalWin.PressKey "Esc"
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{ESC}"
				Set WshShell = Nothing
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find [Edit Schedule Calendar] Dialog")
				Fn_SchMgr_ScheduleCalendarOperations = False
				Set objCalWin = Nothing
				Exit Function
			End If

			Select Case sAction
				Case "SchHrCalendar"
					'Set Working Days for Schedule
					If dicSchCalendar.Item("OnWeekDays") <> "" Then
						aDays = split(dicSchCalendar.Item("OnWeekDays"), "~", -1, 1)

						'Check the Days Check-Boxes for the required week days
						For iCounter = 0 to Ubound(aDays)
							objCalWin.JavaCheckBox("DayCheck").SetTOProperty "attached text", aDays(iCounter)
							objCalWin.JavaCheckBox("DayCheck").Set "ON"
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Week Day [" + aDays(iCounter) + "] as ON")
								Fn_SchMgr_ScheduleCalendarOperations = False
								objCalWin.JavaButton("Cancel").Click micLeftBtn
								Set objCalWin = Nothing
								Exit Function
							End If
						Next
					End If

					'Set Non-Working Days for Schedule
					If dicSchCalendar.Item("OffWeekDays") <> "" Then
						aDays = split(dicSchCalendar.Item("OffWeekDays"), "~", -1, 1)

						'Check the Days Check-Boxes for the required week days
						For iCounter = 0 to Ubound(aDays)
							objCalWin.JavaCheckBox("DayCheck").SetTOProperty "attached text", aDays(iCounter)
							objCalWin.JavaCheckBox("DayCheck").Set "OFF"
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Week Day [" + aDays(iCounter) + "] as OFF")
								Fn_SchMgr_ScheduleCalendarOperations = False
								objCalWin.JavaButton("Cancel").Click micLeftBtn
								Set objCalWin = Nothing
								Exit Function
							End If
						Next
					End If

					Set objDailyWin = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Daily Defaults Details")
					'Set Working Hours for All the Working Days
					For iCounter = 1 to 7
						objCalWin.JavaButton("Details...").SetTOProperty "index", iCounter
						wait(1)
						If objCalWin.JavaButton("Details...").GetROProperty("enabled") Then
							objCalWin.JavaButton("Details...").Click micLeftBtn
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click [Details...] Button for Week Day [" + iCounter + "]")
								Fn_SchMgr_ScheduleCalendarOperations = False
								objCalWin.JavaButton("Cancel").Click micLeftBtn
								Set objCalWin = Nothing
								Set objDailyWin = Nothing
								Exit Function
							End If

							'Add Daily Details
							If objDailyWin.Exist(10) Then
								'Set Specific Times Radio ON
								objDailyWin.JavaRadioButton("SpecificTimes").Set "ON"
	
								'Set the Timings in the Table
								If dicSchCalendar.Item("DayHrDetails") <> "" Then
									aRowData = split(dicSchCalendar.Item("DayHrDetails"), ",", -1, 1)
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Data for Detail Hors not Provided Properly")
									Fn_SchMgr_ScheduleCalendarOperations = False
									objDailyWin.JavaButton("Cancel").Click micLeftBtn
									objCalWin.JavaButton("Cancel").Click micLeftBtn
									Set objCalWin = Nothing
									Set objDailyWin = Nothing
									Exit Function
								End If
	
								For iRow = 0 to Ubound(aRowData)
									objDailyWin.JavaTable("TimeTable").SelectRow iRow
									'Split Data for Entering into Individule Cell
									aCellData = split(aRowData(iRow), "~", -1,1)
									For iCol = 0 to Ubound(aCellData)
										objDailyWin.JavaTable("TimeTable").SelectRow iRow
										wait(1)
										objDailyWin.JavaTable("TimeTable").DoubleClickCell iRow, iCol, "LEFT"
										objDailyWin.JavaEdit("TableCellEdit").Set aCellData(iCol)
										objDailyWin.JavaEdit("TableCellEdit").Activate
										objDailyWin.JavaTable("TimeTable").SelectRow Cint(iRow+1)
										wait(1)
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Cell Data [" + aCellData(iCol) + "]")
											Fn_SchMgr_ScheduleCalendarOperations = False
											objDailyWin.JavaButton("Cancel").Click micLeftBtn
											objCalWin.JavaButton("Cancel").Click micLeftBtn
											Set objCalWin = Nothing
											Set objDailyWin = Nothing
											Exit Function
										End If
									Next
								Next
								'Click OK on Daily Details Dialog
								objDailyWin.JavaButton("OK").Click micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Detail Timing Data")
									Fn_SchMgr_ScheduleCalendarOperations = False
	'								objDailyWin.JavaButton("Cancel").Click micLeftBtn
									objCalWin.JavaButton("Cancel").Click micLeftBtn
									Set objCalWin = Nothing
									Set objDailyWin = Nothing
									Exit Function
								End If
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find [Daily Defaults Details] Dialog")
								Fn_SchMgr_ScheduleCalendarOperations = False
								objCalWin.JavaButton("Cancel").Click micLeftBtn
								Set objCalWin = Nothing
								Set objDailyWin = Nothing
								Exit Function
							End If
						End If
					Next

				'- - - - - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - -- - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - 
					Case "SetCalendar_WorkingDayExceptions"

					'Set  Working HH:MM  RadioButton
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Schedule Calendar").JavaRadioButton("Working HH:MM").Set
					wait(2)
					
					'Click on Deatils Button
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Schedule Calendar").JavaButton("Details...").SetTOProperty "Index",0
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Schedule Calendar").JavaButton("Details...").Click
					wait(2)

					'Set the Timings in the Table
					If dicSchCalendar.Item("WorkingDayHrDetails") <> "" Then
								aRowData = split(dicSchCalendar.Item("WorkingDayHrDetails"), ",", -1, 1)
					Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Data for Detail Hors not Provided Properly")
								Fn_SchMgr_ScheduleCalendarOperations = False
								objWrkingDyExcep.JavaButton("Cancel").Click micLeftBtn
								objCalWin.JavaButton("Cancel").Click micLeftBtn
								Set objCalWin = Nothing
								Set objDailyWin = Nothing
								Exit Function
					End If
	
					For iRow = 0 to Ubound(aRowData)
								objWrkingDyExcep.JavaTable("TimeTable").SelectRow iRow
								'Split Data for Entering into Individule Cell
								aCellData = split(aRowData(iRow), "~", -1,1)
								For iCol = 0 to Ubound(aCellData)
											objWrkingDyExcep.JavaTable("TimeTable").SelectRow iRow
											wait(1)
											objWrkingDyExcep.JavaTable("TimeTable").DoubleClickCell iRow, iCol, "LEFT"
											objWrkingDyExcep.JavaEdit("TableCellEdit").Set aCellData(iCol)
											objWrkingDyExcep.JavaEdit("TableCellEdit").Activate
											objWrkingDyExcep.JavaTable("TimeTable").SelectRow Cint(iRow+1)
											wait(1)
											If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Cell Data [" + aCellData(iCol) + "]")
													Fn_SchMgr_ScheduleCalendarOperations = False
													objWrkingDyExcep.JavaButton("Cancel").Click micLeftBtn
													objWrkingDyExcep.JavaButton("Cancel").Click micLeftBtn
													Set objWrkingDyExcep = Nothing
													Set objWrkingDyExcep = Nothing
													Exit Function
											End If
								Next
					Next
				Case "Verify"
					'Set Working Days for Schedule
								If dicSchCalendar.Item("OnWeekDays") <> "" Then
									aDays = split(dicSchCalendar.Item("OnWeekDays"), "~", -1, 1)
			
									'Check the Days Check-Boxes for the required week days
									For iCounter = 0 to Ubound(aDays)
										objCalWin.JavaCheckBox("DayCheck").SetTOProperty "attached text", aDays(iCounter)
										Wait(1)
										sCheck=objCalWin.JavaCheckBox("DayCheck").GetRoProperty("value")
										If sCheck="1" Then
											Fn_SchMgr_ScheduleCalendarOperations=True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified ["+aDays(iCounter)+"] is Week Day")
										Else
											Fn_SchMgr_ScheduleCalendarOperations=False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify ["+aDays(iCounter)+"] is Week Day")
											Exit  Function 
										End If
									Next
								End If
			
								'Set Non-Working Days for Schedule
								If dicSchCalendar.Item("OffWeekDays") <> "" Then
									aDays = split(dicSchCalendar.Item("OffWeekDays"), "~", -1, 1)
									
									'Check the Days Check-Boxes for the required week days
									For iCounter = 0 to Ubound(aDays)
										objCalWin.JavaCheckBox("DayCheck").SetTOProperty "attached text", aDays(iCounter)
										Wait(1)
										sCheck=objCalWin.JavaCheckBox("DayCheck").GetRoProperty("value")
										If sCheck="0" Then
											Fn_SchMgr_ScheduleCalendarOperations=True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified ["+aDays(iCounter)+"] is Off Week Day")
										Else
											Fn_SchMgr_ScheduleCalendarOperations=False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify ["+aDays(iCounter)+"] is Off Week Day")
											Exit  Function 
										End IF
									Next
								End If
			
								Set objDailyWin = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Daily Defaults Details")
								aHour=split(dicSchCalendar.Item("WorkingHrs"),"~",-1,1)
								'Verify Working Hours for 
								For iCounter = 0 to 6
			'						objCalWin.JavaEdit("DayDetails").SetTOProperty "attached text", "HH:MM"
									objCalWin.JavaEdit("DayDetails").SetTOProperty "index", iCounter
									wait(1)
									sHour=objCalWin.JavaEdit("DayDetails").GetRoProperty("value")
									If sHour=aHour(iCounter) Then
										Fn_SchMgr_ScheduleCalendarOperations=True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified working hours ["+aHour(iCounter)+"]  for ["+aDays(iCounter)+"]")
									Else
										Fn_SchMgr_ScheduleCalendarOperations=False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify working hours ["+aHour(iCounter)+"]  for ["+aDays(iCounter)+"]")
										Exit  Function 
									End If
								Next
								
					Case "Delete"

						If  objCalWin.JavaButton("Delete Calendar").GetROProperty("enabled") = 1Then
								objCalWin.JavaButton("Delete Calendar").Click micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [Delete Calendar ] Button on Calendar Dialog")
									Fn_SchMgr_ScheduleCalendarOperations = False
									objCalWin.JavaButton("Cancel").Click micLeftBtn
									Set objCalWin = Nothing
									Set objDailyWin = Nothing
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click [Delete Calendar ] Button on Calendar Dialog")
									Fn_SchMgr_ScheduleCalendarOperations = True
								End If

								Wait(3)					
								
								If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Delete Resource Calendar").Exist Then
									JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Delete Resource Calendar").JavaButton("Yes").Click micLeftBtn
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [Ok ] Button on Delete Resource Calendar Dialog")
										Fn_SchMgr_ScheduleCalendarOperations = False
										objCalWin.JavaButton("Cancel").Click micLeftBtn
										Set objCalWin = Nothing
										Set objDailyWin = Nothing
										Exit Function
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click [OK ] Button on Delete Resource Calendar Dialog")
										Fn_SchMgr_ScheduleCalendarOperations = True
									End If
								End If
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click [Delete Calendar ] Button on Calendar Dialog")
								Fn_SchMgr_ScheduleCalendarOperations = False
								objCalWin.JavaButton("Cancel").Click micLeftBtn
								Set objCalWin = Nothing
								Set objDailyWin = Nothing
								Exit Function
					   End If
				End Select
				'- - - - - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - 
				If instr(1,sSchName,":") Then
							aValues = split(sSchName,":",-1,1)
							If aValues(1) = "OK" Then
										'Click OK on Working Day  Details Dialog
										objWrkingDyExcep.JavaButton("OK").Click micLeftBtn
										If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Daily Defaults Details").Exist(5) Then
												JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Daily Defaults Details").JavaButton("OK").Click micLeftBtn
										End If
										Fn_SchMgr_ScheduleCalendarOperations = True
										Exit Function
							ElseIf aValues(1) = "Cancel" Then
										'Click Cancel on Working Day  Details Dialog
										objWrkingDyExcep.JavaButton("Cancel").Click micLeftBtn
										If  JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Daily Defaults Details").Exist(5) Then
												JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Daily Defaults Details").JavaButton("OK").Click micLeftBtn
										End If
										Fn_SchMgr_ScheduleCalendarOperations = True
										Exit Function
							End If
				End If
				'- - - - - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - -- - - - - - -- - - - - - - - 
				'Click OK on Calendar Dialog
			If sAction <> "Delete" Then
				'[TC1123-20161031-11_11_2016-VivekA-Maintenance] - Added by Poonam C ---------------------
				' For case verify click on cancel button
			    If sAction = "Verify" Then
			    	If objCalWin.JavaButton("OK").GetROProperty("enabled") = "0" Then
			    		objCalWin.JavaButton("Cancel").Click micLeftBtn
						Set objCalWin = Nothing
						Set objDailyWin = Nothing
						Exit Function
			    	End If
			    End If
			    '-----------------------------------------------------------------------------------------
				objCalWin.JavaButton("OK").Click micLeftBtn
				If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to [OK] Button on Calendar Dialog")
							Fn_SchMgr_ScheduleCalendarOperations = False
							objCalWin.JavaButton("Cancel").Click micLeftBtn
							Set objCalWin = Nothing
							Set objDailyWin = Nothing
							Exit Function
				End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set Calendar Daily Timings as [" + dicSchCalendar.Item("DayHrDetails") + "]")
						Fn_SchMgr_ScheduleCalendarOperations = True
						Set objCalWin = Nothing
						Set objDailyWin = Nothing
			End If

End Function

'*********************************************************		Function to  Creates Workflow Task	***********************************************************************

'Function Name		:					Fn_SchMgr_WrkFlwTaskCreate

'Description			 :		 		  This function is used to create the Workflow Tasks

'Parameters			   :	 			1.  sMode: Mode to Execute [Menu/RMB]
'												2. sTskName: Task(s) name to apply workflow on
'												3. sTrigger: Trigger to be applied
'												4. sTemplate; Template to be selected
											
'Return Value		   : 				 True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:				Call Fn_SchMgr_WrkFlwTaskCreate("Menu", "hhh:t2,hhh:t3", "Schedule start date", "AutoDoDo")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari S.						26-June-2010	   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_WrkFlwTaskCreate(sMode, sTskName, sTrigger, sTemplate)
   GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_WrkFlwTaskCreate"
   Dim objWrkFlwWin, bReturn,aAction,sAction,sGetTemplate,sGetTrigger
   Dim aNode, iCnt

   On Error Resume Next
	
	set objWrkFlwWin = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Workflow Rule Configuration")
	sAction = ""	
	if instr(1,sMode,":") then
		aAction = split(sMode,":",-1,1)
		sMode = aAction(0)
		sAction = aAction(1)
	end if 
	aNode = split(sTskName, ",", -1, 1)
	
'	If objWrkFlwWin.Exist(3) = False Then
If sTrigger="Predecessor complete"  Then
	sTrigger="Predecessors complete"
End If
		Select Case sMode
'
			Case "Menu"
'				'Select/Multiselect Task(s)
				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect", sTskName, "", "", "")
				If bReturn = TRUE Then
					Fn_SchMgr_WrkFlwTaskCreate = TRUE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_WrkFlwTaskCreate:Schedule Table Node(s) [" + sTskName + "] Selected Successfully")
				Else
				    Fn_SchMgr_WrkFlwTaskCreate = FALSE				 
				    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_WrkFlwTaskCreate : Fail to Select Schedule Table Node(s) [" + sTskName + "]")	
				    set objWrkFlwWin = Nothing
					Exit Function
				End If

				'Select Main Menu
				bReturn =  Fn_MenuOperation("Select", "Schedule:Workflow Task")
				If bReturn = TRUE Then
					Fn_SchMgr_WrkFlwTaskCreate = TRUE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_WrkFlwTaskCreate:Menu [Schedule:Workflow Task] Operated Successfully")
				Else
				    Fn_SchMgr_WrkFlwTaskCreate = FALSE				 
				    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_WrkFlwTaskCreate : Fail to Operate Menu [Schedule:Workflow Task]")	
				    set objWrkFlwWin = Nothing
					Exit Function
				End If

			Case "RMB"

				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelectPopup", sTskName, "", "", "Workflow Task")
				If bReturn = TRUE Then
					Fn_SchMgr_WrkFlwTaskCreate = TRUE				 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_WrkFlwTaskCreate:Successfully Selected Popup Menu [Workflow Task] for Schedule Table Node(s) [" + sTskName + "]")
				Else
				    Fn_SchMgr_WrkFlwTaskCreate = FALSE				 
				    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_WrkFlwTaskCreate : Fail to Select Popup Menu [Workflow Task] for Schedule Table Node(s) [" + sTskName + "]")	
				    set objWrkFlwWin = Nothing
					Exit Function
				End If
'
		End Select
'	End If
'	wait(60)


	If sAction = "Verify" then
		For iCnt = 0 to UBound(aNode)
				If sTrigger <> "" then
					sGetTrigger	= objWrkFlwWin.JavaEdit("WorkflowTrigger").getRoProperty("value") 
					if trim(sTrigger) = trim(sGetTrigger) then
						 Fn_SchMgr_WrkFlwTaskCreate = true	
					else
					  	Fn_SchMgr_WrkFlwTaskCreate = false
					  	exit function		
					end if 
				End If
				
				If sTemplate <> "" then
					sGetTemplate	= objWrkFlwWin.JavaEdit("WorkflowTemplate").getRoProperty("value") 
					if trim(sTemplate) = trim(sGetTemplate) then
						Fn_SchMgr_WrkFlwTaskCreate = true	
					else
						Fn_SchMgr_WrkFlwTaskCreate = false
						exit function		
					end if 
				End If
						
			Next
			'Exit Verify Case, if True
			If iCnt = UBound(aNode) Then
				Exit Function
			End If 
		
		objWrkFlwWin.JavaButton("OK").Click micLeftBtn
		If Err.Number < 0 Then
			Fn_SchMgr_WrkFlwTaskCreate = FALSE				 
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_WrkFlwTaskCreate : Fail to Click [OK] Button")	
			set objWrkFlwWin = Nothing
			Exit Function
		End If
		exit function
	End if 	
	
	'Exit if Dialog does NOT exist
	If objWrkFlwWin.Exist(20) = False Then
		Fn_SchMgr_WrkFlwTaskCreate = FALSE				 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_WrkFlwTaskCreate : Fail to Find [Workflow Rule Configuration] Dialog")	
		set objWrkFlwWin = Nothing
		Exit Function
	End If

	'If Dialog exists, Enter Trigger & Template Values
	Err.Clear
	objWrkFlwWin.JavaEdit("WorkflowTrigger").Set ""
	wait 1
	objWrkFlwWin.JavaEdit("WorkflowTrigger").Set sTrigger
	objWrkFlwWin.JavaEdit("WorkflowTrigger").Activate
	
	objWrkFlwWin.JavaEdit("WorkflowTemplate").Set ""
	wait 1
	'objWrkFlwWin.JavaEdit("WorkflowTemplate").Set sTemplate
	objWrkFlwWin.JavaEdit("WorkflowTemplate").Type sTemplate
	objWrkFlwWin.JavaEdit("WorkflowTemplate").Activate

	If Err.Number < 0 Then
		Fn_SchMgr_WrkFlwTaskCreate = FALSE				 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_WrkFlwTaskCreate : Fail to Set [Trigger OR Template] Value")	
		objWrkFlwWin.JavaButton("Cancel").Click micLeftBtn
		set objWrkFlwWin = Nothing
		Exit Function
	End If

	'Click OK Button
	objWrkFlwWin.JavaButton("OK").Click micLeftBtn
	If Err.Number < 0 Then
		Fn_SchMgr_WrkFlwTaskCreate = FALSE				 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_WrkFlwTaskCreate : Fail to Click [OK] Button")	
		set objWrkFlwWin = Nothing
		Exit Function
	End If

	Fn_SchMgr_WrkFlwTaskCreate = TRUE				 
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Fn_SchMgr_WrkFlwTaskCreate : Successfully Created Workflow Task(s) [" + sTskName + "] with Trigger [" + sTrigger + "] & Template [" + sTemplate + "]")	
	set objWrkFlwWin = Nothing

End Function
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'The function is commented because of changes in the OR  and controls of Teamcenter 9.0  a function with the same name is inserted in the VBS
'Can be used again if the changes are reverted.
'By Shreyas
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

''*********************************************************   Function is to verify ,modify summary of cost operation.***********************************************************************
'
''Function Name		:					Fn_SchMgr_CostOperation
'
''Description			 :		 		  This function is to verify ,modify summary of cost operation.
'
''Parameters			   :	 			sAction: Action need to perform.
''                                                   dicProperties: Dictionary object to hold cost value.
'											
''Return Value		   : 			  True/False  
'
''Pre-requisite			:		 		Schedule manger panel need to be open.
''                                                   Scedule need to be selected.
'
''Examples				:			Call Fn_SchMgr_CostOperation("Verify",dicSchCost)
''												
''History:
''										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''										Rupali						30-July-2010	   1.0
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Public Function Fn_SchMgr_CostOperation(sAction,dicSchCost)
'   On Error Resume Next
'   Dim dicCount , dicKeys , dicItems , iCounter,objWin ,bReturn,sIndex,objTable
'   Set objWin =  JavaWindow("ScheduleManagerWindow").JavaWindow("Costs")
'
'	If not objWin.Exist(3) Then
'		bReturn = Fn_MenuOperation("Select","Schedule:Costs")
'		If bReturn = False Then
'				Fn_SchMgr_CostOperation = False
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:Costs]")
'				Set objWin = Nothing
'				Exit Function
'		Else
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:Costs]")
'		End If
'	End If
'
'	If objWin.Exist(5) Then
'
'		 dicCount  = dicSchCost.Count
'		 dicItems = dicSchCost.Items
'		 dicKeys = dicSchCost.Keys
'
'		Select Case sAction
'			Case "Verify"
'				For iCounter = 0 to dicCount - 1
'					If  dicItems(iCounter) <> ""Then
'
'						Select Case dicKeys(iCounter)
'
'							Case "TotalEstimatedCost","TotalAccruedCost","TotalEstimatedWork","TotalAccruedWord"
'								If  dicItems(iCounter) = JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaStaticText(dicKeys(iCounter)).GetROProperty("attached text") Then
'									Fn_SchMgr_CostOperation = TRUE 
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of " + dicKeys(iCounter))
'								Else 
'									Fn_SchMgr_CostOperation = FALSE 
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of " + dicKeys(iCounter))
'									objWin.JavaButton("Cancel").Click micLeftBtn
'									Set objWin = Nothing 
'									Exit function 
'								End If 
'
'							Case "BillCode","BillSub-code","BillType","RateModifier"
'								If dicItems(iCounter) =JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaList(dicKeys(iCounter)).GetROProperty("value") Then
'									Fn_SchMgr_CostOperation = TRUE 
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of " + dicKeys(iCounter))
'								Else 
'									Fn_SchMgr_CostOperation = FALSE 
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of " + dicKeys(iCounter))
'									objWin.JavaButton("Cancel").Click micLeftBtn
'									Set objWin = Nothing 
'									Exit function
'								End If 
'
'							Case "Rollup", "DrillDown"
'								If JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaButton(dicKeys(iCounter)).GetROProperty("enabled") = "1" Then
'									JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaButton(dicKeys(iCounter)).Click micLeftBtn 
'									Fn_SchMgr_CostOperation = TRUE
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click button " + dicKeys(iCounter))
'								Else 
'									Fn_SchMgr_CostOperation = FALSE 
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button " + dicKeys(iCounter))
'									objWin.JavaButton("Cancel").Click micLeftBtn
'									Set objWin = Nothing 
'									Exit function
'								End If
'
'						End Select
'					End If
'				Next
'
'				If dicSchCost.Item("Name") <> "" Then
'					Set objTable = JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaTable("BreakdownTable")
'					sIndex = Fn_SchMgr_TableRowIndex(objTable, dicSchCost.Item("Name"),"#0")
'
'					If sIndex <> False Then
'
'						If dicSchCost.Item("EstimatedHours") <> "" Then
'                            If dicSchCost.Item("EstimatedHours") = objTable.GetCellData(sIndex,"#1") Then
'								Fn_SchMgr_CostOperation = True
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  EstimatedHours") 
'							Else 
'								Fn_SchMgr_CostOperation = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  EstimatedHours") 
'								JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaButton("Cancel").Click micLeftBtn
'								Set objTable = Nothing 
'								Set objWin = Nothing 
'							End If
'						End If
'
'						If dicSchCost.Item("AccruedHours") <> "" Then
'							If dicSchCost.Item("AccruedHours") = objTable.GetCellData(sIndex,"#2") Then
'								Fn_SchMgr_CostOperation = True
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  AccruedHours") 
'							Else 
'								Fn_SchMgr_CostOperation = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  AccruedHours") 
'								JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaButton("Cancel").Click micLeftBtn
'								Set objTable = Nothing 
'								Set objWin = Nothing 
'							End If
'						End If  
'
'						If dicSchCost.Item("EstimatedCost") <> "" Then
'							If dicSchCost.Item("EstimatedCost") = objTable.GetCellData(sIndex,"#3") Then
'								Fn_SchMgr_CostOperation = True
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  EstimatedCost") 
'							Else 
'								Fn_SchMgr_CostOperation = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  EstimatedCost") 
'								JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaButton("Cancel").Click micLeftBtn
'								Set objTable = Nothing 
'								Set objWin = Nothing 
'							End If
'						End If  
'
'						If dicSchCost.Item("AccruedCost") <> "" Then
'							If dicSchCost.Item("AccruedCost") = objTable.GetCellData(sIndex,"#4") Then
'								Fn_SchMgr_CostOperation = True
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  AccruedCost") 
'							Else 
'								Fn_SchMgr_CostOperation = False
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  AccruedCost") 
'								JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaButton("Cancel").Click micLeftBtn
'								Set objTable = Nothing 
'								Set objWin = Nothing 
'							End If
'						End If 
'					Else 
'						Fn_SchMgr_CostOperation = False
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get row from breakdown table for name " + dicSchCost.Item("Name")) 
'						objWin.JavaButton("Cancel").Click micLeftBtn
'						Set objTable = Nothing 
'						Set objWin = Nothing 
'					End If
'				End If
'
'			    JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaButton("Cancel").Click micLeftBtn
'
'				If  Err.Number < 0 Then
'					Fn_SchMgr_CostOperation = False
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get row from breakdown table for name " + dicSchCost.Item("Name")) 
'					JavaWindow("ScheduleManagerWindow").JavaWindow("Costs").JavaButton("Cancel").Click micLeftBtn
'					Set objTable = Nothing 
'					Set objWin = Nothing 
'				End If
'		End Select
'	Else
'		Fn_SchMgr_CostOperation = FALSE
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to displayed Costs dialog.")
'		Exit Function
'	End If 
'	Set objWin = Nothing 
'
'End Function 

'*********************************************************  Function do Creation of Attribute *********************************************************************

'Function Name		:					Fn_SchMgr_Load_Schedule(sButtonName, bCheckbox, sOther)

'Description			 :                  This fuction handles Load Schedule Dialog
'																	
'Parameters			   :	 				 1. sButtonName
'														 2. bCheckbox
'														3.sOther 

'Return Value		   : 			 True/False

'Examples				:			 Call  Fn_SchMgr_Load_Schedule("Yes", True, "")
'												Call  Fn_SchMgr_Load_Schedule("No", True, "")
'												Call  Fn_SchMgr_Load_Schedule("Yes", "", "")
'												Call  Fn_SchMgr_Load_Schedule("No", "", "")
												
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal Tanpure			03-March-2011      1.0												      Prasanna B.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_Load_Schedule(sButtonName, bCheckbox, sOther)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_Load_Schedule"
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Check Loasd Schedule Dialog Exist
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	If  JavaDialog("Load Schedule").Exist(5) 	Then

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'Set Check box
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			If bCheckbox = True Then
					JavaDialog("Load Schedule").JavaCheckBox("ShowMsg").Set "ON"
					If Err.Number < 0 Then
							Fn_SchMgr_Load_Schedule = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Check box" ) 
							Exit Function 
					End If
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Check box [ Load Schedule ] Dialog")
			End If

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'Click on Yes Button
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			If Lcase(sButtonName) = "yes" Then
					JavaDialog("Load Schedule").JavaButton("Yes").Click
					If Err.Number < 0 Then
							Fn_SchMgr_Load_Schedule = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on [ Yes ] button" ) 
							Exit Function 
					End If
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on [ Yes ] button of [ Load Schedule ] Dialog")
			End If

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'Click on No Button
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			If Lcase(sButtonName) = "no" Then
                    JavaDialog("Load Schedule").JavaButton("No").Click
					If Err.Number < 0 Then
							Fn_SchMgr_Load_Schedule = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click on [ No ] button" ) 
							Exit Function 
					End If
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on [ No ] button of [ Load Schedule ] Dialog")
			End If

			Fn_SchMgr_Load_Schedule = True

	Else 

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Dialog Does not Exist so It will return True
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Fn_SchMgr_Load_Schedule = True
	End If

End Function


'**********************************************  Function perform all the actions related to column management 		********************************************************

'Function Name		:					Fn_SchMgr_ProgViewColumnOperations

'Description			 :		 		  The Function perform all the actions related to column management.

'Parameters			   :	 			1.  sAction: Action need to perform. (Add/Remove/Verify)
'											  2.sColType :The Column type e.g Schedule
'											 3.aColName : The name of the column(s) to be added/removed. 
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		 Schedule table should be displayed.

'Examples				:				aColName = Array("finish_date","start_date")
											 'bReturn = Fn_SchMgr_ProgViewColumnOperations("Add","Combined",aColName)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										  Omkar 				15-March-2011   		1.0													Prasanna
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_ProgViewColumnOperations(sAction,sColType,aColName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ProgViewColumnOperations"
   	On Error Resume Next 
	Dim sIndex,iCounter,bReturn,objTable,objColChooser
	
	For iCounter = 0 to Ubound(aColName)
		If aColName(iCounter) = "Status" Then
			aColName(iCounter) = "Schedule Status" 
		End If
	Next

	Set objTable = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable")
	Set objColChooser = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Column Chooser")

	'Select Menu Column Chooser
	If Not objColChooser.Exist(5) Then
		objTable.SelectColumnHeader "Object", "RIGHT"
		JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaMenu("label:=Column Chooser","index:=0").Select
	End If
Wait(10)    'Added by Nilesh
'Select the Column Type from Column Chooser
JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Column Chooser").JavaTree("Type").Select("Types:"+sColType)
	If Err.Number < 0 Then
								Fn_SchMgr_ProgViewColumnOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Column Type :Types:ScheduleTask. from Column Chooser")
								objColChooser.JavaButton("Cancel").Click
                        		Exit Function 
		Else
							Fn_SchMgr_ProgViewColumnOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the Column Type:Types:ScheduleTask Column Chooser")

End If

	If objColChooser.Exist(5) Then
	
		Select Case sAction
	
			Case "Add"
				If IsArray(aColName) Then

					''Added by Sushma: 17-Jul-12
					wait  2
					objColChooser.JavaList("AvailableColumns").Activate

					For iCounter = 0 to Ubound(aColName)
		
						bReturn = Fn_SchMgr_TableColIndex(objTable,aColName(iCounter))
		
						If cBool(bReturn) = False Then
							'Select column from available columns.
							objColChooser.JavaList("AvailableColumns").ExtendSelect aColName(iCounter)
							Wait 1
							If Err.Number < 0 Then
								Fn_SchMgr_ColumnOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to select  " &aColName(iCounter)& " from available columns.")
								objColChooser.JavaButton("Cancel").Click
								Set objTable = Nothing
								Set objColChooser = Nothing
								Exit Function 
							End If
						Else
							Fn_SchMgr_ColumnOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Already  "  & aColName(iCounter) & "column exists in schedule table.")
							objColChooser.JavaButton("Cancel").Click
							Exit Function
						End If
		
					Next
		
					'Click Add button
					objColChooser.JavaButton("AddCol").WaitProperty "enabled",1,20000
					objColChooser.JavaButton("AddCol").Click


					If objColChooser.JavaButton("OK").Exist(5) Then
						objColChooser.JavaButton("OK").WaitProperty "enabled",1,20000
						objColChooser.JavaButton("OK").Click micLeftBtn
						'Click Apply button
					Elseif objColChooser.JavaButton("Apply").Exist(5) Then
						objColChooser.JavaButton("Apply").WaitProperty "enabled",1,20000
						objColChooser.JavaButton("Apply").Click micLeftBtn
					End If

					If Err.Number < 0 Then
						Fn_SchMgr_ColumnOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to add columns to schedule table..")
						objColChooser.JavaButton("Cancel").Click
						Set objTable = Nothing
						Set objColChooser = Nothing
						Exit Function 
					End If
					Fn_SchMgr_ColumnOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully added columns to schedule table .")
		
				End If
	
			Case "Remove"
				If IsArray(aColName) Then
					For iCounter = 0 to Ubound(aColName)
						bReturn = Fn_SchMgr_TableColIndex(objTable,aColName(iCounter))

						If  bReturn <> False Then
							'Select column from displayed columns.
							objColChooser.JavaList("DisplayedColumns").ExtendSelect aColName(iCounter)
							If Err.Number < 0 Then
								Fn_SchMgr_ProgViewColumnOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to select  " &aColName(iCounter)& " from displayed columns.")
								objColChooser.JavaButton("Cancel").Click
								Set objTable = Nothing
								Set objColChooser = Nothing
								Exit Function 
							End If
						Else 
							Fn_SchMgr_ProgViewColumnOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),aColName(iCounter) & " column does not exist in schedule table. ")
							objColChooser.JavaButton("Cancel").Click
							Set objTable = Nothing
							Set objColChooser = Nothing
							Exit Function 
						End If
					Next
					'Click Remove button
					objColChooser.JavaButton("RemoveCol").WaitProperty "enabled",1,20000
					objColChooser.JavaButton("RemoveCol").Click
					'Click Apply button
					objColChooser.JavaButton("Apply").WaitProperty "enabled",1,20000
					objColChooser.JavaButton("Apply").Click
		
					If Err.Number < 0 Then
						Fn_SchMgr_ProgViewColumnOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to remove column from schedule table.")
						Set objTable = Nothing
						Set objColChooser = Nothing
						Exit Function 
					End If
					Fn_SchMgr_ProgViewColumnOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove columns from schedule table .")
				End If

			Case "Verify"
				If IsArray(aColName) Then
					For iCounter = 0 to Ubound(aColName)
						bReturn = Fn_SchMgr_TableColIndex(objTable,aColName(iCounter))
						If bReturn <> False Then
							Fn_SchMgr_ColumnOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColName(iCounter) &" column exist in schedule table.")
						Else  
							Fn_SchMgr_ColumnOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColName(iCounter) &" column does not exist in schedule table.")
							objColChooser.JavaButton("Cancel").Click
							Set objTable = Nothing
							Set objColChooser = Nothing
							Exit Function 
						End If
					Next
					'objColChooser.JavaButton("Cancel").Click 

                    If objColChooser.JavaButton("Close").Exist(5) Then
						objColChooser.JavaButton("Close").WaitProperty "enabled",1,20000
						objColChooser.JavaButton("Close").Click micLeftBtn
						'Click Apply button
					Elseif objColChooser.JavaButton("Cancel").Exist(5) Then
						objColChooser.JavaButton("Cancel").WaitProperty "enabled",1,20000
						objColChooser.JavaButton("Cancel").Click micLeftBtn
					End If

				End If 
		End Select
	End If
	Set objTable = Nothing
	Set objColChooser = Nothing
End Function


'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   :Fn_SchMgr_FixedCostsAction(sNodeName,sMode,sAction,sCostName,sAccuralType,sEstimatedCost,sActualCost,sCurrency,bUseActualCost,sBillCode,sBillSubCode,sBillType,sNewCostName)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function is an replica of the function with the same name which is commented due to changes in controls and OR
'''''/$$$$
'''''/$$$$	
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           24/03/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			24/03/2011         1.0
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


'Public Function Fn_SchMgr_FixedCostsAction(sNodeName,sMode,sAction,sCostName,sAccuralType,sEstimatedCost,sActualCost,sCurrency,bUseActualCost,sBillCode,sBillSubCode,sBillType,sNewCostName)
'   On Error Resume Next 
'   Dim bReturn,objCost,objNewFixed,objTable,sIndex,sValue
'	Set objCost = JavaWindow("ScheduleManagerWindow").Dialog("Costs")
'	Set objNewFixed = JavaWindow("ScheduleManagerWindow").Dialog("Costs").Dialog("Fixed Cost")
'	Set objTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("FixedCostsTable")
'
'    If Not objCost.Exist(5) Then
'
'	   Select Case sMode
'			Case "Menu"
'				bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sNodeName,"","","")
'				If  bReturn <> False Then
'					Fn_SchMgr_FixedCostsAction = True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected node " + sNodeName)
'				ELse
'					Fn_SchMgr_FixedCostsAction = False
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  node " + sNodeName)
'					Set objCost = Nothing
'					Set objNewFixed = Nothing 
'					Set objTable = Nothing 
'					Call Fn_ReadyStatusSync(5)
'					Exit Function
'				End If
'
'				bReturn = Fn_MenuOperation("Select","Schedule:Costs")
'				Wait (25)
'				If bReturn = True Then
'					Fn_SchMgr_FixedCostsAction = True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Costs.]")
'				Else
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Costs.]")
'					Fn_SchMgr_FixedCostsAction = False
'					Set objCost = Nothing 
'					Set objNewFixed = Nothing 
'					Set objTable = Nothing 
'					Exit Function
'				End If
'	
'			Case "RMB"
'				bReturn =  Fn_SchMgr_SchTable_NodeOperation("PopupMenu", sNodeName, "", "", "Costs")
'				If bReturn = True Then
'					Fn_SchMgr_FixedCostsAction = True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked RMB Menu [Costs]")
'				Else
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked RMB Menu [Costs]")
'					Fn_SchMgr_FixedCostsAction = False
'					Set objCost = Nothing
'					Set objNewFixed = Nothing 
'					Set objTable = Nothing 
'					Exit Function
'				End If
'		End Select
'
'   End If
'
'    If objCost.Exist(5) Then
'		Select Case sAction
'
'			Case "Add"
'				objCost.WinButton("New").WaitProperty "enabled",1,20000
'				objCost.WinButton("New").Click micLeftBtn
'
'				If objNewFixed.Exist(5) Then 
'					If sCostName <> "" Then
'						objNewFixed.WinEdit("CostName").Set sCostName
'						wait 1
'					End If
'
'					If sAccuralType <> "" Then
'						objNewFixed.WinComboBox("AccrualType").Select sAccuralType
'						wait 1
'					End If
'	
'					If sEstimatedCost <> ""  Then
'						objNewFixed.WinEdit("EstimatedCost").Set sEstimatedCost
'						wait 1
'					End If
'	
'					If sActualCost <> "" Then
'						objNewFixed.WinEdit("ActualCost").Set sActualCost
'						wait 1
'					End If
'	
'					If  sCurrency <> "" Then
'						objNewFixed.WinComboBox("Currency").Select sCurrency
'						wait 1
'					End If 
'	
'					If bUseActualCost <> "" Then
'						If Cbool(bUseActualCost) = True Then
'							objNewFixed.WinCheckBox("UseActualCost").Set "ON"
'						ElseIf Cbool(bUseActualCost) = False Then
'							objNewFixed.WinCheckBox("UseActualCost").Set  "OFF"
'						End If
'					End If
'					wait 2
'	
'					If sBillCode <> "" Then
'						objNewFixed.WinComboBox("BillCode").Select sBillCode
'					wait 1
'					End If
'	
'					If sBillSubCode <> "" Then
'						objNewFixed.WinComboBox("BillSub-code").Select sBillSubCode
'					wait 1
'					End If
'	
'					If sBillType <> "" Then
'						objNewFixed.WinComboBox("BillType").Select sBillType
'					wait 1
'					End If
'	
'					If Err.Number < 0 Then
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add [Costs] " + sCostName)
'						Fn_SchMgr_FixedCostsAction = False
'						objNewFixed.WinButton("Cancel").Click micLeftBtn
'						objCost.WinButton("Cancel").Click micLeftBtn
'						Set objCost = Nothing
'						Set objNewFixed = Nothing 
'						Set objTable = Nothing 
'						Exit Function
'					End If
'						wait 5
'	
'					objNewFixed.WinButton("Finish").WaitProperty "enabled",1,20000
'					objNewFixed.WinButton("Finish").Click micLeftBtn
'					objCost.WinButton("Finish").WaitProperty "enabled",1,20000
'					objCost.WinButton("Finish").Click micLeftBtn
'	
'					If Err.Number < 0 Then
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add [Cost] " + sCostName)
'						Fn_SchMgr_FixedCostsAction = False
'						objNewFixed.WinButton("Cancel").Click micLeftBtn
'						objCost.WinButton("Cancel").Click micLeftBtn
'						Set objCost = Nothing
'						Set objNewFixed = Nothing 
'						Set objTable = Nothing 
'						Exit Function
'					Else 
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully add [Cost] " + sCostName)
'						Fn_SchMgr_FixedCostsAction = True
'						wait 2
'					End If
'				Else 
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fixed Cost  dialog does not exist.")
'					Fn_SchMgr_FixedCostsAction = False
'					Set objCost = Nothing
'					Set objNewFixed = Nothing 
'					Set objTable = Nothing 
'					Exit Function
'			    End If 
'
'			Case "Modify"
'				
'				sIndex = Fn_SchMgr_WinListViewIndex(objTable,sCostName)
'				If sIndex <> False Then
'					objTable.Select(sIndex)
''					objTable.ClickCell sIndex,"Cost Name" 
'					objCost.WinButton("Details").WaitProperty "enabled",1,20000
'					objCost.WinButton("Details").Click micLeftBtn
'				
'					If objNewFixed.Exist(5) Then 
'						If sNewCostName <> "" Then
'							objNewFixed.WinEdit("CostName").Set sNewCostName
'						End If
'	
'						If sAccuralType <> "" Then
'							objNewFixed.WinComboBox("AccrualType").Select sAccuralType
'						End If
'		
'						If sEstimatedCost <> ""  Then
'							objNewFixed.WinEdit("EstimatedCost").Set sEstimatedCost
'						End If
'		
'						If sActualCost <> "" Then
'							objNewFixed.WinEdit("ActualCost").Set sActualCost
'						End If
'		
'						If  sCurrency <> "" Then
'							objNewFixed.WinComboBox("Currency").Select sCurrency
'						End If 
'		
'						If bUseActualCost <> "" Then
'							If Cbool(bUseActualCost) = True Then
'								objNewFixed.WinCheckBox("UseActualCost").Set "ON"
'							ElseIf Cbool(bUseActualCost) = False Then
'								objNewFixed.WinCheckBox("UseActualCost").Set  "OFF"
'							End If
'						End If
'		
'						If sBillCode <> "" Then
'							objNewFixed.WinComboBox("BillCode").Select sBillCode
'						End If
'		
'						If sBillSubCode <> "" Then
'							objNewFixed.WinComboBox("BillSub-code").Select sBillSubCode
'						End If
'		
'						If sBillType <> "" Then
'							objNewFixed.WinComboBox("BillType").Select sBillType
'						End If
'		
'						If Err.Number < 0 Then
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modified [Costs] " + sCostName)
'							Fn_SchMgr_FixedCostsAction = False
'							objNewFixed.WinButton("Cancel").Click micLeftBtn
'							objCost.WinButton("Cancel").Click micLeftBtn
'							Set objCost = Nothing
'							Set objNewFixed = Nothing 
'							Set objTable = Nothing 
'							Exit Function
'						End If
'		
'						objNewFixed.WinButton("Finish").WaitProperty "enabled",1,20000
'						objNewFixed.WinButton("Finish").Click micLeftBtn
'						objCost.WinButton("Finish").WaitProperty "enabled",1,20000
'						objCost.WinButton("Finish").Click micLeftBtn
'		
'						If Err.Number < 0 Then
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to modified [Cost] " + sCostName)
'							Fn_SchMgr_FixedCostsAction = False
'							objNewFixed.WinButton("Cancel").Click micLeftBtn
'							objCost.WinButton("Cancel").Click micLeftBtn
'							Set objCost = Nothing
'							Set objNewFixed = Nothing 
'							Set objTable = Nothing 
'							Exit Function
'						Else 
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully modified [Cost] " + sCostName)
'							Fn_SchMgr_FixedCostsAction = True
'						End If
'					Else 
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fixed Cost  dialog does not exist.")
'						Fn_SchMgr_FixedCostsAction = False
'						Set objCost = Nothing
'						Set objNewFixed = Nothing 
'						Set objTable = Nothing 
'						Exit Function
'					End If 
'				Else 
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sCostName +  " cost name does not exist in fixed cost table.")
'					Fn_SchMgr_FixedCostsAction = False
'					Set objCost = Nothing
'					Set objNewFixed = Nothing 
'					Set objTable = Nothing 
'					Exit Function
'				End If
'
'			Case "Verify"
'					sIndex = Fn_SchMgr_WinListViewIndex(objTable,sCostName)
'				If sIndex <> False Then
'					objTable.Select(sIndex)
'					objCost.WinButton("Details").WaitProperty "enabled",1,20000
'					objCost.WinButton("Details").Click micLeftBtn
'				
'					If objNewFixed.Exist(5) Then 
'						If sCostName <> "" Then
'							If  sCostName = objNewFixed.WinEdit("CostName").GetROProperty("text") Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Cost Name] value  " + sCostName)
'								Fn_SchMgr_FixedCostsAction = True
'							Else 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Cost Name] value  " + sCostName)
'								Fn_SchMgr_FixedCostsAction = False
'								objNewFixed.WinButton("Cancel").Click micLeftBtn
'								objCost.WinButton("Cancel").Click micLeftBtn
'								Set objCost = Nothing
'								Set objNewFixed = Nothing 
'								Set objTable = Nothing 
'								Exit Function
'							End If
'						End If
'	
'						If sAccuralType <> "" Then
'							If sAccuralType = objNewFixed.WinComboBox("AccrualType").GetROProperty("text") Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Accural Type] value  " + sAccuralType)
'								Fn_SchMgr_FixedCostsAction = True
'							Else 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Accural Type] value  " + sAccuralType)
'								Fn_SchMgr_FixedCostsAction = False
'								objNewFixed.WinButton("Cancel").Click micLeftBtn
'								objCost.WinButton("Cancel").Click micLeftBtn
'								Set objCost = Nothing
'								Set objNewFixed = Nothing 
'								Set objTable = Nothing 
'								Exit Function
'							End If
'						End If
'		
'						If sEstimatedCost <> ""  Then
'							If  sEstimatedCost = objNewFixed.WinEdit("EstimatedCost").GetROProperty("text") Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Estimated Cost] value  " + sEstimatedCost)
'								Fn_SchMgr_FixedCostsAction = True
'							Else 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Estimated Cost] value  " + sEstimatedCost)
'								Fn_SchMgr_FixedCostsAction = False
'								objNewFixed.WinButton("Cancel").Click micLeftBtn
'								objCost.WinButton("Cancel").Click micLeftBtn
'								Set objCost = Nothing
'								Set objNewFixed = Nothing 
'								Set objTable = Nothing 
'								Exit Function
'							End If
'						End If
'		
'						If sActualCost <> "" Then
'							If  sActualCost = objNewFixed.WinEdit("ActualCost").GetROProperty("text") Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Actual Cost ] value  " + sActualCost)
'								Fn_SchMgr_FixedCostsAction = True
'							Else 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Actual Cost] value  " + sActualCost)
'								Fn_SchMgr_FixedCostsAction = False
'								objNewFixed.WinButton("Cancel").Click micLeftBtn
'								objCost.WinButton("Cancel").Click micLeftBtn
'								Set objCost = Nothing
'								Set objNewFixed = Nothing 
'								Set objTable = Nothing 
'								Exit Function
'							End If
'						End If
'		
'						If  sCurrency <> "" Then
'							If sCurrency = objNewFixed.WinComboBox("Currency").GetROProperty("text") Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Currency] value  " + sCurrency)
'								Fn_SchMgr_FixedCostsAction = True
'							Else 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Currency] value  " + sCurrency)
'								Fn_SchMgr_FixedCostsAction = False
'								objNewFixed.WinButton("Cancel").Click micLeftBtn
'								objCost.WinButton("Cancel").Click micLeftBtn
'								Set objCost = Nothing
'								Set objNewFixed = Nothing 
'								Set objTable = Nothing 
'								Exit Function
'							End If
'						End If 
'		
'						If bUseActualCost <> "" Then
'							sValue= JavaWindow("ScheduleManagerWindow").Dialog("Costs").Dialog("Fixed Cost").WinCheckBox("UseActualCost").CheckProperty("Checked","ON")
'							If Cbool(bUseActualCost) = True Then
'									If lcase(bUseActualCost)=lcase(cBool(sValue)) Then
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Use Actual Cost] value  " + bUseActualCost)
'											Fn_SchMgr_FixedCostsAction = True
'									Else
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Use Actual Cost] value  " + bUseActualCost)
'											Fn_SchMgr_FixedCostsAction = False
'											objNewFixed.WinButton("Cancel").Click micLeftBtn
'											objCost.WinButton("Cancel").Click micLeftBtn
'											Set objCost = Nothing
'											Set objNewFixed = Nothing 
'											Set objTable = Nothing 
'											Exit Function
'									End If
'							Elseif Cbool(bUseActualCost) = False Then
'									If lcase(bUseActualCost)=lcase(sValue) Then
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Use Actual Cost] value  " + bUseActualCost)
'											Fn_SchMgr_FixedCostsAction = True
'									Else
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Use Actual Cost] value  " + bUseActualCost)
'											Fn_SchMgr_FixedCostsAction = False
'											objNewFixed.WinButton("Cancel").Click micLeftBtn
'											objCost.WinButton("Cancel").Click micLeftBtn
'											Set objCost = Nothing
'											Set objNewFixed = Nothing 
'											Set objTable = Nothing 
'											Exit Function
'									End If
'							End if
'						End if
'	
'	
'						If sBillCode <> "" Then
'							If sBillCode = objNewFixed.WinComboBox("BillCode").GetROProperty("text") Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Bill Code] value  " + sBillCode)
'								Fn_SchMgr_FixedCostsAction = True
'							Else 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Bill Code] value  " + sBillCode)
'								Fn_SchMgr_FixedCostsAction = False
'								objNewFixed.WinButton("Cancel").Click micLeftBtn
'								objCost.WinButton("Cancel").Click micLeftBtn
'								Set objCost = Nothing
'								Set objNewFixed = Nothing 
'								Set objTable = Nothing 
'								Exit Function
'							End If
'						End If
'		
'						If sBillSubCode <> "" Then
'							If sBillSubCode = objNewFixed.WinComboBox("BillSub-code").GetROProperty("text") Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Bill SubCode] value  " + sBillSubCode)
'								Fn_SchMgr_FixedCostsAction = True
'							Else 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Bill SubCode] value  " + sBillSubCode)
'								Fn_SchMgr_FixedCostsAction = False
'								objNewFixed.WinButton("Cancel").Click micLeftBtn
'								objCost.WinButton("Cancel").Click micLeftBtn
'								Set objCost = Nothing
'								Set objNewFixed = Nothing 
'								Set objTable = Nothing 
'								Exit Function
'							End If
'						End If
'		
'						If sBillType <> "" Then
'							If sBillType = objNewFixed.WinComboBox("BillType").GetROProperty("text") Then
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify [Bill Type] value  " + sBillType)
'								Fn_SchMgr_FixedCostsAction = True
'							Else 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify [Bill Type] value  " + sBillType)
'								Fn_SchMgr_FixedCostsAction = False
'								objNewFixed.WinButton("Cancel").Click micLeftBtn
'								objCost.WinButton("Cancel").Click micLeftBtn
'								Set objCost = Nothing
'								Set objNewFixed = Nothing 
'								Set objTable = Nothing 
'								Exit Function
'							End If
'						End If
'		
'					    objNewFixed.WinButton("Cancel").WaitProperty "enabled",1,20000
'						objNewFixed.WinButton("Cancel").Click micLeftBtn
'						objCost.WinButton("Cancel").WaitProperty "enabled",1,20000
'						objCost.WinButton("Cancel").Click micLeftBtn
'		
'						If Err.Number < 0 Then
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify")
'							Fn_SchMgr_FixedCostsAction = False
'							objNewFixed.WinButton("Cancel").Click micLeftBtn
'							objCost.WinButton("Cancel").Click micLeftBtn
'							Set objCost = Nothing
'							Set objNewFixed = Nothing 
'							Set objTable = Nothing 
'							Exit Function
'						End If
'					Else 
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fixed Cost  dialog does not exist.")
'						Fn_SchMgr_FixedCostsAction = False
'						Set objCost = Nothing
'						Set objNewFixed = Nothing 
'						Set objTable = Nothing 
'						Exit Function
'					End If 
'				Else 
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sCostName +  " cost name does not exist in fixed coast table.")
'					Fn_SchMgr_FixedCostsAction = False
'					Set objCost = Nothing
'					Set objNewFixed = Nothing 
'					Set objTable = Nothing 
'					Exit Function
'				End If
'		End Select
'  Else 
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Cost dialog does not exist.")
'		Fn_SchMgr_FixedCostsAction = False
'		Set objCost = Nothing
'		Set objNewFixed = Nothing 
'		Set objTable = Nothing 
'		Exit Function
'	End If
'
'	Set objCost = Nothing
'	Set objNewFixed = Nothing 
'	Set objTable = Nothing 
'End Function
'

'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_SchMgr_WinListViewIndex(objSchTable, sListItem)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will create mapping and verify if the mapping has happened
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) objSchTable : Path of the WinListView   control
'''''/$$$$                                      2.) sListItem : Item Name to be searched
''''''/$$$$
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           24/03/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			 24/03/2011         1.0
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_SchMgr_WinListViewIndex(objSchTable, sListItem)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_WinListViewIndex"
	Dim sCount ,sItem, IntCounter, ObjTable, StrIndex, ArrNode

	On Error Resume Next

	'Verify that Resource Table is displayed
	If objSchTable.Exist(5) Then

		'Get the itemcount in the list
		sCount= JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("FixedCostsTable").GetItemsCount
		
		'Get the required item name and assign that item name to the function
				For icount =0 to sCount-1
							sItem=JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("FixedCostsTable").GetItem(icount)
							If lcase(sListItem)=lcase(sItem) Then
									Fn_SchMgr_WinListViewIndex=sItem
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully assigned the value [" +sListItem+"] to the function 'Fn_SchMgr_WinListViewIndex'")	
									Exit for
							End If
				Next

		If  cstr(icount) = sCount Then
			Fn_SchMgr_WinListViewIndex = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ":Failed to Get  the required item name")	
		End If
  End If
End Function


'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_SchMgr_ColumnChooserOperations(sAction,sType,sColumn,sButton)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will create mapping and verify if the mapping has happened
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''''/$$$$                                     2.) sType : Type of column to be selected
'''''/$$$$							          3.) sColumn : Valid Column name
'''''/$$$$									 4.) sButton : Button to be clicked
'''''/$$$$
''''''/$$$$
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           	30/03/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			 30/03/2011         1.0
'''''/$$$$
'''''/$$$$	How To Use :  bReturn=Fn_SchMgr_ColumnChooserOperations("Remove","","Description","Apply")
'''''/$$$$						   bReturn=Fn_SchMgr_ColumnChooserOperations("AvailableListVerify","","Description","Cancel")
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SchMgr_ColumnChooserOperations(sAction,sType,sColumn,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ColumnChooserOperations"
   Dim  iCounter, aProperties,objColumn,sDetails
   Dim iItemCount, bFlag, iCounter1, sAppValue
   Set objColumn = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Column Chooser")
   Fn_SchMgr_ColumnChooserOperations=false

If objColumn.Exist(3)=false Then
'Invoke the column chooser dialog
JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectColumnHeader "Object","RIGHT"
JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaMenu("label:=Column Chooser","index:=0").Select
End if

		If objColumn.Exist(3)=false Then

				Fn_SchMgr_ColumnChooserOperations=false
				exit function
		Else

		Select Case sAction

			Case "Add"

'
			If sType<>"" and sColumn <>"" Then
				objColumn.JavaTree("Type").select sType

				'for selecting more than one node
				wait(3)
					If instr(1,sColumn,",")>0 Then
						aProperties=split(sColumn,",",-1,1)
							For iCounter=0 To Ubound(aProperties)
								objColumn.JavaList("AvailableColumns").ExtendSelect aProperties(iCounter)
								If err.number<0 Then
									Fn_SchMgr_ColumnChooserOperations=false
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Add The values "+aProperties(iCounter))
									exit function
								End If

								Fn_SchMgr_ColumnChooserOperations=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Added The values "+aProperties(iCounter))
							
						
						Next
				Else 

				'for selecting just one node
				objColumn.JavaEdit("AvailColsTxt").Type sColumn
				If err.number<0 Then
									Fn_SchMgr_ColumnChooserOperations=false
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Add The values "+sColumn)
									exit function
				End If
				Fn_SchMgr_ColumnChooserOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Added The value "+sColumn)
			End if

		'click on add button
		 If objColumn.JavaButton("AddCol").GetROProperty("enabled")="1" Then
							objColumn.JavaButton("AddCol").Click micLeftBtn
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on the AddCol Button.")
								If err.number<0 Then
							Fn_SchMgr_ColumnChooserOperations=false
							exit function
						End If
						
					End If
		End if

	Case "Remove"

		If  sColumn <>"" Then
			'for selecting more than one node
				If instr(1,sColumn,",")>0 Then
						aProperties=split(sColumn,",",-1,1)
						For iCounter=0 To Ubound(aProperties)
				objColumn.JavaList("DisplayedColumns").ExtendSelect aProperties(iCounter)
				If err.number<0 Then
					Fn_SchMgr_ColumnChooserOperations=false
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Remove The values "+aProperties(iCounter))
					exit function
				End If
				Fn_SchMgr_ColumnChooserOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Removed The values "+aProperties(iCounter))
			Next
				Else 
					'for selecting just one node
					objColumn.JavaEdit("DisplayedColsTxt").Type sColumn
					Fn_SchMgr_ColumnChooserOperations=True
	End if
End if

			'click on remove button
			 If objColumn.JavaButton("RemoveCol").GetROProperty("enabled")="1" Then
								objColumn.JavaButton("RemoveCol").Click micLeftBtn
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on the RemoveCol Button.")
									If err.number<0 Then
								Fn_SchMgr_ColumnChooserOperations=false
								exit function
							End If
							
						End If


	Case "Verify" 

		If sColumn <>"" Then
				
				'[TC1122-2016011300-29_01_2016-VivekA-Maintenance] - Modified code to verify items in Displayed Columns list
				'for verifying more than one node
'				If instr(1,sColumn,",")>0 Then
				aProperties=split(sColumn,",",-1,1)
				iItemCount = CInt(objColumn.JavaList("DisplayedColumns").GetROProperty("items count"))
				For iCounter=0 To Ubound(aProperties)
					bFlag = False				
					For iCounter1 = 0 To iItemCount-1			
						sAppValue = objColumn.JavaList("DisplayedColumns").GetItem(iCounter1)
						If sAppValue = aProperties(iCounter) Then
							bFlag = True
							Exit For
						End If
					Next
				
					If bFlag = False Then
						Fn_SchMgr_ColumnChooserOperations=false
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to verify The values "+aProperties(iCounter))
						exit function
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified The values "+aProperties(iCounter))
					Fn_SchMgr_ColumnChooserOperations=True
				Next
				'------------------------------------------------------------------------------------------------------------
'				 If instr(1,sColumn,",")>0 Then
''				sColumn=Replace(sColumn,",","")
'				sColumn=Replace(sColumn,",",chr(10))
'				End If
'				sDetails=objColumn.JavaList("DisplayedColumns").GetROProperty ("value")
'				If instr(1,sDetails,sColumn)>0 Then
'				Fn_SchMgr_ColumnChooserOperations=True
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified that the values are present")
'				Else
'						Fn_SchMgr_ColumnChooserOperations=false
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify that the values are present")
'						exit function
'				End If

'				Else 
'
'		objColumn.JavaEdit("DisplayedColsTxt").Type sColumn
'		sDetails=objColumn.JavaList("DisplayedColumns").GetROProperty ("value")
'			End if
'			If instr(1,sDetails,sColumn)>0 Then
'					Fn_SchMgr_ColumnChooserOperations=True
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified that the values are present")
'			Else
'					Fn_SchMgr_ColumnChooserOperations=false
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify that the values are present")
'					exit function
'			End If
	End if


	Case "AvailableListVerify" 
	
	If sType<>"" Then
		objColumn.JavaTree("Type").select sType
	End If

		If  sColumn <>"" Then

				'for verifying more than one node
				If instr(1,sColumn,",")>0 Then
						aProperties=split(sColumn,",",-1,1)
						For iCounter=0 To Ubound(aProperties)
				objColumn.JavaList("AvailableColumns").ExtendSelect aProperties(iCounter)
					objColumn.JavaList("AvailableColumns")
				If err.number<0 Then
					Fn_SchMgr_ColumnChooserOperations=false
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to select The values "+aProperties(iCounter))
					exit function
				End If
				Fn_SchMgr_ColumnChooserOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected The values "+aProperties(iCounter))
				Fn_SchMgr_ColumnChooserOperations=True
			
			Next
				sDetails=objColumn.JavaList("AvailableColumns").GetROProperty ("value")
				If instr(1,sDetails,sColumn)>0 Then
				Fn_SchMgr_ColumnChooserOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified that the values are present")
				Else
						Fn_SchMgr_ColumnChooserOperations=false
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify that the values are present")
						exit function
				End If

				Else 

		objColumn.JavaList("AvailableColumns").Type sColumn
		sDetails=objColumn.JavaList("AvailableColumns").GetROProperty ("value")
			End if
			If instr(1,sDetails,sColumn)>0 Then
					Fn_SchMgr_ColumnChooserOperations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified that the values are present")
			Else
					Fn_SchMgr_ColumnChooserOperations=false
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify that the values are present")
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Column Chooser").close
					exit function
			End If
	End if

	End Select
End if


'Click on required button

If sButtons<>"" Then

JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Column Chooser").JavaButton(sButtons).Click micLeftBtn
	If err.number<0 Then
		Fn_SchMgr_ColumnChooserOperations=false
		exit function
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on the button "+sButtons)
Else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: No Button Need to be clicked")
End If
End Function


'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_SchMgr_ChooseSchedulesOperations(sAction,sTreeNode,sSearchText,sButton)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will create mapping and verify if the mapping has happened
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''''/$$$$                                     2.) sType : Type of column to be selected
'''''/$$$$							          3.) sColumn : Valid Column name
'''''/$$$$									 4.) sButton : Button to be clicked
'''''/$$$$
''''''/$$$$
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           	30/03/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			 30/03/2011         1.0
'''''/$$$$
'''''/$$$$	How To Use :  bReturn=Fn_SchMgr_ChooseSchedulesOperations("Remove","MasterSch_20208,MasterSch_21669","","OK")
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public function Fn_SchMgr_ChooseSchedulesOperations(sAction,sTreeNode,sSearchText,sButtons)
 GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ChooseSchedulesOperations"
 Dim  iCounter, aProperties,objChooseSchedules,sDetails,objError,intCount,sItem,intNodeCount
 Dim bFlag,aNodeValues,aValues,sTotalScheduleCounter,sDivision,aDiv,jCounter,aProp,aProp2,sCount,i

Set objChooseSchedules=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Choose Schedules...")
Set objError=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("JavaErrorDialog")
Fn_SchMgr_ChooseSchedulesOperations=false

If objChooseSchedules.Exist(3)=false Then
	'Invoke the column chooser dialog
	Call Fn_ToolbatButtonClick("Choose Schedules")
End if

call Fn_ReadyStatusSync(5)

If objChooseSchedules.Exist(3)=false Then
'to verify if the dialog is opened..
Fn_SchMgr_ChooseSchedulesOperations=false
exit function
End If
 
If objChooseSchedules.JavaButton("LoadAll").GetROProperty("enabled") = 1 Then
	objChooseSchedules.JavaButton("LoadAll").Click micLeftBtn
	call Fn_ReadyStatusSync(3)
End If
sCount= objChooseSchedules.JavaTree("AvailableSchedules").GetROProperty("items count")
  	
Select Case sAction

	Case "DoubleClickAdd"

		bFlag=False

		For i=0 to sCount-1
			sValue=objChooseSchedules.JavaTree("AvailableSchedules").GetItem(i)
			
			If instr(1,sValue,sTreeNode)>0 Then
				objChooseSchedules.JavaTree("AvailableSchedules").Activate("#0:"+sTreeNode)
				bFlag=True
				Exit for
			End if
		Next
		If bFlag=True Then
			Fn_SchMgr_ChooseSchedulesOperations=true
		End If


	Case "Add"

		bFlag=False

		If instr(1,sTreeNode,",")>0 Then
			aProperties=split(sTreeNode,",",-1,1)
				For iCounter=0 To Ubound(aProperties)
					objChooseSchedules.JavaTree("AvailableSchedules").ExtendSelect "#0:"+aProperties(iCounter)
					If err.number<0 Then
						Fn_SchMgr_ChooseSchedulesOperations=false
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Find The values "+aProperties(iCounter))
						exit function
					End If
					
					If iCounter=ubound(aProperties) Then
						objChooseSchedules.JavaButton("Add").Click micLeftBtn
						bFlag=True
					End If
					If err.number<0 Then
						Fn_SchMgr_ChooseSchedulesOperations=false
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Add The values "+aProperties(iCounter))
						exit function
					End If
				Next
	
				Fn_SchMgr_ChooseSchedulesOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Added The values " + sTreeNode)
		Else
			For i=0 to sCount-1
				sValue=objChooseSchedules.JavaTree("AvailableSchedules").GetItem(i)
				If instr(1,sValue,sTreeNode)>0 Then
					objChooseSchedules.JavaTree("AvailableSchedules").Select "#0:"+sTreeNode
					objChooseSchedules.JavaButton("Add").Click micLeftBtn
					bFlag=True
					Exit for
				End if
			Next
		End if
		
			
		If bFlag=True Then
			Fn_SchMgr_ChooseSchedulesOperations=true
		End If
			

	Case "Remove"

		If instr(1,sTreeNode,",")>0 Then
			aProperties=split(sTreeNode,",",-1,1)
			For iCounter=0 To Ubound(aProperties)
				objChooseSchedules.JavaList("SelectedSchedules").ExtendSelect aProperties(iCounter)
				If err.number<0 Then
					Fn_SchMgr_ChooseSchedulesOperations=false
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Remove The values "+aProperties(iCounter))
					exit function
				End If
				Fn_SchMgr_ChooseSchedulesOperations=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Removed The values "+aProperties(iCounter))
			Next
		Else 
			'for selecting just one node
			objChooseSchedules.JavaList("SelectedSchedules").Select sTreeNode
			Fn_SchMgr_ChooseSchedulesOperations=True
		End if


		'click on remove button
		If objChooseSchedules.JavaButton("Remove").GetROProperty("enabled")="1" Then
			objChooseSchedules.JavaButton("Remove").Click micLeftBtn
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on the Remove Button.")
			If err.number<0 Then
				Fn_SchMgr_ChooseSchedulesOperations=false
				exit function
			End If
		End IF

	Case "Search&Add"

		'click on clear button 
		objChooseSchedules.JavaButton("Clear").Click micLeftBtn
		If err.number<0 Then
			Fn_SchMgr_ChooseSchedulesOperations=false
			exit function
		End If

		'Set the schedule name in the searcg edit box
		objChooseSchedules.JavaEdit("SearchText").Type sSearchText
		If err.number<0 Then
			Fn_SchMgr_ChooseSchedulesOperations=false
			exit function
		End If

		'click on the search button
		objChooseSchedules.JavaButton("Find").Click micLeftBtn
		If err.number<0 Then
			Fn_SchMgr_ChooseSchedulesOperations=false
			exit function
		End If

		'check if any error has occured	
		objError.SetTOProperty "title","Object Not Found"
		If objError.Exist(3) Then
			objError.JavaButton("OK").Click micLeftBtn
			Fn_SchMgr_ChooseSchedulesOperations=false
			Exit Function
		End If
			
		
		sValue=objChooseSchedules.JavaTree("AvailableSchedules").GetROProperty("text")
		If instr(1,sValue,sSearchText)>0 Then
			bFlag=True
			objChooseSchedules.JavaButton("Add").Click micLeftBtn
			If err.number<0 Then
				Fn_SchMgr_ChooseSchedulesOperations=false
				exit function
			End If
		Else
			Fn_SchMgr_ChooseSchedulesOperations=false
			Exit Function
		End If
					
'		For i=0 to sCount-1
'			sValue=objChooseSchedules.JavaTree("AvailableSchedules").GetItem(i)
'			If instr(1,sValue,sSearchText)>0 Then
'				objChooseSchedules.JavaTree("AvailableSchedules").Select "#0:"+sSearchText
'				objChooseSchedules.JavaButton("Add").Click micLeftBtn
'				bFlag=True
'				Exit for
'			End if
'		Next
		If bFlag=True Then
			Fn_SchMgr_ChooseSchedulesOperations=true
		End If

		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully found the node "+sTreeNode)
		Fn_SchMgr_ChooseSchedulesOperations=true

	Case "Verify"

		'for verifying multiple nodes
		If instr(1,sTreeNode,",")>0 Then
			aProperties=split(sTreeNode,",",-1,1)
			For iCounter=0 To Ubound(aProperties)
				objChooseSchedules.JavaList("SelectedSchedules").ExtendSelect aProperties(iCounter)
				If err.number<0 Then
					Fn_SchMgr_ChooseSchedulesOperations=false
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to verify that the value "+aProperties(iCounter)+" exists in the list")
					exit function
				End If
					Fn_SchMgr_ChooseSchedulesOperations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully verified that the value "+aProperties(iCounter)+" exists in the list")
			Next
		Else 

			intNodeCount= 	objChooseSchedules.JavaList("SelectedSchedules").GetROProperty("items count")
			For intCount = 0 to intNodeCount - 1
				sItem = objChooseSchedules.JavaList("SelectedSchedules").GetItem(intCount)
				If instr(1,sItem,sTreeNode)>0 Then
					Fn_SchMgr_ChooseSchedulesOperations = TRUE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected The value "+sItem+" is present")
					Exit For
				End If
			Next
			If cint(intCount) = cint(intNodeCount) Then
				Fn_SchMgr_ChooseSchedulesOperations = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to verify that the value "+sItem+" is present")
			End if
		End If

	Case "Set&Clear"
		'Set the schedule name in the searcg edit box
		objChooseSchedules.JavaEdit("SearchText").Type sSearchText
		If Err.number<0 Then
			Fn_SchMgr_ChooseSchedulesOperations=False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed To Set Schedule Name In Search Edit Box.")
			Exit Function
		Else
			Fn_SchMgr_ChooseSchedulesOperations=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Schedule Name In Search Edit Box.")
		End If
		
		'click on clear button 
		objChooseSchedules.JavaButton("Clear").Click micLeftBtn
		If Err.number<0 Then
			Fn_SchMgr_ChooseSchedulesOperations=False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Failed To Click On Clear Button .")
			Exit Function
		Else
			Fn_SchMgr_ChooseSchedulesOperations=True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked On Clear Button .")
		End If

End Select


If sButtons<>"" Then

objChooseSchedules.JavaButton(sButtons).Click micLeftBtn
	If err.number<0 Then
		Fn_SchMgr_ChooseSchedulesOperations=false
		exit function
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on the button "+sButtons)
Else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: No Button Need to be clicked")
End If

End function





'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_SchMgr_SaveProgram(sButton)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will handle the "Save" dialog appearing when trying to close schedule manager while a program view is open
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sButton : Valid Button Name
'''''/$$$$  
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           	05/04/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			 05/04/2011          1.0
'''''/$$$$
'''''/$$$$	How To Use :  bReturn = Fn_SchMgr_SaveProgram("Cancel")
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SchMgr_SaveProgram(sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_SaveProgram"
   Dim objError
Fn_SchMgr_SaveProgram=false
   'Close the schedule manager perspective
   Call Fn_MenuOperation("Select","File:Close")
If err.number<0 Then
	Fn_SchMgr_SaveProgram=false
	Exit function
Else
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully closed the Schedule Manager perspective")
	   Fn_SchMgr_SaveProgram=True
End If

    JavaWindow("ScheduleManagerWindow").JavaWindow("Error").SetTOProperty "title","Save"

	Set objError=JavaWindow("ScheduleManagerWindow").JavaWindow("Error")
	


		If sButton<>"" Then
			If objError.Exist(3) Then
				Select Case sButton
					Case "Yes"
						objError.JavaButton("OK").SetTOProperty "label",sButton
						objError.JavaButton("OK").Click micLeftBtn
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the button "+sButton)
					Case "No"
						objError.JavaButton("OK").SetTOProperty "label",sButton
						objError.JavaButton("OK").Click micLeftBtn
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the button "+sButton)
					Case "Cancel"
						objError.JavaButton("OK").SetTOProperty "label",sButton
						objError.JavaButton("OK").Click micLeftBtn
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the button "+sButton)
					End Select
			End If
		End If

	Set objError=nothing
End Function


'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   :  Fn_SchMgr_GanttChartOperations(sAction, aParam)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will handle the GacttChart Operations
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be Carried out
'''''/$$$$    								      :   2) aParam : Array of Paramertes
'''''/$$$$  
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   Vallari           	06/04/2011         1.0
'''''/$$$$
'''''/$$$$
'''''/$$$$	How To Use :  bReturn =  Fn_SchMgr_GanttChartOperations("DeleteDep_KeyOp", aParam)
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public Function Fn_SchMgr_GanttChartOperations(sAction, aParam)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_GanttChartOperations"
   On Error Resume Next

   Dim bReturn, iDepObj, objDep, iCnt, objBounds
   Dim sFromTask, sToTask
   Dim iX, iY, WshShell
   Dim iStart, iFinish,sTemp,sButton
	If instr(1,sAction,":") Then
		sTemp = split(sAction,":",-1,1)
		sAction = sTemp(0)
		sButton = sTemp(1)
	End If

   Select Case sAction
		 	Case "DeleteDep_KeyOp"

					'Removing focus from GanttChart, if in case any task is selected
					JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectRow "#0"
					wait(1)
'					Set WshShell = CreateObject("WScript.Shell")
'					WshShell.SendKeys "{F5}"
'					wait(2)
'					Set WshShell = nothing


	                Set objDep = Description.Create()
					'objDep("Class Name").Value = "JavaObject"
					objDep("Class Name").Value = "JavaStaticText"
					objDep("toolkit class").Value = "com.teamcenter.rac.schedule.common.gantt.GanttNewDependency"
					iDepObj = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaObject("GanttChart").ChildObjects(objDep).count
					Set objDep = nothing

					For iCnt = 0 to iDepObj-1
	                        JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaStaticText("GanttDependency").SetTOProperty "index", cstr(iCnt)
							sFromTask = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaStaticText("GanttDependency").Object.getFromTask.toString()
							sToTask = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaStaticText("GanttDependency").Object.getToTask.toString()
							If trim(lcase(sFromTask)) = trim(lcase(aParam(0))) And trim(lcase(sToTask)) = trim(lcase(aParam(1))) Then
								Set objBounds = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaStaticText("GanttDependency").Object.getBounds
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaObject("GanttChart").Object.scrollRectToVisible(objBounds)
								iStart = cint(objBounds.getX) + 2
								iFinish =cint( objBounds.getX) + cint(objBounds.getWidth)
								iY = cint(objBounds.getY) + (cint(objBounds.getHeight)/2) - 3
								Set objBounds = nothing
								Exit For
							End If
					Next
					If iCnt = iDepObj Then
							Fn_SchMgr_GanttChartOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find Dependency in GanttChart between Task [" + aParam(0) + "] and Task [" + aParam(1) + "]")
							Exit Function
					End If

					'Set the title for Dependency Delete dialog
					JavaDialog("Confirmation").SetTOProperty "title", "Confirm Dependency Deletion"

					For iX = iStart to iFinish step 5
								'Select the Dependency and hit Delete key
								JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaObject("GanttChart").Click iX, iY,"LEFT"
								wait(1)
								Set WshShell = CreateObject("WScript.Shell")
								WshShell.SendKeys "{DEL}"
								wait(2)
								WshShell.SendKeys "{ESC}"
								wait(2)
								Set WshShell = nothing
			
								If JavaDialog("Confirmation").Exist(2) Then
										If sButton = "" Then
											JavaDialog("Confirmation").JavaButton("Yes").Click micLeftBtn
											wait(1)
											Exit For		
										Else
											JavaDialog("Confirmation").JavaButton("No").Click micLeftBtn											
											wait(1)
											Exit For		
										End If
								End If
								
								
					Next

					If iX >= iFinish Then
							Fn_SchMgr_GanttChartOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find Dependency Delete Dialog")
							Exit Function
					End If

					Fn_SchMgr_GanttChartOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deleted Dependency in GanttChart between Task [" + aParam(0) + "] and Task [" + aParam(1) + "]")

   End Select
End Function



'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   :Fn_SchMgr_CostOperation(sAction,dicSchCost)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function is an replica of the function with the same name which is commented due to changes in controls and OR
'''''/$$$$
'''''/$$$$	
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS          07/04/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			07/04/2011         1.0
'''''/$$$$
'''''/$$$$	How To Use : 
'''''/$$$$									dicSchCost.RemoveAll()
'''''/$$$$									dicSchCost.Add "TotalEstimatedCost","$0.00"
'''''/$$$$									dicSchCost.Add "BD Name","t1"
'''''/$$$$									dicSchCost.Add "BD EstimatedHours","16h"
'''''/$$$$									dicSchCost.Add "BD Accured Cost","0"
'''''/$$$$									dicSchCost.Add "BillCode","unassigned"
'''''/$$$$									dicSchCost.Add "FC CostName","qwerty"
'''''/$$$$									dicSchCost.Add "FC EstimatedCost","$0.00"
'''''/$$$$									dicSchCost.Add "FC AccuredCost","$5.00"
'''''/$$$$
'''''/$$$$									bReturn=Fn_SchMgr_CostOperation("Verify",dicSchCost)
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SchMgr_CostOperation(sAction,dicSchCost)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_CostOperation"
   On Error Resume Next
   Dim dicCount , dicKeys , dicItems , iCounter,objWin ,bReturn,sIndex,objTable,objFixCostTable,sCount,iCount,jCounter
   Dim sNewItemMenu
'   Set objWin =  JavaWindow("ScheduleManagerWindow").Dialog("Costs")
	Set objWin = Fn_SISW_PPM_GetObject("Costs")
	sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"ScheduleCosts")
	If not objWin.Exist(SISW_MINLESS_TIMEOUT) Then
		bReturn = Fn_MenuOperation("Select",sNewItemMenu)
		If bReturn = False Then
				Fn_SchMgr_CostOperation = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:Costs]")
				Set objWin = Nothing
				Exit Function
		Else
				wait 2
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:Costs]")
		End If
	End If

			If instr(1,sAction,":")>0 Then
			 	If objWin.Exist(5) Then

		 dicCount  = dicSchCost.Count
		 dicItems = dicSchCost.Items
		 dicKeys = dicSchCost.Keys
		 aProperties=split(sAction,":",-1,1)
								 Select Case aProperties(0)
								 Case "Verify"
									For iCounter = 0 to dicCount - 1
										If  dicItems(iCounter) <> ""Then
					
											Select Case dicKeys(iCounter)
					
												Case "TotalEstimatedCost","TotalAccruedCost","TotalEstimatedWork","TotalAccruedWord", "TotalAccruedWork"
													If instr(lcase(dicKeys(iCounter)), lcase("Word")) > 0 Then
													  dicKeys(iCounter) = replace(dicKeys(iCounter), "Word", "Work")
													End If
													
													If  dicItems(iCounter) = objWin.Static(dicKeys(iCounter)).GetROProperty("text") Then
														Fn_SchMgr_CostOperation = TRUE 
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of " + dicKeys(iCounter))
													Else 
														Fn_SchMgr_CostOperation = FALSE 
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of " + dicKeys(iCounter))
														objWin.WinButton("Cancel").Click 0,0,micLeftBtn
														Set objWin = Nothing 
														Exit function 
													End If 
					
												Case "BillCode","BillSub-code","BillType","RateModifier"
													If dicItems(iCounter) =objWin.JavaList(dicKeys(iCounter)).GetVisibleText Then
														Fn_SchMgr_CostOperation = TRUE 
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of " + dicKeys(iCounter))
													Else 
														Fn_SchMgr_CostOperation = FALSE 
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of " + dicKeys(iCounter))
														objWin.WinButton("Cancel").Click 0,0,micLeftBtn
														Set objWin = Nothing 
														Exit function
													End If 
					
												Case "Rollup", "DrillDown"
													If objWin.JavaButton(dicKeys(iCounter)).CheckProperty ("enabled",1,20000) = true Then
														objWin.WinButton(dicKeys(iCounter)).Click micLeftBtn 
														Fn_SchMgr_CostOperation = TRUE
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click button " + dicKeys(iCounter))
													Else 
														Fn_SchMgr_CostOperation = FALSE 
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button " + dicKeys(iCounter))
														objWin.WinButton("Cancel").Click 0,0,micLeftBtn
														Set objWin = Nothing 
														Exit function
													End If
					
					
												Case "BD Name"
												
													If dicSchCost.Item("BD Name") <> "" Then
																	Set objTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("BreakdownTable")
																	sCount=objTable.GetItemsCount
																	For iCount=0 to sCount-1
																			sDetails =  objTable.GetSubItem(iCount,"Name")
																			 If lcase(dicSchCost.Item("BD Name"))=lcase(sDetails) Then
																				 jCounter=iCount
																			 End If
																	Next
													Else 
																		Fn_SchMgr_CostOperation = False
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get row from breakdown table for name " + dicSchCost.Item("Name")) 
																		objWin.WinButton("Cancel").Click 0,0,micLeftBtn
																		Set objTable = Nothing 
																		Set objWin = Nothing 
													End If
												
																	 If jCounter<0 then
																		 Fn_SchMgr_CostOperation = False
																		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to check if the name is present") 
																		 Exit function
																	 End If
					
					
					
												Case "BD EstimatedHours"
					
																	If jCounter>=0 Then
																		If dicSchCost.Item("BD EstimatedHours") <> "" Then
																			If dicSchCost.Item("BD EstimatedHours") = objTable.GetSubItem(jCounter,"Estimated Hours") Then
																				Fn_SchMgr_CostOperation = True
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  EstimatedHours") 
																			Else 
																				Fn_SchMgr_CostOperation = False
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  EstimatedHours") 
																				objWin.WinButton("Cancel").Click 0,0,micLeftBtn
																				Set objTable = Nothing 
																				Set objWin = Nothing 
																			End If
																		End If
																	End if
					
												Case "BD AccruedHours"
					
																	If jCounter>=0 Then
																		If dicSchCost.Item("BD AccruedHours") <> "" Then
																			If dicSchCost.Item("BD AccruedHours") = objTable.GetSubItem(jCounter,"Accrued Hours") Then
																				Fn_SchMgr_CostOperation = True
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  AccruedHours") 
																			Else 
																				Fn_SchMgr_CostOperation = False
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  AccruedHours") 
																				objWin.JavaButton.WinButton("Cancel").Click 0,0,micLeftBtn
																				Set objTable = Nothing 
																				Set objWin = Nothing 
																			End If
																		End If  
																	End if
					
													Case "BD EstimatedCost"
					
																		If jCounter>=0 Then
																		If dicSchCost.Item("BD EstimatedCost") <> "" Then
																			If dicSchCost.Item("BD EstimatedCost") = objTable.GetSubItem(jCounter,"Estimated Cost") Then
																				Fn_SchMgr_CostOperation = True
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  EstimatedCost") 
																			Else 
																				Fn_SchMgr_CostOperation = False
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  EstimatedCost") 
																				objWin.JavaButton.WinButton("Cancel").Click 0,0,micLeftBtn
																				Set objTable = Nothing 
																				Set objWin = Nothing 
																			End If
																		End If  
																	End if
					
												Case "BD AccruedCost"
					
																	If jCounter>=0 Then
																		If dicSchCost.Item("BD AccruedCost") <> "" Then
																			If dicSchCost.Item("BD AccruedCost") = objTable.GetSubItem(jCounter,"Accrued Cost") Then
																				Fn_SchMgr_CostOperation = True
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  AccruedCost") 
																			Else 
																				Fn_SchMgr_CostOperation = False
																				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  AccruedCost") 
																			objWin.JavaButton.WinButton("Cancel").Click 0,0,micLeftBtn
																				Set objTable = Nothing 
																				Set objWin = Nothing 
																			End If
																		End If 
																End if
																
					
											Case "FC CostName"
					
												If dicSchCost.Item("FC CostName") <> "" Then
														Set objFixCostTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("FixedCostsTable")
														sCount=objFixCostTable.GetItemsCount
														For iCount=0 to sCount-1
																sDetails =  objFixCostTable.GetSubItem(iCount,"Cost Name")
																 If lcase(dicSchCost.Item("FC CostName"))=lcase(sDetails) Then
																	 jCounter=iCounter
																	 Exit for
																 End If
														Next
													End if
					
					
										Case "FC EstimatedCost"
					
												If jCounter>=0 Then
														If dicSchCost.Item("FC EstimatedCost") <> "" Then
																If dicSchCost.Item("FC EstimatedCost") = objFixCostTable.GetSubItem(jCounter,"Estimated Cost") Then
																	Fn_SchMgr_CostOperation = True
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  EstimatedHours") 
																Else 
																	Fn_SchMgr_CostOperation = False
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  EstimatedHours") 
																	objWin.WinButton("Cancel").Click 0,0,micLeftBtn
																	Set objFixCostTable = Nothing 
																	Set objWin = Nothing 
																End If
														End If
					
												End If
					
										Case "FC AccuredCost"
											If jCounter >= 0  Then
																	If dicSchCost.Item("FC AccuredCost") <> "" Then
																		If dicSchCost.Item("FC AccuredCost") = objFixCostTable.GetSubItem(jCounter,"Accured Cost") Then
																			Fn_SchMgr_CostOperation = True
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  AccruedHours") 
																		Else 
																			Fn_SchMgr_CostOperation = False
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  AccruedHours") 
																			objWin.JavaButton.WinButton("Cancel").Click 0,0,micLeftBtn
																			Set objFixCostTable = Nothing 
																			Set objWin = Nothing 
																		End If
																	End If  
											End If
					
									End Select
								End If
							Next
						End Select
				Else
					Fn_SchMgr_CostOperation = FALSE
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to displayed Costs dialog.")
					Exit Function
				End If 

				If lCase(aProperties(1))="cancel" Then
						objWin.WinButton("Cancel").Click 0,0,micLeftBtn
				Else
				objWin.WinButton("Finish").Click 0,0,micLeftBtn
				End If
Set objWin = Nothing
	Else
		If objWin.Exist(5) Then

		 dicCount  = dicSchCost.Count
		 dicItems = dicSchCost.Items
		 dicKeys = dicSchCost.Keys
		 aProperties=split(sAction,":",-1,1)

		Select Case sAction
			Case "Verify"
				For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> ""Then

						Select Case dicKeys(iCounter)

							Case "TotalEstimatedCost","TotalAccruedCost","TotalEstimatedWork","TotalAccruedWord", "TotalAccruedWork"
								If instr(lcase(dicKeys(iCounter)), lcase("Word")) > 0 Then
								  dicKeys(iCounter) = replace(dicKeys(iCounter), "Word", "Work")
								End If
								
								If  dicItems(iCounter) = objWin.Static(dicKeys(iCounter)).GetROProperty("text") Then
									Fn_SchMgr_CostOperation = TRUE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of " + dicKeys(iCounter))
								Else 
									Fn_SchMgr_CostOperation = FALSE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of " + dicKeys(iCounter))
									objWin.WinButton("Cancel").Click 0,0,micLeftBtn
									Set objWin = Nothing 
									Exit function 
								End If 

							Case "BillCode","BillSub-code","BillType","RateModifier"
								If dicItems(iCounter) =objWin.JavaList(dicKeys(iCounter)).GetVisibleText Then
									Fn_SchMgr_CostOperation = TRUE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of " + dicKeys(iCounter))
								Else 
									Fn_SchMgr_CostOperation = FALSE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of " + dicKeys(iCounter))
									objWin.WinButton("Cancel").Click 0,0,micLeftBtn
									Set objWin = Nothing 
									Exit function
								End If 

							Case "Rollup", "DrillDown"
								If objWin.JavaButton(dicKeys(iCounter)).CheckProperty ("enabled",1,20000) = true Then
									objWin.WinButton(dicKeys(iCounter)).Click micLeftBtn 
									Fn_SchMgr_CostOperation = TRUE
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click button " + dicKeys(iCounter))
								Else 
									Fn_SchMgr_CostOperation = FALSE 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click button " + dicKeys(iCounter))
									objWin.WinButton("Cancel").Click 0,0,micLeftBtn
									Set objWin = Nothing 
									Exit function
								End If

							Case "BD Name"
							
								If dicSchCost.Item("BD Name") <> "" Then
												Set objTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("BreakdownTable")
												sCount=objTable.GetItemsCount
												For iCount=0 to sCount-1
														sDetails =  objTable.GetSubItem(iCount,"Name")
														 If lcase(dicSchCost.Item("BD Name"))=lcase(sDetails) Then
															 	Fn_SchMgr_CostOperation = True
															 jCounter=iCount
														 End If
												Next
								Else 
													Fn_SchMgr_CostOperation = False
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get row from breakdown table for name " + dicSchCost.Item("Name")) 
													objWin.WinButton("Cancel").Click 0,0,micLeftBtn
													Set objTable = Nothing 
													Set objWin = Nothing 
								End If
							
												 If jCounter<0 then
													 Fn_SchMgr_CostOperation = False
													 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to check if the name is present") 
													 Exit function
												 End If


							Case "BD EstimatedHours"

												If jCounter>=0 Then
													If dicSchCost.Item("BD EstimatedHours") <> "" Then
														If dicSchCost.Item("BD EstimatedHours") = objTable.GetSubItem(jCounter,"Estimated Hours") Then
															Fn_SchMgr_CostOperation = True
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  EstimatedHours") 
														Else 
															Fn_SchMgr_CostOperation = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  EstimatedHours") 
															objWin.WinButton("Cancel").Click 0,0,micLeftBtn
															Set objTable = Nothing 
															Set objWin = Nothing 
														End If
													End If
												End if

							Case "BD AccruedHours"

												If jCounter>=0 Then
													If dicSchCost.Item("BD AccruedHours") <> "" Then
														If dicSchCost.Item("BD AccruedHours") = objTable.GetSubItem(jCounter,"Accrued Hours") Then
															Fn_SchMgr_CostOperation = True
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  AccruedHours") 
														Else 
															Fn_SchMgr_CostOperation = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  AccruedHours") 
															objWin.JavaButton.WinButton("Cancel").Click 0,0,micLeftBtn
															Set objTable = Nothing 
															Set objWin = Nothing 
														End If
													End If  
												End if

								Case "BD EstimatedCost"

													If jCounter>=0 Then
													If dicSchCost.Item("BD EstimatedCost") <> "" Then
														If dicSchCost.Item("BD EstimatedCost") = objTable.GetSubItem(jCounter,"Estimated Cost") Then
															Fn_SchMgr_CostOperation = True
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  EstimatedCost") 
														Else 
															Fn_SchMgr_CostOperation = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  EstimatedCost") 
															objWin.JavaButton.WinButton("Cancel").Click 0,0,micLeftBtn
															Set objTable = Nothing 
															Set objWin = Nothing 
														End If
													End If  
												End if

							Case "BD AccruedCost"

												If jCounter>=0 Then
													If dicSchCost.Item("BD AccruedCost") <> "" Then
														If dicSchCost.Item("BD AccruedCost") = objTable.GetSubItem(jCounter,"Accrued Cost") Then
															Fn_SchMgr_CostOperation = True
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  AccruedCost") 
														Else 
															Fn_SchMgr_CostOperation = False
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  AccruedCost") 
														objWin.JavaButton.WinButton("Cancel").Click 0,0,micLeftBtn
															Set objTable = Nothing 
															Set objWin = Nothing 
														End If
													End If 
											End if
											

						Case "FC CostName"

							If dicSchCost.Item("FC CostName") <> "" Then
									Set objFixCostTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("FixedCostsTable")
									sCount=objFixCostTable.GetItemsCount
									For iCount=0 to sCount-1
											sDetails =  objFixCostTable.GetSubItem(iCount,"Cost Name")
											 If lcase(dicSchCost.Item("FC CostName"))=lcase(sDetails) Then
												 jCounter=iCounter
												 Exit for
											 End If
									Next
								End if
			

					Case "FC EstimatedCost"

							If jCounter>=0 Then
									If dicSchCost.Item("FC EstimatedCost") <> "" Then
											If dicSchCost.Item("FC EstimatedCost") = objFixCostTable.GetSubItem(jCounter,"Estimated Cost") Then
												Fn_SchMgr_CostOperation = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  EstimatedHours") 
											Else 
												Fn_SchMgr_CostOperation = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  EstimatedHours") 
												objWin.WinButton("Cancel").Click 0,0,micLeftBtn
												Set objFixCostTable = Nothing 
												Set objWin = Nothing 
											End If
									End If

							End If

					Case "FC AccuredCost"
						If jCounter >= 0  Then
												If dicSchCost.Item("FC AccuredCost") <> "" Then
													If dicSchCost.Item("FC AccuredCost") = objFixCostTable.GetSubItem(jCounter,"Accured Cost") Then
														Fn_SchMgr_CostOperation = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify value of  AccruedHours") 
													Else 
														Fn_SchMgr_CostOperation = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify value of  AccruedHours") 
														objWin.JavaButton.WinButton("Cancel").Click 0,0,micLeftBtn
														Set objFixCostTable = Nothing 
														Set objWin = Nothing 
													End If
									End If  
						End If

						End Select
					
					End If
				Next
' Added [ CancelButtonClick ] code By : Harshal Tanpure , Date=>9-August-2011
If dicSchCost.Item("CancelButtonClick") = "" Then
objWin.WinButton("Cancel").Click 0,0,micLeftBtn
End If
					 Case "GetValue"
				For iCounter = 0 to dicCount - 1
					Select Case dicKeys(iCounter)
	
						Case "TotalEstimatedCost","TotalAccruedCost","TotalEstimatedWork", "TotalAccruedWork"
							sValue = objWin.Static(dicKeys(iCounter)).GetROProperty("text")
							Fn_SchMgr_CostOperation = sValue 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get value of " + dicKeys(iCounter))
				
	
						Case "BD EstimatedHours"
							If dicSchCost.Item("BD Name") <> "" Then
								Set objTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("BreakdownTable")
								sCount=objTable.GetItemsCount
								For iCount=0 to sCount-1
									sDetails =  objTable.GetSubItem(iCount,"Name")
									 If lcase(dicSchCost.Item("BD Name"))=lcase(sDetails) Then
										sValue = objTable.GetSubItem(iCount,"Estimated Hours")
										Fn_SchMgr_CostOperation = sValue 
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get value of " + dicKeys(iCounter))
									 End If
								Next
							End If
	
						Case "BD AccruedHours"
							If dicSchCost.Item("BD Name") <> "" Then
								Set objTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("BreakdownTable")
								sCount=objTable.GetItemsCount
								For iCount=0 to sCount-1
									sDetails =  objTable.GetSubItem(iCount,"Name")
									 If lcase(dicSchCost.Item("BD Name"))=lcase(sDetails) Then
										sValue = objTable.GetSubItem(iCount,"Accrued Hours")
										Fn_SchMgr_CostOperation = sValue 
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get value of " + dicKeys(iCounter))
									 End If
								Next
							End If
	
						Case "BD EstimatedCost"
							If dicSchCost.Item("BD Name") <> "" Then
								Set objTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("BreakdownTable")
								sCount=objTable.GetItemsCount
								For iCount=0 to sCount-1
									sDetails =  objTable.GetSubItem(iCount,"Name")
'									If instr(1,sValue,"Example:")>0 Then
									If instr(1,lcase(dicSchCost.Item("BD Name")),lcase(sDetails))>0 Then
'									 If lcase(dicSchCost.Item("BD Name"))=lcase(sDetails) Then
										sValue = objTable.GetSubItem(iCount,"Estimated Cost")
										Fn_SchMgr_CostOperation = sValue 
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get value of " + dicKeys(iCounter))
									 End If
								Next
							End If
	
	
						Case "BD AccruedCost"
							If dicSchCost.Item("BD Name") <> "" Then
								Set objTable = JavaWindow("ScheduleManagerWindow").Dialog("Costs").WinListView("BreakdownTable")
								sCount=objTable.GetItemsCount
								For iCount=0 to sCount-1
									sDetails =  objTable.GetSubItem(iCount,"Name")
									If instr(1,lcase(dicSchCost.Item("BD Name")),lcase(sDetails))>0 Then
'									 If lcase(dicSchCost.Item("BD Name"))=lcase(sDetails) Then
										sValue = objTable.GetSubItem(iCount,"Accrued Cost")
										Fn_SchMgr_CostOperation = sValue 
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get value of " + dicKeys(iCounter))
									 End If
								Next
							End If		
	
					End Select
				Next
				
' Added [ CancelButtonClick ] code By : Harshal Tanpure , Date=>9-August-2011
If dicSchCost.Item("CancelButtonClick") = "" Then
objWin.WinButton("Cancel").Click 0,0,micLeftBtn
End If
		Case "DialogVerify"					''Added By Vidya
					 If objWin.Exist(3) Then 
						 Fn_SchMgr_CostOperation = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verify The Existance Of The Cost Dialog.") 
					Else 
						Fn_SchMgr_CostOperation = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed ToVerify The Existance Of The Cost Dialog.") 
						Exit Function
					End If 
			objWin.WinButton("Cancel").Click 0,0,micLeftBtn

		End Select
			
	Else
		Fn_SchMgr_CostOperation = FALSE
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to displayed Costs dialog.")
		Exit Function
	End If 
End if
	Set objWin = Nothing 

End Function 



''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''''/$$$$
''''''/$$$$   FUNCTION NAME   :Fn_SchMgr_CostBillValues_Operations(sAction,sValutToSelect,aBillSubCode,sBillType,sRateModifier)
''''''/$$$$
''''''/$$$$   DESCRIPTION        :  This function will perform various operations on the WinComboBoxes in the cost dialog
''''''/$$$$
''''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''''/$$$$										sValutToSelect : Value to beselected from 'BillCode' combobox
''''''/$$$$										aBillSubCode : Array of values to be verified
''''''/$$$$										sBillType : for future use
''''''/$$$$										sRateModifier : for future use
''''''/$$$$	
''''''/$$$$		Return Value : 				True or False
''''''/$$$$
''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''''/$$$$										
''''''/$$$$
''''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''''/$$$$
''''''/$$$$    CREATED BY     :   SHREYAS          11/04/2011         1.0
''''''/$$$$
''''''/$$$$    REVIWED BY     :  Prasanna			11/04/2011         1.0
''''''/$$$$
''''''/$$$$		How To Use :   aValue=array("unassigned","Accounting","Clerical","CorpAdmin","IT")
''''''/$$$$								bReturn=Fn_SchMgr_CostBillValues_Operations("Verify","General",aValue,"","")
''''''/$$$$
''''''/$$$$									
''''''/$$$$
''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$



Public function Fn_SchMgr_CostBillValues_Operations(sAction,sValutToSelect,aBillSubCode,sBillType,sRateModifier)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_CostBillValues_Operations"
Dim iCounter,i,sDetails,sCount,objCost

Fn_SchMgr_CostBillValues_Operations=false
Set objCost = JavaWindow("ScheduleManagerWindow").Dialog("Costs")

'Invoke the costs dialog if not present

 If Not objCost.Exist(3) Then

	 	bReturn = Fn_MenuOperation("Select","Schedule:Costs")
		wait(7)
				If bReturn = True Then
					Fn_SchMgr_CostBillValues_Operations=true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule:Costs.]")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Costs.]")
						Fn_SchMgr_CostBillValues_Operations=false
					Set objCost = Nothing 
					Exit Function
				End if
 End if

 Select Case sAction

	   	Case "Verify"

				If sValutToSelect<>"" Then

					'set the value in the Bill-code WinComboBox
					objCost.WinComboBox("BillCode").Select 	sValutToSelect
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the value " + sValutToSelect)
						objCost.WinButton("Cancel").Click micLeftBtn
						Set objCost = Nothing
						Fn_SchMgr_CostBillValues_Operations=false
						Exit Function
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the value " + sValutToSelect)
						Fn_SchMgr_CostBillValues_Operations=true
					End If
	
				End If
				Wait SISW_MIN_TIMEOUT
						If IsArray(aBillSubCode) Then
								For iCounter = 0 to Ubound(aBillSubCode)
										
										sCount=objCost.WinComboBox("BillSub-code").GetROProperty("items count")
										For i=0 to sCount-1
										sDetails=objCost.WinComboBox("BillSub-code").GetItem(i)
										If lcase(aBillSubCode(i))=lcase(sDetails) Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified  the value " + aBillSubCode(i))
											If cstr(i)=cstr(Ubound(aBillSubCode)) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Verification Process Complete")
		
													'Click on the cancel button
													objCost.WinButton("Cancel").Click micLeftBtn
													If Err.Number < 0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on the cancel button")
																objCost.WinButton("Cancel").Click micLeftBtn
																Set objCost = Nothing
																Fn_SchMgr_CostBillValues_Operations=false
																Exit Function
													Else 
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the cancel button")
																Fn_SchMgr_CostBillValues_Operations=True
																Exit Function
													End If
						
												Exit for
											End If
											
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the value " +aBillSubCode(i))
											objCost.WinButton("Cancel").Click micLeftBtn
											Set objCost = Nothing
											Fn_SchMgr_CostBillValues_Operations=false
											Exit Function
										End If
									Next
								Next
						End if
		


		

					If IsArray(sBillType) Then
							For iCounter = 0 to Ubound(sBillType)
		
									sCount=objCost.WinComboBox("BillType").GetROProperty("items count")
									For i=0 to sCount-1
									sDetails=objCost.WinComboBox("BillType").GetItem(i)
									If lcase(sBillType(i))=lcase(sDetails) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified  the value " + sBillType(i))
										If cstr(i)=cstr(Ubound(sBillType)) Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Verification Process Complete")
	
												'Click on the Cancel button
												objCost.WinButton("Cancel").Click micLeftBtn
												If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on the cancel button")
															objCost.WinButton("Cancel").Click micLeftBtn
															Set objCost = Nothing
															Fn_SchMgr_CostBillValues_Operations=false
															Exit Function
												Else 
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the cancel button")
															Fn_SchMgr_CostBillValues_Operations=True
															Exit Function
												End If
					
											Exit for
										End If
										
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the value " +sBillType(i))
										objCost.WinButton("Cancel").Click micLeftBtn
										Set objCost = Nothing
										Fn_SchMgr_CostBillValues_Operations=false
										Exit Function
									End If
								Next
							Next
					End if
	

		Set objCost = Nothing

Case "SetValues"

	If sValutToSelect<>"" Then

					'set the value in the Bill-code WinComboBox
					objCost.WinComboBox("BillCode").Select 	sValutToSelect
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the value " + sValutToSelect)
						objCost.WinButton("Cancel").Click micLeftBtn
						Set objCost = Nothing
						Fn_SchMgr_CostBillValues_Operations=false
						Exit Function
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the value " + sValutToSelect)
						Fn_SchMgr_CostBillValues_Operations=true
					End If
	
				End If

			

						If aBillSubCode<>"" Then

					'set the value in the Bill-code WinComboBox
					objCost.WinComboBox("BillSub-code").Select 	aBillSubCode
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the value " + aBillSubCode)
						objCost.WinButton("Cancel").Click micLeftBtn
						Set objCost = Nothing
						Fn_SchMgr_CostBillValues_Operations=false
						Exit Function
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the value " + aBillSubCode)
						Fn_SchMgr_CostBillValues_Operations=true
					End If
	
				End If
		


		

						If sBillType<>"" Then

					'set the value in the Bill-code WinComboBox
					objCost.WinComboBox("BillType").Select 	sBillType
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the value " + sBillType)
						objCost.WinButton("Cancel").Click micLeftBtn
						Set objCost = Nothing
						Fn_SchMgr_CostBillValues_Operations=false
						Exit Function
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the value " + sBillType)
						Fn_SchMgr_CostBillValues_Operations=true
					End If
	
				End If

		If sRateModifier<>"" Then

					'set the value in the Bill-code WinComboBox
					objCost.WinComboBox("RateModifier").Select 	sRateModifier
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the value " + sRateModifier)
						objCost.WinButton("Cancel").Click micLeftBtn
						Set objCost = Nothing
						Fn_SchMgr_CostBillValues_Operations=false
						Exit Function
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the value " + sRateModifier)
						Fn_SchMgr_CostBillValues_Operations=true
					End If
	
				End If

			'Click on the Finish button
			objCost.WinButton("Finish").Click micLeftBtn

		Set objCost = Nothing

	End Select
End Function


''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_SchMgr_DefineWBSFormat(sAction,sSequence,sLength,sSeperator,sButtons,sInfo)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will perform various operations on the WinComboBoxes in the cost dialog
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sSequence : Value to be selected in the sequence field
''''/$$$$										sLength : Value to be selected in the Length field
''''/$$$$										sSeperator : Value to be selected in the Seperator field
''''/$$$$										sButtons : Button to be clicked
''''/$$$$										sInfo : For Future Use
''''/$$$$	
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          15/04/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Prasanna			15/04/2011         1.0
''''/$$$$
''''/$$$$		How To Use :   bReturn=Fn_SchMgr_DefineWBSFormat("Verify","Number","","","Cancel","")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("Verify_Added&RemovedEntry","Remove","0","","Cancel","")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("VerifyFormat","asfgjagf","","","Cancel","")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("Set&Verify","","N.N","","","1")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("CheckButtons","","","","","Delete")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("VerifyLevelLimit","","","","","")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("VerifyLevelAfterRemoved","","4","","","")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("Regenerate","","","","","Yes")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("RegenerateAll","","","","","No")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("VerifyFormat&Example","N.N","1.2","","","")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("VerifySequence","uppercases","","","","3")
''''/$$$$								bReturn=Fn_SchMgr_DefineWBSFormat("AddLevel","","","","","1")
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SchMgr_DefineWBSFormat(sAction,sSequence,sLength,sSeperator,sButtons,sInfo)
GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_DefineWBSFormat"
Dim objSelectType,objClassAdmin,WshShell,iCount,sDetails,sValue,aProperties,sStaticCount,objFormat,objRegenerate,bFlag1,bFlag2
Dim sMessage,objSave
Fn_SchMgr_DefineWBSFormat=false

If  JavaWindow("ScheduleManagerWindow").JavaWindow("Define WBS Format").Exist(5) Then
		Set objFormat= JavaWindow("ScheduleManagerWindow").JavaWindow("Define WBS Format")
elseif JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("DefineWBSFormat 2").Exist(5) then
		Set objFormat= JavaWindow("ScheduleManagerWindow").JavaWindow("Shell").JavaWindow("DefineWBSFormat 2")
else
		bReturn = Fn_MenuOperation("Select","Schedule:WBS:Define Format")
		If bReturn = false Then
				Fn_SchMgr_DefineWBSFormat = False		
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:WBS:Define Format]")				
				Exit Function
		Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:WBS:Define Format]")
				Set objFormat= JavaWindow("ScheduleManagerWindow").JavaWindow("Define WBS Format")
		End If
End If




	If objFormat.Exist(3) Then

Select Case sAction


	Case "Verify"	

				If sSequence<>""  Then
				If sInfo<>"" Then
					objFormat.JavaTable("Define").ActivateCell cint(sInfo),"Sequence"
				Else
					objFormat.JavaTable("Define").ActivateCell 0,"Sequence"
				End If
				
									Set objSelectType = description.Create()
											objSelectType("Class Name").value = "JavaList"
											Set objClassAdmin = objFormat.ChildObjects(objSelectType)
											sStaticCount=objFormat.ChildObjects(objSelectType).count
											For iCount=0 to sStaticCount-1
											   objClassAdmin(iCount).Select (sSequence)
											Next
				
																Set WshShell = CreateObject("WScript.Shell")
																WshShell.SendKeys "{ENTER}"
																WshShell.SendKeys "{ENTER}"
																Set WshShell = nothing
				
				objFormat.JavaTable("Define").ClickCell 1,"Level"
		
				Set objSelectType = description.Create()
										objSelectType("Class Name").value = "JavaStaticText"
										Set objClassAdmin = objFormat.ChildObjects(objSelectType)
										sStaticCount= objFormat.ChildObjects(objSelectType).count
										For iCount=0 to sStaticCount-1
										   sValue=  objClassAdmin(iCount).GetROProperty ("label")
										   If instr(1,sValue,"Example:")>0 Then
												iCount=iCount+1
												sDetails=objClassAdmin(iCount).GetROProperty ("label")
												aProperties=split(sDetails,".",-1,1)
												If  isNumeric(aProperties(0))=true and lcase(sSequence)="number" Then
													Fn_SchMgr_DefineWBSFormat=true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the format "+sSequence)
													Exit For
												Elseif isNumeric(aProperties(0))=false and (asc(aProperties(0))>=65 and asc(aProperties(0))<=90) and lcase(sSequence)="uppercase" then
													Fn_SchMgr_DefineWBSFormat = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the format "+sSequence)
													Exit For
												Elseif isNumeric(aProperties(0))=false and (asc(aProperties(0))>=97 and asc(aProperties(0))<=122) and lcase(sSequence)="lowercase" then
													Fn_SchMgr_DefineWBSFormat = true
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the format "+sSequence)
													Exit For
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the format "+sSequence)
													Fn_SchMgr_DefineWBSFormat = False
													objFormat.JavaButton("Cancel").Click micLeftBtn
													Exit function
												End If
										   End 	if	
										Next
							End If
		
							If sLength<>"" Then
								'for future use
							End If
		
					If sSeperator<>"" Then
								'for future use
					End If

		Case "VerifyFormat"

					If sSequence<>"" Then
						'set some value other than the default value in the Format: Field
						objFormat.JavaEdit("Format").Set (sSequence)
						If err.number<0 Then
							Fn_SchMgr_DefineWBSFormat = False
							
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the value ["+sSequence+"] in the Format edit box")
								Set objTask = Nothing
								Exit function
								
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the value ["+sSequence+"] in the Format edit box")
						End If
					End If

					'set some value in the initial value field
					If sLength<>"" Then
						objFormat.JavaEdit("Initial Value").Set (sLength)
								If err.number<0 Then
							Fn_SchMgr_DefineWBSFormat = False
							
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the value ["+sLength+"] in the Initial Value edit box")
								Set objTask = Nothing
								Exit function
								
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the value ["+sLength+"] in the Initial Value edit box")
						End if
					End If

					
		
					'Click on the verify button
					objFormat.JavaButton("Verify").Click
						If err.number<0 Then
							Fn_SchMgr_DefineWBSFormat = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on the Verify Button")
								Set objTask = Nothing
								Exit function
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on the Verify Button")
						End If
			
						'Check the existence of the "Format" dialog in case of wrong entries
						Set objFormatError=JavaWindow("ScheduleManagerWindow").JavaWindow("Define WBS Format").JavaWindow("Format")
						objFormatError.SetTOProperty "title","Format"		''' added objFormatError
						If objFormatError.Exist(3) Then
							objFormatError.JavaButton("OK").Click micLeftBtn
							If err.number<0 Then
									Fn_SchMgr_DefineWBSFormat = true									
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  handle the Format Dialog")
										Set objTask = Nothing
										Exit Function
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully handled the Format Dialog")
										Fn_SchMgr_DefineWBSFormat = false
										objFormat.JavaButton(sButtons).Click micLeftBtn
										Exit Function
								End If
			
						End If

							'Check the existence of the "Confirm" dialog in case of wrong entries
						Set objFormatConfirm=JavaWindow("ScheduleManagerWindow").JavaWindow("Define WBS Format").JavaWindow("Format")
						objFormatConfirm.SetTOProperty "title","Confirm"
						If objFormatConfirm.Exist(3) Then
							'objFormatConfirm.JavaButton("OK").SetTOProperty "label",sInfo
							'objFormatConfirm.JavaButton("OK").Click micLeftBtn
							objFormatConfirm.JavaButton("No").Click micLeftBtn
							If err.number<0 Then
									Fn_SchMgr_DefineWBSFormat = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  handle the Confirm Dialog")
										Set objTask = Nothing
										Exit Function
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully handled the Confirm Dialog")
										Fn_SchMgr_DefineWBSFormat = True
								End If
			
						End If

			Case "Verify_Added&RemovedEntry"

				If sSequence<>"" Then
					If sSequence="Add" Then

						'get the number of rows
						sCount=objFormat.JavaTable("Define").GetROProperty ("rows")

						'click on add button
						objFormat.JavaButton("Add").Click micLeftBtn

						'now again get the row count
						sCount1=objFormat.JavaTable("Define").GetROProperty ("rows")
						If cint(sCount1)=cint(sCount)+1 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that a new row is added")
							Fn_SchMgr_DefineWBSFormat = True
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify that a new row is added")
							Fn_SchMgr_DefineWBSFormat = False
							Exit function
						End If

				Elseif sSequence="Remove" then
						
						'get the number of rows
						sCount=objFormat.JavaTable("Define").GetROProperty ("rows")

						'Delete the required row
						objFormat.JavaTable("Define").SelectRow(sLength)

						'click on Delete button
						objFormat.JavaButton("Delete").Click micLeftBtn

							'now again get the row count
						sCount1=objFormat.JavaTable("Define").GetROProperty ("rows")
						If cint(sCount1)=cInt(sCount-1) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that the row is removed")
							Fn_SchMgr_DefineWBSFormat = True
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify that the row is removed")
							Fn_SchMgr_DefineWBSFormat = False
							Exit function
						End If


					End If
				End If

		Case "Set&Verify"

							If  sSequence<>"" Then
								objFormat.JavaTable("Define").ActivateCell 0,"Sequence"
												Set objSelectType = description.Create()
														objSelectType("Class Name").value = "JavaList"
														Set objClassAdmin = objFormat.ChildObjects(objSelectType)
														sStaticCount=objFormat.ChildObjects(objSelectType).count
														For iCount=0 to sStaticCount-1
														   objClassAdmin(iCount).Select (sSequence)
														Next
							
														Set WshShell = CreateObject("WScript.Shell")
														WshShell.SendKeys "{ENTER}"
														WshShell.SendKeys "{ENTER}"
														Set WshShell = nothing
							
							objFormat.JavaTable("Define").ClickCell 1,"Level"
							Fn_SchMgr_DefineWBSFormat = True
			
						End If

						If sLength<>"" Then
							'for future use
						End If


						If sSeperator<>"" Then

							objFormat.JavaTable("Define").ActivateCell 0,"Separator"
												Set objSelectType = description.Create()
														objSelectType("Class Name").value = "JavaList"
														Set objClassAdmin = objFormat.ChildObjects(objSelectType)
														sStaticCount=objFormat.ChildObjects(objSelectType).count
														For iCount=0 to sStaticCount-1
														   objClassAdmin(iCount).Select (sSeperator)
														Next
							
														Set WshShell = CreateObject("WScript.Shell")
														WshShell.SendKeys "{ENTER}"
														WshShell.SendKeys "{ENTER}"
														Set WshShell = nothing
							
							objFormat.JavaTable("Define").ClickCell 1,"Level"
							Fn_SchMgr_DefineWBSFormat = True

						End If


						If sLength<>"" and sInfo<>"" Then

							'get the value from the Format & initial values field
							sValue=objFormat.JavaEdit("Format").GetROProperty ("value")
							sDetails=objFormat.JavaEdit("Initial Value").GetROProperty ("value")
								If lCase(sValue)=lcase(sLength) and  lCase(sDetails)=lcase(sInfo) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that the initial values & format values are present as per the format")
									Fn_SchMgr_DefineWBSFormat = True
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify that the initial values & format values are present as per the format")
									Fn_SchMgr_DefineWBSFormat = False
									Exit function
								End If
						End If

			Case "CheckButtons"

				Select Case sInfo

					Case "Delete"

						'first take the row count of the table by default there are 2 rows,in this case the delete button is disabled
						sValue=objFormat.JavaTable("Define").GetROProperty ("rows")
						If sValue="2" Then
							sDetails=objFormat.JavaButton("Delete").GetROProperty ("enabled")
							If sDetails="0" Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that the Delete button is greyed out when there are two rows by default")
									Fn_SchMgr_DefineWBSFormat = True
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Button Is Enabled")
							End If
						End If
				End Select

			Case "VerifyLevelLimit"

				'by Default there are 2 entries & the limit is for 10 entries,hence ideally,8 more entries should be added & after 10 entries the Add button gets greyed out

				For iCount=0 to 7
					objFormat.JavaButton("Add").Click micLeftBtn
					wait(1)
								If err.number<0 Then
									Fn_SchMgr_DefineWBSFormat = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on the add button")
									Exit Function
								Else
								If iCount=0  Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click on the add button "+cstr( iCount+1)+"st time")
								Elseif iCount=1 then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click on the add button "+cstr( iCount+1)+"nd time")
								Elseif iCount=2 then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click on the add button "+cstr( iCount+1)+"rd time")
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully click on the add button "+cstr( iCount+1)+"th time")
								End If
									Fn_SchMgr_DefineWBSFormat = True
							  End If
				Next

				'check if the add button is greyed out
				If objFormat.JavaButton("Add").GetROProperty("enabled")="0" Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that the Add button is greyed out")
						Fn_SchMgr_DefineWBSFormat = True
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Add Button Is still Enabled")
						Fn_SchMgr_DefineWBSFormat = False
						objFormat.Close
						Exit Function
				End If

		Case "VerifyLevelAfterRemoved"



						'Delete the required row
						objFormat.JavaTable("Define").SelectRow(sLength)

						'click on Delete button
						objFormat.JavaButton("Delete").Click micLeftBtn

						sDetails=JavaWindow("ScheduleManagerWindow").JavaWindow("Define WBS Format").JavaTable("Define").GetROProperty("rows")
							For iCount=0 to sDetails-1
									sValue=JavaWindow("ScheduleManagerWindow").JavaWindow("Define WBS Format").JavaTable("Define").GetCellData(iCount,"Level")
									If instr(1,cint(sValue),cint(sLength+1))>0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The level "+cstr(sLength+1)+" is still present after deleting it" )
											Fn_SchMgr_DefineWBSFormat = True
											Exit for
									Else
											Fn_SchMgr_DefineWBSFormat = False
									End If
							Next

			Case "Regenerate","RegenerateAll"
				If  objFormat.Exist(3)=true then
					objFormat.Close
				End if

				If sAction="Regenerate" Then
						Call  Fn_MenuOperation("Select","Schedule:WBS:Regenerate")
						If err.number<0 Then
							Fn_SchMgr_DefineWBSFormat = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:WBSRegenerate]")
								Set objTask = Nothing
								Exit Function
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:WBS:Regenerate]")
						End If

					Set objRegenerate=JavaDialog("Confirmation")
					objRegenerate.SetTOProperty "title","Regenerate WBS"
					
						If sInfo="Yes" Then
							objRegenerate.JavaButton("Yes").Click micLeftBtn
							Fn_SchMgr_DefineWBSFormat = True
							Exit function
						Elseif sInfo="No" Then
							objRegenerate.JavaButton("No").Click micLeftBtn
							Fn_SchMgr_DefineWBSFormat = True
							Exit function
						Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Button is to be clicked" )
						End If


						If sInfo <> "" Then
							sMessage =JavaDialog("Confirmation").JavaObject("MLabel").Object.getText()
						  If sInfo=sMessage Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified the confirmation message:"+sInfo)
									Fn_SchMgr_DefineWBSFormat = True
									objRegenerate.Close
									Exit Function

						  Else
									Fn_SchMgr_DefineWBSFormat = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Confirnation Message:"+sInfo)
									objRegenerate.Close
									Exit Function
							End If		
					End If

				Elseif sAction="RegenerateAll" Then
								Call  Fn_MenuOperation("Select","Schedule:WBS:Regenerate All")
						If err.number<0 Then
							Fn_SchMgr_DefineWBSFormat = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:WBSRegenerate All]")
								Set objTask = Nothing
								Exit Function
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:WBS:Regenerate All]")
						End If

					Set objRegenerate=JavaDialog("Confirmation")
					objRegenerate.SetTOProperty "title","Regenerate All"					
						If sInfo="Yes" Then
							objRegenerate.JavaButton("Yes").Click micLeftBtn
							Fn_SchMgr_DefineWBSFormat = True
							Exit Function
						Elseif sInfo="No" Then
							objRegenerate.JavaButton("No").Click micLeftBtn
							Fn_SchMgr_DefineWBSFormat = True
							Exit Function
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Button is to be clicked" )
						End If

						
					  If sInfo <>"" Then
							 sMessage =JavaDialog("Confirmation").JavaObject("MLabel").Object.getText()
							  If sInfo=sMessage Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified the confirmation message:"+sInfo)
										Fn_SchMgr_DefineWBSFormat = True
										objRegenerate.Close
										Exit Function
							  Else
										Fn_SchMgr_DefineWBSFormat = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Confirnation Message:"+sInfo)
										objRegenerate.Close
										Exit Function
								End If	
					End If		
						
				End If

	Case "Save","SaveAll"
				If  objFormat.Exist(5)=True then
					objFormat.Close
				End if

'			  If sAction="Save" OR sAction="SaveAll" Then

						If sAction="Save"	Then
						JavaWindow("ScheduleManagerWindow").JavaMenu("Schedule").JavaMenu("WBS").JavaMenu("Save").Select
						ElseIf sAction="SaveAll" Then
									bReturn =  Fn_MenuOperation("Select","Schedule:WBS:Save All")
									If bReturn  = false Then
											Fn_SchMgr_DefineWBSFormat = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:WBS:"+sAction+"]")
											Set objTask = Nothing
											Exit Function
									Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:WBS:"+sAction+"]")
									End If
						End If
						Wait(2)
						Set objSave=JavaDialog("Confirmation")
						  If sAction="Save" Then									
									objSave.SetTOProperty "title","Save WBS"
						  else
									objSave.SetTOProperty "title","Save All"	
						  End if															  
			
							If sInfo="Yes" Then
										objSave.JavaButton("Yes").Click micLeftBtn
										Fn_SchMgr_DefineWBSFormat = True
										Set objSave=Nothing
										Exit Function
							Elseif sInfo="No" Then
										objSave.JavaButton("No").Click micLeftBtn
										Fn_SchMgr_DefineWBSFormat = True
										Set objSave=Nothing
										Exit Function
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Button is to be clicked" )
							End If
		
							If sInfo <> "" Then
									sMessage =JavaDialog("Confirmation").JavaObject("MLabel").Object.getText()
									  If sInfo=sMessage Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified the confirmation message:"+sInfo)
												Fn_SchMgr_DefineWBSFormat = True
												objSave.Close
												Exit Function
					
									  Else
												Fn_SchMgr_DefineWBSFormat = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Confirnation Message:"+sInfo)
												objSave.Close
												Exit Function
									End If		
							End If  			

		Case "VerifyFormat&Example"

			If sSequence<>""  Then
				bFlag1=False
				sValue=objFormat.JavaEdit("Format").GetROProperty ("value")
				If lCase(sValue)=Lcase(sSequence) then
						bFlag1=True
				End If
			End If

				 If sLength<>"" Then
					 bFlag2=false
				objFormat.JavaStaticText("Static").SetTOProperty "label",sLength
						If objFormat.JavaStaticText("Static").Exist(2)=True Then
						bFlag2=True
						End If
				 End If

			If bFlag1 and bFlag2=True Then
				Fn_SchMgr_DefineWBSFormat = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the values i.e ["+sSequence+"] or ["+sLength+"] in the Format & Example Field are correct")
			Else
				Fn_SchMgr_DefineWBSFormat = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "One of the Two entered values i.e ["+sSequence+"] or ["+sLength+"] are wrong")
			End If

            Case "VerifySequence"
			If sSequence<>"" and  sInfo <>"" Then
			sValue=	JavaWindow("ScheduleManagerWindow").JavaWindow("Define WBS Format").JavaTable("Define").GetCellData(cint(sInfo),1)
				If lCase(sValue)=lcase(sSequence) Then
					Fn_SchMgr_DefineWBSFormat = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the sequense i.e ["+sSequence+"] at the index passed as ["+cstr(cint(sInfo))+" ]")
				Else
					Fn_SchMgr_DefineWBSFormat = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the sequense i.e ["+sSequence+"] at the index passed as ["+cstr(cint(sInfo))+" ]")
				End If


			End If

			Case "AddLevel"
				If  sInfo <>"" Then
					For iCount=0 to cInt(sInfo)-1
						objFormat.JavaButton("Add").Click micLeftBtn
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully added a new level and the level number is ["+cstr(sInfo)+" ]")
						Fn_SchMgr_DefineWBSFormat = True
					Next
				End if


	End Select

		'click on the required buttons
		If sButtons<>"" Then
			objFormat.JavaButton(sButtons).Click micLeftBtn
			If err.number<0 Then
				Fn_SchMgr_DefineWBSFormat = False
				
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  click on button "+sButtons)
					Set objTask = Nothing
					Exit Function
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on button "+sButtons)
					Fn_SchMgr_DefineWBSFormat = True
			End If
		End If
	End if

	Set objFormat=nothing
End Function

'Added by Prasanna 29-Apr-2011

Public Function Fn_WBS_OptionsSettings(sAction, sSeprator, sButton,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_WBS_OptionsSettings"
	Dim ObjDialog, aARelation, i, bFlag, itemsCount, iCounter, itemText,sItemName,sReturnValue

	Select Case sAction
	
		Case "DefineFormat"
			'Select menu [Edit  -> Options...]

			If Not Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Options").Exist(5) Then
					Call Fn_MenuOperation("Select","Edit:Options...")
			End If    
			Call Fn_ReadyStatusSync(1)

'			Set ObjDialog = Fn_UI_ObjectCreate("Fn_GeneralItem_OptionsSettings", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options"))
			Set ObjDialog = Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Options")
			'Code For ItemRevision Option
			If sSeprator="" Then
					'Select General : Item under OptionsTree.
					Call Fn_JavaTree_Select("Fn_GeneralUI_OptionsSettings", Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Options"), "OptionsTree","Options:WBS")	
			End If

			'Click on Define WBS format button
			Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Options").JavaButton("AddRelation").SetToProperty "label","Define WBS Format"
			Call Fn_Button_Click("Fn_GeneralItem_OptionsSettings",Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Options"),"AddRelation")
			wait 3
			If sButton <> "" Then
						Call Fn_Button_Click("Fn_GeneralItem_OptionsSettings",Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Options"),"OK")		
			End If
			Fn_WBS_OptionsSettings = true
			
		Case "Close"
			Call Fn_Button_Click("Fn_GeneralItem_OptionsSettings",Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Options"),"OK")     			
			Fn_WBS_OptionsSettings = true
	End Select	
End function



'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   :  Fn_SchMgr_CriticalPathColourOperations(sAction,sTaskPath,sInfo1,sInfo2)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will set,remove & verify the colour of the critical path applied
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''''/$$$$                                     2.) sTaskPath : Task or the Schedule path
'''''/$$$$							          3.) sInfo1 : For future Use
'''''/$$$$									 4.) sInfo2 : For future Use
'''''/$$$$
''''''/$$$$
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           	05/05/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			05/05/2011         1.0
'''''/$$$$
'''''/$$$$	How To Use :  bReturn=Fn_SchMgr_CriticalPathColourOperations("Menu","qwerty","","")
'''''/$$$$						   bReturn=Fn_SchMgr_CriticalPathColourOperations("VerifyCriticalPath","qwerty:t2","","")
'''''/$$$$						  bReturn=Fn_SchMgr_CriticalPathColourOperations("SetDefaultColour","","","")
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   :  Fn_SchMgr_CriticalPathColourOperations(sAction,sTaskPath,sInfo1,sInfo2)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will set,remove & verify the colour of the critical path applied
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''''/$$$$                                     2.) sTaskPath : Task or the Schedule path
'''''/$$$$							          3.) sInfo1 : For future Use
'''''/$$$$									 4.) sInfo2 : For future Use
'''''/$$$$
''''''/$$$$
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           	05/05/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			05/05/2011         1.0
'''''/$$$$
'''''/$$$$	How To Use :  bReturn=Fn_SchMgr_CriticalPathColourOperations("Menu","qwerty","","")
'''''/$$$$						   bReturn=Fn_SchMgr_CriticalPathColourOperations("VerifyCriticalPath","qwerty:t2","","")
'''''/$$$$						  bReturn=Fn_SchMgr_CriticalPathColourOperations("SetDefaultColour","","","")
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_SchMgr_CriticalPathColourOperations(sAction,sTaskPath,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_CriticalPathColourOperations"
   Dim objTable,iVal,iCount,sValue,bFlag,sRows,aValues,i,sIndex
   Fn_SchMgr_CriticalPathColourOperations=false

   Set objTable=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable")

   Select Case sAction

   Case "SetDefaultColour"

					   'invoke the set critical path colour dialog
					   bReturn = Fn_MenuOperation("Select","View:Critical Path:Set Color")
										If bReturn = False Then
											Fn_SchMgr_CriticalPathColourOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [View:Critical Path:Set Color]")
											Exit Function
									Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [View:Critical Path:Set Color]")
											  Fn_SchMgr_CriticalPathColourOperations=True
									End If
				
					   Set objSelectType = description.Create()
																	objSelectType("Class Name").value = "Static"
																	Set objSchTable =Dialog("Choose a color for Critical").ChildObjects(objSelectType)
																	
																	iCount=Dialog("Choose a color for Critical").ChildObjects(objSelectType).Count
				
					
					For iVal=0 to iCount-1
						If objSchTable(iVal).GetROProperty("attached text")="&Basic colors:" Then
							objSchTable(iVal).click 9,30,micLeftBtn 'the co-ordinates 9,30 are set to select "Red" as default colour and are tried on 4 different resolutions and are working properly
							Exit for												'however,if they do not work well,then this case should not be used
						End If
					Next
				
					'click on the ok button to set the colour
					Dialog("Choose a color for Critical").WinButton("OK").Click 0,0,micLeftBtn
					If err.number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the default colour")
						Dialog("Choose a color for Critical").Close
						Fn_SchMgr_CriticalPathColourOperations=false
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the default colour")
						Fn_SchMgr_CriticalPathColourOperations=True
					End If

	Case "Menu"

					 If sTaskPath<>"" Then
					
					  bReturn=Fn_SchMgr_SchTable_NodeOperation ( "Select", sTaskPath, "", "" , "" )
						   If bReturn = False Then
											Fn_SchMgr_CriticalPathColourOperations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select the task at path "+sTaskPath)
											Exit Function
						Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Select the task at path "+sTaskPath)
										Fn_SchMgr_CriticalPathColourOperations=True
						End If
				 Else
								Fn_SchMgr_CriticalPathColourOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The node name is blank for which the critical path has to be set or removed")
								Exit Function
			   End If
			
			   'now set the critical path
						 bReturn = Fn_MenuOperation("Select","View:Critical Path:View Critical Path")
								If bReturn = False Then
										Fn_SchMgr_CriticalPathColourOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [View:Critical Path:View Critical Path]")
										Exit Function
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [View:Critical Path:View Critical Path]")
										  Fn_SchMgr_CriticalPathColourOperations=True
								End If
								wait(3)

  CAse "VerifyCriticalPath"

		   If sTaskPath<>"" Then
			   If instr(1,sTaskPath,",")>0 Then
				   'select the taskpath specified
				   bReturn=Fn_SchMgr_SchTable_NodeOperation ( "MultiSelect", sTaskPath, "", "" , "" )
					   If bReturn = False Then
										Fn_SchMgr_CriticalPathColourOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  MultiSelect the tasks at path "+sTaskPath)
										Exit Function
					Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully MultiSelected the tasks at path "+sTaskPath)
									Fn_SchMgr_CriticalPathColourOperations=True
					End If
				Else
				  bReturn=Fn_SchMgr_SchTable_NodeOperation ( "Select", sTaskPath, "", "" , "" )
					   If bReturn = False Then
										Fn_SchMgr_CriticalPathColourOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select the task at path "+sTaskPath)
										Exit Function
					Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Select the task at path "+sTaskPath)
									Fn_SchMgr_CriticalPathColourOperations=True
					End If
			   End If
			Else
								Fn_SchMgr_CriticalPathColourOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The node name is blank for which the critical path has to be verified")
								Exit Function
	
		   End If
		   'verify if the colour changes when the critical path is set
		   	sIndex = Fn_SchMgr_TreeTableRowIndex(objTable, sTaskPath, "Object")
			If sIndex = False Then
				Fn_SchMgr_CriticalPathColourOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The node name is blank for which the critical path has to be verified")
								Exit Function
			Else
			bFlag=false
			sIndex = Right(sIndex, Len(sIndex) -1)			
			sValue= objTable.Object.getBackgroundColorForRow(cint(sIndex),False).toString
				aValues=split(sValue,"[",-1,1)
				For i=0 to ubound(aValues)
					If instr(1,aValues(i),"r=255,g=0,b=0")>0 Then
						bFlag=true
					End If
				Next
			End If



			If bFlag=true Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the colour for critical path set for the task at path "+sTaskPath)
				Fn_SchMgr_CriticalPathColourOperations=True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the colour for critical path set for the task at path "+sTaskPath)
				Fn_SchMgr_CriticalPathColourOperations=false
			End If
	End Select
set	objTable=nothing
		
End Function



'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   :  Fn_SchMgr_ConfirmLaunchWorkflow(sAction,sTaskPath,sButton,sInfo1,sInfo2)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will handle the confirm launch workflow dialog
'''''/$$$$
'''''/$$$$    PARAMETERS      :   1.) sAction : Action to be performed
'''''/$$$$                                     2.) sTaskPath : Task or the Schedule path
'''''/$$$$							          3.) sInfo1 : For future Use
'''''/$$$$									 4.) sInfo2 : For future Use
'''''/$$$$
''''''/$$$$
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           	05/05/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Prasanna			05/05/2011         1.0
'''''/$$$$
'''''/$$$$	How To Use :  bReturn=Fn_SchMgr_ConfirmLaunchWorkflow("HandleDialog","qwerty:t1","Yes","","")
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SchMgr_ConfirmLaunchWorkflow(sAction,sTaskPath,sButton,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ConfirmLaunchWorkflow"
   Dim objWorkflow,objError , ArrNodes
  set objWorkflow=JavaDialog("LaunchWorkflow")
   Fn_SchMgr_ConfirmLaunchWorkflow=false


   Select Case sAction

 	Case "HandleDialog"

		 If sTaskPath<>"" Then
					
                    ArrNodes = split(sTaskPath, ",",-1,1)
					If  Ubound(ArrNodes) > 0 Then
						bReturn=Fn_SchMgr_SchTable_NodeOperation ( "MultiSelect", sTaskPath, "", "" , "" )
					Else
					   bReturn=Fn_SchMgr_SchTable_NodeOperation ( "Select", sTaskPath, "", "" , "" )
					End If
						
						If bReturn = False Then
											Fn_SchMgr_ConfirmLaunchWorkflow = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select the task at path "+sTaskPath)
											Exit Function
						Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Select the task at path "+sTaskPath)
										Fn_SchMgr_ConfirmLaunchWorkflow=True
						End If
		End if

			'check the existence of the Confirm Launch Workflow Dialog & handle it if exists
		If objWorkflow.Exist(3) =False then
		 bReturn = Fn_MenuOperation("Select","Schedule:Launch Workflow Now")
		 	If bReturn = False Then
					Fn_SchMgr_ConfirmLaunchWorkflow = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Schedule:Launch Workflow Now]")
					Exit Function
			Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Schedule:Launch Workflow Now]")
					  Fn_SchMgr_ConfirmLaunchWorkflow=True
			End If
		End If

			If objWorkflow.Exist(3) =true then
				objWorkflow.JavaButton(sButton).Click micLeftBtn
				If err.number<0 Then
					Fn_SchMgr_ConfirmLaunchWorkflow = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to handle the Confirm Launch Workflow Dialog")
					JavaDialog("LaunchWorkflow").Close
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully handled the Confirm Launch Workflow Dialog")
					  Fn_SchMgr_ConfirmLaunchWorkflow=True
				End If
			End if
			
			'Added by Omkar to Handle the Launch Workflow Error Dialog
			Set objError=JavaWindow("ScheduleManagerWindow").JavaWindow("Error")
	    	objError.SetTOProperty "title","Launch Workflow Error"
			If objError.Exist(5)  then
						objError.JavaStaticText("Details").SetTOProperty "label",sInfo1
						If  objError.JavaStaticText("Details").Exist(5) Then
								objError.JavaButton("OK").Click micLeftBtn
								If Err.number<0 Then
										Fn_SchMgr_ConfirmLaunchWorkflow = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to handle the [Launch Workflow ] Error Window")
										JavaDialog("LaunchWorkflow").Close
										Set objError=Nothing	
										Exit Function
								Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully handled the [ Launch Workflow] Error Window")
										Fn_SchMgr_ConfirmLaunchWorkflow = false
								End If
						  else
										Fn_SchMgr_ConfirmLaunchWorkflow = true
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify error message")
										Exit Function
							end if 											
			End If			
  End Select
  Set objWorkflow=nothing
  Set objError=Nothing	
End Function


'*********************************************************  Function do Creation of Attribute *********************************************************************

'Function Name		:					 Fn_SchMgr_ErrorDialogMessageVerify(sAction,sWindCaption, sErrMsg )  

'Description			 :                  This fuction Verifies Error message and Details message
'																	
'Parameters			   :	 				 1. sAction
'														 2. sWindCaption
'														3.sErrMsg 

'Return Value		   : 			 True/False

'Examples				:			 Call  Fn_SchMgr_ErrorDialogMessageVerify("DetailsMessageWithoutClose" ,"Delete Rate Modifier Error", "Referenced by  (SchMgtCostFormStorage) Failed on Object Skill_2_Rate_03098 (BillRateImpl)The instance is referenced." )  
'												Call  Fn_SchMgr_ErrorDialogMessageVerify("Error_Message","Delete Rate Modifier Error", "A scheduling object of type 'BillRate' could not be deleted. Please see the server error log file for details." )  
												
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal Tanpure		09-May-2011      1.0											            	      Prasanna B.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Function Fn_SchMgr_ErrorDialogMessageVerify(sAction,sWindCaption, sErrMsg )  

		Dim dicErrorInfo, bReturn 
		Set dicErrorInfo = CreateObject("Scripting.Dictionary")
		With dicErrorInfo 	
		 .Add "Title", 	 sWindCaption
		 .Add "Message", sErrMsg		 
		 .Add "Action", sAction		 
		 .Add "Button", "OK"		 
		End with
		Fn_SchMgr_ErrorDialogMessageVerify = Fn_SISW_SchMgr_ErrorVerify(dicErrorInfo)

End Function

'*********************************************************  Function to Handle Information Dialog *********************************************************************

'Function Name		:					 Fn_SchMgr_InformationDialogHandle(sTitle,sMsg,sButton,sDetails)  

'Description			 :                  This fuction Handles Information Dialog
'																	
'Parameters			   :	 				 1. sTitle   ' Dialog Title
'											 2. sMsg     ' Message to verify
'											 3. sButton  ' Valid Button Name
'											 4. sDetails

'Return Value		   : 			 True/False

'Examples				:			 Call Fn_SchMgr_InformationDialogHandle("Information","Workflow is triggered, do you still want to change ""Status""?","Yes","")
												
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sachin Joshi		16-June-2011            1.0											Prasanna B.
'										Sushma Pagare		20-June-2013            1.0									
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SchMgr_InformationDialogHandle(sTitle,sMessage,sButton,sDetails)

		Dim dicErrorInfo, bReturn 
		Set dicErrorInfo = CreateObject("Scripting.Dictionary")
		With dicErrorInfo 
		 .Add "Title", sTitle
		 .Add "Message", sMessage		 
		 .Add "Button", sButton
		 .Add "Action", "InformationDialogHandle"
		End with
		Fn_SchMgr_InformationDialogHandle = Fn_SISW_SchMgr_ErrorVerify(dicErrorInfo)

End Function



''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_SchMgr_PrintTablePropertyVerify(sAction,sMenu,aColumn,aProperty,sExtra,bClose)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will verify the schedule details in "HTML / Text"  or  "Graphics" View
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sMenu : Valid menu name either HTML / Text or Graphics
''''/$$$$										aColumn : Array of Column names to be verified
''''/$$$$										aProperty : Array of property names to be verified
''''/$$$$										sExtra : For Future Use
''''/$$$$										bClose : Boolean parameter to close the dialog
''''/$$$$	
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          05/09/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Prasanna			05/09/2011         1.0
''''/$$$$
''''/$$$$		How To Use :   aValue=Array("Work Estimate","Work Complete Percent")
''''/$$$$									  bReturn=Fn_SchMgr_PrintTablePropertyVerify("VerifyHTMLColumnName","HTML / Text",aValue,"","","True")					
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_SchMgr_PrintTablePropertyVerify(sAction,sMenu,aColumn,aProperty,sExtra,bClose)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_PrintTablePropertyVerify"
   Fn_SchMgr_PrintTablePropertyVerify=false

   Dim sValue,objScheduleTable,iCounter

   If sMenu="HTML / Text" Then
	   set objScheduleTable= JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Print")
	Else
		'can be used to set for "Graphics view i.e "Page Setup Dialog
		''Page Setup  dialog is not added in the OR
   End If 


	'open the desired  menu

	JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaTable("SchTaskTable").SelectColumnHeader "Object", "RIGHT"
	JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaMenu("label:=Print Table","index:=0").JavaMenu("label:="&sMenu,"index:=0").Select
	wait 5

	If  objScheduleTable.Exist Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The object [" &objScheduleTable.toString()&"] does Exist")

				Select Case sAction
				
						Case "VerifyHTMLColumnName"

				
							 If IsArray(aColumn) Then
				
										 For iCounter=0 to uBound(aColumn)
											 sValue=objScheduleTable.JavaEdit("ScheduleHTMLDetails").GetROProperty ("value")
											 wait 5
											 If instr(1,sValue,aColumn(iCounter))>1 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verifed that the column ["&aColumn(iCounter)&"] is present")
												 Fn_SchMgr_PrintTablePropertyVerify=true
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Failed to verify that the column ["&aColumn(iCounter)&"] is present.")
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Check Column Name / Names again")
												Fn_SchMgr_PrintTablePropertyVerify=false
												Exit function
											 End If
									 Next
						Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Array Should be passed as a parameter to verify column names")
										Fn_SchMgr_PrintTablePropertyVerify=false
										Exit function
						End If
				End Select
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The object " &objScheduleTable.toString()&" does not Exist")
		Fn_SchMgr_PrintTablePropertyVerify=false
		Exit Function
	End if 

		If cBool(bClose)=true Then
			objScheduleTable.Close
		End If

End Function


''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_SchMgr_ProgramViewCreate(sAction,sName,sDesc,sImport,aColumns,bOpenOnCreate,aSchedules,aFilter,aGroupAttributes,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create a program view in simple and detail cases
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sName : Program View Name
''''/$$$$										sDesc : Program View Name Description
''''/$$$$										sImport : Value for the Import File
''''/$$$$										aColumns : Array Of Columns To Be Added
''''/$$$$										bOpenOnCreate : to open or to not open Program View After Creation
''''/$$$$										aSchedules : Array Of Schedules to be Selected And Added
''''/$$$$										aFilter : Array of Filters to be added
''''/$$$$										aGroupAttributes : Array of Group attributes
''''/$$$$										sInfo1 : For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          30/11/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			30/11/2011            1.0
''''/$$$$
'''/$$$$    Modified By        :     Sneha              21/12/2011            2.0
'''/$$$$
'''/$$$$    Changes         :  Added code for column chooser
'''/$$$$
''''/$$$$		How To Use :   
''''/$$$$							aValues=array("Qwerty","Schedule11164")
''''/$$$$						bReturn=Fn_SchMgr_ProgramViewCreate("SimpleCreate","ShreyasView","Test View","","","false",aValues,"","","","")
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_SchMgr_ProgramViewCreate(sAction,sName,sDesc,sImport,aColumns,bOpenOnCreate,aSchedules,aFilter,aGroupAttributes,sInfo1,sInfo2)
		GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ProgramViewCreate"
		Dim objProgram,iCount,bReturn,objNotFound,objSave,objColChooser,iCounter,aValues
		Fn_SchMgr_ProgramViewCreate=false1
				bReturn= Fn_MenuOperation("Select","File:New:Program View")
				Call Fn_ReadyStatusSync(5)
					If bReturn=True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "NewProgramView dialog is Now displayed.")
							Fn_SchMgr_ProgramViewCreate=True
							Call Fn_ReadyStatusSync(5)
							wait 3
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "NewProgramView dialog is still not displayed.")
							Fn_SchMgr_ProgramViewCreate=False
							Exit Function
					End If
'		Set objProgram=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewProgramView")
	    Set objNotFound=Fn_SISW_GetObject("Object Not Found")
        Set objSave=JavaWindow("ScheduleManagerWindow").JavaWindow("Save")
			If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewProgramView").Exist(5) Then
						Set objProgram=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewProgramView")
   			Elseif  Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("NewProgramView").Exist(5)  Then
			          Set objProgram= Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("NewProgramView")
			Else
			       	Fn_SchMgr_ProgramViewCreate=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Hierarchy for New Program View has been changed")
					Exit function
			End If

			Select Case sAction

								Case "SimpleCreate"

									If sName<>"" Then
											objProgram.JavaEdit("Name").Set sName
											If err.number<0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Value ["+sName+"]")
												Fn_SchMgr_ProgramViewCreate=False
												Exit Function
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully to set the Value ["+sName+"]")
												Fn_SchMgr_ProgramViewCreate=True	
											End If
								End If

									If sDesc<>"" Then
											objProgram.JavaEdit("Description").Set sDesc
											If err.number<0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Value ["+sDesc+"]")
												Fn_SchMgr_ProgramViewCreate=False
												Exit Function
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully to set the Value ["+sDesc+"]")
												Fn_SchMgr_ProgramViewCreate=True	
											End If
								End If

									If sImport<>"" Then
										'will be coded as required
									End if

									'To select Columns from the Cloumn chooser Dialog
									If IsArray(aColumns) Then
                                        Set objColChooser=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Column Chooser")
										If objColChooser.Exist(5)=False Then
											JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("NewProgramView").JavaButton("ChooseColumns").Click micLeftBtn
											If Err.Number < 0 Then
														Fn_SchMgr_ProgramViewCreate = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to Open the Coloumn chooser Dialog.")
														objColChooser.JavaButton("Cancel").Click
														Set objColChooser = Nothing
														Exit Function
											Else
													Fn_SchMgr_ProgramViewCreate = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Opened the Coloumn chooser Dialog.")
'													objColChooser.JavaButton("Close").Click
'													Exit Function
											End If
										End If
										If sInfo2<>"" Then
												If instr(1,sInfo2,":") Then			
													sInfo2 = split(sInfo2,":",-1,1)
													objColChooser.JavaTree("Type").Select("Types:"+sInfo2(0))
													If Err.Number < 0 Then
														Fn_ProgViewColumnOperations = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Column Type :Types:ScheduleTask. from Column Chooser")
														objColChooser.JavaButton("Close").Click
														Exit Function
													Else
														Fn_ProgViewColumnOperations = True
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected the Column Type:Types:ScheduleTask Column Chooser")
													End If
											End If
									End If
										For iCounter = 0 to Ubound(aColumns)
												bReturn = Fn_SchMgr_TableColIndex(objTable,aColumns(iCounter))
												If cBool(bReturn) = False Then
													'Select column from available columns.
													Wait 3
													objColChooser.JavaList("AvailableColumns").ExtendSelect aColumns(iCounter)
													If Err.Number < 0 Then
														Fn_SchMgr_ProgramViewCreate = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to select  " &aColName(iCounter)& " from available columns.")
														objColChooser.JavaButton("Close").Click
														Set objColChooser = Nothing
														Exit Function
													End If
											Else
													Fn_SchMgr_ProgramViewCreate = True
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Already  "  & aColumns(iCounter) & "column exists in schedule table.")
													objColChooser.JavaButton("Close").Click
													Exit Function
											End If
										Next
									
										'Click Add button
										objColChooser.JavaButton("AddCol").WaitProperty "enabled",1,20000
										objColChooser.JavaButton("AddCol").Click

										Wait 3				
										If objColChooser.JavaButton("OK").Exist(5) Then
											objColChooser.JavaButton("OK").WaitProperty "enabled",1,20000
											objColChooser.JavaButton("OK").Click micLeftBtn
										'Click Apply button
										Elseif objColChooser.JavaButton("Apply").Exist(5) Then
											objColChooser.JavaButton("Apply").WaitProperty "enabled",1,20000
											objColChooser.JavaButton("Apply").Click micLeftBtn
										End If
									
										Fn_SchMgr_ProgramViewCreate = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully added columns to schedule table .")
										If Err.Number < 0 Then
											Fn_SchMgr_ProgramViewCreate = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Faliled to add columns to schedule table..")
											objColChooser.JavaButton("Close").Click
											Set objColChooser = Nothing
											Exit Function
										End If
								End If

								If cbool(bOpenOnCreate)=true Then
										objProgram.JavaCheckBox("OpenOnCreate").Set "ON"
										If err.number<0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Set the Checkbox  [ OpenOnCreate ] to ON")
											Fn_SchMgr_ProgramViewCreate=False
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Set the Checkbox  [ OpenOnCreate ] to ON")
											Fn_SchMgr_ProgramViewCreate=True	
										End If
								Elseif cbool(bOpenOnCreate)=false then
										objProgram.JavaCheckBox("OpenOnCreate").Set "OFF"
										If err.number<0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Set the Checkbox  [ OpenOnCreate ] to OFF")
											Fn_SchMgr_ProgramViewCreate=False
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Set the Checkbox  [ OpenOnCreate ] to OFF")
											Fn_SchMgr_ProgramViewCreate=True	
										End If
							End If

							'Click on Next button
							If IsArray(sInfo2) Then
										If sInfo2(1)="True" Then
												objProgram.JavaButton("Finish").Click micLeftBtn
												If err.number<0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Finish Button and Create Program View ["+sName+"] from the New Program View Wizard")
													Fn_SchMgr_ProgramViewCreate=False
													Exit Function
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Finish Button and Created Program View ["+sName+"] from the New Program View Wizard")
													Fn_SchMgr_ProgramViewCreate=True	
													Exit Function
												End If	
										End If
							Else
								objProgram.JavaButton("Next").Click micLeftBtn
								If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to go to the next menu of the New Program View Dialog")
									Fn_SchMgr_ProgramViewCreate=False
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully activated the next menu of the New Program View Dialog")
									Fn_SchMgr_ProgramViewCreate=True
									Call Fn_ReadyStatusSync(5)
								End If
							End If

							If IsArray(sInfo2) Then
											If sInfo2(1)="True" Then
												objProgram.JavaButton("Finish").Click micLeftBtn
												If err.number<0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Finish Button and Create Program View ["+sName+"] from the New Program View Wizard")
													Fn_SchMgr_ProgramViewCreate=False
													Exit Function
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Finish Button and Created Program View ["+sName+"] from the New Program View Wizard")
													Fn_SchMgr_ProgramViewCreate=True	
													Exit Function
												End If	
											End If
							Else
											If objProgram.JavaButton("loadall").Exist(5) Then
													If objProgram.JavaButton("loadall").GetRoProperty("enabled")=1 Then objProgram.JavaButton("loadall").Click micLeftBtn
													Wait 3
											End If
							End If

									'choose the Required Schedules
							If IsArray(sInfo2) Then
											If sInfo2(1)="True" Then
												objProgram.JavaButton("Finish").Click micLeftBtn
												If err.number<0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Finish Button andCreate Program View ["+sName+"] from the New Program View Wizard")
													Fn_SchMgr_ProgramViewCreate=False
													Exit Function
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Finish Button and Created Program View ["+sName+"] from the New Program View Wizard")
													Fn_SchMgr_ProgramViewCreate=True	
													Exit Function
												End If	
											End If
							Else
									If  isarray(aSchedules)=true Then
										For iCount=0 to uBound(aSchedules)
												objProgram.JavaButton("Clear").Click micLeftBtn
												objProgram.JavaEdit("FindText").Set aSchedules(iCount)
												objProgram.JavaButton("Find").Click micLeftBtn
												wait 3
												If objNotFound.Exist(3) Then
													objNotFound.JavaButton("OK").Click micLeftBtn
													objProgram.JavaButton("Close").Click micLeftBtn
													Fn_SchMgr_ProgramViewCreate=False
													Exit Function
													Exit for
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Schedule ["+aSchedules(iCount)+"] was not found in the Available Schedules List")
													Fn_SchMgr_ProgramViewCreate=False
													Exit Function
												Else
													objProgram.JavaButton("Add").Click micLeftBtn
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Schedule ["+aSchedules(iCount)+"] was Successfully found in the Available Schedules List and added to the Selected Schedules List")
												End if
										Next
									End If
							End If


							'Click on Finish Button to Create the Program View
							If sInfo2 <>"" Then
								If IsArray(sInfo2) Then
									If sInfo2(1)="True" Then
										objProgram.JavaButton("Finish").Click micLeftBtn
										If err.number<0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Create Program View ["+sName+"] from the New Program View Wizard")
											Fn_SchMgr_ProgramViewCreate=False
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Program View ["+sName+"] from the New Program View Wizard")
											Fn_SchMgr_ProgramViewCreate=True	
										End If
									End If
								End If
							Else
								objProgram.JavaButton("Finish").Click micLeftBtn
								If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Create Program View ["+sName+"] from the New Program View Wizard")
									Fn_SchMgr_ProgramViewCreate=False
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Program View ["+sName+"] from the New Program View Wizard")
									Fn_SchMgr_ProgramViewCreate=True	
								End If
							End If
							'Check the Existence of save as window and handle it if required

							if sInfo1<>"" then
									If objSave.Exist Then
											objSave.JavaButton(sInfo1).Click micLeftBtn
											If err.number<0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Save Program View ["+sName+"] from the New Program View Wizard")
												Fn_SchMgr_ProgramViewCreate=False
												Exit Function
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Saved Program View ["+sName+"] from the New Program View Wizard")
												Fn_SchMgr_ProgramViewCreate=True	
											End If
									End If
							End if

			End Select
		Set objProgram=nothing
		Set objNotFound=nothing
		Set objSave=nothing
End Function


''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_SchMgr_CreateCrossScheduleDependency(sAction,sNode,sSchedule,sTask,sProxyType,sDependencyType,sLag,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create a program view in simple and detail cases
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sNode : Valid Schedule Table Node
''''/$$$$										sSchedule : Schedule to be selected From Available Schedules List
''''/$$$$										sTask : Task to be selected From Available Schedules List
''''/$$$$										sProxyType : Valid proxy Task Type (successor  / Predessor)
''''/$$$$										sDependencyType : Valid Dependency Type
''''/$$$$										sLag : Valid Lag
''''/$$$$										sInfo1 : For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          30/11/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			30/11/2011            1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SchMgr_CreateCrossScheduleDependency("CreateDependency","Schedule2:t4","Qwerty1","t1","Successor","Finish-to-Start","2","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SchMgr_CreateCrossScheduleDependency(sAction,sNode,sSchedule,sTask,sProxyType,sDependencyType,sLag,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_CreateCrossScheduleDependency"
Dim objCrossDependency,sValue,bReturn,sIndex,iCount
Set objCrossDependency=JavaWindow("ScheduleManagerWindow").JavaWindow("CreateCrossScheduleDependency")

Fn_SchMgr_CreateCrossScheduleDependency=false

'Select the Node from the Schedule table
			bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sNode,"","","")

			If  bReturn <> False Then
				Fn_SchMgr_CreateCrossScheduleDependency=true
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sNode)
			ELse
				Fn_SchMgr_CreateCrossScheduleDependency=false
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sNode)
				Exit Function
			End If

			'check if the Cross Schedule DEpendency Dialog Exists
			If objCrossDependency.Exist(5) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "CrossScheduleDependency dialog is displayed.")
			Else
					bReturn= Fn_MenuOperation("Select","Schedule:Link:Create Cross Schedule Dependency")
					If bReturn=True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "CrossScheduleDependency dialog is Now displayed.")
							Fn_SchMgr_CreateCrossScheduleDependency=True
							Call Fn_ReadyStatusSync(5)
							wait 3
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "CrossScheduleDependency dialog is still not displayed.")
							Fn_SchMgr_CreateCrossScheduleDependency=False
							Exit Function
					End If
			End If

			Select Case sAction
			

						Case "CreateDependency"

								If sInfo2<>"" Then
									If sInfo2="Template" Then
										objCrossDependency.JavaWindow("SelectSchedule").JavaTree("Available Schedules").SetTOProperty "attached text", "Available Templates"
									End If
								Else
										objCrossDependency.JavaWindow("SelectSchedule").JavaTree("Available Schedules").SetTOProperty "attached text", "Available Schedules"
								End If

						'Load All the schedules before selecting the required one.
						If objCrossDependency.JavaButton("LoadAll").Exist(5) Then
							If objCrossDependency.JavaButton("LoadAll").GetROProperty("enabled") = 1 then objCrossDependency.JavaButton("LoadAll").Click
						End If
				        Call Fn_ReadyStatusSync(10)
						
						'Select the Schedule from the Available Schedules List
						objCrossDependency.JavaWindow("SelectSchedule").JavaTree("Available Schedules").Select "#0:"+sSchedule
						If err.number<0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the Schedule ["+sSchedule+"] from the Available Schedules List")
							Fn_SchMgr_CreateCrossScheduleDependency=False
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the Schedule ["+sSchedule+"] from the Available Schedules List")
							Fn_SchMgr_CreateCrossScheduleDependency=True	
						End If

						'Click on next button
						objCrossDependency.JavaButton("Next").Click micLeftBtn
						If objCrossDependency.JavaWindow("SelectTask").Exist(3)=false Then
							wait 2
						End If

						'Select the Task from the Task list

                        objCrossDependency.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").Object.expandAll
'						wait 2
'					sIndex = Fn_SchMgr_TreeTableRowIndex(objCrossDependency.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable"), sTask, "Object")
'					If sIndex  <> FALSE Then
'						sIndex = "#" + cstr(sIndex)
'
'					'Select the Expected  scheduleTable Node
'					 objCrossDependency.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").SelectRow sIndex
'							 If Err.Number <  0 Then
'								 Fn_SchMgr_CreateCrossScheduleDependency = FALSE				 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the task at path ["+sTask+"] from the Available Tasks List")
'							Else
'								Fn_SchMgr_CreateCrossScheduleDependency = TRUE				 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the task at path ["+sTask+"] from the Available Tasks List")	
'							End If
'					Else
'							Fn_SchMgr_CreateCrossScheduleDependency = FALSE				 
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task does not Exist at path ["+sTask+"] in the Availaible tasks list")	
'							Exit Function
'				End If

						iRows=objCrossDependency.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").GetROProperty("rows")
						For iCount=0 to iRows-1
							sValue=objCrossDependency.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").Object.getValueAt(iCount,0).toString()
							If lCase(sValue)=lCase(sTask) Then
								objCrossDependency.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").SelectRow iCount
								Fn_SchMgr_CreateCrossScheduleDependency = TRUE				 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the task at path ["+sTask+"] from the Available Tasks List")	
								Exit For
							End If
						Next
						
						If cint(iRows)=cint(iCount) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task does not Exist at path ["+sTask+"] in the Availaible tasks list")	
												Fn_SchMgr_CreateCrossScheduleDependency=False
												Exit Function
						End If

						'Click on next button
						objCrossDependency.JavaButton("Next").Click micLeftBtn

						'Set the Values For Proxy Task

						If sProxyType<>"" Then
								objCrossDependency.JavaStaticText("Dependency_Label").SetTOProperty "label","The Proxy Task.*"
								objCrossDependency.JavaList("ProxyTask").Select sProxyType
								If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Value ["+sProxyType+"] in the Proxy Task List")
									Fn_SchMgr_CreateCrossScheduleDependency=False
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully to set the Value ["+sProxyType+"] in the Proxy Task List")
									Fn_SchMgr_CreateCrossScheduleDependency=True	
								End If
					End If

						'Set the Values For Dependency

						If sDependencyType<>"" Then
							objCrossDependency.JavaStaticText("Dependency_Label").SetTOProperty "label","Dependency Type:"
								objCrossDependency.JavaList("DependencyType").Select sDependencyType
								If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Value ["+sDependencyType+"] in the DependencyType List")
									Fn_SchMgr_CreateCrossScheduleDependency=False
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully to set the Value ["+sDependencyType+"] in the DependencyType List")
									Fn_SchMgr_CreateCrossScheduleDependency=True	
								End If
					End If

						'Set the Values For Lag

						If sLag<>"" Then

							If cInt(sLag)>0 Then
									For iCount=0 to cint(sLag)-1
											objCrossDependency.JavaSpin("Lag").next
									Next
							Else
									For iCount=0 to cint(sLag)-1
											objCrossDependency.JavaSpin("Lag").Prev
									Next							
							End If

								If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Value ["+sLag+"] for Lag")
									Fn_SchMgr_CreateCrossScheduleDependency=False
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully to set the Value ["+sLag+"] for Lag")
									Fn_SchMgr_CreateCrossScheduleDependency=True	
								End If
					End If

					'Click on Finish Button
					objCrossDependency.JavaButton("Finish").Click micLeftBtn

					If err.number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Create Cross Schedule Dependency From tasks at path ["+sNode+"] to task at path ["+sTask+"] from the Create Cross Schedule Dependency Wizard")
						Fn_SchMgr_CreateCrossScheduleDependency=False
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Cross Schedule Dependency From tasks at path ["+sNode+"] to task at path ["+sTask+"] from the Create Cross Schedule Dependency Wizard")
						Fn_SchMgr_CreateCrossScheduleDependency=True	
					End If

			End Select
		Set objCrossDependency=nothing
End Function



''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_SchMgr_CreateProxyTask(sAction,sNode,sSchedule,sTask,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create a proxy task
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sNode : Valid Schedule Table Node
''''/$$$$										sSchedule : Schedule to be selected From Available Schedules List
''''/$$$$										sTask : Task to be selected From Available Schedules List
''''/$$$$										sInfo1 : For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          30/11/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			30/11/2011            1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SchMgr_CreateProxyTask("CreateProxyTask","Schedule2","Qwerty1","t1","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public function Fn_SchMgr_CreateProxyTask(sAction,sNode,sSchedule,sTask,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_CreateProxyTask"
Dim objProxy,sValue,bReturn,sIndex,iCount,iRows
Set objProxy=JavaWindow("ScheduleManagerWindow").JavaWindow("NewProxyTask")

Fn_SchMgr_CreateProxyTask=false

'Select the Node from the Schedule table
			bReturn = Fn_SchMgr_SchTable_NodeOperation("MultiSelect",sNode,"","","")

			If  bReturn <> False Then
				Fn_SchMgr_CreateProxyTask=true
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected tasks" + sNode)
			ELse
				Fn_SchMgr_CreateProxyTask=false
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  task" + sNode)
				Exit Function
			End If

			'check if the Cross Schedule DEpendency Dialog Exists
			If objProxy.Exist(5) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "NewProxyTask dialog is displayed.")
			Else
					bReturn= Fn_MenuOperation("Select","File:New:Proxy Task...")
					If bReturn=True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "NewProxyTask dialog is Now displayed.")
							Fn_SchMgr_CreateProxyTask=True
							Call Fn_ReadyStatusSync(5)
'							wait 3
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "NewProxyTask dialog is still not displayed.")
							Fn_SchMgr_CreateProxyTask=False
							Exit Function
					End If
			End If

			Select Case sAction
			

						Case "CreateProxyTask"


						If sInfo2<>"" Then
								If sInfo2="Template" Then
									objProxy.JavaWindow("SelectSchedule").JavaTree("AvailableSchedules").SetTOProperty "attached text", "Available Templates"
								End If
						Else
								objProxy.JavaWindow("SelectSchedule").JavaTree("AvailableSchedules").SetTOProperty "attached text", "Available Schedules"
						End If

						'Select the Schedule from the Available Schedules List
						If objProxy.JavaWindow("SelectSchedule").JavaButton("LoadAll").Exist(5) Then							
								If objProxy.JavaWindow("SelectSchedule").JavaButton("LoadAll").GetROProperty ("enabled")=1 Then objProxy.JavaWindow("SelectSchedule").JavaButton("LoadAll").Click micLeftBtn
								Call Fn_ReadyStatusSync(5)
								wait 3
						End If

						objProxy.JavaWindow("SelectSchedule").JavaTree("AvailableSchedules").Select "#0:"+sSchedule
						If err.number<0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the Schedule ["+sSchedule+"] from the Available Schedules List")
							Fn_SchMgr_CreateProxyTask=False
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the Schedule ["+sSchedule+"] from the Available Schedules List")
							Fn_SchMgr_CreateProxyTask=True	
						End If

						'Click on next button
						objProxy.JavaButton("Next").Click micLeftBtn
						If objProxy.JavaWindow("SelectTask").Exist(3)=false Then
							wait 1
						End If

						'Select the Task from the Task list

						objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").Object.expandAll
		

'						sValue= objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").GetColumnName (0)
'					sIndex = Fn_SchMgr_TreeTableRowIndex(objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable"), sTask, sValue)
'					If sIndex  <> FALSE Then
'						sIndex = "#" + cstr(sIndex)
'
'					'Select the Expected  scheduleTable Node
'					 objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").SelectRow sIndex
'							 If Err.Number <  0 Then
'								 Fn_SchMgr_CreateCrossScheduleDependency = FALSE				 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the task at path ["+sTask+"] from the Available Tasks List")
'							Else
'								Fn_SchMgr_CreateProxyTask = TRUE				 
'								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the task at path ["+sTask+"] from the Available Tasks List")	
'							End If
'					Else
'							Fn_SchMgr_CreateProxyTask = FALSE				 
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task does not Exist at path ["+sTask+"] in the Availaible tasks list")	
'							Exit Function
'				End If

						iRows=objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").GetROProperty("rows")
						For iCount=0 to iRows-1
							sValue=objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").Object.getValueAt(iCount,0).toString()
							If lCase(sValue)=lCase(sTask) Then
								objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").SelectRow iCount
								Fn_SchMgr_CreateProxyTask = TRUE				 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the task at path ["+sTask+"] from the Available Tasks List")	
								Exit For
							End If
						Next
						
						If cint(iRows)=cint(iCount) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task does not Exist at path ["+sTask+"] in the Availaible tasks list")	
												Fn_SchMgr_CreateProxyTask=False
												Exit Function
						End If


					'Click on Finish Button
					objProxy.JavaButton("Finish").Click micLeftBtn

					If err.number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Create Proxy Task at path ["+sTask+"] from the New Proxy Task Wizard")
						Fn_SchMgr_CreateProxyTask=False
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Proxy Task at path ["+sTask+"] from the New Proxy Task Wizard")
						Fn_SchMgr_CreateProxyTask=True	
					End If

					Case "VerifyTaskExists&Create"


						'Select the Schedule from the Available Schedules List
						If objProxy.JavaWindow("SelectSchedule").JavaButton("LoadAll").Exist(5) Then
							If objProxy.JavaWindow("SelectSchedule").JavaButton("LoadAll").GetROProperty ("enabled")=1 Then objProxy.JavaWindow("SelectSchedule").JavaButton("LoadAll").Click micLeftBtn
							Call Fn_ReadyStatusSync(5)
							wait 1
						End If
						

						objProxy.JavaWindow("SelectSchedule").JavaTree("AvailableSchedules").Select "#0:"+sSchedule
						If err.number<0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select the Schedule ["+sSchedule+"] from the Available Schedules List")
							Fn_SchMgr_CreateProxyTask=False
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected the Schedule ["+sSchedule+"] from the Available Schedules List")
							Fn_SchMgr_CreateProxyTask=True	
						End If

						'Click on next button
						objProxy.JavaButton("Next").Click micLeftBtn
						If objProxy.JavaWindow("SelectTask").Exist(3)=false Then
							wait 1
						End If

						'Select the Task from the Task list

						objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").Object.expandAll
		
						iRows=objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").GetROProperty("rows")
						For iCount=0 to iRows-1
							sValue=objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").Object.getValueAt(iCount,0).toString()
							If lCase(sValue)=lCase(sTask) Then
								objProxy.JavaWindow("SelectTask").JavaTable("ScheduleTreeTable").SelectRow iCount
								Fn_SchMgr_CreateProxyTask = TRUE				 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified the task at path ["+sTask+"] from the Available Tasks List")	
								Exit For
							End If
						Next
						
						If cint(iRows)=cint(iCount) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Task does not Exist at path ["+sTask+"] in the Availaible tasks list")	
												Fn_SchMgr_CreateProxyTask=False
												Exit Function
						End If


					'Click on Finish Button
					objProxy.JavaButton(sInfo1).Click micLeftBtn
					If err.number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click the Button ["+sInfo1+"]")
						Fn_SchMgr_CreateProxyTask=False
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked the Button ["+sInfo1+"]")
						Fn_SchMgr_CreateProxyTask=True	
					End If

			End Select
		Set objProxy=nothing
End Function



''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_SchMgr_CustomizeGroupFilters(sField,sRange,sOrder,sRollupIndex,sRollup,sButton,sExtra1,sExtra2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create a proxy task
''''/$$$$
''''/$$$$  PARAMETERS   : 		sField : Valid Index and Value to be set in the Field Area
''''/$$$$											sRange : Valid Numeric or Date Range  (Will be coded as required)
''''/$$$$											sOrder : Valid Index and Value to be set in the Order Area
''''/$$$$											sRollupIndex : Valid Index and Button Name of the Controls of the Customize Rollup Dialog
''''/$$$$											sRollup=Valid Value to be set in the Field Name and Rollup Condition
''''/$$$$											sButton=Valid button name to be clicked on the Customize Groups By Dialog
''''/$$$$											sExtra1 : For Future Use
''''/$$$$											sExtra2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          06/12/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			   06/12/2011           1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SchMgr_CustomizeGroupFilters("2:Status","","3:Ascending","0:1:Done","State:count","Done","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_SchMgr_CustomizeGroupFilters(sField,sRange,sOrder,sRollupIndex,sRollup,sButton,sExtra1,sExtra2)

GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_CustomizeGroupFilters"
Fn_SchMgr_CustomizeGroupFilters=false
	Dim objGroup,objRollup,aValues,aIndex,iCount,sValues,objButton,sTemplateType,aRollupValues, objGroup1

	Set objGroup=JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Customize Group By")
	Set objGroup1=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Customize Group By")
	Set objRollup=JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("CustomizeRollups")

'Check the Existence of the Customize Groups For dialog

			'check if the Cross Schedule DEpendency Dialog Exists
			If objGroup.Exist(5) Or objGroup1.Exist(2) Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Customize Group By dialog is displayed.")
			Else
					bReturn = Fn_MenuOperation("Select","Program:Group Attributes")
					If bReturn=True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Customize Group By dialog is Now displayed.")
							Fn_SchMgr_CustomizeGroupFilters=True
							Call Fn_ReadyStatusSync(5)
							wait 3
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Customize Group By dialog is still not displayed.")
							Fn_SchMgr_CustomizeGroupFilters=False
							Exit Function
					End If
			End If
			
			If objGroup.exist(5) Then
				Set objGroup=JavaWindow("ScheduleManagerWindow").JavaWindow("SchMgrWindow").JavaDialog("Customize Group By")
			elseIf objGroup1.exist(2) Then
				Set objGroup=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Customize Group By")
			End If


'Set the Desired Value for the Fields Area
	
	If sField<>"" Then
		aValues=split(sField,":",-1,1)
			Set objButton=Description.Create()
					objButton("Class Name").value = "JavaButton"
					set sObjects=objGroup.ChildObjects(objButton)
					sValue= sObjects.Count
					For iCount=0 to cint(sValue)-1 step 1
							If cint(aValues(0))=iCount Then
								objGroup.JavaButton("DropDown").SetTOProperty "Index",iCount
								objGroup.JavaButton("DropDown").Click micLeftBtn
								If aValues(1) = "BLANK" Then
									wait 2
									bReturn = Fn_iComboSet(objGroup," ")
'									bReturn = Fn_UI_JavaObject_Click("Fn_SchMgr_FilterSettings", objGroup.JavaWindow("JavaWindow"), "BlankListItem", 2, 2, "LEFT")
									If bReturn = False Then
										Fn_SchMgr_CustomizeGroupFilters=false
										exit function
									Else
										Fn_SchMgr_CustomizeGroupFilters=True
									End If
								Else
									Set sTemplateType=Description.Create()
									sTemplateType("Class Name").value = "JavaStaticText"
									Set  intNoOfObjects =objGroup.ChildObjects(sTemplateType)
									  For i = 0 to intNoOfObjects.count-1
										   If  intNoOfObjects(i).getROProperty("label") = aValues(1) Then
												intNoOfObjects(i).Click 1,1
												If objGroup.JavaDialog("Information").Exist(5)=true Then
													objGroup.JavaDialog("Information").JavaButton("OK").Click micLeftBtn
													Fn_SchMgr_CustomizeGroupFilters=false
													exit function
												Else
													Fn_SchMgr_CustomizeGroupFilters=True
													Exit for
												End If
												bFlag = True
												Exit for
										   End If
									  Next
								End If
								Exit for
						End If
					Next
	End If

	If sRange<>"" Then
		'For further Use
	End If
	
'Set the Desired Value for the Order Area
	If sOrder<>"" Then
		aValues=split(sOrder,":",-1,1)

				For iCount=1 to cint(sValue) step 1
							If cInt(aValues(0))=iCount Then
								objGroup.JavaButton("DropDown").SetTOProperty "Index",iCount
								 objGroup.JavaButton("DropDown").Click micLeftBtn
								 If aValues(1) = "BLANK" Then
								 	bReturn = Fn_iComboSet(objGroup," ")
'									bReturn = Fn_UI_JavaObject_Click("Fn_SchMgr_FilterSettings", objGroup.JavaWindow("JavaWindow"), "BlankListItem", 4, 4, "LEFT")
									If bReturn = False Then
										Fn_SchMgr_CustomizeGroupFilters=false
										exit function
									Else
										Fn_SchMgr_CustomizeGroupFilters=True
									End If
								Else
									Set sTemplateType=Description.Create()
									sTemplateType("Class Name").value = "JavaStaticText"
									Set  intNoOfObjects =objGroup.ChildObjects(sTemplateType)
									  For i = 0 to intNoOfObjects.count-1
											   If  intNoOfObjects(i).getROProperty("label") = aValues(1) Then
														intNoOfObjects(i).Click 1,1
														Fn_SchMgr_CustomizeGroupFilters=True
														Exit for
											   End If
									  Next
								 End If
							     Exit for
		                   End If
					Next
	End If
	
	
	
'Set the Desired Values in The Customize Rollup Dialog
	
	If sRollup<>"" Then
		aIndex=split(sRollupIndex,":",-1,1)
		aRollupValues=split(sRollup,":",-1,1)
		objGroup.JavaButton("Rollup").SetTOProperty "Index",cInt(aValues(0))
		wait 2
		objGroup.JavaButton("Rollup").Click micLeftBtn
		wait 3

		'Set value for Rollup Field

			Set objButton=Description.Create()
					objButton("Class Name").value = "JavaButton"
					set sObjects=objGroup.ChildObjects(objButton)
					sValue= sObjects.Count
				For iCount=0 to cint(sValue)-1 step 1
						If cInt(aIndex(0))=iCount Then
							objRollup.JavaButton("DropDown").SetTOProperty "Index",iCount
							 objRollup.JavaButton("DropDown").Click micLeftBtn
							Set sTemplateType=Description.Create()
							sTemplateType("Class Name").value = "JavaStaticText"
							Set  intNoOfObjects =objRollup.ChildObjects(sTemplateType)
							  For i = 0 to intNoOfObjects.count-1
								   If  intNoOfObjects(i).getROProperty("label") = aRollupValues(0) Then
											intNoOfObjects(i).Click 1,1
										If objGroup.JavaDialog("Information").Exist(5)=true Then
												objGroup.JavaDialog("Information").JavaButton("OK").Click micLeftBtn
												Fn_SchMgr_CustomizeGroupFilters=false
												exit function
											Else
														Fn_SchMgr_CustomizeGroupFilters=True
														Exit for
											End If
											bFlag = True
											Exit for
								   End If
							  Next
							  Exit for
		End If
					Next

						'Set value for Rollup Condition
	
					For iCount=1 to cint(sValue) step 1
						If cInt(aIndex(1))=iCount Then
							objRollup.JavaButton("DropDown").SetTOProperty "Index",iCount
							objRollup.JavaButton("DropDown").Click micLeftBtn
							Set sTemplateType=Description.Create()
							sTemplateType("Class Name").value = "JavaStaticText"
							Set  intNoOfObjects =objRollup.ChildObjects(sTemplateType)
							  For i = 0 to intNoOfObjects.count-1
								   If  intNoOfObjects(i).getROProperty("label") = aRollupValues(1) Then
											intNoOfObjects(i).Click 1,1
											Fn_SchMgr_CustomizeGroupFilters=True
											Exit for
								   End If
							  Next
							  Exit for
		End If
					Next
	
				'	Click on the Required Button
				objRollup.JavaButton(aIndex(2)).Click micLeftBtn
					If err.number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click the button ["+aIndex(2)+"] from the Customize Rollups By dialog")
						Fn_SchMgr_CustomizeGroupFilters=False
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked the button ["+aIndex(2)+"] from the Customize Rollups By dialog")
						Fn_SchMgr_CustomizeGroupFilters=True	
					End If

	End If
	
					'click on the Required Button of the Customize Group By dialog
					objGroup.JavaButton(sButton).Click micLeftBtn
					If err.number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click the button ["+sButton+"] from the Customize Group By dialog")
						Fn_SchMgr_CustomizeGroupFilters=False
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked the button ["+sButton+"] from the Customize Group By dialog")
						Fn_SchMgr_CustomizeGroupFilters=True	
					End If
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_SchMgr_GetDay(sDate)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function return Day of the Date
''''/$$$$
''''/$$$$  PARAMETERS   : 		sDate: Date whose day is to be returned
''''/$$$$											
''''/$$$$	
''''/$$$$	Return Value : 				Day name of the concerned date
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(),
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   Nilesh                    08/12/2011     1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			   08/12/2011           1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SchMgr_GetDay(now)
''''/$$$$		
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

	Public Function Fn_SchMgr_GetDay(sDate)
		GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_GetDay"
	   Dim sDay,sResult
	   Fn_SchMgr_GetDay=False
	   sDay=Weekday(sDate)
	   Select Case sDay
		   Case 1
				sResult="Sunday"
		   Case 2
			   sResult="Moday"
		   Case 3
			   sResult="Tuesday"
		   Case 4
			   sResult="Wednesday"
		   Case 5
			   sResult="Thursday"
		   Case 6
			   sResult="Friday"
		   Case 7
			   sResult="Saturday"
	   End Select
		Fn_SchMgr_GetDay= sResult
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get the Day ["+Cstr(sResult)+"] of the Date ["+Cstr(sDate)+"]")	
	End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_SchMgr_ProgramViewSaveAs(sName,sDesc,bNewRoot,sButton,sExtra)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will Save As the Program View
''''/$$$$
''''/$$$$  PARAMETERS   : 		sName : Name of Program View
''''/$$$$											sDesc : Description
'''/$$$$	                                         bNewRoot : Set Show as new root checkbox 
''''/$$$$											sButton : Button name of Dialog 
''''/$$$$											sExtra : Reserved for future use
''''/$$$$											
''''/$$$$	Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(),Fn_MenuOperation(), Fn_ReadyStatusSync()
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   Nilesh                    14/12/2011      1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			   14/12/2011           1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SchMgr_ProgramViewSaveAs("Test","Testing","OFF","OK","")
''''/$$$$		
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SchMgr_ProgramViewSaveAs(sName,sDesc,bNewRoot,sButton,sExtra)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ProgramViewSaveAs"
	Dim objSaveDialog,aButtons,bReturn
	
	If Window("SchMgrWin").JavaWindow("JApplet").JavaDialog("Save As").Exist(5) = False Then
		'[TC1123-20161010-24_10_2016-VivekA-Maintenance] - Added for new Hierarchy of "Save As" dialog
		If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Save As").Exist(1) = False Then
			bReturn = Fn_MenuOperation("Select","File:Save As Program view")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Open Save As Dialog")
				Fn_SchMgr_ProgramViewSaveAs = False
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Opened Save As Dialog")
			End If
			Call Fn_ReadyStatusSync(1)
			
			If Window("SchMgrWin").JavaWindow("JApplet").JavaDialog("Save As").Exist(5) Then
				Set objSaveDialog = Window("SchMgrWin").JavaWindow("JApplet").JavaDialog("Save As")
			ElseIf JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Save As").Exist(1) Then
				Set objSaveDialog = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Save As")
			End If
		Else
			Set objSaveDialog = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Save As")
		End If
	Else
		Set objSaveDialog = Window("SchMgrWin").JavaWindow("JApplet").JavaDialog("Save As")
	End If
     
	If  sName <> "" Then
		objSaveDialog.JavaEdit("Name:").Set sName	
		If err.number <0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Name to ["+sName+"]")
			Fn_SchMgr_ProgramViewSaveAs=False
			Exit Function
		End If
	End If

	If  sDesc <> "" Then
		objSaveDialog.JavaEdit("Description:").Set sDesc	
		If err.number <0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Description to ["+sDesc+"]")
			Fn_SchMgr_ProgramViewSaveAs=False
			Exit Function
		End If
	End If


	If  bNewRoot <> "" Then
		If lcase(bNewRoot)="on" Then
				objSaveDialog.JavaCheckBox("Show as new root").Set "ON"
				If err.number <0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Show as new root to  ["+bNewRoot+"]")
					Fn_SchMgr_ProgramViewSaveAs=False
					Exit Function
				End If
		Elseif lcase(bNewRoot)="off" Then
				objSaveDialog.JavaCheckBox("Show as new root").Set "OFF"
				If err.number <0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Show as new root to  ["+bNewRoot+"]")
					Fn_SchMgr_ProgramViewSaveAs=False
					Exit Function
				End If
		End If
	End If
	If  sButton<>""Then
		objSaveDialog.JavaButton(sButton).Click micLeftBtn
		If err.number <0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on button ["+sButton+"]")
			Fn_SchMgr_ProgramViewSaveAs=False
			Exit Function
		End If
	Else
	   aButtons = split(sButtons, ":",-1,1)
	   iCounter = Ubound(aButtons)
	   For iCount=0 to iCounter
			objSaveDialog.JavaButton(aButtons(iCount)).Click micLeftBtn
			Call Fn_ReadyStatusSync(2)
       Next
		If err.number <0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on button ["+sButton+"]")
			Fn_SchMgr_ProgramViewSaveAs=False
			Exit Function
		End If
	End If

Fn_SchMgr_ProgramViewSaveAs=True
Set objSaveDialog=nothing

End Function


''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day(sAction,sWorkOrNonWork,sDate,sYear,sDay,sButtons,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will create a proxy task
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  Edit Schedule Calendar window should be Present
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''/$$$$										sWorkOrNonWork : To make a specific day as Working or a Non Working Day
''''/$$$$										sDate : Date from which the Month Has to be set
''''/$$$$										sYear : Sets the Specified Year
''''/$$$$										sDay : Selects the Specified day
''''/$$$$										sButtons : Clicks the Required Button
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
''''/$$$$    CREATED BY     :   SHREYAS          15/12/2011         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			15/12/2011            1.0
''''/$$$$
''''/$$$$     Modified By      :   Pritam		         16/12/2011       2.0                        Added Verify Case
'''/$$$$
''''/$$$$    REVIWED BY     :  Nilesh			16/12/2011            1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day("PerformOperation","Non Working","04/13/2011",cstr(Year(now)),"22","","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public function Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day(sAction,sWorkOrNonWork,sDate,sYear,sDay,sButtons,sInfo1,sInfo2)

	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day"
	Dim objCalendar,bReturn,sValue,sDifference,sDetails,sCurrent,aButtons,iCounter,iCount
	
	If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Schedule Calendar").Exist(2) = True Then
		Set objCalendar=JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Edit Schedule Calendar")
	ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Resource Calendar").Exist(2) = True Then
		Set objCalendar=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Edit Resource Calendar")
	Else
		Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day = False
		Exit Function
	End If
	

	Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False

	Select Case sAction

		Case "PerformOperation"


				 If sDay<>"" then
						sDetails=month(sDate)
						''Changed By Sushma, as current date displayed should be considered instead of today's date.
						''sCurrent= month(date)
						sCurrent= month(objCalendar.JavaEdit("Date").GetROProperty("value"))

						sDifference=cInt(sCurrent)-cInt(sDetails)
						If sDifference>0 Then
							For iCount=1 to sDifference
									objCalendar.JavaButton("ScrollLeft").Click micLeftBtn
									If err.number<0 Then
										Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
										Exit Function
									Else
										Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
									End If
							Next
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Month as  ["+monthname(month(sDate))+"]")
						Elseif sDifference<0 Then
							For iCount=1 to -(sDifference)
									objCalendar.JavaButton("ScrollRight").Click micLeftBtn
									If err.number<0 Then
										Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
										Exit Function
									Else
										Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
									End If
							Next
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Month as  ["+monthname(month(sDate))+"]")
						End If
						wait 3
			End if
		
				If sYear<>"" Then
							objCalendar.JavaEdit("Year").Set sYear
							If err.number<0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the year as  ["+sYear+"]")
								Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
								Exit Function
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the year as  ["+sYear+"]")
								Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
							End If
							objCalendar.JavaEdit("Year").Activate
				End If
		
			 If sDay<>"" then
			'Select the Desired Day in month
				objCalendar.JavaCheckBox("DayCheck").SetTOProperty "attached text",sDay
			
				objCalendar.JavaCheckBox("DayCheck").Set "ON"
					If err.number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Day as  ["+sDay+"]")
						Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Day as  ["+sDay+"]")
						Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
					End If
				wait 3
			End if
			
			'make it as a non working or working ay
				 If sWorkOrNonWork<>"" then
						If lCase(sWorkOrNonWork)="non working" Then
							objCalendar.JavaRadioButton("Working HH:MM").SetTOProperty "attached text","Non Working"
						
							objCalendar.JavaRadioButton("Working HH:MM").Set "ON"
								If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Work Status as  ["+sWorkOrNonWork+"]")
									Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Work Status as  ["+sWorkOrNonWork+"]")
									Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
								End If
						Elseif lCase(sWorkOrNonWork)="working hh:mm" then
		
							'will be coded later
		
						End If
		
			End If
		
		
				 If sButtons<>"" Then
				   aButtons = split(sButtons, ":",-1,1)
				   iCounter = Ubound(aButtons)
					   For iCount=0 to iCounter
						   objCalendar.JavaButton(aButtons(iCount)).Click micLeftBtn
							If err.number<0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Button ["+aButtons(iCounter)+"]")
								Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
								Exit Function
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Button ["+aButtons(iCounter)+"]")
								Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
							End If
					   Next
				 End If

      Case "Verify"
			
				If sDate<>"" then
						sDetails=month(sDate)
						''Changed By Sushma, as current date displayed should be considered instead of today's date.
						''sCurrent= month(date)
						sCurrent= month(objCalendar.JavaEdit("Date").GetROProperty("value"))
						sDifference=cInt(sCurrent)-cInt(sDetails)
						If sDifference>0 Then
							For iCount=1 to sDifference
									objCalendar.JavaButton("ScrollLeft").Click micLeftBtn
									If err.number<0 Then
										Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
										Exit Function
									Else
										Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
									End If
							Next
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Month as  ["+monthname(month(sDate))+"]")
						Elseif sDifference<0 Then
							For iCount=1 to -(sDifference)
									objCalendar.JavaButton("ScrollRight").Click micLeftBtn
									If err.number<0 Then
										Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
										Exit Function
									Else
										Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
									End If
							Next
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Month as  ["+monthname(month(sDate))+"]")
						End If
						wait 3
			End if
		
				If sYear<>"" Then
							objCalendar.JavaEdit("Year").Set sYear
							If err.number<0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the year as  ["+sYear+"]")
								Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
								Exit Function
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the year as  ["+sYear+"]")
								Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
							End If
							objCalendar.JavaEdit("Year").Activate
				End If
		
			 If sDay<>"" then
			'Select the Desired Day in month
				objCalendar.JavaCheckBox("DayCheck").SetTOProperty "attached text",sDay
			
				objCalendar.JavaCheckBox("DayCheck").Set "ON"
					If err.number<0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to set the Day as  ["+cstr(sDate)+"]")
						Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set the Day as  ["+cstr(sDate)+"]")
						Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True	
					End If
				wait 3
			End if
			
			If sWorkOrNonWork<>"" then
					If lCase(sWorkOrNonWork)="non working" Then
						objCalendar.JavaRadioButton("Working HH:MM").SetTOProperty "attached text","Non Working"
					
						If objCalendar.JavaRadioButton("Working HH:MM").GetRoProperty("value") = 1 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that ["+cstr(sDate)+"] is ["+sWorkOrNonWork+"] day" )
							Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=True
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Verify that ["+cstr(sDate)+"] is ["+sWorkOrNonWork+"] day")
							Fn_SchMgr_ScheduleCalendar_NonWorking_Or_Working_Day=False	
							Exit Function
						End If
					Elseif lCase(sWorkOrNonWork)="working hh:mm" then
	
						'will be coded later
	
					End If
		
			End If

	  End Select
	Set objCalendar=nothing
End Function

''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   : Fn_SchMgr_GetDateByDay()
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function return a Date for immidiate next WeekDay
''''/$$$$
''''/$$$$   PARAMETERS   : 			sAction : Action to be performed
''''/$$$$																		GetDateByDay  :To get the date of particular weekday from the sDate mentioned			 format  (mm/dd/yy) eg. 12/23/2011
''''/$$$$												sDate: From Which date the Day to de returned
''''/$$$$												sWeekDay : one of the WeekDay (Monday or Tuesday or Wednesday or Thursday or Friday or Saturday or Sunday)
''''/$$$$												sDays : How many days to add
''''/$$$$												sParam1 : For Future use
''''/$$$$												sParam2 : For Future use
''''/$$$$											
''''/$$$$	
''''/$$$$	Return Value : 				Day name of the concerned date
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(),
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :            Pritam               08/12/2011     1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			   08/12/2011           1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SchMgr_GetDay("GetDateByDay","Monday",now,"","","")
''''/$$$$		
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

	Public Function Fn_SchMgr_GetDate(sAction,sWeekDay,sDate,sDays,sParam1,sParam2)
		GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_GetDate"
	   Dim iDayValue,sResDate,sDateVal,sValue,sCurrentDate,iCount
	   Fn_SchMgr_GetDate=False
	   Select Case sWeekDay
		
		   Case "Sunday"
			   iDayValue=1
		   Case "Monday"
			   iDayValue=2
		   Case "Tuesday"
			   iDayValue=3
		   Case "Wednesday"
			   iDayValue=4
		   Case "Thursday"
			   iDayValue=5
		   Case "Friday"
			   iDayValue=6
		   Case "Saturday"
			   iDayValue=7
			   
	   End Select
	   
	   Select Case sAction
			
			Case "GetDateByDay"
				   sDateVal=sDate
				   sCurrentDate = Fn_SchMgr_FormatDate(sDateVal)
				   sValue=weekday(sCurrentDate)
					iCount = iDayValue-cint(sValue)
					sDateVal=dateadd("d",iCount,sCurrentDate)
					sResDate=Fn_SchMgr_FormatDate(sDateVal)
					If iCount<0 Then
						sDateVal=dateadd("d",7,sDateVal)
					End If
					sResDate=Fn_SchMgr_FormatDate(sDateVal)
			End Select
			
			Fn_SchMgr_GetDate = sResDate
		
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully get the Date ["+sDayValue+"] of the Day ["+sDay+"]")	

	End Function



'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_SchMgr_DeleteAllExistingRates()
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function will delete all the Existing Rates fgorm the Rate Modifier Dialog
'''''/$$$$
'''''/$$$$	
'''''/$$$$	Return Value : 			True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$										
'''''/$$$$
'''''/$$$$	HISTORY           :   AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS           30/12/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  SHREYAS			30/12/2011         1.0
'''''/$$$$
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SchMgr_DeleteAllExistingRates()
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_DeleteAllExistingRates"
   Dim sValue,bReturn,objRate,iCounter
   Set objRate=JavaWindow("ScheduleManagerWindow").JavaWindow("Manage Rate Modifiers")

    Fn_SchMgr_DeleteAllExistingRates=false

   If objRate.Exist(5)=false Then

			bReturn = Fn_MenuOperation("Select","Schedule:Rate Modifiers")
	
			If bReturn = True Then
				Fn_SchMgr_DeleteAllExistingRates = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Invoked Menu [Schedule-->Rate Modifiers.]")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule-->Rate Modifiers.]")
				Fn_SchMgr_DeleteAllExistingRates = False
				Set objRateModifier = Nothing 
				Exit Function
			End If

	 End If
	
			sValue=objRate.JavaTable("RateModTable").GetROProperty ( "rows")
			For iCounter=0 to sValue-1
			
					sValue=objRate.JavaTable("RateModTable").GetCellData (0,0)
					If sValue<>"" Then
						objRate.JavaTable("RateModTable").SelectRow 0
						 objRate.JavaButton("Delete").Click
						 	wait 1
						 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deleted Rate ["+sValue+"]")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Rates Exist")
						Exit for
					End If
			
			Next
	
	wait 2
	

'click on finish Button	
objRate.JavaButton("Finish").Click micLeftBtn
Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deleted all existing Rates")
Fn_SchMgr_DeleteAllExistingRates=true
Set objRate=nothing
	
End Function

''*********************************************************		Function to  Add or Remove privileged User for workflow task	***********************************************************************

'Function Name		:					Fn_SchMgr_WorkflowPrivilegedUser

'Description			 :		 		  This function is used to add or remove privileged User for workflow task

'Parameters			   :	 			1.  sAction: Add / Remove User
'													2.  sUserName   			
											
'Return Value		   : 				 True/False

'Pre-requisite			:		 		A Task should be selected in Schedule Manager window

'Examples				:				 Fn_SchMgr_WorkflowPrivilegedUser("Add",""Organization:Engineering:Designer:AutoTest1 (autotest1)")
'												     Fn_SchMgr_WorkflowPrivilegedUser("Remove","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sushma Pagare           28-Jun-12           1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SchMgr_WorkflowPrivilegedUser(sAction, sUserName)
	GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_WorkflowPrivilegedUser"
	Dim objWrkFlowDialog, descJavaWindow, objChildObjects, iCounter,  objPrivUserDialog, objTree, aUserInfo
	On Error Resume Next

	Set objWrkFlowDialog  = Fn_SISW_PPM_GetObject("Workflow Rule Configuration")
	If objWrkFlowDialog.Exist(5) = False  Then
			bReturn = Fn_MenuOperation("Select", "Schedule:Workflow Task")
			If bReturn = TRUE Then
                Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_WorkflowPrivilegedUser: Invoked Schedule-> Workflow Task.")
			Else
			   Fn_SchMgr_WorkflowPrivilegedUser = FALSE				 
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SchMgr_WorkflowPrivilegedUser : Failed to invoke Schedule-> Workflow Task")	
			  Exit Function
			End If	
	End If
	wait 2
	Select Case sAction
		'.---------------------------------------This case is used to add a privileged user.----------------------------------------------
		Case 	"Add"
			objWrkFlowDialog.JavaButton("AddPrivilegedUser").Click micLeftBtn

             Set descJavaWindow = Description.Create()
			descJavaWindow("Class Name").value = "JavaWindow"
			Set objChildObjects = JavaWindow("ScheduleManagerWindow").ChildObjects(descJavaWindow)
			For iCounter =0 to objChildObjects.Count-1
				If objChildObjects(iCounter).getROProperty("label") = "User or Resource Pool Selection" OR objChildObjects(iCounter).getROProperty("label") = "Privilege User" Then
						Set objPrivUserDialog =  objChildObjects(iCounter)
						Exit For
				End If
			Next
			Set objTree = objPrivUserDialog.JavaTree("to_class:=JavaTree")
            aUserInfo  = Split(sUserName, ":", -1,1)                                                         '' split  Organization:Engg:Designer:AutoTest1 (autotest1)
            objPrivUserDialog.JavaEdit("to_class:=JavaEdit","toolkit class:=org.eclipse.swt.widgets.Text").Set  trim(aUserInfo(UBound(aUserInfo)))                      ''Type  AutoTest1 (autotest1)" 
			wait 2
			If objTree.GetROProperty("items count") <> 0 Then  											
						objTree.Select sUserName
						wait 1
						objPrivUserDialog.JavaButton("attached text:=OK").WaitProperty  "enabled", 1, 20000
                        objPrivUserDialog.JavaButton("attached text:=OK").Click	
						wait 1
			Else
						Fn_SchMgr_WorkflowPrivilegedUser = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Privilege User")
                        objPrivUserDialog.JavaButton("attached text:=Cancel").Click	
						Set descJavaWindow = Nothing
						Set objChildObjects = Nothing
						Set objPrivUserDialog = Nothing
						Set objTree = Nothing
						Set objWrkFlowDialog = Nothing
                        Exit Function
			End If  
			wait 1
			objWrkFlowDialog.JavaButton("OK").WaitProperty  "enabled", 1, 20000
			objWrkFlowDialog.JavaButton("OK").Click micLeftBtn
			Set descJavaWindow = Nothing
			Set objChildObjects = Nothing
			Set objPrivUserDialog = Nothing
			Set objTree = Nothing
			Set objWrkFlowDialog = Nothing
        If Err.Number < 0 Then
			Fn_SchMgr_WorkflowPrivilegedUser = False																			      
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Click [Ok] Button of Workflow Rule configuration dialog")
			Exit Function
		Else
			Fn_SchMgr_WorkflowPrivilegedUser = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected ["+sUserName+"] as Privileged User and clicked OK")			
		End If
		Case "Remove"
				objWrkFlowDialog.JavaButton("RemoveUser").Click micLeftBtn
				objWrkFlowDialog.JavaButton("OK").WaitProperty  "enabled", 1, 20000
				objWrkFlowDialog.JavaButton("OK").Click micLeftBtn
				Set objWrkFlowDialog = Nothing

				If Err.Number < 0 Then
					Fn_SchMgr_WorkflowPrivilegedUser = False																			      
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Remove privileged user")
					Exit Function
				Else
					Fn_SchMgr_WorkflowPrivilegedUser = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully removed ["+sUserName+"] as Privileged User and clicked OK")			
				End If
		Case Else
			Set Fn_SchMgr_WorkflowPrivilegedUser = Nothing
    End Select

End Function


'*********************************************************		Generic function to handle Error dialogs   	***********************************************************************
'Function Name		:				Fn_SISW_SchMgr_ErrorVerify()

'Description			 :		 		 The function is generic function to handle error dialogs. It is created after combining error dialog functions from ScheduleManager.vbs
'										Fn_SchMgr_SchedulingErrorVerify
'										Fn_SchMgr_DialogMsgVerify
'										Fn_SchMgr_PercentLinkedMsgVerify
'										Fn_SchMgr_WarningMsgVerify
'										Fn_SchMgr_ErrorDialogMessageVerify
'										Fn_SchMgr_InformationDialogHandle

'Parameters			   :	 			1.  dicErrorInfo											
'Return Value		   : 				True/False
'Pre-requisite			:		 		NA.
'Examples				:				
'									Set dicErrorInfo = CreateObject("Scripting.Dictionary")
'									With dicErrorInfo	
'										.Add "Title", "Error"
'										.Add "Message", "The operation failed on one or more of the selected objects"
'										.Add "Button", "OK"
'										.Add "Action","DetailsMessageWithoutClose"
'									End with
'									bReturn = Fn_SISW_SchMgr_ErrorVerify(dicErrorInfo)

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sushma Pagare          20-Jun-2013
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_SchMgr_ErrorVerify(dicErrorInfo)
			GBL_FAILED_FUNCTION_NAME="Fn_SISW_SchMgr_ErrorVerify"
			Dim  dicKeys, dicItems, iCounter
			Dim sAction, sTitle, sErrorMsg,sButton, sAppMsg
			Dim bReturn			
            Dim objWin, objJavaWin, ObjJDialog,ObjWinDialog, objDetailsMsg
			Dim descDialog, descButton, descChild, objChild
            			
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
					Case "Button"
							sButton = dicItems(iCounter)					
				End Select
			Next		

			Fn_SISW_SchMgr_ErrorVerify = FALSE
            On Error Resume Next

			Select Case sAction

				''This covers     Fn_SchMgr_InformationDialogHandle(sTitle,sMessage,sButton,sDetails)
				Case "InformationDialogHandle"					
    					JavaDialog("Confirmation").SetTOProperty "title",sTitle						
						If JavaDialog("Confirmation").Exist(SISW_MIN_TIMEOUT) Then
							sAppMsg = JavaDialog("Confirmation").JavaObject("MLabel").GetROProperty("text")							
							If instr(sAppMsg, sErrorMsg) > 0 Then
								JavaDialog("Confirmation").JavaButton(sButton).Click
								wait(2)
								Call Fn_ReadyStatusSync(3)
								Fn_SISW_SchMgr_ErrorVerify = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully verified error message ["+sErrorMsg +"] on [Information] Dialog.")								
							Else
								GBL_ACTUAL_MESSAGE=sAppMsg
								Fn_SISW_SchMgr_ErrorVerify = False
								JavaDialog("Confirmation").Close
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify error message ["+sErrorMsg +"] on [Information] Dialog.")
							End If
						Else
							Fn_SISW_SchMgr_ErrorVerify = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to verify [Information] Dialog Exists.")
						End If   
						Exit Function    

				'' This covers Fn_SchMgr_PercentLinkedMsgVerify(sMessage)
				Case "PercentLinkedMsgVerify"

						JavaWindow("Percent Linked").Activate			
						sAppMsg = JavaWindow("Percent Linked").JavaStaticText("ErrMessage").GetROProperty("label")
						JavaWindow("Percent Linked").JavaButton("OK").Click micLeftBtn
						If sErrorMsg <> "" Then
								If instr(sAppMsg, sErrorMsg) > 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Validated message on [Percent Linked] Dialog" )
											Fn_SISW_SchMgr_ErrorVerify = True
								Else
											GBL_ACTUAL_MESSAGE=sAppMsg
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Validate message on [Percent Linked] Dialog" )
											Fn_SISW_SchMgr_ErrorVerify = False								
								End If
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked OK on [Percent Linked] Dialog" )
								Fn_SISW_SchMgr_ErrorVerify = True				
						End If     
						Exit Function

					''This covers Fn_SchMgr_ErrorDialogMessageVerify(sAction,sWindCaption, sErrMsg )  
					Case "DetailsMessageWithoutClose", "Error_Message"

							Set objJavaWin = JavaWindow("ScheduleManagerWindow").JavaWindow("Error")
							Set objDetailsMsg= objJavaWin.JavaEdit("Details")
                            objJavaWin.SetTOProperty "title", sTitle

							If objDetailsMsg.Exist(SISW_MIN_TIMEOUT) = False Then
									Call Fn_UI_JavaStaticText_Click("Fn_SISW_ErrorVerify", objJavaWin.JavaStaticText("Details"), "Details", 15, 0, "LEFT")                                                                   
                                    objDetailsMsg.RefreshObject
							End If
							If sAction = "DetailsMessageWithoutClose" Then
								sAppMsg = objDetailsMsg.GetROProperty("value")
							ElseIf sAction = "Error_Message" Then
								If objJavaWin.JavaEdit("ErrMessage").Exist(2) Then		'[TC1121-2015102600-06_11_2015-VivekA-Maintenance] - Added by Nishigandha J
									sAppMsg = objJavaWin.JavaEdit("ErrMessage").GetROProperty("value")
								Else 
									sAppMsg = objJavaWin.JavaStaticText("ErrMessage").GetROProperty("value")								
								End If	
								'sAppMsg = objJavaWin.JavaStaticText("ErrMessage").GetROProperty("value")	
								Call Fn_Button_Click("Fn_ErrorDialogMessageVerify", objJavaWin,sButton)
							End If							
							If instr(sAppMsg, sErrorMsg) > 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Validated message on Error Dialog" )
									Fn_SISW_SchMgr_ErrorVerify = True
							Else
									GBL_ACTUAL_MESSAGE=sAppMsg
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Validate message on Error Dialog" )
									Fn_SISW_SchMgr_ErrorVerify = False								
							End If      							
							Exit Function

					Case Else

							' This covers  	1)Fn_SchMgr_SchedulingErrorVerify(sMesssage, sButton),  2) Fn_SchMgr_DialogMsgVerify(sTitle,sMsg,sButton)   3)Fn_SchMgr_WarningMsgVerify(sMesssage, sButton)				 
							'If Error Message is blank, then take it from global variable
							If sErrorMsg = "" Then
								sErrorMsg = sErrorText
							End If
							Fn_SISW_SchMgr_ErrorVerify = True
				
							Set descDialog=description.Create()
							descDialog("micclass").value="Dialog"
							descDialog("regexpwndtitle").value = sTitle
							descDialog("regexpwndclass").value = "#32770"
							'Description of  Button Object  on  dialog
							Set descButton=description.Create()
							descButton("micclass").value="WinButton"
							descButton("nativeclass").value = "Button"
							descButton("regexpwndtitle").value = sButton
							'General Object description to search all Objects
							Set descChild=description.Create()
				
							JavaWindow("ScheduleManagerWindow").Dialog("ErrorDialog").SetTOProperty "text", sTitle
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title", sTitle				
							JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("JavaErrorDialog").SetTOProperty "title", sTitle						
							Window("SchMgrWin").JavaWindow("JApplet").JavaDialog("Scheduling Error").SetTOProperty "title", sTitle
							JavaDialog("Error").SetTOProperty "title", sTitle
							JavaWindow("ScheduleManagerWindow").JavaWindow("Error").SetTOProperty "title", sCaption
							
							Set objJavaWin =JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("JavaErrorDialog")
							Set ObjJDialog = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog")
							Set ObjWinDialog = Window("SchMgrWin").JavaWindow("JApplet").JavaDialog("Scheduling Error")            
							Set objWin = JavaWindow("ScheduleManagerWindow").Dialog("ErrorDialog")
				
							If Dialog(descDialog).Exist(SISW_MIN_TIMEOUT) Then
									'Capture All runtime objects to find message text
									Set  objChild = Dialog(descDialog).ChildObjects(descChild)
									sAppMsg = objChild(1). getroproperty("text")
									wait(2)
									Dialog(descDialog).WinButton(descButton).Click 10,10,micLeftBtn
									If Dialog(descDialog).Exist(1) Then
										Dialog(descDialog).Close()
									End If									
							ElseIf objJavaWin.Exist(SISW_MICRO_TIMEOUT) Then		     ''This dialog will even serve "Validate Inline Editing" dialog	after setting tiltle		
									sAppMsg = objJavaWin.JavaEdit("JTextArea").GetROProperty("value")					
									objJavaWin.JavaButton("OK").SetTOProperty "label", sButton
									Call  Fn_SISW_UI_JavaButton_Operations("Fn_SISW_SchMgr_ErrorVerify", "DeviceReplay.Click",objJavaWin,"OK")
									'objJavaWin.JavaButton("OK").Click micLeftBtn            									
							ElseIf ObjJDialog.Exist(SISW_MICRO_TIMEOUT) Then					
									sAppMsg = ObjJDialog.JavaEdit("ErrText").GetROProperty("value")
									ObjJDialog.JavaButton("OK").SetTOProperty "label", sButton
									ObjJDialog.JavaButton("OK").Click micLeftBtn														
							 ElseIf ObjWinDialog.Exist(SISW_MICRO_TIMEOUT) Then     ''This dialog will even serve "Validate Inline Editing" dialog after setting tiltle
									sAppMsg = ObjWinDialog.JavaEdit("ErrText").GetROProperty("value")
									ObjWinDialog.JavaButton("OK").SetTOProperty "label", sButton
									ObjWinDialog.JavaButton("OK").Click micLeftBtn														
							ElseIf objWin.Exist(SISW_MICRO_TIMEOUT) Then		
									sAppMsg = objWin.Static("ErrText").GetROProperty("text")
									objWin.WinButton(sButton).Click micLeftBtn					
									Fn_SISW_SchMgr_ErrorVerify = True
							ElseIf JavaWindow("ScheduleManagerWindow").JavaWindow("Error").Exist(SISW_MICRO_TIMEOUT) Then      ''Covers Fn_SchMgr_WarningMsgVerify(sMesssage, sButton)
									sAppMsg = JavaWindow("ScheduleManagerWindow").JavaWindow("Error").JavaStaticText("ErrMessage").GetROProperty("attached text")
									 JavaWindow("ScheduleManagerWindow").JavaWindow("Error").JavaButton(sButton).Click micLeftBtn									 
							ElseIf  JavaDialog("Error").Exist(SISW_MICRO_TIMEOUT) Then                                                                                               ''Covers  Fn_SchMgr_WarningMsgVerify(sMesssage, sButton)
									sAppMsg = JavaDialog("Error").JavaEdit("ErrMsg").GetROProperty("value")
									If sButton <> "" Then
										JavaDialog("Error").JavaButton("OK").SetTOProperty "label", sButton
									End If
									JavaDialog("Error").JavaButton("OK").Click micLeftBtn
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[Scheduling Error] Dialog Does not Exist")
								Fn_SISW_SchMgr_ErrorVerify =False								
							End If
							
							If Fn_SISW_SchMgr_ErrorVerify = True Then
								'Compare message if sent in
								If sErrorMsg <> "" Then
										If instr(sAppMsg,sErrorMsg) > 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Sucessfully Verified Message : "+ sErrorMsg)
										Else
											GBL_ACTUAL_MESSAGE=sAppMsg
											Fn_SISW_SchMgr_ErrorVerify =False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Message Verification Failed : "+ sErrorMsg)
										End If
								End If
							End If
							
							Set objWin = Nothing
							Set objJavaWin = Nothing
							Set ObjJDialog = Nothing
							Set ObjWinDialog = Nothing
							Set descDialog=nothing
							Set descButton=nothing
							Set descChild=nothing
							Set objChild=nothing	
						
			End Select

			
End Function

''*********************************************************		Function to verify "Open by Name" wizard values	***********************************************************************

'Function Name		:					Fn_ScMgr_DeliverableOpenByNameVerify

'Description			 :		 		  This Function is used to verify "Open by Name" wizard values

'Parameters			   :	 			1.  sAction: Action to be performed ie. Verify or else
'										2.  sSchName: Schedule name
'										3.  sDelName:  Deliverable name
'										4.  sDelType:  Deliverable type
'										5.  sSrchName: Search string (wild card or any stand)
'										6.  sColToVerify:  Column name which to be verified
'										7.  sValueToVerify: Valuee to be verified under above selected function
'										8.  sExtra:  future use

											
'Return Value		   : 				 True/False

'Pre-requisite			:		 		A Task should be selected in Schedule Manager window

'Examples				:				 bReturn = Fn_ScMgr_DelvrblOpenByNameVerify("Verify","SchName100","DelName1234","MS Word","*","Type", "MS Word", "")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rajendra Patil           26-May-15           1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ScMgr_DeliverableOpenByNameVerify(sAction,sSchName,sDelName,sDelType,sSrchName,sColToVerify, sValueToVerify, sExtra)
	GBL_FAILED_FUNCTION_NAME="Fn_ScMgr_DeliverableOpenByNameVerify"
	Dim bReturn,objScheduleDel,sIndex,objDelTable,objTable,iTypeIndex, iColIndex, sArr
	Dim sTitle, objOpenByName, iRowCnt, iColCnt, sColName, sCellVal
	sTitle = "Update Schedule Deliverable Error"
	sErrorText = "Schedule deliverables should have unique names."
	Set objScheduleDel = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Deliverables")
	Set objOpenByName = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Open by Name")
	Set objDelTable = objOpenByName.JavaTable("DelTable")
	Fn_ScMgr_DeliverableOpenByNameVerify = False
	   	If Not objScheduleDel.Exist(SISW_MIN_TIMEOUT) Then
			bReturn = Fn_SchMgr_SchTable_NodeOperation("Select",sSchName,"","","")
			If  bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select  schedule " + sSchName)
				Set objScheduleDel = Nothing
				sErrorText = ""
				Exit Function
			End If
	
			bReturn = Fn_MenuOperation("Select","Schedule:Schedule Deliverables")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoked Menu [Schedule:Schedule Deliverables]")
				Set objScheduleDel = Nothing
				sErrorText = ""
				Exit Function
		    End If
		End If
		
	    If objScheduleDel.Exist(SISW_MIN_TIMEOUT) Then
			'Set the value of java table
			Set objTable = JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Schedule Deliverables").JavaTable("SchDeliverablesTable")		
			objScheduleDel.JavaButton("Add").WaitProperty "enabled",1,20000
			objScheduleDel.JavaButton("Add").Click micLeftBtn
			sIndex =Cint( objTable.GetROProperty("rows"))
			If sIndex > 0 Then
					'Set Deliverable name
				If sDelName <> "" Then
					objTable.ClickCell (sIndex-1),"#0","LEFT","NONE"
	                objScheduleDel.JavaEdit("DelTableEdit").Object.setText sDelName
					objTable.ClickCell (sIndex-1),"#2","LEFT","NONE"
				End If
	
				bReturn = Fn_SchMgr_DialogMsgVerify(sTitle,sErrorText,"OK")
				If bReturn Then
					objScheduleDel.JavaButton("Cancel").Click
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add deliverable.[" + sDelName + "]")
					Set objScheduleDel  = Nothing 
					Set objTable = Nothing
					Set objDelTable = Nothing
					sErrorText = ""
					Exit Function
				End If
				'Set Deliverable type
				If  sDelType <> "" Then
					objTable.ClickCell (sIndex-1), "#1","LEFT", "NONE"
					iTypeIndex = objScheduleDel.JavaList("SchDelList").GetItemIndex(sDelType)
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sDelType + " deliverable type not exist.")
						objScheduleDel.JavaButton("Cancel").Click
						Set objScheduleDel  = Nothing 
						Set objTable = Nothing
						Set objDelTable = Nothing
						sErrorText = ""
						Exit Function
					End If
					objScheduleDel.JavaList("SchDelList").Object.setSelectedIndex Cint(iTypeIndex)
				End If
				
				objTable.ClickCell (sIndex-1), "#3","LEFT", "NONE"
				
				If objOpenByName.Exist(SISW_MIN_TIMEOUT) Then
					objOpenByName.JavaEdit("Name").Object.setText sSrchName
					objOpenByName.JavaButton("Find").Object.doClick(1)
					'Handle Nothing object dialog .
					If JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Open by Name").JavaDialog("Nothing found!").Exist(SISW_MIN_TIMEOUT) Then
						JavaWindow("ScheduleManagerWindow").JavaWindow("ScheduleManagerApplet").JavaDialog("Open by Name").JavaDialog("Nothing found!").JavaButton("OK").Click micLeftBtn
						Fn_ScMgr_DeliverableOpenByNameVerify = False
						'Exit Function
					Else
						
						If objOpenByName.JavaButton("LoadAll").Exist(SISW_MICRO_TIMEOUT) Then
							If objOpenByName.JavaButton("LoadAll").GetROProperty("enabled") = 1 Then
								objOpenByName.JavaButton("LoadAll").Object.doClick 1
								wait 3
								Call Fn_ReadyStatusSync(3)
							End If
						End If
		
						iColCnt = objDelTable.GetROProperty("cols")
						For iColCntr = 0 To cInt(iColCnt)-1
						sColName = objDelTable.Object.GetColumnName(iColCntr)
							If sColName = sColToVerify Then
							   iColIndex = iColCntr
							   Exit for
							End If
						Next	
					End If
				End if
				
				Select Case sAction
					Case "Verify"
						Select Case sColToVerify
							Case "Type"															
								iRowCnt = objDelTable.GetROProperty("rows")
								If iRowCnt > 0 Then
									For iRowCntr = 0 To cInt(iRowCnt)-1
										sCellVal = objDelTable.GetCellData(iRowCntr, iColIndex)
										Fn_ScMgr_DeliverableOpenByNameVerify = True
										If sCellVal <> sDelType Then
											Fn_ScMgr_DeliverableOpenByNameVerify = False
											Exit For
										End If								
									Next								
								End If
								
							Case "Name","Object"
								iRowCnt = objDelTable.GetROProperty("rows")
								If iRowCnt > 0 Then
									sArr = Split(sValueToVerify, "~")
									For iValCntr = 0 To Ubound(sArr)
										For iRowCntr = 0 To cInt(iRowCnt)-1
											sCellVal = objDelTable.GetCellData(iRowCntr, iColIndex)
											If sCellVal = sArr(iValCntr) Then
												Fn_ScMgr_DeliverableOpenByNameVerify = True
												Exit For
											End If								
										Next	
									Next									
								End If				
						End Select
				End Select
			End if	

				If objOpenByName.Exist(SISW_MICRO_TIMEOUT) Then
					objOpenByName.Close
					If objScheduleDel.Exist(SISW_MICRO_TIMEOUT) Then
						objScheduleDel.JavaButton("Cancel").Click
					End If
				End If						
		End if
					
		Set objScheduleDel  = Nothing 
		Set objTable = Nothing
		Set objDelTable = Nothing	
End Function
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' 
'   FUNCTION NAME   : Fn_SchMgr_VerityTaskTimeAndDuration()
'	
'	DESCRIPTION     :  This function is used to Compare Time & Duration of Task with Passed CurrentTime and Duration
'
'	Return Value    :  True or False
'
'	HISTORY         :  AUTHOR              	DATE        		Changes 		VERSION
'
'   CREATED BY      :  Poonam Chopade     	20-Jan-2017      	Created 	 	  1.0
'
'	Example			:  Call	Fn_SchMgr_VerityTaskTimeAndDuration("VerifyTime",Now,"20-Jan-2017 16:45","","")
'                      Call Fn_SchMgr_VerityTaskTimeAndDuration("VerifyDuration","","","0.2h","1")
'
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public function Fn_SchMgr_VerityTaskTimeAndDuration(sAction,sCurrentDateTime,sTaskDateTime,sTaskDuration,sDurToCompare)
    GBL_FAILED_FUNCTION_NAME="Fn_SchMgr_VerityTaskTimeAndDuration"
    Dim ActualSec,sTDateTime,sTDuration
       
    Fn_SchMgr_VerityTaskTimeAndDuration=false
     
    Select Case sAction
    	Case "VerifyTime"
	    		If sCurrentDateTime <> "" and sTaskDateTime <> "" Then
			    		sCurrentDateTime = Day(sCurrentDateTime)&"-"&MonthName(Month(sCurrentDateTime),true)&"-"&year(sCurrentDateTime)&" "&Hour(sCurrentDateTime)&":"&Minute(sCurrentDateTime)&":"&Second(sCurrentDateTime)
			    		sTDateTime = sTaskDateTime & ":" & "00"
			    		ActualSec = DateDiff("s",sTDateTime,sCurrentDateTime)
			    		If ActualSec >= 0 and ActualSec <=120 Then
		    				Fn_SchMgr_VerityTaskTimeAndDuration = True
		    			Else
		    				Fn_SchMgr_VerityTaskTimeAndDuration = false
		    				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify time for task.")
		    			End If	
			    End If
		  
    	 Case "VerifyDuration"
	    	   If sTaskDuration <> "" and sDurToCompare <> "" Then
		    	     sTDuration = Replace(sTaskDuration,"h","")
		    	     If sTDuration >=0 and sTDuration <=sDurToCompare Then
		    	     	  Fn_SchMgr_VerityTaskTimeAndDuration = True
		    	     Else
		    	     	 Fn_SchMgr_VerityTaskTimeAndDuration = false
				    	 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify duration for task.")
		    	     End If
	    	   End If
    End Select 
	
End Function
