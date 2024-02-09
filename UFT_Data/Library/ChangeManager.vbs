Option Explicit
'Public  iTimeOut
iTimeOut=180
'-------------------------------'Global variables for Teamcenter Perspective Names-------------------------------------------------------
Public GBL_PERSPECTIVE_CHANGEMANAGER
GBL_PERSPECTIVE_CHANGEMANAGER = "Change Manager"
'-------------------------------'Global variables for Teamcenter Perspective Names-------------------------------------------------------
'*********************************************************	Function List		****************************************************************************************************************************************
'0.Fn_SISW_CM_GetObject()
'1.Fn_CM_ManageSaveSearchOperations()
'2.Fn_CM_ViewTreeOperations()
'3.Fn_CM_ErrorMessageVerify()
'4.Fn_CM_AssignParticipantsOperations()
'5.Fn_CM_AssignParticipantsTreeOperations()
'6.Fn_CM_AssignParticipantsStaticTextOperations()
'7.Fn_CM_AssignParticipantsGUIOprts()
'8.Fn_CM_ChangeTypeTreeOperations()
'9.Fn_CM_ChangeErrorMsgVerify()
'10.Fn_CM_SummaryOperation()
'11.Fn_CM_RandNoGenerate()
'12.Fn_CM_TabOperation()
'13.Fn_CM_CreateChangeInContext()
'14.Fn_CM_SrchResltTreeOperation()   - Unsued function Eliminated 
'15.Fn_CM_ComponentTreeOperations()
'16.Fn_CM_ViewMenuOperations()
'17.Fn_CM_CreateTask()
'18.Fn_CM_TaskTreeOperations()
'19.Fn_CM_CommitRollupDrpDwnOperation()
'20.Fn_CM_RollupTreeOperations()
'21.Fn_CM_DeriveChangeCreate()
'22.Fn_CM_ErrorWindowMsgVerify()
'23 Fn_CM_SummaryPropertyVerify()
'24 Fn_CM_ObjectPropertyPanelVerify()
'25 Fn_CM_TaskAssignment()
'26 Fn_CM_ScheduleMembership()
'27 Fn_CM_SpecifyQueryDetailsAndInvoke()
'28 Fn_Chng_SrchSavedSearchOperation()
'29 Fn_CM_GetTreeItemPath()
'30 Fn_CM_getJavaTreeIndex()
'31 Fn_SISW_CM_ErrorVerify()
'32 Fn_CM_SummaryTabTableOperations
'33 Fn_CM_ChangeInContext_Operations
'*********************************************************	Function List		****************************************************************************************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_CM_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_CM_GetObject("Change Manager")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sukhada Bakshi		 17-5-2013
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_CM_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ChangeManager.xml"
	Set Fn_SISW_CM_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'-------------------------------------------------------------------Function Used to Manage saved searches----------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_ManageSaveSearchOperations

'Description			 :	Function Used to Manage saved searches

'Parameters			   :	1.strInvokeOption: ManageSavedSearches window Invocation Option . ViewMenu Or ToolBar
										'2.strAction: Acion to Perform Eg. Add
										'3.strSrchName: Search Name
										'4.strType:Search Type (This parameter not yet handle in this function)
										'5.bShowOptn:Show Option (This parameter not yet handle in this function)
										'6.strAssgnSearch:Assign Search type
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective

'Examples				:	Fn_CM_ManageSaveSearchOperations("ViewMenu","Add","TestSearch","","","Test Search")
										'Fn_CM_ManageSaveSearchOperations("ToolBar","Add","Test1","","","Test1")
										'Fn_CM_ManageSaveSearchOperations("ToolBar","GetType","My Open Changes","","","")
										'Fn_CM_ManageSaveSearchOperations("ToolBar","Show","My Open Changes","","","")
										'Fn_CM_ManageSaveSearchOperations("ToolBar","Remove","Folder","","","")
										'Fn_CM_ManageSaveSearchOperations("ToolBar","Rename","My Open Searches~New Search Name","","","Test Search")
										'Imp Note- In Rename Case Separate The Search Names By { ~ } (Tilda) first is Existing Name ~ Second is New name which have to set
										'Fn_CM_ManageSaveSearchOperations("ToolBar","Rename","My Open Searches","","","Test Search")
										'Fn_CM_ManageSaveSearchOperations("ToolBar","VerifySearchName","My Open Searches","","","")
										'Fn_CM_ManageSaveSearchOperations("ToolBar","InvalidAssgnSearch","TestSearch","","","Change Notice Revision...")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				03/08/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ManageSaveSearchOperations(strInvokeOption,strAction,strSrchName,strType,bShowOptn,strAssgnSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_ManageSaveSearchOperations"
   'Declaring Variables
	Dim  iRowNumber,iRowCnt,iCount,sSrchName,arrSearchName,strProp,strAssgnSearchName
   Dim ObjSavedSrchWnd,objSelectType,objDialog
	'Function Return False
	Fn_CM_ManageSaveSearchOperations=False
   Select Case strInvokeOption
		 	Case "ViewMenu"
				'Invoking Manage saved searches Window by PopUp Menu
                Call Fn_ToolbatButtonClick("View Menu")
				wait(1)
				JavaWindow("ChangeManager").WinMenu("ContextMenu").Select "Manage saved searches..."
				Call Fn_ReadyStatusSync(5)
            Case "ToolBar"
                If not JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").Exist(5) Then
					'Invoking Manage saved searches Window by clicking on Folder Ikon
					Call Fn_ToolbatButtonClick("Manage Change Home saved searches")
					Wait 1
					Call Fn_ReadyStatusSync(5)
				End If
			Case Else
                Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:Wrong Invoke Option Pass")
				Exit Function
   End Select
   'Verifying existance of "ManageSavedSearches" window
	If Fn_UI_ObjectExist("Fn_CM_ManageSaveSearchOperations",JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches"))=False Then
		Exit Function
	End If
	'Creating object  of "ManageSavedSearches" window
	Set ObjSavedSrchWnd=Fn_UI_ObjectCreate("Fn_CM_ManageSaveSearchOperations",JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches"))
	   Select Case strAction
				Case "Add"
					'Clicking on Add button 
                    Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"Add")
                    iRowCnt=Fn_UI_Object_GetROProperty("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd.JavaTable("SavedSearchesTable"),"rows")
					For iCount=0 To iRowCnt-1
                        Call Fn_UI_JavaTable_SelectRow("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"SavedSearchesTable",iCount)
						sSrchName=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iCount,"Search name")
						If sSrchName="New Search" Then
							iRowNumber=iCount
                            Exit For
						End If
					Next
					If strType<>"" Then
						JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Type"
						'Selecting Type
						Set objSelectType=description.Create()
						objSelectType("Class Name").value = "JavaList"
						Set objDialog =JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").ChildObjects(objSelectType)
						 objDialog(0).Select strType
						 wait 1
						Set objSelectType=Nothing
						Set objDialog=Nothing
					End If
					
					JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Assigned search"
					'Assigning search 
                    Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaList"
					Set objDialog =JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").ChildObjects(objSelectType)
                     objDialog(0).Select strAssgnSearch
					 JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Type"
					 
					 JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Search name"
					 wait 1
				If strSrchName<>"" Then
						'Setting Search Name
						'ObjSavedSrchWnd.JavaTable("SavedSearchesTable").SetCellData iRowNumber,"Search name",strSrchName
						Set objSelectType=description.Create()
						objSelectType("Class Name").value = "JavaEdit"
						Set objDialog =JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").ChildObjects(objSelectType)
                     	objDialog(0).set strSrchName
						wait 1
					End If
					JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber-1 ,"Search name"
					wait 1
					 
					 Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"OK")  
					 Call Fn_ReadyStatusSync(2)
					Fn_CM_ManageSaveSearchOperations=True
				 Case "Show"
					iRowCnt=Fn_UI_Object_GetROProperty("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd.JavaTable("SavedSearchesTable"),"rows")
					For iCount=0 To iRowCnt-1
						Call Fn_UI_JavaTable_SelectRow("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"SavedSearchesTable",iCount)
						sSrchName=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iCount,"Search name")
						If sSrchName=strSrchName Then
							iRowNumber=iCount
							Exit For
						End If
					Next
					JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Show"
					JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Type"
					Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"OK")  
					Fn_CM_ManageSaveSearchOperations=True
				Case "GetType"
						iRowCnt=Fn_UI_Object_GetROProperty("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd.JavaTable("SavedSearchesTable"),"rows")
						For iCount=0 To iRowCnt-1
							Call Fn_UI_JavaTable_SelectRow("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"SavedSearchesTable",iCount)
							wait(2)
							sSrchName=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iCount,"Search name")
							If sSrchName=strSrchName Then
								iRowNumber=iCount
								Exit For
							End If
						Next
						Fn_CM_ManageSaveSearchOperations=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iRowNumber,"Type")
						Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"Cancel")
						wait(2)
                 Case "Remove"				
						iRowCnt=Fn_UI_Object_GetROProperty("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd.JavaTable("SavedSearchesTable"),"rows")
						For iCount=0 To iRowCnt-1
							Call Fn_UI_JavaTable_SelectRow("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"SavedSearchesTable",iCount)
							sSrchName=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iCount,"Search name")
							If sSrchName=strSrchName Then
								iRowNumber=iCount
								Exit For
							End If
						Next
						JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Search name"
						Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"Remove")  
						Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"OK")  
						Fn_CM_ManageSaveSearchOperations=True
				 Case "Rename"				
						iRowCnt=Fn_UI_Object_GetROProperty("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd.JavaTable("SavedSearchesTable"),"rows")
						arrSearchName=Split(strSrchName,"~")
						For iCount=0 To iRowCnt-1
							Call Fn_UI_JavaTable_SelectRow("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"SavedSearchesTable",iCount)
							sSrchName=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iCount,"Search name")
							If sSrchName=arrSearchName(0) Then
								iRowNumber=iCount
								Exit For
							End If
						Next
						If Ubound(arrSearchName)=1 Then
							If arrSearchName(1)<>"" Then
								JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Search name"
								'JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SetCellData iRowNumber,"Search name",arrSearchName(1)
                                Set objSelectType=description.Create()
								objSelectType("Class Name").value = "JavaEdit"
								Set objDialog =JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").ChildObjects(objSelectType)
								objDialog(0).Set arrSearchName(1)
								Set objSelectType=Nothing
								Set objDialog =Nothing
							End If
						End If
						If strAssgnSearch<>"" Then
							JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Assigned search"
							'Assigning search 
							Set objSelectType=description.Create()
							objSelectType("Class Name").value = "JavaList"
							Set objDialog =JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").ChildObjects(objSelectType)
							 objDialog(0).Select strAssgnSearch
						End If
						JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Type"
						Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"OK")  
						Fn_CM_ManageSaveSearchOperations=True
				 Case "VerifySearchName"
						iRowCnt=Fn_UI_Object_GetROProperty("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd.JavaTable("SavedSearchesTable"),"rows")
						For iCount=0 To iRowCnt-1
							Call Fn_UI_JavaTable_SelectRow("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"SavedSearchesTable",iCount)
							sSrchName=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iCount,"Search name")
							If sSrchName=strSrchName Then
								Fn_CM_ManageSaveSearchOperations=True
								Exit For
							End If
						Next
						Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"Cancel")

				Case "InvalidAssgnSearch"
					'Clicking on Add button 
                    Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"Add")
                    iRowCnt=Fn_UI_Object_GetROProperty("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd.JavaTable("SavedSearchesTable"),"rows")
					For iCount=0 To iRowCnt-1
                        Call Fn_UI_JavaTable_SelectRow("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"SavedSearchesTable",iCount)
						sSrchName=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iCount,"Search name")
						If sSrchName="New Search" Then
							iRowNumber=iCount
                            Exit For
						End If
					Next
					If strType<>"" Then
						JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Type"
						'Selecting Type
						Set objSelectType=description.Create()
						objSelectType("Class Name").value = "JavaList"
						Set objDialog =JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").ChildObjects(objSelectType)
						 objDialog(0).Select strType
						Set objSelectType=Nothing
						Set objDialog=Nothing
					End If
					JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Search name"
					If strSrchName<>"" Then
						'Setting Search Name
						ObjSavedSrchWnd.JavaTable("SavedSearchesTable").SetCellData iRowNumber,"Search name",strSrchName
					End If
					JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Assigned search"
					'Assigning search 
                    Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaList"
					Set objDialog =JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").ChildObjects(objSelectType)
                     objDialog(0).Select strAssgnSearch
					 JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").SelectCell iRowNumber,"Type"
					 strProp=Fn_UI_Object_GetROProperty("Fn_CM_ManageSaveSearchOperations", JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaButton("OK"), "enabled")
					If Cint(strProp)=0 Then
							strAssgnSearchName=JavaWindow("ChangeManager").JavaWindow("ManageSavedSearches").JavaTable("SavedSearchesTable").GetCellData(iRowNumber,"Assigned search")
							If strAssgnSearchName="-- Select one --" Then
								Fn_CM_ManageSaveSearchOperations=False
								Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"Cancel") 
                                Call Fn_CM_ErrorMessageVerify("Discarding unsaved changes","OK to exit","OK")
							End If
					Else
							Fn_CM_ManageSaveSearchOperations=True
							Call Fn_Button_Click("Fn_CM_ManageSaveSearchOperations",ObjSavedSrchWnd,"OK")  
					End If

				Case Else
					Set ObjSavedSrchWnd=Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:Wrong Action "& strAction)
					Exit Function
	   End Select       
		'Releasing all objects
        Set ObjSavedSrchWnd=Nothing
		Set objSelectType=Nothing
		Set objDialog =Nothing
End Function

'-------------------------------------------------------------------Function Used to perform operatons on View Tree---------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_ViewTreeOperations

'Description			 :	Function Used to perform operatons on View Tree

'Parameters			   :	1.strAction: Action Name
										'2.strNodeName: Node Name		
										'3.strMenu: Pop Up Menu Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective

'Examples				:	Fn_CM_ViewTreeOperations("Select","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
										'Fn_CM_ViewTreeOperations("DoubleClick","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
										'Fn_CM_ViewTreeOperations("Expand","Change Home:My Open Changes","")
										'Fn_CM_ViewTreeOperations("VerifyNode","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
										'Fn_CM_ViewTreeOperations("PopupMenuSelect","Change Home:Test1","Refresh")
										'Fn_CM_ViewTreeOperations("MultiSelect","Change Home:My Open Changes:ECN-074691/A;1-CM76,Change Home:My Open Changes:ECN-459679/A;1-CM76,Change Home:My Open Changes:PR-581018/A;1-PR130","")   
										'Fn_CM_ViewTreeOperations("MultiSelect","Change Home:Snd:ECR-954847/A;1-ECR1,Change Home:Snd:ECR-775411/A;1-ECR3","")
										'Fn_CM_ViewTreeOperations("MultiSelectPopupMenu","Change Home:Snd:ECR-516067/A;1-ECR2","Derive Change...")
										'Fn_CM_ViewTreeOperations("GetNodeIndex","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
'										Fn_CM_ViewTreeOperations("Collapse","Change Home:My Open Changes","")
'History					 :			
'	Developer Name				Date			Rev. No.	Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sandeep N					05/08/2010		1.0												Tushar B
'	Sandeep N					19/01/2011		1.0			Added Case "GetNodeIndex"		Harshal A
'	Sandeep N					20/01/2011		1.0			Added Case "Collapse"		Harshal A
'	Sandeep N					15/11/2011		2.0			Modified All case by adding "Fn_UI_JavaTreeGetItemPath" UI function
'	Koustubh W					06/04/2012		3.0			Modified All case by adding "Fn_CM_GetTreeItemPath" function
'	Sandeep N					16/04/2012		4.0			Modified case : MultiSelectPopupMenu
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ViewTreeOperations(strAction,StrNodePath,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_ViewTreeOperations"
   'Variable declaration
   Dim arrNodeName,ItemCount,iCount,sNode, aMenuList,intCount
   Dim intNodeCount,sTreeItem,NodeName,StrMultiNodePath,arrMultiNodeName,iCnt
   Dim ObjChMgrWnd,iPath
	Fn_CM_ViewTreeOperations = False
	'Creating Object of ChangeManager window
	Set ObjChMgrWnd=Fn_UI_ObjectCreate("Fn_CM_ViewTreeOperations", JavaWindow("ChangeManager"))
	
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
			iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),sNode,"","")
			If iPath <> False Then ObjChMgrWnd.JavaTree("ViewTree").Expand iPath
			wait 1
		End IF
	End IF
	
	Select Case strAction
		'===================================================================================================
		Case "Select" 		'Fn_CM_ViewTreeOperations("Select","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
				iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),StrNodePath,"","")
				If iPath<>False Then
					ObjChMgrWnd.JavaTree("ViewTree").Select iPath
					Fn_CM_ViewTreeOperations=True
				End If
		'===================================================================================================
		Case "Expand" 'Fn_CM_ViewTreeOperations("Expand","Change Home:My Open Changes","")
				iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),StrNodePath,"","")
				Call Fn_UI_JavaTree_Expand("Fn_CM_ViewTreeOperations", ObjChMgrWnd, "ViewTree",iPath)
				Fn_CM_ViewTreeOperations=True
		'===================================================================================================
		Case "Collapse"
				iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),StrNodePath,"","")
				Call Fn_UI_JavaTree_Collapse("Fn_CM_ViewTreeOperations", ObjChMgrWnd, "ViewTree",iPath)
				Fn_CM_ViewTreeOperations=True
		'===================================================================================================
		Case "PopupMenuSelect","PopupMenuSelectExt"		'Fn_CM_ViewTreeOperations("PopupMenuSelect","Change Home:Test1","Refresh")
				If strAction = "PopupMenuSelectExt" Then
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CM_ViewTreeOperations",ObjChMgrWnd.JavaTree("ViewTree"),StrNodePath,"","@")
				Else
					iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),StrNodePath,"","")
				End If
				
				If iPath<>False Then
					ObjChMgrWnd.JavaTree("ViewTree").Select iPath
					'Open context menu
					wait 1
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_CM_ViewTreeOperations",ObjChMgrWnd, "ViewTree",iPath)
					Wait 2
					'Select Menu action
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Select Case intCount
						Case "0"
							 StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CM_ViewTreeOperations = False
							Set ObjChMgrWnd=Nothing
							Exit Function
					End Select
					ObjChMgrWnd.WinMenu("ContextMenu").Select StrMenu
					Fn_CM_ViewTreeOperations=True
				End If				
		'===================================================================================================
		Case "VerifyNode"		'Fn_CM_ViewTreeOperations("VerifyNode","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
				iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),StrNodePath,"","")
				If iPath=False Then
					Fn_CM_ViewTreeOperations = FALSE
				Else
					Fn_CM_ViewTreeOperations=True
				End If
		'===================================================================================================
		Case "DoubleClick"
						Call Fn_CM_ViewTreeOperations("Select",StrNodePath,"") ' Modified By Vidya 22/2/2013
						wait 1
						Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
						Fn_CM_ViewTreeOperations=True

''				iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),StrNodePath,"","")  '
''				ObjChMgrWnd.JavaTree("ViewTree").Activate iPath
''				Fn_CM_ViewTreeOperations=True
		'===================================================================================================
		 Case "MultiSelect"
				arrMultiNodeName=split(StrNodePath,",",-1,1)
				For iCount = 0 to uBound(arrMultiNodeName)
					iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),arrMultiNodeName(iCount),"","")
					If iCount = 0 Then
						Fn_CM_ViewTreeOperations = Fn_UI_JavaTree_ExtendSelect("Fn_CM_ViewTreeOperations", ObjChMgrWnd, "ViewTree",iPath)
					Else
						Fn_CM_ViewTreeOperations = Fn_UI_JavaTree_ExtendSelect("Fn_CM_ViewTreeOperations", ObjChMgrWnd, "ViewTree",iPath)
					End If
					If Fn_CM_ViewTreeOperations = False Then exit for
				Next		
		'===================================================================================================
		Case "MultiSelectPopupMenu"		'Fn_CM_ViewTreeOperations("MultiSelectPopupMenu","Change Home:Test1","Refresh")
					arrMultiNodeName=split(StrNodePath,",",-1,1)
					For iCount = 0 to uBound(arrMultiNodeName)
						iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("ViewTree"),arrMultiNodeName(iCount),"","")
						If iPath = False Then
							Fn_CM_ViewTreeOperations = False 
							exit for
						End If
						If iCount = 0 Then
							ObjChMgrWnd.JavaTree("ViewTree").select iPath
							Fn_CM_ViewTreeOperations = True
						Else
							Fn_CM_ViewTreeOperations = Fn_UI_JavaTree_ExtendSelect("Fn_CM_ViewTreeOperations", ObjChMgrWnd, "ViewTree",iPath)
						End If
						If Fn_CM_ViewTreeOperations = False Then exit for
					Next
					If Fn_CM_ViewTreeOperations <> False Then
						'Open context menu
						Fn_CM_ViewTreeOperations = Fn_UI_JavaTree_OpenContextMenu("Fn_CM_ViewTreeOperations",ObjChMgrWnd, "ViewTree", iPath)
						wait 2
						'Select Menu action
						aMenuList = split(StrMenu, ":",-1,1)
							Select Case Ubound(aMenuList)
								Case 0
									StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
								Case 1
									StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
								Case 2
									StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
								Case Else
									Fn_CM_ViewTreeOperations = False
									Set ObjChMgrWnd=Nothing
									Exit Function
							End Select
						ObjChMgrWnd.WinMenu("ContextMenu").Select StrMenu				
						Fn_CM_ViewTreeOperations=True
					End If
		'====================================Not Yet implemented the new workaround for "GetNodeIndex" Case
		Case "GetNodeIndex"		'Fn_CM_ViewTreeOperations("GetNodeIndex","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
				Fn_CM_ViewTreeOperations = Fn_CM_getJavaTreeIndex(ObjChMgrWnd.JavaTree("ViewTree"), StrNodePath) 
		Case Else
				Fn_CM_ViewTreeOperations=False
	End Select
	'Rleasing Change Manager Window Object
	Set ObjChMgrWnd=Nothing
End Function

'-------------------------------------------------------------------Function Used to Handle Error Dialog-------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_ErrorMessageVerify

'Description			 :	Function Used to Handle Error Dialog

'Parameters			   :	1.sDilogName: Error Dialog Name
										'2.sErrorMessage: Error message
										'3.sButtonName:Button name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Error Dialog Should be appear on srceen

'Examples				:	Fn_CM_ErrorMessageVerify("Delete search folder","OK to delete","OK")
				
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				05/08/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ErrorMessageVerify(sDialogName,sErrorMessage,sButtonName)

		Dim dicErrorInfo
		 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
		 dicErrorInfo.Add "Action", "ErrorMessageVerify"
		 dicErrorInfo.Add "Title", sDialogName
		 dicErrorInfo.Add "Message", sErrorMessage
		 dicErrorInfo.Add "Button", sButtonName    
		 Fn_CM_ErrorMessageVerify = Fn_SISW_CM_ErrorVerify(dicErrorInfo)

End Function
'-------------------------------------------------------------------Function Used to perform operations on { AssignParticipants } Dialog--------------------------------------------------------------
'Function Name		:	Fn_CM_AssignParticipantsOperations

'Description			 :	Function Used to perform operatons on  Trees which are present on { AssignParticipants } Dialog

'Parameters			   :	1.strAction: Action Name
										'2.strParticipntNode:Participants Name
										'3.strOrgNode: Organisation Name
										'4.strPrjNode:Project Node Name
										'5.strProjName:Project Name
										'strMemName:Member Name
										'5strGrpName:Group Name
										'6arrSearch:

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective and Revision of Object should be selected

'Examples				:	Fn_CM_AssignParticipantsOperations("Add","Participants:Change Specialist I","Organization:Engineering:Designer: AutoTest3 (autotest3)","","","","","")
'										Fn_CM_AssignParticipantsOperations("Remove","Participants:Change Specialist I:AutoTest1 (autotest1)-Engineering/Designer","","","","","","")
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				12/08/2010			           1.0																					 Rajesh G
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_AssignParticipantsOperations(strAction,strParticipntNode,strOrgNode,strPrjNode,strProjName,strMemName,strGrpName,arrSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_AssignParticipantsOperations"
   'Variable declaration
	Dim ObjAssgnParticpntDialog,sMenu,objErrorDialog,sAppMsg,sErrMsg
	Dim bFlag,sArr,iCnt,sNode, sArr1, sArr2
	Fn_CM_AssignParticipantsOperations=False
	'bFlag set to False
	bFlag=False
	'verifying Existance of { AssignParticipants } Dialog
	sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "AssignParticipants")
	If Fn_UI_ObjectExist("Fn_CM_AssignParticipantsOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants"))=False and Fn_UI_ObjectExist("Fn_CM_AssignParticipantsOperations",JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("AssignParticipants"))=False Then
		'Invoking { AssignParticipants } Dialog
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
		wait 2
	End If
	
	'Creating Object of { AssignParticipants } Dialog
	If Fn_UI_ObjectExist("Fn_CM_AssignParticipantsOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants")) Then 
		Set ObjAssgnParticpntDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants")
	ElseIf Fn_UI_ObjectExist("Fn_CM_AssignParticipantsOperations",JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("AssignParticipants")) Then
		Set ObjAssgnParticpntDialog=JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("AssignParticipants")
	End If 

	Select Case strAction
		Case "Add" 'Case to Add Assign participants
            If strParticipntNode<>"" Then
				'Selecting Participants
                Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"ParticipantsTree",strParticipntNode)
				wait(2)
			End If
			If strOrgNode<>"" Then
				'Selecting Organisation
				sArr = Split(strOrgNode,":")
				If Ubound(sArr) > 1 Then
					For iCnt = 0 to Ubound(sArr) - 1
						If iCnt = 0 Then
							sNode = sArr(0)
						Else
							sNode = sNode+":"+sArr(iCnt)
						End If
						Call Fn_UI_JavaTree_Expand("", ObjAssgnParticpntDialog, "OrganizationTree", sNode)
						wait(3)
					Next
				End If
				wait(3)
				Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"OrganizationTree",strOrgNode)
				wait(2)
			End If
			If strProjName<>"" Then
				'Selecting Project Team Tab
                Call Fn_UI_JavaTab_Select("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"JTabbedPane", "Project Teams")
                wait 2
				'Verifying Project is Present in List
				bFlag=Fn_UI_ListItemExist("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "ProjectsList",strProjName)
				If bFlag=False Then
                    Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Close")
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:Invalid Project Name"& strProjName &"Pass by user")
					Set ObjAssgnParticpntDialog=Nothing
					Exit Function
				End If
				'Selecting Project from List
                Call Fn_List_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "ProjectsList",strProjName)
                wait 2
			End If

			If strPrjNode<>"" Then
				'Selecting Project
				'[TC1123-20160915a00-27_09_2016-VivekA-Maintenance] - As per Akshay's mail, and Design change
				If Instr(strPrjNode,"/")>0 Then
					sArr1 = Split(strPrjNode,"/")
					strPrjNode = strPrjNode +" "+ "("+lcase(sArr1(2))+")"
				End If
				'-----------------------------------------------------
				Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"ProjectsTree",strPrjNode)
				wait(2)
			End If
		
			If strMemName<>"" Then
                Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog.JavaRadioButton("Members"),"attached text",strMemName)
				'Selecting Member Type
                Call Fn_UI_JavaRadioButton_SetON("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"Members")
			End If
			If strGrpName<>"" Then
                Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog.JavaRadioButton("Group"),"attached text",strGrpName)
				'Selecting Group Type
                Call Fn_UI_JavaRadioButton_SetON("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"Group")
			End If
			'Adding he participants
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Add")
			wait(3)
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Apply")		
            Call Fn_ReadyStatusSync(3)
            '--------------------------------
            'Added Assign Participants Error Message Call
            
            Set objErrorDialog=JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("AssignParticipants").JavaDialog("AddParticipantsError")
            If objErrorDialog.Exist(2)=True Then
            bReturn=False
			Set descStaticText=Description.Create()
			descStaticText("Class Name").value="JavaEdit"
			Set objStaticText=objErrorDialog.ChildObjects(descStaticText)								
			For iCounter=0 to objStaticText.count-1
				sAppMsg=objStaticText(iCounter).getROProperty("text")
			If Instr(1,Lcase(sAppMsg),Lcase(sErrMsg))>0 Then
				Call Fn_SISW_UI_JavaButton_Operations("", "Click", objErrorDialog, "OK")
				Fn_CM_AssignParticipantsOperations=True
			Else
				Fn_CM_AssignParticipantsOperations=False	
			End If
			Next	
			Else
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "OK")
			Fn_CM_AssignParticipantsOperations=True
			End  If
			'Function Return True
			Fn_CM_AssignParticipantsOperations=True
			
		Case "Remove" 'Case to Add Assign participants
            If strParticipntNode<>"" Then
				'Selecting Participants
                Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"ParticipantsTree",strParticipntNode)
                wait(3)
			End If
			'Adding he participants
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Remove")
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Apply")
'			wait(5)
            Call Fn_ReadyStatusSync(5)
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "OK")
			'Function Return True
			Fn_CM_AssignParticipantsOperations=True
			Case "RemoveWithoutClose" 'Case to Add Assign participants
            If strParticipntNode<>"" Then
				'Selecting Participants
                Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"ParticipantsTree",strParticipntNode)
                wait(3)
			End If
			'Adding he participants
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Remove")
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Apply")
'			wait(5)
            Call Fn_ReadyStatusSync(5)
			'Function Return True
			Fn_CM_AssignParticipantsOperations=True

	Case "Modify" 'Case to Modify Assign participants
            If strParticipntNode<>"" Then
				'Selecting Participants
                Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"ParticipantsTree",strParticipntNode)
                wait(3)
			End If
			If strOrgNode<>"" Then
				'Selecting Organisation
				Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"OrganizationTree",strOrgNode)
                wait(3)
			End If
			If strProjName<>"" Then
				'Selecting Project Team Tab
                Call Fn_UI_JavaTab_Select("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"JTabbedPane", "Project Teams")
				wait(3)
				'Verifying Project is Present in List
				bFlag=Fn_UI_ListItemExist("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "ProjectsList",strProjName)
				If bFlag=False Then
                    Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Close")
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:Invalid Project Name"& strProjName &"Pass by user")
					Set ObjAssgnParticpntDialog=Nothing
					Exit Function
				End If
				'Selecting Project from List
                Call Fn_List_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "ProjectsList",strProjName)
				wait(3)
			End If
			If strPrjNode<>"" Then
				'Selecting Project
				'[TC1123-20160915a00-27_09_2016-VivekA-Maintenance] - As per Akshay's mail, and Design change
				If Instr(strPrjNode,"/")>0 Then
					sArr2 = Split(strPrjNode,"/")
					strPrjNode = strPrjNode + " "+"("+lcase(sArr2(2))+")"
				End If
				'--------------------------------------------------
				Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"ProjectsTree",strPrjNode)
                wait(3)
			End If
			If strMemName<>"" Then
                Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"attached text",strMemName)
				'Selecting Member Type
                Call Fn_UI_JavaRadioButton_SetON("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"Members")
			End If
			If strGrpName<>"" Then
                Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"attached text",strGrpName)
				'Selecting Group Type
                Call Fn_UI_JavaRadioButton_SetON("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"Group")
			End If
			'Adding he participants
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Modify")
			wait(3)
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Apply")
			wait(3)
'			wait(5)
           Call Fn_ReadyStatusSync(5)
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "OK")
			wait(3)
			'Function Return True
			Fn_CM_AssignParticipantsOperations=True
			
		Case "AddParticipantsErrorMessageVerify" 'Case to verify error while Assigning participants
				If strParticipntNode<>"" Then
				'Selecting Participants
                Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"ParticipantsTree",strParticipntNode)
				wait(2)
			End If
			If strOrgNode<>"" Then
				'Selecting Organisation
				sArr = Split(strOrgNode,":")
				If Ubound(sArr) > 1 Then
					For iCnt = 0 to Ubound(sArr) - 1
						If iCnt = 0 Then
							sNode = sArr(0)
						Else
							sNode = sNode+":"+sArr(iCnt)
						End If
						Call Fn_UI_JavaTree_Expand("", ObjAssgnParticpntDialog, "OrganizationTree", sNode)
						wait(3)
					Next
				End If
				wait(3)
				Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"OrganizationTree",strOrgNode)
				wait(2)
			End If
			If strProjName<>"" Then
				'Selecting Project Team Tab
                Call Fn_UI_JavaTab_Select("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"JTabbedPane", "Project Teams")
                wait 2
				'Verifying Project is Present in List
				bFlag=Fn_UI_ListItemExist("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "ProjectsList",strProjName)
				If bFlag=False Then
                    Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Close")
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail:Invalid Project Name"& strProjName &"Pass by user")
					Set ObjAssgnParticpntDialog=Nothing
					Exit Function
				End If
				'Selecting Project from List
                Call Fn_List_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "ProjectsList",strProjName)
                wait 2
			End If

			If strPrjNode<>"" Then
				'Selecting Project
				'[TC1123-20160915a00-27_09_2016-VivekA-Maintenance] - As per Akshay's mail, and Design change
				If Instr(strPrjNode,"/")>0 Then
					sArr1 = Split(strPrjNode,"/")
					strPrjNode = strPrjNode +" "+ "("+lcase(sArr1(2))+")"
				End If
				'-----------------------------------------------------
				Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog,"ProjectsTree",strPrjNode)
				wait(2)
			End If
		
			If strMemName<>"" Then
                Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog.JavaRadioButton("Members"),"attached text",strMemName)
				'Selecting Member Type
                Call Fn_UI_JavaRadioButton_SetON("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"Members")
			End If
			If strGrpName<>"" Then
                Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog.JavaRadioButton("Group"),"attached text",strGrpName)
				'Selecting Group Type
                Call Fn_UI_JavaRadioButton_SetON("Fn_CM_AssignParticipantsOperations",ObjAssgnParticpntDialog,"Group")
			End If
			'Adding he participants
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Add")
			wait(3)
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Apply")		
            Call Fn_ReadyStatusSync(3)
            
            'Verify Error Message
            Set objErrorDialog = Fn_SISW_CM_GetObject("AddParticipantsErrorNew")
            sErrMsg=Fn_GetXMLNodeValue(Environment.Value("sPath")&"\TestData\AutomationXML\ErrorMessageXML\ChangeManager.xml","AddParticipantsError")
            
            sAppMsg=Fn_UI_Object_GetROProperty("Fn_CM_AssignParticipantsOperations",objErrorDialog.JavaEdit("ErrTextArea"), "text")
            If Instr(1,Lcase(sAppMsg),Lcase(sErrMsg))>0 Then
					Call Fn_SISW_UI_JavaButton_Operations("", "Click", objErrorDialog, "OK")
					Fn_CM_AssignParticipantsOperations=True
			Else
					Fn_CM_AssignParticipantsOperations= False	
			End If
			Call Fn_Button_Click("Fn_CM_AssignParticipantsOperations", ObjAssgnParticpntDialog, "Cancel")	
			Call Fn_ReadyStatusSync(3)	
			
			
			
            End Select
	Set ObjAssgnParticpntDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to perform operatons on  Trees which are present on { AssignParticipants } Dialog----------------------------------------------------------------
'Function Name		:	Fn_CM_AssignParticipantsTreeOperations

'Description			 :	Function Used to perform operatons on  Trees which are present on { AssignParticipants } Dialog

'Parameters			   :	1.strAction: Action Name
										'2.strTreeName:Tree Name On which have to perform operation
										'3.strNode: Node Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective

'Examples				:	Fn_CM_AssignParticipantsTreeOperations("Select","ParticipantsTree","Participants:Requestor:Sandeep Navghane (x_navgha)-Engineering/Designer")
										'Fn_CM_AssignParticipantsTreeOperations("Expand","ParticipantsTree","Participants:Requestor")
										'Fn_CM_AssignParticipantsTreeOperations("Verify","OrganizationTree","Participants:Requestor")
										'Fn_CM_AssignParticipantsTreeOperations("Select","ProjectsTree","ProjectsTree:Requestor:Sandeep Navghane (x_navgha)-Engineering/Designer")

										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				12/08/2010			           1.0																					 Rajesh G
'													Sandeep N										   				13/09/2011			           1.1																					 Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_AssignParticipantsTreeOperations(strAction,strTreeName,strNode)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_AssignParticipantsTreeOperations"
   'Variable Declarion
	Dim ObjAssgnParticpntDialog
	Dim intNodeCount,sTreeItem,intCount,sMenu
	sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "AssignParticipants")
	'Verifying Existance Of { AssignParticipants } Dialog
	
	If Not JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants").Exist(SISW_MIN_TIMEOUT) And Not JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("AssignParticipants").Exist(SISW_MIN_TIMEOUT) Then
		'Invoking { AssignParticipants } Dialog
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
	End If
	
	'Creating Object of { AssignParticipants } Dialog
	If Fn_UI_ObjectExist("Fn_CM_AssignParticipantsOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants")) Then 
		Set ObjAssgnParticpntDialog=JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants")
	ElseIf Fn_UI_ObjectExist("Fn_CM_AssignParticipantsOperations",JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("AssignParticipants")) Then
		Set ObjAssgnParticpntDialog=JavaWindow("DefaultWindow").JavaWindow("DefaultEmbeddedFrame").JavaDialog("AssignParticipants")
	End If 
	
	If  strTreeName="ProjectsTree" Then
		'Selecting Project Team Tab
		Call Fn_UI_JavaTab_Select("Fn_CM_AssignParticipantsTreeOperations",ObjAssgnParticpntDialog,"JTabbedPane", "Project Teams")
	End If
	
	Select Case strAction
		Case "Select" 'Fn_CM_AssignParticipantsTreeOperations("Select","ParticipantsTree","Participants:Requestor:Sandeep Navghane (x_navgha)-Engineering/Designer")
            Call Fn_JavaTree_Select("Fn_CM_AssignParticipantsTreeOperations",ObjAssgnParticpntDialog,strTreeName,strNode)
			Fn_CM_AssignParticipantsTreeOperations=True
		Case "Expand"	''Fn_CM_AssignParticipantsTreeOperations("Expand","ParticipantsTree","Participants:Requestor")
			'Expanding The Node
			wait(1)
            Call Fn_UI_JavaTree_Expand("Fn_CM_AssignParticipantsTreeOperations",ObjAssgnParticpntDialog,strTreeName,strNode)
			Fn_CM_AssignParticipantsTreeOperations=True
		Case "Verify" 'Fn_CM_AssignParticipantsTreeOperations("Verify","OrganizationTree","Participants:Requestor")
            intNodeCount = Fn_UI_Object_GetROProperty("Fn_CM_AssignParticipantsTreeOperations",ObjAssgnParticpntDialog.JavaTree(strTreeName), "items count")
			For intCount = 0 to intNodeCount - 1
				sTreeItem = ObjAssgnParticpntDialog.JavaTree(strTreeName).GetItem(intCount)
				If Trim(lcase(sTreeItem)) = Trim(Lcase(strNode)) Then
					Fn_CM_AssignParticipantsTreeOperations = True
					Exit For
				End If
			Next
			If Cint(intCount) = Cint(intNodeCount) Then
				Fn_CM_AssignParticipantsTreeOperations = False
			End If
			'Closing The Dialog
			If ObjAssgnParticpntDialog.JavaButton("Cancel").Exist(5) Then
				 Call Fn_Button_Click("Fn_CM_AssignParticipantsTreeOperations", ObjAssgnParticpntDialog, "Cancel")				
			ElseIf ObjAssgnParticpntDialog.JavaButton("Close").Exist(2) Then
				 Call Fn_Button_Click("Fn_CM_AssignParticipantsTreeOperations", ObjAssgnParticpntDialog, "Close")
			End If
	End Select
	'Releasing Object of { AssignParticipants } Dialog
	Set ObjAssgnParticpntDialog=Nothing
End Function

'-------------------------------------------------------------------Function Used to perform operatons on static Text of  { AssignParticipants } Dialog------------=-----------------------------------------------------------------
'Function Name		:	Fn_CM_AssignParticipantsStaticTextOperations

'Description			 :	Function Used to perform operatons on static Text of  { AssignParticipants } Dialog

'Parameters			   :	1.strAction: Action Name
										'2.strStaticText:Static Text on Wich have to perform operation
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective

'Examples				:	Fn_CM_AssignParticipantsStaticTextOperations("Exist","Assigning participants is not allowed...")
									   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				16/08/2010			           1.0																					 Rajesh G
'													Sandeep N										   				21/09/2011			           1.1					Added Cancel button call					  Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_AssignParticipantsStaticTextOperations(strAction,strStaticText)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_AssignParticipantsStaticTextOperations"
   'Variable Declarion
	Dim ObjAssgnParticpntDialog,ObjJavaStat,ObjAssgnParChld
	Dim iCount,bFlag,sMenu
	bFlag=False
	Fn_CM_AssignParticipantsStaticTextOperations=False
	sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "AssignParticipants")
	'Verifying Existance Of { AssignParticipants } Dialog
	If Fn_UI_ObjectExist("Fn_CM_AssignParticipantsStaticTextOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants"))=False Then
		'Invoking { AssignParticipants } Dialog
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(2)
	End If
	
	'Creating Object of { AssignParticipants } Dialog
	Set ObjAssgnParticpntDialog=Fn_UI_ObjectCreate("Fn_CM_AssignParticipantsStaticTextOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants"))
	Select Case strAction
		Case "Exist" 'Fn_CM_AssignParticipantsStaticTextOperations("Select","ParticipantsTree","Participants:Requestor:Sandeep Navghane (x_navgha)-Engineering/Designer")
            Set ObjJavaStat=Description.Create()
			ObjJavaStat("Class Name").value="JavaStaticText"
			Set ObjAssgnParChld=ObjAssgnParticpntDialog.ChildObjects(ObjJavaStat)
			For iCount=0 To ObjAssgnParChld.Count-1
				If Instr(1,ObjAssgnParChld(iCount).ToString,strStaticText)>=1 Then
					bFlag=True
				End If
			Next
			If bFlag=True Then
				Fn_CM_AssignParticipantsStaticTextOperations=True
			End If
			'Closing The Dialog
			If ObjAssgnParticpntDialog.JavaButton("Cancel").Exist(5) Then
				 Call Fn_Button_Click("Fn_CM_AssignParticipantsStaticTextOperations", ObjAssgnParticpntDialog, "Cancel")				
			ElseIf ObjAssgnParticpntDialog.JavaButton("Close").Exist(2) Then
				 Call Fn_Button_Click("Fn_CM_AssignParticipantsStaticTextOperations", ObjAssgnParticpntDialog, "Close")
			End If
			
	End Select
	'Releasing Object of { AssignParticipants } Dialog
	Set ObjAssgnParticpntDialog=Nothing
	Set ObjJavaStat=Nothing
	Set ObjAssgnParChld=Nothing
End Function

'-------------------------------------------------------------------Function Used to perform operations on { AssignParticipants } Dialog Controls--------------------------------------------------------------
'Function Name		:	Fn_CM_AssignParticipantsGUIOprts

'Description			 :	Function Used to perform operations on { AssignParticipants } Dialog Controls

'Parameters			   :	1.strAction: Action Name
										'2.arrAssgnParticipants:Array of parameters


'Return Value		   : 	True Or False Or Value

'Pre-requisite			:	Assign Participants Dialog Should Be Open

'Examples				:	arrAssgnParticipants=Array(StrTreeName,StrNodeName,StrTabName,StrRadioaBtnName,StrButtonName)
'										arrAssgnParticipants=Array("Participants","Participants:Analyst","Organisazion","Any Member","Add")
'										Fn_CM_AssignParticipantsGUIOprts("VerifyNode",arrAssgnParticipants)
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				12/08/2010			           1.0																					 Rajesh G
'                                                    Snehal Salunkhe                                             02/01/2012                     1.1                          Changed close case						Prasanna				 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'arrAssgnParticipants=Array(StrTreeName,StrNodeName,StrTabName,StrRadioaBtnName,StrButtonName)
'arrAssgnParticipants=Array("Participants","Participants:Analyst","Organisazion","Any Member","Add")
Public Function Fn_CM_AssignParticipantsGUIOprts(SrtAction,arrAssgnParticipants)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_AssignParticipantsGUIOprts"
   Dim bFlag,iTabCnt,arrTabName(1),iCounter,intNodeCount,intCount,sTreeItem
   Dim ObjAssignParticipants
   Fn_CM_AssignParticipantsGUIOprts=False
   Set ObjAssignParticipants=Fn_UI_ObjectCreate("Fn_CM_AssignParticipantsGUIOprts",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("AssignParticipants"))
   Select Case SrtAction
			 	Case "VerifyNode"
						    intNodeCount = Fn_UI_Object_GetROProperty("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants.JavaTree(arrAssgnParticipants(0)), "items count")
							For intCount = 0 to intNodeCount - 1
								sTreeItem = ObjAssignParticipants.JavaTree(arrAssgnParticipants(0)).GetItem(intCount)
								If Trim(lcase(sTreeItem)) = Trim(Lcase(arrAssgnParticipants(1))) Then
									Fn_CM_AssignParticipantsGUIOprts = True
									Exit For
								End If
							Next
							If Cint(intCount) = Cint(intNodeCount) Then
								Fn_CM_AssignParticipantsGUIOprts = False
							End If					                       
				Case "ButtonExist"
                    bFlag=Fn_UI_ObjectExist("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants.JavaButton(arrAssgnParticipants(4)))
					If bFlag=True Then
						Fn_CM_AssignParticipantsGUIOprts=True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully verified Button" & arrAssgnParticipants(4) & " is is exist on Dialog")
					End If
				Case "ResourcePoolOptnValue"
                    Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants.JavaRadioButton("Members"),"attached text",arrAssgnParticipants(3))
					Fn_CM_AssignParticipantsGUIOprts=Fn_UI_Object_GetROProperty("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants.JavaRadioButton("Members"),"value")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Current value of " & arrAssgnParticipants(3) & " radio button is" & Cstr(Fn_CM_AssignParticipantsGUIOprts))
				Case "TabCount"
					Fn_CM_AssignParticipantsGUIOprts=Fn_UI_Object_GetROProperty("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants.JavaTab("JTabbedPane"),"items count")
				Case "SelectTab"
					bFlag=Fn_UI_JavaTab_Select("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants,"JTabbedPane",arrAssgnParticipants(2))
					If bFlag=True Then
						Fn_CM_AssignParticipantsGUIOprts=True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Select tab " & arrAssgnParticipants(2))
					End If
				 Case "Close"
                        'Call Fn_Button_Click("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants,"Close")
						Call Fn_Button_Click("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants,"Cancel")  ' Added for teamcenter changes Tc 9.1-1214 onwards. Close button changed to Cancel
						Fn_CM_AssignParticipantsGUIOprts=True
				 Case "TabName"
						iTabCnt=Fn_UI_Object_GetROProperty("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants.JavaTab("JTabbedPane"),"items count")
						For iCounter=0 To iTabCnt-1							
								ObjAssignParticipants.JavaTab("JTabbedPane").Select(iCounter)
								arrTabName(iCounter)=Fn_UI_Object_GetROProperty("Fn_CM_AssignParticipantsGUIOprts",ObjAssignParticipants.JavaTab("JTabbedPane"),"value")
						Next
						Fn_CM_AssignParticipantsGUIOprts=arrTabName
				Case Else
						Fn_CM_AssignParticipantsGUIOprts=False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "wrong Action Name "& SrtAction & "pass by user")
   End Select
   Set ObjAssignParticipants=Nothing
End Function


'-------------------------------------------------------------------Function Used to perform operatons on  Change Type Tree Which is present on New Change Dialog---------------------------------------------------
'Function Name		:	Fn_CM_ChangeTypeTreeOperations

'Description			 :	Function Used to perform operatons on  Change Type Tree Which is present on New Change Dialog

'Parameters			   :	1.strAction: Action Name
										'2.strTreeName:Tree Name On which have to perform operation
										'3.strNode: Node Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on MyTeamcenter Perspective

'Examples				:	bReturn=Fn_CM_ChangeTypeTreeOperations("VerifyNode","Complete List:Problem Report","")
'										bReturn=Fn_CM_ChangeTypeTreeOperations("VerifyNode","Complete List:Change Request","")
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				25/08/2010			           1.0																					 Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ChangeTypeTreeOperations(strAction,strNodeName,StrMenu)
GBL_FAILED_FUNCTION_NAME="Fn_CM_ChangeTypeTreeOperations"
   'Veriable Declaration
   Dim sTreeItem,intCount,intNodeCount
   Dim ObjChangeWnd
   Fn_CM_ChangeTypeTreeOperations=False
	If Fn_UI_ObjectExist("Fn_CM_ChangeTypeTreeOperations",JavaWindow("DefaultWindow").JavaWindow("New Change"))=False Then
		'Invoking "New Change" Window
		Call Fn_MenuOperation("Select","File:New:Change...")
	End If
    'Creating Object of "New Change" window
	Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_CM_ChangeTypeTreeOperations",JavaWindow("DefaultWindow").JavaWindow("New Change"))
	Call Fn_UI_JavaTree_Expand("Fn_CM_ChangeTypeTreeOperations", ObjChangeWnd, "ChangeTypeTree","Complete List")
	Select Case strAction
		Case "VerifyNode"
		    intNodeCount =Fn_UI_Object_GetROProperty("Fn_CM_ChangeTypeTreeOperations",ObjChangeWnd.JavaTree("ChangeTypeTree"),"items count")    
			For intCount = 0 to intNodeCount - 1
				sTreeItem = ObjChangeWnd.JavaTree("ChangeTypeTree").GetItem(intCount)
				If Trim(lcase(sTreeItem)) = Trim(Lcase(strNodeName)) Then
					Fn_CM_ChangeTypeTreeOperations = True
					Exit For
				End If
			Next
	End Select
	Set ObjChangeWnd=Nothing
End Function

'-------------------------------------------------------------------Function Used to Handle Error Dialog which is come at time of PR,ECR,ECN creation----------------------------------------------------------------------------
'Function Name		:	Fn_CM_ChangeErrorMsgVerify

'Description			 :	Function Used to Handle Error Dialog which is come at time of PR,ECR,ECN creation

'Parameters			   :	1.strAction:Action Name
										'2.sDilogName: Error Dialog Name
										'3.sErrorMessage: Error message
										'4.sButtonName:Button name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Error Dialog Should be appear on srceen

'Examples				:	bReturn=Fn_CM_ChangeErrorMsgVerify("EditBoxMsgVerify","Information","Supplied item_id value","OK")
'										bReturn=Fn_CM_ChangeErrorMsgVerify("StaticMsgVerify","Information","Details","OK")
				
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				25/08/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ChangeErrorMsgVerify(strAction,sDialogName,sErrorMessage,sButtonName)

	Dim dicErrorInfo
	 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	 dicErrorInfo.Add "Action", strAction
	 dicErrorInfo.Add "Title", sDialogName
	 dicErrorInfo.Add "Message", sErrorMessage
	 dicErrorInfo.Add "Button", sButtonName    
	 Fn_CM_ChangeErrorMsgVerify = Fn_SISW_CM_ErrorVerify(dicErrorInfo)


End Function

'-------------------------------------------------------------------Function Used to perform operatons on Summery Tab------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_SummaryOperation

'Description			 :	Function Used to perform operatons on Summery Tab

'Parameters			   :	1.sAction: Action Name
										'2.sObjectName:Object Name		
										'3.sObjectExpectedValue: Expected value of object
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Summery Tab Should be present

'Examples				:	Fn_CM_SummaryOperation ("Analyst","Analyst" ,"Engineering/Designer/AutoTest1","" , "")
									'	Fn_CM_SummaryOperation ("Requestor","Requestor" ,"Engineering/Designer/AutoTest1","" , "")	 
									'	Fn_CM_SummaryOperation ("Specialist","Change Specialist I" ,"Engineering/Designer/AutoTest1","" , "")  
									'	Fn_CM_SummaryOperation ("CheckOut","Checked-Out: ,"","" , "") 
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				26/08/2010			           1.0																						Tushar B
'													Sandeep N										   				02/01/2011			           1.1																						Sandeep N
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_SummaryOperation (sAction, sObjectName, sObjectExpectedValue, sAttachedObjectName, sSendTo)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_SummaryOperation"
   Dim bFlag,strObjExpVal,strObjName, sArr_Specialist,sObjectExpectedValue1
   Fn_CM_SummaryOperation=False
    bFlag=False

	'This Temporary Code added for refresh window Working on Build (20101208) by Rupali [23-Dec-2010]
'	Call Fn_MenuOperation("Select","View:Refresh")
'	wait(1)
'	Call Fn_ReadyStatusSync(2)

   Select Case sAction
   Case "Analyst"
		'Swapnil Added the call to select the summary Tab in cases Analyst ,Requestor,Specialist
		Call  Fn_MyTc_TabSet("Summary")
		 Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text"),"label","Analyst:")
		 strObjExpVal=Fn_UI_Object_GetROProperty("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaObject("Summary_Object"), "text")
		  '[TC1123-20160915a00-27_09_2016-VivekA-Maintenance] - As per Akshay's mail, and Design change
		 If Instr(strObjExpVal,"(") > 0 AND Instr(strObjExpVal,")") > 0 Then
		 	If Instr(sObjectExpectedValue,"/")>0 Then
				sArr_Specialist = Split(sObjectExpectedValue,"/")
				sObjectExpectedValue1 = sObjectExpectedValue +" "+ "("+lcase(sArr_Specialist(2))+")"
				If Trim(strObjExpVal)=Trim(sObjectExpectedValue1) Then
					sObjectExpectedValue = sObjectExpectedValue1
				End If
			End If
		 End If		
		 '-----------------------------------------------------
		 If Trim(lcase(strObjExpVal))=Trim(lcase(sObjectExpectedValue)) Then
			 bFlag=True
		End If
		If bFlag=True Then
				Fn_CM_SummaryOperation=True
		End If
		JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text").SetTOProperty "label",""	
	
   	Case "Requestor"
		Call  Fn_MyTc_TabSet("Summary")
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text"),"label","Requestor:")
		 strObjExpVal=Fn_UI_Object_GetROProperty("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaObject("Summary_Object"), "text")
		  '[TC1123-20160915a00-27_09_2016-VivekA-Maintenance] - As per Akshay's mail, and Design change
		 If Instr(strObjExpVal,"(") > 0 AND Instr(strObjExpVal,")") > 0 Then
		 	If Instr(sObjectExpectedValue,"/")>0 Then
				sArr_Specialist = Split(sObjectExpectedValue,"/")
				sObjectExpectedValue1 = sObjectExpectedValue +" "+ "("+lcase(sArr_Specialist(2))+")"
				If Trim(strObjExpVal)=Trim(sObjectExpectedValue1) Then
					sObjectExpectedValue = sObjectExpectedValue1
				End If
			End If
		 End If		
		 '-----------------------------------------------------
		 If Trim(strObjExpVal)=Trim(sObjectExpectedValue) Then
			 bFlag=True
		End If
    	If bFlag=True Then
			Fn_CM_SummaryOperation=True
		End If
		JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text").SetTOProperty "label",""

	Case "Specialist"
		Call  Fn_MyTc_TabSet("Summary")
		 Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text"),"label","Change Specialist I:")
		 strObjExpVal=Fn_UI_Object_GetROProperty("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaObject("Summary_Object"), "text")
		 '[TC1123-20160915a00-27_09_2016-VivekA-Maintenance] - As per Akshay's mail, and Design change
		 If Instr(strObjExpVal,"(") > 0 AND Instr(strObjExpVal,")") > 0 Then
		 	If Instr(sObjectExpectedValue,"/")>0 Then
				sArr_Specialist = Split(sObjectExpectedValue,"/")
				sObjectExpectedValue1 = sObjectExpectedValue +" "+ "("+lcase(sArr_Specialist(2))+")"
				If Trim(strObjExpVal)=Trim(sObjectExpectedValue1) Then
					sObjectExpectedValue = sObjectExpectedValue1
				End If
			End If
		 End If		
		'-----------------------------------------------------
		 If Trim(strObjExpVal)=Trim(sObjectExpectedValue) Then
			 bFlag=True
		End If
    	If bFlag=True Then
				Fn_CM_SummaryOperation=True
		End If
		JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text").SetTOProperty "label",""
	
	Case "ChangeImplementationBoard"
			 JavaWindow("MyTeamcenter").JavaLink("SummaryType").SetTOProperty "attached text","Change Implementation Board:"
			 strObjExpVal=JavaWindow("MyTeamcenter").JavaLink("SummaryType").GetROProperty("text")
			 If strObjExpVal=sObjectExpectedValue Then
				 bFlag=True
			End If
			 strObjName=JavaWindow("MyTeamcenter").JavaLink("SummaryType").GetROProperty("attached text")
			If strObjName=sObjectName+":" Then
				If bFlag=True Then
					Fn_CM_SummaryOperation=True
				End If
			End If

    Case "ChangeImplementationBoard_Edit"
             Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text"),"label","Change Implementation Board:")

			 	If JavaWindow("MyTeamcenter").JavaEdit("Summary_Text").Exist(5) Then
				strObjExpVal=Fn_UI_Object_GetROProperty("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaEdit("Summary_Text"), "value")
				If Trim(strObjExpVal)=Trim(sObjectExpectedValue) Then
					bFlag=True
				End If
			ElseIf JavaWindow("MyTeamcenter").JavaList("Summary_List").Exist(5)  Then
				strObjExpVal=Split(sObjectExpectedValue,"~")
				For iCounter=0 To UBound(strObjExpVal)
					bFlag=Fn_UI_ListItemExist("Fn_CM_SummaryOperation", JavaWindow("MyTeamcenter"), "Summary_List",strObjExpVal(iCounter))
					If bFlag=False Then
						Exit For
					End If
				Next
			End If
			If bFlag=True Then
					Fn_CM_SummaryOperation=True
			End If
			JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text").SetTOProperty "label",""
	
	Case "ChangeReviewBoard"
		JavaWindow("MyTeamcenter").JavaLink("SummaryType").SetTOProperty "attached text","Change Review Board:"
		 strObjExpVal=JavaWindow("MyTeamcenter").JavaLink("SummaryType").GetROProperty("value")
		 If strObjExpVal=sObjectExpectedValue Then
			 bFlag=True
		End If
		 strObjName=JavaWindow("MyTeamcenter").JavaLink("SummaryType").GetROProperty("attached text")
		If strObjName=sObjectName+":" Then
			If bFlag=True Then
				Fn_CM_SummaryOperation=True
			End If
		End If

	Case "ChangeReviewBoard_JavaEdit"
			Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text"),"label","Change Review Board:")
			If JavaWindow("MyTeamcenter").JavaEdit("Summary_Text").Exist(5) Then
				strObjExpVal=Fn_UI_Object_GetROProperty("Fn_CM_SummaryOperation",JavaWindow("MyTeamcenter").JavaEdit("Summary_Text"), "value")
				If Trim(strObjExpVal)=Trim(sObjectExpectedValue) Then
					bFlag=True
				End If
			ElseIf JavaWindow("MyTeamcenter").JavaList("Summary_List").Exist(5)  Then
				strObjExpVal=Split(sObjectExpectedValue,"~")
				For iCounter=0 To UBound(strObjExpVal)
					bFlag=Fn_UI_ListItemExist("Fn_CM_SummaryOperation", JavaWindow("MyTeamcenter"), "Summary_List",strObjExpVal(iCounter))
					If bFlag=False Then
						Exit For
					End If
				Next
			End If
            		
			If bFlag=True Then
					Fn_CM_SummaryOperation=True
			End If
			JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text").SetTOProperty "label",""

	Case "CheckOut"

		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_SummaryOperation",JavaWindow("ChangeManager").JavaStaticText("Summary_Text"),"label",sObjectName)
		Fn_CM_SummaryOperation=Fn_UI_Object_GetROProperty("Fn_CM_SummaryOperation",JavaWindow("ChangeManager").JavaEdit("Summary_Text"), "value")
		JavaWindow("MyTeamcenter").JavaStaticText("Summary_Text").SetTOProperty "label",""

   End Select
End Function 
'-------------------------------------------------------------------Function Used to Genarate 6 Digit Random Number------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_RandNoGenerate

'Description			 :	Function Used to Genarate 6 Digit Random Number
										
'Return Value		   : 	Random Number

'Examples				:	Fn_CM_RandNoGenerate

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Amol L										   						03/09/2010						1.0																							tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_RandNoGenerate()
'	 Dim iNumber,iScnd,iNum,iHr,iDay
'	 Randomize
'	 iScnd=Second(Now)
'	iHr=Hour(Now)
'	iDay=Day(Now)
'	iNum=987123+iScnd+iHr+iDay
'	 iNumber = Int((iNum * Rnd) + 1)
'	 If Len(Cstr(iNumber)) < 5 Then
'			Fn_CM_RandNoGenerate = "0" + Cstr(iNumber)+"6"
'	 ElseIf Len(Cstr(iNumber)) < 6 Then
'			Fn_CM_RandNoGenerate = "0" + Cstr(iNumber)
'	 ElseIf Len(Cstr(iNumber)) > 6 Then
'			iNumber = Int((900000 * Rnd) + 1)
'				If Len(Cstr(iNumber)) < 5 Then
'					Fn_CM_RandNoGenerate = "0" + Cstr(iNumber)+"6"
'				ElseIf Len(Cstr(iNumber)) < 6 Then
'					Fn_CM_RandNoGenerate = "0" + Cstr(iNumber)
'				Else
'					Fn_CM_RandNoGenerate = Cstr(iNumber)
'				End If
'	 Else
'			Fn_CM_RandNoGenerate = Cstr(iNumber)
'	 End If

	Fn_CM_RandNoGenerate = Cstr(Fn_Setup_RandNoGenerate(6))

End Function
'*********************************************************		Function to Perform Operations on Tab into Teamcenter*****************************************************
'Function Name		:				Fn_MyTc_TabOperation

'Description			 :		 		 This Tab includes summary Tab,Details tab,Viewer Tab,Impact analysis tab ,BOM Changes,Change Effectivity
'													1)Existance of Tab
'                                                   
'Parameters			   :	 			1) sAction: Action to be performed on the Tab
'													 2) sTabName: Tab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Change Manager application should be displayed.

'Examples				:				 Fn_CM_TabOperation("Exist","Impact Analysis")
'													 Fn_CM_TabOperation("Exist","Summary:Details:Viewer:Impact Analysis:BOM Changes:Change Effectivity")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N					03-9-10				1.0			No							Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N					17-04-12				1.0			Added Case ActivateExt						
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_CM_TabOperation(sAction, sTabName)  
	GBL_FAILED_FUNCTION_NAME="Fn_CM_TabOperation"
	Dim iTabCount,iTabIndex,sTabVal,arrTab,iCount,bFlag,iCounter,iItemCount
	Dim iCnt,sxLen,i,objItem,objJavaTab,objSelectType,objIntNoOfObjects
	'Set objJavaTab =Fn_UI_ObjectCreate("Fn_CM_TabOperation", JavaWindow("ChangeManager").JavaObject("MyTcTabObject"))
	Set objJavaTab = JavaWindow("ChangeManager").JavaTab("Tab")
	bFlag=False
    Select Case sAction
                Case  "Exist" 
							arrTab=Split(sTabName,":")
							For iCount=0 To Ubound(arrTab)
								'Counting Number of Tabs 
								 iTabCount = Fn_UI_Object_GetROProperty("Fn_CM_TabOperation",objJavaTab,"items count")
								For iTabIndex = 0 to iTabCount-1
									bFlag=False
									objJavaTab.Select "#"&iTabIndex
									wait(2)
									sTabVal= Fn_UI_Object_GetROProperty("Fn_CM_TabOperation",objJavaTab,"value")
									If sTabVal = arrTab(iCount) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTabName+ "] : Tab is Exist")
										Fn_CM_TabOperation=True
										Exit For
									Else
										bFlag=True
									End If									
								Next     
								If bFlag=True Then
									Fn_CM_TabOperation = False
									Exit For
								End If
							Next
				 Case "VerifyActivate"
								'Check Weather Requested tab is open or not( Activated or Not )
                              	If  sTabName = JavaWindow("ChangeManager").JavaTab("Tab").GetROProperty("value")   Then							
										'Call Log file
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTabName+ "] : Tab is Activate(Open) ")
										Fn_CM_TabOperation = TRUE
								Else							
									'Call Log file
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [" +sTabName+ "] : Tab is NOT Activate (Open)")	
									Fn_CM_TabOperation = FALSE
								End If
				Case "Activate" 
							Call Fn_UI_JavaTab_Select("Fn_CM_TabOperation",  JavaWindow("ChangeManager"), "Tab", sTabName)
							Fn_CM_TabOperation = True

				Case "ActivateExt"  'Added by Sandip N
'							iCnt = objJavaTab.Object.getTabItemCount
'							sxLen = 0
'							For i = 0 to iCnt-1
'								set objItem = objJavaTab.Object.getItem(i)
'								'print objItem.text
'								sxLen = sxLen + objItem.getWidth
'								If trim(objItem.text) = trim(sTabName) Then
'                					sxLen = sxLen - (objItem.getWidth)
'									objJavaTab.Click sxLen, (objItem.getHeight/2), "LEFT"
'									Exit For
'								End If
'							Next
'							If err.number < 0 Then
'									Fn_CM_TabOperation = False
'							Else
'									Fn_CM_TabOperation = True
'							End If
'							Set objItem = Nothing
							'TC 12 - 2017080200 - 21/Aug/17 - NishigandhaJ - Modified code as object changed in application	
							bFlag=False
							Set objSelectType = description.Create()
							objSelectType("Class Name").value = "JavaTab"
							objSelectType("toolkit class").value = "org.eclipse.swt.custom.CTabFolder"						
							Set  objIntNoOfObjects = JavaWindow("DefaultWindow").ChildObjects(objSelectType)
					
							For icount = 0 To objIntNoOfObjects.Count-1 Step 1			
								iItemCount = cInt(objIntNoOfObjects(icount).Object.getItemCount())
								For iCounter = 0 To iItemCount- 1 Step 1
									If trim(sTabName) = trim(objIntNoOfObjects(icount).Object.getItems().mic_arr_get(iCounter).getText()) Then
										objIntNoOfObjects(icount).Select sTabName
										bFlag=True
										Exit For 
									End IF
								Next
								If bFlag=True Then Exit For 
							Next
							
							If bFlag=False Then
								Fn_CM_TabOperation = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Select ["+sItem+"] Tab.")
							Else
								Fn_CM_TabOperation = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Selected ["+sItem+"] Tab.")
							End If
							
							Set objSelectType=Nothing
							Set  objIntNoOfObjects=Nothing
				Case "GetTabSequence" 	
'								sTabVal = ""
'								Set oTabCtl = Nothing								
'								 Set oTabCtl = JavaWindow("ChangeManager").JavaObject("MyTcTabObject")
'								  For iTabCount = 0 to Cint(oTabCtl.Object.getTabItemCount)-1
'									  If iTabCount = 0 Then
'										  sTabVal = oTabCtl.Object.getItem(iTabCount).text
'									  Else
'										  sTabVal = sTabVal+":"+oTabCtl.Object.getItem(iTabCount).text
'									  End If									 									 
'								  Next
'								  Fn_CM_TabOperation = sTabVal
'								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyTc_TabOperation Execution Sucessful")
'									 Set oTabCtl = Nothing
'							'TC 12 - 2017080200 - NishigandhaJ - 21/Aug/17 - Modified code as object changed in application			 
							sTabVal = ""		
							iItemCount = cInt(objJavaTab.Object.getItemCount())
							For iCounter = 0 To iItemCount- 1 Step 1
								 If iCounter = 0 Then
									  sTabVal = objJavaTab.Object.getItems().mic_arr_get(iCounter).getText()
								  Else
									  sTabVal = sTabVal+":"+objJavaTab.Object.getItems().mic_arr_get(iCounter).getText()
								  End If
							Next
							Fn_CM_TabOperation = sTabVal
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyTc_TabOperation Execution Sucessful")
												
				Case Else
							'Call Log file
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Invalid ACTION  [" +sAction+ "] is Requested.")
							Fn_CM_TabOperation = False
	End Select
	Set objJavaTab = Nothing
End Function
'-------------------------------------------------------------------Function Used to Create Proplem Report,Change Notice,Change Request,Deviation Request In Context----------------------------------------------
'Function Name		:	Fn_CM_CreateChangeInContext

'Description			 :	Function Used to Create Proplem Report,Change Notice,Change Request,Deviation Request In Context

'Parameters			   :	1.strNodeName: Node name(Proplem Report Or Change Notice Or Change Request Or Deviation Request)
										'2.strChangeID: ID (ID should be unique and have to be in Proper format)
										'3.strChangeRev: Revision
										'4.strChangeName:Name
										'5.strChangeDesc:Description
										'6.strSrchText:Change Type
										'7.strFilterText

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In

'Examples				:	bReturn=Fn_CM_CreateChangeInContext("Problem Report","PR-000071","A","TestPR","TestDescription","","")
										'bReturn=Fn_CM_CreateChangeInContext("Change Notice","ECN-000071","A","TestECN","ECNDescription","","")
										'bReturn=Fn_CM_CreateChangeInContext("Change Request","ECR-000071","A","TestECR","ECRDescription","","")
										'bReturn=Fn_CM_CreateChangeInContext("Deviation Request","EDR-000071","A","TestEDR","EDRDescription","TestEDRType","")
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				03/09/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_CreateChangeInContext(strNodeName,strChangeID,strChangeRev,strChangeName,strChangeDesc,strChangeType,strSrchText)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_CreateChangeInContext"
   'Declaring Variables
   Dim strNodePath,intNodeCount,sTreeItem,intCount
   Dim ObjChangeWnd
	Fn_CM_CreateChangeInContext=False
	'Verifying "New Change In context" window's existance
	If Fn_UI_ObjectExist("Fn_CM_CreateChangeInContext",JavaWindow("MyTeamcenter").JavaWindow("NewChangeInContext"))=False Then
		Exit Function	
	End If
	'Creating Object of "New Change In context" window
	Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_CM_CreateChangeInContext",JavaWindow("MyTeamcenter").JavaWindow("NewChangeInContext"))
	If strSrchText<>"" Then
		'Setting Deviation Change Type (This Text box is appear only in Deviation Request Case)
		Call Fn_UI_EditBox_Type("Fn_CM_CreateChangeInContext",ObjChangeWnd,"SearchText",strSrchText)
	End If
	Call Fn_JavaTree_Select("Fn_CM_CreateChangeInContext", ObjChangeWnd, "ChangeType","Complete List")
	strNodePath="Complete List:"+strNodeName
	'Verifying Node is present in Tree
    intNodeCount =Fn_UI_Object_GetROProperty("Fn_CM_CreateChangeInContext",ObjChangeWnd.JavaTree("ChangeType"),"items count")    
	For intCount = 0 to intNodeCount - 1
		sTreeItem = ObjChangeWnd.JavaTree("ChangeType").GetItem(intCount)
		If Trim(lcase(sTreeItem)) = Trim(Lcase(strNodePath)) Then
			Fn_CM_CreateChangeInContext = True
			Exit For
		End If
	Next
	If Cint(intCount) =Cint( intNodeCount) Then
		Call Fn_Button_Click("Fn_CM_CreateChangeInContext",ObjChangeWnd,"Cancel")
		Fn_CM_CreateChangeInContext =False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& strNodePath &"Node is not found in View Tree")
		Set ObjChangeWnd=Nothing
		Exit Function
	End If
	'Selecting Node from tree
   Call Fn_JavaTree_Select("Fn_CM_CreateChangeInContext", ObjChangeWnd, "ChangeType",strNodePath)
	'Clicking on Next button to proceed 
	Call Fn_Button_Click("Fn_CM_CreateChangeInContext",ObjChangeWnd,"Next")
	wait 1
	If strChangeID<>"" Then
		'Setting Id
		Call Fn_Edit_Box("Fn_CM_CreateChangeInContext",ObjChangeWnd,"Id",strChangeID)
		wait(2)
	End If
	If strChangeRev<>"" Then
		'Setting Revision
		Call Fn_Edit_Box("Fn_CM_CreateChangeInContext",ObjChangeWnd,"Revision",strChangeRev)
		wait(2)
	End If
	'Setting Name
	Call Fn_Edit_Box("Fn_CM_CreateChangeInContext",ObjChangeWnd,"Name",strChangeName)
	wait(2)
	'Setting Description
	Call Fn_Edit_Box("Fn_CM_CreateChangeInContext",ObjChangeWnd,"Description",strChangeDesc)
	wait(2)
	'Clicking On Finish Button To finish the Operation
	Call Fn_Button_Click("Fn_CM_CreateChangeInContext",ObjChangeWnd,"Finish")
	wait(2)
	'Clicking On Cancel Button To Cancel the Operation
	Call Fn_Button_Click("Fn_CM_CreateChangeInContext",ObjChangeWnd,"Cancel")
	wait(2)
	'function Return True
	Fn_CM_CreateChangeInContext=True
	'Releasing "New Change" window's object
	Set ObjChangeWnd=Nothing
End Function

'-------------------------------------------------------------------Function Used to perform operatons on component Tree------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_ComponentTreeOperations

'Description			 :	Function Used to perform operatons on component Tree

'Parameters			   :	1.strAction: Action Name
										'2.strNodeName: Node Name		
										'3.strMenu: Pop Up Menu Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective

'Examples				:	Fn_CM_ComponentTreeOperations("Select","Task1:Problem Item","")
										'Fn_CM_ComponentTreeOperations("Expand","Task1:Problem Item","")
										'Fn_CM_ComponentTreeOperations("VerifyNode","Task1:Problem Item","")
										'Fn_CM_ComponentTreeOperations("PopupMenuSelect","Task1:Problem Item","Refresh")
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				06/09/2010			           1.0																						Tushar B
'													Sandeep N										   				15/11/2011			           2.0					Added UI function call "Fn_UI_JavaTreeGetItemPath" to get Node Path 																
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ComponentTreeOperations(strAction,strNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_ComponentTreeOperations"
   'Variable declaration
   Dim aMenuList,intCount,iPath
   Dim intNodeCount,sTreeItem
   Dim ObjChMgrWnd
	Fn_CM_ComponentTreeOperations=False
	'Creating Object of ChangeManager window
	Set ObjChMgrWnd=Fn_UI_ObjectCreate("Fn_CM_ComponentTreeOperations", JavaWindow("ChangeManager"))
		Select Case strAction
			'=========================================================================================================
			Case "Select" 		'Fn_CM_ComponentTreeOperations("Select","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
					iPath=Fn_UI_JavaTreeGetItemPath(ObjChMgrWnd.JavaTree("ComponentTree"),strNodeName)
					If iPath<>False Then
						ObjChMgrWnd.JavaTree("ComponentTree").Select iPath
						Fn_CM_ComponentTreeOperations=True
					End If
			'=========================================================================================================
			Case "Expand" 'Fn_CM_ComponentTreeOperations("Expand","Change Home:My Open Changes","")
					iPath=Fn_UI_JavaTreeGetItemPath(ObjChMgrWnd.JavaTree("ComponentTree"),strNodeName)
					If iPath<>False Then
						Call Fn_UI_JavaTree_Expand("Fn_CM_ComponentTreeOperations", ObjChMgrWnd, "ComponentTree",iPath)
						Fn_CM_ComponentTreeOperations=True
					End If
			'=========================================================================================================
			Case "PopupMenuSelect"		'Fn_CM_ComponentTreeOperations("PopupMenuSelect","Change Home:Test1","Refresh")
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					iPath=Fn_UI_JavaTreeGetItemPath(ObjChMgrWnd.JavaTree("ComponentTree"),strNodeName)
					If iPath<>False Then
						ObjChMgrWnd.JavaTree("ComponentTree").Select iPath
					End If
					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_CM_ComponentTreeOperations",ObjChMgrWnd, "ComponentTree",iPath)
					'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CM_ComponentTreeOperations = False
							Set ObjChMgrWnd=Nothing
							Exit Function
					End Select
					ObjChMgrWnd.WinMenu("ContextMenu").Select StrMenu				
					Fn_CM_ComponentTreeOperations=True
			'=========================================================================================================
			Case "VerifyNode"		'Fn_CM_ComponentTreeOperations("VerifyNode","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
					iPath=Fn_UI_JavaTreeGetItemPath(ObjChMgrWnd.JavaTree("ComponentTree"),strNodeName)
					If iPath=False Then
						Fn_CM_ComponentTreeOperations = False
					Else
						Fn_CM_ComponentTreeOperations = True
					End If
			'=========================================================================================================
			Case Else
					Fn_CM_ComponentTreeOperations=False
		End Select
	'Rleasing Change Manager Window Object
	Set ObjChMgrWnd=Nothing
End Function
'-------------------------------------------------------------------Function Used to perform operatons on View Menu ToolBar Button And There with there Context Menu--------------------------------------------------
'Function Name		:	Fn_CM_ViewMenuOperations

'Description			 :	Function Used to perform operatons on View Menu ToolBar Button And There with there Context Menu

'Parameters			   :	1.strAction: Action Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective

'Examples				:		bReturn=Fn_CM_ViewMenuOperations("OpenSchedule")
'											bReturn=Fn_CM_ViewMenuOperations("ViewTaskfolders")
'											bReturn=Fn_CM_ViewMenuOperations("Rollup")
'								   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				06/09/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'This function perform RMB operation.
Public Function Fn_CM_ViewMenuOperations(strAction)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_ViewMenuOperations"
	Select Case strAction
		Case "OpenSchedule"
			Call Fn_ToolbarButtonClick_Ext(1,"View Menu")
			JavaWindow("ChangeManager").WinMenu("ContextMenu").Select "Open Schedule"
			Wait 10
            Call Fn_LoadSchedule("No","")'' // Modified  By Vidya 8/3/2013
'			If JavaDialog("Load Schedule").Exist Then 
'				JavaDialog("Load Schedule").JavaCheckBox("Don't show this message").Set("ON")
'				JavaDialog("Load Schedule").JavaButton("No").Click
'			End If
			Fn_CM_ViewMenuOperations=True
		 Case "RollUp"
			Call Fn_ToolbarButtonClick_Ext(1,"View Menu")
			JavaWindow("ChangeManager").WinMenu("ContextMenu").Select "Rollup"
			Fn_CM_ViewMenuOperations=True
		Case "ViewTaskfolders"
			 Call Fn_ToolbarButtonClick_Ext(2,"View Menu")
		
		  if JavaWindow("ChangeManager").WinMenu("ContextMenu").Exist =True Then
			JavaWindow("ChangeManager").WinMenu("ContextMenu").Select "View Task folders"
			Fn_CM_ViewMenuOperations=True
			Else 
			Fn_CM_ViewMenuOperations=False
			End  IF
			
		Case Else
			Fn_CM_ViewMenuOperations=False
	End Select
End Function

'-------------------------------------------------------------------Function Used to Create Task from Task Tab-----------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_CreateTask

'Description			 :	Function Used to Create Task from Task Tab

'Parameters			   :	1.strAction: Action Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective And present on Task Tab

'Examples				:		bReturn=Fn_CM_CreateTask("Task1","")
'											bReturn=Fn_CM_CreateTask("Task2","9h")
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				06/09/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_CreateTask(strTaskName,strWorkName)
		GBL_FAILED_FUNCTION_NAME="Fn_CM_CreateTask"
		Fn_CM_CreateTask=False
		Dim ObjDialog
		
		Set ObjDialog = JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame")
		wait(4)
		Call Fn_SISW_UI_JavaEdit_Operations("Fn_CM_CreateTask", "Set",ObjDialog.JavaEdit("TaskName"),"", strTaskName )
		'Call Fn_Edit_Box("Fn_CM_CreateTask", JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame"),"TaskName",strTaskName)
		If strWorkName<>"" Then
			Call Fn_SISW_UI_JavaEdit_Operations("Fn_CM_CreateTask", "setText", ObjDialog.JavaEdit("Work"),"", strWorkName )
			'Call Fn_Edit_Box("Fn_CM_CreateTask", JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame"),"Work",strWorkName)
		End If
        If Fn_UI_ObjectExist("Fn_CM_CreateTask", ObjDialog.JavaButton("Create"))=True Then
			Call Fn_Button_Click("Fn_CM_CreateTask", ObjDialog, "Create")
			Fn_CM_CreateTask=True
		End If
		
		Set ObjDialog = Nothing
End Function

'-------------------------------------------------------------------Function Used to Perform Operation On Task Tree------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_TaskTableOperations

'Description			 :	Function Used to Perform Operation On Task Table

'Parameters			   :	1.strAction: Action Name
'										2.strTaskName: Task Name
'										3.strPopupMenu: Menu Name

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective And present on Task Tab

'Examples				:		Fn_CM_TaskTreeOperations("SelectTask","Test:Task1","")
'											
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				06/09/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_TaskTreeOperations(strAction,strTaskName,strPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_TaskTreeOperations"
	Dim iCounter,iRows,sCellData
	Fn_CM_TaskTreeOperations=false
   Select Case strAction
		 	Case "SelectTask"
				JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame").JavaTable("TaskTable").ClickCell 0,"Object"
				wait 2
				iRows=Fn_UI_Object_GetROProperty("Fn_CM_TaskTreeOperations",JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame").JavaTable("TaskTable"),"rows")
				For iCounter=0 To iRows-1
                    Call Fn_UI_JavaTable_SelectRow("Fn_CM_TaskTreeOperations", JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame"), "TaskTable",iCounter)
					wait 1
					sCellData=JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame").JavaTable("TaskTable").GetCellData(iCounter,"Object")
					If sCellData=strTaskName Then
                        Call Fn_UI_JavaTable_SelectRow("Fn_CM_TaskTreeOperations", JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame"), "TaskTable",iCounter)
						Fn_CM_TaskTreeOperations=True
						Exit For
					End If
				Next
   End Select	
End Function
'-------------------------------------------------------------------Function Used to Perform Operation on "Commit Rollup" ToolBar buttons Drop Down Menu----------------------------------------------------------------------------
'Function Name		:	Fn_CM_CommitRollupDrpDwnOperation

'Description			 :	Function Used to Perform Operation on "Commit Rollup" ToolBar buttons Drop Down Menu

'Parameters			   :	1.StrMenuName: Menu Name Which have to Select
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective And Rollup Tree Tab activated

'Examples				:	Fn_CM_CommitRollupDrpDwnOperation("Problem Items")
'										Fn_CM_CommitRollupDrpDwnOperation("Impacted Items")

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				09/09/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_CommitRollupDrpDwnOperation(StrMenuName)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_CommitRollupDrpDwnOperation"
   Dim ObjDesc, ArrLists, iToolCnt, iCounter, sContents
	If JavaWindow("DefaultWindow").Exist(iTimeOut) Then
		'Create Toolbar object
		Set ObjDesc = Description.Create() 
		ObjDesc("to_class").Value = "JavaToolbar" 
		ObjDesc("enabled").Value = 1
		JavaWindow("DefaultWindow").Maximize
		'Get the total of Toolbar objects
		Set ArrLists =JavaWindow("DefaultWindow").ChildObjects(ObjDesc)
		iToolCnt = JavaWindow("DefaultWindow").ChildObjects(ObjDesc).count
		For iCounter = 0 to iToolCnt-1
			sContents = ArrLists(iCounter).GetContent()
			If instr(sContents, "Commit Rollup") > 0 Then
                ArrLists(iCounter).ShowDropdown "Commit Rollup"
				JavaWindow("ChangeManager").WinMenu("ContextMenu").Select StrMenuName
				Fn_CM_CommitRollupDrpDwnOperation =True
				Exit For
			End If
		Next
		If iCounter = iToolCnt Then
			Fn_CM_CommitRollupDrpDwnOperation = False
		End If
		Set ObjDesc = Nothing
		Set ArrLists = Nothing
	Else
		Fn_CM_CommitRollupDrpDwnOperation =False
	End If
End Function

'-------------------------------------------------------------------Function Used to perform operatons on  Rollup Tree--------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_RollupTreeOperations

'Description			 :	Function Used to perform operatons on  Rollup Tree

'Parameters			   :	1.strAction: Action Name
										'2.strTreeName:Tree Name On which have to perform operation
										'3.strNode: Node Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective And Rollup Tree Tab should be open

'Examples				:	Fn_CM_RollupTreeOperations("Select","ECN-000005/A;1-PRN:Impacted Items","")
										'Fn_CM_RollupTreeOperations("Expand","ECN-000005/A;1-PRN:Impacted Items","")
										'Fn_CM_RollupTreeOperations("VerifyNode","ECN-000005/A;1-PRN:Impacted Items","")
										'Fn_CM_RollupTreeOperations("VerifyRelatedTask","ECN-280509/A;1-ECN175:Reference Items:000613-Item1114369","Task1")

										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				09/09/2010			           1.0																					 Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_RollupTreeOperations(strAction,strNode,strMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_RollupTreeOperations"
   'Variable Declarion
	Dim intNodeCount,sTreeItem,intCount,sTask
	
	Select Case strAction
		Case "Select" 
            Call Fn_JavaTree_Select("Fn_CM_RollupTreeOperations",JavaWindow("ChangeManager"),"RollupTree",strNode)
			Fn_CM_RollupTreeOperations=True
		Case "Expand"
            Call Fn_UI_JavaTree_Expand("Fn_CM_RollupTreeOperations",JavaWindow("ChangeManager"),"RollupTree",strNode)
			Fn_CM_RollupTreeOperations=True
		Case "VerifyNode" 
            intNodeCount = Fn_UI_Object_GetROProperty("Fn_CM_RollupTreeOperations",JavaWindow("ChangeManager").JavaTree("RollupTree"),"items count")
			For intCount = 0 to intNodeCount - 1
				sTreeItem = JavaWindow("ChangeManager").JavaTree("RollupTree").GetItem(intCount)
				If Trim(lcase(sTreeItem)) = Trim(Lcase(strNode)) Then
					Fn_CM_RollupTreeOperations = True
					Exit For
				End If
			Next
			If Cint(intCount) = Cint(intNodeCount) Then
				Fn_CM_RollupTreeOperations = False
			End If
		Case "TreeClose"
			JavaWindow("ChangeManager").JavaTab("Tab").SetTOProperty "index","2"
			wait 3
			JavaWindow("ChangeManager").JavaTab("Tab").CloseTab "Roll Up"
        Case "VerifyRelatedTask"
			sTask = JavaWindow("ChangeManager").JavaTree("RollupTree").GetColumnValue( strNode,"Related Tasks")
			If Instr(1,sTask,strMenu,1) > 0 Then
					Fn_CM_RollupTreeOperations = True
			Else
					Fn_CM_RollupTreeOperations = False
			End If
	End Select
End Function
'-------------------------------------------------------------------Function Used to Create Derive Change------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_DeriveChangeCreate

'Description			 :	Function Used to Create Derive Change

'Parameters			   :	1.strNodeName: Node name(Proplem Report Or Change Notice Or Change Request Or Deviation Request)
										'2.strChangeID: ID (ID should be unique and have to be in Proper format)
										'3.strChangeRev: Revision
										'4.strChangeSynopsis:Name
										'5.strChangeDesc:Description
										'6.strChangeType:Change Type
										'7.bPropRelation:Propaget Relation Opetion
										'8.strSrchText:Search Type
										

'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In

'Examples				:	Fn_CM_DeriveChangeCreate("Change Notice","","","Test","TestDesc","","ON","")
'										Fn_CM_DeriveChangeCreate("Change Request","","","Test","TestDesc","","OFF","")
' 										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				10/09/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_DeriveChangeCreate(strNodeName,strChangeID,strChangeRev,strChangeSynopsis,strChangeDesc,strChangeType,bPropRelation,strSrchText)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_DeriveChangeCreate"
	'Declaring Variables
   Dim strNodePath,intNodeCount,sTreeItem,intCount
   Dim ObjDeriveChngWnd
	Fn_CM_DeriveChangeCreate=False
	'Creating Object of "New Change In context" window
	Set ObjDeriveChngWnd=Fn_UI_ObjectCreate("Fn_CM_DeriveChangeCreate", JavaWindow("DefaultWindow").JavaWindow("DeriveChange"))
	
	If ObjDeriveChngWnd.JavaTree("ChangeType").Exist(2) Then
	
	If strSrchText<>"" Then
		'Setting Deviation Change Type (This Text box is appear only in Deviation Request Case)
		Call Fn_UI_EditBox_Type("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"SrchChangeType",strSrchText)
	End If
	Call Fn_JavaTree_Select("Fn_CM_DeriveChangeCreate", ObjDeriveChngWnd, "ChangeType","Complete List")
	Call Fn_UI_JavaTree_Expand("Fn_CM_DeriveChangeCreate", ObjDeriveChngWnd, "ChangeType","Complete List")
	strNodePath="Complete List:"+strNodeName
	'Verifying Node is present in Tree
    intNodeCount =Fn_UI_Object_GetROProperty("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd.JavaTree("ChangeType"),"items count")    
	For intCount = 0 to intNodeCount - 1
		sTreeItem = ObjDeriveChngWnd.JavaTree("ChangeType").GetItem(intCount)
		If Trim(lcase(sTreeItem)) = Trim(Lcase(strNodePath)) Then
			Fn_CM_DeriveChangeCreate = True
			Exit For
		End If
	Next
	If Cint(intCount) =Cint( intNodeCount) Then
		Call Fn_Button_Click("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Cancel")
		Fn_CM_DeriveChangeCreate =False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& strNodePath &"Node is not found in View Tree")
		Set ObjDeriveChngWnd=Nothing
		Exit Function
	End If
	'Selecting Node from tree
   Call Fn_JavaTree_Select("Fn_CM_DeriveChangeCreate", ObjDeriveChngWnd, "ChangeType",strNodePath)
	'Clicking on Next button to proceed 
	Call Fn_Button_Click("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Next")
	
	End If
	
	If strChangeID<>"" Then
		'Setting Id
		Call Fn_UI_EditBox_Type("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Id",strChangeID)
	End If
	
	If strChangeRev<>"" Then
		'Setting Revision
		Call Fn_UI_EditBox_Type("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Revision",strChangeRev)
	End If
	If  strChangeSynopsis<>"" Then
		'Setting Name
		Call Fn_Edit_Box("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Synopsis","")
		Call Fn_UI_EditBox_Type("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Synopsis",strChangeSynopsis)        
	End If
	If strChangeDesc<>""  Then
		'Setting Description
		Call Fn_Edit_Box("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Description","")
		Call Fn_UI_EditBox_Type("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Description",strChangeDesc)
	End If
	If strChangeType<>"" Then
		Call Fn_UI_EditBox_Type("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"ChangeType",strChangeType)
	End If
	If bPropRelation<>"" Then
        Call Fn_CheckBox_Set("Fn_CM_DeriveChangeCreate", ObjDeriveChngWnd, "Propagaterelations", bPropRelation)
	End If
	'Clicking On Finish Button To finish the Operation
	Call Fn_Button_Click("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Finish")
	'Added by Vallari - 13-Apr-2011
	Call Fn_ReadyStatusSync(2)
	JavaWindow("DefaultWindow").JavaWindow("DeriveChange").JavaButton("Cancel").Click micLeftBtn
	'Call Fn_Button_Click("Fn_CM_DeriveChangeCreate",ObjDeriveChngWnd,"Cancel")
	'function Return True
	Fn_CM_DeriveChangeCreate=True
	'Releasing "New Change" window's object
	Set ObjDeriveChngWnd=Nothing
End Function 


'-------------------------------------------------------------------Function Used to Handle Error Window-------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_ErrorWindowMsgVerify

'Description			 :	Function Used to Handle Error Window

'Parameters			   :   '1.sDilogName: Error Dialog Name
										'2.sErrorMessage: Error message
										'3.sButtonName:Button name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Error Window Should be appear on srceen

'Examples				:	Fn_CM_ErrorWindowMsgVerify("","The Requested ","OK")			
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				15/10/2010			           1.0																						Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ErrorWindowMsgVerify(sDilogName,sErrorMessage,sButtonName)

	Dim dicErrorInfo
	 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	 dicErrorInfo.Add "Action", "ErrorWindowMsgVerify"
	 dicErrorInfo.Add "Title", sDialogName
	 dicErrorInfo.Add "Message", sErrorMessage
	 dicErrorInfo.Add "Button", sButtonName    
	 Fn_CM_ErrorWindowMsgVerify = Fn_SISW_CM_ErrorVerify(dicErrorInfo)

End Function

'-------------------------------------------------------------------Function Used to Verify Property with value in Summary tab-------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_SummaryPropertyVerify

'Description			 :	Function Used to Verify Property with value in Summary tab

'Parameters			   :   '1.strProperty:  - propertyname : PropertyValue   
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Summary Tab Should be appear on srceen

'Examples				:	Call Fn_CM_SummaryPropertyVerify("Disposition:None;Maturity:Elaborating")			
'										Call Fn_CM_SummaryPropertyVerify("Change Specialist I:dba/DBA/cmuser01")
'										Call Fn_CM_SummaryPropertyVerify("Change Review Board:Engineering/Designer/cmuser01")  
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Pranav Ingle						   						15/09/2010			           1.0																						Sunny
'													Sandeep Navghane						   		02/01/2011			           1.1																						Sunny
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_CM_SummaryPropertyVerify(strProperty)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_SummaryPropertyVerify"
	Dim ObjTCwindow,arrProperty,iCounter,bCheck,iProperty,strValue,arrPropertyValue,arrPropertyName
	Fn_CM_SummaryPropertyVerify = False
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	Set ObjTCwindow =Nothing
	If JavaWindow("DefaultWindow").JavaTab("GeneralTab").Exist(5) then
		JavaWindow("DefaultWindow").JavaTab("GeneralTab").Select "Overview"
		Call Fn_ReadyStatusSync(1)
	end if
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Checking Existance of [ MyTeamcenter ] or [ ChangeManager ] window
	If JavaWindow("MyTeamcenter").Exist(6) Then
			'Creating Object of [ MyTeamcenter ] window
			Set ObjTCwindow = JavaWindow("MyTeamcenter")
	ElseIf JavaWindow("ChangeManager").Exist(6) Then
			'Creating Object of [ ChangeManager ] window
			Set ObjTCwindow = JavaWindow("ChangeManager")
	Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Window does not Exist")
			Exit Function
	End If
	'Spliting Multiple properties
	arrProperty = Split(strProperty,";")
	For iCounter = 0 to UBound(arrProperty)
			bCheck=False
			'Spliting property name and value
			iProperty = Split(arrProperty(iCounter),":")
			arrPropertyName= iProperty(0)
			arrPropertyValue= iProperty(1)		
			Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_CM_SummaryPropertyVerify",ObjTCwindow.JavaStaticText("Summary_Text"),"label",arrPropertyName+":")
			If ObjTCwindow.JavaEdit("Summary_Text").Exist(6) Then
				strValue=Fn_UI_Object_GetROProperty("Fn_CM_SummaryPropertyVerify",ObjTCwindow.JavaEdit("Summary_Text"),"value")
				If Trim(strValue)=Trim(arrPropertyValue) Then
					bCheck=True
				End If
			ElseIf ObjTCwindow.JavaObject("Summary_Object").Exist(6) Then
				strValue=Fn_UI_Object_GetROProperty("Fn_CM_SummaryPropertyVerify",ObjTCwindow.JavaObject("Summary_Object"), "text")
				If Trim(strValue)=Trim(arrPropertyValue) Then
					bCheck=True
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Property [ "+arrPropertyName+" ] does not exist on Summary Page")
				Set ObjTCwindow =Nothing
				Exit Function
			End If
			If bCheck=False Then
				Exit Function
			End If
	Next
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Fn_CM_SummaryPropertyVerify Completed Successfully")
	Fn_CM_SummaryPropertyVerify = true
	ObjTCwindow.JavaStaticText("Summary_Text").SetTOProperty "label",""

	Set ObjTCwindow =Nothing
End Function
'*********************************************************		Function Checks Out the Teamcenter Obejct		***********************************************************************
'Function Name		:				Fn_CM_ObjectPropertyPanelVerify

'Description			 :		 		This function allows verification of an object property through the Properties Panel.

'Parameters			   : 	               1) sPropName: name of the property for which value needs verification
'														2) sValue: Property value to be verified

'														Note: Multiple values can be passed using Array with ; separation
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:			 		Select object>Window>Show View> Properties
 
'														Once the properties panel is activated for the selected object, verify the required info.

'Examples				:				 Fn_CM_ObjectPropertyPanelVerify("Properties:Current Name;Properties:Current ID","Item1;004260")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	Reviewed date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N					24-Sep-2010						1.0														Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_CM_ObjectPropertyPanelVerify(sPropName,sValue)
			GBL_FAILED_FUNCTION_NAME="Fn_CM_ObjectPropertyPanelVerify"
			Dim objPropertyTree
			Dim aProperties, aValues, iCount, iCounter, sPropVal, bFlag
			Call Fn_SetView("General:Properties")
			Call Fn_ToolbarOperation("Click", "Show Advanced Properties","")
			Call Fn_ReadyStatusSync(3)
			' Multiple Values are seperated and stored in Array. 
			aProperties =Split(sPropName,";",-1,1)
			If trim(sValue) = "" Then
				aValues= Array("")
			Else
				aValues=Split(sValue,";",-1,1)
			End If
			Set  objPropertyTree = Fn_UI_ObjectCreate("Fn_CM_ObjectPropertyPanelVerify",JavaWindow("DefaultWindow").JavaTree("PropertiesTree"))
			
			For iCounter=0 to ubound(aProperties)
					bFlag=FALSE
					For iCount=0 to objPropertyTree.GetROProperty("count_all_items")-1
							'Verify that sPropName is Exist
							If  aProperties(iCounter)=objPropertyTree.GetItem(iCount) then
									bFlag=TRUE
									'Converting values in integer then converting in string
									If Isnumeric(aValues(iCounter)) then
											aValues(iCounter) = Cstr(Cint(aValues(iCounter)))
									End If 
									JavaWindow("DefaultWindow").JavaTab("GeneralTab").SetTOProperty "index",1
									'Verifying Whether Object Property values and Property Panel Values are equal 								
									sPropVal=Cstr(objPropertyTree.GetColumnValue(aProperties(iCounter),"Value"))
									aValues(iCounter) = Replace(aValues(iCounter),"~",";")
									If  Trim(sPropVal) <> Trim(CStr(aValues(iCounter)))  then
											Fn_CM_ObjectPropertyPanelVerify=False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Object Property values and Property Panel Values are not equal ")
											'Close Properties Tab
											Call Fn_TabFolder_Operation("Properties_Close", "Properties", "")
											Exit Function
									End If
									Exit For
                           	End If
					Next
					If bFlag=False Then
							Fn_CM_ObjectPropertyPanelVerify=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Property Name and Property Panel Properties Name are equal ")
							'Close Properties Tab
							Call Fn_TabFolder_Operation("Properties_Close", "Properties", "")
							Exit Function
					End If
			Next
		'Close Properties Tab
		Call Fn_TabFolder_Operation("Properties_Close", "Properties", "")
		Fn_CM_ObjectPropertyPanelVerify=True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Property values and Property Panel Values are equal ")	
		Set objPropertyTree=Nothing
End Function

'-------------------------------------------------------------------Function Used to Assign task to a user-------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_TaskAssignment(sAction, sUser,aResourceLevel)

'Description			 :	Function Used to Assign task to a user

'Parameters			   :    
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Task should be selected

'Examples				:	Call Fn_CM_TaskAssignment("Assign", aLoginUser(0)+" ("+aLoginUser(5)+")","100")		
										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sunny Ruparel						   						27/09/2010			           1.0																						Sunny
			'													Ikhlaque						   						22/06/2012			           1.1																						Sandeep
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
Public Function Fn_CM_TaskAssignment(sAction, sUser,aResourceLevel)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_TaskAssignment"
	'variable declaration
	Dim objDialog,sMenu
	'Creating Object of  [ TaskAssignment ] window
	sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "ScheduleAssignmentsAssigntoTask")
	Set objDialog = JavaWindow("ChangeManager").JavaWindow("TaskAssignment")
'	Set objDialog = Fn_UI_ObjectCreate("Fn_CM_TaskAssignment",JavaWindow("ChangeManager").JavaWindow("TaskAssignment"))
	If  Fn_UI_ObjectExist("Fn_CM_TaskAssignment",objDialog)=False  Then	
		Call Fn_MenuOperation("Select",sMenu)
		Call Fn_ReadyStatusSync(1)
	End If
	
	Select Case sAction

		Case "Assign"
				'Removed all old code by Ikhlaque added new code
				If objDialog.JavaTree("AssignmentTree").Exist(4)  Then
					objDialog.JavaTree("AssignmentTree").select sUser
					Call Fn_Button_Click("Fn_CM_TaskAssignment", objDialog, "Add")
					Call Fn_Button_Click("Fn_CM_TaskAssignment", objDialog, "OK")
					Fn_CM_TaskAssignment = True
				End If			
		Case Else

			Fn_CM_TaskAssignment = False

	End Select

	Set objDialog=Nothing

End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:		Fn_SchMgr_ScheduleMembership 

'Description			 :		 This function is used  to add memeber  to schedule membership

'Parameters			   :	  1.sAction : Action need to perform.
'                                          2.strMemberPath : Member Name Path (Comma Seperated [ , ])
											
'Return Value		   : 	True/False

'Pre-requisite			:	Schedule Should be selected

'Examples				:	Fn_CM_ScheduleMembership("SingleMemberAdd","Organization, User, Amol Lanke (x_lanke)")
'
'Note 						  :	Need to pass member path correctly spaces and comma's				
'
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep N						1-Oct-2010			1.0																	Tushar B	
'									Sandeep N						3-Jul-2012			1.1					Added Code to Expand Parent node 												Sukhada
'									Sandeep N						16-Jul-2012			1.2					Added Code to handle new Object Hierarchy : JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame").JavaDialog("ScheduleMembership")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ScheduleMembership(strAction,strMemberPath)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_ScheduleMembership"
   'Declaring Variables
	Dim ObjSchDialog
	Dim iItemCount,iCount,strNodeName,bFlag
	Dim iCounter1,iCounter,aMember,sNode

	Fn_CM_ScheduleMembership=False
	bFlag=False
	'Verifying Existance Of ScheduleMembership Dialog
	If not JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame").JavaDialog("ScheduleMembership").Exist(6) and not JavaWindow("ChangeManager").JavaWindow("ChangeManager").JavaDialog("ScheduleMembership").Exist(5) Then
		'Invoking ScheduleMembership Dialog if its already not Exist
		Call Fn_MenuOperation("Select","Schedule:Schedule Membership")
		Call Fn_ReadyStatusSync(1)
	End If
	'Creating Object Of ScheduleMembership Dialog
	If JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame").JavaDialog("ScheduleMembership").Exist(5) Then
		Set ObjSchDialog=Fn_UI_ObjectCreate("Fn_CM_ScheduleMembership",JavaWindow("ChangeManager").JavaWindow("WEmbeddedFrame").JavaDialog("ScheduleMembership"))
	ElseIf JavaWindow("ChangeManager").JavaWindow("ChangeManager").JavaDialog("ScheduleMembership").Exist(5) then
		Set ObjSchDialog=Fn_UI_ObjectCreate("Fn_CM_ScheduleMembership",JavaWindow("ChangeManager").JavaWindow("ChangeManager").JavaDialog("ScheduleMembership"))
	Else	
		Exit function
	End If
	Select Case strAction
		Case "SingleMemberAdd" 'Case to select single Member 
			'------ Added code to expand all parent nodes----------------------------------------------------------------------------------------------------------
			aMember=Split(strMemberPath,",")

			For iCounter1=0 to ubound(aMember)-1
				If iCounter1=0 Then
					sNode=aMember(0)
				else
					sNode=sNode+","+aMember(iCounter1)
				End If
				bFlag=False
				iItemCount=Fn_UI_Object_GetROProperty("Fn_CM_ScheduleMembership",ObjSchDialog.JavaTree("ScheduleMemberTree"), "items count")
				For iCounter=0 to iItemCount-1	
					strNodeName=ObjSchDialog.JavaTree("ScheduleMemberTree").Object.getPathForRow(iCounter).tostring
					strNodeName=Mid (strNodeName,2,Len(strNodeName)-2)
					If Trim(strNodeName)=Trim(sNode) Then
						strNodeName=ObjSchDialog.JavaTree("ScheduleMemberTree").GetItem(iCounter)
						wait(1)
						ObjSchDialog.JavaTree("ScheduleMemberTree").Expand strNodeName		
						bFlag=True
						Exit for
					end if
				Next
				If bFlag=false Then
					Exit for
				End If
			Next
			'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

			'Selecting Member from members tree
			If bFlag=true Then
				iItemCount=Fn_UI_Object_GetROProperty("Fn_CM_ScheduleMembership",ObjSchDialog.JavaTree("ScheduleMemberTree"), "items count")
				For iCount=0 To iItemCount-1
					strNodeName=ObjSchDialog.JavaTree("ScheduleMemberTree").Object.getPathForRow(iCount).tostring
					strNodeName=Mid (strNodeName,2,Len(strNodeName)-2)
					If  Trim(strNodeName)=Trim(strMemberPath) Then
						strNodeName=ObjSchDialog.JavaTree("ScheduleMemberTree").GetItem(iCount)
						wait(3)
						ObjSchDialog.JavaTree("ScheduleMemberTree").Select strNodeName
						bFlag=True
						Exit For
					End If
				Next
			End If
			If bFlag=False Then
				Fn_CM_ScheduleMembership=False
				Set ObjSchDialog=Nothing
				Exit Function
			End If
	End Select
	Fn_CM_ScheduleMembership=True
	'Clicking Ok button to add member
	Call Fn_Button_Click("Fn_CM_ScheduleMembership", ObjSchDialog, "OK")
	'Releasing ScheduleMembership Dialog object
	Set ObjSchDialog=Nothing
End Function
'-------------------------------------------------------------------Function Used to Execute Search Query From Change Manager Perspective---------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_SpecifyQueryDetailsAndInvoke

'Description			 :	Function Used to Execute Search Query From Change Manager Perspective

'Parameters			   :	1.cmSearchCriteriaDictionary: Dictionary Object of Key-Value Pair For Search Criteria
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective And Search Tab have to Present

'Examples				:	dicSearchCriteria( "Description")="Test"
'										Fn_CM_SpecifyQueryDetailsAndInvoke(dicSearchCriteria)

'Note						  :	 This Function use dictionary Object which is used for My TC search Query

'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				06/10/2010			           1.0																					   Suuny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_SpecifyQueryDetailsAndInvoke(cmSearchCriteriaDictionary)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_SpecifyQueryDetailsAndInvoke"
'Variable Declaration
Dim DictItems, DictKeys, strAction, iCounter
Dim ObjCMJavaWindow
'Checking Existance Of Change Manager Window
If Fn_UI_ObjectExist("Fn_CM_SpecifyQueryDetailsAndInvoke",JavaWindow("ChangeManager"))=True Then
		'Creating Object Of Change Manager Window
		Set ObjCMJavaWindow=Fn_UI_ObjectCreate("Fn_CM_SpecifyQueryDetailsAndInvoke",JavaWindow("ChangeManager"))
		'Clearing All Search fields
		Call Fn_ToolbatButtonClick("Clear all search fields")
		'Get the keys & items count from data dictionary.	
		DictItems = cmSearchCriteriaDictionary.Items
		DictKeys = cmSearchCriteriaDictionary.Keys
		For iCounter = 0 to cmSearchCriteriaDictionary.Count - 1
				  If IsNull(DictKeys(iCounter))  Then
				 Else
							If  DictItems(iCounter) = "" Then										
							Else
									strAction = DictKeys(iCounter)
									' Set the value as per the data dictioanry key.
									Select case strAction
											'----------------------- EditBox Set ----------------------
											Case "Description","Synopsis"
													Call Fn_Edit_Box("Fn_CM_SpecifyQueryDetailsAndInvoke",ObjCMJavaWindow,strAction,DictItems(iCounter))
													Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: The property "& strAction  & " Changed successfully")
									End select 
							End If
				End If                              
		Next
		'Clicking On Execute Query ToolBar button
		Call Fn_ToolbatButtonClick("Executes the search and displays the results in search result view") 
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Values specified Successfully set and search executed.  ")
		'Function returns true after executing query
		Fn_CM_SpecifyQueryDetailsAndInvoke = True
Else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail : Change Manager Window does not exist.") 
	'Function returns false if Change Manager window does not exist
	Fn_CM_SpecifyQueryDetailsAndInvoke = False
End If 
'releasing Change Manager windows object
Set ObjCMJavaWindow = Nothing	
End Function

'#######################################################################################################################################################~
'########################################################    Set Requisite Search Preference		###############################################################~
'#
'# FUNCTION NAME:	Fn_MyTc_SrchSavedSearchOperation
'#
'#  MyCommunity ID :	320
'#
'# MODULE: 						 Change Manager
'#
'# DESCRIPTION:			 1. Check/UnCheck the[Is Shared] box
'#											2. Click on [Create In] button on [Add Search to My Saved Searches]
'#											3. Expand and select prefered Tree Path under [My Saved Search]
'#											
'#											Case: Add_To_My_Saved_Search  		' 
'#													a. Specify [Name] of the saved search									
'#													b. Specify Folder Name on [Folder Information] dialog and click [OK]
'#													c. Select Path
'#											Case: Saved Search Delete
'#													a. Click on [Delete] button
'#													b. Click [OK] on Warning dialog		
'#											Case: Saved Search Rename
'#												a. Click on [Rename] button
'#												b. Specify New Name through send Key operation
'#												c. Send [Enter] Key						
'#											Case: Saved Search Validate		' Need to pass Full Reference path in "sSearchName" for existance check.							
'#												a. Validate existance of saved search
'#													4. Click [OK]
'#		
'# PRE-REQUISITE:		1. RAC Session accessible and My Teacenter Application Search Pane loaded
'#											 2. Search Criteria applied with various input values 
'#									>>   3. Click on Toolbar button [Add Search To My Saved Searches] under Search Pane
'#							
'# PARAMETERS   :       sAction: Name of the case to exercise pertaining to saved search
'#                                           bIsShared : ON/OFF
'#										 	 sSourceFolderPath: Existing Search Folder Path
'#											 sNewName: New Name for folder OR Search
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				Fn_MyTc_SrchSavedSearchOperation("Add_To_My_Saved_Search",  "ON", "A:B:C", "NewName") 
'#										
'#	History	:						Developer Name			Date			Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#								Prasad Kulkarni																													Deepak 
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'########################################################################################################################################################~
'########################################################    Set Requisite Search Preference		###############################################################~

Public Function  Fn_Chng_SrchSavedSearchOperation(sAction,  bIsShared, sSourceFolderPath, sNewName)
	GBL_FAILED_FUNCTION_NAME="Fn_Chng_SrchSavedSearchOperation"
Dim iNode, sNode, iCount, bExist, aRefPath, iCountOuter, afullRefPath, iEle, bFolder
Dim ObjJavaWndw,ObjQkLnkWndw, ObjAddSrch
Dim ArrLists,ObjDesc,iCounter
Dim ObjJavaTree,intNodeCount,sTreeItem
Dim iNod,sNod
Set ObjJavaWndw = JavaWindow("DefaultWindow") 

		Select Case sAction

				Case "Add_To_My_Saved_Search" 
			
						'	[Example ( last parameter is for name and its compulsary )]			Fn_MyTc_SrchSavedSearchOperation("Add_To_My_Saved_Search",  "ON", "A:B:C", "NewName") 
						' If Dont want to save/create new folder then pass blank "" in parameter "sSourceFolderPath"  ->> this will save query under parent node.
						'<<<<<<<<<<<<<<<<<<<<<<<  Add to my saved search  >>>>>>>>>>>>>>>>>>>>>>>>> 
						If True = Fn_UI_ObjectExist("Fn_Chng_SrchSavedSearchOperation",JavaWindow("ChangeManager").JavaWindow("AddSrchtoMySaved")) Then

								Set ObjAddSrch = Fn_UI_ObjectCreate( "Fn_Chng_SrchSavedSearchOperation",JavaWindow("ChangeManager").JavaWindow("AddSrchtoMySaved") )		
								If sSourceFolderPath <> "" Then
	
											Call Fn_Button_Click("Fn_Chng_SrchSavedSearchOperation", ObjAddSrch, "CreateIn")
				
											aRefPath = Split(sSourceFolderPath,":")
											afullRefPath = "My Saved Searches"
										
											For iCountOuter = 0 To  UBound(aRefPath)
														bExist = False
														'iNode=ObjAddSrch.JavaTree("ExistingSavedSrchs").GetROProperty("items count")				
														For  iCount=1 to ObjAddSrch.JavaTree("ExistingSavedSrchs").GetROProperty("items count")-1
																sNode = ObjAddSrch.JavaTree("ExistingSavedSrchs").GetItem(iCount)
																If  sNode = afullRefPath+":"+aRefPath(iCountOuter) Then
																	bExist = true
																	Exit For
																End If
														Next
														If True =  bExist Then
																	afullRefPath =afullRefPath+":"+aRefPath(iCountOuter) 
																	Call Fn_JavaTree_Select("Fn_Chng_SrchSavedSearchOperation", ObjAddSrch, "ExistingSavedSrchs", afullRefPath)
														Else
																'CreateFolder
																Call Fn_Button_Click("Fn_Chng_SrchSavedSearchOperation", ObjAddSrch, "NewFolder")
																If True = Fn_UI_ObjectExist("Fn_Chng_SrchSavedSearchOperation",JavaWindow("ChangeManager").JavaWindow("Folder Information") ) Then
																		Call Fn_Edit_Box("Fn_Chng_SrchSavedSearchOperation",JavaWindow("ChangeManager").JavaWindow("Folder Information"),"Name","")
																		Call Fn_UI_EditBox_Type("Fn_Chng_SrchSavedSearchOperation",JavaWindow("ChangeManager").JavaWindow("Folder Information"),"Name", aRefPath(iCountOuter) )
																		Call Fn_Button_Click("Fn_Chng_SrchSavedSearchOperation", JavaWindow("ChangeManager").JavaWindow("Folder Information"), "OK")		
																		Call Fn_ReadyStatusSync(5)
																		afullRefPath =afullRefPath+":"+aRefPath(iCountOuter) 
																		Call Fn_JavaTree_Select("Fn_Chng_SrchSavedSearchOperation", ObjAddSrch, "ExistingSavedSrchs", afullRefPath)
																Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Window Folder Create Does Not Exist ")
																		Fn_Chng_SrchSavedSearchOperation = False
																End If
														End If 	
											Next
								End If
								Call Fn_CheckBox_Set("Fn_Chng_SrchSavedSearchOperation", ObjAddSrch, "IsShared", bIsShared )
								Call Fn_Edit_Box("Fn_Chng_SrchSavedSearchOperation",ObjAddSrch,"Name","")
								Call Fn_UI_EditBox_Type("Fn_Chng_SrchSavedSearchOperation",ObjAddSrch,"Name", sNewName )
                                Call Fn_Button_Click("Fn_Chng_SrchSavedSearchOperation", ObjAddSrch, "OK")

								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Passed:Search Successfully Added to My Saved Searches" )
								Fn_Chng_SrchSavedSearchOperation = True
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed: Window Does Not Exist." )
								Fn_Chng_SrchSavedSearchOperation = False
						End If
						'<<<<<<<<<<<<<<<<<<<<<<<  Add to my saved search  >>>>>>>>>>>>>>>>>>>>>>>>>
				Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Requested Action Not Valid. ")   	
							Fn_Chng_SrchSavedSearchOperation = False

				End Select

Set ObjJavaWndw = Nothing
Set ObjAddSrch = Nothing
End Function			

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:				Fn_CM_CompnentTabOperations(sAction,sTCComponentTabName) 

'Description		:		 		Function used to Perform operations on Component Tab of Change Manager


'Parameters			   :				1) sAction: Action to be performed on the Tab
'												  2) sTCComponentTabName: ComponentTab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Required Item Should be Double Clicked in home Tab

'Examples				:				Fn_CM_CompnentTabOperations("Activate", "Home")  

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane		21-09-2011			1.0								No					Sunny R
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_CompnentTabOperations(sAction,sTCComponentTabName) 
	GBL_FAILED_FUNCTION_NAME="Fn_CM_CompnentTabOperations"
	Dim iTabCount, iTabIndex, sTabVal, bTabActive, objJavaTab,tabVal,iCounter,oTabCtl,iCount,objItem,objTabFld,i,iCnt
	Dim sxLen,syLen,sBounds,aBounds, aMenuList, StrMenu
	Select Case sAction
				Case "VerifyActivate"	'Updated By Harshal Agrawal on 07 Dec 2010
								Set objTabFld = JavaWindow("ChangeManager").JavaObject("TCComponentTab")
								i = JavaWindow("ChangeManager").JavaObject("TCComponentTab").Object.getSelectedTabIndex
								set objItem = objTabFld.Object.getItem(i)
								If trim(objItem.text) = trim(sTCComponentTabName) Then
										Fn_CM_CompnentTabOperations = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified that Tab ["+sTCComponentTabName+"] is Activated.")
								Else
										Fn_CM_CompnentTabOperations = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify that Tab ["+sTCComponentTabName+"] is Activated.")
								End If
								Set objItem = Nothing
								Set objTabFld = Nothing
				
				Case "Activate" 		'Updated By Harshal Agrawal on 07 Dec 2010
								Set oTabCtl = Nothing
								JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").SetTOProperty "index",0
								 Set oTabCtl = JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget")
								  For iCount = 0 to Cint(oTabCtl.Object.getTabItemCount)-1
										  If InStr(1, oTabCtl.Object.getItem(iCount).text, sTCComponentTabName, vbTextCompare) Then
												  oTabCtl.Object.setSelectedTabAndNotifyListeners CInt(iCount), true
												 Fn_CM_CompnentTabOperations = True
												 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_CM_CompnentTabOperations Execution Sucessful")
												 Exit For
											Else
												Fn_CM_CompnentTabOperations = False
												 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_CM_CompnentTabOperations Execution Falied Tab not Found")
											End If
									 Next
									 oTabCtl.SetTOProperty "index",1
									 Set oTabCtl = Nothing

				Case "GetTabSequence" 	'Added By Ketan on 14-Feb-2011
								sTabVal = ""
								Set oTabCtl = Nothing
								JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").SetTOProperty "index",0
								 Set oTabCtl = JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget")
								  For iCount = 0 to Cint(oTabCtl.Object.getTabItemCount)-1
									  If iCount = 0 Then
										  sTabVal = oTabCtl.Object.getItem(iCount).text
									  Else
										  sTabVal = sTabVal+":"+oTabCtl.Object.getItem(iCount).text
									  End If									 									 
								  Next
								  Fn_CM_CompnentTabOperations = sTabVal
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_CM_CompnentTabOperations Execution Sucessful")
									 oTabCtl.SetTOProperty "index",1
									 Set oTabCtl = Nothing

				Case  "Close" 'Updated By Harshal Agrawal on 15 Dec 2010
								Set objTabFld = JavaWindow("ChangeManager").JavaObject("TCComponentTab")
								iCnt = objTabFld.Object.getTabItemCount
								sxLen = 0
								syLen = 0
								For i = 0 to iCnt-1
										 set objItem = objTabFld.Object.getItem(i)
										 sxLen = sxLen + objItem.getWidth
										 If trim(objItem.text) = trim(sTCComponentTabName) Then
											   sxLen = sxLen - (objItem.getWidth/2)
											   syLen = (objItem.getHeight/2)
											   objTabFld.Click sxLen, syLen, "LEFT"
											   Fn_ReadyStatusSync(2)
											   sBounds = objItem.getCloseButtonBounds.toString()
											   sBounds = right(sBounds, Len(sBounds)-instr(sBounds, "{"))
											   aBounds = split(sBounds, ",", -1, 1)
											   sxLen = Cint(trim(aBounds(0))) + 5
											   syLen = Cint(trim(aBounds(1))) + 5
											   objTabFld.Click sxLen, syLen, "LEFT"
											   Exit For
										 End If
								Next
								If Err.Number < 0 Then
								  Fn_CM_CompnentTabOperations = False
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Close ["+sItem+"] Tab.")
								Else
								  Fn_CM_CompnentTabOperations = True
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Closed ["+sItem+"] Tab.")
								End If
								Set objItem = Nothing
								Set objTabFld = Nothing

		Case  "CheckDuplicate"

        							'Check Requested Tab is exist or not.
									iCounter = 0
									For iTabIndex = 0 to iTabCount-1
											Call Fn_UI_JavaTab_Select("Fn_CM_CompnentTabOperations",JavaWindow("ChangeManager"),"NavTreeTab", "#"&iTabIndex)
											sTabVal= Fn_UI_Object_GetROProperty("Fn_CM_CompnentTabOperations",objJavaTab,"value")
											If sTabVal = sTCComponentTabName Then
												iCounter = iCounter + 1
											End If
									Next
									If iCounter > 1 Then
										'Function returns True
										Fn_CM_CompnentTabOperations = True
										'Call Log file
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTCComponentTabName+ "] : Tab has "+iCounter+" Instances ")
									Else 
										'Function returns False
										Fn_CM_CompnentTabOperations = False
										'Call Log file
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: [" +sTCComponentTabName+ "] : Tab has no duplicate instance ")										
									End If

				Case "TabRMBMenuSelect" 
								Call Fn_UI_JavaObject_Click("Fn_CM_CompnentTabOperations", JavaWindow("DefaultWindow"), "RACTabFolderWidget",1,1,"RIGHT")
								aMenuList = split(sTCComponentTabName, ":",-1,1)
								iTabCount = Ubound(aMenuList)
								'Select Menu action
								Select Case iTabCount
												Case "0"
													 StrMenu = JavaWindow("ChangeManager").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
												Case "1"
													StrMenu = JavaWindow("ChangeManager").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
												Case "2"
													StrMenu = JavaWindow("ChangeManager").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
												Case Else
													Fn_CM_CompnentTabOperations = FALSE
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Invalid PopupMenu  [" +sTCComponentTabName+ "] is Requested.")
													Exit Function
								End Select
								JavaWindow("ChangeManager").WinMenu("ContextMenu").Select StrMenu
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully selected [" +sTCComponentTabName+ "] PopupMenu.")
								Fn_CM_CompnentTabOperations = TRUE									
									
				Case Else
								 Fn_CM_CompnentTabOperations = FALSE
								'Call Log file
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Invalid ACTION  [" +sAction+ "] is Requested.")
	End Select
	Set objJavaTab = Nothing
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_GetTreeItemPath(objTree, sNode, sDelimiter, sInstanceHandler)

'Description		:	Function used to Perform operations on Component Tab of Change Manager


'Parameters			:	1) objTree: Object of a tree.
'						2) sNode: Node Path
'						3) sDelimiter: for future use.
'						4) sInstanceHandler: for future use.
											
'Return Value		: 	item path \ FALSE

'Pre-requisite		:	change Manager perspective should be present

'Examples			:	Call Fn_CM_GetTreeItemPath(JavaWindow("ChangeManager").JavaTree("ViewTree"), "Change Home:My Open Changes:PR-000003/A;1-test", "", "")

'History:
'	Developer Name			Date			Rev. No.		sChanges Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe		06-Apr-2012			1.0				Created
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_GetTreeItemPath(objTree, sNode, sDelimiter, sInstanceHandler)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_GetTreeItemPath"
	Dim iCnt, aNodeArr, iArrCnt, iItemCnt
	Dim objCurrTreeItm, sTreeNodeToStr2
	Dim sItmPath
	If sDelimiter = "" Then
		sDelimiter = ":"
	End If

	Fn_CM_GetTreeItemPath = False

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
			Fn_CM_GetTreeItemPath = False
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
			Fn_CM_GetTreeItemPath = False
			Exit Function
		End If
	Next
	set objCurrTreeItm = Nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Retrieved Node Index" + sItmPath + "of Change Manager Tree Node [" + sNode + "]" )	
	Fn_CM_GetTreeItemPath = sItmPath
End Function
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@
'@@    Function Name		:	Fn_CM_getJavaTreeIndex
'@@
'@@    Description			:	Function Used to retrieve Index Of Java Tree Node
'@@
'@@    Parameters			:	1. objTree: Tree object
'@@											  2. StrNode: Full Node 
'@@								
'@@    Return Value		   	: 	Node index / -1
'@@
'@@    Pre-requisite		:	Tree Should Exist							
'@@
'@@    Examples				:	Call Fn_CM_getJavaTreeIndex(JavaWindow("MyTeamcenter").JavaTree("NavTree"), "Home:Newstuff:000136-new")
'@@								Call Fn_CM_getJavaTreeIndex(JavaWindow("MyTeamcenter").JavaTree("SearchResultTree"),  "Item... (1):000108-Top")
'@@
'@@	   History					 	:	
'@@					Developer Name			Date				Rev. No.		Changes Done								Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe			08-12-2011			1.0				Created
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@					Koustubh Watwe			04-Apr-2012			2.0				Created
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_CM_getJavaTreeIndex(objTree, sNode)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_getJavaTreeIndex"
	Dim iPath, arriPath, iArrCnt, iNodeCnt, iLimit, obj
	Fn_UI_getTreeIndex_iGblCnt = -1
	Fn_UI_getTreeIndex_CompareString = ""
	Fn_UI_getTreeIndex_bFound = False
	iPath = Fn_CM_GetTreeItemPath(objTree, sNode ,"","")
	If iPath = False Then
		Fn_UI_getJavaTreeIndex = Fn_UI_getTreeIndex_iGblCnt
		exit function
	End If
	Fn_UI_getTreeIndex_iGblCnt = 0
	iPath=Replace(iPath,"#","")
	arriPath=Split(iPath,":")
	For iArrCnt = 0 to UBound(arriPath)-1
		arriPath(iArrCnt) = cInt(arriPath(iArrCnt))
		If iArrCnt = 0 Then
			Set obj = objTree.Object.getItem(arriPath(iArrCnt))
		Else
			Set obj = obj.getItem(arriPath(iArrCnt)) 
		End If
		iLimit = cInt(arriPath(iArrCnt + 1))
		If iArrCnt <> uBound(arriPath) Then
			iLimit = iLimit - 1 
		End If
		For iNodeCnt = 0 to iLimit
			Call Fn_UI_getTreeIndex(obj.getItem(iNodeCnt), "")
		Next
		Fn_UI_getTreeIndex_iGblCnt = Fn_UI_getTreeIndex_iGblCnt + 1
	Next
	Fn_CM_getJavaTreeIndex = Fn_UI_getTreeIndex_iGblCnt
End Function


'*********************************************************		Generic function to handle Error dialogs   	***********************************************************************
'Function Name		:				Fn_SISW_CM_ErrorVerify()

'Description			 :		 		 The function is generic function to handle error dialogs. It is created after combining error dialog functions from ChangeManager.vbs
'										Fn_CM_ErrorMessageVerify
'										Fn_CM_ChangeErrorMsgVerify
'										Fn_CM_ErrorWindowMsgVerify

'Parameters			   :	 			1.  dicErrorInfo											
'Return Value		   : 				True/False
'Pre-requisite			:		 		NA.
'Examples				:				
'									Set dicErrorInfo = CreateObject("Scripting.Dictionary")
'									With dicErrorInfo	
'										.Add "Title", "Error"
'										.Add "Message", "The operation failed on one or more of the selected objects"
'										.Add "Button", "OK"										
'									End with
'									bReturn = Fn_SISW_CM_ErrorVerify(dicErrorInfo)

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sushma Pagare          28-Jun-2013
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_SISW_CM_ErrorVerify(dicErrorInfo)
			GBL_FAILED_FUNCTION_NAME="Fn_SISW_CM_ErrorVerify"
			Dim  dicKeys, dicItems, iCounter
			Dim sAction, sTitle, sErrorMsg,sButton, sAppMsg
			Dim descStaticText, objStaticText,objErrorDialog
            			
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

			Fn_SISW_CM_ErrorVerify = FALSE
            On Error Resume Next

			Select Case sAction

				''This covers  Fn_CM_ErrorMessageVerify(sDilogName,sErrorMessage,sButtonName) and Also
				' '         Fn_CM_ChangeErrorMsgVerify(strAction,sDilogName,sErrorMessage,sButtonName)   - Case "StaticMsgVerify"
				Case "ErrorMessageVerify","StaticMsgVerify"					
    					
						JavaWindow("ChangeManager").JavaWindow("ErrorDialog").SetTOProperty "title",sTitle
						JavaWindow("ErrorDialog").SetTOProperty "title",sTitle
						'Added by Nilesh on 4-Jul-2013
                        JavaWindow("MyTeamcenter").JavaWindow("Error").SetTOProperty "title",sTitle
						If JavaWindow("ChangeManager").JavaWindow("ErrorDialog").Exist(8) Then
							Set objErrorDialog=JavaWindow("ChangeManager").JavaWindow("ErrorDialog")
						Elseif JavaWindow("ErrorDialog").Exist(8) then
							Set objErrorDialog=JavaWindow("ErrorDialog")
'						Added by Nilesh on 4-Jul-2013
						Elseif JavaWindow("MyTeamcenter").JavaWindow("Error").Exist(8) then
							Set objErrorDialog=JavaWindow("MyTeamcenter").JavaWindow("Error")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sTitle & " Error not appear")
							Exit function
						End If
						
						'Getting the label of  ErrMsg in sAppMsg
						Set descStaticText=Description.Create()
						descStaticText("Class Name").value="JavaStaticText"
						'Taking Child object of present Error Dialog box
						Set objStaticText=objErrorDialog.ChildObjects(descStaticText)								
						For iCounter=0 to objStaticText.count-1
							'Checking Error message 
							sAppMsg=objStaticText(iCounter).getROProperty("text")
							If Instr(1,Lcase(sAppMsg),Lcase(sErrorMsg))>0 Then
								objErrorDialog.JavaButton("OK").SetTOProperty "label",sButton
								'Clicking on ok button
								Call Fn_Button_Click("Fn_SISW_CM_ErrorVerify",objErrorDialog,"OK")
								Fn_SISW_CM_ErrorVerify=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Successfully Verified Message : "+ sErrorMsg)
								Exit For
							Else
								GBL_ACTUAL_MESSAGE=sAppMsg
								Fn_SISW_CM_ErrorVerify=False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Failed to verify Message : "+ sErrorMsg)
							End If
						Next				
						Set descStaticText=Nothing
						Set objStaticText=Nothing
						Set objErrorDialog=Nothing 
						Exit Function    

					'' This covers Fn_CM_ChangeErrorMsgVerify(strAction,sDilogName,sErrorMessage,sButtonName)   - Case "EditBoxMsgVerify"
				Case "EditBoxMsgVerify"
							JavaWindow("MyTeamcenter").JavaWindow("Error").SetTOProperty "title",sTitle
							'Checking Error Dialog Exist or not 
							If  JavaWindow("MyTeamcenter").JavaWindow("Error").Exist(3) =False  Then	
										Fn_SISW_CM_ErrorVerify=False
										Exit Function
							End If	
							sAppMsg = JavaWindow("MyTeamcenter").JavaWindow("Error").JavaEdit("Details").getROProperty("value")
							If Instr(1,Lcase(sAppMsg),Lcase(sErrorMsg))>0 Then
								JavaWindow("MyTeamcenter").JavaWindow("Error").JavaButton("OK").SetTOProperty "label",sButton
								Call Fn_Button_Click("Fn_SISW_CM_ErrorVerify",JavaWindow("MyTeamcenter").JavaWindow("Error"),"OK")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Sucessfully Verified Message : "+ sErrorMsg)
								Fn_SISW_CM_ErrorVerify=True
							Else
								GBL_ACTUAL_MESSAGE=sAppMsg
								Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Failed to Verify Message : "+ sErrorMsg)
								Fn_SISW_CM_ErrorVerify=False
							End If											
							Exit Function

''					This Case covers Fn_CM_ErrorWindowMsgVerify(sDilogName,sErrorMessage,sButtonName)
					Case "ErrorWindowMsgVerify"

							JavaWindow("ChangeManager").JavaWindow("Error_Paste").JavaDialog("Paste").JavaDialog("Error").SetTOProperty "title",sDialogName
							Window("TeamcenterWindow").JavaDialog("New Schedule").JavaDialog("New Schedule").SetTOProperty "title",sDialogName
							JavaWindow("ChangeManager").JavaWindow("ErrorWindow").SetTOProperty "title" ,sTitle
							If JavaWindow("ChangeManager").JavaWindow("Error_Paste").JavaDialog("Paste").JavaDialog("Error").Exist(3) Then
									Set ObjErrDialog =JavaWindow("ChangeManager").JavaWindow("Error_Paste").JavaDialog("Paste").JavaDialog("Error")
									sAppMsg = ObjErrDialog.JavaEdit("Error_msg").GetROProperty("value")
							ElseIf Window("TeamcenterWindow").JavaDialog("New Schedule").JavaDialog("New Schedule").Exist(3) Then
									Set ObjErrDialog=Window("TeamcenterWindow").JavaDialog("New Schedule").JavaDialog("New Schedule")
									sAppMsg = ObjErrDialog.JavaEdit("Error_msg").GetROProperty("value")
							ElseIf JavaWindow("ChangeManager").JavaWindow("ErrorWindow").Exist(2) Then
							  Set ObjErrDialog=JavaWindow("ChangeManager").JavaWindow("ErrorWindow")
							    sAppMsg =JavaWindow("ChangeManager").JavaWindow("ErrorWindow").JavaStaticText("ErrMSgText").GetROProperty("label")
							    
							    If sAppMsg = "" Then
							     	 sAppMsg =  JavaWindow("ChangeManager").JavaWindow("ErrorWindow").JavaStaticText("Error_static1").GetROProperty("label")
							    Else
									Fn_SISW_CM_ErrorVerify=False
									Exit Function
								End If
                                       


							Else
									Fn_SISW_CM_ErrorVerify=False
									Exit Function
							End If
							If sErrorMsg <> ""  Then
									If Instr(1,Lcase(sAppMsg),Lcase(sErrorMsg))>0 Then
                                        objErrorDialog.JavaButton("OK").SetTOProperty "label",sButton
										Call Fn_Button_Click("Fn_SISW_CM_ErrorVerify",ObjErrDialog,"OK")
										Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Sucessfully Verified Message : "+ sErrorMsg)
										Fn_SISW_CM_ErrorVerify=True
									Else
										GBL_ACTUAL_MESSAGE=sAppMsg
										Call Fn_WriteLogFile(Environment.Value("TestLogFile")," Failed to Verify Message : "+ sErrorMsg)
										Fn_SISW_CM_ErrorVerify=False
									End If                                          									
							Else
									Fn_SISW_CM_ErrorVerify=True
							End If
							Set objErrorDialog=Nothing
							Exit Function                                          
													
			End Select
			
End Function
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@    Function Name		:	Fn_CM_SummaryTabTableOperations
'@@
'@@    Description			:	Function Used to perform operation on Tables in summary tab
'@@
'@@    Parameters			:	1.sTab = Tab Name
'@@								2. sAction = Action Name
'@@								3. sObjColName = Main object Name on which other values of column verification is done
'@@								4. sObjColVal = Main object Name on which other values of column verification is done 
'@@								5. sColName = Column Name 
'@@								6.sColVal = Column value 
'@@								7.sMenu = for future operation like popup or any menu operation	
'@@
'@@    Return Value		   	: 	True/False
'@@
'@@    Pre-requisite		:	Table Should Exist							
'@@
'@@    Examples				:	Call Fn_CM_SummaryTabTableOperations("Change History","Verify_ChangeHistory_CellData","Solution Item","000755/B;1-ItemTest","Impacted Item","000755/A;1-ItemTest","")
'@@
'@@	   History				:	Developer Name			Date			Rev. No.		Changes Done			Reviewer
'@@					
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@								Poonam Chopade		 30-Oct-2017		1.0				Created				TC11.4(20171023.00)_NewDevelopment_PoonamC_30Oct2017
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_SummaryTabTableOperations(sTab,sAction,sObjColName,sObjColVal,sColName,sColVal,sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_SummaryTabTableOperations"
	
	Dim ObjChangeWindow,iRows,sColsNames,aCols,iCnt,iColIndex,sAppValue,iRowIndex,ObjDefaultWindow
	Set ObjChangeWindow = Fn_SISW_CM_GetObject("ChangeManager")
	Fn_CM_SummaryTabTableOperations = False
	Set ObjDefaultWindow = JavaWindow("DefaultWindow")
	'Select Change History Tab
	Call Fn_UI_JavaTab_Select("Fn_MyTc_ItemSummaryOperation",ObjDefaultWindow,"GeneralInnerTab",sTab)
	Call Fn_ReadyStatusSync(1)
	
	Select Case sAction
		Case "Verify_ChangeHistory_CellData"
			iRows = ObjChangeWindow.JavaTable("ChangeHistoryTable").GetROProperty("rows")
			sColsNames = ObjChangeWindow.JavaTable("ChangeHistoryTable").GetROProperty("column names")
			aCols = Split(sColsNames,";")
			For iCnt = 0 To UBound(aCols)
				If trim(sObjColName) = trim(aCols(iCnt)) Then
					iColIndex = iCnt
					Exit For
				End If
			Next
			'Get Row Index of Object Value to verify column value
			For iCnt = 0 To cint(iRows)-1
				sAppValue = ObjChangeWindow.JavaTable("ChangeHistoryTable").object.getitem(iCnt).getData().getcellcomps(iColIndex).tostring()
				sAppValue = Replace(Replace(sAppValue,"[",""),"]","")
				If trim(sAppValue) = trim(sObjColVal) Then
					iRowIndex = iCnt
					Exit For
				End If
			Next
			'Get Column Index to verify value
			For iCnt = 0 To UBound(aCols)
				If trim(sColName) = trim(aCols(iCnt)) Then
					iColIndex = iCnt
					Exit For
				End If
			Next
			'Get Column Value
			sAppValue = ObjChangeWindow.JavaTable("ChangeHistoryTable").object.getitem(iRowIndex).getData().getcellcomps(iColIndex).tostring()
			sAppValue = Replace(Replace(sAppValue,"[",""),"]","")
			If trim(sAppValue) = trim(sColVal) Then
				Fn_CM_SummaryTabTableOperations = True
			Else
				Fn_CM_SummaryTabTableOperations = False
			End if
End Select

Set ObjChangeWindow = Nothing

End Function
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@    Function Name		:	Fn_CM_ChangeInContext_Operations
'@@
'@@    Description			:	Function Used to perform operation on New Change Context dialog
'@@
'@@    Parameters			:	1.sAction = Action Name
'@@								2. dicChangeInfo = dictionary object contains change info
'@@								3. sButtons = buttons name
'@@
'@@    Return Value		   	: 	True/False				
'@@
'@@    Examples				:	Set dicChangeInfo = CreateObject("Scripting.Dictionary")
'@@									dicChangeInfo("ChangeType") = "Change Notice"
'@@									dicChangeInfo("ChangeID") = "ECN-123456"
'@@									dicChangeInfo("ChangeRev") = "A"
'@@									dicChangeInfo("ChangeName") = "CN1"
'@@									dicChangeInfo("ChangeDescription") = "CN Desc"
'@@									dicChangeInfo("ProcessTemplateList") = "AutoSimpleReview"
'@@									dicChangeInfo("ProcessAssignmentList") = ""
'@@								bReturn = Fn_CM_ChangeInContext_Operations("Create",dicChangeInfo,"Finish:Cancel")
'@@
'@@	   History				:	Developer Name			Date			Rev. No.		Changes Done			Reviewer
'@@					
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@								Poonam Chopade		 12-Dec-2017		1.0				Created				TC11.4(20171201.00)_NewDevelopment_PoonamC_12Dec2017
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_ChangeInContext_Operations(sAction,dicChangeInfo,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_ChangeInContext_Operations"
   'Declaring Variables
   Dim strNodePath,intNodeCount,sTreeItem,intCount,iRowCount
   Dim ObjChangeWnd,iCount,cFileName,objSelectType,objIntNoOfObjects
	Fn_CM_ChangeInContext_Operations = False
	
	Select Case sAction
		Case "Create"
			'Verifying "New Change In context" window's existance
			If Fn_UI_ObjectExist("Fn_CM_ChangeInContext_Operations",JavaWindow("MyTeamcenter").JavaWindow("NewChangeInContext"))=False Then
				Exit Function	
			End If
			'Creating Object of "New Change In context" window
			Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_CM_ChangeInContext_Operations",JavaWindow("MyTeamcenter").JavaWindow("NewChangeInContext"))
			If dicChangeInfo("SearchText")<>"" Then
				'Setting Deviation Change Type (This Text box is appear only in Deviation Request Case)
				Call Fn_UI_EditBox_Type("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"SearchText",dicChangeInfo("SearchText"))
			End If
			Call Fn_JavaTree_Select("Fn_CM_ChangeInContext_Operations", ObjChangeWnd, "ChangeType","Complete List")
			strNodePath="Complete List:"+dicChangeInfo("ChangeType")
			'Verifying Node is present in Tree
		    intNodeCount =Fn_UI_Object_GetROProperty("Fn_CM_ChangeInContext_Operations",ObjChangeWnd.JavaTree("ChangeType"),"items count")    
			For intCount = 0 to intNodeCount - 1
				sTreeItem = ObjChangeWnd.JavaTree("ChangeType").GetItem(intCount)
				If Trim(lcase(sTreeItem)) = Trim(Lcase(strNodePath)) Then
					Fn_CM_ChangeInContext_Operations = True
					Exit For
				End If
			Next
			If Cint(intCount) =Cint( intNodeCount) Then
				Call Fn_Button_Click("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"Cancel")
				Fn_CM_ChangeInContext_Operations =False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& strNodePath &"Node is not found in View Tree")
				Set ObjChangeWnd=Nothing
				Exit Function
			End If
			
			'Selecting Node from tree
		   Call Fn_JavaTree_Select("Fn_CM_ChangeInContext_Operations", ObjChangeWnd, "ChangeType",strNodePath)
		   Wait 3
			'Clicking on Next button to proceed 
			Call Fn_Button_Click("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"Next")
			Wait 3
			If dicChangeInfo("ChangeID")<>"" Then
				'Setting Id
				Call Fn_Edit_Box("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"Id",dicChangeInfo("ChangeID"))
			End If
			
			If dicChangeInfo("ChangeRev")<>"" Then
				'Setting Revision
				Call Fn_Edit_Box("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"Revision",dicChangeInfo("ChangeRev"))
			End If
			
			'Setting Name
			Call Fn_Edit_Box("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"Name",dicChangeInfo("ChangeName"))
			
			'Setting Description
			Call Fn_Edit_Box("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"Description",dicChangeInfo("ChangeDescription"))
			
			If dicChangeInfo("ProcessTemplateList") <> "" or dicChangeInfo("ProcessAssignmentList") <> "" Then
				'Clicking on Next button to proceed 
				Call Fn_Button_Click("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"Next")
				Call Fn_ReadyStatusSync(1)
				
				'select process template list
				If Fn_UI_ObjectExist("Fn_CM_ChangeInContext_Operations",JavaWindow("MyTeamcenter").JavaWindow("NewChangeInContext").JavaList("ProcessTemplateList")) Then
					Call Fn_SISW_UI_JavaList_Operations("Fn_CM_ChangeInContext_Operations","Select",ObjChangeWnd,"ProcessTemplateList", dicChangeInfo("ProcessTemplateList"), "", "")
					wait 1
				Else
					Call Fn_Button_Click("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,"SrchButton")
					Set objSelectType = description.Create()
					objSelectType("Class Name").value = "JavaTable"	
					objSelectType("class_path").value = ".*Table.*"
					Set objIntNoOfObjects = JavaWindow("MyTeamcenter").JavaWindow("Shell").ChildObjects(objSelectType)
					iRowCount = objIntNoOfObjects(0).GetROProperty("rows")
					For iCount=0 to iRowCount-1
						cFileName= objIntNoOfObjects(0).GetCellData(iCount,0)
						If trim(cFileName)= trim(dicChangeInfo("ProcessTemplateList")) Then	
							objIntNoOfObjects(0).ActivateRow iCount
							objIntNoOfObjects(0).PressKey " "
							Set objSelectType = Nothing
							Set objIntNoOfObjects = Nothing
							Exit For
						End If
					Next
				End If
				
				'If dicChangeInfo("ProcessTemplateList") <> "" Then
					'Call Fn_SISW_UI_JavaList_Operations("Fn_CM_ChangeInContext_Operations","Select",ObjChangeWnd,"ProcessTemplateList", dicChangeInfo("ProcessTemplateList"), "", "")
					'wait 1
				'End If
				
				'select process Assigment list
				If dicChangeInfo("ProcessAssignmentList") <> "" Then
					Call Fn_SISW_UI_JavaList_Operations("Fn_CM_ChangeInContext_Operations","Select",ObjChangeWnd,"ProcessAssignmentList", dicChangeInfo("ProcessAssignmentList"), "", "")
					wait 1
				End If
			End If
	End Select
	
	If sButtons <> "" Then
		sButtons = Split(sButtons,":")
		For iCount = 0 To UBound(sButtons)
			'Clicking On button
			Call Fn_Button_Click("Fn_CM_ChangeInContext_Operations",ObjChangeWnd,sButtons(iCount))
			Call Fn_ReadyStatusSync(1)
		Next
	End If
	
	Fn_CM_ChangeInContext_Operations=True
	'Releasing "New Change" window's object
	Set ObjChangeWnd=Nothing
End Function


'-------------------------------------------------------------------Function Used to perform operatons on Change BOM Tree---------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_CM_BOMTreeOperations

'Description			 :	Function Used to perform operatons on View Tree

'Parameters			   :	1.strAction: Action Name
										'2.strNodeName: Node Name		
										'3.strMenu: Pop Up Menu Name
										
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be Log In present on Change Manager Perspective

'Examples				:	Fn_CM_BOMTreeOperations("Select","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
										'Fn_CM_BOMTreeOperations("Expand","Change Home:My Open Changes","")
										'Fn_CM_BOMTreeOperations("PopupMenuSelect","Change Home:Test1","Refresh")
										'Fn_CM_BOMTreeOperations("Collapse","Change Home:My Open Changes","")
'History					 :			
'	Developer Name				Date			Rev. No.	Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sandeep N					05/08/2010		1.0												Tushar B
'	Sandeep N					19/01/2011		1.0			Added Case "GetNodeIndex"		Harshal A
'	Sandeep N					20/01/2011		1.0			Added Case "Collapse"		Harshal A
'	Sandeep N					15/11/2011		2.0			Modified All case by adding "Fn_UI_JavaTreeGetItemPath" UI function
'	Koustubh W					06/04/2012		3.0			Modified All case by adding "Fn_CM_GetTreeItemPath" function
'	Sandeep N					16/04/2012		4.0			Modified case : MultiSelectPopupMenu
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_CM_BOMTreeOperations(strAction,StrNodePath,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_CM_BOMTreeOperations"
   'Variable declaration
   Dim arrNodeName,ItemCount,iCount,sNode, aMenuList,intCount
   Dim intNodeCount,sTreeItem,NodeName,StrMultiNodePath,arrMultiNodeName,iCnt
   Dim ObjChMgrWnd,iPath
	Fn_CM_BOMTreeOperations = False
	'Creating Object of ChangeManager window
	Set ObjChMgrWnd=Fn_UI_ObjectCreate("Fn_CM_BOMTreeOperations", JavaWindow("ChangeManager"))
	
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
			iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("BOMTree"),sNode,"","")
			If iPath <> False Then ObjChMgrWnd.JavaTree("BOMTree").Expand iPath
			wait 1
		End IF
	End IF
	
	Select Case strAction
		'===================================================================================================
		Case "Select" 		'Fn_CM_BOMTreeOperations("Select","Change Home:My Open Changes:PR-000006/A;1-Test PR","")
				iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("BOMTree"),StrNodePath,"","")
				If iPath<>False Then
					ObjChMgrWnd.JavaTree("BOMTree").Select iPath
					Fn_CM_BOMTreeOperations=True
				End If
		'===================================================================================================
		Case "Expand" 'Fn_CM_BOMTreeOperations("Expand","Change Home:My Open Changes","")
				iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("BOMTree"),StrNodePath,"","")
				Call Fn_UI_JavaTree_Expand("Fn_CM_BOMTreeOperations", ObjChMgrWnd, "BOMTree",iPath)
				Fn_CM_BOMTreeOperations=True
		'===================================================================================================
		Case "PopupMenuSelect","PopupMenuSelectExt"		'Fn_CM_BOMTreeOperations("PopupMenuSelect","Change Home:Test1","Refresh")
				If strAction = "PopupMenuSelectExt" Then
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_CM_BOMTreeOperations",ObjChMgrWnd.JavaTree("BOMTree"),StrNodePath,"","@")
				Else
					iPath = Fn_CM_GetTreeItemPath(ObjChMgrWnd.JavaTree("BOMTree"),StrNodePath,"","")
				End If
				
				If iPath<>False Then
					ObjChMgrWnd.JavaTree("BOMTree").Select iPath
					'Open context menu
					wait 1
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_CM_BOMTreeOperations",ObjChMgrWnd, "BOMTree",iPath)
					Wait 2
					'Select Menu action
					aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Select Case intCount
						Case "0"
							 StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu =ObjChMgrWnd.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_CM_BOMTreeOperations = False
							Set ObjChMgrWnd=Nothing
							Exit Function
					End Select
					ObjChMgrWnd.WinMenu("ContextMenu").Select StrMenu
					Fn_CM_BOMTreeOperations=True
				End If				
		'===================================================================================================
		Case Else
				Fn_CM_BOMTreeOperations=False
	End Select
	'Rleasing Change Manager Window Object
	Set ObjChMgrWnd=Nothing
End Function




