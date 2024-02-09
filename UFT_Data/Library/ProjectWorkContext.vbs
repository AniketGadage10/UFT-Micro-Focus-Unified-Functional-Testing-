Option Explicit
'----------------------------'Global variables for Teamcenter Perspective Names----------------------------------------------------------
Public GBL_PERSPECTIVE_PROJECT
GBL_PERSPECTIVE_PROJECT = "Project"
'----------------------------'Global variables for Teamcenter Perspective Names----------------------------------------------------------
'=======================================================================================================================================================
' Function List
'=======================================================================================================================================================
'0.  Fn_SISW_PWC_GetObject(sObjectName)
'1. Fn_PWC_TabOperation(sAction, sTabName)
'2. Fn_PWC_TabSet(StrTabName)
'3. Fn_PWCRules_TreeOpeartion(sAction,sNodeName, sMenu)
'4. Fn_PWCProject_TreeOpeartion(sAction,sNodeName, sMenu)
'5. Fn_PWCOrgGrp_TreeOpeartion(sAction,sNodeName, sMenu)
'6. Fn_PWCSelMem_TreeOpeartion(sAction,sNodeName, sMenu)
'7. Fn_PWC_MemberSelection(sAction, sNodeName, sExtra)
'8. Fn_PWC_ProjectOperations(sAction, sID, sName, sDescription, sStatus, sSecurity, sMemberAction, sNodeName, sButton)
'9. Fn_PWC_CheckButtonsdisabled(sReferencePath,sButtons)
'10.Fn_PWC_NonProjectAdministratorMsgVerify(sErrorText)
'11.Fn_PWC_TeamAdmin(sUser,sButtons)
'12.Fn_PWC_SearchObjectVerify(sAction,sMsg,sbutton)
'13.Fn_PWC_WorkContextMessageVerify(sAction, sMessage)   - Eliminated. Not used anywhere. By Sushma Pagare [ 8-Jul-13]
'14.Fn_PWC_AssignWorkContext(sAction, sWorkCnxt, sButtons)
'15.Fn_PWC_PropertyOperations(sAction, bSubGroup, SComments, sDateArch, sDateCreated, sDateLstBack, sDesc, sGroup, sGroupID, sGroupMem, sLstModiDate, sLstModiUser, sName1, sName2, sObject, sOwner, sOwningSite, sProject, sRole, sUserSetModi, sButtons)
'16.Fn_PWC_ObjectPropertyIsEditable(sAction, sPropertyName, sButtons)
'17.Fn_PWC_DialogMsgVerify(sTitle,sMsg,sButton) 
'18.Fn_PWC_CheckOutMessageVerify(sAction, aObject, aMessage)
'19.Fn_PWC_SmartFolder_TreeOpeartion(sAction,sNodeName)
'20.Fn_PWC_InsufficentPrivilegeErrorVerify(sTitle,sDetails)
'21.Fn_PWC_FilterAssociationTable_Operations(sAction, sName, sSourceType, sProperty, sValue, iRow, iColumn)
'22.Fn_ProjectTabOperations(sAction,sTabName)
'23.Fn_PWC_CumulativeTable_Operations(sAction, bContribute, sName, sSourceType, sProperty, sValue, iRow, iColumn)
'24.Fn_PWC_ItemDetailsCreate(sSelectType, bConfItem, sItemInfo, sAddItemInfo, sAddItemRevInfo, sAttachFileInfo, sWorkFlowInfo, sIdentifierBasicInfo, sAddIDInfo, sAddRevInfo, sAssignProj, sDefineOptions, sButtons)
'25.Fn_PWC_SelectedMemberNodePath(sUserToSelect)
'26.Fn_PWC_LibraryTreeOpeartion(sAction,sNodeName, sMenu)
'27.Fn_SISW_PWC_ErrorVerify(dicErrorInfo)
'28.Fn_PWC_ChangeOwnership(sAction,sNodeName,sNewOwningUser)
'=======================================================================================================================================================

'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_PWC_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_PWC_GetObject("Project - Teamcenter 8")

'History:
'	Developer Name			Date						Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Shreyas 		     17-Jly-2012							1.0								
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_PWC_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ProjectWorkContext.xml"
	Set Fn_SISW_PWC_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'*********************************************************		Function to select  the Tab into Project		**********************************************************************
'Function Name		:				Fn_PWC_TabOperation

'Description			 :		 		 This Tab includes Definition Tab,AM Rule Tab
'													1)Click on Tab
'													2)Verify the Tab is open
'													3)Close the Tab

'Parameters			   :	 			1) sAction: Action to be performed on the Tab
'													 2) sTabName: Tab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Project prespective should be displayed.

'Examples				:				Fn_PWC_TabOperation("Activate", "AM Rules")  

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					05-08-10			1.0																Harshal	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_PWC_TabOperation(sAction, sTabName)  
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_TabOperation"
	Dim iTabCount, iTabIndex, sTabVal	
	Dim objJavaTab
'	Set objJavaTab =	Fn_UI_ObjectCreate("Fn_PWC_TabOperation", JavaWindow("Project - Teamcenter 8").JavaObject("RACTabFolderWidget"))	
'	Set objJavaTab =	Fn_UI_ObjectCreate("Fn_PWC_TabOperation", JavaWindow("Project - Teamcenter 8").JavaTab("ConfigTab"))
    Select Case sAction
				Case "VerifyActivate"
								'Check Weather Requested tab is open or not( Activated or Not )
'                              	If  sTabName = objJavaTab.GetROProperty("value")   Then							
'										'Call Log file
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTabName+ "] : Tab is Activate(Open) ")
'										Fn_PWC_TabOperation = TRUE
'								else							
'									'Call Log file
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [" +sTabName+ "] : Tab is NOT Activate (Open)")	
'									Fn_PWC_TabOperation = FALSE
'								End If
								bFlag = False
								bFlag = Fn_TabFolder_Operation("VerifyActive", sTabName, "")
								
								If bFlag = False Then
									Fn_PWC_TabOperation = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify ["+sItem+"] Tab.")
								Else
									Fn_PWC_TabOperation = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified ["+sItem+"] Tab.")
								End If
				Case "Activate" 
								'If True =  Fn_PWC_TabSet(sTabName) Then
								 If True =  Fn_TabFolder_Operation("Select", sTabName, "")	Then			'''--------------Changed function By:- Pranav S[03-July-2012]								
                                    'Call Log file
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTabName+ "] : Tab is Activated ")
									Fn_PWC_TabOperation = TRUE
								else
									'Call Log file
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [" +sTabName+ "] : Tab is Not Available ")	
									Fn_PWC_TabOperation = FALSE
								End if
				Case  "Close" 
									'Counting Number of Tabs 
'									 iTabCount = objJavaTab.GetROProperty("items count")
'									For iTabIndex = 0 to iTabCount-1
'										objJavaTab.Select "#"&iTabIndex
'										sTabVal=objJavaTab.GetROProperty("value") 
'										If sTabVal = sTabName Then
'											objJavaTab.CloseTab  sTabName
'											'Call Log file
'											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTabName+ "] : Tab is Successfully Closed ")
'											Fn_PWC_TabOperation = TRUE		
'											Exit For
'										End If
'									Next     
'									If iTabIndex = iTabCount Then
'										'Call Log file
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Tab  [" +sAction+ "] is Not Available.")
'										Fn_PWC_TabOperation = FALSE
'									End If
								bFlag = False
								bFlag = Fn_TabFolder_Operation("VerifyActive", sTabName, "")
								
								If bFlag = False Then
									Fn_PWC_TabOperation = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Verify ["+sItem+"] Tab.")
								Else
									Fn_PWC_TabOperation = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Verified ["+sItem+"] Tab.")
								End If
				Case Else
								'Call Log file
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Invalid ACTION  [" +sAction+ "] is Requested.")
								Fn_PWC_TabOperation = FALSE
	End Select
	'Set objJavaTab = Nothing
End Function
'*********************************************************		Function to select  the Tab into Project		***********************************************************************
'Function Name		:				Fn_PWC_TabSet(StrTabName)

'Description			 :		 		 This function is used to select the required Tab.

'Parameters			   :	 			1.  StrTabName:Name of the Tab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Project prespective should be displayed.

'Examples				:				 Fn_PWC_TabSet("AM Rules")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					05-08-10			1.0																Harshal	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_PWC_TabSet(StrTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_TabSet"
	Dim objJavaWindow
	Set objJavaWindow = Fn_UI_ObjectCreate( "Fn_PWC_TabSet", JavaWindow("Project - Teamcenter 8"))
	   Select Case StrTabName
				'For selecting Definition Tab
			   Case "Definition" 				
						Call Fn_UI_JavaTab_Select("Fn_PWC_TabSet", objJavaWindow, "ConfigTab", "Definition")
						Fn_PWC_TabSet = TRUE				
			    'For selecting AM Rules Tab
				Case "AM Rules" 				
						Call Fn_UI_JavaTab_Select("Fn_PWC_TabSet", objJavaWindow, "ConfigTab", "AM Rules")
						Fn_PWC_TabSet = TRUE								
				'Error message If the above Tab is not selected
				Case Else 
						 Fn_PWC_TabSet = FALSE
	   End Select
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Select Tab " & StrTabName & " succeeded")
		Set objJavaWindow = Nothing 
End Function
''*********************************************************		Function to perform action on Project Tree	***********************************************************************
'Function Name		:				Fn_PWCRules_TreeOpeartion()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. GetIndex	

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 
'												  3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE / FALSE and Index Value in "GetIndex" case.

'Pre-requisite			:		 		Project Prespective is Open.

'Examples				:				Case "Select" : Call Fn_PWCRules_TreeOpeartion("Select","In Project(  ) -> Projects:Has Class( Schedule ) -> Scheduling Objects","")
'													Case "PopupMenuSelect" : Call Fn_PWCRules_TreeOpeartion("PopupMenuSelect","In Project(  ) -> Projects:Has Class( Schedule ) -> Scheduling Objects","Copy	Ctrl+C")
'													Case "PopupMenuExist" : Call Fn_PWCRules_TreeOpeartion("PopupMenuExist","In Project(  ) -> Projects:Has Class( Schedule ) -> Scheduling Objects","Copy	Ctrl+C")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Ketan Raje													05-08-10						1.0																							Harshal	
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_PWCRules_TreeOpeartion(sAction,sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_PWCRules_TreeOpeartion"
	Dim objJavaWindowProj, objJavaTreeProj, intNodeCount, intCount, sTreeItem, aMenuList
	Set objJavaWindowProj = Fn_UI_ObjectCreate( "Fn_PWCRules_TreeOpeartion",JavaWindow("Project - Teamcenter 8"))
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"					
                    Call Fn_JavaTree_Select("Fn_PWCRules_TreeOpeartion", objJavaWindowProj, "AMRuleTree",sNodeName)
					Fn_PWCRules_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_PWCRules_TreeOpeartion",objJavaWindowProj,"AMRuleTree",sNodeName)
					Fn_PWCRules_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_PWCRules_TreeOpeartion", objJavaWindowProj,"AMRuleTree",sNodeName)
					Fn_PWCRules_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_PWCRules_TreeOpeartion", objJavaWindowProj.JavaTree("AMRuleTree"))
					intNodeCount = objJavaTreeProj.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeProj.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_PWCRules_TreeOpeartion = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_PWCRules_TreeOpeartion = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuSelect"
			Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_PWCRules_TreeOpeartion", JavaWindow("Project - Teamcenter 8").JavaTree("AMRuleTree"))
					'Build the Popup menu to be selected
					aMenuList = split(sMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					'Select node
                    Call Fn_JavaTree_Select("Fn_PWCRules_TreeOpeartion",objJavaWindowProj,"AMRuleTree",sNodeName)
					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCRules_TreeOpeartion",objJavaWindowProj,"AMRuleTree",sNodeName)
					'Select Menu action
					Select Case intCount
						Case "0"
							 sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_PWCRules_TreeOpeartion = FALSE
							Exit Function
					End Select
					If JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").Exist Then
						JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").Select sMenu
						Fn_PWCRules_TreeOpeartion = TRUE
					Else
						Fn_PWCRules_TreeOpeartion = FALSE
					End If					
		'----------------------------------------------------------------------- CHECK EXISTANCE OF POP-UP MENU-------------------------------------------------------------------------
		Case "PopupMenuExist"
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCRules_TreeOpeartion",objJavaWindowProj,"AMRuleTree",sNodeName)
				If JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").GetItemProperty (sMenu,"Exists") = True Then
					Fn_PWCRules_TreeOpeartion = TRUE
				Else
					Fn_PWCRules_TreeOpeartion = FALSE
			  	End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowProj.JavaTree("AMRuleTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowProj.JavaTree("AMRuleTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_PWCRules_TreeOpeartion = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_PWCRules_TreeOpeartion = FALSE
				End If

		Case Else
						Fn_PWCRules_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PWCRules_TreeOpeartion function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_PWCRules_TreeOpeartion")
	Set objJavaWindowProj = nothing
	Set objJavaTreeProj = nothing
End Function
''*********************************************************		Function to perform action on Project Tree	***********************************************************************
'Function Name		:				Fn_PWCProject_TreeOpeartion()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. GetIndex	

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 
'												  3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE / FALSE and Index Value in "GetIndex" case.

'Pre-requisite			:		 		Project Prespective is Open.

'Examples				:				Case "Select" : Call Fn_PWCProject_TreeOpeartion("Select","Project:AutoProj_45832","")
'													Case "Expand" : Call Fn_PWCProject_TreeOpeartion("Expand","Project:AutoProj_45832","")
'													Case "Collapse" : Call Fn_PWCProject_TreeOpeartion("Collapse","Project:AutoProj_45832","")
'													Case "Exist" : Call Fn_PWCProject_TreeOpeartion("Exist","Project:AutoProj_45832","")
'													Case "GetIndex" : Call Fn_PWCProject_TreeOpeartion("GetIndex","Project:AutoProj_45832","")
'													Case "PopupMenuSelect" : Call Fn_PWCProject_TreeOpeartion("PopupMenuSelect","Project:AutoProj_45832","Copy	Ctrl+C")
'													Case "PopupMenuExist" : Call Fn_PWCProject_TreeOpeartion("PopupMenuExist","Project:AutoProj_45832","Copy	Ctrl+C")
  												
'History					 :		
'	Developer Name				Date						Rev. No.						Changes Done						Reviewer
'	------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ketan Raje					06-08-10						1.0																							Harshal	
'	------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh Watwe				16-10-15						2.0						Modified code to get path of the node
'	------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_PWCProject_TreeOpeartion(sAction,sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_PWCProject_TreeOpeartion"
	Dim objJavaWindowProj, objJavaTreeProj, intNodeCount, intCount, sTreeItem, aMenuList,aNodeName 
	Dim iPath, sInstanceHandler, arrStrNode, oCurrentNode, iCount,arrNodes
	Set objJavaWindowProj = Fn_UI_ObjectCreate( "Fn_PWCProject_TreeOpeartion",JavaWindow("Project - Teamcenter 8"))
	Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_PWCProject_TreeOpeartion", objJavaWindowProj.JavaTree("ProjectTree"))
	Fn_PWCProject_TreeOpeartion = False
	'If Project node name contains '@', use ~ as instance handler. - Added by Koustubh W ---------------
	sInstanceHandler = "@"
	If instr(sAction,"@") > 0 Then
		sInstanceHandler = "~"
	End If
	'---------------------------------------------------------------------------------------------------
	'Code Commented as "Projects" root node is changed to "Project"										------------- Pranav S[Build:0620]
  	'Added by Prasanna for change in Parent node of Project Tree. Project is replaced by 'Projects'
	aNodeName = split(sNodeName, ":",-1,1)
	If aNodeName(0) = "Projects" Then
			sNodeName=Replace(sNodeName, "Projects", "Project", 1, 1, 1)		
			aNodeName(0) = "Project"
	End If

	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Search"
'                    Call Fn_Button_Click("Fn_PWCProject_TreeOpeartion", objJavaWindowProj,"ReloadPrjct")
					'Added by Nilesh on 29-Jun-12
                    Call Fn_ToolbatButtonClick("Refresh")
					
					Wait(5)
					If JavaDialog("Search").Exist(3) Then
						JavaDialog("Search").JavaButton("Yes").Click
					End If
					If sNodeName <> "" Then
						Call Fn_Edit_Box("Fn_PWCProject_TreeOpeartion",objJavaWindowProj,"PrjctSrch", sNodeName)
					End If
'                    Call Fn_Button_Click("Fn_PWCProject_TreeOpeartion", objJavaWindowProj,"PrjctFind")
'				Added by Nilesh on 29-Ju n-12
					Wait (2)
					Call Fn_ToolbatButtonClick("Enter a name to do Find in Display")
                    Fn_PWCProject_TreeOpeartion = TRUE
        '----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select", "Select_@"
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, aNodeName(0) , "", sInstanceHandler)
					'Call Fn_PWCProject_TreeOpeartion("Expand",aNodeName(0),"")
					If iPath <> False Then
						objJavaTreeProj.Expand iPath
					else
						Exit function 
					End If
					Call Fn_ReadyStatusSync(1)
					'Call Fn_JavaTree_Select("Fn_PWCProject_TreeOpeartion", objJavaWindowProj, "ProjectTree",sNodeName)
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, sNodeName , "", sInstanceHandler)
					'Call Fn_PWCProject_TreeOpeartion("Expand",aNodeName(0),"")
					If iPath <> False Then
						objJavaTreeProj.Select iPath
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit function 
					End If
					
					Fn_PWCProject_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand", "Expand_@"
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, aNodeName(0) , "", sInstanceHandler)
					'Call Fn_PWCProject_TreeOpeartion("Expand",aNodeName(0),"")
					If iPath <> False Then
						objJavaTreeProj.Expand iPath
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + aNodeName(0) + "] of JavaTree does not exist")
						Exit function 
					End If
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, sNodeName , "", sInstanceHandler)
					'Call Fn_PWCProject_TreeOpeartion("Expand",aNodeName(0),"")
					If iPath <> False Then
						objJavaTreeProj.Expand iPath
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit function 
					End If
					Fn_PWCProject_TreeOpeartion = True
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse", "Collapse_@"
                    iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, sNodeName , "", sInstanceHandler)
					'Call Fn_PWCProject_TreeOpeartion("Expand",aNodeName(0),"")
					If iPath <> False Then
						objJavaTreeProj.Collapse iPath
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit function 
					End If
					Fn_PWCProject_TreeOpeartion = True
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist", "Exist_@"
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, aNodeName(0) , "", sInstanceHandler)
					'Call Fn_PWCProject_TreeOpeartion("Expand",aNodeName(0),"")
					If iPath <> False Then
						objJavaTreeProj.Expand iPath
					else
						Exit function 
					End If
					Call Fn_ReadyStatusSync(1)
					'Call Fn_JavaTree_Select("Fn_PWCProject_TreeOpeartion", objJavaWindowProj, "ProjectTree",sNodeName)
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, sNodeName , "", sInstanceHandler)
					'Call Fn_PWCProject_TreeOpeartion("Expand",aNodeName(0),"")
					If iPath <> False Then
						Fn_PWCProject_TreeOpeartion = TRUE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
					else
						Fn_PWCProject_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
					End If
		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuSelect", "PopupMenuSelect_@"
			Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_PWCProject_TreeOpeartion", objJavaTreeProj)
					'Build the Popup menu to be selected
					aMenuList = split(sMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					'Select node
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, sNodeName , "", sInstanceHandler)
					If iPath = False Then
						Fn_PWCProject_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit function
					End If
					objJavaTreeProj.Select iPath
                    'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCProject_TreeOpeartion",objJavaWindowProj,"ProjectTree",iPath)
					'Select Menu action
					Select Case intCount
						Case "0"
							 sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_PWCProject_TreeOpeartion = FALSE
							Exit Function
					End Select
					If JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").Exist Then
						JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").Select sMenu
						Fn_PWCProject_TreeOpeartion = TRUE
					Else
						Fn_PWCProject_TreeOpeartion = FALSE
					End If					
		'----------------------------------------------------------------------- CHECK EXISTANCE OF POP-UP MENU-------------------------------------------------------------------------
		Case "PopupMenuExist"
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCProject_TreeOpeartion",objJavaWindowProj,"ProjectTree",sNodeName)
				If JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").GetItemProperty (sMenu,"Exists") = True Then
					Fn_PWCProject_TreeOpeartion = TRUE
				Else
					Fn_PWCProject_TreeOpeartion = FALSE
			  	End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex","GetIndex_Ext"								'[TC1122-2016010600-15_01_2016-VivekA-NewDevelopment] - Added new case "GetIndex_Ext" to get tree index
				bFlag = False
				If sAction = "GetIndex_Ext" Then
					bReturn = Fn_UI_getJavaTreeIndex(objJavaWindowProj.JavaTree("ProjectTree"),sNodeName)
					Fn_PWCProject_TreeOpeartion = bReturn
					bFlag = True
				Else		
					For intCount=0 to objJavaWindowProj.JavaTree("ProjectTree").GetROProperty ("items count")-1
						sTreeItem = objJavaWindowProj.JavaTree("ProjectTree").GetItem (intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_PWCProject_TreeOpeartion = intCount
							bFlag = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
							Exit For
						End If
					Next
					If bFlag = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
						Fn_PWCProject_TreeOpeartion = FALSE
					End If
				End If

		Case "ClickReload"
                    Call Fn_Button_Click("Fn_PWCProject_TreeOpeartion", objJavaWindowProj,"ReloadPrjct")
					Fn_PWCProject_TreeOpeartion = TRUE
		'[TC1015-2015101300-02_11_2015-VivekA-NewDevelopment] - Added by Snehal S
		Case "GetChildItemCount"
				If Fn_PWCProject_TreeOpeartion("Expand",sNodeName,"")=True Then
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", JavaWindow("Project - Teamcenter 8").JavaTree("ProjectTree"),sNodeName, "", sInstanceHandler)
					If iPath <> False Then
						iPath = replace(iPath, "#", "") 
						arrStrNode = Split(iPath,":",-1,1)
						Set oCurrentNode = JavaWindow("Project - Teamcenter 8").JavaTree("ProjectTree").Object.getItem(cInt(arrStrNode(0)))
						For iCount = 1 to uBound(arrStrNode)
							Set oCurrentNode = oCurrentNode.getItem(cInt(arrStrNode(iCount)))
						Next
						Fn_PWCProject_TreeOpeartion = cInt(oCurrentNode.getItemCount())
						Set oCurrentNode=Nothing
					Else
						Fn_PWCProject_TreeOpeartion = False
					End If
				Else
					Fn_PWCProject_TreeOpeartion = False
				End If
		'----------------------------------------------------------------------- For selecting Multiple nodes -------------------------------------------------------------------------
		'[TC11.3(20170509d)_NewDevelopment_PoonamC_25July2017: Added Case to multi select Nodes from project tree.]
		Case "MultiSelect"
					arrNodes = Split(sNodeName,"~")
					For iCount = 0 To UBound(arrNodes)
							iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, arrNodes(iCount) , "", sInstanceHandler)
							If iPath <> False Then
								objJavaTreeProj.ExtendSelect iPath
								Call Fn_ReadyStatusSync(1)
							else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + arrNodes(iCount) + "] of JavaTree does not exist")
								Exit function 
							End If	
					Next
					Fn_PWCProject_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For deselecting nodes -------------------------------------------------------------------------
		Case "Deselect"	
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCProject_TreeOpeartion", objJavaTreeProj, sNodeName , "", sInstanceHandler)
					If iPath=False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to DeSelect Node [" + sNodeName + "] of NavTree")
						  Fn_PWCProject_TreeOpeartion = False
					Else
						objJavaTreeProj.Deselect iPath
						Call Fn_ReadyStatusSync(SISW_MICRO_TIMEOUT)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DeSelected Node [" + sNodeName + "] of NavTree")
						Fn_PWCProject_TreeOpeartion = True
					End If
		Case Else
						Fn_PWCProject_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PWCProject_TreeOpeartion function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_PWCProject_TreeOpeartion")
	Set objJavaWindowProj = nothing
	Set objJavaTreeProj = nothing
End Function
''*********************************************************		Function to perform action on Project Tree	***********************************************************************
'Function Name		:				Fn_PWCOrgGrp_TreeOpeartion()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. GetIndex	

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 
'												  3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE / FALSE and Index Value in "GetIndex" case.

'Pre-requisite			:		 		Project Prespective is Open.

'Examples				:				Case "Select" : Call Fn_PWCOrgGrp_TreeOpeartion("Select","Project:AutoProj_45832","")
'													Case "Expand" : Call Fn_PWCOrgGrp_TreeOpeartion("Expand","Project:AutoProj_45832","")
'													Case "Collapse" : Call Fn_PWCOrgGrp_TreeOpeartion("Collapse","Project:AutoProj_45832","")
'													Case "Exist" : Call Fn_PWCOrgGrp_TreeOpeartion("Exist","Project:AutoProj_45832","")
'													Case "GetIndex" : Call Fn_PWCOrgGrp_TreeOpeartion("GetIndex","Project:AutoProj_45832","")
'													Case "PopupMenuSelect" : Call Fn_PWCOrgGrp_TreeOpeartion("PopupMenuSelect","Project:AutoProj_45832","Copy	Ctrl+C")
'													Case "PopupMenuExist" : Call Fn_PWCOrgGrp_TreeOpeartion("PopupMenuExist","Project:AutoProj_45832","Copy	Ctrl+C")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Ketan Raje													06-08-10						1.0																							Harshal	
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_PWCOrgGrp_TreeOpeartion(sAction,sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_PWCOrgGrp_TreeOpeartion"
	Dim objJavaWindowProj, objJavaTreeProj, intNodeCount, intCount, sTreeItem, aMenuList
	Dim sRootNode,aNodes,bReturn
	ReDim aNodes(2)
	Set objJavaWindowProj = Fn_UI_ObjectCreate( "Fn_PWCOrgGrp_TreeOpeartion",JavaWindow("Project - Teamcenter 8"))

	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"					
                    Call Fn_JavaTree_Select("Fn_PWCOrgGrp_TreeOpeartion", objJavaWindowProj, "OrgGrpTree",sNodeName)
					Fn_PWCOrgGrp_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_PWCOrgGrp_TreeOpeartion",objJavaWindowProj,"OrgGrpTree",sNodeName)
					Fn_PWCOrgGrp_TreeOpeartion = TRUE
		Case "LogicalExpand"
                    Call Fn_UI_JavaTree_Expand("Fn_PWCOrgGrp_TreeOpeartion",objJavaWindowProj,"OrgGrpTree",sNodeName)
					sRootNode = JavaWindow("Project - Teamcenter 8").JavaTree("SelMemTree").GetItem(0)
					If instr(1,sNodeName,":")<>0 Then
						aNodes = Split(sNodeName,":")
						bReturn = Fn_PWCSelMem_TreeOpeartion("Exist",sRootNode+":"+aNodes(1),"")
						If bReturn = True Then
							JavaWindow("Project - Teamcenter 8").JavaTree("SelMemTree").Select sRootNode+":"+aNodes(1)
							JavaWindow("Project - Teamcenter 8").JavaButton("Remove").Click
						End If
					End If
					Fn_PWCOrgGrp_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_PWCOrgGrp_TreeOpeartion", objJavaWindowProj,"OrgGrpTree",sNodeName)
					Fn_PWCOrgGrp_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_PWCOrgGrp_TreeOpeartion", objJavaWindowProj.JavaTree("OrgGrpTree"))
					intNodeCount = objJavaTreeProj.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeProj.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_PWCOrgGrp_TreeOpeartion = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_PWCOrgGrp_TreeOpeartion = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuSelect"
			Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_PWCOrgGrp_TreeOpeartion", JavaWindow("Project - Teamcenter 8").JavaTree("OrgGrpTree"))
					'Build the Popup menu to be selected
					aMenuList = split(sMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					'Select node
                    Call Fn_JavaTree_Select("Fn_PWCOrgGrp_TreeOpeartion",objJavaWindowProj,"OrgGrpTree",sNodeName)
					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCOrgGrp_TreeOpeartion",objJavaWindowProj,"OrgGrpTree",sNodeName)
					'Select Menu action
					Select Case intCount
						Case "0"
							 sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							sMenu = JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_PWCOrgGrp_TreeOpeartion = FALSE
							Exit Function
					End Select
					If JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").Exist Then
						JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").Select sMenu
						Fn_PWCOrgGrp_TreeOpeartion = TRUE
					Else
						Fn_PWCOrgGrp_TreeOpeartion = FALSE
					End If					
		'----------------------------------------------------------------------- CHECK EXISTANCE OF POP-UP MENU-------------------------------------------------------------------------
		Case "PopupMenuExist"
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCOrgGrp_TreeOpeartion",objJavaWindowProj,"OrgGrpTree",sNodeName)
				If JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").GetItemProperty (sMenu,"Exists") = True Then
					Fn_PWCOrgGrp_TreeOpeartion = TRUE
				Else
					Fn_PWCOrgGrp_TreeOpeartion = FALSE
			  	End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowProj.JavaTree("OrgGrpTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowProj.JavaTree("OrgGrpTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_PWCOrgGrp_TreeOpeartion = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_PWCOrgGrp_TreeOpeartion = FALSE
				End If

		Case Else
						Fn_PWCOrgGrp_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PWCOrgGrp_TreeOpeartion function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_PWCOrgGrp_TreeOpeartion")
	Set objJavaWindowProj = nothing
	Set objJavaTreeProj = nothing
End Function
''*********************************************************		Function to perform action on Project Tree	***********************************************************************
'Function Name		:				Fn_PWCSelMem_TreeOpeartion()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. GetIndex	

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 
'												  3. StrMenu: Context menu to be selected

'Return Value		   : 				TRUE / FALSE and Index Value in "GetIndex" case.

'Pre-requisite			:		 		Project Prespective is Open.

'Examples				:				Case "Select" : Call Fn_PWCSelMem_TreeOpeartion("Select","AutoProj_23496:dba","")
'													Case "Expand" : Call Fn_PWCSelMem_TreeOpeartion("Expand","AutoProj_23496:dba","")
'													Case "Collapse" : Call Fn_PWCSelMem_TreeOpeartion("Collapse","AutoProj_23496:dba","")
'													Case "DoubleClick" : Call Fn_PWCSelMem_TreeOpeartion("DoubleClick","AutoProj_23496:dba:DBA:dba/DBA/AutoTest7","")
'													Case "Exist" : Call Fn_PWCSelMem_TreeOpeartion("Exist","AutoProj_23496:dba","")
'													Case "GetIndex" : Call Fn_PWCSelMem_TreeOpeartion("GetIndex","AutoProj_23496:dba","")
'													Case "PopupMenuSelect" : Call Fn_PWCSelMem_TreeOpeartion("PopupMenuSelect","AutoProj_23496:dba","Set Privileged Users")
'													Case "PopupMenuExist" : Call Fn_PWCSelMem_TreeOpeartion("PopupMenuExist","AutoProj_23496:dba","Set Privileged Users")
'													Case "GetStatus" : Call Fn_PWCSelMem_TreeOpeartion("GetStatus","AutoProj_23496:dba","")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Ketan Raje													06-08-10						1.0																							Harshal	
'														Sandeep N														10-07-12						1.1				modified case : DoubleClick			    Pranav S
'																															modified case DoubleClick to expand all parent node before avtivating Child node
'														Sandeep N														17-07-12						1.2				added case : GetStatus			    		Pranav S
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_PWCSelMem_TreeOpeartion(sAction,sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_PWCSelMem_TreeOpeartion"
	Dim objJavaWindowProj, objJavaTreeProj, intNodeCount, intCount, sTreeItem, aMenuList
	Dim aNodeList, aPersonName, iIndex, sText,iCounter
	Dim sExpand,aNodeName,aMemName
	Dim sPath,aPath,sNode,iCnt,sValue
	Dim var,childObjects,Objbutton
	Set objJavaWindowProj = Fn_UI_ObjectCreate( "Fn_PWCSelMem_TreeOpeartion",JavaWindow("Project - Teamcenter 8"))


	'Temporary solution for PR#6528447 - Considering autotest users are having same person & login names
	'Vallari - 01Jun11 - Commenting this as test  scripts ahve been updated for the same
	'aNodeList = split(sNodeName, ":", -1, 1)
	'If UBound(aNodeList) > 1 Then
		'aPersonName = split(sNodeName, "/", -1, 1)
		'iIndex = UBound(aPersonName)
		'sText = Lcase(aPersonName(iIndex))
		'sNodeName = sNodeName + " (" + sText + ")"
	'End If

	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"					
					'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
					'Added line by : Sandeep : 23-May-2013
					objJavaWindowProj.JavaTree("SelMemTree").Select sNodeName
					wait 1
					'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
                    Call Fn_JavaTree_Select("Fn_PWCSelMem_TreeOpeartion", objJavaWindowProj, "SelMemTree",sNodeName)
					Fn_PWCSelMem_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sNodeName)
                    Fn_PWCSelMem_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_PWCSelMem_TreeOpeartion", objJavaWindowProj,"SelMemTree",sNodeName)
					Fn_PWCSelMem_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
		Case "DoubleClick"		
					'' Swapnil: This code is added to expaned the member selection tree and select the appropriate member to make him previliges user.
					
'					aNodeName = split(sNodeName,":",-1,1)   ' Added by Prasanna 18-06-12 for Tc10.0(0606)
'					aMemName = Split(aNodeName(0),".")			'Changed for hierarchy by Pranav S[27-Jun-2012] TC 10.0 (0620)
'					sExpand = aMemName(0)+"."+aMemName(1)
'					'Expand the OrgGrp Tree
'					Call Fn_UI_JavaTree_Expand("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sExpand)
'					wait(3)
'					sNodeName = sExpand + ":"+aNodeName(1)
'					Call Fn_JavaTree_Node_Activate("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sNodeName)
'					Fn_PWCSelMem_TreeOpeartion = TRUE
					aNodeName = split(sNodeName,":",-1,1)
					For iCounter=0 to ubound(aNodeName)-1
						If iCounter=0 Then
							sExpand=aNodeName(0)
						else
							sExpand=sExpand+":"+aNodeName(iCounter)
						End If
						wait 1					
						Call Fn_UI_JavaTree_Expand("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sExpand)
					Next
					'wait 2
					Call Fn_JavaTree_Node_Activate("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sNodeName)
					Fn_PWCSelMem_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_PWCSelMem_TreeOpeartion", objJavaWindowProj.JavaTree("SelMemTree"))
					intNodeCount = objJavaTreeProj.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeProj.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_PWCSelMem_TreeOpeartion = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_PWCSelMem_TreeOpeartion = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuSelect"
			Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_PWCSelMem_TreeOpeartion", JavaWindow("Project - Teamcenter 8").JavaTree("SelMemTree"))
					'Build the Popup menu to be selected
					aMenuList = split(sMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					'Select node
                    Call Fn_JavaTree_Select("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sNodeName)
					'Open context menu
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sNodeName)
					'Select Menu action
                    If JavaWindow("Project - Teamcenter 8").JavaMenu("label:="&sMenu).Exist Then
						JavaWindow("Project - Teamcenter 8").JavaMenu("label:="&sMenu).Select
						Fn_PWCSelMem_TreeOpeartion = TRUE
					Else
						Fn_PWCSelMem_TreeOpeartion = FALSE
					End If					
		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
		Case "PopupMenuIsEnabled"
			Dim bolState
			ArrMenu=Split(sMenu,":")
			NumObjects = ubound(ArrMenu) 
			Call Fn_JavaTree_Select("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sNodeName)
			Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sNodeName)
			bolState = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaMenu("label:="&ArrMenu(0)&"","index:=0").GetROProperty("enabled")				
						If  Cbool(bolState) Then
							Fn_PWCSelMem_TreeOpeartion = True
					    Else 
							Fn_PWCSelMem_TreeOpeartion = False
						End If
            		'----------------------------------------------------------------------- CHECK EXISTANCE OF POP-UP MENU-------------------------------------------------------------------------
		Case "PopupMenuExist"
				Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWCSelMem_TreeOpeartion",objJavaWindowProj,"SelMemTree",sNodeName)
				If JavaWindow("Project - Teamcenter 8").WinMenu("ContextMenu").GetItemProperty (sMenu,"Exists") = True Then
					Fn_PWCSelMem_TreeOpeartion = TRUE
				Else
					Fn_PWCSelMem_TreeOpeartion = FALSE
			  	End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowProj.JavaTree("SelMemTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowProj.JavaTree("SelMemTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_PWCSelMem_TreeOpeartion = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_PWCSelMem_TreeOpeartion = FALSE
				End If
		'----------------------------------------------------------------------- Case to get current status of User-------------------------------------------------------------------------
		Case "GetStatus"
				If sNodeName<>"" Then
					Fn_PWCSelMem_TreeOpeartion=objJavaWindowProj.JavaTree("SelMemTree").GetColumnValue(sNodeName,"Status")
				else
					Fn_PWCSelMem_TreeOpeartion=false
				End If
				'      ------------------------------------------------------------------------Case to Multi-Select Nodes---------------------------------------------------------------------------------
	
	Case "Multiselect"
		Dim NodeLists
'			Split the string where "'," exist
			NodeLists = Split(sNodeName,",")
			intNodeCount =ubound(NodeLists)
			
'
				For intCount =0 to intNodeCount				
							objJavaWindowProj.JavaTree("SelMemTree").ExtendSelect NodeLists(intCount)
							Fn_PWCSelMem_TreeOpeartion = TRUE
				Next

'----------------------------------------------------------------------- For doble clicking on a particular  node-------------------------------------------------------------------------
	Case "MultiSelectContextMenu"
					NodeLists = split(sNodeName,",",-1,1)
					aMenuList = split(sMenu, ":",-1,1)
					intCount = Ubound(aMenuList)


					'Select multiple node
					Call Fn_PWCSelMem_TreeOpeartion("Multiselect", sNodeName, "")

					'Open context menu
                	Call Fn_UI_JavaTree_OpenContextMenu("Fn_MyTc_NavTree_NodeOperation",objJavaWindowProj,"SelMemTree",NodeLists(0))
					Select Case intCount
						Case "0"
							 
							 StrMenu = objJavaWindowProj.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu =objJavaWindowProj.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = objJavaWindowProj.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_PWCSelMem_TreeOpeartion = FALSE
							Exit Function
					End Select
					objJavaWindowProj.WinMenu("ContextMenu").Select sMenu
					If err.number < 0 Then
						Fn_PWCSelMem_TreeOpeartion = false
					else
						Fn_PWCSelMem_TreeOpeartion = TRUE
					End If
		Case "ListSelect"					
					objJavaWindowProj.JavaList("LibraryList").Select sNodeName
					wait 1		
					Fn_PWCSelMem_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- CHECK FOR EXPANDED ITEM-------------------------------------------------------------------------	  	
		Case "IsExpanded"
				sPath = Fn_UI_JavaTreeGetItemPathExt("Fn_PWCSelMem_TreeOpeartion", JavaWindow("Project - Teamcenter 8").JavaTree("SelMemTree"),sNodeName, "", "")
				aPath = Split(Replace(sPath,"#",""),":")
				For iCnt = 0 To Ubound(aPath)
					If iCnt = 0 Then
						Set sNode = JavaWindow("Project - Teamcenter 8").JavaTree("SelMemTree").Object.getItem(aPath(iCnt))
					Else
						Set sNode  = sNode.getItem(aPath(iCnt)) 
					End If
				Next
				sValue = sNode.getExpanded()
				If CBool(sValue) = True  Then
					Fn_PWCSelMem_TreeOpeartion = TRUE
				Else
					Fn_PWCSelMem_TreeOpeartion = FALSE
				End If
			'----------------------------------------------------------------------- CHECK FOR sELECTED ITEM-------------------------------------------------------------------------	  			
		Case "IsSelected"
			 set oCurrentNode = JavaWindow("Project - Teamcenter 8").JavaTree("SelMemTree").Object.getFocusItem()
			 strNode=oCurrentNode.getNameText()
			 If strNode=sNodeName Then
			 	Fn_PWCSelMem_TreeOpeartion=True
			 Else
				Fn_PWCSelMem_TreeOpeartion=False			 
			 End If
		'----------------------------------------------------------------------- Search for ITEM-------------------------------------------------------------------------	  				 
		'[TC11.3_NewDevelopment_PoonamC_01Aug2017: Added New Cases to serch User,Group & Role from Selected Menmber tree]
		Case "SearchGroup","SearchUser","SearchRole"
				' Click on Clear button
				 Call Fn_Button_Click("Fn_PWCSelMem_TreeOpeartion", objJavaWindowProj, "ClearSelMem")
				 Call Fn_ReadyStatusSync(1)
				 
				 'Set Serach Criteria
			 	 Call Fn_SISW_UI_JavaEdit_Operations("Fn_PWCSelMem_TreeOpeartion","Type",objJavaWindowProj,"SelectedMemSrch",sNodeName)
			 	 Call Fn_ReadyStatusSync(1)
			 	 
			 	 '================== [ TC11.6_20180814b00_Maintenance_PoonamC_21Aug2018 : Added code to click on button ] =================================================
			 	 If sAction = "SearchGroup" Then
			 	 		'Fn_PWCSelMem_TreeOpeartion = Fn_Button_Click("Fn_PWCSelMem_TreeOpeartion", objJavaWindowProj, "GroupSelMem")
						sValue = "Find groups"									 	
			 	ElseIf sAction = "SearchUser" Then
			 		   'Fn_PWCSelMem_TreeOpeartion = Fn_Button_Click("Fn_PWCSelMem_TreeOpeartion", objJavaWindowProj, "UserSelMem")
						sValue = "Find users (Based on user ID and user name)"
				ElseIf sAction = "SearchRole" Then
				 	  ' Fn_PWCSelMem_TreeOpeartion = Fn_Button_Click("Fn_PWCSelMem_TreeOpeartion", objJavaWindowProj, "RoleSelMem")
			 	 		sValue = "Find roles"	
			 	End If
			 	
			 	Set var = Description.Create()
					var("Class Name").value = "JavaButton"
					var("toolkit class").value = "org.eclipse.swt.widgets.Button"
				 Set childObjects = objJavaWindowProj.ChildObjects(var)
			 	
	 	 		intCount = 1
				For iCounter = 0 To childObjects.count - 1  
					If childObjects(iCounter).Object.getToolTipText = sValue Then
						If intCount = 2 Then
							Set Objbutton = childObjects(iCounter)
							Exit For
						Else
							intCount = intCount + 1
						End If
					End If
				Next
			 	Objbutton.Click
			 	Call Fn_ReadyStatusSync(1)
			 	
			 	 If Err.Number < 0 Then
			 	 	Fn_PWCSelMem_TreeOpeartion = False
			 	 Else
			 	 	Fn_PWCSelMem_TreeOpeartion = True
			 	 End If
			 	 
				 Set childObjects = Nothing
				 Set var = Nothing
				 Set Objbutton = Nothing
		'==========================================================================================	
		Case Else
						Fn_PWCSelMem_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PWCSelMem_TreeOpeartion function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_PWCSelMem_TreeOpeartion")
	Set objJavaWindowProj = nothing
	Set objJavaTreeProj = nothing
End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Fn_PWC_MemberSelection()  
'###
'###    DESCRIPTION        :   Add/Remove Members
'###	Prequisite 					:	1.Project Prespective is Open.
'###
'###    PARAMETERS      :  1.sAction:Add/Remove
'###  											2.sNodeName:
'###                                           3.sExtra:
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_PWCOrgGrp_TreeOpeartion(), Fn_PWCSelMem_TreeOpeartion(), Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION			Build
'###
'###    CREATED BY     :   Ketan Raje           			06/08/2010       	1.0
'###
'###    REVIWED BY     :    Harshal
'###
'###    MODIFIED BY   :  	Harshal							07/09/2011			1.1						Tc91(2011082400)
'###
'###    EXAMPLE          : 		Case "Add" : Call Fn_PWC_MemberSelection("Add", "Organization:Project Administration:Project Administrator:AutoProjAdmin (autoprojadmin)", "")	[Traditional Method]
'###																																														OR
'###																	Call Fn_PWC_MemberSelection("Add",Environment.Value("TcUser1"),"") [Recommended]
'###
'###										 Case "Remove" : Call Fn_PWC_MemberSelection("Remove", "Team:Project Administration", "")
'###										Case "	DoubleClick": Call Fn_PWC_MemberSelection("DoubleClick", "Organization:Engineering", "")	'To add group by double clicking
'###										Case "MultiSelect" : Call Fn_PWC_MemberSelection("MultiSelect", "Organization:Engineering,Organization:Manufacturing", "")	'	To select more than one group
'#############################################################################################################
Public Function Fn_PWC_MemberSelection(sAction, sNodeName, sExtra)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_MemberSelection"
	Dim objMember,aNodeName,iCount,sExpand
	Dim sNode,aMember,aOrgGrpNode
	Dim var,childObjects,Objbutton,sValue
	Set objMember = Fn_UI_ObjectCreate("Fn_PWC_MemberSelection", JavaWindow("Project - Teamcenter 8"))
	'Commented as it is not required	
	'Relaod organization tree
'	Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Refresh")
'	wait(3)
		Select Case sAction
        			Case "Add"
        			        Call Fn_ProjectTabOperations("Activate","Member Selection")
							Wait(1)
							aNodeName = split(sNodeName,":",-1,1)
							If ubound(aNodeName)>1 Then
								For iCount=0 to ubound(aNodeName)-1
									If iCount  = 0 Then
										sExpand = aNodeName(iCount)
									Else
										sExpand = sExpand+":"+aNodeName(iCount)
									End If
								'Expand the OrgGrp Tree
								Call Fn_PWCOrgGrp_TreeOpeartion("Expand",sExpand,"")
								Next
							End If
							wait(1)
							'Select the OrgGrp Tree
							Call Fn_PWCOrgGrp_TreeOpeartion("Select",sNodeName,"")
							For iCount=0 to 2
								Call Fn_ReadyStatusSync(1)
								objMember.JavaButton("Add").WaitProperty "enabled", "1"
								If objMember.JavaButton("Add").GetROProperty("enabled")=1 Then
									'Click on Add button
									Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Add")
									Exit for
								End If
								Wait(1)
							Next
							'Handle Null Pointer exception									
							Call Fn_PWC_DialogMsgVerify("Error","","OK") 
				Case "Remove"
							aNodeName = split(sNodeName,":",-1,1)
							If ubound(aNodeName)>1 Then
								For iCount=0 to ubound(aNodeName)-1
									If iCount  = 0 Then
										sExpand = aNodeName(iCount)
									Else
										sExpand = sExpand+":"+aNodeName(iCount)
									End If
								'Expand the OrgGrp Tree
								Call Fn_PWCSelMem_TreeOpeartion("Expand",sExpand,"")
								Next
							End If
							'Select the SelMem Tree
								Call Fn_PWCSelMem_TreeOpeartion("Expand",aNodeName(0),"")
								wait(2)
							Call Fn_PWCSelMem_TreeOpeartion("Select",sNodeName,"")
							wait(2)
							'Click on Add button
							bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_PWC_MemberSelection", "Click",  objMember, "Remove")
							   If bGblFuncRetVal = False Then
							   	 Fn_PWC_MemberSelection = False
							   	 Exit Function
							   End If				
							Wait(3)
							'Handle Null Pointer exception									
							Call Fn_PWC_DialogMsgVerify("Error","","OK") 
							If objMember.JavaWindow("Remove User").Exist(2) Then
								Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember.JavaWindow("Remove User"), "Yes")
								Call Fn_ReadyStatusSync(1)
							End If

				Case "AddGroup","AddUser","AddRole"
							ReDim aOrgGrpNode(6)
                            Set objMember = Nothing
							Set objMember = JavaWindow("Project - Teamcenter 8")
							Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Refresh")
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Refreshed Organisation Tree")
							If sAction = "AddGroup" Then
								aNodeName = Split(sNodeName,":")
								sNode = aNodeName(Ubound(aNodeName))
							Elseif sAction = "AddUser" Then
							'Commented by Prasanna 15-June-12 for 10.0(0606 build)
'									If instr(1,sNodeName,"Organization" ) = 0Then 'Harshal 07Sept2011 : 4Line Code Added for Envoirnment Variable Compatibility.
'										aOrgGrpNode = split(sNodeName,":",-1,1)
'										sNodeName = "Organization:"+aOrgGrpNode(2)+":"+aOrgGrpNode(3)+":"+aOrgGrpNode(0)+" ("+aOrgGrpNode(5)+")" 
'									End If
									aNodeName = Split(sNodeName,"(")
									aMember = Split(aNodeName(1),")")
									sNode = aMember(0)
						ElseIf sAction="AddRole" Then	
							aNodeName = Split(sNodeName,":")
							sNode=aNodeName(Ubound(aNodeName))
                             End If

							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Setting KeyWord to search")
							JavaWindow("Project - Teamcenter 8").JavaEdit("OrgGrpSrch").Set sNode
							
							'================== [ TC11.6_20180814b00_Maintenance_PoonamC_21Aug2018 : Added code to click on button ] =================================================
							If sAction = "AddGroup" Then
								'Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Group")
								 sValue = "Find groups"
							Elseif sAction = "AddUser" Then
								'Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "User")
								 sValue = "Find users (Based on user ID and user name)"
							Elseif sAction = "AddRole" Then
								'Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Role")
								sValue = "Find roles"
							End If
							
							Set var = Description.Create()
								var("Class Name").value = "JavaButton"
								var("toolkit class").value = "org.eclipse.swt.widgets.Button"
							Set childObjects = JavaWindow("Project - Teamcenter 8").ChildObjects(var)
							
							For iCount = 0 To childObjects.count - 1  
								If childObjects(iCount).Object.getToolTipText = sValue Then
									Set Objbutton = childObjects(iCount)
									Exit For
								End If
							Next
							Objbutton.Click
							Call Fn_ReadyStatusSync(1)
							
							Set var = Nothing
							Set childObjects = Nothing
							Set Objbutton = Nothing
							'===========================================================================================
							If JavaWindow("Project - Teamcenter 8").JavaWindow("Search").Exist(5)  Then
								Call Fn_Button_Click("Fn_PWC_MemberSelection", JavaWindow("Project - Teamcenter 8").JavaWindow("Search"), "OK")
								Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Refresh")
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Object not found in the Organisation Tree")
								Fn_PWC_MemberSelection = False
								Exit Function
							End If
							Wait(6)
							If sAction = "AddGroup" Then
								Call Fn_ProjectTabOperations("Activate","Member Selection")
								Wait(1)
							End If
							JavaWindow("Project - Teamcenter 8").JavaTree("OrgGrpTree").Select sNodeName
							Wait(6)
							For iCount=0 to 2
								Call Fn_ReadyStatusSync(1)
								objMember.JavaButton("Add").WaitProperty "enabled", "1"
								If objMember.JavaButton("Add").GetROProperty("enabled")=1 Then
									'Click on Add button
									Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Add")
									Wait(3)
									Exit for
								End If
								Wait(3)
							Next
							'Handle Null Pointer exception									
							'Call Fn_PWC_DialogMsgVerify("Error","","OK")
							Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Refresh")
							If objMember.JavaButton("Clear_Search").Exist(1) Then
								If objMember.JavaButton("Clear_Search").GetROProperty("enabled") = 1 Then
									Call Fn_Button_Click("Fn_PWC_MemberSelection", objMember, "Clear_Search")
								End If
							End If
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),sAction+" Sucessfull")
							Fn_PWC_MemberSelection = True

				Case "DoubleClick"
							Call Fn_JavaTree_Node_Activate("Fn_PWC_MemberSelection",objMember,"OrgGrpTree",sNodeName)
							Fn_PWC_MemberSelection = True

				Case "MultiSelect"
							Call Fn_UI_JavaTree_ExtendSelect("Fn_PWC_MemberSelection",objMember,"OrgGrpTree", sNodeName)
							Fn_PWC_MemberSelection = True

				Case Else						
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PWC_MemberSelection function failed")
							Fn_PWC_MemberSelection = FALSE
							Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PWC_MemberSelection")
	Fn_PWC_MemberSelection = TRUE
	Set objMember = nothing
End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Fn_PWC_ProjectOperations()  
'###
'###    DESCRIPTION        :   Add/Remove Members
'###	Prequisite 					:	1.Project Prespective is Open.
'###
'###    PARAMETERS      :  1.sAction:
'###  											2.sID:
'###                                           3.sName:
'###                                           4.sDescription:
'###                                           5.sStatus:
'###                                           6.sSecurity:
'###                                           7.sMemberAction:
'###                                           8.sNodeName:
'###                                           9.sButton:
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), Fn_Edit_Box(), Fn_UI_ObjectCreate(), Fn_UI_JavaRadioButton_SetON(), Fn_UI_Object_SetTOProperty(), Fn_CheckBox_Set(), Fn_PWC_MemberSelection(), Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :   Ketan Raje           			10/08/2010       	1.0
'###
'###    REVIWED BY     :    Harshal
'###
'###    MODIFIED BY   		:  Shweta Rathod		8-Aug-17		1.0        	Added case to verify status of Program/project type [ TC11.3_20170509d_NewDev_PWC ]  				
'###
'###    EXAMPLE          : 		Case "Create" : Call Fn_PWC_ProjectOperations("Create", "KetanTest", "KetanTest", "Testing1", "Active", "ON", "Add", "Organization:Engineering")
'###										 Case "Modify" : Call Fn_PWC_ProjectOperations("Modify", "KetanTest2", "KetanTest2", "Testing2", "Active", "OFF", "", "")
'###										 Case "Delete" : Call Fn_PWC_ProjectOperations("Delete", "", "", "", "", "", "", "")
'###										 Case "Verify" : Call Fn_PWC_ProjectOperations("Verify", "AutoProj_982010191427", "AutoProj_982010191427", "Testing", "", "", "", "")	
'###										 Case "Exist" : Call Fn_PWC_ProjectOperations("Exist", "", "", "", "Active:Inactive:Inactive And Invisible", "sSecurity", "", "")
'###										 case "ExistWithStatus":Call Fn_PWC_ProjectOperations("ExistWithStatus","","", "", "Program", "", "","")
'#############################################################################################################
Public Function Fn_PWC_ProjectOperations(sAction, sID, sName, sDescription, sStatus, sSecurity, sMemberAction, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_ProjectOperations"
	Dim objProject, aStatus, bFlag,aNodeName, aSecurity

	If instr(1,sNodeName,"Organization",1)  Then              ' Added by prasanna for 10.0 - 0606 build.
			aNodeName = Split(sNodeName,":",-1,1)
            If aNodeName(0) = "Organization" Then
					sNodeName = Replace(sNodeName,"Organization:","",1,1,1)
			End If
	End If
	If sSecurity<>"" Then
		aSecurity = split(sSecurity,"~")
		sSecurity=aSecurity(0)
	End If
	
	Set objProject = Fn_UI_ObjectCreate("Fn_PWC_ProjectOperations", JavaWindow("Project - Teamcenter 8"))
		Select Case sAction
				Case "Create","Modify","Create_Ext"
							If Trim(sAction) = "Create" Then
								Call Fn_Button_Click("Fn_PWC_ProjectOperations", objProject, "Clear")
'								Call Fn_Button_Click("Fn_PWC_ProjectOperations", objProject, "Refresh")		---------Pranav S.[Build:0606]								Call Fn_Button_Click("Fn_PWC_ProjectOperations", objProject, "Refresh")
							End If
							'Added code to click on Create button
							If Trim(sAction) = "Create_Ext" then
								sAction = "Create" 
							End If 							
							'Enter ID 									
							If sID<>"" Then
								Call Fn_Edit_Box("Fn_PWC_ProjectOperations",objProject,"ID",sID)
							End If
							'Enter Name
							If sName<>"" Then
								Call Fn_Edit_Box("Fn_PWC_ProjectOperations",objProject,"Name",sName)
							End If
							'Enter Description
							If sDescription<>"" Then
								Call Fn_Edit_Box("Fn_PWC_ProjectOperations",objProject,"Desc",sDescription)
							End If
							'Select Status
							If sStatus<>"" Then								
								Call Fn_UI_Object_SetTOProperty("Fn_PWC_ProjectOperations",objProject.JavaRadioButton("StatusActive"),"attached text",sStatus)
								Call Fn_UI_JavaRadioButton_SetON("Fn_PWC_ProjectOperations",objProject, "StatusActive")
							End If
							'Set Use Program security
							If sSecurity<>"" Then
								If Ucase(sSecurity)="ON" Then
									Call Fn_UI_JavaRadioButton_SetON("Fn_PWC_ProjectOperations",objProject, "Program")	  ' Design Change - Tc11.3-20170125-31_1_2017-JotibaT- Checkbox changed to Radio button - Discussed with Akshay J. 
								ElseIf Ucase(sSecurity)="OFF" Then
							 		Call Fn_UI_JavaRadioButton_SetON("Fn_PWC_ProjectOperations",objProject, "Project")	
								End If
								If UBound(aSecurity)>0Then
									objProject.JavaCheckBox("UsePrgmSec").SetToProperty "attached text","Inherit member selection from parent"
									  If aSecurity(1)="ON" Then
									  	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_PWC_ProjectOperations", "Set", objProject, "UsePrgmSec", aSecurity(1))
									  ElseIf  aSecurity(1)="OFF" Then
									  	Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_PWC_ProjectOperations", "Set", objProject, "UsePrgmSec", aSecurity(1))
									  End If
									  
									JavaWindow("DefaultWindow").JavaWindow("Warn").setToProperty "title","Warning"
									If JavaWindow("DefaultWindow").JavaWindow("Warn").Exist(1) Then
										JavaWindow("DefaultWindow").JavaWindow("Warn").JavaButton("OK").Click
									End If
								End If
							End If
'								Call Fn_CheckBox_Set("Fn_PWC_ProjectOperations", objProject, "UsePrgmSec", sSecurity)
'								JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin").SetTOProperty "title","Program Level Security"
'								If JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin").Exist(1) Then
'									JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin").JavaButton("OK").Click micLeftBtn
'								End If
							IF sMemberAction<>"" and sNodeName<>"" Then
								'Call Member selection function
								Call Fn_PWC_MemberSelection(sMemberAction, sNodeName, "")
							End if
'							Call Fn_ReadyStatusSync(1)
							'Click on Create button
							Call Fn_Button_Click("Fn_PWC_ProjectOperations", objProject, sAction)
							'Handle Dailog Box
							JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin").SetTOProperty "title","Missing Project Team Admin in Team"
							'If JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin").Exist Then
							If Fn_SISW_UI_Object_Operations("Fn_PWC_ProjectOperations","Exist",JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin"),SISW_MICRO_TIMEOUT) Then
								Call Fn_Button_Click("Fn_PWC_ProjectOperations", JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin"), "OK")
							Else
								JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin").SetTOProperty "title","Missing Program Team Admin in Team"
								'If JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin").Exist Then
								If Fn_SISW_UI_Object_Operations("Fn_PWC_ProjectOperations","Exist",JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin"),SISW_MICRO_TIMEOUT) Then
									Call Fn_Button_Click("Fn_PWC_ProjectOperations", JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin"), "OK")
								End If
							End If
							Call Fn_ReadyStatusSync(1)
				Case "Delete"
							'Click on delete button
							Call Fn_Button_Click("Fn_PWC_ProjectOperations", objProject, sAction)
							'Handle Confirm Delete Dialog Box.
							JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin").SetTOProperty "title","Confirm Delete"
							'Click on yes button
							Call Fn_Button_Click("Fn_PWC_ProjectOperations", JavaWindow("Project - Teamcenter 8").JavaWindow("Missing Project Team Admin"), "OK")
				Case "Verify"
						'Check Project ID
						If sID<>"" Then
							If Trim(Lcase(sID)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PWC_ProjectOperations",objProject,"ID"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "ID value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "ID value does not matches")
								Fn_PWC_ProjectOperations = False
								Exit Function
							End If
						End If
						'Check Project Name
						If sName<>"" Then							
							If Trim(Lcase(sName)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PWC_ProjectOperations",objProject,"Name"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Name value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Name value does not matches")
								Fn_PWC_ProjectOperations = False
								Exit Function
							End If
						End If
						'Check Project Description
						If 	trim(sDescription) <> "" Then
							If Trim(Lcase(sDescription)) = Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PWC_ProjectOperations",objProject,"Desc"))) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Description value matches")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Description value does not matches")
								Fn_PWC_ProjectOperations = False
								Exit Function
							End If
						End If 
						'Log to specify all values match correctly
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All specified values match actual values")
				Case "Exist","ExistWithStatus" 'added case to verify status of Program/project type [ ShwetaR_TC11.3_20170509d_NewDev_PWC ] 
							'The case is used to verify wheather the controls exist on the UI.
							bFlag = True
							'Check existence of RadioButtons of Status
							If sStatus<>"" Then
								aStatus = split(sStatus, ":",-1,1)
								For iCount=0 to ubound(aStatus)
									objProject.JavaRadioButton("StatusActive").SetTOProperty "attached text",aStatus(iCount)
									If objProject.JavaRadioButton("StatusActive").Exist(2) = True Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aStatus(iCount) &" JavaRadioButton is Visible")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aStatus(iCount) &" JavaRadioButton is not Visible")
										bFlag = False
									End If
								Next
							End If
							'Check existence of the UseProgramSecurity CheckBox.
							If sSecurity<>"" Then					
								If objProject.JavaRadioButton("Program").Exist(1) AND objProject.JavaRadioButton("Project").Exist(1) Then  ' Design Change - Tc11.3-20170125-31_1_2017-JotibaT- Checkbox changed to Radio button - Discussed with Akshay J. 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Program & Project JavaRadioButton is Visible")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Program & Project JavaRadioButton is not Visible")
									bFlag = False
								End If
							End If
							If sAction = "ExistWithStatus"  then
								if objProject.JavaRadioButton("StatusActive").GetROProperty("value") = 1 then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sStatus&" JavaRadioButton is true")
									bFlag = True
								else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sStatus&" Failed to verify JavaRadioButton is true")
									bFlag = False
								End if
							End if
						If bFlag=True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All values verified successfully")
						ElseIf bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All values does not match successfully")
							Fn_PWC_ProjectOperations = FALSE
							Set objProject = nothing 
							Exit Function
						End If												
				Case Else						
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PWC_ProjectOperations function failed")
							Fn_PWC_ProjectOperations = FALSE
							Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PWC_ProjectOperations")
	Fn_PWC_ProjectOperations = TRUE
	Set objProject = nothing
End Function
'*********************************************************		Function to Check the buttons are enabled / disabled		**********************************************************************
'Function Name		:				Fn_PWC_CheckButtonsdisabled(sReferencePath,sButtons)

'Description			 :		 		 To Check the button is enabled or disabled.

'Parameters			   :	 			1) sReferencePath
'													 2) sButtons
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Project prespective should be displayed.

'Examples				:				Call Fn_PWC_CheckButtonsdisabled(JavaWindow("Project - Teamcenter 8"),"Create:Modify:Copy")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					10-08-10			1.0	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_PWC_CheckButtonsdisabled(sReferencePath,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_CheckButtonsdisabled"
   Dim aButtons,intCount,iCounter,bFlag
	bFlag = True
		aButtons = split(sButtons, ":",-1,1)
		intCount = Ubound(aButtons)
		For iCounter=0 to intCount
				If Fn_UI_Object_GetROProperty("Fn_PWC_CheckButtonsdisabled",sReferencePath.JavaButton(aButtons(iCounter)), "enabled")="0" Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The "&aButtons(iCounter)&" button is disabled")					
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The "&aButtons(iCounter)&" button is enabled")
					bFlag = False							
				End If
		Next
		If bFlag=False Then
			Fn_PWC_CheckButtonsdisabled = False
		Else
			Fn_PWC_CheckButtonsdisabled = True
		End If		
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The function Fn_PWC_CheckButtonsdisabled completed succesfully")
End Function
'*********************************************************		Fn_PWC_NonProjectAdministratorMsgVerify 	**********************************************************************
'Function Name		:				Fn_PWC_NonProjectAdministratorMsgVerify(sErrorText)

'Description			 :		 		 To verify the message in NonProjectAdministrator Dialog and Click Ok .

'Parameters			   :	 		sErrorText
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		NonProjectAdministrator Dialog Should Exist

'Examples				:				Call Fn_PWC_NonProjectAdministratorMsgVerify("You are not a Project Administrator")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Harshal																			Handle one more msgbox
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_PWC_NonProjectAdministratorMsgVerify(sErrorText)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_NonProjectAdministratorMsgVerify"
	GBL_EXPECTED_MESSAGE=sErrorText
   Dim sMsg, objproject, iCount
		For iCount=0 to 0
					JavaWindow("Project - Teamcenter 8").JavaWindow("non-Project Administrator").SetTOProperty "title","non-Project Administrator Access"
					If JavaWindow("Project - Teamcenter 8").JavaWindow("non-Project Administrator").Exist(3) Then
						Set objproject = Fn_UI_ObjectCreate("Fn_PWC_NonProjectAdministratorMsgVerify", JavaWindow("Project - Teamcenter 8").JavaWindow("non-Project Administrator"))
						Exit For
					End If
					JavaWindow("Project - Teamcenter 8").JavaWindow("non-Project Administrator").SetTOProperty "title","Project - Teamcenter 8"
					If JavaWindow("Project - Teamcenter 8").JavaWindow("non-Project Administrator").Exist(3) Then
						Set objproject = Fn_UI_ObjectCreate("Fn_PWC_NonProjectAdministratorMsgVerify", JavaWindow("Project - Teamcenter 8").JavaWindow("non-Project Administrator"))
						Exit For
					End If
					'[TC1015-20151013b00-18_11_2015-VivekA-Maintenance] - Added by Poonam C
					JavaWindow("Project - Teamcenter 8").Dialog("non-Project Administrator Dialog").SetTOProperty "title","non-Project Administrator Access"
					If JavaWindow("Project - Teamcenter 8").Dialog("non-Project Administrator Dialog").Exist(3) Then
						Set objproject = Fn_UI_ObjectCreate("Fn_PWC_NonProjectAdministratorMsgVerify", JavaWindow("Project - Teamcenter 8").Dialog("non-Project Administrator Dialog"))
						Exit For
					End If
					'---------------------------------------------------
		Next
	If objproject.Exist(5)Then
	 
		  If JavaWindow("Project - Teamcenter 8").JavaWindow("non-Project Administrator").JavaStaticText("ErrorMessage").Exist(3) Then
		  	 sMsg = objproject.JavaStaticText("ErrorMessage").GetROProperty("attached text")
		  Else
             sMsg = objproject.Static("ErrorMessage").GetROProperty("text")      		  
		  End If
	 
		If instr(1,sMsg,sErrorText)<> 0 Then
		    If JavaWindow("Project - Teamcenter 8").JavaWindow("non-Project Administrator").JavaButton("OK").Exist(3) Then
		   		objproject.JavaButton("OK").Click
		   	Else
		   		objproject.WinButton("OK").Click
            End if	
			Fn_PWC_NonProjectAdministratorMsgVerify = True
		Else
			GBL_ACTUAL_MESSAGE=sMsg
			Fn_PWC_NonProjectAdministratorMsgVerify = False
		End If
	Else
		Fn_PWC_NonProjectAdministratorMsgVerify = False
	End If
End Function
'#########################################################################################################
'###
'###    FUNCTION  NAME   :   Fn_PWC_TeamAdmin(sUser,sButtons)
'###
'###    DESCRIPTION        :   Select a user as Team Admin.
'###
'###    PARAMETERS      :   1. sUser: 
'###											 2.	sButtons:
'###                                         
'###    Function Calls       :   Fn_WriteLogFile(), Fn_Button_Click(), Fn_UI_ObjectCreate()
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           11/08/2010         1.0
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		
'#############################################################################################################
Public Function Fn_PWC_TeamAdmin(sUser,sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_TeamAdmin"
	Dim objproject, iCounter, bReturn, aButtons, iCount, bFlag
	bFlag = False
	'Click on Select Team Admin button
	Call Fn_Button_Click("Fn_PWC_TeamAdmin", JavaWindow("Project - Teamcenter 8"), "PrivilegedMem")
	Set objproject = Fn_UI_ObjectCreate("Fn_PWC_TeamAdmin", JavaWindow("Project - Teamcenter 8").JavaDialog("TeamAdministrator"))
			If sUser<>"" Then
					bReturn = objproject.JavaList("TeamAdminList").GetROProperty("items count")
					'Extract the index of row at which the object exist.
						For iCounter=0 to bReturn-1
							If Trim(lcase(objproject.JavaList("TeamAdminList").GetItem(iCounter))) = Trim(lcase(sUser)) then
								objproject.JavaList("TeamAdminList").Select sUser
								bFlag = True
								Exit For 
							End If
						Next
						If iCounter=bReturn-1 and bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add Team Admin")
							Fn_PWC_TeamAdmin = False
							Set objproject = nothing
							Exit Function
						End If
			End If
			'Click on Buttons
			If sButtons<>"" Then
					aButtons = split(sButtons, ":",-1,1)
					iCounter = Ubound(aButtons)
					For iCount=0 to iCounter
						'Click on Add Button
						Call Fn_Button_Click("Fn_PWC_TeamAdmin", objproject, aButtons(iCount))
					Next
			End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully completed function Fn_PWC_TeamAdmin")
	Fn_PWC_TeamAdmin = TRUE
    Set objproject = nothing 	
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_PWC_SearchObjectVerify
'###
'###    DESCRIPTION    : Operation related to search in project prespective
'###    
'###    PARAMETERS      : sAction,sMsg,sButton
'###               
'###	 HISTORY       :   AUTHOR     		DATE        	VERSION
'###
'###    CREATED BY     :    Harshal Agrawal 24/08/2010		1.0
'###
'###    REVIWED BY     :    Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE       : 'MsgBox Fn_PWC_SearchObjectVerify("ReloadTreeDialog","149 in the project tree","Yes")
'### 					'MsgBox Fn_PWC_SearchObjectVerify("InvalidCriteria","No projects found","OK")
'### 					'MsgBox Fn_PWC_SearchObjectVerify("NoCriteria","","")
'### 					'MsgBox Fn_PWC_SearchObjectVerify("ReloadButtonToolTipVerify","project","")
'### 					'MsgBox Fn_PWC_SearchObjectVerify("SearchButtonToolTipVerify","project","")
'### 					'MsgBox Fn_PWC_SearchObjectVerify("SearchEditBoxToolTipVerify","project","") 		
'#############################################################################################################
Function Fn_PWC_SearchObjectVerify(sAction,sMsg,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_SearchObjectVerify"
Dim sAppMsg,ObjSearch,objShell
Set ObjSearch = Nothing
Select Case sAction
	Case "ReloadTreeDialog"
			If JavaDialog("Search").Exist(5) Then
				If sMsg <> "" Then
					sAppMsg = JavaDialog("Search").JavaObject("MLabel").Object.getText
					If instr(1,sAppMsg,sMsg)<> 0 Then
						JavaDialog("Search").JavaButton(sButton).Click
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Sucessful")
						Fn_PWC_SearchObjectVerify = True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Failed")
						Fn_PWC_SearchObjectVerify = False
					End If
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Failed")
				Fn_PWC_SearchObjectVerify = False
			End If
	Case "NoCriteria"
		JavaWindow("Project - Teamcenter 8").Dialog("ErrorDialog").Close
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Sucessful")
		Fn_PWC_SearchObjectVerify = True
		
	Case "InvalidCriteria"
        Set ObjSearch = JavaWindow("Project - Teamcenter 8").JavaWindow("Search")
		If ObjSearch.Exist(8)Then
				If sMsg <> "" Then
					sAppMsg = ObjSearch.JavaStaticText("No projects found").GetROProperty("attached text")
					If instr(1,sAppMsg,sMsg)<> 0 Then
						ObjSearch.JavaButton(sButton).Click
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Sucessful")
						Fn_PWC_SearchObjectVerify = True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Failed")
						Fn_PWC_SearchObjectVerify = False
					End If
				End If
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Failed")
			Fn_PWC_SearchObjectVerify = False
		End If
    Case "ReloadButtonToolTipVerify"
		Set ObjSearch = JavaWindow("Project - Teamcenter 8")
		sAppMsg =ObjSearch.JavaButton("ReloadPrjct").GetROProperty("tool_tip_text")
		If instr(1,sAppMsg,sMsg)<> 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Sucessful")
			Fn_PWC_SearchObjectVerify = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Failed")
			Fn_PWC_SearchObjectVerify = False
		End if
    Case "SearchButtonToolTipVerify"
		Set ObjSearch = JavaWindow("Project - Teamcenter 8")
		sAppMsg = ObjSearch.JavaButton("PrjctFind").GetROProperty("tool_tip_text")
		If instr(1,sAppMsg,sMsg)<> 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Sucessful")
			Fn_PWC_SearchObjectVerify = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Failed")
			Fn_PWC_SearchObjectVerify = False
		End if
    Case "SearchEditBoxToolTipVerify"
		Set ObjSearch = JavaWindow("Project - Teamcenter 8")
		sAppMsg = ObjSearch.JavaEdit("PrjctSrch").GetROProperty("tool_tip_text")
		If instr(1,sAppMsg,sMsg)<> 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Sucessful")
			Fn_PWC_SearchObjectVerify = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function:Fn_PWC_SearchObjectVerify Failed")
			Fn_PWC_SearchObjectVerify = False
		End if
   End Select
Set ObjSearch = Nothing
End Function

'#########################################################################################################
'###    FUNCTION NAME   :   Fn_PWC_WorkContextMessageVerify(sAction, sMessage)
'###    Eliminated .Not used anywhere.
'#############################################################################################################

'*********************************************************		Function Assigning a Work Context		***********************************************************************
'Function N	ame		:				Fn_PWC_AssignWorkContext

'Description			 :		 		 Assign 

'Parameters			   :                1. sAction : 
'													2. sWorkCnxt : 
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		An Item should be selected.

'Examples				:				Call Fn_PWC_AssignWorkContext("Assign", "TestingK10", "OK")

'History:
'										Developer Name			Date				Rev. No.			Changes Done				Reviewer		Reviewer Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					03-Sept-2010		1.0																	Harshal			03-Sept-2010
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_PWC_AssignWorkContext(sAction, sWorkCnxt, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_AssignWorkContext"
	Dim breturn, objWC, iCounter, sText, bFlag, iCount, aButtons, iRowData
	bFlag = false
	Select Case sAction
		Case "Assign"
				breturn = Fn_UI_ObjectExist("Fn_FormCreate", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Assign Work Context"))
				If breturn = false Then
					Call Fn_MenuOperation("Select","Tools:Assign Work Context...")
					Call Fn_ReadyStatusSync(3)
				End If
				Set objWC = Fn_UI_ObjectCreate("Fn_FormCreate", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Assign Work Context"))
	            'Set Name
                call Fn_Edit_Box("Fn_PWC_AssignWorkContext",objWC,"Name",sWorkCnxt)
				Wait(5)
				'Find the Work Context
				objWC.JavaEdit("Name").Activate
				'Wait still the status is Ready
				Call Fn_ReadyStatusSync(4)
				If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Assign Work Context").JavaButton("LoadAll").GetROProperty("enabled") = 1 Then
					'Click on LoadAll Button
					call Fn_Button_Click("Fn_PWC_AssignWorkContext",objWC,"LoadAll")
				End If
				'Wait still the status is Ready
				Call Fn_ReadyStatusSync(4)
				'Count number of rows of Table
				bReturn = objWC.JavaTable("Name").GetROProperty("rows")	
				'Extract the index of row at which the object exist.
				For iCounter=0 to bReturn - 1
				sText = objWC.JavaTable("Name").GetCellData(iCounter,"Object")						
					If IsNumeric(sWorkCnxt) Then
						 If cstr(sText) = cstr(cint(sWorkCnxt))  Then
							 objWC.JavaTable("Name").ClickCell iCounter,"Object","LEFT"
							 bFlag = True
							 Exit for
						End If
					elseIf cstr(sText) = cstr(sWorkCnxt)  Then
							 objWC.JavaTable("Name").ClickCell iCounter,"Object","LEFT"
							 bFlag = True
							 Exit for
					End If									
				Next
				'Click on Apply/OK or Cancel button
				If sButtons<>"" Then
						aButtons = split(sButtons, ":",-1,1)
						iCount = Ubound(aButtons)
						For iRowData=0 to iCount
							'Click on Add Button
							Call Fn_Button_Click("Fn_PWC_AssignWorkContext", objWC, aButtons(iRowData))
						Next
				End If
				If bFlag = false Then
						Fn_PWC_AssignWorkContext = FALSE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_PWC_AssignWorkContext : Row with Object "&sWorkCnxt&" does not exist")	
				Else 
						Fn_PWC_AssignWorkContext = TRUE				 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_MyTc_DetailTableContentOperation: Row with Object "&sWorkCnxt&" exist")	
				End If						
		Case else
			Fn_PWC_AssignWorkContext =False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to assign Work Context in Fn_PWC_AssignWorkContext")	
	End Select
	Set objWC = Nothing
End Function
'######################################################################################################################################
'###    FUNCTION NAME   :  Fn_PWC_PropertyOperations(sAction, bSubGroup, SComments, sDateArch, sDateCreated, sDateLstBack, sDesc, sGroup, sGroupID, sGroupMem, sLstModiDate, sLstModiUser, sName1, sName2, sObject, sOwner, sOwningSite, sProject, sRole, sUserSetModi, sButtons)
'###
'###    DESCRIPTION     :  	Edit Properties of Work-Context
'###
'###    Function Calls  :  		Fn_UI_ObjectCreate(), Fn_UI_ObjectExist(), Fn_Edit_Box(), Fn_WriteLogFile()
'###
'###	 HISTORY         :   		AUTHOR                 DATE        VERSION
'###
'###    CREATED BY      :     Ketan Raje      		   06/09/10      1.0
'###
'###    REVIWED BY      :    Harshal							 	 
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Call Fn_PWC_PropertyOperations("Set", "", "", "", "", "", "", "", "", "", "", "", "ModiName", "", "", "", "", "", "", "", "OK")
'######################################################################################################################################
Function Fn_PWC_PropertyOperations(sAction, bSubGroup, SComments, sDateArch, sDateCreated, sDateLstBack, sDesc, sGroup, sGroupID, sGroupMem, sLstModiDate, sLstModiUser, sName1, sName2, sObject, sOwner, sOwningSite, sProject, sRole, sUserSetModi, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_PropertyOperations"
   Dim objWrkCtxt, aButtons, iCounter, iCount
   Set objWrkCtxt = Fn_UI_ObjectCreate("Fn_PWC_PropertyOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"))
	'Click on Static text
	If Fn_UI_ObjectExist("Fn_PWC_PropertyOperations", objWrkCtxt.JavaStaticText("BottomLink")) = True Then
		objWrkCtxt.JavaStaticText("BottomLink").Click 1,1,"LEFT"
	End If
	Select Case sAction
		Case "Set"
				'The Above part is to be coded as required.
				If sName1<>"" Then
					'Set the Name Value
					Call Fn_Edit_Box("Fn_PWC_PropertyOperations", objWrkCtxt, "Name", sName1)	
				End If
				'The Below part is to be coded as required.
				Fn_PWC_PropertyOperations = True
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to execute Function [ Fn_PWC_PropertyOperations ] Invalid case [ " & sAction & " ] ")
				Fn_PWC_PropertyOperations = False
	End Select
	'Click on Buttons
	If sButtons<>"" Then
			aButtons = split(sButtons, ":",-1,1)
			iCounter = Ubound(aButtons)
			For iCount=0 to iCounter
				'Click on Add Button
				Call Fn_Button_Click("Fn_PWC_PropertyOperations", objWrkCtxt, aButtons(iCount))
			Next
	End If
Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS : Function [ Fn_PWC_PropertyOperations ] executed successfully with case [ " & sAction & " ] ")   	
Set objWrkCtxt = Nothing
End Function
'###########################################     (Function to Check Property Editable)      ###############################################
'#
'# 	Function Name		:				Fn_PWC_ObjectPropertyIsEditable()

'#	Description			 :		 		     Validate that the Object Property value filed associated to Property is editable 
'#
'#	Parameters			   :	 		   1) sAction
'#												  2) sPropertyName:Name of the Property	
'#												  3) sButtons

'#	Return Value		   : 				TRUE (if property is enabled)\ FALSE (if property is disabled)
'#
'#	Pre-requisite			:		 		  Object Properties window is open.
'#
'#	Examples				:			     Msgbox Fn_PWC_ObjectPropertyIsEditable("JavaEdit","Name:","Cancel")
'#
'#	History:
'#	Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'###############################################################################################################################
'#	Ketan Raje				15-Sept-2010   		1.0													Harshal
'###############################################################################################################################
Public Function Fn_PWC_ObjectPropertyIsEditable(sAction, sPropertyName, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_ObjectPropertyIsEditable"
Dim objDesc, intNoOfObjects, sPropName, arrPropName, iCnt, aButtons, iCounter, iCount
Select Case sAction
			Case "JavaEdit"
							Set objDesc = Description.Create()
							objDesc("Class Name").value = "JavaEdit"
							Set  intNoOfObjects =  JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").ChildObjects(objDesc)
							  For iCnt = 0 to intNoOfObjects.count-1
								  sPropName = intNoOfObjects(iCnt).getROProperty("attached text")
								  arrPropName = Split(sPropName,"~")
								   If  arrPropName(0) = sPropertyName Then
										If intNoOfObjects(iCnt).getROProperty("enabled") = 1 Then
											Fn_PWC_ObjectPropertyIsEditable = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sPropertyName + " Property is disabled")
											Exit for
										Else
											Fn_PWC_ObjectPropertyIsEditable = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sPropertyName + " Property is enabled")
										End If
									End If
							  Next
							  Set objDesc = Nothing
							  Set intNoOfObjects = Nothing
			Case Else						
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PWC_ObjectPropertyIsEditable function failed")
							Fn_PWC_ObjectPropertyIsEditable = FALSE
							Exit Function											
		End Select
		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)
				iCounter = Ubound(aButtons)
				For iCount=0 to iCounter
					'Click on Add Button
					Call Fn_Button_Click("Fn_PWC_ObjectPropertyIsEditable", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"), aButtons(iCount))
				Next
		End If
End Function
'*********************************************************		Function to verify  dialog error message.	***********************************************************************
'Function Name		:					Fn_PWC_DialogMsgVerify

'Description			 :		 		  This function is used to get Schedule table Row Index.

'Parameters			   :	 			1.  sTitle:Title of dialog.
'													2. sMsg : Message to verify. (Optional)
'													3. sButton : Button Name.
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:			  Msgbox Fn_PWC_DialogMsgVerify("Error","does not have any group member for the given group","OK") 
'											Msgbox Fn_PWC_DialogMsgVerify("No Group Member","no group member required by the selected work context","OK") 

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje				17-Sept-2010		1.0														Harshal
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PWC_DialogMsgVerify(sTitle,sMsg,sButton) 

    Dim dicErrorInfo
	 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	 dicErrorInfo.Add "Action", "VerifyUsingDescription"
	 dicErrorInfo.Add "Title", sTitle
	 dicErrorInfo.Add "Message", sMsg
	 dicErrorInfo.Add "Button", sButton    
	 Fn_PWC_DialogMsgVerify = Fn_SISW_PWC_ErrorVerify(dicErrorInfo)
	 Set dicErrorInfo = Nothing

End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_PWC_CheckOutMessageVerify(sAction, aObject, aMessage)
'###
'###    DESCRIPTION     :   Verify the error messages popping up after Check-Out operation
'###
'###    PARAMETERS      :   sAction
'###									aObject
'###									aMessage
'###
'###    Return Value  	:   	True/False 
'###
'###    HISTORY         :   	AUTHOR              	DATE        		VERSION
'###
'###    CREATED BY      :   Ketan Raje				15-Sept-2010   			1.0
'###
'###    REVIWED BY      :	Harshal					15-Sept-2010   			
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Msgbox Fn_PWC_CheckOutMessageVerify("ButtonToolTip", "StartTestWC", "Object  is of a type not supported by Check-Out facility.")
'#############################################################################################
Public Function Fn_PWC_CheckOutMessageVerify(sAction, aObject, aMessage)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_CheckOutMessageVerify"
	GBL_EXPECTED_MESSAGE=aMessage
Dim objDialog, sObj, txtMsg, sElement, iCount, bFlag, sMsg, iCountChecked
iCountChecked = 0
If Fn_UI_ObjectExist("Fn_RefreshWindow", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))=False Then
    Call Fn_MenuOperation("Select","Edit:Properties")
End If
Set objDialog = Fn_UI_ObjectCreate("Fn_PWC_CheckOutMessageVerify", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out"))
'Click on yes button 
Call Fn_Button_Click("Fn_PWC_CheckOutMessageVerify", objDialog, "Yes")
sObj = Split(aObject, ":", -1, 1)
txtMsg = Split(aMessage, ":", -1, 1)
Select Case sAction
    Case "ButtonToolTip"
    			objDialog.JavaButton("CheckOutErrorBtn").SetTOProperty "Index",iCount
				objDialog.Click 1,1,"LEFT"
				sMsg = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Check-Out").JavaButton("CheckOutErrorBtn").GetROProperty("tool_tip_text")
						If InStr(1, sMsg, txtMsg(iCount), 1) > 0 Then
							bFlag = True
						Else
							GBL_ACTUAL_MESSAGE=sMsg
							bFlag = False
						End If			
    Case "Label"
				If InStr(1, objDialog.JavaObject("CheckOutErrorMessage").GetROProperty("attached text"), aMessage, 1) > 0 Then
						bFlag = True
				Else 
						GBL_ACTUAL_MESSAGE=objDialog.JavaObject("CheckOutErrorMessage").GetROProperty("attached text")
				End If
End Select
	Call Fn_Button_Click("Fn_PWC_CheckOutMessageVerify", objDialog, "OK")
		If bFlag = True Then
				Fn_PWC_CheckOutMessageVerify = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass: CheckOut Error Message Verified Successfully. ")
		Else
				Fn_PWC_CheckOutMessageVerify = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"FAIL: CheckOut Error Message Not Verified. ")				
		End If
Set objDialog = Nothing
Set sElement = Nothing
End Function
''*********************************************************		Function to perform action on Smart Folder Filter Configuration Tree	***********************************************************************
'Function Name		:				Fn_PWC_SmartFolder_TreeOpeartion()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. GetIndex	
'																	5. IsSelected

'Parameters			   :	 			1. sAction: Action to be performed
'											   2. sNodeName: Fully qulified tree Path (delimiter as ':') 

'Return Value		   : 				TRUE / FALSE and Index Value in "GetIndex" case.

'Pre-requisite			:		 		AccessManager Prespective is Open.

'Examples				:			 Case "Select" : Call Fn_PWC_SmartFolder_TreeOpeartion("Select","AutoPrj1892010155455")
'											Case "Expand" : Call Fn_PWC_SmartFolder_TreeOpeartion("Expand","Y11")
'											Case "Collapse" : Call Fn_PWC_SmartFolder_TreeOpeartion("Collapse","Y11")
'											Case "Exist" : Call Fn_PWC_SmartFolder_TreeOpeartion("Exist","Y11:Y11_A")
'											Case "GetIndex" : Call Fn_PWC_SmartFolder_TreeOpeartion("GetIndex","Y11:Y11_A")
'											Case "IsSelected" : Call Fn_PWC_SmartFolder_TreeOpeartion("IsSelected","Y11:Y11_A")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			20/09/2010			              1.0								Created									Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_PWC_SmartFolder_TreeOpeartion(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_SmartFolder_TreeOpeartion"
	Dim objJavaWindowSmt, objJavaTreeSmtFold, intNodeCount, intCount, sTreeItem, sItemName, iRow
	Set objJavaWindowSmt = Fn_UI_ObjectCreate( "Fn_PWC_SmartFolder_TreeOpeartion",JavaWindow("Project - Teamcenter 8"))
	
	'Swapnil:: Added the call to select the "Smart Folder Filter Configuration" Tab
	'
	'If  JavaWindow("Project - Teamcenter 8").JavaTree("SmartFolderTree").GetROProperty ("width") < 15 Then					'----Added code to check whether Tree is expanded or not  By:- Pranav S[04-07-12]
	'	Call Fn_TabFolder_Operation("DoubleClickTab", "Smart Folder Filter Configuration", "")
	'End If
	If Fn_SISW_UI_RACTabFolderWidget_Operation("IsMaximized","Smart Folder Filter Configuration","")=False Then
		Call Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", "Smart Folder Filter Configuration", "")
	End If
	
	Select Case sAction
		'------- For selecting single node --------------------------------
		Case "Select"
				Fn_PWC_SmartFolder_TreeOpeartion = Fn_JavaTree_Select("Fn_PWC_SmartFolder_TreeOpeartion", objJavaWindowSmt, "SmartFolderTree",sNodeName)
		'------- For expanding a particular  node--------------------------
		Case "Expand"
				Fn_PWC_SmartFolder_TreeOpeartion = Fn_UI_JavaTree_Expand("Fn_PWC_SmartFolder_TreeOpeartion",objJavaWindowSmt,"SmartFolderTree",sNodeName)
		'------- For collapssing a particular  node------------------------
		Case "Collapse"
				Fn_PWC_SmartFolder_TreeOpeartion = Fn_UI_JavaTree_Collapse("Fn_PWC_SmartFolder_TreeOpeartion", objJavaWindowSmt,"SmartFolderTree",sNodeName)
		'------- For Checking existance of a particular  node--------------
		Case "Exist"
				Set objJavaTreeSmtFold = Fn_UI_ObjectCreate( "Fn_PWC_SmartFolder_TreeOpeartion", objJavaWindowSmt.JavaTree("SmartFolderTree"))
					intNodeCount = objJavaTreeSmtFold.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeSmtFold.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_PWC_SmartFolder_TreeOpeartion = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_PWC_SmartFolder_TreeOpeartion = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowSmt.JavaTree("SmartFolderTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowSmt.JavaTree("SmartFolderTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_PWC_SmartFolder_TreeOpeartion = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_PWC_SmartFolder_TreeOpeartion = FALSE
				End If
		'----------------------------------------------------------------------- Check wheather particular node is Selected or Not-------------------------------------------------------------------------
		Case "IsSelected"
			wait(3)
			Set objJavaTreeSmtFold = Fn_UI_ObjectCreate( "Fn_PWC_SmartFolder_TreeOpeartion", objJavaWindowSmt.JavaTree("SmartFolderTree"))				
				If Trim(Lcase(objJavaTreeSmtFold.GetROProperty("value"))) = Trim(Lcase(sNodeName)) Then
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Java Tree Node ["+sNodeName+"] is Selected .")
				   Fn_PWC_SmartFolder_TreeOpeartion = TRUE
				Else
				   'Writing Log
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Java Tree Node ["+sNodeName+"] is Not Selected .")
				   Fn_PWC_SmartFolder_TreeOpeartion = FALSE
			End If
		Case Else
			Fn_PWC_SmartFolder_TreeOpeartion = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PWC_SmartFolder_TreeOpeartion function failed >> Invalid Case")
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_PWC_SmartFolder_TreeOpeartion")
	
	If Fn_SISW_UI_RACTabFolderWidget_Operation("IsMaximized","Smart Folder Filter Configuration","")=True Then
	  Call Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", "Smart Folder Filter Configuration", "")
	End If 
	Set objJavaWindowSmt = nothing
	Set objJavaTreeSmtFold = nothing
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_PWC_InsufficentPrivilegeErrorVerify(sTitle,sDetails)
'###
'###    DESCRIPTION     :   Verify the error messages popping up after Check-Out operation
'###
'###    PARAMETERS      :   sTitle
'###									sDetails
'###									
'###
'###    Return Value  	:   	True/False 
'###
'###    HISTORY         :   	AUTHOR              	DATE        		VERSION
'###
'###    CREATED BY      :   Harshal				20-Sept-2010   			1.0
'###
'###    REVIWED BY      :	Harshal					20-Sept-2010   			
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Msgbox Fn_PWC_InsufficentPrivilegeErrorVerify("Project","You have insufficient privilege to assign object")
'#############################################################################################
Function Fn_PWC_InsufficentPrivilegeErrorVerify(sTitle,sDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_InsufficentPrivilegeErrorVerify"
	GBL_EXPECTED_MESSAGE=sDetails
	Dim objErr,sAppDetails
	Dim objWin
	Set objErr = JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").JavaWindow("ProjectErrorDialog")
	Set objWin = JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").JavaDialog("Project")

	If objWin.Exist(5) Then
		objWin.JavaCheckBox("More...").Set "ON"
		sAppDetails = objWin.JavaEdit("DetailedMsg").GetROProperty("value")
		objWin.JavaButton("OK").Click
	ElseIf objErr.Exist(5) Then
		sAppDetails = objErr.JavaEdit("Details").GetROProperty("value")
		objErr.JavaButton("OK").Click micLeftBtn
	Else
		Fn_PWC_InsufficentPrivilegeErrorVerify = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Falied verify the Error Dialog")
	End If

	If instr(1,sAppDetails,sDetails)<>0 Then
		Fn_PWC_InsufficentPrivilegeErrorVerify =True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Sucessfully verified the Msg")	
	Else
		GBL_ACTUAL_MESSAGE=sAppDetails
		Fn_PWC_InsufficentPrivilegeErrorVerify = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Falied verify the Msg")
	End If

	Set objErr = nothing
	Set objWin = nothing
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_PWC_FilterAssociationTable_Operations(sAction, bContribute, sName, sSourceType, sProperty, sValue, iRow, iColumn)
'###
'###    DESCRIPTION     :   Cases Related to AssociationTable
'###
'###    Return Value  	:   	True/False 
'###
'###    HISTORY         :   	AUTHOR              	DATE        		VERSION 	Build
'###
'###    CREATED BY      :   Ketan Raje				21-Sept-2010   			1.0			902
'###
'###    REVIWED BY      :	Harshal					21-Sept-2010   			1.0			902
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Case "IsList" : Msgbox Fn_PWC_FilterAssociationTable_Operations("IsList", "", "", "", "", "", "0", "2")
'### 								 Case "AddFilter" : Msgbox Fn_PWC_FilterAssociationTable_Operations("AddFilter", "OFF", "", "", "", "", "", "")
'### 								 Case "RemoveFilter" : Msgbox Fn_PWC_FilterAssociationTable_Operations("RemoveFilter", "ON", "", "", "", "", "0", "")
'### 								 Case "Verify" : Msgbox Fn_PWC_FilterAssociationTable_Operations("Verify", "", "G10293-FrameBacklit Glass", "Item", "object_type", "Ketan", "", "")
'#############################################################################################
Public Function Fn_PWC_FilterAssociationTable_Operations(sAction, bContribute, sName, sSourceType, sProperty, sValue, iRow, iColumn)
		GBL_FAILED_FUNCTION_NAME="Fn_PWC_FilterAssociationTable_Operations"
		Dim iRORows,iLen,i,WshShell,objTable,iRowCount,iColCount,iRows
		'Call Fn_SISW_UI_RACTabFolderWidget_Operation("RMBMenuSelect", "Smart Folder Filter Configuration", "Maximize")
		Call Fn_SISW_UI_RACTabFolderWidget_Operation("DoubleClick", "Smart Folder Filter Configuration", "")
	 	If bContribute="ON" Then
				'Set the Contribute Check Box.
				Call Fn_CheckBox_Set("Fn_PWC_FilterAssociationTable_Operations", JavaWindow("Project - Teamcenter 8"), "Contribute", bContribute)
				Call Fn_ReadyStatusSync(2)
		End If
		Select Case sAction
				Case "Cleanup"
						iRowCount = cInt(JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable").GetROProperty("rows"))
						if iRowCount > 0 then
							For iRows=0 to iRowCount-1
								JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable").ClickCell 0, 0
								wait 1
								Call Fn_Button_Click("Fn_PWC_FilterAssociationTable_Operations", JavaWindow("Project - Teamcenter 8"), "RemoveFilter")	
							Next
							'Call Fn_Button_Click("Fn_PWC_FilterAssociationTable_Operations", JavaWindow("Project - Teamcenter 8"), "SaveFilterSetting")
							Call Fn_SISW_UI_JavaButton_Operations("Fn_PWC_FilterAssociationTable_Operations", "Object.click", JavaWindow("Project - Teamcenter 8"),"SaveFilterSetting")
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Failed to comple of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = True
							End If
						Else
							Fn_PWC_FilterAssociationTable_Operations = True
						End If
				
				Case "IsList"
							'Select particular Cell of Filter Association Table
							JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable").SelectCell iRow,iColumn
							Wait(2)
							'Click particular Cell of Filter Association Table
							JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable").SelectCell iRow,iColumn
							Wait(2)
							If JavaWindow("Project - Teamcenter 8").JavaList("FilterAssociationList").Exist(2) = True Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : JavaList is present at the given location")
								Fn_PWC_FilterAssociationTable_Operations = True
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : JavaList is not present at the given location")
								Fn_PWC_FilterAssociationTable_Operations = False
							End If		
				Case "AddFilter"
						'Click on Add Filter Button
						Call Fn_Button_Click("Fn_PWC_FilterAssociationTable_Operations", JavaWindow("Project - Teamcenter 8"), "AddFilter")
'							While JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").GetROProperty("enabled") = 1
'								'Click on Save Filter Settings Button
'								Call Fn_Button_Click("Fn_PWC_FilterAssociationTable_Operations", JavaWindow("Project - Teamcenter 8"), "SaveFilterSetting")
'							Wend
						JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").WaitProperty "enabled", 1, 20
						If CBool(JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").GetROProperty("enabled")) = True Then
							JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").Click micLeftBtn
						End If
						'Write Log to Return Value.
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Failed to comple of function Fn_PWC_FilterAssociationTable_Operations")
							Fn_PWC_FilterAssociationTable_Operations = False
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_PWC_FilterAssociationTable_Operations")
							Fn_PWC_FilterAssociationTable_Operations = True
						End If
				Case "RemoveFilter"
						'Click particular Row of Filter Association Table
						JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable").ActivateRow iRow
						'Click on Remove Filter Button
						Call Fn_Button_Click("Fn_PWC_FilterAssociationTable_Operations", JavaWindow("Project - Teamcenter 8"), "RemoveFilter")

						JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").WaitProperty "enabled", 1, 20
						If CBool(JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").GetROProperty("enabled")) = True Then
							'JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").Click micLeftBtn
							Call Fn_SISW_UI_JavaButton_Operations("Fn_PWC_FilterAssociationTable_Operations", "Object.click", JavaWindow("Project - Teamcenter 8"),"SaveFilterSetting")
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Failed to comple of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = True
							End If
						End If
						'Write Log to Return Value.
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Failed to comple of function Fn_PWC_FilterAssociationTable_Operations")
							Fn_PWC_FilterAssociationTable_Operations = False
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_PWC_FilterAssociationTable_Operations")
							Fn_PWC_FilterAssociationTable_Operations = True
						End If
				Case "AddFilterCriteria"
							Set objTable = JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable")
							If Fn_SISW_UI_Object_Operations("","Enabled", JavaWindow("Project - Teamcenter 8").JavaButton("AddFilter"), SISW_MAX_TIMEOUT) = True Then
								JavaWindow("Project - Teamcenter 8").JavaButton("AddFilter").Click micLeftBtn
							End If
							iRORows = objTable.GetROProperty("rows") - 1
							objTable.SelectCell iRORows,"Source Type"
							Wait(3)
							objTable.SelectCell iRORows,"Source Type"
							JavaWindow("Project - Teamcenter 8").JavaList("FilterAssociationList").Select sSourceType
							objTable.SelectCell iRORows,"Property"
							Wait(3)
							objTable.SelectCell iRORows,"Property"
							JavaWindow("Project - Teamcenter 8").JavaList("FilterAssociationList").Select sProperty
							
							wait(2)
							objTable.SelectCell iRORows,"Value"
							wait 1
							objTable.type sValue
							objTable.SelectCell iRORows,"Name"
							Set objTable = Nothing 

							JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").WaitProperty "enabled", 1, 20

							Dim exitCnt
							exitCnt = 1
							While JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").GetROProperty ("enabled") = 1
									      'bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_PWC_FilterAssociationTable_Operations", "Click",  JavaWindow("Project - Teamcenter 8"), "SaveFilterSetting")
											' As Save Filter Settings is not displayed due to scroll bar added Send Keys to click on the button
											If Environment.value("TestName") = "SmartFolderConfig_Custom_prop" Then
												bGblFuncRetVal = Fn_KeyBoardOperation("SendKeys", "{TAB}~{TAB}~{TAB}~ ")
											Else
												bGblFuncRetVal = Fn_KeyBoardOperation("SendKeys", "{TAB}~{TAB}~{TAB}~{TAB}~ ")
											End If
											
									      	If bGblFuncRetVal = False Then
									      		Fn_PWC_FilterAssociationTable_Operations = False
									      	Exit Function
									      End If
									wait(3)
									exitCnt = exitCnt + 1
									if exitCnt = 10 then
										Fn_PWC_FilterAssociationTable_Operations = False
										exit function
									End If
							Wend
							
							'Write Log to Return Value.
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Failed to comple of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = True
							End If
				Case "EditFilterCriteria"
						Set objTable = JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable")
						'iRORows = iRow
						objTable.SelectCell iRow,"Source Type"
						Wait(3)
						JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable").SelectCell iRow,"Source Type"
						JavaWindow("Project - Teamcenter 8").JavaList("FilterAssociationList").Select sSourceType
						objTable.SelectCell iRow,"Property"
						Wait(3)
						JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable").SelectCell iRow,"Property"
						JavaWindow("Project - Teamcenter 8").JavaList("FilterAssociationList").Select sProperty
						JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable").SelectCell iRow,"Value"

						'Swapnil:18-JUNE-2012: 0606: Instead of for loop use Type Method.
						wait(2)
						objTable.Type sValue

						Set WshShell = Nothing
						Set objTable = Nothing 

						JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").WaitProperty "enabled", 1, 20
						If CBool(JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").GetROProperty("enabled")) = True Then
							JavaWindow("Project - Teamcenter 8").JavaButton("SaveFilterSetting").Object.click
							Call Fn_ReadyStatusSync(1)
						End If
						'Write Log to Return Value.
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Failed to comple of function Fn_PWC_FilterAssociationTable_Operations")
							Fn_PWC_FilterAssociationTable_Operations = False
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_PWC_FilterAssociationTable_Operations")
							Fn_PWC_FilterAssociationTable_Operations = True
						End If

				Case "Verify"
					Set objTable = JavaWindow("Project - Teamcenter 8").JavaTable("FilterAssociationTable")
					iRowCount = objTable.GetROProperty("rows")
					iColCount = objTable.GetROProperty("cols")
					For iRows=0 to iRowCount-1
							If Trim(Lcase(objTable.GetCellData(iRows,0))) = Trim(Lcase(sName)) Then
								If Trim(Lcase(objTable.GetCellData(iRows,1))) = Trim(Lcase(sSourceType)) Then
									If Trim(Lcase(objTable.GetCellData(iRows,2))) = Trim(Lcase(sProperty)) Then
										If Trim(Lcase(objTable.GetCellData(iRows,3))) = Trim(Lcase(sValue)) Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Data verified successfully")
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_PWC_FilterAssociationTable_Operations function passed successfully")
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_PWC_FilterAssociationTable_Operations function passed successfully")
											Fn_PWC_FilterAssociationTable_Operations = iRows												
											Exit Function
										End If										
									End If									
								End If								
							End If
					Next
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : The given Data is not present in the Table")
					Set objTable = Nothing 
					Fn_PWC_FilterAssociationTable_Operations = FALSE
				'[TC11.4(20170912a00)_NewDevelopment_PoonamC_21Sept2017:Added new Case to enter filter criteria and Save it.]
				Case "AddFilterCriteriaAndSave"
							Set objTable = JavaWindow("Project - Teamcenter 8_Old").JavaTable("FilterAssociationTable")
							If Fn_SISW_UI_Object_Operations("","Enabled", JavaWindow("Project - Teamcenter 8_Old").JavaButton("AddFilter"), SISW_MAX_TIMEOUT) = True Then
								JavaWindow("Project - Teamcenter 8_Old").JavaButton("AddFilter").Click micLeftBtn
							End If
							iRORows = objTable.GetROProperty("rows") - 1
							objTable.SelectCell iRORows,"Source Type"
							Wait(3)
							objTable.SelectCell iRORows,"Source Type"
							JavaWindow("Project - Teamcenter 8_Old").JavaList("FilterAssociationList").Select sSourceType
							objTable.SelectCell iRORows,"Property"
							Wait(3)
							objTable.SelectCell iRORows,"Property"
							JavaWindow("Project - Teamcenter 8_Old").JavaList("FilterAssociationList").Select sProperty
							
							wait(2)
							objTable.SelectCell iRORows,"Value"
							wait 1
							objTable.type sValue
							objTable.SelectCell iRORows,"Name"
							Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							Set objTable = Nothing 

							JavaWindow("Project - Teamcenter 8_Old").JavaButton("SaveFilterSetting").WaitProperty "enabled", 1, 20
							
							JavaWindow("Project - Teamcenter 8_Old").JavaButton("SaveFilterSetting").Object.setFocus()
							Wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							
							'Call Fn_Button_Click("Fn_PWC_FilterAssociationTable_Operations",JavaWindow("Project - Teamcenter 8_Old"),"SaveFilterSetting")
							Call Fn_ReadyStatusSync(2)
							Wait 1
							'Write Log to Return Value.
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Failed to complete of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = True
							End If	
				Case "CleanupExt"
						iRowCount = cInt(JavaWindow("Project - Teamcenter 8_Old").JavaTable("FilterAssociationTable").GetROProperty("rows"))
						if iRowCount > 0 then
							For iRows=0 to iRowCount-1
								JavaWindow("Project - Teamcenter 8_Old").JavaTable("FilterAssociationTable").ClickCell 0, 0
								wait 1
								Call Fn_Button_Click("Fn_PWC_FilterAssociationTable_Operations", JavaWindow("Project - Teamcenter 8_Old"), "RemoveFilter")	
							Next
							JavaWindow("Project - Teamcenter 8_Old").JavaButton("SaveFilterSetting").Object.setFocus()
							wait 1
							Call Fn_KeyBoardOperation("SendKeys", "{ENTER}")
							Call Fn_ReadyStatusSync(2)
							If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Failed to comple of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = False
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_PWC_FilterAssociationTable_Operations")
								Fn_PWC_FilterAssociationTable_Operations = True
							End If
						Else
							Fn_PWC_FilterAssociationTable_Operations = True
						End If
				Case Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_PWC_FilterAssociationTable_Operations function failed")
					Fn_PWC_FilterAssociationTable_Operations = False
	End Select
	Call Fn_SISW_UI_RACTabFolderWidget_Operation("RMBMenuSelect", "Smart Folder Filter Configuration", "Restore")
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_ProjectTabOperations(sAction,sTabName)
'###
'###    DESCRIPTION     :   Tab operation cases
'###
'###    PARAMETERS      :   sAction: Activate,VerifyActivate
'###									sTabName: Project,Smart Folder Filter Configuration
'###									
'###
'###    Return Value  	:   	True/False 
'###
'###    HISTORY         :   	AUTHOR              	DATE        		VERSION		Build
'###
'###    CREATED BY      :   Harshal				23-Sept-2010   			1.0					902
'###
'###    REVIWED BY      :	Harshal					23-Sept-2010   		1.0					902
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Msgbox Fn_ProjectTabOperations("Activate","Smart Folder Filter Configuration")
'###  								'Msgbox Fn_ProjectTabOperations("VerifyActivate","Smart Folder Filter Configuration")
'###  								'Msgbox Fn_ProjectTabOperations("Activate","Project")
'###  								'Msgbox Fn_ProjectTabOperations("VerifyActivate","Project")
'#############################################################################################
Function Fn_ProjectTabOperations(sAction,sTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_ProjectTabOperations"
Dim objTab
Set objTab = JavaWindow("Project - Teamcenter 8").JavaTab("SmartFolderTab")
Select Case sAction
Case "Activate"
	objTab.Select sTabName
	Fn_ProjectTabOperations = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"True: Activated Tab Succesfully. ")
Case "VerifyActivate"
	If trim(cstr(sTabName)) = trim(cstr(objTab.GetROProperty("value"))) Then
		Fn_ProjectTabOperations = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"True: Verified Activated Tab Succesfully. ")
	Else
		Fn_ProjectTabOperations = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"False: fail to Verify Activated Tab.")
	End If
Case Else
	Fn_ProjectTabOperations = False
	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"False:InAppropiate sAction parameter")
End Select
Set objTab = Nothing
End Function
'#######################################################################################
'###     FUNCTION NAME   :   Fn_PWC_CumulativeTable_Operations(sAction, bContribute, sName, sSourceType, sProperty, sValue, iRow, iColumn)
'###
'###    DESCRIPTION     :   Function is used to perform operations on Cumulative Table of Filter Association.
'###
'###    Return Value  	:   	True/False 
'###
'###    HISTORY         :   	AUTHOR              	DATE        		VERSION		Build 
'###
'###    CREATED BY      :   Ketan Raje				23-Sept-2010   			1.0			902
'###
'###    REVIWED BY      :	Harshal					23-Sept-2010   			1.0			902   			
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Case "Verify" : 'Msgbox Fn_PWC_CumulativeTable_Operations("Verify", "", "G10293-FrameBacklit Glass", "Item", "object_type", "Ketan", "", "")
'#############################################################################################
Public Function Fn_PWC_CumulativeTable_Operations(sAction, bContribute, sName, sSourceType, sProperty, sValue, iiRow, iiColumn)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_CumulativeTable_Operations"
	Dim iRowCount, iColCount, iRows, ObjCumTable
	Set ObjCumTable =	Fn_UI_ObjectCreate("Fn_PWC_CumulativeTable_Operations", JavaWindow("Project - Teamcenter 8").JavaTable("CumulativeTable"))
	 	If bContribute<>"" Then
				'Set the Contribute Check Box.
				Call Fn_CheckBox_Set("Fn_PWC_CumulativeTable_Operations", JavaWindow("Project - Teamcenter 8"), "Contribute", bContribute)
		End If
Select Case sAction
			Case "Verify"
						iRowCount = ObjCumTable.GetROProperty("rows")
						iColCount = ObjCumTable.GetROProperty("cols")
						For iRows=0 to iRowCount-1
							'For iCols=0 to iColCount-1
								If Trim(Lcase(ObjCumTable.GetCellData(iRows,0))) = Trim(Lcase(sName)) Then
									If Trim(Lcase(ObjCumTable.GetCellData(iRows,1))) = Trim(Lcase(sSourceType)) Then
										If Trim(Lcase(ObjCumTable.GetCellData(iRows,2))) = Trim(Lcase(sProperty)) Then
											If Trim(Lcase(ObjCumTable.GetCellData(iRows,3))) = Trim(Lcase(sValue)) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Data verified successfully")
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Fn_PWC_CumulativeTable_Operations function passed successfully")
												Fn_PWC_CumulativeTable_Operations = iRows
												Exit Function
											End If										
										End If									
									End If								
								End If
							'Next
						Next
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : The given Data is not present in the Table")
						Fn_PWC_CumulativeTable_Operations = FALSE
			Case Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fn_PWC_CumulativeTable_Operations function failed")
						Fn_PWC_CumulativeTable_Operations = FALSE						
End Select
	Set ObjCumTable = Nothing
End Function
'*********************************************************		Function to create details Item		***********************************************************************
'Function Name	        :				Fn_PWC_ItemDetailsCreate

'Description			  :		 		  Creats an Item with detail information

'Return Value		    : 				True / False  

'Pre-requisite			:		 		"MyTeamCenter" prespective should be open. 

'Examples			  :					Call Fn_PWC_ItemDetailsCreate("Item", "OFF", "None:None:Ketan:Testing:None", "", "", "", "", "OFF", "", "", "sam", "", "Finish:Close")
'Example for Design :			  Call Fn_PWC_ItemDetailsCreate("Design", "OFF", "None:None:TestDesign:Testing:None", "Design:None:DOC:None:None:000088", "", "", "", "", "", "", "", "", "Finish:Close")
'Example for Drawing :			Call Fn_PWC_ItemDetailsCreate("Drawing", "OFF", "None:None:TestDrawing:Testing:None", "aaa:123:AutoGrp1:A:xyz", "", "", "", "", "", "", "", "", "Finish:Close")
'Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Ketan Raje			01/10/2010		1.0			Created								Harshal
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Aniruddha Bhoot		7/06/2012			        
'---------------------------------------------------------------------------------------------
' Shrikant N			7/06/2012	    			Changes : Modified hierchy For (New Item)Design		        
'---------------------------------------------------------------------------------------------
' Sandeep N				27/07/2012	    			Changes : added code to click on "Enter Additional Drawing Information" link
'------------------------------------------------------------------------------------------------
' Koustubh W			28/09/2012	    			Changes : Commented code If Trim(Lcase(sSelectType))="drawing" 
'------------------------------------------------------------------------------------------------
'*****************************************************************************************************************************************************************************************************************************
Function Fn_PWC_ItemDetailsCreate(sSelectType, bConfItem, sItemInfo, sAddItemInfo, sAddItemRevInfo, sAttachFileInfo, sWorkFlowInfo, sIdentifierBasicInfo, sAddIDInfo, sAddRevInfo, sAssignProj, sDefineOptions, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_ItemDetailsCreate"
   Fn_PWC_ItemDetailsCreate = False
   on error Resume Next
	Dim ObjStaticText, objDialogNewItem, aItemInfo, sItemId, sRevId, aAddItemInfo, aProjectName, iRowData, iCount, iCounter, sOptions, aButtons,WshShell,iLen
	Dim objSelectType,bFlag,sNewItemMenu
	Dim iPreviousSelectedItem,iCurrentSelectedItem
	bFlag=False

	If Trim(Lcase(sSelectType))="design" OR Trim(Lcase(sSelectType))="subdesign" Then
		sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"FileNewDesign")
		'Select menu [File -> New -> Design...]
		Window("TeamcenterWindow").JavaDialog("New Item").SetTOProperty "title","New Design"
		If Fn_UI_ObjectExist("Fn_PWC_ItemDetailsCreate",Window("TeamcenterWindow").JavaDialog("New Item"))=False Then
			Call Fn_MenuOperation("Select",sNewItemMenu)
			Call Fn_ReadyStatusSync(2)
		End If
'	ElseIf Trim(Lcase(sSelectType))="drawing" OR Trim(Lcase(sSelectType))="subdrawing" Then
		'Select menu [File -> New -> Design...]
		'Window("TeamcenterWindow").JavaDialog("New Item").SetTOProperty "title","New Drawing"
		'If Fn_UI_ObjectExist("Fn_PWC_ItemDetailsCreate",Window("TeamcenterWindow").JavaDialog("New Item"))=False Then
		'	Call Fn_MenuOperation("Select","File:New:Drawing...")
		'	Call Fn_ReadyStatusSync(2)
		'End If
		
	ElseIf Trim(Lcase(sSelectType))="data requirement item"  OR Trim(Lcase(sSelectType))="p4_subdri" Then
		sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"FileNewDataRequirementItem")
		Window("TeamcenterWindow").JavaDialog("New Item").SetTOProperty "title","New Data Requirement Item"
		If Fn_UI_ObjectExist("Fn_PWC_ItemDetailsCreate",Window("TeamcenterWindow").JavaDialog("New Item"))=False Then
			Call Fn_MenuOperation("Select",sNewItemMenu)
			Call Fn_ReadyStatusSync(2)
		End If
	Else
		'Select menu [File -> New -> Item...]
		'If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("New Item").Exist = False Then
		sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"),"NewItem")
		If Fn_UI_ObjectExist("Fn_ItemBasicCreate",JavaWindow("DefaultWindow").JavaWindow("NewItem"))=False Then
				Call Fn_MenuOperation("Select",sNewItemMenu)
				Call Fn_ReadyStatusSync(2)
		End If
		bFlag=True
        
	End If
	If bFlag=False Then
		'Creating Object of links on the left side of the window
		Set ObjStaticText =Window("TeamcenterWindow").JavaDialog("New Item").JavaStaticText("Stpes")
		'Check the existence of "New Item " window
		Set objDialogNewItem=Window("TeamcenterWindow").JavaDialog("New Item")
	Else
		Set objDialogNewItem=JavaWindow("DefaultWindow").JavaWindow("NewItem")
	End If

	'Select Item Type
	If sSelectType <> "" Then
		Call Fn_UI_JavaList_ExtendSelect("Fn_PWC_ItemDetailsCreate", objDialogNewItem,"SelectedProject",sSelectType)
	End If
	'checked Configuration item or not
	If Trim(bConfItem) <> "" Then
	 Call Fn_CheckBox_Set("Fn_PWC_ItemDetailsCreate", objDialogNewItem,"Configuration Item",bConfItem)
	End If
	'Click on "Next" button
	 Call Fn_Button_Click("Fn_PWC_ItemDetailsCreate", objDialogNewItem,"Next")
		'Enter Item Information
		If sItemInfo<>"" Then
				aItemInfo = split(sItemInfo, ":",-1,1)
				'click on assign button
				If  aItemInfo(0) = "None" or aItemInfo(1) = "None" Then	
					Call Fn_Button_Click("Fn_PWC_ItemDetailsCreate", objDialogNewItem,"Assign")
				Else
					'Set Item ID
					Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemID", aItemInfo(0))
					If sSelectType="P4_AAUItem1" Then
						objDialogNewItem.JavaButton("UnitOfMeasure").SetTOProperty "index",1
						If objDialogNewItem.JavaButton("UnitOfMeasure").Exist Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified that Revision is a ListBox")
						End If
					End If
					'Set Revision ID
					Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"RevisionID", aItemInfo(1))
				End If				
				'Set Item name
				 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate", objDialogNewItem,"ItemName",aItemInfo(2))
				'Set description
				If aItemInfo(3)<>"None" Then
					Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate", objDialogNewItem,"Description",aItemInfo(3))
				End If
				'Set UOM
				If aItemInfo(4) <> "None" Then
				  Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate", objDialogNewItem,"Unit of Measure",aItemInfo(4))
				End If 
		End If
		'Extract Creation data
		sItemId = objDialogNewItem.JavaEdit("ItemID").GetROProperty("value")
		sRevId = objDialogNewItem.JavaEdit("RevisionID").GetROProperty("value")
		'Entering Additional Item Information
			If sAddItemInfo<>"" Then				
				' Click on Next Button
				If Trim(Lcase(sSelectType))="design" OR Trim(Lcase(sSelectType))="subdesign" Then
					ObjStaticText.SetTOProperty "label", "Enter Additional Design Information"
					ObjStaticText.Click 1, 1
				ElseIf Trim(Lcase(sSelectType))="drawing" OR Trim(Lcase(sSelectType))="subdrawing" Then
					ObjStaticText.SetTOProperty "label", "Enter Additional Drawing Information"
					ObjStaticText.Click 1, 1
	                  ElseIf Trim(Lcase(sSelectType))="data requirement item" OR Trim(Lcase(sSelectType))="p4_subdri" Then
                    ObjStaticText.SetTOProperty "label", "Enter Additional Data Requirement Item Information"
					ObjStaticText.Click 1, 1
				Else
					ObjStaticText.SetTOProperty "label", "Enter Additional Item Information"
					ObjStaticText.Click 1, 1
				End If
				aAddItemInfo = split(sAddItemInfo, ":",-1,1)	
                    	If sSelectType="Contract" OR sSelectType="P4_SubContract"  Then
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("New Item").JavaEdit("Contract Category").Activate
							Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"Contract Category", aAddItemInfo(0))
						ElseIf sSelectType = "ADS Tec Document" OR sSelectType ="Technical Document" OR sSelectType = "SubTechDoc" Then
							If	aAddItemInfo(0)<>"None" Then
								 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"Category", aAddItemInfo(0))
							End If
							If	aAddItemInfo(1)<>"None" Then
								 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"Technical Document Category", aAddItemInfo(1))
							End If
						ElseIf Trim(Lcase(sSelectType))="design" OR Trim(Lcase(sSelectType))="subdesign" Then
							If	aAddItemInfo(0)<>"None" Then
								 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"Design Category", aAddItemInfo(0))
							End If
							If	aAddItemInfo(1)<>"None" Then
								 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourceDocCategory", aAddItemInfo(1))
							End If
							If	aAddItemInfo(2)<>"None" Then
								 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourceDocID", aAddItemInfo(2))
							End If
							If aAddItemInfo(3)<>"None" Then
								Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourceDocOrgNT", aAddItemInfo(3))
							End If
							If aAddItemInfo(4)<>"None" Then
								Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourcDocRev", aAddItemInfo(4))
							End If
							If aAddItemInfo(5)<>"None" Then
								Call Fn_UI_Object_SetTOProperty("Fn_PWC_ItemDetailsCreate",objDialogNewItem.JavaEdit("SourcTecDocCategoryNT"),"Index","5")
								Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourcTecDocCategoryNT", aAddItemInfo(5))
							End If

						ElseIf sSelectType = "Data Requirement Item" OR sSelectType = "P4_SubDRI" Then
							If aAddItemInfo(0)<>"None" Then
									Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"Contract Line Item Number", aAddItemInfo(0))
							End If
							If aAddItemInfo(1)<>"None" Then
									Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"Contract Reference", aAddItemInfo(1))
							End If
					ElseIf sSelectType = "Data Item Description" OR sSelectType = "P4_SubDID" Then
							If aAddItemInfo(0)<>"None" Then
									Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"DID Type", aAddItemInfo(0))
							End If
							If aAddItemInfo(1)<>"None" Then
									Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"Program Phases", aAddItemInfo(1))
							End If
					ElseIf sSelectType = "Standard Note" Then
							If aAddItemInfo(0)<>"None" Then
									Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"Note Category", aAddItemInfo(0))
							End If
					ElseIf sSelectType = "Drawing" Then
							If aAddItemInfo(0)<>"None" Then
								Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourceDocCategory", aAddItemInfo(0))
							End If
							If aAddItemInfo(1)<>"None" Then
								Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourceDocID", aAddItemInfo(1))
							End If
							If aAddItemInfo(2)<>"None" Then
								Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourceDocOrgNT", aAddItemInfo(2))
							End If
							If aAddItemInfo(3)<>"None" Then
								Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourcDocRev", aAddItemInfo(3))
							End If
							If aAddItemInfo(4)<>"None" Then
								Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"SourcTecDocCategoryNT", aAddItemInfo(4))
							End If
					Else
						If sAddItemInfo(0) <>"None" Then
							'Code need to be updated
						End If
						If sAddItemInfo(1) <>"None" Then
							 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemPreviousID", aAddItemInfo(1))
						End If
						If sAddItemInfo(2) <>"None" Then
							 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemSerialNumber", aAddItemInfo(2))
						End If
						If sAddItemInfo(3) <>"None" Then
							 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemComment", aAddItemInfo(3))
						End If
						If sAddItemInfo(4) <>"None" Then
							 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData1", aAddItemInfo(4))
						End If
						If sAddItemInfo(5) <>"None" Then
							 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData2", aAddItemInfo(5))
						End If
						If sAddItemInfo(6) <>"None" Then
							 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData3", aAddItemInfo(6))
						End If
						'This Edit Box is added only for Requirement Specification
						If sAddItemInfo(7) <>"None" Then
							 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ReqSpecSubject", aAddItemInfo(7))
						End If
				End If
			End If	
			'Enter Additional Item Revision Information
			If sAddItemRevInfo<>"" Then
				' Click on Next Button		
				If Trim(Lcase(sSelectType))="design" OR Trim(Lcase(sSelectType))="subdesign" Then
					ObjStaticText.SetTOProperty "label", "Enter Additional Design Revision Information"
					ObjStaticText.Click 1, 1
				Else
					ObjStaticText.SetTOProperty "label", "Enter Additional Item Revision Information"
					ObjStaticText.Click 1, 1
				End If
				aAddItemRevInfo = split(sAddItemRevInfo, ":",-1,1)	
					If aAddItemRevInfo(0) <>"None" Then
						If sSelectType="Correspondence"  OR sSelectType="P4_SubCorresp" Then
							Window("TeamcenterWindow").JavaDialog("New Item").JavaEdit("Category").Activate
							Window("TeamcenterWindow").JavaDialog("New Item").JavaEdit("Category").Set aAddItemRevInfo(0)						
						Else
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemProjectID", aAddItemRevInfo(0))
						End if
					End If
					If aAddItemRevInfo(1) <>"None" Then
						If sSelectType="Correspondence" Then
							Window("TeamcenterWindow").JavaDialog("New Item").JavaEdit("Correspondence Direction").Activate
							Window("TeamcenterWindow").JavaDialog("New Item").JavaEdit("Correspondence Direction").Set aAddItemRevInfo(0)													
						Else
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemRevPreviousID", aAddItemRevInfo(1))
						End if
					End If
					If aAddItemRevInfo(2) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemSerialNumber", aAddItemRevInfo(2))
					End If
					If aAddItemRevInfo(3) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemComment", aAddItemRevInfo(3))
					End If
					If aAddItemRevInfo(4) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData1", aAddItemRevInfo(4))
					End If
					If aAddItemRevInfo(5) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData2", aAddItemRevInfo(5))
					End If
					If aAddItemRevInfo(6) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData3", aAddItemRevInfo(6))
					End If
			End If
			'Enter Identifier Basic Information
			If sIdentifierBasicInfo<>"" Then
				' Click on Next Button		
				ObjStaticText.SetTOProperty "label", "Enter Identifier Basic Information"
				ObjStaticText.Click 1, 1
				Wait(2)
				If sIdentifierBasicInfo(0)<>"None" Then
					'Set TOProperties of Dialog Box
					JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("No Assign Privilege").SetTOProperty "title","New Item ..."
					'Set the "Don't show this message" Status
					Call Fn_CheckBox_Set("Fn_PWC_ItemDetailsCreate", JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("No Assign Privilege"), "Don't show this message", sIdentifierBasicInfo(0))
					'Click on OK button
					Call Fn_Button_Click("Fn_PWC_ItemDetailsCreate", JavaWindow("DefaultWindow").JavaWindow("Shell").JavaWindow("No Assign Privilege"), "OK")
				End If
					If sIdentifierBasicInfo(1) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemProjectID", sIdentifierBasicInfo(1))
					End If
					If sIdentifierBasicInfo(2) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemRevPreviousID", sIdentifierBasicInfo(2))
					End If
					If sIdentifierBasicInfo(3) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemSerialNumber", sIdentifierBasicInfo(3))
					End If
					If sIdentifierBasicInfo(4) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemComment", sIdentifierBasicInfo(4))
					End If
					If sIdentifierBasicInfo(5) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData1", sIdentifierBasicInfo(5))
					End If
					If sIdentifierBasicInfo(6) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData2", sIdentifierBasicInfo(6))
					End If
					If sIdentifierBasicInfo(7) <>"None" Then
						 Call Fn_Edit_Box("Fn_PWC_ItemDetailsCreate",objDialogNewItem,"ItemUserData3", sIdentifierBasicInfo(7))
					End If
			End If
			'Assign to Project
			If sAssignProj<>"" Then
				' Click on Next Button
				ObjStaticText.SetTOProperty "label", "Assign to Project"
				ObjStaticText.Click 1, 1
				Call Fn_ReadyStatusSync(3)
				bReturn = objDialogNewItem.JavaList("ProjectForSelect").GetROProperty("items count")
				'Extract the index of row at which the object exist.
				aProjectName = split(sAssignProj, ":",-1,1)
				iCount = Ubound(aProjectName)
				For iRowData=0 to iCount
					For iCounter=0 to bReturn-1
						If Trim(lcase(objDialogNewItem.JavaList("ProjectForSelect").GetItem(iCounter))) = Trim(lcase(aProjectName(iRowData))) then
							objDialogNewItem.JavaList("ProjectForSelect").Select aProjectName(iRowData)
							'Click on Remove Button
							Call Fn_Button_Click("Fn_PWC_ItemDetailsCreate", objDialogNewItem, "AddProject")											
							Exit For 
						End If
					Next
				Next
			End If
			If sDefineOptions<>"" Then
				' Click on Next Button
					ObjStaticText.SetTOProperty "label", "Define Options"
					ObjStaticText.Click 1, 1	
						sOptions = split(sDefineOptions, ":",-1,1)					
							If sOptions(0) <> "" Then
								Call Fn_CheckBox_Set("Fn_PWC_ItemDetailsCreate" ,objDialogNewItem,"ShowAsNwRt", sOptions(0)) 
							End If
							If sOptions(1) <> "" Then
								Call Fn_CheckBox_Set("Fn_PWC_ItemDetailsCreate" ,objDialogNewItem,"UsItIdentifierAs", sOptions(1)) 
							End If
							If sOptions(2) <> "" Then
								Call Fn_CheckBox_Set("Fn_PWC_ItemDetailsCreate" ,objDialogNewItem,"UsRevIdentifier", sOptions(2)) 
							End If
							If sOptions(3) <> "" Then
								Call Fn_CheckBox_Set("Fn_PWC_ItemDetailsCreate" ,objDialogNewItem,"ChkOutItmRevOnCr", sOptions(3)) 
							End If
			End If
			objDialogNewItem.JavaButton("Next").WaitProperty "enabled", 1, 20000
			'Click on Buttons
			If sButtons<>"" Then
					aButtons = split(sButtons, ":",-1,1)
					iCounter = Ubound(aButtons)
					For iCount=0 to iCounter
						'Click on Add Button
						Call Fn_Button_Click("Fn_PWC_ItemDetailsCreate", objDialogNewItem, aButtons(iCount))
'						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("New Item").Activate
'						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("New Item").JavaButton("Finish").Click micLeftBtn						
						Wait(5)
						Call Fn_ReadyStatusSync(3)
						If Trim(Lcase(aButtons(iCount))) = "finish" OR Trim(Lcase(aButtons(iCount))) = "close" Then
							If objDialogNewItem.Exist Then
								If objDialogNewItem.JavaButton(aButtons(iCount)).GetROProperty("enabled") = 1 Then
									Call Fn_Button_Click("Fn_PWC_ItemDetailsCreate", objDialogNewItem, aButtons(iCount))  
								End If
							End If
						End If
					Next
			End If
			'Function Returns Item ID and True
			Fn_PWC_ItemDetailsCreate = sItemId & "-" & sRevId
			Window("TeamcenterWindow").JavaDialog("New Item").SetTOProperty "title","New Item"
			Call Fn_ReadyStatusSync(1)
			'Write Log
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Completed the function Fn_PWC_ItemDetailsCreate")
 Set ObjStaticText = Nothing
 Set objDialogNewItem = Nothing
 End Function
''*********************************************************		Function to perform action on Project Tree	***********************************************************************
'Function Name		:				Fn_PWC_SelectedMemberNodePath()
'Description			 :		 		 Returns the path for the Node selection in Selected member tree.
'Parameters			   :	 			sUserToSelect[Envoirnment variable from EnvVar_Ext.xml  else in Format "AutoTest1:AutoTest1:Engineering:Designer::autotest1"]
'Return Value		   : 				Returns the path for the Node selection in Selected member tree.
'Pre-requisite			:		 		Project Prespective is Open.
'Examples				:				Call Fn_PWC_SelectedMemberNodePath(Environment.Value("TcUser1"))
'History					 :		
'													Developer Name			Date						Rev. No.			Build
'												-----------------------------------------------------------------------------------------------------------------
'													Harshal Agrawal			07 Sept 2011			1.0						Tc91(2011082400)
'												-----------------------------------------------------------------------------------------------------------------
'*******************************************************************************************************************************************************************************************
Function Fn_PWC_SelectedMemberNodePath(sUserToSelect)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_SelectedMemberNodePath"
   	Dim objSelectedMemberTree,aUserToSelect
	ReDim aUserToSelect(6)
	 Set objSelectedMemberTree = Fn_UI_ObjectCreate("Fn_PWC_MemberSelection", JavaWindow("Project - Teamcenter 8").JavaTree("SelMemTree"))
		aUserToSelect = split(sUserToSelect, ":", -1, 1)
		Fn_PWC_SelectedMemberNodePath = objSelectedMemberTree.GetItem(0)+":"+aUserToSelect(3)+":"+aUserToSelect(2)+"/"+aUserToSelect(3)+"/"+aUserToSelect(0)+" ("+aUserToSelect(5)+")"
End Function

''*********************************************************		Function to perform action on Library Tree	***********************************************************************
'Function Name		:				Fn_PWC_LibraryTreeOpeartion()
'Description			 :		 		 For library tree in Project
'Parameters			   :	 			sAction,sNodeName, sMenu
'Return Value		   : 				true/false.
'Pre-requisite			:		 		Project Prespective is Open.
'Examples				:				Fn_PWC_LibraryTreeOpeartion("Select","Classificaion Root", "")
'													bReturn=Fn_PWC_LibraryTreeOpeartion("AddNode","Classification Root:MyGroup_1010:MyLibs  [5]:MySignals  [3]:Signals1  [2]", "")
'History					 :		
'													Developer Name			Date						Rev. No.			Build
'												-----------------------------------------------------------------------------------------------------------------
'													Prasanna B.			23 Jan 2012			1.0						Tc91(2011011100)
'													Shreyas					03-02-2012			1.1						Tc91(2011011100)						
'												-----------------------------------------------------------------------------------------------------------------
'*******************************************************************************************************************************************************************************************

Public Function Fn_PWC_LibraryTreeOpeartion(sAction,sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_LibraryTreeOpeartion"
	Dim objLibTree,iRowCounter,iItemCount,objApplet,iCounter,sTreeItem,bReturn,aNodeName

	Set objLibTree = JavaWindow("Project - Teamcenter 8")
'		Set objLibTree = JavaWindow("Project - Teamcenter 8").JavaTree("LibraryTree")
'	Set objApplet = JavaWindow("Project - Teamcenter 8")

	Select Case sAction
  			   Case "Select"
									objLibTree.JavaTree("LibraryTree").Activate sNodeName
									If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_PWCRules_TreeOpeartion ] Failed case [ " & sAction & " ] Specified node ["+sNodeName+"] can not be selected.")
											Set objLibTree = nothing
'											Set objApplet = nothing
											Fn_PWC_LibraryTreeOpeartion = false
											Exit function
									Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_PWCRules_TreeOpeartion ] Successfully selected node ["+sNodeName+"].")
											Fn_PWC_LibraryTreeOpeartion = true
									End If
				Case "Expnad"        					
									iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_PWCRules_TreeOpeartion", objLibTree , "LibraryTree", sNodeName, "", "")
									If iRowCounter = -1 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_PWCRules_TreeOpeartion ] Failed case [ " & sAction & " ] Specified node ["+sNodeName+"] can not be expanded.")
											Set objLibTree = nothing
'											Set objApplet = nothing
											Fn_PWC_LibraryTreeOpeartion = false
											Exit function							
									End If
									objLibTree.JavaTree("LibraryTree").Object.setSelectionRow iRowCounter
									objLibTree.JavaTree("LibraryTree").Object.setExpandedState objLibTree.JavaTree("LibraryTree").Object.getSelectionPath(), true					
									 If Err.Number <  0  Then
											Fn_PWC_LibraryTreeOpeartion = false
'											Set objApplet = nothing
											Exit function
									End If
									
									Fn_PWC_LibraryTreeOpeartion = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_PWCRules_TreeOpeartion ] Successfully expanded node ["+sNodeName+"].")
				Case "Exist"
									iItemCount = objLibTree.JavaTree("LibraryTree").GetROProperty( "items count")
	
									For iCounter=0 To (iItemCount-1)
										sTreeItem = objLibTree.JavaTree("LibraryTree").GetItem(iCounter)
										If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
											Fn_PWC_LibraryTreeOpeartion = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully found node " + sNodeName + "of Library Tree." )	
											Exit For
										End If
									Next 
	
									If  Cint(iCounter) = Cint (iItemCount) Then
										Fn_PWC_LibraryTreeOpeartion = false
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  found node [" + sNodeName + " ] of Library Tree." )	
										Set objLibTree = nothing
										Exit Function 
									End If
					Case "PopupMenuSelect"
									iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_PWC_LibraryTreeOpeartion", objLibTree , "LibraryTree", sNodeName, "", "")
									If iRowCounter = -1 Then
										'If Node name contains numeric index for any child nodes
										objLibTree.JavaTree("LibraryTree").Select(sNodeName) 
									Else
										objLibTree.JavaTree("LibraryTree").Object.setSelectionRow iRowCounter
									End If
					
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_PWC_LibraryTreeOpeartion ] Failed case [ " & sAction & " ] for node [" + sNodeName + " ] in Library tree.")
'										Set objApplet = nothing
										Fn_PWC_LibraryTreeOpeartion = false
										Exit function
									End If


									 JavaWindow("Project - Teamcenter 8").JavaTree("LibraryTree").DblClick 0,0,"LEFT"
									wait 1
									JavaWindow("Project - Teamcenter 8").JavaTree("LibraryTree").DblClick 0,0,"LEFT"
									wait 2

									objLibTree.JavaTree("LibraryTree").OpenContextMenu(sNodeName)
									wait 2
									objLibTree.JavaMenu("RMBMenuSelect").SetTOProperty "label", sMenu
									If not objLibTree.JavaMenu("RMBMenuSelect").Exist(3) Then
										objLibTree.JavaTree("LibraryTree").Select(sNodeName) 
										wait 1
										Call Fn_UI_JavaTree_OpenContextMenu("Fn_PWC_LibraryTreeOpeartion",objLibTree,"LibraryTree",sNodeName)
										wait 1
									End If
									objLibTree.JavaMenu("RMBMenuSelect").Select
									If Err.Number <  0  Then
										Fn_PWC_LibraryTreeOpeartion = false
'										Set objApplet = nothing
										Exit function
									End If
									Fn_PWC_LibraryTreeOpeartion = True

				Case "AddNode"

					'Expand all The Nodes of the Liabrary Tree
					bReturn=Fn_PWC_LibraryTreeOpeartion("PopupMenuSelect","Classification Root", "ExpandAll")
					If bReturn=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_PWC_LibraryTreeOpeartion ] Failed case [ " & sAction & " ] for node [" + sNodeName + " ] in Library tree.")
						Fn_PWC_LibraryTreeOpeartion = false
						Exit function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_PWC_LibraryTreeOpeartion ] Passed case [ " & sAction & " ] for node [" + sNodeName + " ] in Library tree By Expanding All Nodes")
						Fn_PWC_LibraryTreeOpeartion = true
					End If
					Call Fn_ReadyStatusSync(3)
					'Select the Desired Node
					bReturn=Fn_PWC_LibraryTreeOpeartion("PopupMenuSelect",sNodeName, "Select")
					If bReturn=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_PWC_LibraryTreeOpeartion ] Failed case [ " & sAction & " ] for node selection [" + sNodeName + " ] in Library tree.")
						Fn_PWC_LibraryTreeOpeartion = false
						Exit function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_PWC_LibraryTreeOpeartion ] Passed case [ " & sAction & " ] for node selection [" + sNodeName + " ] in Library")
						Fn_PWC_LibraryTreeOpeartion = true
					End If
					Call Fn_ReadyStatusSync(1)
					'Click on Add Button
					bReturn= Fn_Button_Click("Fn_PWC_LibraryTreeOpeartion", objLibTree, "AddClass") 
					If bReturn=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_PWC_LibraryTreeOpeartion ] Failed case [ " & sAction & " ] for Clicking Button [Add]")
						Fn_PWC_LibraryTreeOpeartion = false
						Exit function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_PWC_LibraryTreeOpeartion ] Passed case [ " & sAction & " ] for Clicking Button [Add]")
						Fn_PWC_LibraryTreeOpeartion = true
					End If

					'Modified code to handle the Duplicate Node dialog										------------------------BY- Anjali M[30 Jul 2012]
                    JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").SetTOProperty "Index",1
					If JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").JavaDialog("DuplicateNode").Exist(5) Then 
							JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").JavaDialog("DuplicateNode").JavaButton("OK").Click micLeftBtn
							JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").RefreshObject
							Fn_PWC_LibraryTreeOpeartion = false
                            Exit function
					End If

					'Unset the index of Shell window
					JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").RefreshObject

					'Click on Save Button
					bReturn= Fn_Button_Click("Fn_PWC_LibraryTreeOpeartion", objLibTree, "Save") 
					If bReturn=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [ Fn_PWC_LibraryTreeOpeartion ] Failed case [ " & sAction & " ] for Clicking Button [Save]")
						Fn_PWC_LibraryTreeOpeartion = false
						Exit function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [ Fn_PWC_LibraryTreeOpeartion ] Passed case [ " & sAction & " ] for Clicking Button [Save]")
						Fn_PWC_LibraryTreeOpeartion = true
					End If

				Case "RemoveNode"

						If instr(1,sNodeName,"~") Then												    
									aNodeName = split(sNodeName,"~",-1,1)
						Else
									aNodeName = Array(sNodeName)
						End If

						For  iCounter=0 to uBound(aNodeName)
								objLibTree.JavaList("LibraryList").ExtendSelect(aNodeName(iCounter))
								Call Fn_ReadyStatusSync(1)
								If Err.number<0 Then
										Fn_PWC_LibraryTreeOpeartion=False
										Exit function
								End If
						Next

						'Click on Remove Button
						bReturn= Fn_Button_Click("Fn_PWC_LibraryTreeOpeartion", objLibTree, "RemoveClass") 
						Call Fn_ReadyStatusSync(1)
						If Err.number<0 Then
								Fn_PWC_LibraryTreeOpeartion=False
								Exit function
						End If

						'Click on Save
						bReturn= Fn_Button_Click("Fn_PWC_LibraryTreeOpeartion", objLibTree, "Save")
						Call Fn_ReadyStatusSync(1)
						If Err.number<0 Then
								Fn_PWC_LibraryTreeOpeartion=False
								Exit function
						Else
								Fn_PWC_LibraryTreeOpeartion=True
								Call Fn_ReadyStatusSync(1)
						End If

				Case "SetFilterCriteria"     ' Added by Pooja S:  14-Feb-2012

						'Click on drop Down button
						'objApplet.JavaObject("JLabelDropDown").Click 4,6,"LEFT"
						objLibTree.JavaObject("JLabelDropDown").Click 4,6,"LEFT"
						If err.number<0 Then
									Fn_SE_DataDictionarySearchDialogOperations=False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on [JLabelDropDown] button")
									Exit function
						End If
						'Select from Java Menu "Filter Classification Hierarchy"
						objLibTree.JavaMenu("MenuSelect").SetTOProperty "label", "Filter Classification Hierarchy"
						objLibTree.JavaMenu("MenuSelect").Select
						wait(3)
						'Select  "Select All"  from the 'SetFilterCriteria'  Dialog 
						JavaWindow("Project - Teamcenter 8").JavaWindow("Search_FilterCriteria").JavaDialog("SetFilterCriteria").JavaList("SelectLibraryTypes").Select sMenu
						wait(3)
						If err.number<0 Then
									Fn_PWC_LibraryTreeOpeartion=False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Select [ Select All ] from [ SetFilterCriteria]  Dialog")
									Exit function
						End If
						'Click on 'OK' button
						JavaWindow("Project - Teamcenter 8").JavaWindow("Search_FilterCriteria").JavaDialog("SetFilterCriteria").JavaButton("OK").Click
						wait(3)
						If err.number<0 Then
									Fn_PWC_LibraryTreeOpeartion=False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Fail to Click on [ OK ] from [ SetFilterCriteria]  Dialog")
									Exit function
						Else
								Fn_PWC_LibraryTreeOpeartion=True
								Call Fn_ReadyStatusSync(1)
						End If

    		 End Select
End Function

'*********************************************************		Function to verify  information error message while Creating New Change.	***********************************************************************
'Function Name		:					Fn_PWC_VerifyInformationMessage

'Description			 :		 		  This function is used to verify  information error message while Creating New Change

'Parameters			   :	 			1.  sTitle:Title of dialog.
'													2. sMsg : Message to verify. (Optional)
'													3. sButton : Button Name.
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		

'Examples				:			  Msgbox Fn_PWC_VerifyInformationMessage("Error","does not have any group member for the given group","OK") 


'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ganesh B				21-May-2014			1.0					CFraeted New Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PWC_VerifyInformationMessage(sTitle,sMsg,sButton) 

    Dim dicErrorInfo
	 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	 dicErrorInfo.Add "Action", "VerifyInformationMessage"
	 dicErrorInfo.Add "Title", sTitle
	 dicErrorInfo.Add "Message", sMsg
	 dicErrorInfo.Add "Button", sButton    
	 Fn_PWC_VerifyInformationMessage = Fn_SISW_PWC_ErrorVerify(dicErrorInfo)
	 Set dicErrorInfo = Nothing

End Function
'*********************************************************	Generic function to handle Error dialogs in PSE Module  	***********************************************************************
'Function Name		:		Fn_SISW_PWC_ErrorVerify()

'Description		:	The function is generic function to handle error dialogs. It is created after combining error dialog functions from ProjectWorkContext.vbs
'							Fn_PWC_DialogMsgVerify

'Parameters			 :	 			1.  dicErrorInfo
											
'Return Value		 : 				True/False

'Pre-requisite		 :		 		NA.

'Examples			 :		Dim dicErrorInfo
'										 Set dicErrorInfo = CreateObject("Scripting.Dictionary")
'										 dicErrorInfo.Add "Action", "DeleteWCMessageVerify"
'										 dicErrorInfo.Add "Message", sMessage	    
'										 bReturn = Fn_SISW_PWC_ErrorVerify(dicErrorInfo)										 
'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sushma Pagare          4-Jul-2013
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ganesh B         	21-May-2014			1.1				added Case "VerifyInformationMessage"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
public Function Fn_SISW_PWC_ErrorVerify(dicErrorInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PWC_ErrorVerify"
	Dim dicKeys, dicItems, iCounter, bReturn
	Dim sAction, sTitle, sErrorMsg,sButton, sAppMsg
	Dim descDialog, descButton, descChild, objChild
    Dim objWrkCtxt
	Dim objInformation
	On Error Resume Next
			
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
	
	Select Case sAction

          	''  This covers Fn_PWC_DialogMsgVerify(sTitle,sMsg,sButton) 
		Case "VerifyUsingDescription"
					
					Fn_SISW_PWC_ErrorVerify = True
					' Create Object Description of  Dialog 
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
					If Dialog(descDialog).Exist(5) Then
		
							'Capture All runtime objects to find message text
							Set  objChild = Dialog(descDialog).ChildObjects(descChild)
							'Set message text to variable 
							sAppMsg = objChild(1). getroproperty("text")
							'compare run time message to verify  the error message
							If sErrorMsg <> "" Then						
								If Instr(1,sErrorMsg,sAppMsg) <> 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")
								Else
									GBL_ACTUAL_MESSAGE=sAppMsg
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
									Fn_SISW_PWC_ErrorVerify = False
								End If
							End If
						' To Click "OK" Button after verification
							wait(2)
							Dialog(descDialog).WinButton(descButton).Click 10,10,micLeftBtn
							If Dialog(descDialog).Exist(5) Then
								Dialog(descDialog).Close()
							End If
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The " + sTitle + " Dialog does not Exist")
							Fn_SISW_PWC_ErrorVerify = False
						End If
		
					Set descDialog=nothing
					Set descButton=nothing
					Set descChild=nothing
					Set objChild=nothing
					Exit Function
		Case "VerifyInformationMessage"
			Set objInformation = JavaWindow("DefaultWindow").JavaWindow("New Change").JavaWindow("Information")
			objInformation.SetTOProperty "title", sTitle
			If objInformation.exist(1) Then
				 sAppMsg = objInformation.JavaEdit("Details").GetROProperty("value")
				If sErrorMsg <> "" Then						
					If StrComp(trim(sErrorMsg), trim(sAppMsg)) = 0 Then
						Fn_SISW_PWC_ErrorVerify = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")
					Else
						GBL_ACTUAL_MESSAGE=sAppMsg
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
						Fn_SISW_PWC_ErrorVerify = False
					End If
				End If
				objInformation.JavaButton("OK").SetTOProperty "label", sButton
				objInformation.JavaButton("OK").Click	
				Set objInformation = Nothing
			End If
	End Select

End Function

'*********************************************************		Function to change ownership	***********************************************************************

'Function Name		:					Fn_PWC_ChangeOwnership

'Description			 :		 		  This function is used to change the ownership

'Parameters			   :	 			1. sAction : Action need to perform.
'                                       2. sNodeName : Node need to select from tree node.
'										3. sNewOwningUser : New owning user. ( Full colon separated path)
											
'Return Value		   : 				 True/False

'Pre-requisite			:		 		 Project pane should be open.

'Examples				:				 Fn_PWC_ChangeOwnership("ChangeOwner","Project:Project_name",GroupName:Role:UserName (username))

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vivek Ahirrao        20-April-2015
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PWC_ChangeOwnership(sAction,sNodeName,sNewOwningUser)
	GBL_FAILED_FUNCTION_NAME="Fn_PWC_ChangeOwnership"
	On Error Resume Next
	'Decalre Variables 
	Dim bReturn,aNewOwner,iCount,sNode
	Fn_PWC_ChangeOwnership = False
	'Set Object for Change Ownership & Organisation selection dialog
	Set objChangeOwner = JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").JavaDialog("ChangeOwnership")
	Set objOrgSelection = JavaWindow("Project - Teamcenter 8").JavaWindow("Shell").JavaDialog("OrganizationSelection")
	'Select tree node name
	If Not objChangeOwner.Exist(5) Then
		If  sNodeName <> "" Then
			bReturn = Fn_PWCProject_TreeOpeartion("Select",sNodeName,"")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to select node " + sNodeName)
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully select node " + sNodeName)
			End If
		End If
	'Open Change Ownership dialog
		Call Fn_MenuOperation("Select","Edit:Change Ownership...")
		Call Fn_ReadyStatusSync(2)
	End If
	
	If objChangeOwner.Exist(5) Then
		Select Case sAction
	'Case for change owner user
			Case "ChangeOwner"
				Call Fn_Button_Click("Fn_PWC_ChangeOwnership", objChangeOwner,"OwingUser")
				Call Fn_ReadyStatusSync(1)
				If objOrgSelection.Exist(5) Then
					aNewOwner = Split(sNewOwningUser,":",-1,1)
					If Ubound(aNewOwner) > 0 Then
						sNode = aNewOwner(0)
						For iCount =1 to (Ubound(aNewOwner)-1)
							sNode = sNode + ":" +  aNewOwner(iCount)
							objOrgSelection.JavaTree("NewOwningTree").Expand sNode
							Wait(2)
						Next
					End If      
					Wait(3)
	'Implemented UI call to Select Tree Node
					Call Fn_JavaTree_Select("Fn_PWC_ChangeOwnership", objOrgSelection , "NewOwningTree",sNewOwningUser )
					Call Fn_ReadyStatusSync(3)
					Call Fn_Button_Click("Fn_PWC_ChangeOwnership", objOrgSelection,"OK")
					Call Fn_ReadyStatusSync(3)		   
	'Click on Yes button
					bReturn = Fn_Button_Click("Fn_PWC_ChangeOwnership", objChangeOwner,"Yes")
					If bReturn = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to click on 'Yes' button.")
						Fn_PWC_ChangeOwnership = False
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully clicked on 'Yes' button.")
						Fn_PWC_ChangeOwnership = True
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Organization Selection dialog does not exist..")
					Exit Function
				End If
	'Invalid case		
			Case Else
				Fn_PWC_ChangeOwnership = False
				Exit Function
		End Select
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Change ownership dialog does not exist.")
		Fn_PWC_ChangeOwnership = False
		Exit Function
	End If
	
	Set objChangeOwner = nothing
	Set objOrgSelection = nothing
End Function
