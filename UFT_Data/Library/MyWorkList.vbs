Option Explicit
Private objItmBounds
iTimeOut = 120
'--------------------------------------------------------------------------------------
'Global variables for myworklist
'--------------------------------------------------------------------------------------
Public GBL_MW_TN_TASKSTOPERFORM,GBL_MW_TN_PERFORMSIGNOFFS,GBL_MW_TN_SELECTSIGNOFFTEAM,GBL_MW_PARTICIPANT_PROFILES, GBL_MW_TN_MYWORKLIST, GBL_MW_TN_INBOX
Public GBL_MW_TN_NewDoTask, GBL_MW_TN_DoTask, GBL_MW_TN_COMPONENTPROCESS,GBL_MW_TN_IndependentSubWorkflows,GBL_MW_TN_ReviseImpacted,GBL_MW_NEWREVIEWTASK
Public GBL_MW_PARTICIPANT_SIGNOFF_TEAM,GBL_MW_TN_TARGET,GBL_MW_PARTICIPANT_PROPOSEDRESPONSIBLEPARTY,GBL_MW_PARTICIPANT_USERS, GBL_MW_TN_TASKSTOTRACK
Public GBL_MW_PT_PERFORMSIGNOFFS, GBL_MW_PT_SELECTSIGNOFFTEAM,GBL_MW_NEWROUTETASK,GBL_MW_PT_NEWDOTASK,GBL_MW_PT_NEWCONDITIONTASK,GBL_MW_PT_ACKNOWLEDGETASK,GBL_MW_PT_NOTIFYTASK
Public GBL_MW_PT_REVIEWTASK
GBL_MW_TN_MYWORKLIST="My Worklist"
GBL_MW_TN_INBOX=" Inbox"
GBL_MW_TN_TASKSTOPERFORM="Tasks To Perform"
GBL_MW_TN_TASKSTOTRACK = "Tasks To Track"
GBL_MW_TN_NewDoTask=" (New Do Task 1)"
GBL_MW_TN_TARGET ="Targets"
GBL_MW_TN_DoTask="(Do task)"
GBL_MW_TN_COMPONENTPROCESS = "Component Process"
GBL_MW_TN_IndependentSubWorkflows = "Independent Sub-Workflows"
GBL_MW_TN_ReviseImpacted = "(Revise Impacted)"
GBL_MW_NEWREVIEWTASK  = "New Review Task "
GBL_MW_PARTICIPANT_PROPOSEDRESPONSIBLEPARTY = "Proposed Responsible Party"
GBL_MW_PARTICIPANT_PROFILES = "Profiles"
GBL_MW_PARTICIPANT_SIGNOFF_TEAM = "Signoff Team"
GBL_MW_PARTICIPANT_USERS = "Users"
GBL_MW_TN_PERFORMSIGNOFFS=" (perform-signoffs)"
GBL_MW_PT_PERFORMSIGNOFFS = "perform-signoffs" 'Process tree(PT) Node "perform-signoffs"   
GBL_MW_TN_SELECTSIGNOFFTEAM=" (select-signoff-team)"       
GBL_MW_PT_SELECTSIGNOFFTEAM = "select-signoff-team " 'Process Tree (PT) Node "select-signoff-team "    
GBL_MW_PT_NEWDOTASK = "New Do Task " 
GBL_MW_PT_NEWCONDITIONTASK = "New Condition Task "
GBL_MW_PT_REVIEWTASK = "Review Task"
GBL_MW_PT_ACKNOWLEDGETASK="Acknowledge Task"
GBL_MW_PT_NOTIFYTASK="Notify Task"
GBL_MW_NEWROUTETASK ="New Route Task "

'*********************************************************	Function List		***********************************************************************
'Fn_SISW_MyWorkList_GetObject()
'1. Fn_MyWorkList_TreeNodeOperations()
'2. Fn_MyWorkList_SignoffTeam_TreeNodeOperations()
'3. Fn_MyWorkList_Org_TreeNodeOperations()
'4. Fn_MyWorklist_TaskComplete()
'5. Fn_MyWorklist_PerformSignOff()
'6. Fn_MyWorkList_AssignResponsibleParty()
'7. Fn_MyWorkList_AssignParticipant()
'8. Fn_MyWorkList_SignoffTeamSelect() 
'9. Fn_MyWorkList_ResourcePoolSubscription()
'10. Fn_MyWorkList_TaskPromote(sTaskName, sComment)
'11. Fn_MyWorkList_DelegateSignoff()
'12. Fn_MyWorkList_ViewAuditLog(sAction,aLog)
'13. Fn_MyWorkList_FormatDate()
'14. Fn_MyWorkList_WorkflowSurrogate()
'15. Fn_MyWorkList_OutOfOfficeAssit()
'16. Fn_MyWorkList_TaskDemote()
'17. Fn_MyWorkList_WorkflowSubProcessAssign()
'18. Fn_MyWorkList_TaskSuspend(sTaskName, sComment)
'19. Fn_MyWorkList_ProcessView_Attributes(sAction, dicProcessViewAttributes)
'20. Fn_WorkflowSubProcess_Operations(sAction, dicNewSubProcess, sBtnName)
'21. Fn_MyWorkList_ViwerPaneProperties()
'22. Fn_MyWorkList_Menu_TaskComplete(sAction, sTaskName, sTaskInstruction, sProcessDesc, sComment, radComplete, sPassword)
'23. Fn_MyWorkList_ProcessView_ProcessTreeOperations(sAction, sNodeName, sMenu)
'24. Fn_MyWorkList_HTML_PropertyVerify()
'25. Fn_MyWorkList_TaskActions()
'26. Fn_MyWorkList_SignoffTeamSelectDialog_Opearations()
'27. Fn_MyWorkList_DialogMsgVerify(sAction, sTitle,sMsg,sButton) 
'28. Fn_MyWorkList_SurrogateActions(sAction, sTaskName, sSignoffMember, sActiveSurrogate, sStandORRelease, bCheckOut, sBtnClick)
'29. Fn_MyWorkList_Signoff_AddressList_Operations(sAction, dicSelectSignOff)
'30. Fn_MyWorkList_ClaimPerformSignoffOperations
'31. Fn_MyWorkList_TaskAbort
'32. Fn_MyWorkList_TaskPromoteOperations
'*********************************************************	Function List		***********************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_MyWorkList_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_MyWorkList_GetObject("Workflow Surrogate")

'History:

'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Nilesh Gadekar		 07-Aug-2012		1.0											Created
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_MyWorkList_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\MyWorkList.xml"
	Set Fn_SISW_MyWorkList_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'*********************************************************  Function do Operation on WorkList Tree *********************************************************************

'Function Name		:					Fn_MyWorkList_TreeNodeOperations

'Description			 :		 		    Action  performed :-
'																	1. Node Select
'																	2. Node multi-select
'																	3. Node Expand
'																	4. Node Collapse
'																	5. Node Popup menu select
'																    6.Exist

'Parameters			   :	 			1. sAction: Action to be performed
'													2.sNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
' 												   3. sMenu: Context menu to be selected

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 MyWorklist pane should be displayed.

'Examples				:			 Fn_MyWorkList_TreeNodeOperations("PopupMenuSelect", "My Worklist:Rupali Palhade (x_palhad) Inbox","Send To:Schedule Manager")
												' Fn_MyWorkList_TreeNodeOperations("ExistWithTild","My Worklist~AutoTest4 (autotest4) Inbox~Tasks to Perform~Component Process:1 (select-signoff-team)","")
												' Fn_MyWorkList_TreeNodeOperations("Exist","My Worklist:AutoTest4 (autotest4) Inbox:Tasks to Perform","")
												' Fn_MyWorkList_TreeNodeOperations("ExpandWithTilda","My Worklist~AutoTest4 (autotest4) Inbox~Tasks to Perform~Component Process:1 (select-signoff-team)","")
												' Fn_MyWorkList_TreeNodeOperations("Expand","My Worklist:AutoTest4 (autotest4) Inbox:Tasks to Perform","")
												' Fn_MyWorkList_TreeNodeOperations("ExistWithoutRefresh","My Worklist:AutoTest4 (autotest4) Inbox:Tasks to Perform","")
												' Fn_MyWorkList_TreeNodeOperations("ExistWithoutRefreshWithTild","My Worklist:AutoTest4 (autotest4) Inbox:Tasks to Perform","")
												' Fn_MyWorkList_TreeNodeOperations("SelectWithoutRefresh","My Worklist:AutoTest4 (autotest4) Inbox:Tasks to Perform","")
												' Fn_MyWorkList_TreeNodeOperations("SelectWithoutRefreshWithTild","My Worklist:AutoTest4 (autotest4) Inbox:Tasks to Perform","")
'												Fn_MyWorkList_TreeNodeOperations("GetChildItemCount","My Worklist:AutoTest7 (autotest7) Inbox:Tasks To Perform:Sub Process for Attaches (select-signoff-team):References","")


'History:
'										Developer Name			Date					Rev. No.			Changes Done																				Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Rupali				21-Jun-2010	       		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Shrikant N			27-Mar-2012         	1.1             Modied case ExistWithTild , Exist, ExpandWithTilda , Expand	 								Koustubh W															
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'						        			Ashok				29-Mar-2012     	    1.1			Modified Case  SelectSameInstanceWithTild, SelectSameInstance 								Koustubh W
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Shrikant N			06-Apr-2012         	1.1             Added condition for refresh tree	 								                       Koustubh W		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Shweta Rathod		04-Oct-2017				1.0			Added case "GetChildItemCount" to get child count from the node								Shweta R	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_TreeNodeOperations(sAction,sNodeName,sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TreeNodeOperations"
   On Error Resume Next
   Dim arrNodeList,objSearchTree,iItemCount,iCounter,sTreeItem,arrNode,iOuterCount,aMenuList
   Dim aNodePath, iGlobalCnt, iInstance, bFlag, sPath, strItemName,intCounter, iItemCount1, iItemCount2
   Dim App,sNewNodeName
   Dim i, sRefrshNode,sNode,sDel,oCurrentNode
	sDel = ":"
	If instr(sNodeName,"~") Then
		sDel = "~"
	End If

	'In Tc10.1_0313 build letter case got changed for addressing translation PR
	sNodeName = replace(sNodeName, "Tasks to Perform", "Tasks To Perform")
	sNodeName = replace(sNodeName, "Tasks to Track", "Tasks To Track")

   Set objSearchTree =  JavaWindow("MyWorkListWindow").JavaTree("MyWorkListTree")
   objSearchTree.Object.SetFocus
   If objSearchTree.Exist(5) Then
'		Added code by Shrikant
		 If instr(sAction,"Tild") >0 Then sDel="~"
		arrNodeList=Split(sNodeName, sDel,-1,1) 

		'Added by Prasanna to refresh the tree
		If instr(1,arrNodeList(ubound(arrNodeList)),"Inbox") > 0 Then
				objSearchTree.Select "#0"
				wait 3
				If instr(sAction,"WithoutRefresh") >0 Then 'Added By Shrikant
					' do nothing
				Else
					call Fn_MenuOperation("Select", "View:Refresh")
				End If
				wait 1
                Call Fn_ReadyStatusSync(3)
		End If

		
		'If  arrNodeList(ubound(arrNodeList)) = "Tasks to Track" or arrNodeList(ubound(arrNodeList)) = "Tasks to Perform" Then
		If InStr(sNodeName, "Tasks to Track") > 0 or InStr(sNodeName, "Tasks to Perform") > 0 Then
					For i = 0 to Ubound(arrNodeList)
						sRefrshNode = sRefrshNode + sDel + arrNodeList(i)'		Modified code by Shrikant
						If arrNodeList(i) = "Tasks to Track" OR arrNodeList(i) = "Tasks to Perform" Then
							Exit For
						End If
					Next
					sRefrshNode = Right(sRefrshNode, Len(sRefrshNode)-1)
					sRefrshNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sRefrshNode, sDel)
					If sRefrshNode = False Then
						Fn_MyWorkList_TreeNodeOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get Node  Index for [" + sRefrshNode + "] of MyworkList Tree." )	
						Set objSearchTree = Nothing
						Exit Function
					End If
					objSearchTree.Select Trim(sRefrshNode)
					wait 3
					If instr(sAction,"WithoutRefresh") >0 Then 'Added By Shrikant
					 
					Else
						call Fn_MenuOperation("Select","View:Refresh")
					End If
					wait 1
					Call Fn_ReadyStatusSync(3)
		End If

		Select Case sAction

				Case "ExistWithTild", "ExistWithTilda", "Exist" , "ExistWithoutRefresh","ExistWithoutRefreshWithTild" '		Modified case by Shrikant
						sTreeItem = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, sDel)
						If sTreeItem = False Then
							Fn_MyWorkList_TreeNodeOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to find node " + sNodeName + "of MyworkList Tree." )	
							Set objSearchTree = Nothing
							Exit Function
						Else
							Fn_MyWorkList_TreeNodeOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully found node " + sNodeName + "of MyworkList Tree." )	
						End If
			Case  "Select", "SelectWithTild", "SelectSameInstanceWithTild","SelectSameInstance","SelectWithoutRefresh"	'		Modified case by Shrikant
				sNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, sDel)
				If sNode = False Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get Node  Index for [" + sNodeName + "] of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function
				End If
				Call Fn_ReadyStatusSync(3)
				objSearchTree.Select sNode
				wait(1)
'				objSearchTree.Click 0,0
				If Err.Number < 0 Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node  " + sNodeName + "of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_TreeNodeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected  Node  " + sNodeName + "of MyworkList Tree.")	
				End If

			Case "MultiSelect" 
				arrNode = Split(sNodeName,",")
				For iOuterCount = 0 to  Ubound(arrNode)
					sTreeItem = Fn_MyWorkList_GetItemPathIndex(objSearchTree, arrNode(iOuterCount), sDel)					
					If sTreeItem = False Then
						Exit function
					End If
					If iOuterCount <> 0 Then
						objSearchTree.ExtendSelect sTreeItem
					Else
						objSearchTree.Select sTreeItem
					End If
					If Err.Number < 0 Then
						Fn_MyWorkList_TreeNodeOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node  " + sNodeName + "of MyworkList Tree." )	
						Set objSearchTree = Nothing
						Exit Function 
					Else
						Fn_MyWorkList_TreeNodeOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected  Nodes  " + sNodeName  + "of MyworkList Tree.")	
					End If
				Next

			Case  "Expand","ExpandWithTilda"					'		Modified case by Shrikant
				sNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, sDel)
				If sNode = False Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get Node  Index for [" + sNodeName + "] of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function
				End If
				objSearchTree.Expand sNode
				Call Fn_ReadyStatusSync(3)
				If Err.Number < 0 Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to expand node   " + sNodeName + "of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_TreeNodeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully expand node  " + sNodeName  + "of MyworkList Tree.")	
				End If

			Case "Collapse"
				sNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, sDel)
				If sNode = False Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get Node  Index for [" + sNodeName + "] of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function
				End If
				objSearchTree.Collapse sNode
				If Err.Number < 0 Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Collapse node   " + sNodeName + "of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_TreeNodeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapse node  " + sNodeName  + "of MyworkList Tree.")	
				End If

			Case "PopupMenuSelect"

				aMenuList = split(sMenu, ":",-1,1)
				iCounter = Ubound(aMenuList)
				sNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, "")
				If sNode = False Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get Node  Index for [" + sNodeName + "] of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function
				End If
				objSearchTree.Select sNode
				objSearchTree.OpenContextMenu sNode
				Select Case iCounter
					Case "0"
							 sMenu = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						    sMenu = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					 Case Else
						    Fn_MyWorkList_TreeNodeOperations = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Wrong Parameter for Popup menu Select [" + sMenu + "]")	
						   Exit Function
				End Select
				JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").WinMenu("ContextMenu").Select sMenu  

				If Err.Number < 0 Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select menu   " + sMenu + "of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_TreeNodeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected menu  " + sMenu  + "of MyworkList Tree.")	
				End If
			Case "PopupMenuSelectWithTilda"
					
				aMenuList = split(sMenu, ":",-1,1)
				iCounter = Ubound(aMenuList)
				sNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, "~")
				If sNode = False Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get Node  Index for [" + sNodeName + "] of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function
				End If
				objSearchTree.Select sNode
				objSearchTree.OpenContextMenu sNode
				Select Case iCounter
					Case "0"
							 sMenu = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
					Case "1"
						    sMenu = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
					 Case Else
						    Fn_MyWorkList_TreeNodeOperations = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Wrong Parameter for Popup menu Select [" + sMenu + "]")	
						   Exit Function
				End Select
				JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").WinMenu("ContextMenu").Select sMenu  

				If Err.Number < 0 Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select menu   " + sMenu + "of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_TreeNodeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected menu  " + sMenu  + "of MyworkList Tree.")	
				End If
			Case  "DoubleClick"
				Dim intX, intY, intWidth, intHeight
				sNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, "")
				If sNode = False Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get Node  Index for [" + sNodeName + "] of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function
				End If
				If Trim(sNodeName) <> "" Then
					objSearchTree.Select sNode
				End If

				'objSearchTree.Activate sNode
				intX = objItmBounds.x
				intY = objItmBounds.y
				intWidth = objItmBounds.width
				intHeight = objItmBounds.height

				Set objItmBounds = nothing
				objSearchTree.DblClick intX+intWidht+15, intY+intHeight/2, "LEFT"

				If Err.Number < 0 Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Double Clicked the Selected Node " + sNodeName + " of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_TreeNodeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Double Clicked on the Selected  Node " + sNodeName + " of MyworkList Tree.")	
				End If
			Case "GetIndexOfNode"
				sNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, "")
				If sNode = False Then
					Fn_MyWorkList_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to get Node Index for [" + sNodeName + "] of MyworkList Tree." )	
					Set objSearchTree = Nothing
					Exit Function
				End If
				
				sNodeIndex = Replace(sNode,"#","")
				aNodeIndex = Split(sNodeIndex,":")
				Fn_MyWorkList_TreeNodeOperations = aNodeIndex(UBound(aNodeIndex))
			Case "GetChildItemCount"
				if Fn_MyWorkList_TreeNodeOperations("Expand",sNodeName,"")=True Then
					sNode = Fn_MyWorkList_GetItemPathIndex(objSearchTree, sNodeName, "")
					If sNode <> False Then
						sNode = replace(sNode, "#", "") 
						arrNodeList = Split(sNode,":",-1,1)
						Set oCurrentNode = objSearchTree.Object.getItem(cInt(arrNodeList(0)))
						For iCounter = 1 to uBound(arrNodeList)
							Set oCurrentNode = oCurrentNode.getItem(cInt(arrNodeList(iCounter)))
						Next
						Fn_MyWorkList_TreeNodeOperations = cInt(oCurrentNode.getItemCount())
						Set oCurrentNode=Nothing
					Else
						Fn_MyWorkList_TreeNodeOperations = False
					End If
				Else
					Fn_MyWorkList_TreeNodeOperations = False
				End If	
			Case Else
				Fn_MyWorkList_TreeNodeOperations = False

		End Select
   Else
		Fn_MyWorkList_TreeNodeOperations = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "My WorkList Tree does not exist.")	
		Exit Function
   End If
	Set objSearchTree = Nothing
End Function

'*********************************************************  Function do Operation on WorkList Tree *********************************************************************
'	Function Name			:				Fn_MyWorkList_SignoffTeam_TreeNodeOperations

'	Description			 	:		 		Action  performed :-
'											1. Node Select																	
'											2. Node Expand
'											3. Node Collapse																	
'											4. Exist
'											5. GetSelectedNode

'	Parameters			   	:	 			1. sAction		: Action to be performed
'											2. sNodeName	: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
' 												   
'	Return Value		   	: 			 	True/False/Selected Node name

'	Pre-requisite			:		 	 	Signfoff tree should be displayed.

'	Examples				:			 	Fn_MyWorkList_SignoffTeam_TreeNodeOperations("Select", "Profiles")
'											Fn_MyWorkList_SignoffTeam_TreeNodeOperations("GetSelectedNode","")

'	History:
'
'	Developer Name			Date				Rev. No.			Changes Done											Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Prasanna			03-Aug-2010	       		1.0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam			25-Feb-2016	       		1.0			Added New Case "GetSelectedNode"				[Tc1122:2016021000:25Feb2016:AnkitN:NewDevelopment]
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_SignoffTeam_TreeNodeOperations(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_SignoffTeam_TreeNodeOperations"
	On Error Resume Next
	Dim arrNodeList,iItemCount,iCounter,sTreeItem,arrNode,iOuterCount,aMenuList
	Dim objSignOffTree,iAll,iNodeCount,iRowCounter
	set objDialogView=JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
	objDialogView.JavaRadioButton("ViewOptions").SetTOProperty "Attached Text","Task View"
	objDialogView.JavaRadioButton("ViewOptions").Set "ON"
	If Err.Number < 0 Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select task View  from Viewer tab.")
	Fn_MyWorkList_SignoffTeam_TreeNodeOperations = False
	Set objDialog = Nothing
	Set objDialogView = Nothing
	Exit Function
	Else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Task View from Viewer tab.")
	Call Fn_ReadyStatusSync(2)
	Wait(2)
	End If

	Set objDialogView = Nothing

   	If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("SignOffTeamTree").Exist Then
			Set objSignOffTree =  JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("SignOffTeamTree")
			iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_MyWorkList_SignoffTeam_TreeNodeOperations", JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow") , "SignOffTeamTree", sNodeName, "", "")
	Else
			Set objSignOffTree =  JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("RouteSignOffTeamTree")
			iRowCounter = Fn_JavaTree_NodeIndexExt("Fn_MyWorkList_SignoffTeam_TreeNodeOperations", JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow") , "RouteSignOffTeamTree", sNodeName, "", "")
	
	End If

If objSignOffTree.Exist(5) Then
	   Select Case sAction
				Case  "Select"		
				iNodeCount = objSignOffTree.GetROProperty("items count")
				If instr(1,sNodeName,":") Then
						aSplitNodeName=split(sNodeName,":",-1,1)
						sNodeName=aSplitNodeName(ubound(aSplitNodeName))
				End If


				For iAll = 0 to iNodeCount-1
						sTempNode = objSignOffTree.Object.getPathForRow(iAll).tostring()
							If instr(1,sTempNode,sNodeName) Then
								iRowCounter = iAll
							Exit for
						End If
				Next

				objSignOffTree.Object.setSelectionRow iRowCounter
	
				 If Err.Number < 0 Then
							Fn_MyWorkList_SignoffTeam_TreeNodeOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node  " + sNodeName + "of Signoff Tree." )	
							Set objSignOffTree = Nothing
							Exit Function 
				Else
							Fn_MyWorkList_SignoffTeam_TreeNodeOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected  Node  " + sNodeName + "of Signoff Tree.")	
				End If  

			Case "Exist"
							
				iNodeCount = objSignOffTree.GetROProperty("items count")
				If instr(1,sNodeName,":") Then
						aSplitNodeName=split(sNodeName,":",-1,1)
						sNodeName=aSplitNodeName(ubound(aSplitNodeName))
				End If
				For iAll = 0 to iNodeCount-1
					sTempNode = objSignOffTree.Object.getPathForRow(iAll).tostring()
						If instr(1,sTempNode,sNodeName) Then
							Fn_MyWorkList_SignoffTeam_TreeNodeOperations = True
						Exit for
					End If
				Next
				
				If  Cint(iAll) = Cint (iNodeCount) Then
					Fn_MyWorkList_SignoffTeam_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Find node " + sNodeName + "of Signoff Tree." )	
					Set objSignOffTree = Nothing
					Exit Function 
				End If

			Case  "Expand"
				
				iNodeCount = objSignOffTree.GetROProperty("items count")
				If instr(1,sNodeName,":") Then
						aSplitNodeName=split(sNodeName,":",-1,1)
						sNodeName=aSplitNodeName(ubound(aSplitNodeName))
				End If
				For iAll = 0 to iNodeCount-1
						sTempNode = objSignOffTree.Object.getPathForRow(iAll).tostring()
							If instr(1,sTempNode,sNodeName) Then
								objSignOffTree.Object.expandPath(objSignOffTree.Object.getPathForRow(iAll))
								wait(2)
							Exit for
						End If
				Next
				
				If Err.Number < 0 Then
					Fn_MyWorkList_SignoffTeam_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to expand node   " + sNodeName + "of Sign off Tree." )	
					Set objSignOffTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_SignoffTeam_TreeNodeOperations = True
					wait(2)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully expanded node  " + sNodeName  + "of Sign off Tree.")	
				End If

			Case  "Collapse"
				sNodeName = "#0:"+sNodeName
				objSignOffTree.Collapse sNodeName
				If Err.Number < 0 Then
					Fn_MyWorkList_SignoffTeam_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Collapse node   " + sNodeName + "of Sign off Tree." )	
					Set objSignOffTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_SignoffTeam_TreeNodeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapsed node  " + sNodeName  + "of Sign off Tree.")	
					wait(2)
				End If
'----------------------------------------------------------------------------------------------------------------------------------------------------------------			
			Case "GetSelectedNode"					'[Tc1122:2016021000:25Feb2016:AnkitN:NewDevelopment] - Added Case to get selected node in signoff Tree
					Fn_MyWorkList_SignoffTeam_TreeNodeOperations = objSignOffTree.Object.getSelectedNode().toString()
'----------------------------------------------------------------------------------------------------------------------------------------------------------------								
			Case else
					Fn_MyWorkList_SignoffTeam_TreeNodeOperations = False

		End Select
Else
		Fn_MyWorkList_SignoffTeam_TreeNodeOperations = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sign Off Tree Tree does not exist.")	
		Exit Function
End If

	Set objSignOffTree = Nothing

End Function 

'*********************************************************  Function do Operation on WorkList Tree *************************************************************
'	Function Name			:				Fn_MyWorkList_Org_TreeNodeOperations

'	Description			 	:		 		Action  performed :-
'											1. Node Select																	
'											2. Node Expand
'											3. Node Collapse																	
'											4. Exist
'											5. MultiSelect

'	Parameters			   	:	 			1. sAction		: Action to be performed
'											2. sNodeName	: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 

'	Return Value		  	: 			 	True/False

'	Pre-requisite			:		 	 	Signfoff tree should be displayed.

'	Examples				:			 	Fn_MyWorkList_Org_TreeNodeOperations("MultiSelect","AutoGrp1:AutoRole1,Engineering:Designer:AutoTest5 (autotest5)")
'											Fn_MyWorkList_Org_TreeNodeOperations("Search","User$AutoTest1~Group$Designer~Role$Engineering")
'											Fn_MyWorkList_Org_TreeNodeOperations("VerifySearchEditBoxToolTip","Role$Role name criteria. Wildcard characters are allwed.~Group$Group name criteria. Wildcard characters are alowed.")
'
'	History					:
'										
'	Developer Name		Date		Rev. No.							Changes Done												Reviewer	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Prasanna		03-Aug-2010	      1.0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam		25-Feb-2016	      1.0			Added new Case "Search" from Tc1015							[Tc1122:2016021000:25Feb2016:AnkitN:NewDevelopment]
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam		03-Mar-2016	      1.0			Added new Case "VerifySearchEditBoxToolTip"					[Tc1122:2016021000:03Mar2016:AnkitN:NewDevelopment]
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_Org_TreeNodeOperations(sAction, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_Org_TreeNodeOperations"
  	On Error Resume Next
  	Dim iItemCount, iCounter, iOuterCount
  	Dim sExpandNode, sCategory, sText, sTreeItem
  	Dim arrNodeList, aMember, arrNode, aMenuList
  	Dim objOrgTree, objWApplet
  	Dim bReturn

  	Set objOrgTree =  JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("OrgTree")

	If objOrgTree.Exist(5) Then
		Select Case sAction
			Case  "Select"					
					sNodeName = "#0:" + sNodeName
					objOrgTree.Select sNodeName 

					If Err.Number < 0 Then
						Fn_MyWorkList_Org_TreeNodeOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node  " + sNodeName + "of Oraganization Tree." )	
						Set objOrgTree = Nothing
						Exit Function 
					Else
						Fn_MyWorkList_Org_TreeNodeOperations = True
						Call Fn_ReadyStatusSync(5)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected  Node  " + sNodeName + "of Oraganization Tree.")	
					End If	
			
			Case "Exist"
					iItemCount = objOrgTree.GetROProperty( "items count")
					If sNodeName = "" Then      'TC1122-2016032300-05_04_2016-VivekA-NewDevelopment - Changed by AnkitN
						sNodeName = "Organization"
					Else 
						sNodeName = "Organization:"+sNodeName
					End If
					
					For iCounter=0 To (iItemCount-1)
						sTreeItem = objOrgTree.GetItem(iCounter)
						
						If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_MyWorkList_Org_TreeNodeOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully found node " + sNodeName + "of Oraganization Tree." )	
							Exit For
						End If
					Next 	

					If  Cint(iCounter) = Cint (iItemCount) Then
						Fn_MyWorkList_Org_TreeNodeOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  find node " + sNodeName + "of Oraganization Tree." )	
						Set objOrgTree = Nothing
						Exit Function 
					End If

			Case  "Expand"
				sNodeName = "#0:"+sNodeName
				objOrgTree.Expand sNodeName
				If Err.Number < 0 Then
					Fn_MyWorkList_Org_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to expand node   " + sNodeName + "of Oraganization Tree." )	
					Set objOrgTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_Org_TreeNodeOperations = True
					Call Fn_ReadyStatusSync(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully expanded node  " + sNodeName  + "of Oraganization Tree.")	
				End If

			Case  "Collapse"
				sNodeName = "#0:"+sNodeName
				objOrgTree.Collapse sNodeName
				If Err.Number < 0 Then
					Fn_MyWorkList_Org_TreeNodeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Collapse node   " + sNodeName + "of Oraganization Tree." )	
					Set objOrgTree = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_Org_TreeNodeOperations = True
					Call Fn_ReadyStatusSync(5)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapsed node  " + sNodeName  + "of Oraganization Tree.")	
				End If

			Case "MultiSelect" 
				arrNodeList = split(sNodeName, ",",-1,1)
				
				For iCounter = 0 to UBound(arrNodeList)
					arrNode = split(arrNodeList(iCounter),":",-1,1)
					sExpandNode = "#0:"
					For iItemCount = 0 to UBound(arrNode)-1
						sExpandNode = sExpandNode + arrNode(iItemCount)
						wait(2)
						objOrgTree.Expand sExpandNode
						If Err.Number < 0 Then
								Fn_MyWorkList_Org_TreeNodeOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Expand node   " + sExpandNode + "of Oraganization Tree." )	
								Set objOrgTree = Nothing
								Exit Function						
						End If
						Call Fn_ReadyStatusSync(5)
						sExpandNode = sExpandNode + ":"
					Next						
				Next
				objOrgTree.Object.clearSelection
				For iCounter = 0 to UBound(arrNodeList)
							If iCounter = 0 Then
									Wait(2)
									objOrgTree.Select "#0:" + arrNodeList(iCounter)
							Else
									Wait(2)
									objOrgTree.ExtendSelect "#0:" + arrNodeList(iCounter)
							End If
							If Err.Number < 0 Then
								Fn_MyWorkList_Org_TreeNodeOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Mulitiselected node  " + arrNodeList(iCounter)  + "of Oraganization Tree.")
								Exit Function
							End If
				Next
				Fn_MyWorkList_Org_TreeNodeOperations = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Mulitiselected node  " + sNodeName  + "of Oraganization Tree.")
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			'[Tc1122:2016021000:25Feb2016:AnkitN:NewDevelopment] - Added new case to Search User, Group and Role under organization tree from Tc1015
			Case  "Search"
					Set objWApplet = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
					If Instr(sNodeName, "~") > 0 Then
						sCategory = Split(sNodeName, "~")
						For iCounter = 0 To UBound(sCategory)
							aMember = Split(sCategory(iCounter), "$")
							If aMember(0) = "User" Then
								sText = "Enter User ID or User Name"
							ElseIf aMember(0) = "Group" Then
								sText = "Enter Group Name"
							ElseIf aMember(0) = "Role" Then
								sText = "Enter Role Name"
							End If
							objWApplet.JavaEdit("UsrRoleGrpSrch").SetTOProperty "text", sText
							Wait 1
							bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_MyWorkList_Org_TreeNodeOperations", "Set", objWApplet, "UsrRoleGrpSrch", aMember(1))
							If bReturn = False Then
								Fn_MyWorkList_Org_TreeNodeOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to enter User ID or User Name  " + aMember(1) + "of Organization Tree." )	
								Set objWApplet = Nothing
								Exit Function 
							End If
						Next
					Else
						sCategory = sNodeName
						aMember = Split(sCategory,"$")
						If aMember(0) = "User" Then
							sText = "Enter User ID or User Name"
						ElseIf aMember(0) = "Group" Then
							sText = "Enter Group Name"
						ElseIf aMember(0) = "Role" Then
							sText = "Enter Role Name"
						End If
						objWApplet.JavaEdit("UsrRoleGrpSrch").SetTOProperty "text",sText
						Wait 1
						bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_MyWorkList_Org_TreeNodeOperations", "Set",  objWApplet, "UsrRoleGrpSrch", aMember(1))
						If bReturn = False Then
							Fn_MyWorkList_Org_TreeNodeOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to enter User ID or User Name  " + aMember(1) + "of Organization Tree." )	
							Set objWApplet = Nothing
							Exit Function 
						End If
					End If
					
					objWApplet.JavaButton("FindUser").SetTOProperty "label","search_16"
					Wait 1
					bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_Org_TreeNodeOperations", "Click", objWApplet, "FindUser")
					If bReturn <> True Then
						Fn_MyWorkList_Org_TreeNodeOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to click on FindUser button." )	
						Set objWApplet = Nothing
						Exit Function
					Else
						Set objWApplet = Nothing
						Fn_MyWorkList_Org_TreeNodeOperations = True
					End If				
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			'[Tc1122:2016021000:03Mar2016:AnkitN:NewDevelopment] - Added new case to Verify Tool Tip Text of Search Edit Boxes
			Case "VerifySearchEditBoxToolTip"
					Set objWApplet = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
					If Instr(sNodeName, "~") > 0 Then
						sCategory = Split(sNodeName, "~")
						For iCounter = 0 To UBound(sCategory)
							aMember = Split(sCategory(iCounter), "$")
							If aMember(0) = "User" Then
								sText = "Enter User ID or User Name"
							ElseIf aMember(0) = "Group" Then
								sText = "Enter Group Name"
							ElseIf aMember(0) = "Role" Then
								sText = "Enter Role Name"
							End If
							objWApplet.JavaEdit("UsrRoleGrpSrch").SetTOProperty "text", sText
							Wait 1
							bFlag = Fn_VerifyToolTipText(objWApplet.JavaEdit("UsrRoleGrpSrch"), aMember(1))
							If bFlag = False Then
								Fn_MyWorkList_Org_TreeNodeOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Tooltip text of Object ." )	
								Set objWApplet = Nothing
								Exit Function 
							End If
						Next
					Else
						sCategory = sNodeName
						aMember = Split(sCategory,"$")
						If aMember(0) = "User" Then
							sText = "Enter User ID or User Name"
						ElseIf aMember(0) = "Group" Then
							sText = "Enter Group Name"
						ElseIf aMember(0) = "Role" Then
							sText = "Enter Role Name"
						End If
						objWApplet.JavaEdit("UsrRoleGrpSrch").SetTOProperty "text",sText
						Wait 1
						bReturn = Fn_VerifyToolTipText(objWApplet.JavaEdit("UsrRoleGrpSrch"), aMember(1))
						If bReturn = False Then
							Fn_MyWorkList_Org_TreeNodeOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Tooltip text of Object ." )	
							Set objWApplet = Nothing
							Exit Function 
						End If
					End If
	
					Fn_MyWorkList_Org_TreeNodeOperations = True
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------			
			Case Else
				Fn_MyWorkList_Org_TreeNodeOperations = False
		End Select
	Else
		Fn_MyWorkList_Org_TreeNodeOperations = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Oraganization  Tree does not exist.")	
		Exit Function
	End If

	Set objOrgTree = Nothing

End Function 

'*********************************************************  Function do Operation on WorkList Tree *********************************************************************

'Function Name		:					Fn_MyWorklist_TaskComplete

'Description			 :		 		   Function is perform new do to task  																

'Parameters			   :	 			 1. sTaskInstruction: Instruction to be set
'												 2. sProcessDesc: Process description to be added
' 												 3. sComment: Comment to be added
'												 4. radComplete : Radio button to set the Complete value	

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 To Do Task should be selected

'Examples				:			 Fn_MyWorklist_TaskComplete("Test1","Test2","Test3",true)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Prasanna				10-Aug-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MyWorklist_TaskComplete(sTaskInstruction,sProcessDesc,sComment,radComplete)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorklist_TaskComplete"
	Dim iRow, iCol, aComment, iCnt,sPwd

   On error resume next
	'Set Default View to Process View 
	 JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ViewOptions").SetTOProperty "Attached Text","Task View"
	 JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ViewOptions").Set "ON"
	 Call Fn_ReadyStatusSync(5)
	 wait(2)
 ' Set the Instructions	
   If sTaskInstruction <> "" Then		 
	   If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("TaskInstructions").GetROProperty("editable") =1 Then
					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("TaskInstructions").set sTaskInstruction  				
					If Err.Number < 0 Then
							Fn_MyWorklist_TaskComplete = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to set instructions for Task" ) 						
							Exit Function 
					Else
							Fn_MyWorklist_TaskComplete = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set instructions for Task")	
					End If  
	   End If
   End If

' Set the Process Description
	If  sProcessDesc <> "" Then
				JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("VwerPaneProcessDesc").Set sProcessDesc
				If Err.Number < 0 Then
						Fn_MyWorklist_TaskComplete = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to set Process Description for Task" ) 						
						Exit Function 
				Else
						Fn_MyWorklist_TaskComplete = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set Process Description for Task")	
				End If  
	End If

	'Check whether comments contains password
	If  instr(1,sComment,":")  Then
					aComments = split(sComment,":",-1,1)
					sPwd = aComments(1)           
					sComment = aComments(0)

					'Set the Password ( set blank )### Changes by Sukhada 22 March 2011
					'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("Password").Set ""
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					'Set the Password 
					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("Password").Set sPwd
					 If Err.Number < 0 Then
							Fn_MyWorklist_TaskComplete = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set the Password")     							
							Exit Function						
					 Else
							Fn_MyWorklist_TaskComplete = true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set the Password")	
							wait(1)
					End If		
			End If
	
' Set the Comment
	If sComment <> "" Then
		If 	JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("TaskComments").GetROProperty("editable") = "1" Then
			If Instr(sComment, ":") > 0 Then
					'Split to extract Row & Column info
					aComment = Split(sComment, ":", -1,1)
					sComment = aComment(0)
					iRow = Cint(aComment(1))
					If Ubound(aComment) > 1 Then
						iCol = Cint(aComment(2))
					Else
						iCol = 0
					End If

					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("TaskComments").SetCaretPos iRow, iCol
					For iCnt = iCol to Len(sComment) - 1
						JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("TaskComments").Insert Mid(sComment, iCnt, 1), iRow, iCnt
					Next
					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("TaskComments").Insert "" + vbLf + "", iRow, iCnt

			Else
					'To remove comments
					If lcase(sComment) = "blank" Then
						sComment = ""
					End If
					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("TaskComments").Set	sComment		
			End If
			'Capture Error on inserting comment
			If Err.Number < 0 Then
					Fn_MyWorklist_TaskComplete = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to set Comment for Task" )
					Exit Function 
			Else
					Fn_MyWorklist_TaskComplete = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set Comment for Task")	
			End If  
		End If
	End If

' Set the Complete Radio button
	Wait 5

	If radComplete <> "" Then
		'Commented by Vallari - 19-Aug-2010
'			If radComplete = true Then
'					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").javaradiobutton("taskcomplete").set  "ON"
'			Else
'					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").javaradiobutton("taskcomplete").set  "OFF"
'			End If

			'Added by Vallari - 19-Aug-2010
			'To make descision  selection generic
			JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").javaradiobutton("taskcomplete").SetTOProperty "attached text", radComplete
			JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").javaradiobutton("taskcomplete").set  "ON"

			If Err.Number < 0 Then
				Fn_MyWorklist_TaskComplete = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set [" + cstr(radComplete) + "] Option as True")
				Exit Function 
			Else
				Fn_MyWorklist_TaskComplete = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set Set [" + cstr(radComplete) + "] Option as True")	
			End If      
	End If

	' Click on Apply button
	'Below line is commented by Rupali (29-Dec-2010) [TC9.0 Build 1220 not required . It gives error operation can not performed]
'	JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").Activate


	'Commented by Vallari - 24-Dec-10
	'bReturn=Fn_Button_Click("Fn_MyWorklist_TaskComplete",JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow"),"Apply")

	'Added By Vallari - 24-Dec-10 - Synch for Apply button to get Enabled
	JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").WaitProperty "enabled", "1", 60000
	JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").Click micLeftBtn

	Wait(5)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Check Existance of Apply Button // Modified by : Harshal Tanpure //  Date 29-March-2011
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	If JavaWindow("MyTeamcenter").JavaWindow("Error").Exist Then
'		JavaWindow("MyTeamcenter").JavaWindow("Error").JavaButton("OK").Click micLeftBtn
'		Fn_MyWorklist_TaskComplete = False
'		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Apply Changes Due to Error")
'		Exit Function 
'	Else
'			If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").Exist  AND JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").GetROProperty ("enabled")="1" AND JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").GetROProperty ("displayed")="1" Then
'					
'						JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").javaradiobutton("taskcomplete").SetTOProperty "attached text", radComplete
'						JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").javaradiobutton("taskcomplete").set  "ON"
'						Wait(5)
'						If Err.Number < 0 Then
'							Fn_MyWorklist_TaskComplete = False
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Set [" + cstr(radComplete) + "] Option as True")
'							Exit Function 
'						Else
'							Fn_MyWorklist_TaskComplete = True
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully set Set [" + cstr(radComplete) + "] Option as True")	
'						End If      
'			
'						JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").WaitProperty "enabled", "1", 60000
'						JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").Click micLeftBtn
'						Wait(5)
'						If Err.Number < 0 Then
'									Fn_MyWorklist_TaskComplete = False
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Apply Chnages" )
'									Exit Function 
'						Else
'									Fn_MyWorklist_TaskComplete = True
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Applied Changes")	
'						End If      
'			
'				Else
'						
'						If Err.Number < 0 Then
'									Fn_MyWorklist_TaskComplete = False
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Apply Chnages" )
'									Exit Function 
'						Else
'									Fn_MyWorklist_TaskComplete = True
'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Applied Changes")	
'						End If      
'					
'				End If
'	End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   

	'If Error Window Appears Handle it 
	JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Error").SetTOProperty "title","Error"
	If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Error").Exist Then
'	If JavaWindow("MyTeamcenter").JavaWindow("Error").Exist Then
'				JavaWindow("MyTeamcenter").JavaWindow("Error").JavaButton("OK").Click micLeftBtn
				JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Error").JavaButton("OK").Click 
				Fn_MyWorklist_TaskComplete = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Apply Changes Due to Error")
				Exit Function 
	End If 
	
End Function

'*********************************************************  Function perform the signoff operation *********************************************************************

'Function Name		:					Fn_MyWorklist_PerformSignOff

'Description			 :		 		   Function perform the signoff operation
'												 There will be 2 ways to perform this action which should be incorporated in one function.
'                                                                                                                                                                                                         
'Parameters			   :	 			1.  sAction: Action to be performed
'												2. sDecision : Decision need to take to perform signoff opration.
' 												3. sComment: Comment need to add.   

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Perform Signfoff Item should be displayed.

'Examples				:			 Call Fn_MyWorklist_PerformSignOff("ViewerPane","Approve","Approved By User")
'											     Call Fn_MyWorklist_PerformSignOff("Menu","Approve","Approved By User")
'												Password case :  Call Fn_MyWorklist_PerformSignOff("Menu","Approve","Approved By User:AutoTest1")


'History:
'										Developer Name			Date				Rev. No.			Changes Done																									Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Prasanna				11-Aug-2010	       1.0
'											 Prasanna				12-Oct-2010         1.1
'											 Veena					22-Mar-2017         1.0          Added Case"ViewerPane_WittErrorDialog" - toverify error dialog is exist 											Shweta rathod
'											Shweta Rathod           13-Nov-17			1.0			 Added case "VerifyChngUsrSetDlg","VerifySignoffDlg" - to verify Change user setting dialog and signoff dialog.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MyWorklist_PerformSignOff(sAction , sDecision , sComment) 
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorklist_PerformSignOff"
   On Error Resume Next
	
   Dim bReturn,objParent,sDelInfo,sDelRow
   Dim aComments, sPwd
   Select Case sAction
			 	Case "ViewerPane","ViewerPane_WittErrorDialog"
							' Set the Viewer tab
							bReturn = Fn_MyTc_TabSet("Viewer")
							If Err.Number < 0 Then
									Fn_MyWorklist_PerformSignOff = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set Viewer tab") 		
									Exit Function						
							Else
									Fn_MyWorklist_PerformSignOff = true
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Viewer tab")
									Wait(2)
							End If
							Call Fn_ReadyStatusSync(1)	
							' Click on No Desicion Link
                            If sDecision <> "" Then
									If Instr(1,sDecision,":") Then
											sDelInfo = split(sDecision,":",-1,1)
											sDelRow = sDelInfo(1)
											sDecision = sDelInfo(0)
											JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet").JavaTable("StoreProcTable").SelectCell "#"+sDelRow,"2"
									Else
											JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet").JavaTable("StoreProcTable").SelectCell "#0","2"											
									End If
							End If

							wait(2)
							If Err.Number < 0 Then
									Fn_MyWorklist_PerformSignOff = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Decision Link") 		
									Exit Function						
							Else
									Fn_MyWorklist_PerformSignOff = true
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Decision Link")
									Wait(5)
							End If				

								
							'Set the parent Object
							If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Signoff Decision").Exist Then
								Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
							Else
								Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter")		
							End If

				Case "ChngUsrSet_ViewerPane"
							' Set the Viewer tab
							bReturn = Fn_MyTc_TabSet("Viewer")
							If Err.Number < 0 Then
									Fn_MyWorklist_PerformSignOff = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set Viewer tab") 		
									Exit Function						
							Else
									Fn_MyWorklist_PerformSignOff = true
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Viewer tab")
									Wait(2)
							End If
	
							' Click on No Desicion Link
                            If sDecision <> "" Then
										If Instr(1,sDecision,":") Then
												sDelInfo = split(sDecision,":",-1,1)
												sDelRow = sDelInfo(1)
												sDecision = sDelInfo(0)												
												JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("VwerPaneSignOffTable").ClickCell sDelRow,1,"LEFT","NONE"
										Else
												JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("VwerPaneSignOffTable").ClickCell 0,1,"LEFT","NONE"
										End If
							End If
							'JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("VwerPaneSignOffTable").ClickCell 0,1,"LEFT","NONE"
							If Err.Number < 0 Then
									Fn_MyWorklist_PerformSignOff = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Decision Link") 		
									Exit Function						
							Else
									Fn_MyWorklist_PerformSignOff = true
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Decision Link")
									Wait(5)
							End If	

							If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Change User Setting").Exist(10) Then
									JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Change User Setting").JavaButton("Yes").Click micLeftBtn
									If Err.Number < 0 Then
										Fn_MyWorklist_PerformSignOff = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Yes Button on Change User Setting Dialog") 		
										Exit Function
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Yes Button on Change User Setting Dialog") 
									End If
							Else
								Fn_MyWorklist_PerformSignOff = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Find Change User Setting Dialog") 		
								Exit Function
							End If
	
							'Set the parent Object
							If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Signoff Decision").Exist Then
								Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
							Else
								Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter")		
							End If

				Case "Menu","MenuExt"
							bReturn = Fn_MenuOperation("Select", "Actions:Perform")
							Call Fn_ReadyStatusSync(5)
							If bReturn = False Then
									Fn_MyWorklist_PerformSignOff = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Action --> Perform ") 		
									Exit Function						
							Else
									Fn_MyWorklist_PerformSignOff = true
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Action --> Perform")
									Wait(5)
							End If
	
							'Set the parent Object
						If JavaWindow("MyWorkListWindow").Exist(10) Then
									If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Perform Signoff").Exist(10) Then
										Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
									Else
										Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter")	
									End If
							Else
									Set objParent = JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks")	
							End If

							If objParent.JavaDialog("Perform Signoff").Exist(10) = False Then
							    Set objParent=JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame")
							End If

							' Click on No Desicion Link							
                            If sDecision <> "" Then
									If Instr(1,sDecision,":") Then
											sDelInfo = split(sDecision,":",-1,1)
											sDelRow = sDelInfo(1)
											sDecision = sDelInfo(0)
											objParent.JavaDialog("Perform Signoff").JavaTable("SignOffTable").ClickCell sDelRow,2,"LEFT","NONE"
									Else
											objParent.JavaDialog("Perform Signoff").JavaTable("SignOffTable").ClickCell 0,2,"LEFT","NONE"											
									End If
							End If
							wait(2)
							'objParent.JavaDialog("Perform Signoff").JavaTable("SignOffTable").ClickCell 0,1,"LEFT","NONE"							
							If Err.Number < 0 Then
									Fn_MyWorklist_PerformSignOff = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Decision Link") 		
									objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
									objParent = nothing
									Exit Function						
							Else
									Fn_MyWorklist_PerformSignOff = true
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Decision Link")
									Wait(5)
							End If

							'Check whether Change User Settings Exists or not
							If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Change User Setting").Exist(10) Then
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Change User Setting").JavaButton("Yes").Click micLeftBtn
										If Err.Number < 0 Then
											Fn_MyWorklist_PerformSignOff = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Yes Button on Change User Setting Dialog") 		
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Yes Button on Change User Setting Dialog") 
										End If
							End if
				Case "VerifyChngUsrSetDlg","VerifySignoffDlg"	
					bReturn = Fn_MenuOperation("Select", "Actions:Perform")
					Call Fn_ReadyStatusSync(5)
					If bReturn = False Then
						Fn_MyWorklist_PerformSignOff = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Action --> Perform ") 		
						Exit Function						
					Else
						Fn_MyWorklist_PerformSignOff = true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Action --> Perform")
						Wait(5)
					End If

					'Set the parent Object
					If JavaWindow("MyWorkListWindow").Exist(10) Then
						If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Perform Signoff").Exist(10) Then
							Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
						Else
							Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter")	
						End If
					Else
						Set objParent = JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks")	
					End If

					If objParent.JavaDialog("Perform Signoff").Exist(10) = False Then
					    Set objParent=JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame")
					End If

					' Click on No Desicion Link							
                    If sDecision <> "" Then
						If Instr(1,sDecision,":") Then
							sDelInfo = split(sDecision,":",-1,1)
							sDelRow = sDelInfo(1)
							sDecision = sDelInfo(0)
							objParent.JavaDialog("Perform Signoff").JavaTable("SignOffTable").ClickCell sDelRow,2,"LEFT","NONE"
						Else
							objParent.JavaDialog("Perform Signoff").JavaTable("SignOffTable").ClickCell 0,2,"LEFT","NONE"											
						End If
					End If
					wait(2)
					If sAction = "VerifyChngUsrSetDlg" then  
						If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Change User Setting").Exist(10) Then
							Fn_MyWorklist_PerformSignOff = True
							Exit function
						else
							Fn_MyWorklist_PerformSignOff = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Yes Button on Change User Setting Dialog") 		
							Exit Function
						End if
					End if	
					If sAction = "VerifySignoffDlg" then
						If objParent.JavaDialog("Signoff Decision").Exist(5) then
							If not JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Change User Setting").Exist(1) Then
								Fn_MyWorklist_PerformSignOff = True
								Exit function
							else
								Fn_MyWorklist_PerformSignOff = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Yes Button on Change User Setting Dialog") 		
								Exit Function
							End if
						else
							Fn_MyWorklist_PerformSignOff = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Yes Button on Change User Setting Dialog") 		
							Exit Function
						End if
					End if				
		     Case Else
					Fn_MyWorklist_PerformSignOff = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Operation for Action " + sAction + "Does not Exists")
					Exit Function					
   End Select
   

	' Perform the Signoff Operation
	'Set the comments
	If  objParent.JavaDialog("Signoff Decision").Exist Then		
	
			'Check whether Comments contains Passwords or not 
			If  instr(1,sComment,":")  Then
					aComments = split(sComment,":",-1,1)
					sPwd = aComments(1)           
					sComment = aComments(0)

					'Set the Password 
					objParent.JavaDialog("Signoff Decision").JavaEdit("Password").Set ""
					objParent.JavaDialog("Signoff Decision").JavaEdit("Password").Type sPwd
					wait 1
					 If Err.Number < 0 Then
							Fn_MyWorklist_PerformSignOff = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set the Password") 	
							objParent = nothing	
							Exit Function						
					 Else
							Fn_MyWorklist_PerformSignOff = true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set the Password")	
							wait(1)
					End If		
			End If

			'Set the Comments	
			If  instr(1,sComment,",")  Then
					aComments = split(sComment,",",-1,1)
					objParent.JavaDialog("Signoff Decision").JavaEdit("Comments").Set aComments(0)
					 If Err.Number < 0 Then
						Fn_MyWorklist_PerformSignOff = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set the Comments") 	
						objParent = nothing	
						Exit Function						
					Else
						Fn_MyWorklist_PerformSignOff = true
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set the Comments")	
						wait(1)
				   End If	
		 Else
			objParent.JavaDialog("Signoff Decision").JavaEdit("Comments").Set sComment
			 If Err.Number < 0 Then
					Fn_MyWorklist_PerformSignOff = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set the Comments") 	
					objParent = nothing	
					Exit Function						
			Else
					Fn_MyWorklist_PerformSignOff = true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set the Comments")	
					wait(1)
			End If	
	  End If

			'Set the Decision value
			objParent.JavaDialog("Signoff Decision").JavaRadioButton("DecisionOpt").SetTOProperty "attached text",sDecision 
			objParent.JavaDialog("Signoff Decision").JavaRadioButton("DecisionOpt").Set "ON"

			 If Err.Number < 0 Then
					Fn_MyWorklist_PerformSignOff = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set the Decision " +sDecision) 		
					objParent = nothing	
					Exit Function						
			Else
					Fn_MyWorklist_PerformSignOff = true
					wait(2)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set the Decision  " +sDecision)										
			End If	 
	Else
			Fn_MyWorklist_PerformSignOff = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Perform Decision Dialog Doe not Exist") 		
			objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
            objParent = nothing	
			Exit Function
	End If

	' Click on ok button
	wait(10)
	objParent.JavaDialog("Signoff Decision").JavaButton("OK").Click micLeftBtn
	Call Fn_ReadyStatusSync(20)
	
	If Err.Number < 0 Then
			Fn_MyWorklist_PerformSignOff = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button of Perform Signoff Decision dialog") 		
			objParent = nothing	
			Exit Function						
	Else
			Fn_MyWorklist_PerformSignOff = true
			wait(2)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on OK button of Perform Signoff Decision dialog")										
	End If	 
	wait 5
	'Handled Error Dialog if Appears
	If  instr(1,sComment,",")  Then
					aComments = split(sComment,",",-1,1)
					If  aComments(1)="ON" Then
						Fn_MyWorklist_PerformSignOff = true
					Else
						 If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(5) then
								bReturn=Fn_ErrorDialogHandle("Error","","OK")
								if bReturn = True then
										Fn_MyWorklist_PerformSignOff = False
										objParent.JavaDialog("Signoff Decision").JavaButton("Cancel").Click micLeftBtn
										wait(3)
										objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
										wait(3)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Close Error Dialog") 		
										objParent = nothing	
										Exit Function						
								End if 
						End if		
				End If
			Else
			
			If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(3) and sAction = "ViewerPane_WittErrorDialog" then
				JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Close
				Wait 1
				Fn_MyWorklist_PerformSignOff = True
				objParent.JavaDialog("Signoff Decision").JavaButton("Cancel").Click micLeftBtn
				objParent = nothing	
				Exit Function						
			End If
			
			If sAction<>"MenuExt" Then
				if JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(1) then
					bReturn=Fn_ErrorDialogHandle("Error","","OK")
					if bReturn = True then
							Fn_MyWorklist_PerformSignOff = False
							objParent.JavaDialog("Signoff Decision").JavaButton("Cancel").Click micLeftBtn
							objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Close Error Dialog") 		
							objParent = nothing	
							Exit Function						
					End if 
				End if	
			End If
	End If	

	'if sAction = Menu then Need to click on Close button of Perform Signoff Dialog
	If  instr(1,sComment,",")  Then
		aComments = split(sComment,",",-1,1)
		If  aComments(1)="ON" Then
			Fn_MyWorklist_PerformSignOff = true
		Else
			objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
			If Err.Number < 0 Then
					Fn_MyWorklist_PerformSignOff = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Close Perform Signoff Dialog") 		
					objParent = nothing	
					Exit Function						
			Else
					Fn_MyWorklist_PerformSignOff = true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Closed Perform Signoff Dialog")										
			End If	 
		End If
	Else
		If sAction = "Menu" Then
			wait(5)	
			objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
			If Err.Number < 0 Then
					Fn_MyWorklist_PerformSignOff = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Close Perform Signoff Dialog") 		
					objParent = nothing	
					Exit Function						
			Else
					Fn_MyWorklist_PerformSignOff = true
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Closed Perform Signoff Dialog")										
			End If	 
	End If
End If

End Function


'*********************************************************  Function perform the signoff operation *********************************************************************

'Function Name		:					Fn_MyWorkList_AssignResponsibleParty

'Description			 :		 		   Assign Responsible Party for Perform Sign-off tasks
'												 
'                                                                                                                                                                                                         
'Parameters			   :	 			sAction: Viewer/Menu
'												sNode: Worklist tree node
'												sUser: User to be clicked on to invoke dialog
'												aOrgUser: Array of Organization Users
'												sProject: Project to be selected from the list
'												aProjUser: Array of Project team members
'												sResPoolOpt: Resource Pool option  

'Return Value		   : 			 	True/False

'Examples				:			 	Call Fn_MyWorkList_AssignResponsibleParty("Viewer","Tasks to Perform:005185/A;1-mmm1 (perform-signoffs)","AutoTest1 (autotest1)","Engineering:Designer:AutoTest2 (autotest2)","","","")
'											   Call Fn_MyWorkList_AssignResponsibleParty("Menu","Tasks to Perform:005185/A;1-mmm1 (perform-signoffs)","","Engineering:Designer:AutoTest2 (autotest2)","","","")


'History:
'											Developer Name			Date				Rev. No.			Changes Done															Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Prasanna				17-Aug-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Shrikant					06-Apr-2012	       1.0					Modified code to select sNode									Koustubh
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_AssignResponsibleParty(sAction,sNode,sUser,aOrgUser,sProject,aProjUser,sResPoolOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_AssignResponsibleParty"
		On Error Resume Next

		Dim objAssignRP,arrNode,iItemCount, sExpandNode, objSearchTree, sPath, bRet

		'Select the Node from Worklist tree node
					If  sNode <> "" Then			
									If instr(1,sNode, "My Worklist")Then      
										sPath = sNode
										bRet = Fn_MyWorkList_TreeNodeOperations("Select", sPath,"")
									Else
										' added by Shrikant.
										Set objSearchTree =  JavaWindow("MyWorkListWindow").JavaTree("MyWorkListTree")
										sPath = objSearchTree.Object.getItem(0).getData().toString() & ":" & objSearchTree.Object.getItem(0).getItem(0).getData().toString() & ":" & sNode
										bRet = Fn_MyWorkList_TreeNodeOperations("Select", sPath ,"")
									End If

									If NOT(bRet) Then
												Fn_MyWorkList_AssignResponsibleParty = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Filed to Select the Node " + sPath +"  from WorkList tree") 		
												Exit Function						
									Else
												Fn_MyWorkList_AssignResponsibleParty = true
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected the Node " + sPath +"  from WorkList tree") 	
												Wait(5)
									End If
					End If


		' Select the Way to Assign Resopnsible Party
		Select Case sAction 
					Case "Viewer"
									' Select the Viewer tab
								    Call Fn_MyTc_TabSet("Viewer")
									If Err.Number < 0 Then
											Fn_MyWorkList_AssignResponsibleParty = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set Viewer tab") 		
											Exit Function						
									Else
											Fn_MyWorkList_AssignResponsibleParty = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Viewer tab")
											Wait(2)
									End If
								
									'Set the label of the sUser link
									JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaStaticText("ViwrResponsibleParty").SetToProperty "label",sUser
									JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaStaticText("ViwrResponsibleParty").Click 0,0
									If Err.Number < 0 Then
											Fn_MyWorkList_AssignResponsibleParty = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click On User Name " + sUser + " Link") 		
											Exit Function						
									Else
											Fn_MyWorkList_AssignResponsibleParty = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked On User Name " + sUser + " Link")
											Wait(2)
									End If

					Case "Menu"
									bReturn = Fn_MenuOperation("Select", "Actions:Assign...")
									Call Fn_ReadyStatusSync(5)
									If Err.Number < 0 Then
											Fn_MyWorkList_AssignResponsibleParty = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Menu Action --> Assign ") 		
											Exit Function						
									Else
											Fn_MyWorkList_AssignResponsibleParty = true
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Menu Action --> Assign")
											Wait(5)
									End If	
					Case else
									Fn_MyWorkList_AssignResponsibleParty = fasle
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Can Not Perform Action   " + sAction)									

		End Select

		'check Whether Window Exist or not 
		If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Responsible Party").Exist(10) then
		
					Set objAssignRP= JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Responsible Party")

				     ' Select the User from Organization tree.
					 If aOrgUser  <> "" Then                                                                    							
								'Expand the node
								arrNode = split(aOrgUser,":",-1,1)
								sExpandNode = "#0:"
								For iItemCount = 0 to UBound(arrNode)-1
										sExpandNode = sExpandNode + arrNode(iItemCount)
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Responsible Party").JavaTree("OrgTree").Expand sExpandNode
										wait(3)
										sExpandNode = sExpandNode + ":"
								Next
														
								'Select the Node
				       			JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Responsible Party").JavaTree("OrgTree").Select "#0:" + aOrgUser								
								If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + aOrgUser  + " from Oraganization Tree")
												Fn_MyWorkList_AssignResponsibleParty = False
												objAssignRP.JavaButton("Cancel").Click micLeftBtn
												Set objAssignRP = Nothing
												Exit Function				
								Else	
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Node " + aOrgUser  + " from Oraganization Tree")
												Fn_MyWorkList_AssignResponsibleParty = true
								End If									
					 End If

					' Select the Project Team tab	
					 If sProject  <> "" Then
								objAssignRP.JavaTab("Tab").Select "Project Teams"
								Call Fn_ReadyStatusSync(1)
								If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select tab Project Teams from Assign Responsible Party.")
												Fn_MyWorkList_AssignResponsibleParty = False
												objAssignRP.JavaButton("Cancel").Click micLeftBtn
												Set objAssignRP = Nothing
												Exit Function
								Else	
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected tab Project Teams from Assign Responsible Party.")
												Fn_MyWorkList_AssignResponsibleParty = true				
								End If							

                            	'Select the Project                              
								JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Responsible Party").JavaList("ProjectList").Select sProject
								If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Project "+ sProject + " from Assign Responsible Party.")
												Fn_MyWorkList_AssignResponsibleParty = False
												objAssignRP.JavaButton("Cancel").Click micLeftBtn
												Set objAssignRP = Nothing
												Exit Function
								Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Failed to Select Project " + sProject + " from Assign Responsible Party.")
												Fn_MyWorkList_AssignResponsibleParty = true			
								End If

								'Expand the node
								arrNode = split(aProjUser,":",-1,1)
								sExpandNode = "#0:"
								For iItemCount = 0 to UBound(arrNode)-1
										sExpandNode = sExpandNode + arrNode(iItemCount)
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Responsible Party").JavaTree("ProjectsTree").Expand sExpandNode
										wait(3)
										sExpandNode = sExpandNode + ":"
								Next

								'Select the Project Member
								If  aProjUser <> "" Then
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Responsible Party").JavaTree("ProjectsTree").Select "#0:" + aProjUser												
												If Err.Number < 0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + aProjUser  + " from Projects Tree")
																Fn_MyWorkList_AssignResponsibleParty = False
																objAssignRP.JavaButton("Cancel").Click micLeftBtn
																Set objAssignRP = Nothing
																Exit Function
												Else	
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Node " + aProjUser  + " from Projects Tree")
																Fn_MyWorkList_AssignResponsibleParty = true
												End If
													
								End If
					End if

					'Set the Resource Pool Option											
					If sResPoolOpt <> "" Then
									objAssignRP.JavaRadioButton("ResPoolOpt").SetTOProperty "attached text",sResPoolOpt
									objAssignRP.JavaRadioButton("ResPoolOpt").Set "ON"
									If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Resource Pool Option " + sResPoolOpt)
													Fn_MyWorkList_AssignResponsibleParty = False
													objAssignRP.JavaButton("Cancel").Click micLeftBtn													
													Set objAssignRP = Nothing
													Exit Function
									Else	
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected  Resource Pool Option " + sResPoolOpt)
													Fn_MyWorkList_AssignResponsibleParty = true
									End If												 
					End If
					       					 
		End if 

		JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Responsible Party").JavaButton("OK").Click
		If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button")
					Fn_MyWorkList_AssignResponsibleParty = False
					Set objAssignRP = Nothing
					Exit Function
		Else	
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Ok button")
					Fn_MyWorkList_AssignResponsibleParty = true
		End If  	
							
End Function
		

'*********************************************************  Function perform the signoff operation *********************************************************************

'Function Name		:					Fn_MyWorkList_AssignParticipant

'Description			 :		 		   Assign Responsible Party for Perform Sign-off tasks
'												 
'                                                                                                                                                                                                         
'Parameters			   :	 			sAction: Add/Remove/Modify/Verify
'												sParticipantNode: Reviewer / Responsible party 
'												aUser: Array of users to be assigned from Org Tree (':' separated fully qualified path)    
'												sProject: Project to be selected (Projects Team)  
'												aProjUser: Array of project users
'												sMemberOpt: Any/All 
'												sGroupOpt: Any/All  

'Return Value		   : 			 	True/False

'Examples				:			 	  'Add
'														bReturn = Fn_MyWorkList_AssignParticipant ("Add","Proposed Reviewers",aUser,"12345-Proj_WrkflwPreReq","Engineering - ALL:Designer:Engineering/Designer/AutoTest2","Any","Any")
'												Modify
'														bReturn = Fn_MyWorkList_AssignParticipant ("Modify","Proposed Reviewers:AutoTest1 (autotest1)-Engineering/Designer",aUse","","","All","Specific")
'												Remove
'														bReturn = Fn_MyWorkList_AssignParticipant ("Remove","Proposed Reviewers:AutoTest2 (autotest2)-Engineering/Designer","","","","","")
'												UserSeacrh
'														bReturn = Fn_MyWorkList_AssignParticipant ("UserSearch","Proposed Reviewers",aUser,"","","","")
'												RoleSeacrh
'														bReturn = Fn_MyWorkList_AssignParticipant ("RoleSearch","Proposed Reviewers","aUser,"","","","")
'												GroupSearch
'														bReturn = Fn_MyWorkList_AssignParticipant ("GroupSearch","Proposed Reviewers",aUser,"","","","")
'												Verify
'														bReturn = Fn_MyWorkList_AssignParticipant ("Verify","Proposed Reviewers:AutoTest1 (autotest1)-Engineering/Designer","","","","","")
'												VerifyWithNOduplicates
'														bReturn = Fn_MyWorkList_AssignParticipant ("VerifyWithNOduplicates","Proposed Reviewers:AutoTest1 (autotest1)-Engineering/Designer","","","","","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done																					Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Prasanna			16-Aug-2010	       1.0
'										Shweta Rathod			13-Mar-2017		   1.0 				Added case 	VerifyWithNOduplicates - to check duplicate values from the both the node				Shweta Rathod
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_AssignParticipant (sAction,sParticipantNode,aUser,sProject,aProjUser,sMemberOpt,sGroupOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_AssignParticipant"
   On Error Resume Next
   Dim arrNodeList,iItemCount,iCounter,sTreeItem,arrNode,iOuterCount,aMenuList
   Dim objAssignDialog,objOrgTree,sExpandNode,objProjectTree,sUserName,aParticipantNode
   Dim sNodeName,sHeadertxt, sArr,sNewItemMenu
   Dim jItemCount,jCnt,jCounter
	Fn_MyWorkList_AssignParticipant = true

	Set objAssignDialog =  JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants")
	sNewItemMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Workflow_Menu"),"Assign Participants")
	
	'Open the Assign Participants dialog
	If objAssignDialog.Exist(SISW_MIN_TIMEOUT) = false Then
		bReturn = Fn_MenuOperation("Select",sNewItemMenu)	  
		Call Fn_ReadyStatusSync(2)		
		If bReturn = False Then
			Fn_MyWorkList_AssignParticipant = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  Select Menu [Tools:Assign Participants...]")
			Set objAssignDialog  = Nothing
			Exit Function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Tools:Assign Participants...]")
		End If
	End If
	
	'Error hadled for Only One revision handled dialog
	JavaWindow("MyTeamcenter").JavaWindow("Error").SetTOProperty "title","Only one revision allowed"
	If Fn_SISW_UI_Object_Operations("Fn_MyWorkList_AssignParticipant","Exist",JavaWindow("MyTeamcenter").JavaWindow("Error"),"2") Then
''		JavaWindow("MyTeamcenter").JavaWindow("Error").JavaButton("OK").Click micLeftBtn
		Call Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_AssignParticipant", "Fn_Button_Click", JavaWindow("MyTeamcenter").JavaWindow("Error"), "OK")
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on OK button of Only one revision allowed dialog")
			Fn_MyWorkList_AssignParticipant = False
			Set objAssignDialog = Nothing
			Exit Function
		End If
		Fn_MyWorkList_AssignParticipant = False			
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Open Menu [Tools:Assign Participants...] Due to MultiSelect.")
		Set objAssignDialog = Nothing
		Exit Function
	End If	

	Select Case sAction

		Case "Add" 

					'Select the node in Participants tree
					If  sParticipantNode <> "" Then
								If Instr(1,sParticipantNode,"~") > 0 Then
									aParticipantNode = Split(sParticipantNode,"~",-1,1)
									sParticipantNode = aParticipantNode(0)
								End If

								'Expand the Participant tree node 
								arrNode = split(sParticipantNode,":",-1,1)
								sExpandNode = "#0:"
								For iItemCount = 0 to UBound(arrNode)-1
										sExpandNode = sExpandNode + arrNode(iItemCount)
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Expand sExpandNode
										wait(1)
										sExpandNode = sExpandNode + ":"
								Next

								'Select the Participants node
								JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Select "#0:" + sParticipantNode 	
								If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Participnats " + sParticipantNode  + " from Assign participants dialog")
										Fn_MyWorkList_AssignParticipant = False
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
										Set objAssignDialog = Nothing
										Exit Function
								Else
										Wait(5)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Participnats " + sParticipantNode  + " from Assign participants dialog")
								End If
					End If 

					'Add the User from Organization tree
					If isarray(aUser)  Then
'						arrNodeList = Array(aUser)
'						If instr(1,aUser ,",", 1) Then
'							arrNodeList = split(aUser, ",",-1,1)		
'						else
'							arrNodeList = Array(aUser)
'						End If							
							objOrgTree = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree")
							For iCounter = 0 to UBound(aUser)
										arrNode = split(aUser(iCounter),":",-1,1)
										sExpandNode = "#0:"
										For iItemCount = 0 to UBound(arrNode)-1
												sExpandNode = sExpandNode + arrNode(iItemCount)
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Expand sExpandNode
												wait(5)
												sExpandNode = sExpandNode + ":"
										Next					
										
'										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Object.clearSelection	
										'JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").ExtendSelect sExpandNode + arrNode(iItemCount)
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").ExtendSelect sExpandNode + arrNode(iItemCount)
                                  
										Wait(5)
							Next
'							JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Object.clearSelection
'							For iCounter = 0 to UBound(arrNodeList)
'										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").ExtendSelect sExpandNode + arrNode(iCounter)
'										Wait(3)
			'							If Err.Number < 0 Then
			'									Fn_MyWorkList_Org_TreeNodeOperations = False
			'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Mulitiselect node   " + arrNodeList(iCounter) + "of Oraganization Tree." )	
			'									Set objOrgTree = Nothing
			'									Exit Function 							
					'					End If
	
'							Next							
										If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node " +  arrNode(iCounter)   + " from Oraganization Tree")
										Fn_MyWorkList_AssignParticipant = False
										objAssignDialog.JavaButton("Close").Click micLeftBtn
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Close Button")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Close Button")
										End if
										Set objAssignDialog = Nothing
										Exit Function
							 Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected node   " +  arrNode(iCounter) + " of Oraganization Tree." )								
							End If
'					End if

					Else
						If instr(1,aUser ,",", 1) Then
							arrNodeList = split(aUser, ",",-1,1)
						End if

						objOrgTree = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree")
							For iCounter = 0 to UBound(arrNodeList)
										arrNode = split(arrNodeList(iCounter),":",-1,1)
										sExpandNode = "#0:"
										For iItemCount = 0 to UBound(arrNode)-1
												sExpandNode = sExpandNode + arrNode(iItemCount)
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Expand sExpandNode
												wait(1)
												sExpandNode = sExpandNode + ":"
										Next						
							Next
							JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Object.clearSelection
							For iCounter = 0 to UBound(arrNodeList)
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").ExtendSelect "#0:" + arrNodeList(iCounter)
										Wait(3)
			'							If Err.Number < 0 Then
			'									Fn_MyWorkList_Org_TreeNodeOperations = False
			'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Mulitiselect node   " + arrNodeList(iCounter) + "of Oraganization Tree." )	
			'									Set objOrgTree = Nothing
			'									Exit Function 							
					'					End If
	
							Next							
												If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node " + arrNodeList(iCounter)   + " from Oraganization Tree")
										Fn_MyWorkList_AssignParticipant = False
										objAssignDialog.JavaButton("Close").Click micLeftBtn
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Close Button")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Close Button")
										End if
										Set objAssignDialog = Nothing
										Exit Function
							 Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected node   " + arrNodeList(iCounter) + " of Oraganization Tree." )								
							End If
					End if
'							If Err.Number < 0 Then
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node " + arrNodeList(iCounter)   + " from Oraganization Tree")
'										Fn_MyWorkList_AssignParticipant = False
'										objAssignDialog.JavaButton("Close").Click micLeftBtn
'										If Err.Number < 0 Then
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Close Button")
'												Fn_MyWorkList_AssignParticipant = False
'												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
'												Set objAssignDialog = Nothing
'												Exit Function
'										Else
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Close Button")
'										End if
'										Set objAssignDialog = Nothing
'										Exit Function
'							 Else
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected node   " + arrNodeList(iCounter) + " of Oraganization Tree." )								
'							End If

							If IsArray(aParticipantNode) Then
										sHeadertxt =JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaStaticText("HeaderText").GetROProperty("attached text")
										If Trim(aParticipantNode(1)) = Trim(sHeadertxt) Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verify header text.")
											Fn_MyWorkList_AssignParticipant = True
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify header text.")
											Fn_MyWorkList_AssignParticipant = False
											JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
											Set objAssignDialog = Nothing
											Exit Function
										End If
							End If
	
'					End If
	
					If sProject <> "" Then
							' Select the Project Teams Tab
							JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTab("Tab").Select "Project Teams"

							' Select the Project 
							JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaList("ProjectsList").Select  sProject

							If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Project " + sProject   + " from Project List")
									Fn_MyWorkList_AssignParticipant = False
									JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
									Set objAssignDialog = Nothing
									Exit Function
							Else
									Call Fn_ReadyStatusSync(2)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Project " + sProject   + " from Project List")				
							End If						

							objProjectTree = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree")

							'arrNodeList = split(aProjUser, ",",-1,1)		
							'[TC1123-20160915a00-27_09_2016-VivekA-Maintenance] - As per Akshay's mail, and Design change
							If aProjUser<>""  Then
								If Instr(aProjUser(0),"/")>0 Then
									sArr = split(aProjUser(0),"/")
									aProjUser(0) = aProjUser(0) +" "+"("+lcase(sArr(2))+")"
								End If
							End If
							'--------------------------------------------------
							arrNodeList = aProjUser				
							For iCounter = 0 to UBound(arrNodeList)
									arrNode = split(arrNodeList(iCounter),":",-1,1)
									sExpandNode = "#0:"
									For iItemCount = 0 to UBound(arrNode)-1
											sExpandNode = sExpandNode + arrNode(iItemCount)
											JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree").Expand sExpandNode
											Call Fn_ReadyStatusSync(2)
											sExpandNode = sExpandNode + ":"
									Next						
							Next
							JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree").Object.clearSelection
							Call Fn_ReadyStatusSync(2)
							For iCounter = 0 to UBound(arrNodeList)
									If iCounter = 0 Then
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree").Select "#0:" + arrNodeList(iCounter)
									Else
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree").ExtendSelect "#0:" + arrNodeList(iCounter)
									End If
									Call Fn_ReadyStatusSync(2)
							Next
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Multiselected node   " + arrNodeList(iCounter) + " of Project Teams Tree." )   

							If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node " + arrNodeList(iCounter)   + " from Project User Tree")
									Fn_MyWorkList_AssignParticipant = False
									JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
									Set objAssignDialog = Nothing
									Exit Function				
							End If
					End If

					'Select the group
					If sGroupOpt <>"" Then
							Select Case sGroupOpt
										Case "Any"
											JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").SetTOProperty "attached text","Any Group"
											JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").Set "ON"                                                      
										Case "Specific"
											JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").SetTOProperty "attached text","Specific Group"
											JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").Set "ON" 														
							End Select
							If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Group Type" +   sGroupOpt)
									Fn_MyWorkList_AssignParticipant = False
									JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn 
									Set objAssignDialog = Nothing
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Group Type" +   sGroupOpt)				
							End If
					End If

							' Add the Member  
					If sMemberOpt <>"" Then
									Select Case sMemberOpt
													Case "Any"
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").SetTOProperty "attached text","Any Member"
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").Set "ON"                                                      							
													Case "All"
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").SetTOProperty "attached text","All Members"
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").Set "ON"                                                      														
										End Select

										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Member Type" +   sMemberOpt)
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Member Type" +   sMemberOpt)				
										End If
					End If                            					
								
					JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Add").Click micLeftBtn
					If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Add the Project User")
									Fn_MyWorkList_AssignParticipant = False
									JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
									Set objAssignDialog = Nothing
									Exit Function
					End if					
					'End If

			Case "Modify"	
					If sParticipantNode <> "" Then
										'Expand the Participant tree node 
										arrNode = split(sParticipantNode,":",-1,1)
										sExpandNode = "#0:"
										For iItemCount = 0 to UBound(arrNode)-1
												sExpandNode = sExpandNode + arrNode(iItemCount)
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Expand sExpandNode
												wait(1)
												sExpandNode = sExpandNode + ":"
										Next						
			
										'Select the node from Participant tree
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Select "#0:" + sParticipantNode 	
										wait(3)
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Participnats " + sParticipantNode  + " from Assign participants dialog")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Participnats " + sParticipantNode  + " from Assign participants dialog")									
										End If
					Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Participnat has been Selected")
										Fn_MyWorkList_AssignParticipant = False
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
										Set objAssignDialog = Nothing
										Exit Function
					End If
					If aUser <>	"" Then
							arrNodeList = split(aUser, ",",-1,1)
							If Ubound(arrNodeList) > 0  Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Wrong Input For Organisation Members.")
										Fn_MyWorkList_AssignParticipant = False
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
										Set objAssignDialog = Nothing
										Exit Function
							End If
															
							objOrgTree = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree")
							For iCounter = 0 to UBound(arrNodeList)
										arrNode = split(arrNodeList(iCounter),":",-1,1)
										sExpandNode = "#0:"
										For iItemCount = 0 to UBound(arrNode)-1
												sExpandNode = sExpandNode + arrNode(iItemCount)
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Expand sExpandNode
												wait(3)
												sExpandNode = sExpandNode + ":"
										Next						
							Next
							JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Object.clearSelection
							For iCounter = 0 to UBound(arrNodeList)
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").ExtendSelect "#0:" + arrNodeList(iCounter)
										wait(1)
			'							If Err.Number < 0 Then
			'									Fn_MyWorkList_Org_TreeNodeOperations = False
			'									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Mulitiselect node   " + arrNodeList(iCounter) + "of Oraganization Tree." )	
			'									Set objOrgTree = Nothing
			'									Exit Function 							
					'					End If
	
							Next
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected node   " + arrNodeList(iCounter) + " of Oraganization Tree." )	
	
							If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node " + arrNodeList(iCounter)   + " from Oraganization Tree")
										Fn_MyWorkList_AssignParticipant = False
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
										Set objAssignDialog = Nothing
										Exit Function				
							End If
	
							JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Modify").Click micLeftBtn
							If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Modify the User From Oraganization")
													Fn_MyWorkList_AssignParticipant = False
													JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
													Set objAssignDialog = Nothing
													Exit Function
							End if
	
						End If
	
						If sProject <> "" Then
								' Select the Project Teams Tab
								JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTab("Tab").Select "Project Teams"
	
								' Select the Project 
								JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaList("ProjectsList").Select  sProject
	
								If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Project " + sProject   + " from Project List")
										Fn_MyWorkList_AssignParticipant = False
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
										Set objAssignDialog = Nothing
										Exit Function				
								End If						
	
								objProjectTree = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree")
	
								'arrNodeList = split(aProjUser, ",",-1,1)								
								arrNodeList  = 	aProjUser
								If Ubound(arrNodeList) > 0  Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Wrong Input For Project Team Members.")
										Fn_MyWorkList_AssignParticipant = False
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
										Set objAssignDialog = Nothing
										Exit Function
								End If

								For iCounter = 0 to UBound(arrNodeList)
										arrNode = split(arrNodeList(iCounter),":",-1,1)
										sExpandNode = "#0:"
										For iItemCount = 0 to UBound(arrNode)-1
												sExpandNode = sExpandNode + arrNode(iItemCount)
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree").Expand sExpandNode
												wait(1)
												sExpandNode = sExpandNode + ":"
										Next						
								Next
								JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree").Object.clearSelection
								For iCounter = 0 to UBound(arrNodeList)
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ProjectsTree").ExtendSelect "#0:" + arrNodeList(iCounter)
										wait(1)		
								Next
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Multiselected node   " + arrNodeList(iCounter) + " of Project Teams Tree." )   
	
								If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node " + arrNodeList(iCounter)   + " from Project User Tree")
										Fn_MyWorkList_AssignParticipant = False
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
										Set objAssignDialog = Nothing
										Exit Function				
								End If

								'Select the Group option
                                If sGroupOpt <>"" Then
										Select Case sGroupOpt
														Case "Any"
															JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").SetTOProperty "attached text","Any Group"
															JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").Set "ON"                                                      
														Case "Specific"
															JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").SetTOProperty "attached text","Specific Group"
															JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").Set "ON" 														
										End Select
								End If
								If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Group Type" +   sGroupOpt)
													Fn_MyWorkList_AssignParticipant = False
													JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
													Set objAssignDialog = Nothing
													Exit Function
								Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Group Type" +   sGroupOpt)				
								End If

								' Add the member from Project Teams								
								If sMemberOpt <>"" Then
										Select Case sMemberOpt
													Case "Any"
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").SetTOProperty "attached text","Any Member"
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").Set "ON"                                                      							
													Case "All"
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").SetTOProperty "attached text","All Members"
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaRadioButton("MemberOpt").Set "ON"                                                      														
											End Select
	
											If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Member Type" +   sMemberOpt)
													Fn_MyWorkList_AssignParticipant = False
													JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
													Set objAssignDialog = Nothing
													Exit Function
											Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Member Type" +   sMemberOpt)				
											End If
								End If
	
								
								JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Modify").Click micLeftBtn
								If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Modify the Project User")
													Fn_MyWorkList_AssignParticipant = False
													JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
													Set objAssignDialog = Nothing
													Exit Function
								End if
								
					End If

			Case "Remove"                   					
							If sParticipantNode <> "" Then
										'Expand the Participant tree node 
										arrNode = split(sParticipantNode,":",-1,1)
										sExpandNode = "#0:"
										For iItemCount = 0 to UBound(arrNode)-1
												sExpandNode = sExpandNode + arrNode(iItemCount)
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Expand sExpandNode
												wait(1)
												sExpandNode = sExpandNode + ":"
										Next						
			
										'Select the node from Participant tree
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Select "#0:" + sParticipantNode 	
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Participnats " + sParticipantNode  + " from Assign participants dialog")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Participnats " + sParticipantNode  + " from Assign participants dialog")									
										End If
	
										'Click on Remove Button
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Remove").Click micLeftBtn
										If Err.Number < 0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Remove the User From Participants")
																Fn_MyWorkList_AssignParticipant = False
																JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
																Set objAssignDialog = Nothing
																Exit Function
										End If
							End if

	
			Case "UserSearch" 
									If sParticipantNode <> "" Then
												'Expand the Participant tree node 
												arrNode = split(sParticipantNode,":",-1,1)
												sExpandNode = "#0:"
												For iItemCount = 0 to UBound(arrNode)-1
														sExpandNode = sExpandNode + arrNode(iItemCount)
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Expand sExpandNode
														wait(1)
														sExpandNode = sExpandNode + ":"
												Next						
					
												'Select the node from Participant tree
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Select "#0:" + sParticipantNode 	
												wait(2)
												If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Participnats " + sParticipantNode  + " from Assign participants dialog")
														Fn_MyWorkList_AssignParticipant = False
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
														Set objAssignDialog = Nothing
														Exit Function
												Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Participnats " + sParticipantNode  + " from Assign participants dialog")									
												End If
									Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Participnat has been Selected")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
							End If
							If  aUser <> "" Then
										 arrNodeList = split(aUser,":",-1,1)
										 sUserName =arrNodeList(Ubound(arrNodeList))
										If instr(1,sUserName ,"(", 1) Then
												sUserName = Mid(sUserName,instr(1,sUserName,"(", 1),len(sUserName))
												sUserName = Replace(sUserName,"(","") 
												If instr(1,sUserName,")", 1) Then
														sUserName = Replace(sUserName,")","") 
												End If
										End If
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaEdit("FindUser").Set sUserName 
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Search").Click micLeftBtn
		
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on User Search Button")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										End If
										Wait(2)
'										If JavaWindow("My WorkList - Teamcenter").JavaWindow("Search").Exists Then
'												JavaWindow("My WorkList - Teamcenter").JavaWindow("Search").JavaButton("OK").Click micLeftBtn
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No User Found")
'												Fn_MyWorkList_AssignParticipant = False
'												Set objAssignDialog = Nothing
'												Exit Function
'										End If
		
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Select "#0:" + aUser
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select User from Organisation tree")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										End If
'										Wait(2)
'										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Add").Click micLeftBtn
'										If Err.Number < 0 Then
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Add the User From Oraganization tree")
'												Fn_MyWorkList_AssignParticipant = False
'												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
'												Set objAssignDialog = Nothing
'												Exit Function
'										End if
							End if			

			Case "RoleSearch"
									If sParticipantNode <> "" Then
												'Expand the Participant tree node 
												arrNode = split(sParticipantNode,":",-1,1)
												sExpandNode = "#0:"
												For iItemCount = 0 to UBound(arrNode)-1
														sExpandNode = sExpandNode + arrNode(iItemCount)
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Expand sExpandNode
														wait(1)
														sExpandNode = sExpandNode + ":"
												Next						
					
												'Select the node from Participant tree
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Select "#0:" + sParticipantNode 	
												wait(2)
												If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Participnats " + sParticipantNode  + " from Assign participants dialog")
														Fn_MyWorkList_AssignParticipant = False
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
														Set objAssignDialog = Nothing
														Exit Function
												Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Participnats " + sParticipantNode  + " from Assign participants dialog")									
												End If
							Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Participnat has been Selected")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
							End If
							If  aUser <> "" Then
										 arrNodeList = split(aUser,":",-1,1)
										 sUserName =arrNodeList(Ubound(arrNodeList))
										If instr(1,sUserName ,"(", 1) Then
												sUserName = Mid(sUserName,instr(1,sUserName,"("),len(sUserName), 1)
												sUserName = Replace(sUserName,"(","") 
												If instr(0,sUserName,")") Then
														sUserName = Replace(sUserName,")","") 
												End If
										End If
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaEdit("FindRole").Set sUserName 
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Search").Click micLeftBtn
		
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Role Search Button")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										End If
										Wait(2)
'										If JavaWindow("My WorkList - Teamcenter").JavaWindow("Search").Exists Then
'												JavaWindow("My WorkList - Teamcenter").JavaWindow("Search").JavaButton("OK").Click micLeftBtn
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No User Found")
'												Fn_MyWorkList_AssignParticipant = False
'												Set objAssignDialog = Nothing
'												Exit Function
'										End If
		
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Select "#0:" + aUser
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Role from Organisation tree")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										End If

'										Wait(2)
'										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Add").Click micLeftBtn
'										If Err.Number < 0 Then
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Add the Role From Oraganization tree")
'												Fn_MyWorkList_AssignParticipant = False
'												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
'												Set objAssignDialog = Nothing
'												Exit Function
'										End if
							End if			

			Case "GroupSearch" 
									If sParticipantNode <> "" Then
												'Expand the Participant tree node 
												arrNode = split(sParticipantNode,":",-1,1)
												sExpandNode = "#0:"
												For iItemCount = 0 to UBound(arrNode)-1
														sExpandNode = sExpandNode + arrNode(iItemCount)
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Expand sExpandNode
														wait(1)
														sExpandNode = sExpandNode + ":"
												Next						

												'Select the node from Participant tree
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Select "#0:" + sParticipantNode 	
												wait(2)
												If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Participnats " + sParticipantNode  + " from Assign participants dialog")
														Fn_MyWorkList_AssignParticipant = False
														JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
														Set objAssignDialog = Nothing
														Exit Function
												Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Participnats " + sParticipantNode  + " from Assign participants dialog")									
												End If
							Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Participnat has been Selected")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
							End If
							If  Not IsEmpty(aUser) Then
										 arrNodeList = split(aUser,":",-1,1)
										 sUserName =arrNodeList(Ubound(arrNodeList))
										If instr(1,sUserName ,"(", 1) Then
												sUserName = Mid(sUserName,instr(1,sUserName,"("),len(sUserName), 1)
												sUserName = Replace(sUserName,"(","") 
												If instr(0,sUserName,")") Then
														sUserName = Replace(sUserName,")","") 
												End If
										End If
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaEdit("FindGroup").Set sUserName 
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Search").Click micLeftBtn
		
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Group Search Button")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										End If
										Wait(2)
										If JavaWindow("MyWorkListWindow").JavaWindow("Search").Exist(3) Then
												JavaWindow("MyWorkListWindow").JavaWindow("Search").JavaButton("OK").Click micLeftBtn
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No User Found")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										End If
		
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Select "#0:" + aUser
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Group from Organisation tree")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										End If

'										Wait(2)
'										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Add").Click micLeftBtn
'										If Err.Number < 0 Then
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Add the Group From Oraganization tree")
'												Fn_MyWorkList_AssignParticipant = False
'												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
'												Set objAssignDialog = Nothing
'												Exit Function
'										End if
							End if                      

			Case "VerifyOrgNodeExist"

							If  Not IsEmpty(aUser) Then

										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("OrganizationTree").Select "#0:" + aUser
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Group ["+CStr(aUser)+"] from Organisation tree.")
												Fn_MyWorkList_AssignParticipant = False
												JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
												Set objAssignDialog = Nothing
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: To Select Group ["+CStr(aUser)+"] from Organisation tree.")
												Fn_MyWorkList_AssignParticipant = True
										End If

							End If

			Case "Verify","VerifyWithNOduplicates","VerifyWithoutClose"
						If  sParticipantNode <> "" Then

								'Expand the All the nodes first
                        		arrNode = split(sParticipantNode,":",-1,1)
								sExpandNode = "#0:"
								For iItemCount = 0 to UBound(arrNode)-1
										sExpandNode = sExpandNode + arrNode(iItemCount)
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").Expand sExpandNode
										wait(1)
										sExpandNode = sExpandNode + ":"
								Next
								
								'Get the count of Tree items & compare input string with each parameter 
								iItemCount = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").GetROProperty( "items count")
								sNodeName = "Participants:"+sParticipantNode
								For iCounter=0 To (iItemCount-1)
									sTreeItem = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").GetItem(iCounter)
									
									If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
										Fn_MyWorkList_AssignParticipant = True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified Participnats " + sParticipantNode  + " from Assign participants dialog")
										Exit For
									End If
								Next   							
								'--------------------- Start - Added by ShwetaR:NewDev:WorkFlow:09Mar17 - added code to check duplicate values from the both the node ------------------------------------
								If sAction = "VerifyWithNOduplicates" then 
									jItemCount = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").GetROProperty( "items count")
									jCnt = 0
									For jCounter=0 To (jItemCount-1)
										sTreeItem = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaTree("ParticipantsTree").GetItem(jCounter)
										If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
											jCnt = jCnt + 1
											If jCnt > 1 then
												Fn_MyWorkList_AssignParticipant = false
												Exit function
											End if
										End if
									next
									
									if jCnt = 1 then
										Fn_MyWorkList_AssignParticipant = True
									else
										Fn_MyWorkList_AssignParticipant = false
										Exit function
									End if
								End if
								' ---------------------End of - Added by ShwetaR:NewDev:WorkFlow:09Mar17 - added code to check duplicate values from the both the node --------------------------------
								If  Cint(iCounter) = Cint (iItemCount) Then
										Fn_MyWorkList_AssignParticipant = False
										JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("Close").Click micLeftBtn
											If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on Close Button")
													Fn_MyWorkList_AssignParticipant = False
													Set objAssignDialog = Nothing
													Exit Function
											Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on Close Button")
											End if
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify Participnats " + sParticipantNode  + " from Assign participants dialog")
										Set objAssignDialog = Nothing
										Exit Function 
								End If
						End If
	End Select

	If Trim(sAction) <> "UserSearch" AND  Trim(sAction) <> "RoleSearch" AND  Trim(sAction) <> "GroupSearch" and Trim(sAction) <> "VerifyOrgNodeExist" and trim(sAction) <> "VerifyWithNOduplicates" and trim(sAction) <> "VerifyWithoutClose" Then	
				JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Assign Participants").JavaButton("OK").Click micLeftBtn
				If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on OK Button")
						Fn_MyWorkList_AssignParticipant = False
						Set objAssignDialog = Nothing
						Exit Function
				End if
	End If
End Function 

'*********************************************************  Function perform the signoff team select operation *********************************************************************

'Function Name		:					Fn_MyWorkList_SignoffTeamSelect

'Description			 :		 		   This function select sign off team.												 
'                                                                                                                                                                                                         
'Parameters			   :	 			sAction: Add/Remove/Modify/Verify
'												dicSelectSignOff : Refer DictionaryDeclaration.vbs for the defination & keys included

'Return Value		   : 			 	True/False

'Examples				:			 	 dicSelectSignOff("WorkListTreeNode") = Node from Worklist Tree
'												dicSelectSignOff("SignOffTeamSelect") = "Profiles:Engineering/Designer/1" ' To be used in case of Profiles Case
'												dicSelectSignOff("UsersName") = "Engineering:Designer:AutoTest2 (autotest2)"
'												dicSelectSignOff("ProjectName") = "12345-Proj_WrkflwPreReq"
'												dicSelectSignOff("ProjectUsers") = "Designer:Engineering/Designer/AutoTest1"
'												dicSelectSignOff("MemberOption") = "Any"
'												dicSelectSignOff("GroupOption") = "Any"
'												dicSelectSignOff("Quorum") = "Nuemeric:10"
'												dicSelectSignOff("Wait") = true
'												dicSelectSignOff("Adhoc") = true
'												dicSelectSignOff("ProcessDescription") = "Process Desc New"
'												dicSelectSignOff("Comments") = "comments for test"
'
'												Fn_MyWorkList_SignoffTeamSelect("Users",dicSelectSignOff)
'												Fn_MyWorkList_SignoffTeamSelect("Profiles",dicSelectSignOff)
'												--------------------------------------------------------------------------------------------
'												dicSelectSignOff("Required") = "VerifyRequiredCheckBox"
'												dicSelectSignOff("Required") = "SetRequiredCheckBox"
'												dicSelectSignOff("VerifyAction") = "test123"
'												--------------------------------------------------------------------------------------------
'												dicSelectSignOff("VerifyStaticText") = "test123"
'												dicSelectSignOff("VerifyProcessDesc") = "test123"
'												
'History:
'										Developer Name			Date				Rev. No.				Reviewer						Changes Done
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										 Prasanna				16-Aug-2010	        1.0
'										Prasanna				07-Sep-2010			1.0												         Added profiles Case	
'										Mahendra				14-Sep-2010																Added menu Case			
'										Chaitali Rane			13-Mar-2017		    1.0						Shweta Rathod				Added subcase Case "Required" ,"VerifyAction"			
'										shweta Rathod			06-APR-2017			1.0                     Shweta Rathod				Added Case "VerifyStaticText" - to  Verify static text
'										shweta Rathod			16-NOV-2017			1.0                     Shweta Rathod				Added Case "VerifyProcessDesc" - to  Verify process description value
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_MyWorkList_SignoffTeamSelect(sAction,dicSelectSignOff)
	
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_SignoffTeamSelect"
	On Error Resume Next
	Dim dicCount , dicKeys , dicItems , iCounter ,bReturn,arrQuorum,arrNodeList,arrNode,sExpnadNode,aSplitNodeName
	Dim iNodeCounter,iCounter1,sActionSelect,iCnt,iAll
	Dim objSelectType,intNoOfObjects, objDialog, iCounter2, sListHierarchy, sNodeName, sTempNode, sCase

	dicCount  = dicSelectSignOff.Count
	dicItems = dicSelectSignOff.Items
	dicKeys = dicSelectSignOff.Keys

	sActionSelect = sAction
	Fn_MyWorkList_SignoffTeamSelect = true
	sCase = ""
    Select Case sAction         

			Case "Users"
				For iCounter = 0 to dicCount - 1
	                    If  dicItems(iCounter) <> "" Then
	                            Select Case dicKeys(iCounter)

								 Case "WorkListTreeNode"	

											'Select the WorkList Tree Node		
											bReturn =  Fn_MyWorkList_TreeNodeOperations("Select",dicItems(iCounter),"")
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + dicItems(iCounter) +  " From WorkList Tree")
													Exit Function
											End If
											Call Fn_ReadyStatusSync(5)
											
											'Set the Viewer Tab
											bReturn =  Fn_MyTc_TabSet("Viewer")	
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Viewer Tab")
													Exit Function
											End If
											
											Call Fn_ReadyStatusSync(10)
											wait(5)
											 'Set Default View to Task View 
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ViewOptions").SetTOProperty "Attached Text","Task View"
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ViewOptions").Set "ON"
											If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Task View from Viewer tab")
													Fn_MyWorkList_SignoffTeamSelect = False									
													Exit Function
											Else																						
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Task View from Viewer tab")
													Call Fn_ReadyStatusSync(5)
													wait(2)
											End If

											 ' Select the Node in Signoff Team tree											 
											 If  instr(1,dicSelectSignOff.Item("SignOffTeamSelect"),"#0") > 0 Then
														If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("SignOffTeamTree").Exist Then
																JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("SignOffTeamTree").Select dicSelectSignOff.Item("SignOffTeamSelect")
														Else		
																JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("RouteSignOffTeamTree").Select dicSelectSignOff.Item("SignOffTeamSelect")
														End If
														If Err.Number < 0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Sign off Team tree node")
																Fn_MyWorkList_SignoffTeamSelect = False									
																Exit Function
														Else																						
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Sign off Team tree node")
														End If
											 Else
														bReturn =  Fn_MyWorkList_SignoffTeam_TreeNodeOperations("Select",sAction)          
														If bReturn = false Then
																Fn_MyWorkList_SignoffTeamSelect = False									
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select "+ sAction +" from Signoff Team Tree")
																Exit Function
														End If
											 End If
											Call Fn_ReadyStatusSync(5)

								  Case "UsersName"                  ' Select the nodes from Organization tree

											 ' Select the Organization Tab
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTab("OrgProjTeamTab").Select "Organization"
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Organization Tab")
													Exit Function
											Else
													Fn_MyWorkList_SignoffTeamSelect = true									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Organization Tab")
											End If

											Call Fn_ReadyStatusSync(6)

											'Function Called to Select Organization Tree Node	
											bReturn = Fn_MyWorkList_Org_TreeNodeOperations("MultiSelect", dicItems(iCounter))
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Organization Tree Node " + dicItems(iCounter))
													Exit Function
											End If	
											
											'Case Member Option
											If dicKeys("MemberOption") <> "" Then															
															Select Case dicSelectSignOff.Item("MemberOption")
																	Case "Any"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").SetTOProperty "attached text","Any Member"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").Set "ON"                                                      							
																	Case "All"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").SetTOProperty "attached text","All Members"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").Set "ON"                                                      														
															End Select

															If Err.Number < 0 Then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Member Type " +  dicSelectSignOff.Item("MemberOption") + " Member")
																	Fn_MyWorkList_SignoffTeamSelect = False
																	
																	Exit Function
															Else
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Member Type " +   dicSelectSignOff.Item("MemberOption") + " Member")				
															End If			
											End If

											'Select Group Option	
											If dicKeys("GroupOption") <> "" Then
															Select Case dicSelectSignOff.Item("GroupOption")
																	Case "Any"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Any Group"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").Set "ON"                                                      
																	Case "Specific"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Specific Group"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").Set "ON" 														
															End Select
														
															If Err.Number < 0 Then
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group Type" +   dicSelectSignOff.Item("GroupOption") + " Group")
																			Fn_MyWorkList_SignoffTeamSelect = False
																			
																			Exit Function
															Else
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group Type" +   dicSelectSignOff.Item("GroupOption") + " Group")				
															End If		
											End If

									Call Fn_ReadyStatusSync(5)
									'Select Action from the list
															If dicSelectSignOff.Item("Action") <> "" Then										
																			Select Case dicSelectSignOff.Item("Action")
						
																							Case "Review"
																							JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaList("Action").Select "Review"
																					
																							Case "Acknow"
																							JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaList("Action").Select "Acknow"
																										
																							Case "Notify"
																							JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaList("Action").Select "Notify"
																						
																				End Select
															End If
        								
															If Err.Number < 0 Then
																	Fn_MyWorkList_SignoffTeamSelect = False									
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Action from the Java  list")
																	Exit Function
															Else
																	Fn_MyWorkList_SignoffTeamSelect = true									
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Action from the Java list")
															End If

											Call Fn_ReadyStatusSync(10) 
											Wait 3
											'Click on Add button
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Add").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Add User " + dicItems(iCounter)   + " from Organization Tree")
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Added User " + dicItems(iCounter)   + " from Organization Tree")				
											End If	

											Call Fn_ReadyStatusSync(5)
											Wait 3
									Case "ProjectName"				
											 ' Select the Project Teams Tab
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTab("OrgProjTeamTab").Select "Project Teams"
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Project Teams Tab")
													Exit Function
											Else																					
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Project Teams Tab")
											End If

											' Select the Project 
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaList("Projects").Select  dicItems(iCounter)	
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Project " + dicItems(iCounter)   + " from Project List")
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Project " + dicItems(iCounter)   + " from Project List")				
											End If                                          

											Call Fn_ReadyStatusSync(6)

									Case "ProjectUsers"
											' Select the Node in Signoff Team tree
											bReturn = Fn_MyWorkList_SignoffTeam_TreeNodeOperations("Select",sActionSelect)          
											If bReturn = false Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Signoff Team tree Node" + sActionSelect)
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Signoff Team tree Node " + sActionSelect)				
										
											End If
											Call Fn_ReadyStatusSync(5)

											'Expand the node	
											arrNodeList = split(dicItems(iCounter), ",",-1,1)								
											For iNodeCounter = 0 to UBound(arrNodeList)
													arrNode = split(arrNodeList(iNodeCounter),":",-1,1)
													'sExpandNode = "#0:"
													For iItemCount = 0 to UBound(arrNode)-1
															sExpandNode = sExpandNode + arrNode(iItemCount)
															JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProjectsTree").Expand sExpandNode
															wait(1)
															sExpandNode = sExpandNode + ":"
													Next						
											Next
											Call Fn_ReadyStatusSync(5)
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProjectsTree").Object.clearSelection
											For iNodeCounter = 0 to UBound(arrNodeList)
													JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProjectsTree").ExtendSelect "#0:" + arrNodeList(iNodeCounter)
													wait(1)		
											Next
											
											If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + arrNodeList(iNodeCounter)   + " from Project User Tree")
													Fn_MyWorkList_SignoffTeamSelect = False                                                 
													Exit Function				
											Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Multiselected node   " + arrNodeList(iNodeCounter) + " of Project Teams Tree." )   
											End If

											Call Fn_ReadyStatusSync(5)
																						
											' Member Option
											If dicKeys("MemberOption") <> "" Then															
															Select Case dicSelectSignOff.Item("MemberOption")
																	Case "Any"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").SetTOProperty "attached text","Any Member"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").Set "ON"                                                      							
																	Case "All"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").SetTOProperty "attached text","All Members"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").Set "ON"                                                      														
															End Select
							
															If Err.Number < 0 Then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Member Type " + dicSelectSignOff.Item("MemberOption") + " Member")
																	Fn_MyWorkList_SignoffTeamSelect = False
																	
																	Exit Function
															Else
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Member Type " + dicSelectSignOff.Item("MemberOption") + " Member")				
															End If			
											End If

											'Select Group Option	
											If dicKeys("GroupOption") <> "" Then
															Select Case dicSelectSignOff.Item("GroupOption")
																	Case "Any"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Any Group"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").Set "ON"                                                      
																	Case "Specific"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Specific Group"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").Set "ON" 														
															End Select
														
															If Err.Number < 0 Then
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group Type" +  dicSelectSignOff.Item("GroupOption") + " Group")
																			Fn_MyWorkList_SignoffTeamSelect = False
																			
																			Exit Function
															Else
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group Type" +  dicSelectSignOff.Item("GroupOption") + " Group")				
															End If		
											 End If

											Call Fn_ReadyStatusSync(8)
											'Click on Add button
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Add").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Add User to Project " + dicItems(iCounter)   + " from Organization Tree")
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Added User to Project " + dicItems(iCounter)   + " from Organization Tree")				
											End If	

											Call Fn_ReadyStatusSync(5)

										Case "Quorum"
											If  dicItems(iCounter) <> "" Then
														arrQuorum = split(dicItems(iCounter),":",-1,1) 
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ReviewQuorumOption").SetTOProperty "attached text",arrQuorum(0) 
														wait(1)
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ReviewQuorumOption").Set "ON"
														wait(3)
															JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ReviewQuorumOption").Object.setSelected(true)
															wait 5
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Review Quorum option " + arrQuorum(0))
																		Fn_MyWorkList_SignoffTeamSelect = False
																		
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Review Quorum option " + arrQuorum(0))
														End if
			
														If arrQuorum(0)  = "Percent" Then
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ReviewQuorumPercentage").Set arrQuorum(1)
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ReviewQuorumPercentage").Activate
                                                                        'JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ProcDesc").Click 1,1,"LEFT"
																		wait 2
														Else				
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ReviewQuorumNumeric").Set arrQuorum(1)
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ReviewQuorumNumeric").Activate
                                                                        'JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ProcDesc").Click 1,1,"LEFT"
																		wait 2
														End If
														Call Fn_ReadyStatusSync(5)			
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Review Quorum value " + arrQuorum(1))
																		Fn_MyWorkList_SignoffTeamSelect = False
																		
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Review Quorum value " + arrQuorum(1))
														End If
											End if

								Case "Wait"	 
											 If dicItems(iCounter) = true Then
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("WaitForUndecidedReviewers").Set  "ON"
											Else
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("WaitForUndecidedReviewers").Set  "OFF"
											End If

											Call Fn_ReadyStatusSync(5)

											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Wait For Undecided Reviewers value to " + dicItems(iCounter))
															Fn_MyWorkList_SignoffTeamSelect = False
															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Wait For Undecided Reviewers value to " + dicItems(iCounter))
											End if
								
								Case "Adhoc"	 
											 If dicItems(iCounter) = true Then
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("Ad-hocDone").Set  "ON"
											Else
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("Ad-hocDone").Set  "OFF"
											End If

											'Call Fn_ReadyStatusSync(5) [Harshal Tanpure]:To handle the Negative Error number

											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Ad-hoc Done value to " + CStr(dicItems(iCounter)))
															Fn_MyWorkList_SignoffTeamSelect = False
															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Ad-hoc Done value to " + CStr(dicItems(iCounter)))
											End if

											Call Fn_ReadyStatusSync(6)
											If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").Exist Then
												JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").Click micLeftBtn
												If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Apply Sign off Team Settings")
															Fn_MyWorkList_SignoffTeamSelect = False															
															Exit Function
												Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Applied Sign off Team Settings")
												End if
											End If

											Call Fn_ReadyStatusSync(5)
								Case "ProcessDescription"    											

											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ProcDesc").Set dicItems(iCounter)
											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Description value to " + dicItems(iCounter))
															Fn_MyWorkList_SignoffTeamSelect = False															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Description  value to " + dicItems(iCounter))
											End if

											Call Fn_ReadyStatusSync(5)

								Case "Comments" 
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("Comment").Set dicItems(iCounter)
											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Comment value to " + dicItems(iCounter))
															Fn_MyWorkList_SignoffTeamSelect = False															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Comment value to " + dicItems(iCounter))
											End if
   
					 End Select
			End if		
	Next

	Case "Profiles"
				For iCounter = 0 to dicCount - 1
	                    If  dicItems(iCounter) <> "" Then
	                            Select Case dicKeys(iCounter)

								 Case "WorkListTreeNode"	

											'Select the WorkList Tree Node		
											bReturn =  Fn_MyWorkList_TreeNodeOperations("Select",dicItems(iCounter),"")
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + dicItems(iCounter) +  " From WorkList Tree")
													Exit Function
											End If
											Call Fn_ReadyStatusSync(5)

											'Set the Viewer Tab
											bReturn =  Fn_MyTc_TabSet("Viewer")	
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Viewer Tab")
													Exit Function
											End If

											Call Fn_ReadyStatusSync(10)

											 'Set Default View to Task View 
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ViewOptions").SetTOProperty "Attached Text","Task View"
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ViewOptions").Set "ON"
											If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Task View from Viewer tab")
													Fn_MyWorkList_SignoffTeamSelect = False									
													Exit Function
											Else																						
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Task View from Viewer tab")
													Call Fn_ReadyStatusSync(5)
													wait(2)
											End If
											Call Fn_ReadyStatusSync(5)

								Case "SignOffTeamSelect"

											' Expand the Node in Signoff Team tree
                                             If  instr(1,dicSelectSignOff.Item("SignOffTeamSelect"),"#0") > 0 Then
															
															If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("SignOffTeamTree").Exist Then
																	JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("SignOffTeamTree").Expand "#0:#0"
															Else		
																	JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("RouteSignOffTeamTree").Expand "#0:#0"
															End If

                                                            If Err.Number < 0 Then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Profiles Node")
																	Fn_MyWorkList_SignoffTeamSelect = False									
																	Exit Function
															Else																						
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expanded Profiles Node")
															End If
											 Else
															bReturn =  Fn_MyWorkList_SignoffTeam_TreeNodeOperations("Expand","Profiles")          
															If bReturn = false Then
																	Fn_MyWorkList_SignoffTeamSelect = False									
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Profiles from Signoff Team Tree")
																	Exit Function
															Else
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Expanded Profiles from Signoff Team Tree")
															End If
											 End If
											Call Fn_ReadyStatusSync(5)

											'Select the Node Under Profiles	
											bReturn =  Fn_MyWorkList_SignoffTeam_TreeNodeOperations("Select", dicItems(iCounter))          
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select "+ dicItems(iCounter) +" from Signoff Team Tree")
													Exit Function
											Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Successfully Selected "+ dicItems(iCounter) +" from Signoff Team Tree")
											End If
											Call Fn_ReadyStatusSync(5)

								  Case "UsersName"                  ' Select the nodes from Organization tree

											 ' Select the Organization Tab
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTab("OrgProjTeamTab").Select "Organization"
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Organization Tab")
													Exit Function
											Else
													Fn_MyWorkList_SignoffTeamSelect = true									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Organization Tab")
											End If

											Call Fn_ReadyStatusSync(5)
											Wait 2
											'Function Called to Select Organization Tree Node	
											bReturn = Fn_MyWorkList_Org_TreeNodeOperations("MultiSelect", dicItems(iCounter))
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Organization Tree Node " + dicItems(iCounter))
													Exit Function
											End If	
											Wait 2
											'Case Member Option
											If dicKeys("MemberOption") <> "" Then															
															Select Case dicSelectSignOff.Item("MemberOption")
																	Case "Any"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").SetTOProperty "attached text","Any Member"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").Set "ON"                                                      							
																	Case "All"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").SetTOProperty "attached text","All Members"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").Set "ON"                                                      														
															End Select

															If Err.Number < 0 Then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Member Type " +  dicSelectSignOff.Item("MemberOption") + " Member")
																	Fn_MyWorkList_SignoffTeamSelect = False
																	
																	Exit Function
															Else
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Member Type " +   dicSelectSignOff.Item("MemberOption") + " Member")				
															End If			
											End If

											'Select Group Option	
											If dicKeys("GroupOption") <> "" Then
															Select Case dicSelectSignOff.Item("GroupOption")
																	Case "Any"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Any Group"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").Set "ON"                                                      
																	Case "Specific"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Specific Group"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").Set "ON" 														
															End Select
														
															If Err.Number < 0 Then
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group Type" +   dicSelectSignOff.Item("GroupOption") + " Group")
																			Fn_MyWorkList_SignoffTeamSelect = False
																			
																			Exit Function
															Else
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group Type" +   dicSelectSignOff.Item("GroupOption") + " Group")				
															End If		
											End If
                                            Wait 3					
											'Click on Add button
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Add").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Add User " + dicItems(iCounter)   + " from Organization Tree")
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Added User " + dicItems(iCounter)   + " from Organization Tree")				
											End If
											Call Fn_ReadyStatusSync(10)
											Wait 3
											
									Case "ProjectName"				
											 ' Select the Project Teams Tab
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTab("OrgProjTeamTab").Select "Project Teams"
											If bReturn = false Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Project Teams Tab")
													Exit Function
											Else																					
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Project Teams Tab")
											End If

											' Select the Project 
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaList("Projects").Select  dicItems(iCounter)	
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Project " + dicItems(iCounter)   + " from Project List")
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Project " + dicItems(iCounter)   + " from Project List")				
											End If                                          

											Call Fn_ReadyStatusSync(5)

									Case "ProjectUsers"
											' Select the Node in Signoff Team tree
											bReturn = Fn_MyWorkList_SignoffTeam_TreeNodeOperations("Select",sActionSelect)          
											If bReturn = false Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Signoff Team tree Node" + sActionSelect)
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Signoff Team tree Node " + sActionSelect)				
										
											End If
											Call Fn_ReadyStatusSync(5)

											'Expand the node	
											arrNodeList = split(dicItems(iCounter), ",",-1,1)								
											For iNodeCounter = 0 to UBound(arrNodeList)
													arrNode = split(arrNodeList(iNodeCounter),":",-1,1)
													'sExpandNode = "#0:"
													For iItemCount = 0 to UBound(arrNode)-1
															sExpandNode = sExpandNode + arrNode(iItemCount)
															JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProjectsTree").Expand sExpandNode
															wait(1)
															sExpandNode = sExpandNode + ":"
													Next						
											Next

											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProjectsTree").Object.clearSelection
											For iNodeCounter = 0 to UBound(arrNodeList)
													JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProjectsTree").ExtendSelect "#0:" + arrNodeList(iNodeCounter)
													wait(1)		
											Next
											
											If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + arrNodeList(iNodeCounter)   + " from Project User Tree")
													Fn_MyWorkList_SignoffTeamSelect = False                                                 
													Exit Function				
											Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Multiselected node   " + arrNodeList(iNodeCounter) + " of Project Teams Tree." )   
											End If

											Call Fn_ReadyStatusSync(5)
																						
											' Member Option
											If dicKeys("MemberOption") <> "" Then															
															Select Case dicSelectSignOff.Item("MemberOption")
																	Case "Any"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").SetTOProperty "attached text","Any Member"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").Set "ON"                                                      							
																	Case "All"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").SetTOProperty "attached text","All Members"
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolMemberOption").Set "ON"                                                      														
															End Select
							
															If Err.Number < 0 Then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Member Type " + dicSelectSignOff.Item("MemberOption") + " Member")
																	Fn_MyWorkList_SignoffTeamSelect = False
																	
																	Exit Function
															Else
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Member Type " + dicSelectSignOff.Item("MemberOption") + " Member")				
															End If			
											End If

											'Select Group Option	
											If dicKeys("GroupOption") <> "" Then
															Select Case dicSelectSignOff.Item("GroupOption")
																	Case "Any"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Any Group"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").Set "ON"                                                      
																	Case "Specific"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Specific Group"
																					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ResPoolGroupOption").Set "ON" 														
															End Select
														
															If Err.Number < 0 Then
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group Type" +  dicSelectSignOff.Item("GroupOption") + " Group")
																			Fn_MyWorkList_SignoffTeamSelect = False
																			
																			Exit Function
															Else
																			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group Type" +  dicSelectSignOff.Item("GroupOption") + " Group")				
															End If		
											 End If

											Call Fn_ReadyStatusSync(6)
											'Click on Add button
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Add").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Add User to Project " + dicItems(iCounter)   + " from Organization Tree")
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Added User to Project " + dicItems(iCounter)   + " from Organization Tree")				
											End If	

											Call Fn_ReadyStatusSync(6)

									Case "Quorum"
											If  dicItems(iCounter) <> "" Then
														arrQuorum = split(dicItems(iCounter),":",-1,1) 
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ReviewQuorumOption").SetTOProperty "attached text",arrQuorum(0) 
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ReviewQuorumOption").Set "ON"
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Review Quorum option " + arrQuorum(0))
																		Fn_MyWorkList_SignoffTeamSelect = False
																		
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Review Quorum option " + arrQuorum(0))
														End if
			
														If arrQuorum(0)  = "Percent" Then
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ReviewQuorumPercentage").Set arrQuorum(1)
														Else				
																		JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ReviewQuorumNumeric").Set arrQuorum(1)
														End If
														Call Fn_ReadyStatusSync(5)			
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Review Quorum value " + arrQuorum(1))
																		Fn_MyWorkList_SignoffTeamSelect = False
																		
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Review Quorum value " + arrQuorum(1))
														End If
											End if

								Case "Wait"	 
											 If dicItems(iCounter) = true Then
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("WaitForUndecidedReviewers").Set  "ON"
											Else
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("WaitForUndecidedReviewers").Set  "OFF"
											End If

											Call Fn_ReadyStatusSync(5)

											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Wait For Undecided Reviewers value to " + dicItems(iCounter))
															Fn_MyWorkList_SignoffTeamSelect = False
															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Wait For Undecided Reviewers value to " + dicItems(iCounter))
											End if

								Case "Adhoc"	 
											If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("Ad-hocDone").Exist(5) = false Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify exixtance of  [Ad-hocDone] button " )
														Fn_MyWorkList_SignoffTeamSelect = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verifed exixtance of  [Ad-hoc Done] button")
											End If

											 If dicItems(iCounter) = true Then
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("Ad-hocDone").Set  "ON"
											Else
														JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("Ad-hocDone").Set  "OFF"
											End If

											Call Fn_ReadyStatusSync(5)

											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Ad-hoc Done value to " + dicItems(iCounter))
															Fn_MyWorkList_SignoffTeamSelect = False
															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Ad-hoc Done value to " + dicItems(iCounter))
											End if

											Call Fn_ReadyStatusSync(5)
																						
											If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").Exist Then
												JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaButton("Apply").Click micLeftBtn
												If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Apply Sign off Team Settings")
															Fn_MyWorkList_SignoffTeamSelect = False															
															Exit Function
												Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Applied Sign off Team Settings")
												End if
											End If

											Call Fn_ReadyStatusSync(5)
								Case "ProcessDescription"    											

											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("ProcDesc").Set dicItems(iCounter)
											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Description value to " + dicItems(iCounter))
															Fn_MyWorkList_SignoffTeamSelect = False															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Description  value to " + dicItems(iCounter))
											End if

											Call Fn_ReadyStatusSync(5)

								Case "Comments" 
											JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("Comment").Set dicItems(iCounter)
											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Comment value to " + dicItems(iCounter))
															Fn_MyWorkList_SignoffTeamSelect = False															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Comment value to " + dicItems(iCounter))
											End if
   
					 End Select
			End if		
	Next						


	Case "Menu"

		If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").Exist Then
					Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("SelectSignoffTeamDialog")
		ElseIf JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").exist Then
					Set objDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("SelectSignoffTeamDialog")
		else
					Set objDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("SelectSignoffTeamDialog")		
		End If
		'Added by Nilesh on 13-Jun-12 For TC10 Build 0606 change
        If objDialog.Exist(5)=False Then
			Set objDialog=JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("SelectSignoffTeamDialog")
		End If

			For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> "" Then
							Select Case dicKeys(iCounter)

							Case "WorkListTreeNode"

										'Select the WorkList Tree Node		
										bReturn =  Fn_MyWorkList_TreeNodeOperations("Select",dicItems(iCounter),"")
										If bReturn = false Then
												Fn_MyWorkList_SignoffTeamSelect = False					
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + dicItems(iCounter) +  " From WorkList Tree")
												Exit Function
										End If
										Call Fn_ReadyStatusSync(5)

										' Call menu Operation
										If objDialog.Exist = False Then
												bReturn = Fn_MenuOperation("Select", "Actions:Perform")
												Call Fn_ReadyStatusSync(5)
												If bReturn = False Then
														Fn_MyWorklist_PerformSignOff = False
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Action --> Perform ") 
														Exit Function						
												Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Action --> Perform")
														Wait(5)
												End If
										Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Select SignOff Team Dialog Dialog Exist. ") 		
										End If
						' SignOffTeamSelectWithoutClose case is added to select signoff team without close
							Case "SignOffTeamSelect","SignOffTeamSelectWithoutClose"
										sCase = dicKeys(iCounter)
										If dicKeys(iCounter) = "SignOffTeamSelectWithoutClose" Then
											dicSelectSignOff("SignOffTeamSelect") = dicSelectSignoff("SignOffTeamSelectWithoutClose")
										End If
'										objDialog.RefreshObject
'										If dicSelectSignOff("SignOffTeamSelect") <> ""  Then
'											arrNodeList = Split(dicSelectSignOff("SignOffTeamSelect"), ":", -1, 1)
'											If IsArray(arrNodeList) = True Then
'												sExpnadNode = "#0:"+arrNodeList(0)
'											End If
'										End If
'										If sExpnadNode <> "" Then
'										   ' Select the Node in Signoff Team tree
'											objDialog.JavaTree("SignOffTeamTree").Select sExpnadNode    
'											If  Err.Number < 0 Then
'													Fn_MyWorkList_SignoffTeamSelect = False			
'													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select "+ sExpnadNode +" from Signoff Team Tree")
'													objDialog.JavaButton("Close").Click micLeftBtn
'													Exit Function
'											End If
'											Call Fn_ReadyStatusSync(5)
'									
'										   ' Select the Node in Signoff Team tree
'											objDialog.JavaTree("SignOffTeamTree").Expand sExpnadNode           
'											If  Err.Number < 0 Then
'													Fn_MyWorkList_SignoffTeamSelect = False
'													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand "+ sExpnadNode +" from Signoff Team Tree")
'													objDialog.JavaButton("Close").Click micLeftBtn
'													Exit Function
'											End If
'											Call Fn_ReadyStatusSync(5)
'										End If
'
'										objDialog.JavaTree("SignOffTeamTree").Select "#0:"+dicItems(iCounter)
'										If  Err.Number < 0 Then
'												Fn_MyWorkList_SignoffTeamSelect = False							
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ["+dicItems(iCounter)+"] SignOff Team")
'												objDialog.JavaButton("Close").Click micLeftBtn
'												Exit Function
'										Else		
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected ["+dicItems(iCounter)+"] Sign Off Team")
'										End If
'
'										Call Fn_ReadyStatusSync(5)
									'------------------------------------------------------------------------------------------------------------------------
									'Changed code for Selection of SignoffTeam due to new Required field added into SignOff tree 
										objDialog.RefreshObject
										sNodeName = dicSelectSignOff("SignOffTeamSelect")
										aSplitNodeName=split(dicSelectSignOff("SignOffTeamSelect"),":",-1,1)
										sNodeName=aSplitNodeName(ubound(aSplitNodeName))
										iCnt = 0
										For iAll = 0 to Cint(objDialog.JavaTree("SignOffTeamTree").GetROProperty("items count"))-1
											sTempNode = objDialog.JavaTree("SignOffTeamTree").Object.getPathForRow(iAll).tostring()
											Err.Clear
											If instr(1,sTempNode,sNodeName) Then
													objDialog.JavaTree("SignOffTeamTree").Object.setSelectionRow iAll
													If  Err.Number < 0 Then
															Fn_MyWorkList_SignoffTeamSelect = False							
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ["+dicSelectSignOff("SignOffTeamSelect")+"] SignOff Team")
															objDialog.JavaButton("Close").Click micLeftBtn
															Exit Function
													End If
													Call Fn_ReadyStatusSync(5)
													Exit for
											Else
											 		If instr(1,sTempNode,aSplitNodeName(iCnt)) Then
														objDialog.JavaTree("SignOffTeamTree").Object.expandPath(objDialog.JavaTree("SignOffTeamTree").Object.getPathForRow(iAll))
														Call Fn_ReadyStatusSync(5)
														iCnt = iCnt + 1
													End If
											End If
										Next
										
										If iAll = Cint(objDialog.JavaTree("SignOffTeamTree").GetROProperty("items count")) Then
												Fn_MyWorkList_SignoffTeamSelect = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Exist "+ dicSelectSignOff("SignOffTeamSelect") +" from Signoff Team Tree")
												objDialog.JavaButton("Close").Click micLeftBtn
												Exit Function
										End If
										Call Fn_ReadyStatusSync(1)
									'------------------------------------------------------------------------------------------------------------------------	

							  Case "UsersName" 

										 ' Select the Organization Tab
										objDialog.JavaTab("OrgProjTeamTab").Select "Organization"
										If  Err.Number < 0 Then
												Fn_MyWorkList_SignoffTeamSelect = False									
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Organization Tab")
												objDialog.JavaButton("Close").Click micLeftBtn
												Exit Function
										Else							
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Organization Tab")
										End If

										Call Fn_ReadyStatusSync(5)

										'Expand the Nodes
									    arrNodeList = split(dicItems(iCounter), ",",-1,1)								
										For iNodeCounter = 0 to UBound(arrNodeList)
												arrNode = split(arrNodeList(iNodeCounter),":",-1,1)
												sExpandNode = "#0:"
												For iItemCount = 0 to UBound(arrNode)-1
														sExpandNode = sExpandNode + arrNode(iItemCount)
														objDialog.JavaTree("OrgProjectTree").Expand sExpandNode
														wait(1)
														sExpandNode = sExpandNode + ":"
												Next						
										Next

										'Select the Users
										objDialog.JavaTree("OrgProjectTree").Object.clearSelection
										For iNodeCounter = 0 to UBound(arrNodeList)
												objDialog.JavaTree("OrgProjectTree").ExtendSelect "#0:" + arrNodeList(iNodeCounter)
												wait(1)		
										Next
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User " + dicItems(iCounter))
												Fn_MyWorkList_SignoffTeamSelect = False
												objDialog.JavaButton("Close").Click micLeftBtn
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User " + dicItems(iCounter))
										End If	

										'Case Member Option dicSelectSignOff
										If dicSelectSignOff.Item("MemberOption") <> "" Then															
														Select Case dicSelectSignOff.Item("MemberOption")
																Case "Any"
																	objDialog.JavaRadioButton("MemberOption").SetTOProperty "attached text","Any Member"
																	objDialog.JavaRadioButton("MemberOption").Set "ON"                                                      							
																Case "All"
																	objDialog.JavaRadioButton("MemberOption").SetTOProperty "attached text","All Members"
																	objDialog.JavaRadioButton("MemberOption").Set "ON"                                                      														
														End Select

														If Err.Number < 0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Member Type " +  dicSelectSignOff.Item("MemberOption") + " Member")
																Fn_MyWorkList_SignoffTeamSelect = False
																objDialog.JavaButton("Close").Click micLeftBtn
																Exit Function
														Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Member Type " +   dicSelectSignOff.Item("MemberOption") + " Member")				
														End If			
										End If

										'Select Group Option	
										If dicSelectSignOff.Item("GroupOption") <> "" Then
														Select Case dicSelectSignOff.Item("GroupOption")
																Case "Any"
																				objDialog.JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Any Group"
																				objDialog.JavaRadioButton("ResPoolGroupOption").Set "ON"                                                      
																Case "Specific"
																				objDialog.JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Specific Group"
																				objDialog.JavaRadioButton("ResPoolGroupOption").Set "ON" 														
														End Select
													
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group Type" +   dicSelectSignOff.Item("GroupOption") + " Group")
																		Fn_MyWorkList_SignoffTeamSelect = False
																		objDialog.JavaButton("Close").Click micLeftBtn
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group Type" +   dicSelectSignOff.Item("GroupOption") + " Group")				
														End If		
										End If
																		
										'Click on Add button
										objDialog.JavaButton("Add").Click micLeftBtn
										If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Add User " + dicItems(iCounter)   + " from Organization Tree")
													Fn_MyWorkList_SignoffTeamSelect = False
													objDialog.JavaButton("Close").Click micLeftBtn
													Exit Function
										Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Added User " + dicItems(iCounter)   + " from Organization Tree")				
										End If	

										Call Fn_ReadyStatusSync(5)

								Case "ProjectName"
										 ' Select the Project Teams Tab
										objDialog.JavaTab("OrgProjTeamTab").Select "Project Teams"
										If  Err.Number < 0 Then
												Fn_MyWorkList_SignoffTeamSelect = False									
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Project Teams Tab")
												objDialog.JavaButton("Close").Click micLeftBtn
												Exit Function
										Else																					
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Project Teams Tab")
										End If

										' Select the Project 
										objDialog.JavaList("Projects").Select  dicItems(iCounter)	
										If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Project " + dicItems(iCounter)   + " from Project List")
													Fn_MyWorkList_SignoffTeamSelect = False
													objDialog.JavaButton("Close").Click micLeftBtn
													Exit Function
										Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Project " + dicItems(iCounter)   + " from Project List")				
										End If                                          

										Call Fn_ReadyStatusSync(5)

								Case "ProjectUsers"
										' Select the Node in Signoff Team tree
'											bReturn = Fn_MyWorkList_SignoffTeam_TreeNodeOperations("Select",sActionSelect)          
'											If bReturn = false Then
'														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Signoff Team tree Node" + sActionSelect)
'														Fn_MyWorkList_SignoffTeamSelect = False
'														Exit Function
'											Else
'														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Signoff Team tree Node " + sActionSelect)				
'										
'											End If
'											Call Fn_ReadyStatusSync(3)

										'Expand the node	
										If instr(1,dicItems(iCounter),",") Then
											arrNodeList = split(dicItems(iCounter), ",",-1,1)								
										else
											arrNodeList = Array(dicItems(iCounter))								
										End If
										For iNodeCounter = 0 to UBound(arrNodeList)
												arrNode = split(arrNodeList(iNodeCounter),":",-1,1)
												'sExpandNode = "#0:"
												For iItemCount = 0 to UBound(arrNode)-1
														sExpandNode = sExpandNode + arrNode(iItemCount)
														objDialog.JavaTree("OrgProjectTree").Expand sExpandNode
														wait(1)
														sExpandNode = sExpandNode + ":"
												Next						
										Next

										objDialog.JavaTree("OrgProjectTree").Object.clearSelection
										For iNodeCounter = 0 to UBound(arrNodeList)
'												objDialog.JavaTree("OrgProjectTree").Select "#0:" + arrNodeList(iNodeCounter)
												objDialog.JavaTree("OrgProjectTree").Select  arrNodeList(iNodeCounter) ' Modified by : Harshal Tanpure 22-March-2013 Porting Tc 10.1 build : Teamcenter 10 (20130306.00)
												wait(1)		
										Next
										
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + arrNodeList(iNodeCounter)   + " from Project User Tree")
												Fn_MyWorkList_SignoffTeamSelect = False
												objDialog.JavaButton("Close").Click micLeftBtn
												Exit Function
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Multiselected node   " + arrNodeList(iNodeCounter) + " of Project Teams Tree." )   
										End If

										Call Fn_ReadyStatusSync(1)
																					
										' Member Option
										If dicKeys("MemberOption") <> "" Then															
														Select Case dicSelectSignOff.Item("MemberOption")
																Case "Any"
																	objDialog.JavaRadioButton("MemberOption").SetTOProperty "attached text","Any Member"
																	objDialog.JavaRadioButton("MemberOption").Set "ON"                                                      							
																Case "All"
																	objDialog.JavaRadioButton("MemberOption").SetTOProperty "attached text","All Members"
																	objDialog.JavaRadioButton("MemberOption").Set "ON"                                                      														
														End Select
						
														If Err.Number < 0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Member Type " + dicSelectSignOff.Item("MemberOption") + " Member")
																Fn_MyWorkList_SignoffTeamSelect = False
																objDialog.JavaButton("Close").Click micLeftBtn
																Exit Function
														Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Member Type " + dicSelectSignOff.Item("MemberOption") + " Member")				
														End If			
										End If

										'Select Group Option	
										If dicKeys("GroupOption") <> "" Then
														Select Case dicSelectSignOff.Item("GroupOption")
																Case "Any"
																				objDialog.JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Any Group"
																				objDialog.JavaRadioButton("ResPoolGroupOption").Set "ON"                                                      
																Case "Specific"
																				objDialog.JavaRadioButton("ResPoolGroupOption").SetTOProperty "attached text","Specific Group"
																				objDialog.JavaRadioButton("ResPoolGroupOption").Set "ON" 														
														End Select
													
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group Type" +  dicSelectSignOff.Item("GroupOption") + " Group")
																		Fn_MyWorkList_SignoffTeamSelect = False
																		objDialog.JavaButton("Close").Click micLeftBtn																			
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group Type" +  dicSelectSignOff.Item("GroupOption") + " Group")				
														End If		
										 End If

										'Click on Add button
										objDialog.JavaButton("Add").Click micLeftBtn
										If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Add User to Project " + dicItems(iCounter)   + " from Organization Tree")
													Fn_MyWorkList_SignoffTeamSelect = False
													objDialog.JavaButton("Close").Click micLeftBtn
													Exit Function
										Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Added User to Project " + dicItems(iCounter)   + " from Organization Tree")				
										End If	

										Call Fn_ReadyStatusSync(5)

								Case "Quorum", "Quorum_ext"
										If  dicItems(iCounter) <> "" Then
													arrQuorum = split(dicItems(iCounter),":",-1,1) 
													objDialog.JavaRadioButton("ReviewQuorumOption").SetTOProperty "attached text",arrQuorum(0) 
													objDialog.JavaRadioButton("ReviewQuorumOption").Set "ON"
													If Err.Number < 0 Then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Review Quorum option " + arrQuorum(0))
																	Fn_MyWorkList_SignoffTeamSelect = False
																	objDialog.JavaButton("Close").Click micLeftBtn
																	Exit Function
													Else
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Review Quorum option " + arrQuorum(0))
													End if
		
													If arrQuorum(0)  = "Percent" Then
																	objDialog.JavaEdit("ReviewQuorumPercentage").Set arrQuorum(1)
																	
													ElseIf dicKeys(iCounter) = "Quorum_ext" Then
																	objDialog.JavaEdit("ReviewQuorumNumeric").Type arrQuorum(1)
													Else
																	objDialog.JavaEdit("ReviewQuorumNumeric").Set arrQuorum(1)
'																	objDialog.JavaEdit("ReviewQuorumNumeric").Type arrQuorum(1)
																	wait 2
													End If
													Call Fn_ReadyStatusSync(3)			
													If Err.Number < 0 Then
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Review Quorum value " + arrQuorum(1))
																	Fn_MyWorkList_SignoffTeamSelect = False
																	objDialog.JavaButton("Close").Click micLeftBtn
																	Exit Function
													Else
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Review Quorum value " + arrQuorum(1))
													End If
										End if

							Case "Wait"	 
										 If dicItems(iCounter) = true Then
													objDialog.JavaCheckBox("WaitForUndecidedReviewers").Set  "ON"
										Else
													objDialog.JavaCheckBox("WaitForUndecidedReviewers").Set  "OFF"
										End If

										Call Fn_ReadyStatusSync(3)

										If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Wait For Undecided Reviewers value to " + CStr(dicItems(iCounter)))
														Fn_MyWorkList_SignoffTeamSelect = False
														objDialog.JavaButton("Close").Click micLeftBtn
														Exit Function
										Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Wait For Undecided Reviewers value to " + CStr(dicItems(iCounter)))
										End if

							Case "Adhoc"
										Err.Clear
										If dicItems(iCounter) = true Then
													objDialog.JavaCheckBox("Adhoc").Set  "ON"
										Else
													objDialog.JavaCheckBox("Adhoc").Set  "OFF"
										End If

										Call Fn_ReadyStatusSync(3)
										If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Ad-hoc Done value to " + dicItems(iCounter))
														Fn_MyWorkList_SignoffTeamSelect = False
														objDialog.JavaButton("Close").Click micLeftBtn
														Exit Function
										Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Ad-hoc Done value to " + dicItems(iCounter))
										End If
										Call Fn_ReadyStatusSync(5)
																					
										If objDialog.JavaButton("Apply").Exist Then
											objDialog.JavaButton("Apply").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Apply Sign off Team Settings")
														Fn_MyWorkList_SignoffTeamSelect = False
														objDialog.JavaButton("Close").Click micLeftBtn
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Applied Sign off Team Settings")
											End If
										End If

										Call Fn_ReadyStatusSync(5)									
										

							Case "ProcessDescription"    											

										objDialog.JavaEdit("Process Description:").Set dicItems(iCounter)
										If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Description value to " + dicItems(iCounter))
														Fn_MyWorkList_SignoffTeamSelect = False	
														objDialog.JavaButton("Close").Click micLeftBtn
														Exit Function
										Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Description  value to " + dicItems(iCounter))
										End if

										Call Fn_ReadyStatusSync(5)

							Case "Comments" 
								objDialog.JavaEdit("Comments").Activate
								objDialog.JavaEdit("Comments").SetFocus
								objDialog.JavaEdit("Comments").Set ""
								objDialog.JavaEdit("Comments").Set dicItems(iCounter)
								Call Fn_ReadyStatusSync(2)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Comment value to " + dicItems(iCounter))
									Fn_MyWorkList_SignoffTeamSelect = False	
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Comment value to " + dicItems(iCounter))
								End if
							
							 '[TC1123-20161205-13_03_2017-ChaitaliR-NewDevelopment] - Case added to Verify & Set RequiredCheckBox							
							Case "Required"
										If dicItems(iCounter) = "VerifyRequiredCheckBox" Then
													bReturn = Fn_UI_Object_GetROProperty("Fn_MyWorkList_SignoffTeamSelect",objDialog.JavaCheckBox("Required"), "enabled")
													If bReturn <> 0 Then
															Fn_MyWorkList_SignoffTeamSelect = False									
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify [ Required ] checkbox is enabled.")
															Exit Function
													End If
													bReturn = Fn_UI_Object_GetROProperty("Fn_MyWorkList_SignoffTeamSelect",objDialog.JavaCheckBox("Required"), "value")
													If bReturn <> 1 Then
															Fn_MyWorkList_SignoffTeamSelect = False									
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify [ Required ] checkbox is checked.")
															Exit Function
													End If													
										End If
										
										Err.Clear
										If dicItems(iCounter) = "SetRequiredCheckBox" Then
													objDialog.JavaCheckBox("Required").Set  "ON"
										End If

										Call Fn_ReadyStatusSync(3)
										
										If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Required value to " + dicItems(iCounter))
														Fn_MyWorkList_SignoffTeamSelect = False
														objDialog.JavaButton("Close").Click micLeftBtn
														Exit Function
										Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Required value to " + dicItems(iCounter))
										End If
										Call Fn_ReadyStatusSync(2)
										
								'  [TC1123-20161205-13_03_2017-ChaitaliR-NewDevelopment] - Case added to  Verify Action List Value
								Case "VerifyAction"
											objDialog.JavaList("Projects").SetTOProperty "attached text","Action"
											bReturn = Fn_UI_Object_GetROProperty("Fn_MyWorkList_SignoffTeamSelect",objDialog.JavaList("Projects"), "text")
											If bReturn <> dicItems(iCounter) Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify [ Action ] list text.")
													Exit Function
											End If
											
											bReturn = Fn_UI_Object_GetROProperty("Fn_MyWorkList_SignoffTeamSelect",objDialog.JavaList("Projects"), "value")
											If bReturn <> dicItems(iCounter) Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify [ Action ] list value.")
													Exit Function
											End If
								'  [TC1123-20161205-06_APR_2017-ShwetaR-NewDevelopment] - Case added to  Verify static text
								Case "VerifyStaticText"
										objDialog.JavaStaticText("StaticText").SetTOProperty "label",dicItems(iCounter)
										bReturn = Fn_SISW_UI_Object_Operations("Fn_MyWorkList_SignoffTeamSelect","Exist", objDialog.JavaStaticText("StaticText"),SISW_MICRO_TIMEOUT)
										If bReturn = False Then
													Fn_MyWorkList_SignoffTeamSelect = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify [ static text ] .")
													Exit Function
											End If
								Case "VerifyProcessDesc"
									bReturn = objDialog.JavaEdit("Process Description:").GetROProperty("text")
									if bReturn <> dicItems(iCounter) then
										If bReturn = False Then
											Fn_MyWorkList_SignoffTeamSelect = False									
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify [ Process Description: ] .")
											Exit Function
										End If
									End if
							 End Select
					End if
			Next
			
			If sCase <> "SignOffTeamSelectWithoutClose" Then
				objDialog.JavaButton("Close").Click micLeftBtn
				If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Close Button")
								Fn_MyWorkList_SignoffTeamSelect = False	
								Exit Function
				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Close Button ")
				End if
			End If
			
			

			Fn_MyWorkList_SignoffTeamSelect = True
			Set objDialog = Nothing
	End Select

End Function 

'*********************************************************  Function is Subscribe to any Resource Pool *********************************************************************

'Function Name		:					Fn_MyWorkList_ResourcePoolSubscription

'Description			 :		 		   Subscribe to any Resource Pool														

'Parameters			   :	 			 1. sAction: Add/Remove
'													 2. sInboxNode: Resource pool node getting displayed in the MyWorklist tree
' 												 	3. sOption: Accessible/All
'												   4.sGroup: Resource Pool group
'												  5.sRole: Resource Pool role

'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Myworklist tab should selected.

'Examples				:			Fn_MyWorkList_ResourcePoolSubscription("Add","My Worklist:dba/CostDBA","Accessible","dba","CostDBA")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										    Rupali				        16-Aug-2010	        1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_ResourcePoolSubscription(sAction,sInboxNode,sOption,sGroup,sRole)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_ResourcePoolSubscription"
   On Error Resume Next 
   Dim bReturn,objDialog,iItemCount,arrNodeList,sTreeItem,iCounter, aRdBOption, iCnt
   Dim bFlag
   Dim objTree

   Select Case sAction 
		Case "Add"
			bReturn = Fn_MyWorkList_TreeNodeOperations("Exist",sInboxNode,"")
			If bReturn = False Then
				'Select the Menu.
				bReturn = Fn_MenuOperation("Select","Tools:Resource Pool Subscription...")
				Call Fn_ReadyStatusSync(3)		   	    
				If bReturn = False Then
					Fn_MyWorkList_ResourcePoolSubscription = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fail to Select Menu Tools:Resource Pool Subscription...") 
					Exit Function
				Else 
					Fn_MyWorkList_ResourcePoolSubscription = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Menu Tools:Resource Pool Subscription...") 
				End If

				' Check  existance of the Dialog
				Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Resource Pool Subscription")
			
				If objDialog.Exist(5) Then		
					bFlag = False
					Set objTree = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Resource Pool Subscription").JavaTree("ResPoolTree")
					For i = 0 To CInt(objTree.GetROProperty("items count")) - 1
						If instr(objTree.getItem(i) ,sGroup & "/" & sRole) > 0 Then
							bFlag = True
						End If
					Next					

					If bFlag = False Then
						If sOption <> "" Then
							If sOption = "All" Then
								objDialog.JavaRadioButton("ResPoolOpt").SetTOProperty "attached text","All"
								objDialog.JavaRadioButton("ResPoolOpt").Set "ON"
							ElseIf sOption = "Accessible" Then
								objDialog.JavaRadioButton("ResPoolOpt").SetTOProperty "attached text","Accessible"
								objDialog.JavaRadioButton("ResPoolOpt").Set "ON"
							Else
								Fn_MyWorkList_ResourcePoolSubscription = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : " + sOption + " radiobutton does not exist.") 
								objDialog.JavaButton("Cancel").Click micLeftBtn
								Set objDialog = Nothing
								Exit Function 
							End If
						End If
				
						If sGroup <> "" Then
							bReturn =  Fn_SISW_UI_JavaList_Operations("", "Select", objDialog, "Group", sGroup, "", "")
				'			objDialog.JavaList("Group").Select sGroup
							If bReturn = False Then
								Fn_MyWorkList_ResourcePoolSubscription = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Failed to select Group " + sGroup) 
								objDialog.JavaButton("Cancel").Click micLeftBtn
								Set objDialog = Nothing
								Exit Function 
							End If
						End If
						If sRole <> "" Then
							bReturn =  Fn_SISW_UI_JavaList_Operations("", "Select", objDialog, "Role", sRole, "", "")
				'			objDialog.JavaList("Role").Select sRole
							If bReturn = False Then
								Fn_MyWorkList_ResourcePoolSubscription = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Failed to select Role " + sRole) 
								objDialog.JavaButton("Cancel").Click micLeftBtn
								Set objDialog = Nothing
								Exit Function 
							End If
						End If
						objDialog.JavaButton("Add").WaitProperty "enabled",1,20000
						objDialog.JavaButton("Add").Click micLeftBtn
						Wait(5)
				
						If Err.Number < 0Then
							Fn_MyWorkList_ResourcePoolSubscription = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Failed to click ADD button") 
							objDialog.JavaButton("Cancel").Click micLeftBtn
							Set objDialog = Nothing
							Exit Function 
						End If
				
						''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						'Added code for handling Information Dialog  Added by : Harshal Tanpure. 25-March-2011
						'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						If  JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Resource Pool Subscription").JavaDialog("Information").Exist (5) Then
							 JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Resource Pool Subscription").JavaDialog("Information").JavaButton("OK").Click
							 If Err.Number < 0Then
								Fn_MyWorkList_ResourcePoolSubscription = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Failed to click OK button of Informatin Dialog") 
								objDialog.JavaButton("Cancel").Click micLeftBtn
								Set objDialog = Nothing
								Exit Function 
							End If
						End If
				
						If sInboxNode <> "" Then
							arrNodeList = split(sInboxNode,":")
							iItemCount = objDialog.JavaTree("ResPoolTree").GetROProperty( "items count")
							For iCounter=0 To (iItemCount-1)
								If Instr(1,arrNodeList(1),"Inbox") > 0 Then
									sTreeItem = objDialog.JavaTree("ResPoolTree").GetItem(iCounter)
									If  Instr(1,sTreeItem,arrNodeList(1)) > 0 Then
										sTreeItem = Split(sTreeItem,":")
										arrNodeList(1) =sTreeItem(1)
										sInboxNode = Join(arrNodeList,":")
									End If
								End If
							Next
				
							For iCounter=0 To (iItemCount-1)
								sTreeItem = objDialog.JavaTree("ResPoolTree").GetItem(iCounter)
								If Trim (Lcase(sTreeItem)) = Trim(Lcase(sInboxNode)) Then
									Fn_MyWorkList_ResourcePoolSubscription = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully add node " + sInboxNode )	
									Exit For
								End If
							Next
				
							If  Cint(iCounter) = Cint (iItemCount) Then
								Fn_MyWorkList_ResourcePoolSubscription = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to add node " +  sInboxNode)	
								objDialog.JavaButton("Cancel").Click micLeftBtn
								Set objDialog = Nothing
								Exit Function 
							 End If 
						End If						
					End If    
			
					objDialog.JavaButton("Cancel").Click micLeftBtn
				Else
					Fn_MyWorkList_ResourcePoolSubscription = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Resource Pool Subscription dialog does not exist.") 
					Set objDialog = Nothing
					Exit Function 
				End If 
			Else
				Fn_MyWorkList_ResourcePoolSubscription = True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully "+sInboxNode+" Exist in My Worklist.") 
			End If

		Case "Remove"
			bReturn = Fn_MyWorkList_TreeNodeOperations("Exist",sInboxNode,"")
			If bReturn = True Then
				'Select the Menu.
				bReturn = Fn_MenuOperation("Select","Tools:Resource Pool Subscription...")
				Call Fn_ReadyStatusSync(3)		   	    
				If bReturn = False Then
					Fn_MyWorkList_ResourcePoolSubscription = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fail to Select Menu Tools:Resource Pool Subscription...") 
					Exit Function
				Else 
					Fn_MyWorkList_ResourcePoolSubscription = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Menu Tools:Resource Pool Subscription...") 
				End If

				' Check  existance of the Dialog
				Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Resource Pool Subscription")
				wait 2
				 If objDialog.Exist(5) Then
					 arrNodeList = split(sInboxNode,":")
						'iItemCount = objDialog.JavaTree("ResPoolTree").GetROProperty( "items count")
						iItemCount= Fn_UI_Object_GetROProperty("Fn_MyWorkList_ResourcePoolSubscription",objDialog.JavaTree("ResPoolTree"),"items count")
						For iCounter=0 To (iItemCount-1)
							If Instr(1,arrNodeList(1),"Inbox") > 0 Then
								sTreeItem = objDialog.JavaTree("ResPoolTree").GetItem(iCounter)
								If  Instr(1,sTreeItem,arrNodeList(1)) > 0 Then
									sTreeItem = Split(sTreeItem,":")
									arrNodeList(1) =sTreeItem(1)
									sInboxNode = Join(arrNodeList,":")
								End If
							End If
						Next

						objDialog.JavaTree("ResPoolTree").Select sInboxNode

						If Err.Number < 0 Then
							Fn_MyWorkList_ResourcePoolSubscription = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select node from resource tree")	
							objDialog.JavaButton("Cancel").Click micLeftBtn
							Set objDialog = Nothing
							Exit Function 
						End If

						objDialog.JavaButton("Remove").WaitProperty "enabled",1,200000
						objDialog.JavaButton("Remove").Click micLeftBtn

						For iCounter=0 To (iItemCount-1)
							sTreeItem = objDialog.JavaTree("ResPoolTree").GetItem(iCounter)
							If Trim (Lcase(sTreeItem)) = Trim(Lcase(sInboxNode)) Then
								Fn_MyWorkList_ResourcePoolSubscription = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to remove node " +  sInboxNode )	
								Exit For
							End If
						Next

						If  Cint(iCounter) = Cint (iItemCount) Then
							Fn_MyWorkList_ResourcePoolSubscription = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully remove node " +  sInboxNode)	
							objDialog.JavaButton("Cancel").Click micLeftBtn
							Set objDialog = Nothing
							Exit Function 
						 End If 
				 Else
					Fn_MyWorkList_ResourcePoolSubscription = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Resource Pool Subscription dialog does not exist.") 
					Set objDialog = Nothing
					Exit Function 
				 End If 
			End If

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'Added Verify Case for Radio button Options  Added by : Harshal Tanpure. 11-April-2011
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Case "Verify"
				bReturn = Fn_MenuOperation("Select","Tools:Resource Pool Subscription...")
				Call Fn_ReadyStatusSync(3)
				If bReturn = False Then
					Fn_MyWorkList_ResourcePoolSubscription = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fail to Select Menu Tools:Resource Pool Subscription...") 
					Exit Function
				Else 
					Fn_MyWorkList_ResourcePoolSubscription = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Menu Tools:Resource Pool Subscription...") 
				End If

				' Check  existance of the Dialog
				Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Resource Pool Subscription")

				 If objDialog.Exist(5) Then

					If sOption <> "" Then
						aRdBOption = split (sOption,":",-1,1)

							For iCnt =0 to ubound (aRdBoption)
						
									    objDialog.JavaRadioButton("ResPoolOpt").SetTOProperty "attached text", aRdBoption(iCnt)
										If objDialog.JavaRadioButton("ResPoolOpt").Exist(5) then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Radio Button " + aRdBoption(iCnt) + " Exist ") 
											Fn_MyWorkList_ResourcePoolSubscription = True
										Else
											Fn_MyWorkList_ResourcePoolSubscription = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : " + aRdBoption(iCnt) + " radiobutton does not exist.") 
											objDialog.JavaButton("Cancel").Click micLeftBtn
											Set objDialog = Nothing
											Exit Function 
										End If
		
							Next

					End If
					objDialog.JavaButton("Cancel").Click micLeftBtn
					If Err.Number < 0 Then
							Fn_MyWorkList_ResourcePoolSubscription = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to click Cancel button" ) 
							Exit Function 
					End If

					Set objDialog = Nothing
				End If
			
    End Select
End Function



'*********************************************************  Function is Subscribe to any Resource Pool *********************************************************************

'Function Name		:					Fn_MyWorkList_TaskPromote

'Description			 :		 		   Promote Do task									

'Parameters			   :	 			 1. sTaskName: Task to be selected
'												2. sComment: Promote comments
'
'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Myworklist tab should selected.

'Examples				:			Fn_MyWorkList_TaskPromote("My Worklist:dba/CostDBA","Promote Do task")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										    Vallari				        23-Aug-2010	        1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_TaskPromote(sTaskName, sComment)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TaskPromote"
	Dim sMenu, bReturn, ObjPromoteWin

	On Error Resume Next

	sMenu = "Actions:Promote"
	If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").Exist Then
			set ObjPromoteWin = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Promote Action Comments")
	Else
			set ObjPromoteWin = Fn_SISW_WorkflowViewer_GetObject("Promote Action Comments")
	End If

	'Select MyWorklist Tree Node
	If sTaskName <> "" Then
		bReturn = Fn_MyWorkList_TreeNodeOperations("Select",sTaskName,"")
		If bReturn = False Then
			Fn_MyWorkList_TaskPromote = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select  " + sTaskName )	
			Exit Function
		End If
	End If

	bReturn = Fn_MenuOperation("Select", sMenu)
	Call Fn_ReadyStatusSync(3)
	If bReturn = False Then
		Fn_MyWorkList_TaskPromote = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke  Menu " + sMenu )	
		Exit Function
	End If

	If ObjPromoteWin.Exists(10) Then
        'Added by Rima Patil on 01-Aug-2012
		Err.Clear
		ObjPromoteWin.JavaEdit("Comment").Object.setText sComment
		'ObjPromoteWin.JavaEdit("Comment").Set sComment
'        ObjPromoteWin.JavaEdit("Comment").Object.Set sComment
		If Err.Number < 0 Then
			Fn_MyWorkList_TaskPromote = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Commnet " + sComment)	
			ObjPromoteWin.JavaButton("Cancel").Click micLeftBtn
			Set ObjPromoteWin = Nothing
			Exit Function
		End If

		ObjPromoteWin.JavaButton("OK").Click micLeftBtn
		Fn_MyWorkList_TaskPromote = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Promoted Task " + sTaskName)	
		Set ObjPromoteWin = Nothing

	Else
		Fn_MyWorkList_TaskPromote = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find  [Promote Action Comments] Dialog" )	
		Exit Function
	End If

End Function



'*******************************  Function is User can delegate his signoff responsibility to other user or resource pool *********************************************************************

'Function Name		:					Fn_MyWorkList_DelegateSignoff

'Description			 :		 		User can delegate his signoff responsibility to other user or resource pool														

'Parameters			   :	 			 1.sAction: Menu/ViewerPane
'                                        2.sDelLink: Original user name, who will delegate 
'                                        3.sOrgUser: User to whom to be delegated
'                                        4.sProject: Project name on the project tab 
'                                        5.sProjUser: Project user to be delegated to
'                                        6.sResPoolOpt: Resource Pool option
	
'Return Value		   : 			 True/False

'Pre-requisite			:		 	 Perform-signoff node is selected from the myworklist tree

'Examples				:			Fn_MyWorkList_DelegateSignoff("ViewerPane","AutoTest (autotest1)-Engineering/Designer","Engineering:Designer","","","Any Group")
'									Fn_MyWorkList_DelegateSignoff("VerifyUserGroupRole","AutoTest (autotest1)-Engineering/Designer","","","","")

'History:
'										Developer Name			Date				Rev. No.			Changes Done														Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										    Bharti				01-Sep-2010          1.0
'										    Shweta Rathod		30-Mar-2017          1.0        Added case VerifyUserGroupRole - to verify group / role value in table.     Shweta Rathod  
'																								Added case VerifyDeligateDlgExist - to check the existece iof dialog Shweta Rathod	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_DelegateSignoff(sAction,sDelLink,sOrgUser,sProject,sProjUser,sResPoolOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_DelegateSignoff"
    On Error Resume Next	
    Dim bReturn, objDialog,arrNode,sExpandNode, aResPool, objPerformSignOff
	Dim sDelInfo,sDelRow,iItemCount
	Dim objTable
	Select Case sAction
			Case "ViewerPane","VerifyUserGroupRole"
			    bReturn = Fn_MyTc_TabSet("Viewer")
				If bReturn = False Then
							Fn_MyWorkList_DelegateSignoff = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set Viewer tab") 		
							Exit Function						
				Else
							Fn_MyWorkList_DelegateSignoff = true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Viewer tab")
							Wait(2)
				End If
				Call Fn_ReadyStatusSync(6)
				If sDelLink <> "" Then
					If Instr(1,sDelLink,":") Then
						sDelInfo = split(sDelLink,":",-1,1)
						sDelRow = sDelInfo(1)
						sDelLink = sDelInfo(0)
					Else
						sDelRow = "0"
					End If
					
					Call Fn_ReadyStatusSync(6)
					'Verify sDelLink Node exists in the Table
					Set objTable = JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet").JavaTable("StoreProcTable")
					bReturn = Fn_MyWorkList_TableRowIndex(objTable, sDelLink,"User-Group/Role")
					If bReturn = -1 Then
						Call Fn_UpdateLogFiles("Action - WARNING  | Failed to Find [" + sDelLink + "] node in the PErform Signoff table", "WARNING:Expected node not found in Perform Signoff table")
						Set objTable = nothing
					End If
					
					'Added by shweta to verify "User-Group/Role node exist in the table.
					If sAction = "VerifyUserGroupRole" then
						If bReturn = -1 then
							Fn_MyWorkList_DelegateSignoff = false
						else
							Fn_MyWorkList_DelegateSignoff = true
						End if
						Exit function
					End if
					'Added by Vallari - Tc Hang was observed for below table click action
					wait(2)
					Call Fn_ReadyStatusSync(1)
					'Call Fn_MyTc_TabOperation("Close", "Viewer")
					Call Fn_MyTc_TabSet("Summary")
					Call Fn_MyTc_TabSet("Viewer")

					'JavaWindow("MyTeamcenter").JavaWindow("MyTcJApplet").JavaTable("StoreProcTable").SelectCell "#"+sDelRow,"User-Group/Role"
					objTable.SelectCell "#"+cstr(bReturn),"User-Group/Role"
					Call Fn_ReadyStatusSync(5)
					Set objTable = nothing
				End If
				
			Case "Menu"
				'Select the Menu.
				bReturn = Fn_MenuOperation("Select","Actions:Perform")
				Call Fn_ReadyStatusSync(5)
				If bReturn = False Then
					Fn_MyWorkList_DelegateSignoff = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Menu Actions:Perform") 
					Exit Function
				Else 
					Fn_MyWorkList_DelegateSignoff = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Menu Actions:Perform") 
				End If

                Call Fn_ReadyStatusSync(5)
                ' Check  existance of the Perform Dialog
				Set objPerformSignOff = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff")
                 If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").Exist(5) Then
					If Instr(1,sDelLink,":") Then
						sDelInfo = split(sDelLink,":",-1,1)
						sDelRow = sDelInfo(1) 
						sDelLink = sDelInfo(0) 
					Else
						sDelRow = 0
					End If
					Wait(5)
					'Verify sDelLink Node exists in the Table
					Set objTable = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaTable("SignOffTable")
					bReturn = Fn_MyWorkList_TableRowIndex(objTable, sDelLink, 0)
					If bReturn = -1 Then
						Call Fn_UpdateLogFiles("Action - WARNING  | Failed to Find [" + sDelLink + "] node in the Perform Signoff table", "WARNING:Expected node not found in Perform Signoff table")
						Set objTable = nothing
					End If

					JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaTable("SignOffTable").ClickCell cint(sDelRow),0
					wait(5)
					Call Fn_ReadyStatusSync(6)
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Actions:Perform Sign off Dialog did not appear")
					Fn_MyWorkList_DelegateSignoff = False
					Set objPerformSignOff = Nothing
					Exit Function
				End if
			Case "VerifyDeligateDlgExist"
				'Check  existance of the Delegate Dialog
					If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Delegate signoff").Exist(5) Then
						Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Delegate signoff")
					Else
						Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Delegate signoff")
					End If
					If objDialog.Exist(1) Then
						Fn_MyWorkList_DelegateSignoff = true
					else
						Fn_MyWorkList_DelegateSignoff = false
					End if
					Exit function
			Case else
				Fn_MyWorkList_DelegateSignoff = false
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Can Not Perform Action   " + sAction)
				Exit Function
			End Select

			If sDelLink <> "" Then
						'Check  existance of the Delegate Dialog
						If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Delegate signoff").Exist(5) Then
							Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Delegate signoff")
						Else
							Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Delegate signoff")
						End If
						If objDialog.Exist(5) Then
							If sOrgUser  <> "" Then
								If cstr(trim(objDialog.JavaTree("OrganizationTree").GetROProperty("value"))) = "" Then
									Call Fn_UpdateLogFiles("Action - WARNING  | Failed to Find [" + sDelLink + "] node already selected in Delegate Signoff Tree", "WARNING:Expected node not selected in Delegate Signoff Tree")
								End If

								arrNode=split(sOrgUser,":",-1,1)
								sExpandNode = "#0:"
								For iItemCount = 0 to UBound(arrNode)-1
									sExpandNode = sExpandNode + arrNode(iItemCount)
									objDialog.JavaTree("OrganizationTree").Expand sExpandNode
									wait(5)
									Call Fn_ReadyStatusSync(5)
									sExpandNode = sExpandNode + ":"
								Next
								'Select the Node
				       			objDialog.JavaTree("OrganizationTree").Select "#0:" + sOrgUser								
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + sOrgUser  + " from Oraganization Tree")
									Fn_MyWorkList_DelegateSignoff = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									If objPerformSignOff.Exist(5) Then
					                               JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
												   wait(3)
				                    End if
									Set objPerformSignOff = Nothing
									Set objDialog = Nothing
									Exit Function				
								Else	
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Node " + sOrgUser  + " from Oraganization Tree")
									Fn_MyWorkList_DelegateSignoff = True
								End If	
                            End If
                            ' Select the Project Team tab	
					        If sProject  <> "" Then
								objDialog.JavaTab("Tab").Select "Project Teams"
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select tab Project Teams from Delegate Sign Off.")
									Fn_MyWorkList_DelegateSignoff = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									If objPerformSignOff.Exist(5) Then
					                          JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
											  wait(3)
				                     End if
									 Set objPerformSignOff = Nothing
									Set objDialog = Nothing
									Exit Function
								Else	
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected tab Project Teams from Delegate Sign Off.")
									Fn_MyWorkList_DelegateSignoff = true				
								End If							

                            	'Select the Project                              
								objDialog.JavaList("ProjectsList").Select sProject
								Call Fn_ReadyStatusSync(1)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Project "+ sProject + " from Delegate Sign off.")
									Fn_MyWorkList_DelegateSignoff = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									If objPerformSignOff.Exist(5) Then
					                                JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
													wait(3)
				                    End if
									Set objPerformSignOff = Nothing
									Set objDialog = Nothing
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Project " + sProject + " from Delegate Sign Off.")
									Fn_MyWorkList_DelegateSignoff = true			
								End If

								'Expand the node
								arrNode = split(aProjUser,":",-1,1)
								sExpandNode = "#0:"
								For iItemCount = 0 to UBound(arrNode)-1
									sExpandNode = sExpandNode + arrNode(iItemCount)
									objDialog.JavaTree("ProjectsTree").Expand sExpandNode
									wait(2)
									Call Fn_ReadyStatusSync(5)
									sExpandNode = sExpandNode + ":"
								Next

								'Select the Project Member
								If  sProjUser <> "" Then
									objDialog.JavaTree("ProjectsTree").Select "#0:" + sProjUser												
									If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + sProjUser  + " from Projects Tree")
											Fn_MyWorkList_DelegateSignoff = False
									 		objDialog.JavaButton("Cancel").Click micLeftBtn
											If objPerformSignOff.Exist(5) Then
					                                    JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
														wait(3)
				                            End if
											Set objPerformSignOff = Nothing
											Set objDialog = Nothing
											Exit Function
									Else	
											Call Fn_ReadyStatusSync(5)
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Node " + sProjUser  + " from Projects Tree")
											Fn_MyWorkList_DelegateSignoff = true
									End If
								End If
							End if
							'Set the Resource Pool Option											
					        If sResPoolOpt <> "" Then
									objDialog.JavaRadioButton("ResPoolOption").SetTOProperty "attached text",sResPoolOpt
									objDialog.JavaRadioButton("ResPoolOption").Set "ON"
									wait(1)
									If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Resource Pool Option " + sResPoolOpt)
													Fn_MyWorkList_DelegateSignoff = False
													objDialog.JavaButton("Cancel").Click micLeftBtn
													wait(2)		
													If objPerformSignOff.Exist(5) Then
					                                      JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
														  wait(3)
				                                    End if	
													Set objPerformSignOff = Nothing										
													Set objDialog = Nothing
													Exit Function
									Else	
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected  Resource Pool Option " + sResPoolOpt)
													Fn_MyWorkList_DelegateSignoff = true
									End If												 
					        End If
							
                          Call Fn_ReadyStatusSync(6)
                          objDialog.JavaButton("OK").Click micLeftBtn
						  wait(3)

						 'Handled Duplicate Reviewer warning dialog - Observed afetr Tc10_0611 build
						  If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("PerformSignoffWarning").Exist(5) Then
							  If instr(JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("PerformSignoffWarning").JavaEdit("MsgText").GetROProperty("value"), "Duplicate Reviewer") > 0 Then
								  JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("PerformSignoffWarning").JavaButton("OK").Click micLeftBtn
							  End If
						  End If

						  Call Fn_ReadyStatusSync(6)
							
                         If Err.Number < 0 Then
							   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK button")
							   Fn_MyWorkList_DelegateSignoff = False
                               If objPerformSignOff.Exist(5) Then
					                            JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
												wait(3)
				               End if
							   Set objPerformSignOff = Nothing
							  Set objDialog = Nothing
				              Exit Function
	                    Else	
				               Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Ok button")
							   Call Fn_ReadyStatusSync(5)
				               Fn_MyWorkList_DelegateSignoff = true
						End If
                          						
					Else 
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Delegate Sign Off Dialog did not appear")
						Fn_MyWorkList_DelegateSignoff = False
'						Set objPerformSignOff = Nothing
'						Set objDialog = Nothing	    
					End If
				End if
				If objPerformSignOff.Exist(5) Then
						JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
						wait(5)
				End if

				Set objPerformSignOff = Nothing
				Set objDialog = Nothing
End Function
		      
'*********************************************************  Function do Operation on WorkList Tree *********************************************************************

'Function Name		:					Fn_MyWorkList_ViewAuditLog

'Description			 :		 		    View and verifies audit log for the objects in process
'											
'Parameters			   :	 				sAction: More/Less or ""
'													aLog: Array of log statements to be verified
' 												   

'Return Value		   : 			 		True/False

'Pre-requisite			:		 	 		 None

'Examples				:			 		 arrLog = Array("Process Name              003784/A;1-UVWX","Description","Process Template Name     AutoSimpleReview","","Date Created              2010-09-03","","","Start            AutoSimpleReview          2010-09-03   autotest3")
'												   Fn_MyWorkList_ViewAuditLog("",arrLog)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Prasanna				06-Sep-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_ViewAuditLog(sAction,aLog)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_ViewAuditLog"
   On Error Resume Next
   Dim arrNodeList,iItemCount,iCounter,sTreeItem,arrNode,iOuterCount,aMenuList
   Dim objSignOffTree,sAuditText,arrAuditLine,sDateString,aDate

    If  JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").Exist(5) = True Then
			Set  objAuditDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("View Audit File")
	Else      
			Set  objAuditDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("View Audit File")
	End If

	If objAuditDialog.Exist(5) = false Then
				' Open the View --> Audit --> File
		   	    bReturn = Fn_MenuOperation("Select","View:Audit:File")	
				Call Fn_ReadyStatusSync(5)		   	    
				If bReturn = False Then
						Fn_MyWorkList_ViewAuditLog = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  Select Menu [View:Audit:File].")
						Set objAuditDialog = Nothing
						Exit Function
				Else
						Call Fn_ReadyStatusSync(3)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Menu [View:Audit:File].")
				End If
	End If
	If sAction <> "" Then
		Select Case sAction
				
					Case "More"
								objAuditDialog.JavaCheckBox("MoreLessOption").SetTOProperty "Attached Text","More..."
								objAuditDialog.JavaCheckBox("MoreLessOption").Set "ON"							
								Wait(2)
								If Err.Number < 0 Then											
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on More Button.")
											objAuditDialog.JavaButton("Cancel").Click micLeftBtn
											Fn_MyWorkList_ViewAuditLog = false
											Set objAuditDialog = Nothing
											Exit Function
								Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Failed to Click on More Button.")
								End If	
							
					Case "Less"
								objAuditDialog.JavaCheckBox("MoreLessOption").SetTOProperty "Attached Text","Less..."
								objAuditDialog.JavaCheckBox("MoreLessOption").Set "ON"
								Wait(2)
								If Err.Number < 0 Then											
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Less Button")
											objAuditDialog.JavaButton("Cancel").Click micLeftBtn
											Fn_MyWorkList_ViewAuditLog = false
											Set objAuditDialog = Nothing
											Exit Function
								Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Failed to Click on Less Button")
								End If  				
		   End Select
	End If

		If Ubound(aLog) < 0  Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : The Input Array is Empty.")
				objAuditDialog.JavaButton("Cancel").Click micLeftBtn
				Fn_MyWorkList_ViewAuditLog = false
				Set objAuditDialog = Nothing
				Exit Function
		End If
		
		sAuditText = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("View Audit File").JavaEdit("AuditLog").Object.getText()

		'Split the array with delimeter Char(10)
		arrAuditLine = split(trim(sAuditText),Chr(10),-1,1)

		For iCounter = 0 to Ubound(aLog)
				If aLog(iCounter) <> ""  Then

						'Check Whether string Contain Date & time
						If  instr(1,arrAuditLine(iCounter),":") Then
									aDate = split(arrAuditLine(iCounter), ":", -1,1)
									sDateString = left(aDate(0), len(aDate(0))-2) +Right(aDate(2), len(aDate(2))-2)
									arrAuditLine(iCounter) = sDateString
						End If

						'Check the Input String & Log from Audit
						If trim(aLog(iCounter)) = trim(arrAuditLine(iCounter)) Then
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Line Verified Successfully : " + aLog(iCounter))                      								 
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Line : " + aLog(iCounter))
								objAuditDialog.JavaButton("Cancel").Click micLeftBtn
								Fn_MyWorkList_ViewAuditLog = false
								Set objAuditDialog = Nothing
								Exit Function
						End If
						
				 End If
		Next

        Fn_MyWorkList_ViewAuditLog = true
		'Close the dialog
		If  JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").Exist(2) = True Then
				JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("View Audit File").JavaButton("Cancel").Click micLeftBtn
		Else      
				JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("View Audit File").JavaButton("Cancel").Click  micLeftBtn
		End If
		
		If Err.Number < 0 Then					
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Cancel Button.")					
					Fn_MyWorkList_ViewAuditLog = false
					Set objAuditDialog = Nothing
					Exit Function
		Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on More Button.")
		End If	
		Set objAuditDialog = Nothing

End Function

'*********************************************************		Function to  date in Schedule Manager Required Format	***********************************************************************

'Function Name		:					Fn_MyWorkList_FormatDate

'Description			 :		 		  This function is used to get date in Schedule Manager Required Format.

'Parameters			   :	 			1.  sDate:date value to provide. 
											
'Return Value		   : 				 Date in the format YYYY-MM-DD (2010-09-03)

'Examples				:				 sDate = cstr(now)
'												call Fn_MyWorkList_FormatDate(date)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Prasanna.					08-Sep-2010	   		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_FormatDate(sDate)
   Dim sDay, sMonth, sYear

   On Error Resume Next

   sDay = Day(sDate)
   sMonth = Month(sDate) 
   sYear = Year(sDate)
   
	If len(sDay) = 1 Then
		sDay = "0" + cstr(sDay)
	End If

	If len(sMonth) = 1 Then
		 sMonth = "0" + cstr(sMonth)
	End If

	 Fn_MyWorkList_FormatDate = cstr(sYear) + "-"+ cstr(sMonth) + "-"+ cstr(sDay)  
End Function

'*********************************************************  Function do Operation on WorkList Workflow Surrogate Dialog *********************************************************************

'Function Name		:			Fn_MyWorkList_WorkflowSurrogate

'Description		:		 	Sets or verifies surrogate user's inbox:-
'								Case "Add" - Add all the inputs accordignly & click on Add
'								Case "Modify" - Add all the inputs accordignly & click on Modify
'								Case "Remove" - Add all the inputs accordignly & click on Remove
'								Case "Verify" - Verify all the inputs if supplied in & click on Cancel button

'Parameters			:		 	1. sAction: Add/Modify/Remove/Verify
'								2. sOrgGrp, sOrgRole & sOrgUser: Optional params. Set if needed in. Displayed in dba login
'								3. sFrmDt: From date & time
'								4. sToDt: To date & time
'								5. sNewGrp: New group to whom signoff is delegated to.
'								6. sNewRole: New role to whom signoff is delegated to.
'								7. sNewUser: New user to whom signoff is delegated to. 
'								8. aSurrogateUsers: array of users to be surrogated to.

'Return Value		: 			True/False

'Pre-requisite		:		 	Pre-Requisite: Myworklist tree is displayed in MyTc module.

'Examples			:			Call Fn_MyWorkList_WorkflowSurrogate("Add", "Engineering", "Designer" , "Mahendra Bhandarkar(x_bhanda)", "14-Oct-2010 11:43", "14-Nov-2010 11:43", "Engineering", "Designer", "Mahendra Bhandarkar(x_bhanda)",array(""))
'								Call Fn_MyWorkList_WorkflowSurrogate("Verify", "dba", "DBA" , "AutoTestDBA (autotestdba)", "14-Sep-2010 11:43", "", "dba", "DBA", "AutoTest7 (autotest7)",array(""))
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Mahendra Bhandarkar		15-Sep-2010	       	1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Madhura P				14-Apr-2015	       	1.0				Modified code to handle Date control a/c to designed change		Paresh
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_WorkflowSurrogate(sAction, sOrgGrp, sOrgRole, sOrgUser, sFrmDt, sToDt, sNewGrp, sNewRole, sNewUser, aSurrogateUsers)
	
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_WorkflowSurrogate"
	Dim objDialog, bReturn, iCounter, objErrDialog, iCounter2, iCntRows, iCount, sItem, arrDate
	Dim WshShell
	Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Workflow Surrogate")
	Set objErrDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Error")
	If objDialog.Exist(5) = False Then 
		bReturn = Fn_MenuOperation("Select", "Tools:Workflow Surrogate...")
		Call Fn_ReadyStatusSync(5)		   	    
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Tools --> Workflow Surrogate... ") 		
			Fn_MyWorkList_WorkflowSurrogate = False
			Set objDialog = Nothing
			Exit Function					
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Tools --> Workflow Surrogate... ")
			Call Fn_ReadyStatusSync(2)
		End If
	End If	

	If objDialog.Exist(5) Then

		Select Case sAction

		Case "Add"

			If Trim(sOrgGrp) <> "" Then
				objDialog.JavaEdit("OrgGroup").Set sOrgGrp
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sOrgGrp)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sOrgGrp)+"]")
				End If
			End If

			If Trim(sOrgRole) <> "" Then
				objDialog.JavaEdit("OrgRole").Set sOrgRole
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Role ["+CStr(sOrgRole)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Role ["+CStr(sOrgRole)+"]")
				End If
			End If

			If Trim(sOrgUser) <> "" Then
				objDialog.JavaEdit("OrgUser").Set sOrgUser
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sOrgUser)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sOrgUser)+"]")
				End If
			End If

			If Trim(sFrmDt) <> "" Then

				'objDialog.JavaCheckBox("FromDate").Object.setDate Trim(sFrmDt)
				arrDate = Split(sFrmDt," ")
							objDialog.JavaEdit("FromDate").RefreshObject
							wait 1
							objDialog.JavaEdit("FromDate").Click 1,1
							wait 7
							objDialog.JavaEdit("FromDate").Set arrDate(0)
							wait 2
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							objDialog.JavaList("FromDate").Select arrDate(1)

				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select From Date ["+CStr(sFrmDt)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected From Date ["+CStr(sFrmDt)+"]")
				End If
			End If

			If Trim(sToDt) <> "" Then
				'objDialog.JavaCheckBox("ToDate").Object.setDate Trim(sToDt)
				arrDate = Split(sToDt," ")
							objDialog.JavaEdit("ToDate").RefreshObject
							wait 1
							objDialog.JavaEdit("ToDate").Click 1,1
							wait 7
							objDialog.JavaEdit("ToDate").Set arrDate(0)
							wait 2
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							objDialog.JavaList("ToDate").Select arrDate(1)

				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select To Date ["+CStr(sToDt)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected To Date ["+CStr(sToDt)+"]")
				End If
			End If

			If Trim(sNewGrp) <> "" Then
				objDialog.JavaEdit("NewGroup").Set sNewGrp
				objDialog.JavaEdit("NewGroup").Click 0,0
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(5)
				Call Fn_ReadyStatusSync(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sNewGrp)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sNewGrp)+"]")
				End If
			End If

			If Trim(sNewRole) <> "" Then
				objDialog.JavaEdit("NewRole").Set sNewRole
				objDialog.JavaEdit("NewRole").Click 0,0
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(5)
				Call Fn_ReadyStatusSync(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Role ["+CStr(sNewRole)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Role ["+CStr(sNewRole)+"]")
				End If
			End If

			If Trim(sNewUser) <> "" Then
				objDialog.JavaEdit("NewUser").Set sNewUser
				objDialog.JavaEdit("NewUser").Click 0,0
				Wait(5)
				Call Fn_ReadyStatusSync(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sNewUser)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sNewUser)+"]")
				End If
			End If

			objDialog.JavaButton("Add").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Add Button ") 		
				Fn_MyWorkList_WorkflowSurrogate = False
				objDialog.JavaButton("Close").Click micLeftBtn
				Set objDialog = Nothing
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Add Button")
				Call Fn_ReadyStatusSync(1)
			End If

				'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				'Added by : Harshal Tanpure. Date : 07-April-2011
				
				 If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Surrogate Error_Exists").Exist(5) Then
						JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Surrogate Error_Exists").JavaButton("OK").Click
						If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK Button of Surrogate Exists Dialog") 		
								Fn_MyWorkList_WorkflowSurrogate = False
								Set objDialog = Nothing
								Exit Function
						Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on OK Button of Surrogate Exists Dialog")
						End If
				End If 
				
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			' Handling Error Dialog of Invalid Input Dates
		 	objErrDialog.SetTOProperty "title", "Invalid input dates ..."
			If objErrDialog.Exist(5) = True Then
					Fn_MyWorkList_WorkflowSurrogate = False
					Do While objErrDialog.JavaButton("OK").Exist = True
					objErrDialog.JavaButton("OK").Click micLeftBtn
					Loop
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error Invalid input dates... Dialog Existance Verified ")
					wait(2)
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Set objErrDialog = Nothing
					Exit Function			
			End If

			'Handling Error Dialog of Invalid Surrogate
			objErrDialog.SetTOProperty "title", "Invalid Surrogate..."
			If objErrDialog.Exist = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error Invalid Surrogate... Dialog Existance Verified ")
					Fn_MyWorkList_WorkflowSurrogate = False
					Do While objErrDialog.JavaButton("OK").Exist = True
					objErrDialog.JavaButton("OK").Click micLeftBtn
					Loop
					wait(2)
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Set objErrDialog = Nothing
					Exit Function			
			End If

			'Handling Error Dialog of Surrogate Exists
			objErrDialog.SetTOProperty "title", "Surrogate Exists..."
			If objErrDialog.Exist = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error Surrogate Exists... Dialog Existance Verified ")
					Fn_MyWorkList_WorkflowSurrogate = True
					Do While objErrDialog.JavaButton("OK").Exist = True
					objErrDialog.JavaButton("OK").Click micLeftBtn
					Loop
					wait(2)
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Set objErrDialog = Nothing
					Exit Function			
			End If

		Case "Modify"

			If Trim(sOrgGrp) <> "" Then
				objDialog.JavaEdit("OrgGroup").Set sOrgGrp
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sOrgGrp)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sOrgGrp)+"]")
				End If
			End If

			If Trim(sOrgRole) <> "" Then
				objDialog.JavaEdit("OrgRole").Set sOrgRole
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Role ["+CStr(sOrgRole)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Role ["+CStr(sOrgRole)+"]")
				End If
			End If

			If Trim(sOrgUser) <> "" Then
				objDialog.JavaEdit("OrgUser").Set sOrgUser
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sOrgUser)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sOrgUser)+"]")
				End If
			End If

			If IsArray(aSurrogateUsers) = True Then
				For iCounter = 0 To UBound(aSurrogateUsers)
					If iCounter = 0 Then
						objDialog.JavaList("UserList").Select Trim(aSurrogateUsers(iCounter))
					Else
						objDialog.JavaList("UserList").ExtendSelect Trim(aSurrogateUsers(iCounter))
					End If
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ["+aSurrogateUsers(iCounter)+"] from Users List ") 		
						Fn_MyWorkList_WorkflowSurrogate = False
						objDialog.JavaButton("Close").Click micLeftBtn
						Set objDialog = Nothing
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected ["+aSurrogateUsers(iCounter)+"] from Users List ")
						Call Fn_ReadyStatusSync(1)
					End If
				Next
			End If

			If Trim(sFrmDt) <> "" Then
				'objDialog.JavaCheckBox("FromDate").Object.setDate Trim(sFrmDt)
				arrDate = Split(sFrmDt," ")
							objDialog.JavaEdit("FromDate").RefreshObject
							wait 1
							objDialog.JavaEdit("FromDate").Click 1,1
							wait 7
							objDialog.JavaEdit("FromDate").Set arrDate(0)
							wait 2
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							objDialog.JavaList("FromDate").Select arrDate(1)
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select From Date ["+CStr(sFrmDt)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected From Date ["+CStr(sFrmDt)+"]")
				End If
			End If

			If Trim(sToDt) <> "" Then
				'objDialog.JavaCheckBox("ToDate").Object.setDate Trim(sToDt)
				arrDate = Split(sToDt," ")
							objDialog.JavaEdit("ToDate").RefreshObject
							wait 1
							objDialog.JavaEdit("ToDate").Click 1,1
							wait 7
							objDialog.JavaEdit("ToDate").Set arrDate(0)
							wait 2
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							wait 2
							objDialog.JavaList("ToDate").Select arrDate(1)
							WshShell.SendKeys "{TAB}"			'Call added to reflect time in application through automation
							Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select To Date ["+CStr(sToDt)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected To Date ["+CStr(sToDt)+"]")
				End If
			End If

			If Trim(sNewGrp) <> "" Then
				objDialog.JavaEdit("NewGroup").Set sNewGrp
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sNewGrp)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sNewGrp)+"]")
				End If
			End If

			If Trim(sNewRole) <> "" Then
				objDialog.JavaEdit("NewRole").Set sNewRole
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Role ["+CStr(sNewRole)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Role ["+CStr(sNewRole)+"]")
				End If
			End If

			If Trim(sNewUser) <> "" Then
				objDialog.JavaEdit("NewUser").Type sNewUser
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sNewUser)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sNewUser)+"]")
				End If
			End If

			objDialog.JavaButton("Modify").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Modify Button ") 		
				Fn_MyWorkList_WorkflowSurrogate = False
				objDialog.JavaButton("Close").Click micLeftBtn
				Set objDialog = Nothing
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Modify Button")
				Call Fn_ReadyStatusSync(1)
			End If

			' Handling Error Dialog of Invalid Input Dates
		 	objErrDialog.SetTOProperty "text", "Invalid input dates ..."
			If objErrDialog.Exist(5) = True Then
					Fn_MyWorkList_WorkflowSurrogate = False
					Do While objErrDialog.WinButton("OK").Exist = True
					objErrDialog.WinButton("OK").Click micLeftBtn
					Loop
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error Invalid input dates... Dialog Existance Verified ")
					wait(2)
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Set objErrDialog = Nothing
					Exit Function			
			End If

			'Handling Error Dialog of Invalid Surrogate
			objErrDialog.SetTOProperty "text", "Invalid Surrogate..."
			If objErrDialog.Exist = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error Invalid Surrogate... Dialog Existance Verified ")
					Fn_MyWorkList_WorkflowSurrogate = False
					Do While objErrDialog.WinButton("OK").Exist = True
					objErrDialog.WinButton("OK").Click micLeftBtn
					Loop
					wait(2)
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Set objErrDialog = Nothing
					Exit Function			
			End If

			'Handling Error Dialog of Surrogate Exists
			objErrDialog.SetTOProperty "text", "Surrogate Exists..."
			If objErrDialog.Exist = True Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error Surrogate Exists... Dialog Existance Verified ")
					Fn_MyWorkList_WorkflowSurrogate = False
					Do While objErrDialog.WinButton("OK").Exist = True
					objErrDialog.WinButton("OK").Click micLeftBtn
					Loop
					wait(2)
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Set objErrDialog = Nothing
					Exit Function			
			End If
		Case "SurrogateExistSelect"
			If IsArray(aSurrogateUsers) = True Then
				For iCounter = 0 To UBound(aSurrogateUsers)
					If iCounter = 0 Then
						objDialog.JavaList("UserList").Select Trim(aSurrogateUsers(iCounter))
					Else
						objDialog.JavaList("UserList").ExtendSelect Trim(aSurrogateUsers(iCounter))
					End If
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ["+aSurrogateUsers(iCounter)+"] from Users List ") 		
						Fn_MyWorkList_WorkflowSurrogate = False
						objDialog.JavaButton("Close").Click micLeftBtn
						Set objDialog = Nothing
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected ["+aSurrogateUsers(iCounter)+"] from Users List ")
						Call Fn_ReadyStatusSync(1)
						Fn_MyWorkList_WorkflowSurrogate = True
						Exit Function
					End If
				Next
			End If
		Case "Remove"

			If Trim(sOrgGrp) <> "" Then
				objDialog.JavaEdit("OrgGroup").Set sOrgGrp
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sOrgGrp)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sOrgGrp)+"]")
				End If
			End If

			If Trim(sOrgRole) <> "" Then
				objDialog.JavaEdit("OrgRole").Set sOrgRole
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Role ["+CStr(sOrgRole)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Role ["+CStr(sOrgRole)+"]")
				End If
			End If

			If Trim(sOrgUser) <> "" Then
				objDialog.JavaEdit("OrgUser").Type sOrgUser
				Set WshShell = CreateObject("WScript.Shell")
				WAIT(3)
				WshShell.SendKeys "{ENTER}"
				Set WshShell = Nothing
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sOrgUser)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sOrgUser)+"]")
				End If
			End If

			If IsArray(aSurrogateUsers) = True Then
				For iCounter = 0 To UBound(aSurrogateUsers)
					If iCounter = 0 Then
						objDialog.JavaList("UserList").Select Trim(aSurrogateUsers(iCounter))
					Else
						objDialog.JavaList("UserList").ExtendSelect Trim(aSurrogateUsers(iCounter))
					End If
					If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ["+aSurrogateUsers(iCounter)+"] from Users List ") 		
						Fn_MyWorkList_WorkflowSurrogate = False
						objDialog.JavaButton("Close").Click micLeftBtn
						Set objDialog = Nothing
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected ["+aSurrogateUsers(iCounter)+"] from Users List ")
						Call Fn_ReadyStatusSync(1)
					End If
				Next
			End If

			If Trim(sFrmDt) <> "" Then
				'objDialog.JavaCheckBox("FromDate").Object.setDate Trim(sFrmDt)
				arrDate = Split(sFrmDt," ")
							objDialog.JavaEdit("FromDate").RefreshObject
							wait 1
							objDialog.JavaEdit("FromDate").Click 1,1
							wait 7
							objDialog.JavaEdit("FromDate").Set arrDate(0)
							wait 2
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							objDialog.JavaList("FromDate").Select arrDate(1)
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select From Date ["+CStr(sFrmDt)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected From Date ["+CStr(sFrmDt)+"]")
				End If
			End If

			If Trim(sToDt) <> "" Then
				'objDialog.JavaCheckBox("ToDate").Object.setDate Trim(sToDt)
				arrDate = Split(sToDt," ")
							objDialog.JavaEdit("ToDate").RefreshObject
							wait 1
							objDialog.JavaEdit("ToDate").Click 1,1
							wait 7
							objDialog.JavaEdit("ToDate").Set arrDate(0)
							wait 2
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ESC}"
							Set WshShell = Nothing
							wait 2
							objDialog.JavaList("ToDate").Select arrDate(1)
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select To Date ["+CStr(sToDt)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected To Date ["+CStr(sToDt)+"]")
				End If
			End If

			'Click on Remove button
			objDialog.JavaButton("Remove").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Remove Button ") 		
				Fn_MyWorkList_WorkflowSurrogate = False
				objDialog.JavaButton("Close").Click micLeftBtn
				Set objDialog = Nothing
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Remove Button")
				Call Fn_ReadyStatusSync(1)
			End If
		
		Case "Verify"

			If Trim(sOrgGrp) <> "" Then
					objDialog.JavaEdit("OrgGroup").Set sOrgGrp
					Set WshShell = CreateObject("WScript.Shell")
					WAIT(3)
					WshShell.SendKeys "{ENTER}"
					Set WshShell = Nothing
				If objDialog.JavaEdit("OrgGroup").GetROProperty("text") <> Trim(sOrgGrp) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Group ["+CStr(sOrgGrp)+"] ") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Group ["+CStr(sOrgGrp)+"]")
					Wait(2)
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sOrgRole) <> "" Then
					objDialog.JavaEdit("OrgRole").Set sOrgRole
					Set WshShell = CreateObject("WScript.Shell")
					WAIT(3)
					WshShell.SendKeys "{ENTER}"
					Set WshShell = Nothing

				If objDialog.JavaEdit("OrgRole").GetROProperty("text") <> Trim(sOrgRole) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Role ["+CStr(sOrgRole)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Role ["+CStr(sOrgRole)+"]")
					Wait(2)
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sOrgUser) <> "" Then
					objDialog.JavaEdit("OrgUser").Set sOrgUser
					Set WshShell = CreateObject("WScript.Shell")
					WAIT(3)
					WshShell.SendKeys "{ENTER}"
					Set WshShell = Nothing

				If objDialog.JavaEdit("OrgUser").GetROProperty("text") <> Trim(sOrgUser) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sOrgUser)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sOrgUser)+"]")
					Wait(2)
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If IsArray(aSurrogateUsers) = True Then
					iCount = 0
					iCntRows = objDialog.JavaList("UserList").GetROProperty("items count")
					For iCounter = 0 To Cint(iCntRows) - 1
						sItem = objDialog.JavaList("UserList").GetItem(iCounter)
						For iCounter2 = 0 To UBound(aSurrogateUsers)
							If aSurrogateUsers(iCounter2) = sItem Then
								iCount = iCount + 1
								Exit For
							End If
						Next
					Next

					If iCount <> UBound(aSurrogateUsers)+1 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select [" + Join(aSurrogateUsers, ",") + "] from Users List ")
						Fn_MyWorkList_WorkflowSurrogate = False
						objDialog.JavaButton("Close").Click micLeftBtn
						Set objDialog = Nothing
						Exit Function		
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified [" + Join(aSurrogateUsers, ",") +"] from Users List ")
						Call Fn_ReadyStatusSync(2)
					End If

					If iCount = 1 Then
						objDialog.JavaList("UserList").Select Trim(aSurrogateUsers(0))
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ["+aSurrogateUsers(0)+"] from Users List ") 		
							Fn_MyWorkList_WorkflowSurrogate = False
							objDialog.JavaButton("Close").Click micLeftBtn
							Set objDialog = Nothing
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected ["+aSurrogateUsers(0)+"] from Users List ")
							Call Fn_ReadyStatusSync(1)
						End If
					End If
			End If

			If Trim(sFrmDt) <> "" Then
	'			If InStr(1, objDialog.JavaCheckBox("FromDate").GetROProperty("label"), Trim(sFrmDt), 1) > 0 Then
				If InStr(1, trim(objDialog.JavaEdit("FromDate").GetROProperty("value")+" "+objDialog.JavaList("FromDate").GetROProperty("value")), Trim(sFrmDt), 1) > 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected From Date ["+CStr(sFrmDt)+"]")
					Call Fn_ReadyStatusSync(2)
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select From Date ["+CStr(sFrmDt)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				End If
			End If

			If Trim(sToDt) <> "" Then
	'			If InStr(1, objDialog.JavaCheckBox("ToDate").GetROProperty("label"), Trim(sToDt), 1) > 0 Then
				If InStr(1, trim(objDialog.JavaEdit("ToDate").GetROProperty("value")+" "+objDialog.JavaList("ToDate").GetROProperty("value")), Trim(sToDt), 1) > 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected To Date ["+CStr(sToDt)+"]")
					Call Fn_ReadyStatusSync(2)
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select To Date ["+CStr(sToDt)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				End If
			End If

			If Trim(sNewGrp) <> "" Then
				If objDialog.JavaEdit("NewGroup").GetROProperty("text") <> Trim(sNewGrp) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "WARNING : Failed to Verify Group ["+CStr(sNewGrp)+"]") 
					Call Fn_UpdateLogFiles("WARNING : Failed to Verify Group ["+CStr(sNewGrp)+"]", "")		
'					Fn_MyWorkList_WorkflowSurrogate = False
'					objDialog.JavaButton("Close").Click micLeftBtn
'					Set objDialog = Nothing
'					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Group ["+CStr(sNewGrp)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sNewRole) <> "" Then
				If objDialog.JavaEdit("NewRole").GetROProperty("text") <> Trim(sNewRole) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "WARNING : Failed to Verify Role ["+CStr(sNewRole)+"]") 	
					Call Fn_UpdateLogFiles("WARNING : Failed to Verify Group ["+CStr(sNewRole)+"]", "")			
'					Fn_MyWorkList_WorkflowSurrogate = False
'					objDialog.JavaButton("Close").Click micLeftBtn
'					Set objDialog = Nothing
'					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Role ["+CStr(sNewRole)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sNewUser) <> "" Then
				If objDialog.JavaEdit("NewUser").GetROProperty("text") <> Trim(sNewUser) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sNewUser)+"]") 		
					Fn_MyWorkList_WorkflowSurrogate = False
					objDialog.JavaButton("Close").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sNewUser)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

		End Select

		objDialog.JavaButton("Close").Click micLeftBtn
		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Close Button ") 		
			Fn_MyWorkList_WorkflowSurrogate = False
			Set objDialog = Nothing
			Exit Function		
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Close Button")
			Call Fn_ReadyStatusSync(2)
			Fn_MyWorkList_WorkflowSurrogate = True			
		End If
	
	Else

		Fn_MyWorkList_WorkflowSurrogate = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: MyWorkList Workflow Surrogate... function failed")
	End If
		
	Set objDialog = Nothing
	Set objErrDialog = Nothing

End Function
'*********************************************************  Function do Operation on WorkList Out Of Office Assistant Dialog *********************************************************************

'Function Name		:			Fn_MyWorkList_OutOfOfficeAssit

'Description		:		 	Sets or verifies Out of Office assitance:-
'								Case "Set" - Set all the inputs accordignly & click on OK
'								Case "Verify" - Verify all the inputs if supplied in & click on Cancel button

'Parameters			:		 	1. sAction: Set/Verify
'								2. sOrgGrp, sOrgRole & sOrgUser: Optional params. Set if needed in. Displayed in dba login
'								3. sFrmDt: From date & time
'								4. sToDt: To date & time
'								5. sNewGrp: New group to whom signoff is delegated to.
'								6. sNewRole: New role to whom signoff is delegated to.
'								7. sNewUser: New user to whom signoff is delegated to. 

'Return Value		: 			True/False

'Pre-requisite		:		 	Myworklist tree is displayed in MyTc module.

'Examples			:			Call Fn_MyWorkList_OutOfOfficeAssit("Set", "Engineering", "Designer" , "Mahendra Bhandarkar(x_bhanda)", "14-Oct-2010 11:43", "14-Nov-2010 11:43", "Engineering", "Designer", "Mahendra Bhandarkar(x_bhanda)")
'								Call Fn_MyWorkList_OutOfOfficeAssit("Verify", "dba", "DBA" , "AutoTestDBA (autotestdba)", "14-Sep-2010 11:43", "", "dba", "DBA", "AutoTest7 (autotest7)")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Mahendra Bhandarkar		14-Sep-2010	       	1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_OutOfOfficeAssit(sAction, sOrgGrp, sOrgRole , sOrgUser, sFrmDt, sToDt, sNewGrp, sNewRole, sNewUser)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_OutOfOfficeAssit"
	Dim objDialog, bReturn,adate
	Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Out Of Office Assistant")

	If objDialog.Exist(5) = False Then 
		bReturn = Fn_MenuOperation("Select", "Tools:Out Of Office Assistant...")
		Call Fn_ReadyStatusSync(5)		   	    
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Tools --> Out Of Office Assistant... ") 		
			Fn_MyWorkList_OutOfOfficeAssit = False
			Set objDialog = Nothing
			Exit Function					
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Tools --> Out Of Office Assistant... ")
			Call Fn_ReadyStatusSync(4)
		End If
	End If	

	If objDialog.Exist(5) Then

		Select Case sAction

		Case "Set"

			If Trim(sOrgGrp) <> "" Then
				objDialog.JavaButton("OrgGrpDrpDwnBtn").Click
				objDialog.JavaEdit("OrgGroup").Set sOrgGrp
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sOrgGrp)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sOrgGrp)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sOrgRole) <> "" Then
				objDialog.JavaButton("OrgRoleDrpDwnBtn").Click
				objDialog.JavaEdit("OrgRole").Set sOrgRole
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Role ["+CStr(sOrgRole)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Role ["+CStr(sOrgRole)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sOrgUser) <> "" Then
				objDialog.JavaButton("OrgUsrDrpDwnBtn").Click
				objDialog.JavaEdit("OrgUser").Set sOrgUser
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sOrgUser)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sOrgUser)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sFrmDt) <> "" Then
				adate = split(sFrmDt," ")
				bReturn = Fn_Edit_Box("Fn_PSE_RevisionRuleSetDate", objDialog, "FromDate", adate(0) )
				If bReturn = False  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select From Date ["+CStr(adate(0))+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				End If
				Wait(2)
				bReturn = Fn_Edit_Box("Fn_PSE_RevisionRuleSetDate", objDialog, "FromTime", adate(1) )
				If bReturn = False  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select From Time ["+CStr(adate(1))+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				End If
				Wait(2)
			
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected From Date ["+CStr(sFrmDt)+"]")
				Call Fn_ReadyStatusSync(2)
			End If

			If Trim(sToDt) <> "" Then
				adate = split(sToDt," ")
				bReturn = Fn_Edit_Box("Fn_PSE_RevisionRuleSetDate", objDialog, "ToDate", adate(0) )
				If bReturn = False  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select From Date ["+CStr(adate(0))+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				End If
				
				bReturn = Fn_Edit_Box("Fn_PSE_RevisionRuleSetDate", objDialog, "ToTime", adate(1) )
				If bReturn = False  Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select From Time ["+CStr(adate(1))+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				End If
				Wait(2)
				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected To Date ["+CStr(sToDt)+"]")
				Call Fn_ReadyStatusSync(2)
		
			End If

			If Trim(sNewGrp) <> "" Then
				objDialog.JavaButton("NewGrpDrpDwnBtn").Click
				objDialog.JavaEdit("NewGroup").Set sNewGrp
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sNewGrp)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sNewGrp)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sNewRole) <> "" Then
				objDialog.JavaButton("NewRoleDrpDwnBtn").Click
				objDialog.JavaEdit("NewRole").Set sNewRole
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Role ["+CStr(sNewRole)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Role ["+CStr(sNewRole)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sNewUser) <> "" Then
				objDialog.JavaButton("NewUsrDrpDwnBtn").Click
				objDialog.JavaEdit("NewUser").Set sNewUser
				Wait(2)
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select User ["+CStr(sNewUser)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected User ["+CStr(sNewUser)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If
			
			wait(3)
			objDialog.JavaButton("Apply").Click micLeftBtn
			wait(3)

			objDialog.JavaButton("OK").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK Button ") 		
				Fn_MyWorkList_OutOfOfficeAssit = False
				objDialog.JavaButton("Cancel").Click micLeftBtn
				Set objDialog = Nothing
				Exit Function		
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on OK Button")
				Call Fn_ReadyStatusSync(2)
			End If
		
		Case "Verify"

			If Trim(sOrgGrp) <> "" Then
				If objDialog.JavaEdit("OrgGroup").GetROProperty("text") <> Trim(sOrgGrp) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Group ["+CStr(sOrgGrp)+"] ") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Group ["+CStr(sOrgGrp)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sOrgRole) <> "" Then
				If objDialog.JavaEdit("OrgRole").GetROProperty("text") <> Trim(sOrgRole) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Role ["+CStr(sOrgRole)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Role ["+CStr(sOrgRole)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sOrgUser) <> "" Then
				If objDialog.JavaEdit("OrgUser").GetROProperty("text") <> Trim(sOrgUser) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify User ["+CStr(sOrgUser)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified User ["+CStr(sOrgUser)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sFrmDt) <> "" Then
				adate = split(sFrmDt," ")
				If objDialog.JavaEdit("FromDate").GetROProperty("text") = Trim(adate(0)) and objDialog.JavaEdit("FromTime").GetROProperty("text") = Trim(adate(1)) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified From Date ["+CStr(sFrmDt)+"]")
					Call Fn_ReadyStatusSync(2)
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify From Date ["+CStr(sFrmDt)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				End If
			End If

			If Trim(sToDt) <> "" Then
				adate = split(sToDt," ")
				If objDialog.JavaEdit("ToDate").GetROProperty("text") = Trim(adate(0)) and objDialog.JavaEdit("ToTime").GetROProperty("text") = Trim(adate(1)) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified To Date ["+CStr(sToDt)+"]")
					Call Fn_ReadyStatusSync(2)
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify To Date ["+CStr(sToDt)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				End If
			End If

			If Trim(sNewGrp) <> "" Then
				If objDialog.JavaEdit("NewGroup").GetROProperty("text") <> Trim(sNewGrp) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Group ["+CStr(sNewGrp)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Group ["+CStr(sNewGrp)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sNewRole) <> "" Then
				If objDialog.JavaEdit("NewRole").GetROProperty("text") <> Trim(sNewRole) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Role ["+CStr(sNewRole)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function	
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Role ["+CStr(sNewRole)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If

			If Trim(sNewUser) <> "" Then
				If objDialog.JavaEdit("NewUser").GetROProperty("text") <> Trim(sNewUser) Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify User ["+CStr(sNewUser)+"]") 		
					Fn_MyWorkList_OutOfOfficeAssit = False
					objDialog.JavaButton("Cancel").Click micLeftBtn
					Set objDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified User ["+CStr(sNewUser)+"]")
					Call Fn_ReadyStatusSync(2)
				End If
			End If
			
			wait(2)

			objDialog.JavaButton("Cancel").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Cancel Button ") 		
				Fn_MyWorkList_OutOfOfficeAssit = False
				Set objDialog = Nothing
				Exit Function		
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Cancel Button")
				Call Fn_ReadyStatusSync(2)
			End If
	
		End Select

			Fn_MyWorkList_OutOfOfficeAssit = True

	Else

		Fn_MyWorkList_OutOfOfficeAssit = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: MyWorkList Out Of Office Assit... function failed")
	End If

	Set objDialog = Nothing
End Function

'*********************************************************  Function is Subscribe to any Resource Pool *********************************************************************

'Function Name		:					Fn_MyWorkList_TaskDemote

'Description			 :		 		  	Demote the Do task			

'Parameters			   :	 			  1. sTaskName: Task to be selected from worklist tree
'														2. sComment: promote comment
'
'Return Value		   : 			 	True/False

'Pre-requisite			:		 	 Myworklist tab should selected.

'Examples				:			Fn_MyWorkList_TaskDemote("My Worklist:AutoTestDBA (autotestdba) Inbox:Tasks to Perform:000565/A;1-ccc (perform-signoffs)","Promote Do task")

'History:
'										Developer Name									Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										    Mahendra Bhandarkar			        24-Sep-2010	        1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Modified By							Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Vidya Kulkarni			        14-July-2011       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_TaskDemote(sTaskName, sComment)
		GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TaskDemote"
		Dim sMenu, bReturn, ObjPromoteWin
		On Error Resume Next
	
		
		If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").Exist Then
				set ObjPromoteWin = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Demote Action Comments")
				sMenu = "Actions:Demote"
		Else
				set ObjPromoteWin = Fn_SISW_WorkflowViewer_GetObject("Demote Action Comments")
				sMenu = "Actions:Demote"
		End If
		'Select MyWorklist Tree Node
		If Trim(sTaskName) <> "" Then
			 bReturn = Fn_MyWorkList_TreeNodeOperations("Select", sTaskName,"")
			If bReturn = False Then
						Fn_MyWorkList_TaskDemote = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select  [" + sTaskName+"] from Worklist" )	
						Exit Function
			Else
						 wait(3)
						 Call Fn_ReadyStatusSync(5) 
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected  [" + sTaskName+"] from WorkList" )	
			End If

		End If

		'Demote Window Exists or Not	
		If ObjPromoteWin.Exist(5) = False Then				
				bReturn = Fn_MenuOperation("Select", sMenu)
				Call Fn_ReadyStatusSync(3)
				If bReturn = False Then
					Fn_MyWorkList_TaskDemote = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke  Menu [" + sMenu+"]." )	
					Exit Function
				End If
		End If

		If ObjPromoteWin.Exists(10) Then
			If Trim(sComment) <> "" Then
					Err.Clear
					ObjPromoteWin.JavaEdit("Comments").Set Trim(sComment)
					If Err.Number < 0 Then
							Fn_MyWorkList_TaskDemote = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Comment [" + sComment+"].")	
							ObjPromoteWin.JavaButton("Cancel").Click micLeftBtn
							Set ObjPromoteWin = Nothing
							Exit Function
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Entered Comment   [" + sComment+"].")
							Call Fn_ReadyStatusSync(2)
					End If
			End If
	
			ObjPromoteWin.JavaButton("OK").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK Button ") 		
				Fn_MyWorkList_TaskDemote = False
				ObjPromoteWin.JavaButton("Cancel").Click micLeftBtn
				Set ObjPromoteWin = Nothing
				Exit Function		
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on OK Button")
				Call Fn_ReadyStatusSync(2)
			End If

			Fn_MyWorkList_TaskDemote = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Demoted Task [" + sTaskName+"]")	
			Set ObjPromoteWin = Nothing
	
		Else
			Fn_MyWorkList_TaskDemote = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find  [Demote Action Comments] Dialog" )	
			Exit Function
		End If

End Function


'**********************		Function to Assign Workflow Sub Process To Object		***********************************************************************
'Function Name		:				Fn_MyWorkList_WorkflowSubProcessAssign

'Description			 :		 		 Assign Workflow Sub-Process To Object

'Parameters			   :	 			1. sProcName: Name of the Sub-Process 
'													2. sDescription: Description of the Sub-Process
'													3. sProcessTemplate: Workflow Template to be applied.

'Return Value		   : 				True / False

'Pre-requisite			:		 		Strucure Node is already Selected.

'Examples				:				 Fn_MyWorkList_WorkflowSubProcessAssign("New","TestDesc","TCM Release Process")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Omkar Kulkarni										    27/09/2010			              1.0								Created										Mahendra B.
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_WorkflowSubProcessAssign(sProcessTemplateFilter, sProcName, sDescription, sProcessTemplate)
	
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_WorkflowSubProcessAssign"
	Dim sTemplateType, intNoOfObjects, iCounter, bFlag

	If sProcessTemplateFilter = "" Then
		sProcessTemplateFilter = "All"
	End If

	bFlag = False
    'Select menu [File -> New - > Workflow Sub-Process ...]
	If Not JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").Exist(5) Then
			Call Fn_MenuOperation("Select","File:New:Workflow Sub-Process ...")
			Call Fn_ReadyStatusSync(5)		   	    
	End If
	Wait(5)
    If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").Exist(iTimeOut) Then

			'Check Whether to Select All templates or Assigned templates
			If  Trim(sProcessTemplateFilter) <> "" Then
					JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaRadioButton("ProcessTemplateFilter").SetTOProperty "attached text", Trim(sProcessTemplateFilter)
					JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaRadioButton("ProcessTemplateFilter").Set "ON" 
					If Err.Number < 0 Then
								Fn_MyWorkList_WorkflowSubProcessAssign = False
								JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("Cancel").Click micLeftBtn
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set the Filter " + sProcessTemplateFilter)	
								sProcessTemplateFilter = ""
								Exit Function 
						Else								
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set the Filter " + sProcessTemplateFilter)	
								sProcessTemplateFilter = ""
						End If
			End If

			'Set  Process Name
			If Trim(sProcName) <> "" Then
						JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaEdit("Sub-Process Name").Set sProcName
						If Err.Number < 0 Then
									Fn_MyWorkList_WorkflowSubProcessAssign = False
									JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("Cancel").Click micLeftBtn
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set the Process Name  " + sProcName)
									Exit Function 
							Else								
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set the Process Name " + sProcName)
							End If
			End If

			'Set Process description
			If Trim(sDescription) <> "" Then
							JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaEdit("Description").Set sDescription
						If Err.Number < 0 Then
									Fn_MyWorkList_WorkflowSubProcessAssign = False
									JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("Cancel").Click micLeftBtn
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set the Process Description  " + sDescription)
									Exit Function 
							Else								
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set the Process Description " + sDescription)
							End If
			End If

			'Set Process Template
			JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("ProcTempBtn").Click
			Set sTemplateType=Description.Create()
			sTemplateType("Class Name").value = "JavaStaticText"
			Set  intNoOfObjects = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").ChildObjects(sTemplateType)
			  For iCounter = 0 to intNoOfObjects.count-1
				   If  intNoOfObjects(iCounter).getROProperty("label") = sProcessTemplate Then
							intNoOfObjects(iCounter).Click 1,1
                            bFlag = True
							Exit for
				   End If
			  Next

		  	 If Trim(sProcessTemplate) <> "" and bFlag = False Then
					Fn_MyWorkList_WorkflowSubProcessAssign = False
					JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("Cancel").Click micLeftBtn
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Process Template ["+CStr(sProcessTemplate)+"] does not exist on New Sub Process Dialog " )
					Exit Function 
			 End If

			'Click on "OK" button
            If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("OK").GETROProperty("disabled") < 1 Then
						JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("OK").Click micLeftBtn
						wait(5)
						If Err.Number < 0 Then
								Fn_MyWorkList_WorkflowSubProcessAssign = False
								JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("Cancel").Click micLeftBtn
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on OK button" )	
								Exit Function 
						Else					
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on OK button")	
						End If
			Else
						Fn_MyWorkList_WorkflowSubProcessAssign = False
						JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").JavaButton("Cancel").Click micLeftBtn
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on OK button because OK Button is Disabled" ) 
						Exit Function 
            End If
			Fn_MyWorkList_WorkflowSubProcessAssign = TRUE
	Else
			Fn_MyWorkList_WorkflowSubProcessAssign = FALSE
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Workflow Sub Process Dialog does not exist." ) 
	End If

	Set sTemplateType = Nothing
	Set intNoOfObjects = Nothing

End Function

'*********************************************************  Function is Subscribe to any Resource Pool *********************************************************************

'Function Name		 :					Fn_MyWorkList_TaskSuspend

'Description			:		 		  Suspend the Do task	

'Parameters			   :	 			 1. sTaskName: Task to be selected from worklist tree
'						 						2. sComment: promote comment
'
'Return Value		   : 			 	True/False

'Pre-requisite			:		 	 	Myworklist tab should selected.

'Examples				:				Fn_MyWorkList_TaskSuspend("My Worklist:Engineering/Designer Inbox:Tasks to Track:a (perform-signoffs)","Suspended Do task")

'History:s
'										Developer Name									Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										    Mahendra Bhandarkar			        29-Sep-2010	        1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_TaskSuspend(sTaskName, sComment)
		GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TaskSuspend"
		Dim sMenu, bReturn, ObjSuspendWin
		On Error Resume Next
	
		Set ObjSuspendWin = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Suspend Action Comments")

		'Select MyWorklist Tree Node
		If Trim(sTaskName) <> "" Then
			 bReturn = Fn_MyWorkList_TreeNodeOperations("Select", sTaskName,"")
			If bReturn = False Then
						Fn_MyWorkList_TaskSuspend = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select  [" + sTaskName+"] from Worklist" )	
						Exit Function
			Else
						 wait(3)
						 Call Fn_ReadyStatusSync(5) 
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected  [" + sTaskName+"] from WorkList" )	
			End If

		End If

		'Demote Window Exists or Not	
		If ObjSuspendWin.Exist(5) = False Then
				sMenu = "Actions:Suspend"
				bReturn = Fn_MenuOperation("Select", sMenu)
				Call Fn_ReadyStatusSync(5)		   	    
				If bReturn = False Then
					Fn_MyWorkList_TaskSuspend = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke  Menu [" + sMenu+"]." )	
					Exit Function
				End If
		End If

		If ObjSuspendWin.Exists(10) Then
			If Trim(sComment) <> "" Then
					Err.Clear
					ObjSuspendWin.JavaEdit("Comments").Set Trim(sComment)
					If Err.Number < 0 Then
							Fn_MyWorkList_TaskSuspend = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Comment [" + sComment+"].")	
							ObjSuspendWin.JavaButton("Cancel").Click micLeftBtn
							Set ObjSuspendWin = Nothing
							Exit Function
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Entered Comment [" + sComment+"].")
							Call Fn_ReadyStatusSync(2)
					End If
			End If
	
			ObjSuspendWin.JavaButton("OK").Click micLeftBtn
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK Button ") 		
				Fn_MyWorkList_TaskSuspend = False
				ObjSuspendWin.JavaButton("Cancel").Click micLeftBtn
				Set ObjSuspendWin = Nothing
				Exit Function		
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on OK Button")
				Call Fn_ReadyStatusSync(2)
			End If

			Fn_MyWorkList_TaskSuspend = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Suspended Task [" + sTaskName+"]")	
			Set ObjSuspendWin = Nothing
	
		Else
			Fn_MyWorkList_TaskSuspend = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find  [Suspend Action Comments] Dialog" )	
			Exit Function
		End If

End Function

'*********************************************************  Function perform the worklist process view attributes operation *********************************************************************
'Function Name  :   Fn_MyWorkList_ProcessView_Attributes
'
'Description    :        MyWorkList Process View Attributes Operation
' 
'Parameters      :     sAction: Add/Remove/Modify/Verify
'           				 		dicProcessViewAttributes: Refer DictionaryDeclaration.vbs for the defination & keys included
' 
'Return Value     :   True/False
'
'Examples    :      
'								dicProcessViewAttributes.RemoveAll
'								dicProcessViewAttributes.Add("WorkListTreeNode") = "My Worklist:AutoTestDBA (autotestdba) Inbox:Tasks to Perform:000110/A;1-test (select-signoff-team)"
'								dicProcessViewAttributes.Add("ProcessTree") = "AutoRevFailPath:New Review Task 1:select-signoff-team"
'								dicProcessViewAttributes.Add("Attributes") = "True" 
'								dicProcessViewAttributes.Add("State") = "Started"
'								dicProcessViewAttributes.Add("ResParty") = "AutoTestDBA (autotestdba)"
'								dicProcessViewAttributes.Add("NameACL") = ""
'								dicProcessViewAttributes.Add("SignOffsQuorum") = ""
'								dicProcessViewAttributes.Add("DueDate") = ""
'								dicProcessViewAttributes.Add("Duration") = ""
'
'           					Call Fn_MyWorkList_ProcessView_Attributes(sAction, dicProcessViewAttributes)
' 
'History:
'          Developer Name   Date    Rev. No.   Changes Done   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           Mahendra    		04-Oct-2010   1.0             Prasanna  
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_ProcessView_Attributes(sAction, dicProcessViewAttributes)
	
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_ProcessView_Attributes"
	 On Error Resume Next
	 Dim dicCount , dicKeys , dicItems
	 Dim iCounter, bReturn
	 Dim arrNodeList, iNodeCounter, arrNode, sExpnadNode
	 Dim iCounter1, sActionSelect
	 Dim objSelectType, intNoOfObjects, objDialog, iCounter2, sListHierarchy, arrQuorum
	
	 dicCount  = dicProcessViewAttributes.Count
	 dicItems = dicProcessViewAttributes.Items
	 dicKeys = dicProcessViewAttributes.Keys

   Select Case sAction         

   Case "Verify"

    For iCounter = 0 to dicCount - 1
          If  dicItems(iCounter) <> "" Then
			   Select Case dicKeys(iCounter)

			      Case "WorkListTreeNode"
					'Select the WorkList Tree Node
					If Trim(dicItems(iCounter)) <> "" Then
								 bReturn =  Fn_MyWorkList_TreeNodeOperations("Select",dicItems(iCounter),"")
								 If bReturn = false Then
									   Fn_MyWorkList_ProcessView_Attributes = False         
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + dicItems(iCounter) +  " From WorkList Tree")
									   Exit Function
								 End If
'								 Call Fn_ReadyStatusSync(5)
								 Wait(5)				 
			
								 'Set the Viewer Tab
								 bReturn =  Fn_MyTc_TabSet("Viewer") 
								 If bReturn = false Then
									   Fn_MyWorkList_ProcessView_Attributes = False         
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Viewer Tab")
									   Exit Function
								  Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Viewer Tab")
'										Call Fn_ReadyStatusSync(3)
										Wait(3)
								 End If								 
			
								 'Set Default View to Process View 
								 JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ViewOptions").SetTOProperty "Attached Text","Process View"
								 JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaRadioButton("ViewOptions").Set "ON"
								 If Err.Number < 0 Then
									   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Process View from Viewer tab.")
									   Fn_MyWorkList_ProcessView_Attributes = False         
									   Exit Function
								 Else                      
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Process View from Viewer tab.")
										Call Fn_ReadyStatusSync(5)
										Wait(2)
								 End If
					End If

					 
				  Case "ProcessTree"
					arrNode = Split(dicItems(iCounter), ":", -1, 1)
					For iNodeCounter = 0 To UBound(arrNode)
					  If iNodeCounter = 0 Then
						   sExpnadNode = arrNode(iNodeCounter)
					  Else
						   sExpnadNode = sExpnadNode+":"+arrNode(iNodeCounter)
					  End If
					If iNodeCounter <>  UBound(arrNode) Then
						If iNodeCounter <> 0 Then
								JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProcessTree").Expand sExpnadNode
								wait(2)
								If Err.Number < 0 Then
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Node ["+sExpnadNode+"] in Process Tree.")
									  Fn_MyWorkList_ProcessView_Attributes = False         
									  Exit Function
								Else
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expand Node ["+sExpnadNode+"] in Process Tree.")
									  Call Fn_ReadyStatusSync(1)
									   Wait(3)
								End If
							Else
								JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProcessTree").Select sExpnadNode
								wait(2)
								If Err.Number < 0 Then
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node ["+sExpnadNode+"] in Process Tree.")
									  Fn_MyWorkList_ProcessView_Attributes = False         
									  Exit Function
								Else                      
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Node ["+sExpnadNode+"] in Process Tree.")
									  Call Fn_ReadyStatusSync(1)
									   Wait(3)
								End If
							End If
					End If
					Next

					JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProcessTree").Select sExpnadNode
					wait(2)
					If Err.Number < 0 Then
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node ["+sExpnadNode+"] in Process Tree.")
						  Fn_MyWorkList_ProcessView_Attributes = False         
						  Exit Function
					Else                      
						  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Node ["+sExpnadNode+"] in Process Tree.")
						  Call Fn_ReadyStatusSync(3)
						   Wait(3)
					End If
		 			
				   Case "Attributes"
					' Select the Attributes Dialog
					If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").Exist(2) = False  Then
						JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaCheckBox("AttributesBtn").Set "ON"
							 If Err.Number < 0 Then
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Checkbox.")
								   Fn_MyWorkList_ProcessView_Attributes = False
								   Exit Function
							 Else                      
								   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Checkbox.")
								   Call Fn_ReadyStatusSync(2)
									Wait(3)
							 End If
					End If
					
					If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").Exist(2) = True Then
					     JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").Activate
						  If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Attribute Dialog Does not exist.")
								Fn_MyWorkList_ProcessView_Attributes = False
								Exit Function
						  Else
		  						 wait(3)
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Attributes Dialog.")
						  End If
					End If
					'Added by Nilesh on 15-Jun-2012 
					If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").Exist(2) = False Then
						JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaStaticText("Attributes").DblClick 1, 1, "LEFT"
					End If
					'End
					If Err.Number < 0 Then
							  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Text.")
							  Fn_MyWorkList_ProcessView_Attributes = False         
							  Exit Function
					Else                      
							  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Text.")
							  Wait(2)					
					End If
					 

				 Case "State"
				   Set objComp = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").JavaObject("State").Object.getComponent(0)
				   If objComp.getText() = dicItems(iCounter) Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : State Selection ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
				   Else
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the State Selected Text ["+dicItems(iCounter)+"] in Attributes Dialog. ")
							 Fn_MyWorkList_ProcessView_Attributes = False       
							 Exit Function
				   End If
				   Set objComp = Nothing
		
				 Case "ResParty"
		
				   Set objComp = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").JavaObject("ResponsibleParty").Object.getComponent(0)
				   If objComp.getText() = dicItems(iCounter) Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Responsible Party Selection ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
				   Else
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the Responsible Party Selected Text ["+dicItems(iCounter)+"] in Attributes Dialog. ")
							 Fn_MyWorkList_ProcessView_Attributes = False       
							 Exit Function
				   End If
				   Set objComp = Nothing
		
				 Case "NameACL"

				 Case "SignOffsQuorum"

				 Case "DueDate"
					If InStr(1, JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").JavaCheckBox("DueDate").GetROProperty("attached text"), Trim(dicItems(iCounter)), 1) > 0  Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified the Due Date ["+CStr(dicItems(iCounter))+"]")
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify The Due Date ["+CStr(dicItems(iCounter))+"]") 		
							Fn_MyWorkList_ProcessView_Attributes = False
							Set objDialog = Nothing
							Exit Function
					End If

				 Case "Duration"

				   Set objComp = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").JavaEdit("Duration")
				   If objComp.GetROProperty("text") = dicItems(iCounter) Then
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Duration Text ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
				   Else
							 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the Duration Text ["+dicItems(iCounter)+"] in Attributes Dialog.")
							 Fn_MyWorkList_ProcessView_Attributes = False       
							 Exit Function
				   End If
				   Set objComp = Nothing
		
				 Case "Recipients"
		 
       End Select

    End if

  Next

 Case "Modify"

 Case "Remove"

  Case "Add"

 End Select

' Close the Attributes Dialog
If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").Exist(2) = True  Then
	JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Attributes").Close
		 If Err.Number < 0 Then
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Close Attributes Dialog.")
			   Fn_MyWorkList_ProcessView_Attributes = False
			   Exit Function
		 Else                      
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Closed Attributes Dialog.")
			   Call Fn_ReadyStatusSync(2)
				Wait(2)
		 End If
End If

Fn_MyWorkList_ProcessView_Attributes = True  
Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified Attributes Dialog.")
End Function

'*********************************************************		Function to Creates Log File and folder.		***********************************************************************
'Function Name		:				Fn_WorkflowSubProcess_Operations(sAction, dicNewSubProcess, sBtnName)

'Description			 :		 		 Perform Operations on Workflow Sub-Process Dialog

'Parameters			   :	 			1. sAction : Actions Verify/ Add, etc.
'													 2.	dicNewSubProcess -  Sub Process Property Dictionary
'													 3.	sBtnName -  Click on the Button [Ok/ Cancel]

'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Examples				:
' Add Case
'												dicNewSubProcess.RemoveAll
'												dicNewSubProcess.Add "Description", "Description here"
'												dicNewSubProcess.Add "ProcessTemplate", "AutoDoDo"
'												Call Fn_WorkflowSubProcess_Operations("Add", dicNewSubProcess, "")
' ' Verify Case
'												dicNewSubProcess.RemoveAll
'												dicNewSubProcess.Add "InheritTargets", "True"
'												dicNewSubProcess.Add "ProcessName", "000539/A;1-TestItem_74052"
'												dicNewSubProcess.Add "Attachments", "Task Attachments:Targets:000539/A;1-TestItem_74052"
'												Call Fn_WorkflowSubProcess_Operations("Verify", dicNewSubProcess, "OK")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Mahendra Bhandarkar								   	  08/10/2010			     1.0									Created
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WorkflowSubProcess_Operations(sAction, dicNewSubProcess, sBtnName)
	
	GBL_FAILED_FUNCTION_NAME="Fn_WorkflowSubProcess_Operations"
	On Error Resume Next
	Dim dicCount , dicKeys , dicItems
	Dim iCounter, bReturn
	Dim objDialog, sCmpVal
	Dim sTemplateType, intNoOfObjects, bFlag, iCounter1	
	Dim arrNode, iNodeCounter, sExpnadNode, arrHeadNode

	dicCount  = dicNewSubProcess.Count
	dicItems = dicNewSubProcess.Items
	dicKeys = dicNewSubProcess.Keys

	'Select menu [File -> New - > Workflow Sub-Process ...]
	If   JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").Exist(5)=False AND JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("New Sub-Process").Exist(5) =False AND JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("New Sub-Process").Exist(5) = False Then
			bReturn = Fn_MenuOperation("Select","File:New:Workflow Sub-Process ...")
			Call Fn_ReadyStatusSync(5)		   	    
			 If bReturn = false Then
				   Fn_WorkflowSubProcess_Operations = False         
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Dialog Workflow Sub-Process ... From WorkList Tree")
				   Exit Function
			 End If
	End If

	'Added by Nilesh for Hierachy change on TC10.0 0606 Build
	If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process").Exist(5) Then
		Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("New Sub-Process")
	ElseIf JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("New Sub-Process").Exist(5) Then
		Set objDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("New Sub-Process")
	ElseIf  JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("New Sub-Process").Exist(5) Then
		Set objDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("New Sub-Process")
	End If
	

	Wait(5)

    If objDialog.Exist(iTimeOut) Then
		   Select Case sAction
		   Case "Verify"
			For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> "" Then
						   Select Case dicKeys(iCounter)
			
							  Case "InheritTargets"
								'Select the WorkList Tree Node
								If Trim(dicItems(iCounter)) <> "" Then
											bReturn =  Fn_MyWorkList_TreeNodeOperations("Select",dicItems(iCounter),"")
											If CBool(dicItems(iCounter)) = True Then
												sCmpVal = "1"
											Else
												sCmpVal = "0"
											End If
											If objDialog.JavaCheckBox("InheritTargets").GetROProperty("value") <> sCmpVal Then
												   Fn_WorkflowSubProcess_Operations = False         
												   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To verify that Inherit Targets marked [" + CStr(dicItems(iCounter)) +  "] in WorkList New Sub-Process Dialog.")
													objDialog.JavaButton("Cancel").Click micLeftBtn
												   Exit Function
											Else
												   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Inherit Targets verified to mark Checked to [" + CStr(dicItems(iCounter)) +  "] in WorkList New Sub-Process Dialog.")
												  Call Fn_ReadyStatusSync(1)
												   Wait(1)
											End If
								End If
		
							  Case "ProcessName"
								'Select the WorkList Tree Node
								If Trim(dicItems(iCounter)) <> "" Then
											If objDialog.JavaEdit("Sub-Process Name").GetROProperty("value") <> dicItems(iCounter) Then
												   Fn_WorkflowSubProcess_Operations = False         
												   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To verify that Sub-Process Name value ["+CStr(objDialog.JavaEdit("Sub-Process Name").GetROProperty("value"))+"] not matched with [" + dicItems(iCounter) +  "] in WorkList New Sub-Process Dialog.")
													objDialog.JavaButton("Cancel").Click micLeftBtn
												   Exit Function
											Else
												   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Sub-Process Name value matched verified to [" + dicItems(iCounter) +  "] in WorkList New Sub-Process Dialog.")
												  Call Fn_ReadyStatusSync(1)
												   Wait(1)
											End If
								End If

								Case "Attachments"
								'Select the WorkList Tree Node
								If Trim(dicItems(iCounter)) <> "" Then
									arrNode = Split(dicItems(iCounter), ":", -1, 1)
									For iNodeCounter = 0 To UBound(arrNode)
									  If iNodeCounter = 0 Then
										   sExpnadNode = arrNode(iNodeCounter)
									  Else
										   sExpnadNode = sExpnadNode+":"+arrNode(iNodeCounter)
									  End If
									If iNodeCounter <>  UBound(arrNode) Then
										objDialog.JavaTree("AttachmentsTree").Expand sExpnadNode
										If Err.Number < 0 Then
											  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Node ["+sExpnadNode+"] in Attachments Tree.")
											  Fn_WorkflowSubProcess_Operations = False         
											  Exit Function
										Else
											  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expand Node ["+sExpnadNode+"] in Attachments Tree.")
											  Call Fn_ReadyStatusSync(1)
											   Wait(1)
										End If
									End If
									Next
				
									objDialog.JavaTree("AttachmentsTree").Select sExpnadNode
									If Err.Number < 0 Then
										  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node ["+sExpnadNode+"] in Process Tree.")
										  Fn_WorkflowSubProcess_Operations = False         
										  Exit Function
									Else                      
										  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Node ["+sExpnadNode+"] in Process Tree.")
										  Call Fn_ReadyStatusSync(1)
										   Wait(1)
									End If
								End If					
											 
				   End Select
		
			End if
		
		 Next
		
		 Case "Modify"
		
		 Case "Remove"
		
		  Case "Add"

			For iCounter = 0 to dicCount - 1
					If  dicItems(iCounter) <> "" Then
						   Select Case dicKeys(iCounter)
							  Case "ProcessName"
									'Set SubProcess Name
									If Trim(dicItems(iCounter)) <> "" Then
												objDialog.JavaEdit("Sub-Process Name").Set Trim(dicItems(iCounter))
												If Err.Number < 0 Then
													   Fn_WorkflowSubProcess_Operations = False         
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To verify that Sub-Process Name value ["+CStr(objDialog.JavaEdit("Sub-Process Name").GetROProperty("value"))+"] not matched with [" + dicItems(iCounter) +  "] in WorkList New Sub-Process Dialog.")
														objDialog.JavaButton("Cancel").Click micLeftBtn
													   Exit Function
												Else
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Sub-Process Name value matched verified to [" + dicItems(iCounter) +  "] in WorkList New Sub-Process Dialog.")
													  Call Fn_ReadyStatusSync(1)
													   Wait(1)
												End If
									End If

								Case "Description"
									'Set the Description 
									If Trim(dicItems(iCounter)) <> "" Then
												objDialog.JavaEdit("Description").Set Trim(dicItems(iCounter))
												If Err.Number < 0 Then
													   Fn_WorkflowSubProcess_Operations = False         
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Set Description to [" + CStr(dicItems(iCounter)) +  "] in WorkList New Sub-Process Dialog.")
														objDialog.JavaButton("Cancel").Click micLeftBtn
													   Exit Function
												Else
													   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Set Description to [" + CStr(dicItems(iCounter)) +  "] in WorkList New Sub-Process Dialog.")
													  Call Fn_ReadyStatusSync(1)
													   Wait(1)
												End If
									End If

								Case "ProcessTemplate"
									'Set Process Template
									objDialog.JavaButton("ProcTempBtn").Click
									Set sTemplateType=Description.Create()
									sTemplateType("Class Name").value = "JavaStaticText"
									sTemplateType("label").value = Trim(dicItems(iCounter))
									wait(1)
									Set  intNoOfObjects = objDialog.ChildObjects(sTemplateType)
									  If intNoOfObjects.count > 0 Then
											intNoOfObjects(0).Click 1,1
											bFlag = True
									  End If
									  Set sTemplateType=Nothing
									  Set  intNoOfObjects = Nothing
									 If Trim(Trim(dicItems(iCounter))) <> "" and bFlag = False Then
											Fn_WorkflowSubProcess_Operations = False
											objDialog.JavaButton("Cancel").Click micLeftBtn
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Process Template ["+CStr(dicItems(iCounter))+"] does not exist on New Sub Process Dialog " )
											Exit Function 
									Else
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Select Process Template ["+CStr(dicItems(iCounter))+"]  in WorkList New Sub-Process Dialog.")
										  Call Fn_ReadyStatusSync(1)
										   Wait(1)
									End If
							  Case "InheritTargets"
								' Checked/ Unchecked the Inherit Targets CheckBox
								If Trim(dicItems(iCounter)) <> "" Then
												If CBool(dicItems(iCounter)) = False Then
													objDialog.JavaCheckBox("InheritTargets").Set "OFF"
													sCmpVal = "Checked"
												ElseIf CBool(dicItems(iCounter)) = True Then
													objDialog.JavaCheckBox("InheritTargets").Set "ON"
													sCmpVal = "Unchecked"
												End If
												If Err.Number < 0 Then
												   Fn_WorkflowSubProcess_Operations = False         
												   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To ["+sCmpVal+"] Inherit Targets CheckBox in WorkList New Sub-Process Dialog.")
												   objDialog.JavaButton("Cancel").Click micLeftBtn
												   Exit Function
												Else
												   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully ["+sCmpVal+"] Inherit Targets CheckBox in WorkList New Sub-Process Dialog.")
												   Call Fn_ReadyStatusSync(1)
												   Wait(1)
												End If
								End If

								Case "Attachments"
								'Select the WorkList Tree Node
								If Trim(dicItems(iCounter)) <> "" Then
									arrNode = Split(dicItems(iCounter), ":", -1, 1)
									For iNodeCounter = 0 To UBound(arrNode)
									  If iNodeCounter = 0 Then
										   sExpnadNode = arrNode(iNodeCounter)
									  Else
										   sExpnadNode = sExpnadNode+":"+arrNode(iNodeCounter)
									  End If
									If iNodeCounter <>  UBound(arrNode) Then
										objDialog.JavaTree("AttachmentsTree").Expand sExpnadNode
										If Err.Number < 0 Then
											  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Node ["+sExpnadNode+"] in Attachments Tree.")
											  Fn_WorkflowSubProcess_Operations = False
											  Exit Function
										Else
											  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expand Node ["+sExpnadNode+"] in Attachments Tree.")
											  Call Fn_ReadyStatusSync(1)
											   Wait(1)
										End If
									End If
									Next
				
									objDialog.JavaTree("AttachmentsTree").Select sExpnadNode
									If Err.Number < 0 Then
										  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node ["+sExpnadNode+"] in Process Tree.")
										  Fn_WorkflowSubProcess_Operations = False         
										  Exit Function
									Else                      
										  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Node ["+sExpnadNode+"] in Process Tree.")
										  Call Fn_ReadyStatusSync(1)
										   Wait(1)
									End If
								End If					

								Case "AttachmentsMultiSelect"
								'Select the WorkList Tree Node
								arrHeadNode = Split(dicItems(iCounter), ",", -1, 1)
								For iCounter1 = 0 To UBound(arrHeadNode)
										If Trim(arrHeadNode(iCounter1)) <> "" Then
											arrNode = Split(arrHeadNode(iCounter1), ":", -1, 1)
											For iNodeCounter = 0 To UBound(arrNode)
											  If iNodeCounter = 0 Then
												   sExpnadNode = arrNode(iNodeCounter)
											  Else
												   sExpnadNode = sExpnadNode+":"+arrNode(iNodeCounter)
											  End If
											If iNodeCounter <>  UBound(arrNode) Then
												objDialog.JavaTree("AttachmentsTree").Expand sExpnadNode
												If Err.Number < 0 Then
													  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Node ["+sExpnadNode+"] in Attachments Tree.")
													  Fn_WorkflowSubProcess_Operations = False         
													  Exit Function
												Else
													  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expand Node ["+sExpnadNode+"] in Attachments Tree.")
													  Call Fn_ReadyStatusSync(1)
													   Wait(1)
												End If
											End If
											Next
											If iCounter1= 0 Then
													objDialog.JavaTree("AttachmentsTree").Select sExpnadNode
											Else
													objDialog.JavaTree("AttachmentsTree").ExtendSelect sExpnadNode
											End If
											If Err.Number < 0 Then
												  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node ["+sExpnadNode+"] in Process Tree.")
												  Fn_WorkflowSubProcess_Operations = False         
												  Exit Function
											Else                      
												  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Node ["+sExpnadNode+"] in Process Tree.")
												  Call Fn_ReadyStatusSync(1)
												   Wait(1)
											End If
										End If
								Next

								Case "AttachmentsButtonClick"
										'Select the WorkList Tree Node
										If Trim(dicItems(iCounter)) <> "" Then		
											objDialog.JavaButton(dicItems(iCounter)).Click micLeftBtn
											If Err.Number < 0 Then
												  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Button ["+dicItems(iCounter)+"] in Process Tree.")
												  Fn_WorkflowSubProcess_Operations = False         
												  Exit Function
											Else                      
												  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Button ["+dicItems(iCounter)+"] in Process Tree.")
												  Call Fn_ReadyStatusSync(1)
												   Wait(1)
											End If
										End If

				   End Select
			End if
		 Next
		 End Select

		' Close the Attributes Dialog
		If Trim(sBtnName) <> "" and (Trim(sBtnName) = "OK" OR Trim(sBtnName) = "Cancel")  Then
			objDialog.JavaButton(sBtnName).Click
				 If Err.Number < 0 Then
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to ["+CStr(sBtnName)+"] Button in  Workflow Sub-Process ... Dialog.")
					   Fn_WorkflowSubProcess_Operations = False
					   Exit Function
				 Else                      
					   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully ["+CStr(sBtnName)+"] Button in  Workflow Sub-Process ... Dialog.")
					   Call Fn_ReadyStatusSync(2)
						Wait(2)
				 End If
		End If
	Else
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Workflow Sub-Process ... Dialog does not exist.")
			   Fn_WorkflowSubProcess_Operations = False
			   Exit Function
	End If
	Fn_WorkflowSubProcess_Operations = True
End Function


'*********************************************************		Function to Perform Operations on the java table in the Viewer Tab		***********************************************************************
'Function Name		:				Fn_MyWorkList_ViwerPaneProperties(sAction,sValue)

'Description			 :		 		 Perform operations on Java table in the Viewer Tab

'Parameters			   :	 			1. sAction : Action Verify User, Description...													 
'													 2.	aValue -  Name of the User in the Table

'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Example				:			Call Fn_MyWorkList_ViwerPaneProperties("VerifyUser","AutoTest1 (autotest1)-Engineering/")

'												Call Fn_MyWorkList_ViwerPaneProperties("VerifyDecision",Array("0:Approve,"1:Reject"))	

'History					 :		
'													Developer Name									Date							Rev. No.						Changes Done						Reviewer
'								------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Omkar Kulkarni							   	  08/10/2010			 			    1.0										Created								Prasanna
'								------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public function Fn_MyWorkList_ViwerPaneProperties(sAction,aValue)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_ViwerPaneProperties"
	Dim bFlag,iCount,aValSplit,iRowCount,iCounter,jCounter, sUser,sDescription
	bFlag=False
	Fn_MyWorkList_ViwerPaneProperties=False

'Check the Existance of the Viewer Tab

	bReturn =  Fn_MyTc_TabSet("Viewer")	
	If bReturn = false Then
							Fn_MyWorkList_ViwerPaneProperties = False									
						   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :  The Viewer tab does not Exist")
							Exit Function
					Else 
							Fn_MyWorkList_ViwerPaneProperties = true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS :  The Viewer tab Opened Successfully")
							Call Fn_ReadyStatusSync(3)
							wait(3)
	End If	

Select Case  sAction

    Case "VerifyUser"

					If IsArray(aValue)  Then
								'If the value of User matches the one specified as Parameter ,log the Success 
								iRowCount = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("UserTable").GetROProperty("rows")
								For iCounter = 0 to Ubound(aValue)
											For jCounter = 0 to iRowCount 
													   If  JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("UserTable").GetCellData(jCounter,"User-Group/Role")=aValue(jCounter) Then            																								
																bFlag = true
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Successfully Verifed  that the user in the Viewer tab"+sUser)
																Exit For													
													 End If                                            
											Next
								Next						 
					End If
			
					If bFlag = false Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL :  Verification failed for User-Group Role Value")
								Fn_MyWorkList_ViwerPaneProperties=false		
					 Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Verified Successfully User-Group Role Value")
								Fn_MyWorkList_ViwerPaneProperties=True		
					End If

	   Case "VerifyDecision"

					If IsArray(aValue)  Then
								'If the value of User matches the one specified as Parameter ,log the Success 
								iRowCount = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("UserTable").GetROProperty("rows")
								For iCounter = 0 to Ubound(aValue)											
													   aValSplit = split(aValue(jCounter),":",-1,1)
													   If  JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("UserTable").GetCellData(aValSplit(0),"Decision")=aValSplit(1) Then            																								
																bFlag = true
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Successfully Verifed  that the Decision in the Viewer tab"+sUser)
																Exit For													
													 End If                                            											
								Next						 
					End If
			
					If bFlag = false Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL :  Verification failed for Decision Value")
								Fn_MyWorkList_ViwerPaneProperties=false		
					 Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Verified Successfully Decision Value")
								Fn_MyWorkList_ViwerPaneProperties=True		
					End If
		
		Case "VerifyComments"

					If IsArray(aValue)  Then
								'If the value of User matches the one specified as Parameter ,log the Success 
								iRowCount = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("UserTable").GetROProperty("rows")
								For iCounter = 0 to Ubound(aValue)											
													   aValSplit = split(aValue(jCounter),":",-1,1)
													   If  JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTable("UserTable").GetCellData(aValSplit(0),"Comments")=aValSplit(1) Then            																								
																bFlag = true
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Successfully Verifed  that the Comments in the Viewer tab"+sUser)
																Exit For													
													 End If                                            											
								Next						 
					End If
			
					If bFlag = false Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL :  Verification failed for comment Value")
								Fn_MyWorkList_ViwerPaneProperties=false		
					 Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Verified Successfully comment Value")
								Fn_MyWorkList_ViwerPaneProperties=True		
					End If

		Case "VerifyState"

							If IsArray(aValue) Then
									'If the state matches with that of the parameter passed by the user.
									iRowCount=JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaStaticText("StateStatus").GetROProperty("label")
									 For  iCounter=0 to Ubound(aValue)
													If iRowCount =aValue(iCounter) Then
															bFlag=true
															 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Successfully Verifed  that the state of the task in the viewer tab is "+iRowCount)
															Exit For
													End If
									Next
							End If
	
							If bFlag = false Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL :  Verification failed for Status of the task")
														Fn_MyWorkList_ViwerPaneProperties=false		
												 Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Verified Successfully the Status of the task")
														Fn_MyWorkList_ViwerPaneProperties=True		
							End If

			Case "VerifyProcessDescription"
							sDescription = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaEdit("Process Description").GetROProperty("value")
							If Trim(Lcase(aValue)) = Trim(Lcase(sDescription)) Then
									' writing Log 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS:Fail To  Verify  that the Process Description of the task in the viewer tab is ["+sDescription+ "] ")
									Fn_MyWorkList_ViwerPaneProperties = TRUE
								else 
									 ' writing Log 
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Successfully Verifed  that the Process Description of the task in the viewer tab is ["+sDescription+ "] ")	
									Fn_MyWorkList_ViwerPaneProperties = FALSE
									
							End If

      End Select

End Function

'*********************************************************		Function to Perform Operations on the java table in the Viewer Tab		***********************************************************************
'Function Name		:				Fn_MyWorkList_Menu_TaskComplete(sAction, sTaskName, sTaskInstruction, sProcessDesc, sComment, radComplete, sPassword)

'Description			 :		 		 Perform Task Complete from Menu (Action: Perform)

'Parameters			   :	 			1. sAction : Action Verify / SignOff									 
'													 2.	sTaskName - Task Name
'													 3.	sTaskInstruction - Task Instructions
'													 4.	sProcessDesc - Process Description
'													 5.	sComment - Comments
'													 6.	radComplete - Radio Button Name
'													 7.	sPassword - Password

'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Example				:			Call Fn_MyWorkList_Menu_TaskComplete("Verify", "New Do Task 1", "", "Test Description", "Test Comments", "Complete", "")

'											Call Fn_MyWorkList_Menu_TaskComplete("SignOff", "", "", "Test Description", "Test Comments", "Complete", "")

'History					 :		
'													Developer Name									Date							Rev. No.						Changes Done						Reviewer
'								------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Mahendra Bhandarkar				   	  14/10/2010			 			    1.0										Created								Prasanna
'								------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_Menu_TaskComplete(sAction, sTaskName, sTaskInstruction, sProcessDesc, sComment, radComplete, sPassword)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_Menu_TaskComplete"
	Dim objDialog

	If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Do Task").Exist(2) = False And JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Condition Task").Exist(2) = False Then
		If JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Perform Do Task").Exist(2) = False And JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Perform Condition Task").Exist(2) = False Then
			bReturn = Fn_MenuOperation("Select", "Actions:Perform")
			Call Fn_ReadyStatusSync(5)		   	    
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Perform Menu Operation [Actions:Perform]") 		
				Fn_MyWorkList_Menu_TaskComplete = False
				Set objDialog = Nothing
				Set objCondDialog = Nothing
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Performed Menu Operation [Actions:Perform]")
				Call Fn_ReadyStatusSync(2)
			End If
	  End If
   End If



        If 	JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Do Task").Exist Then
					Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Do Task")
		Elseif  JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Condition Task").Exist Then
					Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Condition Task")
		 Elseif JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("Perform Do Task").Exist Then
					Set objDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("Perform Do Task")
		Elseif JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Perform Do Task").Exist Then
					Set objDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Perform Do Task")
		 Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL :Perform Task Dialog does not Exist") 	
					Exit function
	    End If 

	If objDialog.Exist(2) = True Then

		Select Case sAction
						Case "SignOff"
									If  Trim(sTaskName) <> "" Then
										objDialog.JavaEdit("Task Name").Set sTaskName
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sTaskName+"] in Task Name.") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sTaskName+"] in Task Name.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(sTaskInstruction) <> "" Then
										objDialog.JavaEdit("Task Instructions").Set sTaskInstruction
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sTaskInstruction+"] in Task Instructions.") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sTaskInstruction+"] in Task Instructions.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(sProcessDesc) <> "" Then
										objDialog.JavaEdit("Process Description").Set sProcessDesc
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sProcessDesc+"] in Process Description.") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sProcessDesc+"] in Process Description.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(sComment) <> "" Then
										objDialog.JavaEdit("Comments").Set sComment
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sComment+"] in Comment.") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sComment+"] in Comment.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(radComplete) <> "" Then
										objDialog.JavaRadioButton("Complete").SetTOProperty "attached text", radComplete
										objDialog.JavaRadioButton("Complete").Set "ON"
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Checked Radio Button ["+radComplete+"] ") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Checked Radio Button ["+radComplete+"].")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(sPassword) <> "" Then
										objDialog.JavaEdit("Password").Set sPassword
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sPassword+"] in Password ") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sPassword+"] in Password.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									objDialog.JavaButton("OK").Click micLeftBtn
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Ok Button ") 		
										Fn_MyWorkList_Menu_TaskComplete = False
										objDialog.JavaButton("Cancel").Click micLeftBtn
										Set objDialog = Nothing
										Exit Function
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on OK Button.")
										Call Fn_ReadyStatusSync(5)
									End If

						Case "Verify"
									If  Trim(sTaskName) <> "" Then
										If objDialog.JavaEdit("Task Name").GetROProperty("value") <> sTaskName Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sTaskName+"] in Task Name.") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sTaskName+"] in Task Name.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(sTaskInstruction) <> "" Then
										If objDialog.JavaEdit("Task Instructions").GetROProperty("value") <> sTaskInstruction Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sTaskInstruction+"] in Task Instructions.") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sTaskInstruction+"] in Task Instructions.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(sProcessDesc) <> "" Then
										If objDialog.JavaEdit("Process Description").GetROProperty("value") <> sProcessDesc Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sProcessDesc+"] in Process Description.") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sProcessDesc+"] in Process Description.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(sComment) <> "" Then
										If Replace(objDialog.JavaEdit("Comments").GetROProperty("value"),chr(10),"") <> sComment Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sComment+"] in Comment.") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sComment+"] in Comment.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(radComplete) <> "" Then
										objDialog.JavaRadioButton("Complete").SetTOProperty "attached text", radComplete
										If objDialog.JavaRadioButton("Complete").GetROProperty("value") = "0" Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Checked Radio Button ["+radComplete+"] ") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Checked Radio Button ["+radComplete+"].")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									If  Trim(sPassword) <> "" Then
										If objDialog.JavaEdit("Password").GetROProperty("value") <> sPassword Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ["+sPassword+"] in Password ") 		
											Fn_MyWorkList_Menu_TaskComplete = False
											objDialog.JavaButton("Cancel").Click  micLeftBtn
											Set objDialog = Nothing
											Exit Function
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set ["+sPassword+"] in Password.")
											Call Fn_ReadyStatusSync(2)
										End If
									End If

									objDialog.JavaButton("Cancel").Click
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Cancel Button ") 		
										Fn_MyWorkList_Menu_TaskComplete = False
										objDialog.JavaButton("Cancel").Click micLeftBtn
										Set objDialog = Nothing
										Exit Function
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Cancel Button.")
										Call Fn_ReadyStatusSync(2)
									End If
		Case "VerifyMessage"
				If instr(1,sPassword,":")>1 Then
					aError=Split(sPassword,":",-1,1)
				End If
			If JavaWindow("MyTeamcenter").JavaWindow("Error").Exist then
				JavaWindow("MyTeamcenter").JavaWindow("Error").JavaStaticText("ErrLabel").SetTOProperty "label",aError(0)
				sValue= JavaWindow("MyTeamcenter").JavaWindow("Error").JavaEdit("Details").GetROProperty ("value")
				 If JavaWindow("MyTeamcenter").JavaWindow("Error").JavaStaticText("ErrLabel").Exist and sValue=aError(1) Then
					 Fn_MyWorkList_Menu_TaskComplete = True
					 JavaWindow("MyTeamcenter").JavaWindow("Error").Close
				Else
					 Fn_MyWorkList_Menu_TaskComplete = false
					 JavaWindow("MyTeamcenter").JavaWindow("Error").Close
				 End If
			End if

		End Select

	Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Perform Do Task Dialog does not exist. ") 		
									Fn_MyWorkList_Menu_TaskComplete = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
	End If

	Fn_MyWorkList_Menu_TaskComplete = True
	Set objDialog = Nothing

End Function

'*********************************************  Fn_MyWorkList_ProcessView_ProcessTreeOperations**************************************************************

'Function Name		:					Fn_MyWorkList_ProcessView_ProcessTreeOperations(sAction, sNodeName, sMenu)

'Description			 :		 		  Node Operations on Process Tree in Process View Panel

'Parameters			   :	 			sAction - Action to be performed
'													sNodeName - Node name
'													sMenu - For ContextMenu (Implementation Pending)

'Return Value		   : 				True/False

'Pre-requisite			:		 		Node Selected in MyWorkList 

'Examples				:				MsgBox Fn_MyWorkList_ProcessView_ProcessTreeOperations("DoubleClick", "AutoDoReview:New Review Task 1:select-signoff-team", "")
'
'History:
'										Developer Name							Date				Rev. No.			Build			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'											Mahendra Bhandarkar				19-Oct-2010	   		1.0				902
'-------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function  Fn_MyWorkList_ProcessView_ProcessTreeOperations(sAction, sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_ProcessView_ProcessTreeOperations"
   On Error Resume Next
   Dim arrNodeList,objDialog,iItemCount,iCounter,sTreeItem,arrNode,iOuterCount,aMenuList, objDialogView

   Set objDialog =  JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaTree("ProcessTree")

	If objDialog.Exist(2) = False Then
			' Set the Viewer tab
			bReturn = Fn_MyTc_TabSet("Viewer")
			If Err.Number < 0 Then
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to set Viewer tab")
					Set objDialog = Nothing
					Exit Function						
			Else
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Viewer tab")
					Wait(2)
			End If

			 'Set Default View to Process View 
			 Set objDialogView = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
			 objDialogView.JavaRadioButton("ViewOptions").SetTOProperty "Attached Text","Process View"
			 objDialogView.JavaRadioButton("ViewOptions").Set "ON"
			 If Err.Number < 0 Then
				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Process View from Viewer tab.")
				   Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
					Set objDialog = Nothing
					Set objDialogView = Nothing
				   Exit Function
			 Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Process View from Viewer tab.")
					Call Fn_ReadyStatusSync(2)
					 Wait(2)
			 End If
			 Set objDialogView = Nothing
	End If

   If objDialog.Exist(5) Then
	
		Select Case sAction

			Case  "Select"
				objDialog.Select sNodeName
				If Err.Number < 0 Then
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node [" + sNodeName + "] of Process Tree." )	
					Set objDialog = Nothing
					Exit Function
				Else
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected  Node [" + sNodeName + "] of Process Tree.")	
				End If

			Case "MultiSelect" 
				arrNode = Split(sNodeName,",")
				For iOuterCount = 0 to  Ubound(arrNode)
					sTreeItem = arrNode(iOuterCount)
					If iOuterCount = 0 Then
							objDialog.Select sTreeItem
					Else
							objDialog.ExtendSelect sTreeItem
					End If
					If Err.Number < 0 Then
						Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select Node [" + sTreeItem + "] of Process Tree." )	
						Set objDialog = Nothing
						Exit Function 
					Else
						Fn_MyWorkList_ProcessView_ProcessTreeOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected Node [" + sTreeItem  + "] of Process Tree.")	
					End If
				Next

			Case  "Expand"
				objDialog.Expand sNodeName
				If Err.Number < 0 Then
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to expand node [" + sNodeName + "] of Process Tree." )	
					Set objDialog = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully expand node [" + sNodeName  + "] of Process Tree.")	
				End If

			Case "Collapse"
				objDialog.Collapse sNodeName
				If Err.Number < 0 Then
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Collapse node [" + sNodeName + "] of Process Tree." )	
					Set objDialog = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Collapse node [" + sNodeName  + "] of Process Tree.")	
				End If

			Case "Exist"
				iItemCount = objDialog.GetROProperty( "items count")
				For iCounter=0 To (iItemCount-1)
					sTreeItem = objDialog.GetItem(iCounter)
					If Trim (Lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_MyWorkList_ProcessView_ProcessTreeOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully found node [" + sNodeName + "] of Process Tree." )	
						Exit For
					End If
				Next 

				If  Cint(iCounter) = Cint (iItemCount) Then
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to  found node [" + sNodeName + "] of Process Tree." )	
					Set objDialog = Nothing
					Exit Function 
			    End If

			Case  "DoubleClick"
				If Trim(sNodeName) <> "" Then
					objDialog.Select sNodeName
				End If
				objDialog.Activate sNodeName
				If Err.Number < 0 Then
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Double Clicked the Selected Node [" + sNodeName + "] of Process Tree." )	
					Set objDialog = Nothing
					Exit Function 
				Else
					Fn_MyWorkList_ProcessView_ProcessTreeOperations = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Double Clicked on the Selected  Node [" + sNodeName + "] of Process Tree.")	
				End If

			Case Else
				Fn_MyWorkList_ProcessView_ProcessTreeOperations = False

		End Select
   Else
		Fn_MyWorkList_ProcessView_ProcessTreeOperations = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Process Tree does not exist.")	
		Exit Function
   End If
	Set objDialog = Nothing

End Function


'*********************************************  Fn_MyWorkList_HTML_PropertyVerify **************************************************************

'Function Name		:					Fn_MyWorkList_HTML_PropertyVerify(sTitle, sPropertyHead1,sPropertyValue1, sPropertyHead2, sPropertyValue2, bClose)

'Description			 :		 		  HTML Property verifies

'Parameters			   :	 			sTitle - Browser's Title
'													 sPropertyHead1 - Heading of PropertyValue1
'													 sPropertyValue1 - Property Value of Heading 1 [ Seperated by , for Multiple Values]
'													 sPropertyHead2 - Heading of PropertyValue2
'													 sPropertyValue2 - Property Value of Heading 2 [ Seperated by , for Multiple Values]
'													 bClose - Close the Browser or not [ True/ False]

'Return Value		   : 				True/False

'Pre-requisite			:		 		Property Browser should be opened.

'Examples				:				MsgBox Fn_MyWorkList_HTML_PropertyVerify("Admin - Audit", "Action","Modify: Modify", "Date", "13-Oct-2010:13-Oct-2010", False)
'
'History:
'										Developer Name							Date				Rev. No.			Build			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'											Mahendra Bhandarkar				19-Oct-2010	   		1.0				902
'-------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_HTML_PropertyVerify(sTitle, sPropertyHead1,sPropertyValue1, sPropertyHead2, sPropertyValue2, bClose)
		GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_HTML_PropertyVerify"
		Dim ObjInv, arrPropertyName, arrPropertyVal, Rwcnt, objTable, iCounter, jRowConter, ObjClose, arrResult,  bCheckProp, iValueCounter , iArrSize        
		Dim arrSpecificString,iCountSpecific
		Dim sMatchHead, sPNIndex, sPVIndex, ColCnt, getHeadTitle, bFlag

		If Browser("Browser").page("Page").Exist(5)Then

			If Trim(sTitle) <> "" Then
				Browser("Browser").Page("Page").WebElement("HeadTitle").SetTOProperty "innertext", Trim(sTitle)
				If Browser("Browser").Page("Page").WebElement("HeadTitle").Exist Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Opened Page Heading Title ["+CStr(sTitle)+"] verified successfully.")
				Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Opened Page Heading Title ["+CStr(sTitle)+"] failed to verify.")
							Fn_MyWorkList_HTML_PropertyVerify = False
							Exit Function
				End If
			End If

			arrPropertyName = Split(sPropertyValue1,":",-1)
			arrPropertyVal = Split(sPropertyValue2,":",-1)				

			' Setting the Column Index for PropertyHeading1 and PropertyHeading2
			ColCnt = Browser("Browser").Page("Page").WebTable("PropertyTable").GetROProperty("cols")
			' For PropertyHeading1
			If Trim(sPropertyHead1) <> "" Then
				For iCounter = 1 To ColCnt
					getHeadTitle = Browser("Browser").Page("Page").WebTable("PropertyTable").GetCellData(1, iCounter)
					getHeadTitle = REPLACE(getHeadTitle, CHR(130), " ")
					getHeadTitle = REPLACE(getHeadTitle, CHR(32), " ")
					getHeadTitle = Trim(getHeadTitle)					
					If LCase(getHeadTitle) = LCase(sPropertyHead1) Then
						sPNIndex = iCounter
						Exit For
					End If
				Next
			End If

			' For PropertyHeading2
			If Trim(sPropertyHead2) <> "" Then
				For iCounter = 1 To ColCnt
					getHeadTitle = Browser("Browser").Page("Page").WebTable("PropertyTable").GetCellData(1, iCounter)
					getHeadTitle = REPLACE(getHeadTitle, CHR(130), " ")	
					getHeadTitle = REPLACE(getHeadTitle, CHR(32), " ")
					getHeadTitle = Trim(getHeadTitle)					
					If LCase(getHeadTitle) = LCase(sPropertyHead2) Then
						sPVIndex = iCounter
						Exit For
					End If
				Next
			End If

				'************************Here we count the Number of rows present in that perticular table**************************
				Rwcnt = Browser("Browser").Page("Page").WebTable("PropertyTable").RowCount
				Set objTable = Browser("Browser").Page("Page").WebTable("PropertyTable")		
				iValueCounter  = 0
				iCounter = 0
				iArrSize = UBound(arrPropertyName)													
				arrResult = arrPropertyName
				Do
						  bCheckProp = False
						  For jRowConter = 1 to Rwcnt
								' Comparing the CellData with the arguments	  
								If Trim(objTable.GetCellData(jRowConter, sPNIndex)) = Trim(arrPropertyName(iCounter)) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Property Found for [" + arrPropertyName(iCounter)+"]")
										If InStr(1, Trim(objTable.GetCellData(jRowConter, sPVIndex)), Trim(arrPropertyVal(iValueCounter)), 1) > 0 Then 
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Object Value Verified for [" +  arrPropertyName(iCounter)+"]")
													iValueCounter = iValueCounter + 1
													arrResult(iCounter) = True
													bCheckProp = True
													Exit For                                     																						
										 End If
								   End If
							Next
							If bCheckProp = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Object Property Not Found for [" + arrPropertyName(iCounter)+"]")
								arrResult(iCounter) = False
							End If
							iCounter  = iCounter +1									
				Loop While iCounter <= iArrSize
		Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Browser window does not exist.")
		End If                          				

		bFlag = True
		For iCounter = 0 to iArrSize
					If arrResult(iCounter) = False Then
							Fn_MyWorkList_HTML_PropertyVerify = False
							bFlag = False
							Exit for
					End If
		Next

		If CBool(bClose) = True Then
			Browser("Browser").Close
			If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Close Opened Browser.")
						Fn_MyWorkList_HTML_PropertyVerify = False
						Exit Function
			Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Browser Closed Successfully.")
			End If
		End If

		If bFlag = True Then
			Fn_MyWorkList_HTML_PropertyVerify = True
		End If

End Function

'*********************************************  Fn_MyWorkList_TaskActions **************************************************************

'Function Name		:					Fn_MyWorkList_TaskActions(sAction, sWorkListNode, sOption, sComment, sButton)

'Description			 :		 		  The function performs Actions over MyWorkList Tasks

'Parameters			   :	 			sAction - Add/ Verify
'													sWorkListNode - On Which Task to be performed
'													sOption - Perform [Resume/Complete]
'
'Return Value		   : 				True/False

'Pre-requisite			:		 		MyWorkList Tree Opened

'Examples				:				Call Fn_MyWorkList_TaskActions("Add", "My Worklist:AutoTest1 (autotest1) Inbox:Tasks to Perform:001511/A;1-kkkk (New Route Task 1)", "Resume", "Resume comments", "OK")
'
'History:
'										Developer Name							Date				Rev. No.			Build			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'											Mahendra Bhandarkar				21-Oct-2010	   		1.0				916
'-------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_TaskActions(sAction, sWorkListNode, sOption, sComment, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TaskActions"
	Dim objDialog, bReturn, sMenu

		If Trim(sWorkListNode) <> "" Then
					bReturn = Fn_MyWorkList_TreeNodeOperations("Select", sWorkListNode,"")
					If bReturn = False Then
								Fn_MyWorkList_TaskActions = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select  [" + sWorkListNode+"] from Worklist" )
								Set objDialog = Nothing
								Exit Function
					Else
								 wait(3)
								 Call Fn_ReadyStatusSync(5) 
								 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected  [" + sWorkListNode+"] from WorkList" )	
					End If
		End If

		If Trim(sOption) = "Resume" Then
					Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Resume Action Comments")
					sMenu = "Actions:Resume"
		ElseIf Trim(sOption) = "Complete" Then
					Set objDIalog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Complete Action Comments")
					sMenu = "Actions:Complete"
		Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Wrong action ["+CStr(sTitle)+"] entered.")
					Fn_MyWorkList_TaskActions = False
					Exit Function
		End If

		If objDialog.Exist(2) = False Then
					bReturn = Fn_MenuOperation("Select", sMenu)
					Call Fn_ReadyStatusSync(3)
					If bReturn = False Then
								Fn_MyWorkList_TaskActions = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Invoke  Menu [" + sMenu+"]." )	
								Exit Function
					Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Invoked  Menu [" + sMenu+"]." )	
					End If
		End If

		Wait 2

		If objDialog.Exist Then
					Select Case Trim(sAction)
								Case "Add"
												If Trim(sComment) <> "" Then
														'Set the Comments
														objDialog.JavaEdit("Comments").Set Trim(sComment)
														If Err.Number < 0 Then
																Fn_MyWorkList_TaskActions = False
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Comment [" + sComment+"].")	
																objDialog.JavaButton("Cancel").Click micLeftBtn
																Set objDialog = Nothing
																Exit Function
														Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Entered Comment [" + sComment+"].")
																Call Fn_ReadyStatusSync(2)
														End If
												End If

												'Click on button
												If Trim(sButton) <> "" Then
														objDialog.JavaButton(sButton).Click micLeftBtn
														If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+CStr(sButton)+"] Button ") 		
															Fn_MyWorkList_TaskActions = False
															objDialog.JavaButton("Cancel").Click micLeftBtn
															Set objDialog = Nothing
															Exit Function		
														Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on ["+CStr(sButton)+"] Button")
															Call Fn_ReadyStatusSync(2)
														End If
												End If

												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Performed ["+CStr(sAction)+"] for Task [" + sWorkListNode+"]")
	
						Case "Verify"						' For feature use			
			
				 End Select
		Else
				Fn_MyWorkList_TaskActions = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Find ["+CStr(sAction)+" Action Comments] Dialog" )	
				Exit Function
		End If

		Fn_MyWorkList_TaskActions = True
		Set objDialog = Nothing
End Function

'-------------------------------------------------------------------Function Used to perform operations on { SelectSignoffTeamDialog } Dialog----------------------------------------------------------------
'Function Name		:	Fn_MyWorkList_SignoffTeamSelectDialog_Opearations

'Description			 :	Function Used to perform operatons on  Trees which are present on { SelectSignoffTeamDialog } Dialog

'Parameters			   :	1.strAction: Action Name
'									2.dicSelectSignoff:
										
										
'Return Value		   : 	True Or False

'Pre-requisite			:	
'									dicSelectSignoff("SignOffTeamSelect")	= "Signoff Team:Profiles:dba/DBA/1"
'Examples				:	Fn_MyWorkList_SignoffTeamSelectDialog_Opearations("VerifySignOffTeamTreeForSST",dicSelectSignoff)
										   
'History				:			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				25/10/2010			           1.0																					 Prasanna B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_SignoffTeamSelectDialog_Opearations(strAction,dicSelectSignoff)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_SignoffTeamSelectDialog_Opearations"
	'Declaring variables
	Dim intNodeCount,intCount,strTreeItem
	Dim ObjSignOffTmDialog,strNode
	Dim sUserName,iCounter, iCount
	Dim objSignOffTree, sTreeName
	Dim aSplitNodeName, sTempNode
	
	'Creating Object of SelectSignoffTeamDialog
	If JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").Exist Then
				Set ObjSignOffTmDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("SelectSignoffTeamDialog")
	else
				Set ObjSignOffTmDialog = JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks").JavaDialog("SelectSignoffTeamDialog")	
	End If
   Select Case strAction
		 	Case "VerifySignOffTeamTreeForSST","VerifySignOffTeamTree"
					
					If  dicSelectSignOff.Item("WorkListTreeNode") <> "" Then
					
							bReturn =  Fn_MyWorkList_TreeNodeOperations("Select",dicSelectSignOff.Item("WorkListTreeNode"),"")
								If bReturn = false Then
										Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = False									
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + dicSelectSignOff.Item("WorkListTreeNode") +  " From WorkList Tree")
										Exit Function
								End If
								Call Fn_ReadyStatusSync(5)
					
					End If
					
					'Checking Existance of SelectSignoffTeamDialog 
					If  ObjSignOffTmDialog.Exist(5)=False Then
							'Invoking SelectSignoffTeamDialog by Calling Menu
							Call Fn_MenuOperation("Select","Actions:Perform")
							Call Fn_ReadyStatusSync(5)
					End If

					'Added by Nilesh on 13-Jun-12 TC10 Build 0606 
					If ObjSignOffTmDialog.Exist(5)=False  Then
						Set ObjSignOffTmDialog=JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("SelectSignoffTeamDialog")
					End If

					If  dicSelectSignOff.Item("SignOffTeamSelect") <> "" Then
							
							strNode = dicSelectSignOff.Item("SignOffTeamSelect")
							'Taking Items Count Of tree Nodes
							intNodeCount = ObjSignOffTmDialog.JavaTree("SignOffTeamTree").GetROProperty("items count")
							For intCount = 0 to intNodeCount - 1
								'Taking Item From Tree
								'strTreeItem =ObjSignOffTmDialog.JavaTree("SignOffTeamTree").GetItem(intCount)
								strTreeItem =ObjSignOffTmDialog.JavaTree("SignOffTeamTree").object.getPathForRow(intCount).tostring()
								strTreeItem = Replace(strTreeItem,"[","")
								strTreeItem = Replace(strTreeItem,"]","")
								strTreeItem = Replace(strTreeItem,", ",":")
								'matching Current Item With Users Item
								If Trim(lcase(strTreeItem)) = Trim(Lcase(strNode)) Then
									'If item Found In Tree Function returns true
									Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = True
									Exit For
								End If
							Next
							If Cint(intCount) = Cint(intNodeCount) Then
								'If item Not Found In Tree Function returns true
								Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = False
							End If
							
							If strAction<> "VerifySignOffTeamTree"Then
								'Closing The Dialog
								ObjSignOffTmDialog.JavaButton("Close").Click
							End If
						
					End if
			'[TC11.4(20171201.00)_NewDevelopment_Minal N_18Jan2018 : Added Case to remove user from signoffteamtree]				
			Case "RemoveUserFromSignOffTeamTree"				
					If  dicSelectSignOff.Item("WorkListTreeNode") <> "" Then					
						bReturn =  Fn_MyWorkList_TreeNodeOperations("Select",dicSelectSignOff.Item("WorkListTreeNode"),"")
						If bReturn = false Then
							Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = False									
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + dicSelectSignOff.Item("WorkListTreeNode") +  " From WorkList Tree")
							Exit Function
						End If
						Call Fn_ReadyStatusSync(5)					
					End If
					
					'Checking Existance of SelectSignoffTeamDialog 
					If  ObjSignOffTmDialog.Exist(5)=False Then
						'Invoking SelectSignoffTeamDialog by Calling Menu
						Call Fn_MenuOperation("Select","Actions:Perform")
						Call Fn_ReadyStatusSync(5)
					End If

					'Added by Nilesh on 13-Jun-12 TC10 Build 0606 
					If ObjSignOffTmDialog.Exist(5)=False  Then
						Set ObjSignOffTmDialog=JavaWindow("WorkflowViewerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("SelectSignoffTeamDialog")
					End If

					If dicSelectSignOff.Item("UsersName") <> "" Then		
						Set objSignOffTree = ObjSignOffTmDialog.JavaTree("SignOffTeamTree")
						For iRowCounter = 0 to Cint(objSignOffTree.GetROProperty("items count"))-1
							sTempNode = objSignOffTree.Object.getPathForRow(iRowCounter).tostring()
							Err.Clear
							If instr(1,sTempNode,dicSelectSignOff.Item("UsersName")) Then
								objSignOffTree.Object.setSelectionRow iRowCounter
								If  Err.Number < 0 Then
									Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = False							
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ["+dicSelectSignOff.Item("UsersName")+"] SignOff Team")
									ObjSignOffTmDialog.JavaButton("Close").Click micLeftBtn
									Exit Function
								End If
								Call Fn_ReadyStatusSync(5)
								Exit for
							End If
						Next
						If iRowCounter = Cint(objSignOffTree.GetROProperty("items count")) Then
							Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Exist "+ dicSelectSignOff.Item("UsersName") +" in Signoff Team Tree")
							ObjSignOffTmDialog.JavaButton("Close").Click micLeftBtn
							Exit Function
						End If
						Call Fn_ReadyStatusSync(1)
						'Click on Remove button
						If Fn_SISW_UI_Object_Operations("Fn_MyWorkList_SignoffTeamSelectDialog_Opearations", "Enabled", objSignOffTree, "") = False Then 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to removed User " + dicSelectSignOff.Item("UsersName")  + " from SignOff Tree")
							Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = False
							Exit Function
						End If
						ObjSignOffTmDialog.JavaButton("Remove").Click micLeftBtn
						Call Fn_ReadyStatusSync(5)
						'Click on OK button
						ObjSignOffTmDialog.JavaButton("OK").Click micLeftBtn
						Call Fn_ReadyStatusSync(2)
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to removed User " + dicSelectSignOff.Item("UsersName")  + " from SignOff Tree")
							Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = False
							Exit Function
						Else
							Fn_MyWorkList_SignoffTeamSelectDialog_Opearations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully removed User " + dicSelectSignOff.Item("UsersName") + " from SignOff Tree")				
						End If
					End if
   End Select
   'releasing ObjSignOffTmDialog object
   Set ObjSignOffTmDialog=Nothing
End Function

'*********************************************************		Function to verify  dialog error message.	***********************************************************************

'Function Name		:					Fn_MyWorkList_DialogMsgVerify

'Description			 :		 		  This function is used to Error Message Verify.

'Parameters			   :	 			1. sAction: Action to be performed.
'										2. sTitle:Title of dialog.
'										3. sMsg : Message to verify. (Optional)
'										4. sButton : Button Name.
											
'Return Value		   : 				True/False

'Pre-requisite			:		 		Error Message window should be displayed .

'Examples				:				Fn_MyWorkList_DialogMsgVerify("", "Information","The action was successful.","OK")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rima Patil		      03-Jan-2011	   		1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_DialogMsgVerify(sAction, sTitle,sMsg,sButton) 
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_DialogMsgVerify"
	GBL_EXPECTED_MESSAGE=sMsg
   On Error Resume Next
	Dim sResult, diaCreatePref, btnOK, lblMsg, tmp, sErrorMsg

	Fn_MyWorkList_DialogMsgVerify = True

	' Create Object Description of  Dialog 
	Set diaCreatePref=description.Create()
	diaCreatePref("micclass").value="Dialog"
	diaCreatePref("regexpwndtitle").value = sTitle
	diaCreatePref("regexpwndclass").value = "#32770"

	'Description of  Button Object  on  dialog
	Set btnOK=description.Create()
	btnOK("micclass").value="WinButton"
	btnOK("nativeclass").value = "Button"
	btnOK("regexpwndtitle").value = sButton

	'Set Titles of different types of error dialogs
	JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title", sTitle
	JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("JavaErrorDialog").SetTOProperty "title", sTitle
    JavaDialog("Error").SetTOProperty "title", sTitle

	'General Object description to search all Objects
	Set lblMsg=description.Create()

	Select Case sAction
	
			Case "DetailsMsg"
			
					If Dialog(diaCreatePref).Exist(5) Then
		
						'Capture All runtime objects to find message text
						Set lblMsg=description.Create()
						
						Set  tmp = Dialog(diaCreatePref).ChildObjects(lblMsg)
		
						'Set message text to variable 
						sErrorMsg = tmp(6). getroproperty("text")
		
						'compare run time message to verify  the error message
						If sMsg <> "" Then
							If (sMsg = sErrorMsg ) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")
							Else
								GBL_ACTUAL_MESSAGE=sErrorMsg
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
								Fn_MyWorkList_DialogMsgVerify = False
							End If
						End If
		
					' To Click "OK" Button after verification
						wait(2)
						Dialog(diaCreatePref).WinButton(btnOK).Click 10,10,micLeftBtn
							If Dialog(diaCreatePref).Exist(5) Then
								Dialog(diaCreatePref).Close()
							End If

						ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(5)   Then
							If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaCheckBox("More").Exist(1) Then
								JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaCheckBox("More").Set "ON"
							End If
							Wait 2
							If sMsg <> "" Then
									sErrorMsg = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaEdit("DetailMsg").GetROProperty("value")								
		
									If Trim(sErrorMsg) = Trim(sMsg) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")										
									Else
										GBL_ACTUAL_MESSAGE=sErrorMsg
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
										Fn_MyWorkList_DialogMsgVerify = False
									End If
							End if
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").SetTOProperty "label",  sButton
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
							Wait 2
					Else			
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The " + sTitle + " Dialog does not Exist")
						Fn_MyWorkList_DialogMsgVerify = False
					End If
					
			Case "MessageVerifyWithoutOk"
			
					If Dialog(diaCreatePref).Exist(5) Then
		
						'Capture All runtime objects to find message text
						Set  tmp = Dialog(diaCreatePref).ChildObjects(lblMsg)
		
						'Set message text to variable 
						sErrorMsg = tmp(1). getroproperty("text")  
		
						'compare run time message to verify  the error message
						If sMsg <> "" Then
							If (sMsg = sErrorMsg ) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")
							Else
								GBL_ACTUAL_MESSAGE=sErrorMsg
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
								Fn_MyWorkList_DialogMsgVerify = False
							End If
						End If
		
						ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(5)   Then
							If sMsg <> "" Then
									sErrorMsg = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaEdit("ErrText").GetROProperty("value")								
		
									If Trim(sErrorMsg) = Trim(sMsg) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")										
									Else
										GBL_ACTUAL_MESSAGE=sErrorMsg
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
										Fn_MyWorkList_DialogMsgVerify = False
									End If                                
							End if
		
					ElseIf 	JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("JavaErrorDialog").Exist(5) Then
							If sMsg <> "" Then
									sErrorMsg = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("JavaErrorDialog").JavaEdit("ErrText").GetROProperty("value")
		
									If Trim(sErrorMsg) = Trim(sMsg) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")										
									Else
										GBL_ACTUAL_MESSAGE=sErrorMsg
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
										Fn_MyWorkList_DialogMsgVerify = False
									End If                                
							End if
	
					ElseIf JavaDialog("Error").Exist(5) Then
							If sMsg <> "" Then
									sErrorMsg = JavaDialog("Error").JavaEdit("ErrMsg").GetROProperty("value")
		
									If Trim(sErrorMsg) = Trim(sMsg) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")										
									Else
										GBL_ACTUAL_MESSAGE=sErrorMsg
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
										Fn_MyWorkList_DialogMsgVerify = False
									End If                                
							End if
							
					Else			
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The " + sTitle + " Dialog does not Exist")
						Fn_MyWorkList_DialogMsgVerify = False
					End If
		
			Case Else
			
				If Dialog(diaCreatePref).Exist(5) Then
		
						'Capture All runtime objects to find message text
						Set  tmp = Dialog(diaCreatePref).ChildObjects(lblMsg)
		
						'Set message text to variable 
						sErrorMsg = tmp(1). getroproperty("text")  
		
						'compare run time message to verify  the error message
						If sMsg <> "" Then
							If (sMsg = sErrorMsg ) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")
							Else
								GBL_ACTUAL_MESSAGE=sErrorMsg
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
								Fn_MyWorkList_DialogMsgVerify = False
							End If
						End If
		
					' To Click "OK" Button after verification
						wait(2)
						Dialog(diaCreatePref).WinButton(btnOK).Click 10,10,micLeftBtn
							If Dialog(diaCreatePref).Exist(5) Then
								Dialog(diaCreatePref).Close()
							End If
						ElseIf JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(5)   Then
							If sMsg <> "" Then
									sErrorMsg = JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaEdit("ErrText").GetROProperty("value")								
		
									If Trim(sErrorMsg) = Trim(sMsg) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")										
									Else
										GBL_ACTUAL_MESSAGE=sErrorMsg
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
										Fn_MyWorkList_DialogMsgVerify = False
									End If                                
							End if
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").SetTOProperty "label",  sButton
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
		
					ElseIf 	JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("JavaErrorDialog").Exist(5) Then
							If sMsg <> "" Then
									sErrorMsg = JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("JavaErrorDialog").JavaEdit("ErrText").GetROProperty("value")
		
									If Trim(sErrorMsg) = Trim(sMsg) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")										
									Else
										GBL_ACTUAL_MESSAGE=sErrorMsg
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
										Fn_MyWorkList_DialogMsgVerify = False
									End If                                
							End if
		
							JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("JavaErrorDialog").JavaButton("OK").SetTOProperty "label", sButton
							JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("JavaErrorDialog").JavaButton("OK").Click micLeftBtn
		
					ElseIf JavaDialog("Error").Exist(5) Then
							If sMsg <> "" Then
									sErrorMsg = JavaDialog("Error").JavaEdit("ErrMsg").GetROProperty("value")
		
									If Trim(sErrorMsg) = Trim(sMsg) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")										
									Else
										GBL_ACTUAL_MESSAGE=sErrorMsg
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
										Fn_MyWorkList_DialogMsgVerify = False
									End If                                
							End if
		
							JavaDialog("Error").JavaButton("OK").SetTOProperty "label", sButton
							JavaDialog("Error").JavaButton("OK").Click micLeftBtn
				
					Else			
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The " + sTitle + " Dialog does not Exist")
						Fn_MyWorkList_DialogMsgVerify = False
					End If
					
			End Select
	
	Set diaCreatePref=nothing
	Set btnOK=nothing
	Set lblMsg=nothing
	Set tmp=nothing

End Function    
'*********************************************************  Function perform the stand-In operation *********************************************************************
'Function Name		:					Fn_MyWorkList_SurrogateActions

'Description			 :		 		   This function perform operations on Stand-In Dialog										 
'                                                                                                                                                                                                         
'Parameters			   :	 			sAction: Set / Verify
'													sTaskName : Check Task Name Value
'													sSignoffMember : Check Signoff Member Value
'													sActiveSurrogate : Check Active Surrogate Value
'													sStandORRelease : Stand / Release
'													bCheckOut : Target Check Out
'													sBtnClick : Button Click

'Return Value		   : 			 	True/False

'Examples				:                Fn_MyWorkList_SurrogateActions("Verify", "SubProcess001 (perform-signoffs)", "Engineering/Designer/autotest1", "", "Stand-In", False, "Cancel")
'													Fn_MyWorkList_SurrogateActions("Set", "", "", "", "Stand-In", True, "OK")
'
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Mahendra				19-Nov-2010				1					Script								Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_SurrogateActions(sAction, sTaskName, sSignoffMember, sActiveSurrogate, sStandORRelease, bCheckOut, sBtnClick)
		GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_SurrogateActions"
		Dim objDialog, sMsgValue, sSetValue, bReturn
		
		Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("SurrogateActions")
		
		If objDialog.Exist(1) = False Then
				bReturn = Fn_MenuOperation("Select","Actions:Stand-In")
				Call Fn_ReadyStatusSync(5)
				If bReturn = False Then
					Fn_MyWorkList_SurrogateActions = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Menu [Actions:Stand-In]")
					Set objDialog  = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Menu [Actions:Stand-In]")
				End If
		End If
		
		If objDialog.Exist(1) = True Then
		
			Select Case sAction
			Case "Set"
						If  sStandORRelease <> "" Then
							objDialog.JavaRadioButton("StandIn").SetTOProperty "attached text", sStandORRelease
							objDialog.JavaRadioButton("StandIn").Set "ON"
							If objDialog.JavaRadioButton("StandIn").GetROProperty("value") = "1" Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set Checked to the ["+sStandORRelease+"] Radio Button.")
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Checked to the ["+sStandORRelease+"] Radio Button.")
									Fn_MyWorkList_SurrogateActions = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
							End If
						End If
		
						If  bCheckOut <> "" Then
							If bCheckOut = True Then
								sSetValue = "ON"
								sMsgValue = "Checked"
							ElseIf bCheckOut = False Then
								sSetValue = "OFF"
								sMsgValue = "Un-Checked"
							End If
							objDialog.JavaCheckBox("TransferCheckOut").Set sSetValue
							If objDialog.JavaCheckBox("TransferCheckOut").GetROProperty("value") = "1" Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Set ["+sMsgValue+"] to the Transfer Check-Out CheckBox.")
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set ["+sMsgValue+"] to the Transfer Check-Out CheckBox.")
									Fn_MyWorkList_SurrogateActions = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
							End If
						End If
		
			Case "Verify"
						If Trim(sTaskName) <> "" Then
							objDialog.JavaStaticText("TaskName").SetTOProperty "label", sTaskName
							If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the existence of ["+sTaskName+"] Task Label")
									Fn_MyWorkList_SurrogateActions = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the existence of ["+sTaskName+"] Task Label")	
							End If
						End If
		
						If Trim(sSignoffMember) <> "" Then
							objDialog.JavaStaticText("TaskName").SetTOProperty "label", sSignoffMember
							If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the existence of ["+sSignoffMember+"] Signoff Member Label")
									Fn_MyWorkList_SurrogateActions = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the existence of ["+sSignoffMember+"] Signoff Member Label")	
							End If
						End If
		
						If Trim(sActiveSurrogate) <> "" Then
							objDialog.JavaStaticText("TaskName").SetTOProperty "label", sActiveSurrogate
							If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify the existence of ["+sActiveSurrogate+"] Active Surrogate Label")
									Fn_MyWorkList_SurrogateActions = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified the existence of ["+sActiveSurrogate+"] Active Surrogate Label")	
							End If
						End If
		
						If  sStandORRelease <> "" Then
							objDialog.JavaRadioButton("StandIn").SetTOProperty "attached text", sStandORRelease
							If objDialog.JavaRadioButton("StandIn").GetROProperty("value") = "1" Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified the ["+sStandORRelease+"] Radio Button is marked Checked.")
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify that ["+sStandORRelease+"] Radio Button is marked Checked.")
									Fn_MyWorkList_SurrogateActions = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
							End If
						End If
		
						If  bCheckOut <> "" Then
							If bCheckOut = True Then
								sSetValue = "1"
								sMsgValue = "Checked"
							ElseIf bCheckOut = False Then
								sSetValue = "0"
								sMsgValue = "Un-Checked"
							End If
							If objDialog.JavaCheckBox("TransferCheckOut").GetROProperty("value") = sSetValue Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified that Transfer Check-Out CheckBox is marked ["+sMsgValue+"].")
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify that Transfer Check-Out CheckBox is marked ["+sMsgValue+"].")
									Fn_MyWorkList_SurrogateActions = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
							End If
						End If
		
			End Select
		
			If  sBtnClick <>"" Then
				objDialog.JavaButton(sBtnClick).Click micLeftBtn
				If Err.Number < 0 Then
						Fn_MyWorkList_SurrogateActions = False			
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Click on Button ["+sBtnClick+"]")
						objDialog.JavaButton("Cancel").Click micLeftBtn
						Set objDialog = Nothing
						Exit Function
				Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Clicked on Button ["+sBtnClick+"]")
				End If
			End If
		
		Else
					Fn_MyWorkList_SurrogateActions = False			
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Surrogate Actions dialog does not exist..")
					Set objDialog = Nothing
					Exit Function
		End If
		
		Fn_MyWorkList_SurrogateActions = True
		Set objDialog = Nothing
End Function


'*********************************************************  Function perform the signoff team select operation *********************************************************************
'Function Name		:					Fn_MyWorkList_Signoff_AddressList_Operations

'Description			 :		 		   This function select sign off team addresslists												 
'                                                                                                                                                                                                         
'Parameters			   :	 			sAction: ViewerAdd
'													dicSelectSignOff : Refer DictionaryDeclaration.vbs for the defination & keys included

'Return Value		   : 			 	True/False

'Examples				:                dicSelectSignOff.RemoveAll
'													dicSelectSignOff("WorkListTreeNode") = "My Worklist:AutoTestDBA (autotestdba) Inbox:Tasks to Perform:SubProcess002 (select-signoff-team)"
'													dicSelectSignOff("SignOffTeamSelect") = "#0:#1" ' To be used in case of Address Lists Case
'													dicSelectSignOff("SelectAddress") = "SignOffAddressList"
'													dicSelectSignOff("AddressListsContent") = "autotest1,autotest3"                                                 													
'													dicSelectSignOff("ProcessDescription") = "Process Desc New"
'													dicSelectSignOff("Comments") = "comments for test"
'													dicSelectSignOff("Quorum") = "Numeric:10"
'													dicSelectSignOff("Wait") = true
'													dicSelectSignOff("Adhoc") = true
'
'												Fn_MyWorkList_SignoffTeamSelect("ViewerAdd",dicSelectSignOff)
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Mahendra				19-Nov-2010				1					Script								Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_Signoff_AddressList_Operations(sAction, dicSelectSignOff)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_Signoff_AddressList_Operations"
	On Error Resume Next
	Dim objDialog, dicCount , dicKeys , dicItems , iCounter ,bReturn
	Dim arrQuorum,arrNodeList,arrNode,sExpnadNode
	Dim iNodeCounter, iCount, iCnt

	dicCount  = dicSelectSignOff.Count
	dicItems = dicSelectSignOff.Items
	dicKeys = dicSelectSignOff.Keys
    Select Case sAction
			Case "ViewerAdd"
				Set objDialog = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
				For iCounter = 0 to dicCount - 1
	                    If  dicItems(iCounter) <> "" Then
	                            Select Case dicKeys(iCounter)

								 Case "WorkListTreeNode"	

											'Select the WorkList Tree Node		
											bReturn =  Fn_MyWorkList_TreeNodeOperations("Select",dicItems(iCounter),"")
											If bReturn = false Then
													Fn_MyWorkList_Signoff_AddressList_Operations = False					
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node " + dicItems(iCounter) +  " From WorkList Tree")
													Exit Function
											End If
											Call Fn_ReadyStatusSync(10)
											
											'Set the Viewer Tab
											bReturn =  Fn_MyTc_TabSet("Viewer")	
											If bReturn = false Then
													Fn_MyWorkList_Signoff_AddressList_Operations = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Viewer Tab")
													Exit Function
											End If
											
											Call Fn_ReadyStatusSync(10)
											wait(5)
											 'Set Default View to Task View 
											objDialog.JavaRadioButton("ViewOptions").SetTOProperty "Attached Text","Task View"
											objDialog.JavaRadioButton("ViewOptions").Set "ON"
											If Err.Number < 0 Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Task View from Viewer tab")
													Fn_MyWorkList_Signoff_AddressList_Operations = False									
													Exit Function
											Else																						
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Task View from Viewer tab")
											End If

											 ' Select the Node in Signoff Team tree											 
											 If  InStr(1,dicSelectSignOff.Item("SignOffTeamSelect"),"#0", 1) > 0 Then
												 objDialog.JavaTree("RouteSignOffTeamTree").Select "#0:#1"
														If objDialog.JavaTree("SignOffTeamTree").Exist(1) = True Then
																objDialog.JavaTree("SignOffTeamTree").Select dicSelectSignOff.Item("SignOffTeamSelect")
														Else
																objDialog.JavaTree("RouteSignOffTeamTree").Select dicSelectSignOff.Item("SignOffTeamSelect")
														End If
														If Err.Number < 0 Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Sign off Team tree node [Address Lists]")
																Fn_MyWorkList_SignoffTeamSelect = False									
																Exit Function
														Else																						
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Sign off Team tree node [Address Lists]")
														End If
											 Else
														bReturn =  Fn_MyWorkList_SignoffTeam_TreeNodeOperations("Select","Address Lists")          
														If bReturn = false Then
																Fn_MyWorkList_SignoffTeamSelect = False									
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select "+ sAction +" from Signoff Team Tree")
																Exit Function
														End If
											 End If
											Call Fn_ReadyStatusSync(10)

									Case "SelectAddress"				
											 ' Select the Address
											objDialog.JavaList("AddressLists").Select dicItems(iCounter)
											If bReturn = false Then
													Fn_MyWorkList_Signoff_AddressList_Operations = False									
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ["+dicItems(iCounter)+"]")
													Exit Function
											Else																					
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected ["+dicItems(iCounter)+"]")
											End If
											Call Fn_ReadyStatusSync(10)

								Case "AddressListsContent"
											'Select the Node	
											arrNodeList = split(dicItems(iCounter), ",",-1,1)						
											For iNodeCounter = 0 to UBound(arrNodeList)
													If iNodeCounter = 0 Then
																objDialog.JavaList("AddressListContent").Select arrNodeList(iNodeCounter)
													Else
																objDialog.JavaList("AddressListContent").ExtendSelect arrNodeList(iNodeCounter)
													End If
													If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select [" + arrNodeList(iNodeCounter)   + "] from Address Lists.")
															Fn_MyWorkList_Signoff_AddressList_Operations = False                                                 
															Exit Function				
													Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select [" + arrNodeList(iNodeCounter) + "] from Address Lists." )   
													End If
													wait(1)
											Next
											Call Fn_ReadyStatusSync(10)

											'Click on Add button
											objDialog.JavaButton("Add").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Add User to Project " + dicItems(iCounter)   + " from Organization Tree")
														Fn_MyWorkList_Signoff_AddressList_Operations = False
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Added User to Project " + dicItems(iCounter)   + " from Organization Tree")				
											End If	

											Call Fn_ReadyStatusSync(3)

									Case "Quorum" ' Cases has OR Conflict and wrong code
											If  dicItems(iCounter) <> "" Then
														arrQuorum = split(dicItems(iCounter),":",-1,1) 
														objDialog.JavaRadioButton("ReviewQuorumOption").SetTOProperty "attached text",arrQuorum(0) 
														objDialog.JavaRadioButton("ReviewQuorumOption").Set "ON"
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Review Quorum option " + arrQuorum(0))
																		Fn_MyWorkList_Signoff_AddressList_Operations = False
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Review Quorum option " + arrQuorum(0))
														End if
			
														If arrQuorum(0)  = "Percent" Then
																		objDialog.JavaEdit("ReviewQuorumPercentage").Set arrQuorum(1)
														Else				
																		objDialog.JavaEdit("ReviewQuorumNumeric").Set arrQuorum(1)
														End If
														Call Fn_ReadyStatusSync(3)			
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Review Quorum value " + arrQuorum(1))
																		Fn_MyWorkList_Signoff_AddressList_Operations = False
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Review Quorum value " + arrQuorum(1))
														End If
											End if

									Case "AcknwQuorum" ' Cases has OR Conflict and wrong code
											If  dicItems(iCounter) <> "" Then
														arrQuorum = split(dicItems(iCounter),":",-1,1) 
														objDialog.JavaRadioButton("AcknwQuorumOption").SetTOProperty "attached text",arrQuorum(0) 
														objDialog.JavaRadioButton("AcknwQuorumOption").Set "ON"
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Acknowledge Quorum option " + arrQuorum(0))
																		Fn_MyWorkList_Signoff_AddressList_Operations = False
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Acknowledge Quorum option " + arrQuorum(0))
														End if
			
														If arrQuorum(0)  = "Percent" Then
																		objDialog.JavaEdit("AcknwQuorumPercentage").Set arrQuorum(1)
														Else				
																		objDialog.JavaEdit("AcknwQuorumNumeric").Set arrQuorum(1)
														End If
														Call Fn_ReadyStatusSync(3)			
														If Err.Number < 0 Then
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Review Quorum value " + arrQuorum(1))
																		Fn_MyWorkList_Signoff_AddressList_Operations = False
																		
																		Exit Function
														Else
																		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Review Quorum value " + arrQuorum(1))
														End If
											End if

								Case "Wait"	 
											 If dicItems(iCounter) = true Then
														objDialog.JavaCheckBox("WaitForUndecidedReviewers").Set  "ON"
											Else
														objDialog.JavaCheckBox("WaitForUndecidedReviewers").Set  "OFF"
											End If

											Call Fn_ReadyStatusSync(6)

											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Wait For Undecided Reviewers value to " + dicItems(iCounter))
															Fn_MyWorkList_Signoff_AddressList_Operations = False
															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Wait For Undecided Reviewers value to " + dicItems(iCounter))
											End if

								Case "Adhoc"	 
											 If dicItems(iCounter) = true Then
														objDialog.JavaCheckBox("Ad-hocDone").Set  "ON"
											Else
														objDialog.JavaCheckBox("Ad-hocDone").Set  "OFF"
											End If

											Call Fn_ReadyStatusSync(6)

											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Ad-hoc Done value to " + dicItems(iCounter))
															Fn_MyWorkList_Signoff_AddressList_Operations = False
															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Ad-hoc Done value to " + dicItems(iCounter))
											End if

											Call Fn_ReadyStatusSync(6)
																						
											objDialog.JavaButton("Apply").Click micLeftBtn
											If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Apply Sign off Team Settings")
														Fn_MyWorkList_Signoff_AddressList_Operations = False															
														Exit Function
											Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Applied Sign off Team Settings")
														Fn_MyWorkList_Signoff_AddressList_Operations = True
											End if

											Call Fn_ReadyStatusSync(6)

								Case "ProcessDescription"    											

											objDialog.JavaEdit("ProcDesc").Set dicItems(iCounter)
											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Description value to " + dicItems(iCounter))
															Fn_MyWorkList_Signoff_AddressList_Operations = False															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Description  value to " + dicItems(iCounter))
											End if

											Call Fn_ReadyStatusSync(6)

								Case "Comments" 
											objDialog.JavaEdit("Comments").Set dicItems(iCounter)
											If Err.Number < 0 Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Comment value to " + dicItems(iCounter))
															Fn_MyWorkList_Signoff_AddressList_Operations = False															
															Exit Function
											Else
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Comment value to " + dicItems(iCounter))
											End if
   
								End Select
			End If		
	Next
	
		Case "AvailableListVerify","AvailableListVerifyandAdd"
				Set objDialog=JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("SelectSignoffTeamDialog")
				For iCounter = 0 to dicCount - 1
		                    If  dicItems(iCounter) <> "" Then
		                            Select Case dicKeys(iCounter)
									 Case "AddressListsName"	
									 	objDialog.JavaList("AddressLists").SetTOProperty "attached text", "Address Lists:"
									 	If Fn_SISW_UI_JavaList_Operations("Fn_MyWorkList_Signoff_AddressList_Operations", "Select", objDialog, "AddressLists", dicItems(iCounter), "", "") = False Then
					 						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify address Lists " + dicItems(iCounter))
											Fn_MyWorkList_Signoff_AddressList_Operations = False															
											Exit Function
									 	End If
									 Case "AvailableList"
									 	objDialog.JavaList("AddressLists").SetTOProperty "attached text", "%"
									 	bReturn=objDialog.JavaList("AddressLists").GetROProperty("items count")
									 	arrNodeList = split(dicItems(iCounter), ",",-1,1)
									 		
							 		For iNodeCounter=0 to UBound(arrNodeList)
										For iCount=0 to Cint(bReturn)-1
											If Trim(lcase(objDialog.JavaList("AddressLists").GetItem(iCount))) = Trim(lcase(arrNodeList(iNodeCounter))) Then
												iCnt = iCnt + 1
												Exit For
											ElseIf iCount = Cint(bReturn)-1 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aProjectName(iRowData) &" not found in Available List")
											End If
										Next
									Next
									If iCnt <> UBound(arrNodeList)+1 Then
										Fn_MyWorkList_Signoff_AddressList_Operations = False
										Exit Function
									End If
							End Select 
						End If 
				Next
				If sAction="AvailableListVerifyandAdd" Then
					'click on add button
					bReturn= Fn_SISW_UI_JavaButton_Operations("", "Click",objDialog,"Add")
					If bReturn=False Then
						Exit Function 
					End If
				End If
								
								
	Fn_MyWorkList_Signoff_AddressList_Operations = True
	Set objDialog = Nothing

	End Select
End Function 


'*********************************************************		Function to Get  Schedule Table Column Index 		***********************************************************************

'Function Name		:					Fn_MyWorkList_TableColIndex

'Description			 :		 		  This function is used to get the schdule Table column Index.

'Parameters			   :	 			1.  StrColName:Name of the Col to retrieve Index for.
											
'Return Value		   : 				 Col index/-1

'Pre-requisite			:		 		Stchdule Manager window should be displayed .

'Examples				:				Fn_MyWorkList_TableColIndex("Descision")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari							 28-Jan-2011   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_TableColIndex(ObjTable, sColName)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TableColIndex"
	Dim iCols , iCounter, sColIndex, sName

	On Error Resume Next

	If IsNumeric(sColName) Then
		Fn_MyWorkList_TableColIndex = cint(sColName)
		Exit Function
	End If

	'Verify that  scheduleTable is displayed
	If ObjTable.Exist(5) Then

		'Get the No. of cols present in the schedule Table

		iCols = ObjTable.GetROProperty("cols")
		
		'Get the Col No. of required Column
		For iCounter = 0 to iCols -1
			sName =ObjTable.Object.getColumnName(iCounter)
		  
			If Trim(sName) = Trim(sColName) Then
				Fn_MyWorkList_TableColIndex = iCounter
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyWorkList_TableColIndex:The Column Index for Column [" + sColName +"] is [" + iCounter + "]")	
				Exit For
			End If
		Next
		If Cint(iCounter) = Cint(iCols) Then
			Fn_MyWorkList_TableColIndex = -1
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyWorkList_TableColIndex:The Column [" + sColName + "] dose not exist in schedule  table")	
		End If

	End If
End Function


'*********************************************************		Function to  get  Schedule table Row Index	***********************************************************************

'Function Name		:					Fn_MyWorkList_TableRowIndex

'Description			 :		 		  This function is used to get Schedule table Row Index.

'Parameters			   :	 			1.  sNodeName:Name of the Node to retrieve Index for.
											
'Return Value		   : 				 Node index

'Pre-requisite			:		 		Schedule Manager window should be displayed .

'Examples				:				 Fn_MyWorkList_TableRowIndex("Sch1:Task1:T2")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali							19-May-2010	   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_TableRowIndex(objTreeTable, sNodeName, sColname)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TableRowIndex"
	Dim IntRows ,sNodePath, IntCounter
	Dim bReturn

	On Error Resume Next

	'Verify that PSE BOM Table is displayed
	If objTreeTable.Exist(10) Then

		'Get the No. of rows present in the BOM Table
		IntRows = objTreeTable.GetROProperty("rows")

		bReturn = Fn_MyWorkList_TableColIndex(objTreeTable, sColname)
		If bReturn = -1 Then
			Fn_MyWorkList_TableRowIndex = -1
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyWorkList_TableRowIndex:Failed to Get  Column Index of [" + sColname +"]")
			Exit Function
		End If

       'Get the Row No. of required Node
	   For IntCounter = 0 to cint(IntRows -1)
			sNodePath = objTreeTable.Object.getValueAt(IntCounter,0).toString
			If trim(sNodePath) = trim(sNodeName) Then
				Fn_MyWorkList_TableRowIndex = IntCounter
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyWorkList_TableRowIndex: Row Index of [" + sNodeName +"] Node is [" + IntCounter + "]")	
				Exit for
			End If
	   Next

		If  cint( IntCounter) = cint(IntRows) Then
			Fn_MyWorkList_TableRowIndex = -1
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_MyWorkList_TableRowIndex:Failed to Get  Row Index of [" + sNodeName +"]")	
		End If

  End If
End Function


'*********************************************************		Function to  verify values in Perform Signoff Values	***********************************************************************

'Function Name		:					Fn_MyWorkList_VerifyPerformSignoff

'Description			 :		 		  This function is used to get Schedule table Row Index.

'Parameters			   :	 			1.  sAction:Menu
'												  2. sColumnName : Column name
'												  3.. iRow : Integer value of row	
'												  4. sExpectedValue : Value to verify	
'												  5. sOther : future use
'												  6. bClose : Boolean value if dialog need to be close													
											
'Return Value		   : 				 true/false

'Pre-requisite			:		 		Perform signoff task should be selected

'Examples				:				 Fn_MyWorkList_VerifyPerformSignoff("Menu","User-Group/Role",0,"AutoTestDBA (autotestdba)-dba/DBA","",true)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali				 19-May-2010         	   1.0
'Modified By							Nishigandha 		 1-Feb-2018		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_VerifyPerformSignoff(sAction,sColumnName,iRow,sExpectedValue,dicDetails,bClose)
		GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_VerifyPerformSignoff"
		Dim iRowCount,jCounter,sGetValue,bFlag,sValue,sProperty,sSubAction,iCounter
		Dim aProperty
	 	Select Case sAction
				Case "Menu", "Menu_SignoffDecisionDialog"
							If Not JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter").JavaDialog("Perform Signoff").Exist Then
								bReturn = Fn_MenuOperation("Select", "Actions:Perform")
								If Err.Number < 0 Then
										Fn_MyWorkList_VerifyPerformSignoff = False
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Action --> Perform ") 		
										Exit Function						
								Else
										Fn_MyWorkList_VerifyPerformSignoff = true
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Action --> Perform")
										Wait(5)
										Call Fn_ReadyStatusSync(5)		   	    
								End If
							End If
							
							'Set the parent Object
							If JavaWindow("MyWorkListWindow").Exist Then
									If JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow").JavaDialog("Perform Signoff").Exist Then
										Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyWorkListWindow")
									Else
										Set objParent = JavaWindow("MyWorkListWindow").JavaWindow("MyTeamcenter")	
									End If
							Else
									Set objParent = JavaWindow("WorkflowViewerWindow").JavaWindow("QuickLinks")	
							End If
							
							If sAction = "Menu" Then
									iRowCount = objParent.JavaDialog("Perform Signoff").JavaTable("SignOffTable").GetROProperty("rows")
									If cstr(iRow) = "" Then											
													For jCounter = 0 to iRowCount 
																		sGetValue = objParent.JavaDialog("Perform Signoff").JavaTable("SignOffTable").GetCellData(jCounter,sColumnName)
																		If trim(sGetValue) = trim(sExpectedValue) Then
																				   bFlag = true
																				   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Successfully Verifed  that the value is present [ "+sExpectedValue +" ]")
																				   Exit For      
																		End If																			                                              													 
													Next
													
									Else
													sGetValue = objParent.JavaDialog("Perform Signoff").JavaTable("SignOffTable").GetCellData(cint(iRow),sColumnName)
													If trim(sGetValue) = trim(sExpectedValue) Then
															   bFlag = true
															   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS :  Successfully Verifed  that the value is present [ "+sExpectedValue +" ]")													   
													End If	
									End If
									If bFlag = false Then
										  Fn_MyWorkList_VerifyPerformSignoff = false
										  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL :  Failed to Verify  that the value is present [ "+sExpectedValue +" ]")
										  objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
										  Set objParent = nothing
										  Exit Function						
									Else
										  Fn_MyWorkList_VerifyPerformSignoff = true			
									End If
							' Added new case for TC1140 NEW_DEV : WORKFLOW : NishigandhaJ : 2017120100
							ElseIf sAction = "Menu_SignoffDecisionDialog" Then
									dicCount = dicDetails.Count
									dicItems = dicDetails.Items
									dicKeys = dicDetails.Keys
													
									For iCounter = 0 to dicCount - 1
										
										If Instr(dicKeys(iCounter),"EditBox")>0 Then
											sSubAction = "EditBox"
										ElseIf Instr(dicKeys(iCounter),"RadioButton")>0 Then
											sSubAction = "RadioButton"
										Else
											sSubAction = dicKeys(iCounter)
										End If
										
										sProperty = dicItems(iCounter)
										bFlag = False
									
										Select Case sSubAction
										
											'Verify Edit box existence on form
											Case "EditBox"
	
												If sProperty<>"" Then
													aProperty = Split(sProperty,":")
													If aProperty(0) = "Comments" Then
														If objParent.JavaDialog("Signoff Decision").JavaEdit("Comments").Exist Then
															sAppValue = objParent.JavaDialog("Signoff Decision").JavaEdit("Comments").GetROProperty("value")
															wait 1															
														End If
													ElseIf aProperty(0) = "Password" Then
														'for future
													End If
													
													sValue = aProperty(1)
													If Trim(sAppValue)=Trim(sValue) Then
														bFlag = True
													End If
												End If
												
											'Verify Radio Button existence on form
											Case "RadioButton"
												If sProperty<>"" Then
													aProperty = Split(sProperty,":")
													objParent.JavaDialog("Signoff Decision").JavaRadioButton("DecisionOpt").SetTOProperty "attached text",aProperty(0)
													If objParent.JavaDialog("Signoff Decision").JavaRadioButton("DecisionOpt").Exist Then
														sAppValue = objParent.JavaDialog("Signoff Decision").JavaRadioButton("DecisionOpt").GetROProperty("value")
														wait 1
													End If
													
													sValue = aProperty(1)
													If Trim(sAppValue)=Trim(sValue) Then
														bFlag = True
													End If
												End If													
										End Select
										
									If bFlag = false Then
										  Fn_MyWorkList_VerifyPerformSignoff = false
										  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL :  Failed to Verify  that the value is present [ "+sExpectedValue +" ]")
										  objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
										  Set objParent = nothing
										  Exit Function						
									Else
										  Fn_MyWorkList_VerifyPerformSignoff = true			
									End If
									Next	
							
							End If	
							
							If Cbool(bClose)  Then
									If objParent.JavaDialog("Signoff Decision").Exist Then
										objParent.JavaDialog("Signoff Decision").JavaButton("Cancel").Click micLeftBtn
									End If
									objParent.JavaDialog("Perform Signoff").JavaButton("Close").Click micLeftBtn
									Set objParent = nothing
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Successfully Closed dialog.")
							End If

		End Select
End Function



'*********************************************************		Function to  get Index values of Worklist Tree Nodes	***********************************************************************

'Function Name		:					Fn_MyWorkList_GetItemPathIndex

'Description			 :		 		  This function is used to get Index values of Worklist Tree Nodes

'Parameters			   :	 			1.  objTree:Tree Object
'												  2. sNode : Full NOde Path
'												  3.. sDelimiter : Delimiter if used other than ":"												
											
'Return Value		   : 				 Node Index/false

'Pre-requisite			:		 		Worklist Tree is visible

'Examples				:				 Fn_MyWorkList_GetItemPathIndex(JavaWindow("My Teamcenter - Teamcenter").JavaTree("WorklistTree"), "My Worklist:Engineering/Tc_QALead Inbox:Tasks to Perform:000077/A;1-gg (perform-signoffs):Targets", ":")

'History:
'										Developer Name			Date				Rev. No.			Changes Done																						Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Vallari							13-Mar-2012	   1.0

'										Shrikant					26-Mar-2012      1.1			Added variable sTreeNodeToStr2															  Koustubh
'						        		Shrikant					27-Mar-2012      1.1			Added variable iItemCnt					                    											Koustubh
'						        		Ashok						27-Mar-2012      1.1			Added Condition for Verifying sNode is Empty or Not							Koustubh
'						        		Ashok						29-Mar-2012      1.1			Added  code to handle multiple instances				 					Koustubh
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_GetItemPathIndex(objTree, sNode, sDelimiter)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_GetItemPathIndex"
	Dim iCnt, aNodeArr, iArrCnt, iItemCnt
	Dim objCurrTreeItm, sTreeNodeToStr2
	Dim sItmPath
	If sDelimiter = "" Then
		sDelimiter = ":"
	End If
	Set objItmBounds = nothing
 On Error Resume Next
	Fn_MyWorkList_GetItemPathIndex = False

	aNodeArr = split(sNode, sDelimiter, -1, 1)
	set objCurrTreeItm = objTree.Object.getItem(0)

	'Modified by Ashok kakade
	If sNode <> "" Then
		If trim(objTree.Object.getItem(0).getData().toString()) = trim(aNodeArr(0)) Then
			sItmPath = "#0"
		 Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Capture Root Node  " + aNodeArr(0) + "of MyworkList Tree." )	
			Fn_MyWorkList_GetItemPathIndex = False
			Exit Function
		End If
	Else
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

			If trim(objCurrTreeItm.getItem(iCnt).getData().toString()) = aNodePath(0) OR instr(aNodePath(0), "#") > 0 Then
				If  iInstance = 1 Then
						If instr(aNodePath(0), "#") > 0 Then
							sItmPath = sItmPath + ":" +  aNodePath(0)
						Else
							sItmPath = sItmPath + ":#" +  cstr(iCnt)
						End If
						set objCurrTreeItm = objCurrTreeItm.getItem(iCnt)
						Exit For
				Else
						iInstance = iInstance - 1
				End If
			ElseIf instr(aNodeArr(iArrCnt), ")") > 0  Then
				sTreeNodeToStr2 = ""   	'	Added By Koustubh
				sTreeNodeToStr2 = trim(objCurrTreeItm.getItem(iCnt).getData().getComponent().toString2()) 
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
        If iCnt = iItemCnt Then 	'Modified  condition By Koustubh
			set objCurrTreeItm = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Retrieve Node Index for MyworkList Tree Node [" + aNodeArr(iArrCnt) + "]" )
			Fn_MyWorkList_GetItemPathIndex = False
			Exit Function
		End If
	Next
	Set objItmBounds = objCurrTreeItm.getBounds()
	set objCurrTreeItm = Nothing
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Retrieved Node Index" + sItmPath + "of MyworkList Tree Node [" + sNode + "]" )	
	Fn_MyWorkList_GetItemPathIndex = sItmPath
End Function



'*********************************************************		function is used to perform operation on Claim Perform Signoff dialog	***********************************************************************
'Function Name		:					Fn_MyWorkList_ClaimPerformSignoffOperations

'Description			 :		 		  This function is used to perform operation on Claim Perform Signoff dialog

'Parameters			   :	 			1.  sAction:Action  to be performed
'										2. sInvokeOption : invoke option to open dialog(Menu/ContextMenu)
'										3. sWorkListNode : Hierarchy path of node to be selected from My Worklist Tree 
										'4.dicSelectClaimPerformSignoff	: Dictionary object for Claim Perform Signoff operation
'Return Value		   : 				 True/false

'Pre-requisite			:		 		My Worklist Tree is visible

'Examples				:				 Set dicSelectClaimPerformSignoff = CreateObject("Scripting.Dictionary")
'										dicSelectClaimPerformSignoff("User") = "* - Engineering/Designer"
'										dicSelectClaimPerformSignoff("Button") = "Claim"
'										sNode = "My Worklist:Autotest3 (autotest3) Inbox:Tasks to Perform:000064/A;1-Item (perform-signoffs)"	
'										bReturn = Fn_MyWorkList_ClaimPerformSignoffOperations("ClaimUser", "Menu",sNode, dicSelectClaimPerformSignoff)

'History:
'										Developer Name			Date				Rev. No.			Changes Done																						Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Madhura P			09-Jul-2015	  			 1.0				Created																						Ganesh B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_MyWorkList_ClaimPerformSignoffOperations(sAction, sInvokeOption, sWorkListNode, dicSelectClaimPerformSignoff)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_ClaimPerformSignoffOperations"
	Dim bReturn, sMenu, iRow, sAutoDir, objClaimDialog
	Fn_MyWorkList_ClaimPerformSignoffOperations = false
	Set objClaimDialog = JavaWindow("MyWorkListWindow").JavaWindow("ClaimPerformSignoff")
	If objClaimDialog.Exist(2) = False Then
		If Trim(sWorkListNode) <> "" Then
			bReturn = Fn_MyWorkList_TreeNodeOperations("Select", sWorkListNode,"")
			If bReturn = False Then
				Fn_MyWorkList_ClaimPerformSignoffOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select  [" + sWorkListNode+"] from Worklist" )
				Set objClaimDialog = Nothing
				Exit Function
			Else
				 Call Fn_ReadyStatusSync(5) 
				 Fn_MyWorkList_ClaimPerformSignoffOperations = True
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select  [" + sWorkListNode+"] from Worklist" ) 
			End If
		End If
			
		Select Case lCase(sInvokeOption)
			Case "menu", ""
				sAutoDir = Fn_GetEnvValue("User", "AutomationDir")
				sMenu = Fn_GetXMLNodeValue(sAutoDir + "\TestData\AutomationXML\MenuXML\Workflow_Menu.xml", "ActionsClaimTask")
				bReturn = Fn_MenuOperation("Select", sMenu)
				Call Fn_ReadyStatusSync(3)
				If bReturn = False Then
					Fn_MyWorkList_ClaimPerformSignoffOperations = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Invoke  Menu [" + sMenu+"]." )	
					Set objClaimDialog = Nothing
					Exit Function
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Successfully Invoked  Menu [" + sMenu+"]." )	
				End If	
			Case Else
				'do nothing
		End Select

	End If
	
	If objClaimDialog.Exist(2) Then
		Select Case lcase(sAction)
			Case "claimuser"
				For iRow = 0 To cInt(objClaimDialog.JavaTable("ClaimPerformSignoff").GetROProperty("rows")) -1
					If dicSelectClaimPerformSignoff("User") <> "" Then
						If trim(dicSelectClaimPerformSignoff("User")) = trim(objClaimDialog.JavaTable("ClaimPerformSignoff").GetCellData(iRow,"User")) Then
							bReturn = Fn_UI_JavaTable_SelectRow("Fn_MyWorkList_ClaimPerformSignoffOperations", objClaimDialog, "ClaimPerformSignoff",iRow)
							Fn_MyWorkList_ClaimPerformSignoffOperations = True
							Exit For
						Else
							Fn_MyWorkList_ClaimPerformSignoffOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: To Select Row." )	
						End If
					End If
				Next
								
				If dicSelectClaimPerformSignoff("Button") <> "" Then
					bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_ClaimPerformSignoffOperations","Click",objClaimDialog,dicSelectClaimPerformSignoff("Button"))
					Fn_MyWorkList_ClaimPerformSignoffOperations = bReturn
				Else
					bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_ClaimPerformSignoffOperations","Click",objClaimDialog,"Claim")
					Fn_MyWorkList_ClaimPerformSignoffOperations = bReturn
				End If
				If objClaimDialog.Exist(2) Then
					Call Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_ClaimPerformSignoffOperations","Click",objClaimDialog,"Close")
				End If
			Case else
				Fn_MyWorkList_ClaimPerformSignoffOperations = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL:invalid case" )	
		End Select
	End If
	Set objClaimDialog = Nothing
End Function


'*********************************************************  Function is Subscribe to any Resource Pool *********************************************************************

'Function Name		 :					Fn_MyWorkList_TaskAbort

'Description			:		 		Abort the Do task	

'Parameters			   :	 			1. sTaskName: Task to be selected from worklist tree
'						 				2. sComment: promote comment
'										3. sButton : button to click
'
'Return Value		   : 			 	True/False

'Pre-requisite			:		 	 	Myworklist tab should selected.

'Examples				:				Fn_MyWorkList_TaskAbort("My Worklist:Engineering/Designer Inbox:Tasks to Track:a (perform-signoffs)","Aborted Do task")

'History:s
'										Developer Name					Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shweta Rathod		        	27-Jan-2017	        1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_TaskAbort(sTaskName, sComment,sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TaskAbort"
	Dim sMenu, bReturn, ObjSuspendWin
	
	Fn_MyWorkList_TaskAbort = false	
	Set ObjSuspendWin = Fn_SISW_MyWorkList_GetObject("AbortActionComments")
	If Fn_Setup_GetActivePerspectiveName("")="Workflow Viewer" Then
		Set ObjSuspendWin = Fn_SISW_WorkflowViewer_GetObject("AbortAction")
	End if
		'Select MyWorklist Tree Node
		If Trim(sTaskName) <> "" Then
			bReturn = Fn_MyWorkList_TreeNodeOperations("Select", sTaskName,"")
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select  [" + sTaskName+"] from Worklist" )	
				Exit Function
			Else
				 Call Fn_ReadyStatusSync(1) 
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected  [" + sTaskName+"] from WorkList" )	
			End If
		End If

		'Demote Window Exists or Not	
		If ObjSuspendWin.Exist(1) = False Then
			sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Workflow_Menu"),"ActionAbort")
			bReturn = Fn_MenuOperation("Select", sMenu)
			Call Fn_ReadyStatusSync(5)		   	    
			If bReturn = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke  Menu [" + sMenu+"]." )	
				Exit Function
			End If
		End If

		If ObjSuspendWin.Exist(5) Then
			' eneter comment in comment box
			If Trim(sComment) <> "" Then
				bReturn = Fn_SISW_UI_JavaEdit_Operations("Fn_MyWorkList_TaskAbort", "Set",  ObjSuspendWin, "comment", Trim(sComment))
				If bReturn = false Then
					bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_TaskAbort", "Click", ObjSuspendWin,"Cancel")
					Exit Function
				End If
			End If
			' click on button 
			If sButton <> "" Then
				bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_TaskAbort", "Click", ObjSuspendWin,sButton)
			else
				bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_TaskAbort", "Click", ObjSuspendWin,"OK")
			End if
			
			If bReturn = false Then
				ObjSuspendWin.JavaButton("Cancel").Click micLeftBtn
				Exit Function		
			End If
			Call Fn_ReadyStatusSync(1)
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully aborted Task [" + sTaskName+"]")	
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find  [Abort Action Comments] Dialog" )	
			Exit Function
		End If
		Set ObjSuspendWin = Nothing
		Fn_MyWorkList_TaskAbort = true
End Function


'*********************************************************  Function is to perform various opearations on Task:Pronmote dialog*********************************************************************
'Function Name		:			Fn_MyWorkList_TaskPromoteOperations
'
'Description		:		 	perform various operations on Task:Promote dialog								
'
'Parameters			:	 		1. sTaskName: Task to be selected
'								2. sComment: Promote comments
'								3. sButton : Click on button
'								4. dicTaskPromote : for verification purpose
'
'Return Value		 : 			True/False
'
'Pre-requisite		 :		 	Myworklist tab should selected.
'
'Examples			 :			1. To Verify Fields Exist or not:
'@@    								Set dicTaskPromote = CreateObject( "Scripting.Dictionary" )
'@@    								dicTaskPromote("RadioButton") = "Approve~Reject"
'@@    								bReturn = Fn_MyWorkList_TaskPromoteOperations("VerifyUiObjectsExist","My worklist:Task To perform:0001/A;1-Item (select signoff)","","",dicTaskPromote)
'@@								2. To approve/reject the promote task
'@@									bReturn = Fn_MyWorkList_TaskPromoteOperations("Reject",""," Rejected Action","OK","")
'
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										 Shweta Rathod		   02-Mar-2017	        1.0                                         Shweta Rathod
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_MyWorkList_TaskPromoteOperations(sAction,sTaskName,sComment,sButton,dicTaskPromote)
	Dim sMenu, bReturn, ObjPromoteWin,ObjPromoteWin1,ObjPromoteWin2
	Dim dicCount,dicItems,dicKeys,iCounter,aButton,iCount,sName
	
	Fn_MyWorkList_TaskPromoteOperations = false
	GBL_FAILED_FUNCTION_NAME="Fn_MyWorkList_TaskPromoteOperations"
	
	'Select MyWorklist Tree Node
	If sTaskName <> "" Then
		bReturn = Fn_MyWorkList_TreeNodeOperations("Select",sTaskName,"")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select  " + sTaskName )	
			Exit Function
		End If
	End If
	
	'checking exitence of dialog if not perform menu operation to invoke
	sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Workflow_Menu"),"ActionsPromote")
	set ObjPromoteWin1 = Fn_SISW_MyWorkList_GetObject("PromoteActionComments@2")
	set ObjPromoteWin2 = Fn_SISW_WorkflowViewer_GetObject("Promote Action Comments")
	
	If Fn_SISW_UI_Object_Operations("Fn_MyWorkList_TaskPromoteOperations","Exist", ObjPromoteWin1, SISW_MICRO_TIMEOUT) = false and Fn_SISW_UI_Object_Operations("Fn_MyWorkList_TaskPromoteOperations","Exist", ObjPromoteWin2, SISW_MICRO_TIMEOUT) = false Then
		bReturn = Fn_MenuOperation("Select", sMenu)
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Invoke  Menu " + sMenu )	
			Exit Function
		End If	
		Call Fn_ReadyStatusSync(2)
	End if	
	
	If Fn_SISW_UI_Object_Operations("Fn_MyWorkList_TaskPromoteOperations","Exist", ObjPromoteWin1, SISW_MIN_TIMEOUT) = true then	
		set ObjPromoteWin = Fn_SISW_MyWorkList_GetObject("PromoteActionComments@2")
	Elseif Fn_SISW_UI_Object_Operations("Fn_MyWorkList_TaskPromoteOperations","Exist", ObjPromoteWin2, SISW_MICRO_TIMEOUT) = true Then
		set ObjPromoteWin = Fn_SISW_WorkflowViewer_GetObject("Promote Action Comments")
	else	
		Call Fn_WriteLogFile("Fn_MyWorkList_TaskPromoteOperations", "Failed to Find  [Promote Action Comments] Dialog" )	
		Exit Function
	End If
	
	
	Select Case sAction
		Case "Approve","Reject"                   'Approve or reject on promote dialog
			If sAction = "Approve" then
				ObjPromoteWin.JavaRadioButton("Decision").SetTOProperty "attached text","Approve"
			Elseif sAction = "Reject" then
				ObjPromoteWin.JavaRadioButton("Decision").SetTOProperty "attached text","Reject"
			End if
			
			bReturn = Fn_SISW_UI_JavaRadioButton_Operations("Fn_MyWorkList_TaskPromoteOperations", "Set", ObjPromoteWin, "Decision", "ON")
			If bReturn = false then exit function
			
			If sComment <> "" then
				Err.Clear
				ObjPromoteWin.JavaEdit("Comment").Object.setText sComment
				If Err.Number < 0 Then 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile")&"Fn_MyWorkList_TaskPromoteOperations", "Failed to Set Commnet " + sComment)	
					Exit Function
				End if
			End if
			
		Case "VerifyUiObjectsExist"               ' Verify fields 
			dicCount = dicTaskPromote.Count
			dicItems = dicTaskPromote.Items
			dicKeys = dicTaskPromote.Keys
			For iCounter = 0 To dicCount - 1
				Select Case dicKeys(iCounter)
					Case "RadioButton"
						aButton = Split(dicItems(iCounter),"~")
						For iCount = 0 To UBound(aButton)
							Select Case aButton(iCount)
								Case "Approve"
									sName = "Approve"
								Case "Reject"
									sName = "Reject"
							End Select
			
							ObjPromoteWin.JavaRadioButton("Decision").SetTOProperty "attached text",sName
							bReturn = Fn_SISW_UI_Object_Operations("Fn_MyWorkList_TaskPromoteOperations","Exist", ObjPromoteWin.JavaRadioButton("Decision"),SISW_MICRO_TIMEOUT)
							If bReturn=False Then Exit Function
						Next			
				End Select
			Next		
	End Select
	
	If sButton <> "" then
		bReturn = Fn_SISW_UI_JavaButton_Operations("Fn_MyWorkList_TaskPromoteOperations", "Click",ObjPromoteWin,sButton)
		If bReturn = false then exit function
	End if

	set ObjPromoteWin = nothing
	Set ObjPromoteWin1 = Nothing
	Set ObjPromoteWin1 = nothing
	Fn_MyWorkList_TaskPromoteOperations = true
End Function


