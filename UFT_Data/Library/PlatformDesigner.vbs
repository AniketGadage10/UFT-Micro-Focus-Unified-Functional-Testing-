Option Explicit

' Function List
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'0.  Fn_SISW_PD_GetObject
'1.  Fn_PD_PlatformDesignerTreeNodeOperation
'2.  Fn_PD_ArchitectureBreakdownBasicCreate
'3.  Fn_PD_ArchitectureBreakdownDetailsInfo
'4.  Fn_SISW_PD_TableTabOperations
'5.  Fn_SISW_PD_PasteOperations
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_PD_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_PD_GetObject("Remove")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Jeevan M						8-Nov-2012       1.0
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_PD_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\PlatformDesigner.xml"
	Set Fn_SISW_PD_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_PD_PlatformDesignerTreeNodeOperation

'Description			 :	Function Used to perform operations on [ Platform Designer ] tree

'Parameters			   :   1.StrAction: Action Name
'										2.StrNodeName: Node path
'										3.StrMenu: Popup menu
'
'Return Value		   : 	true/false

'Pre-requisite			:	Should be present in Platform Designer perspective

'Examples				:   bReturn=Fn_PD_PlatformDesignerTreeNodeOperation("Select","Products:000168-Item1","")
'										bReturn=Fn_PD_PlatformDesignerTreeNodeOperation("Exist","Products:000168-Item1","")
'										bReturn=Fn_PD_PlatformDesignerTreeNodeOperation("Expand","Products:000168-Item1","")
'										bReturn=Fn_PD_PlatformDesignerTreeNodeOperation("PopupMenuSelect","Products:000018-Project","Send To:Classification")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												08-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_PD_PlatformDesignerTreeNodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_PD_PlatformDesignerTreeNodeOperation"
 	'Declaring variables
	Dim objPDTree
	Dim iItemCount, aNodePath,  iInstance, instCount, aNodes,aMenuList,bReturn
	Dim sPath, sEle ,iCnt, bFlag,iCounter,objApplet
	Fn_PD_PlatformDesignerTreeNodeOperation=False
	'checking existance of [ Platform Designer ] tree
	If Fn_SISW_UI_Object_Operations("Fn_PD_PlatformDesignerTreeNodeOperation","Exist",JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").JavaTree("PlatformDesignerTree"),SISW_MINLESS_TIMEOUT) = False Then
'	If not JavaWindow("PlatformDesigner").JavaApplet("JApplet").JavaTree("PlatformDesignerTree").Exist(4) then
		JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1
		If Fn_SISW_UI_Object_Operations("Fn_PD_PlatformDesignerTreeNodeOperation","Exist",JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").JavaTree("PlatformDesignerTree"),SISW_MINLESS_TIMEOUT) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Platform Designer Tree not exist")
			JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0
			Exit function
		End If		
	End if
	'Creating object [ Platform Designer ] tree
	Set objPDTree=JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").JavaTree("PlatformDesignerTree")
	Set objApplet=JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet")
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			If instr(1,StrNodeName,"~") Then
				aNodePath = split(StrNodeName, "~",-1, 1)
				StrNodeName = trim(aNodePath(0))
				iInstance = cint(aNodePath(1))
				aNodes = split(StrNodeName,":")
				sPath = ""
				For iCounter = 0 to uBound(aNodes) - 1
					If sPath = "" Then
							sPath = aNodes(iCounter)
					Else
							sPath = sPath & ":" & aNodes(iCounter)
					End If
				Next
				sEle = aNodes( UBound(aNodes) )
				bFlag = False
				iItemCount = cInt(objPDTree.GetROProperty("items count"))
				instCount = 0
				For iCounter = 0 to iItemCount - 1
					If objPDTree.GetItem(iCounter) = sPath then
						For iCnt = 0 to  iItemCount - 1 
								iCounter = iCounter +1
								If  iCounter >=  iItemCount Then
										Exit for
								End If
								If objPDTree.GetItem(iCounter) = ( sPath &":" & sEle ) Then
										instCount = instCount + 1
									If instCount = iInstance  Then
											objPDTree.Select sPath & ":#" & iCnt
											bFlag = True
											Exit for
									End If
								End If
							Next
					End If
				If bFlag Then Exit for
				Next
				Fn_PD_PlatformDesignerTreeNodeOperation=bFlag
			else
				objPDTree.Select StrNodeName
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to select node [ "+StrNodeName+" ] from Platform Designer Tree not exist")
					Fn_PD_PlatformDesignerTreeNodeOperation=False
				else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Successfully selected node [ "+StrNodeName+" ] from Platform Designer Tree not exist")
					Fn_PD_PlatformDesignerTreeNodeOperation=true
				End If
			End If
			' Added for Product Context / Architecture Breakdown
			Call Fn_MenuOperation("Select","View:Refresh")
			wait 1
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Exist"
			bFlag=false
			iItemCount = cInt(objPDTree.GetROProperty("items count"))
			For iCounter=0 to iItemCount-1
				If trim(objPDTree.GetItem(iCounter))=trim(StrNodeName) Then
					bFlag=true
					Exit for
				End If
			Next
			Fn_PD_PlatformDesignerTreeNodeOperation=bFlag
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Expand"
			objPDTree.Expand StrNodeName
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to expand node [ "+StrNodeName+" ] from Platform Designer Tree not exist")
				Fn_PD_PlatformDesignerTreeNodeOperation=False
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : Successfully expanded node [ "+StrNodeName+" ] from Platform Designer Tree not exist")
				Fn_PD_PlatformDesignerTreeNodeOperation=true
			End If

		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PopupMenuSelect"
				bReturn=	Fn_PD_PlatformDesignerTreeNodeOperation("Select",StrNodeName,"")
		
				If bReturn=false Then
						Fn_PD_PlatformDesignerTreeNodeOperation = False
				Else
						Fn_PD_PlatformDesignerTreeNodeOperation = True
				End If
				            
				            If Fn_SISW_UI_Object_Operations("Fn_PD_PlatformDesignerTreeNodeOperation","Exist",objPDTree,SISW_MINLESS_TIMEOUT) = False Then
				            	JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1
				            End If
				            Set objApplet=JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet")
							aMenuList = split(StrMenu, ":",-1,1)
							intCount = Ubound(aMenuList)
							'Open context menu
							Call Fn_UI_JavaTree_OpenContextMenu("Fn_PD_PlatformDesignerTreeNodeOperation",objApplet,"PlatformDesignerTree",StrNodeName)
							
							'Select Menu action
							Select Case intCount
								Case "0"
									 StrMenu = objApplet.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
								Case "1"
									StrMenu =objApplet.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
								Case "2"
									StrMenu = objApplet.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
								Case Else
									Fn_PD_PlatformDesignerTreeNodeOperation = FALSE
									Exit Function
							End Select
		
							objApplet.WinMenu("ContextMenu").Select StrMenu
							If Err.number < 0 Then
								Fn_PD_PlatformDesignerTreeNodeOperation = False
							Else
								Fn_PD_PlatformDesignerTreeNodeOperation = True
							End If
	End Select
	'Releasing object [ Platform Designer ] tree
	JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0
	Set objPDTree=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_PD_ArchitectureBreakdownBasicCreate

'Description			 :	Function Used to create basic Architecture Breakdown

'Parameters			   :   1.StrType: Parameter  Defenation group type
'										2.StrID: Parameter  Defenation group ID
'										3.StrRevision: Parameter  Defenation group Revision
'										4.StrName: Parameter  Defenation group Name
'										5.StrDescription: Parameter  Defenation group Description
'										6.StrGenCompID: Generic Component ID
'										7.StrButtonName: Button Name
'
'Return Value		   : 	Item Id - revision or False

'Pre-requisite			:	Should be log in RAC

'Examples				:   bReturn=Fn_PD_ArchitectureBreakdownBasicCreate("ParmGrpDef","","","Arch1","Desc","","Next")
'										bReturn=Fn_PD_ArchitectureBreakdownBasicCreate("ParmGrpDef","","","Arch2","Desc","","Finish")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												08-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_PD_ArchitectureBreakdownBasicCreate(StrType,StrID,StrRevision,StrName,StrDescription,StrGenCompID,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_PD_ArchitectureBreakdownBasicCreate"
 	'Variable declaration
	Dim objArchDialog
	Dim bFlag,crrID,crrRevision
     StrType = Fn_SISW_MechCurrentobjName(StrType)
	Fn_PD_ArchitectureBreakdownBasicCreate=false
 	'Checking existance of [ New Architecture Breakdown ] dialog
 	If Fn_SISW_UI_Object_Operations("Fn_PD_ArchitectureBreakdownBasicCreate","Exist",Window("PDWindow").JavaDialog("NewArchitectureBreakdown"),SISW_MIN_TIMEOUT) = False AND Fn_SISW_UI_Object_Operations("Fn_PD_ArchitectureBreakdownBasicCreate","Exist",JavaDialog("NewArchitectureBreakdown"),SISW_MIN_TIMEOUT) = False Then
'	If not Window("PDWindow").JavaDialog("NewArchitectureBreakdown").Exist(6) AND  not JavaDialog("NewArchitectureBreakdown").Exist(6)Then
		'calling menu : File=>New=>Architecture Breakdown... : to invoke [ New Architecture Breakdown ] dialog
		bFlag = Fn_MenuOperation("Select","File:New:Architecture Breakdown...")
		Call Fn_ReadyStatusSync(1)
		If bFlag=false Then
			exit function
		End If
	End If
	'creating object of [ New Architecture Breakdown ] dialog
	If Fn_SISW_UI_Object_Operations("Fn_PD_ArchitectureBreakdownBasicCreate","Exist",Window("PDWindow").JavaDialog("NewArchitectureBreakdown"),SISW_MIN_TIMEOUT) = True Then 
'	If  Window("PDWindow").JavaDialog("NewArchitectureBreakdown").Exist(5)Then
		Set objArchDialog=Window("PDWindow").JavaDialog("NewArchitectureBreakdown")
	Else
		Set objArchDialog=JavaDialog("NewArchitectureBreakdown")

	End If
 	
	'selecting Architecture Breakdown Type
	Call Fn_List_Select("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,"ArchitectureBreakdownList",StrType)
	wait 1
    'Wait till  Button is Enabled
	objArchDialog.JavaButton("Next").WaitProperty "enabled", 1, 60000
	'Click on "Next" button
	objArchDialog.JavaButton("Next").Click micLeftBtn
	Call Fn_ReadyStatusSync(1)
	wait 1
	'setting ID
	If StrID<>"" Then
		Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"ID",StrID)
	End If
	'setting Revision
	If StrRevision<>"" Then
		Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"Revision",StrRevision)
	End If
	'clicking on assign button to assign ID and Revision
	If StrID="" or StrRevision="" Then
		Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, "Assign")
		wait 1
	End If
	'retriving ID and Revision
	crrID=Fn_Edit_Box_GetValue("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"ID")
	crrRevision=Fn_Edit_Box_GetValue("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"Revision")
'	setting Name
	If StrName<>"" Then
		Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"Name",StrName)
	End If
	'setting Description
	If StrDescription<>"" Then
		Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"Description",StrDescription)
	End If
	'setting Generic Component ID
	If StrGenCompID<>"" Then
		Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"GenericComponentID",StrGenCompID)
	End If
	Fn_PD_ArchitectureBreakdownBasicCreate=crrID+"-"+crrRevision
	If StrButtonName<>"" Then
        Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, StrButtonName)
        Call Fn_ReadyStatusSync(1)
        wait 1
		If lcase(StrButtonName)="finish" Then
			Call  Fn_ReadyStatusSync(1)
			objArchDialog.Highlight
			Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,"Close")
			Call Fn_ReadyStatusSync(1)
		end if
	End If
	'Releasing object of [ New Architecture Breakdown ] dialog
	Set objArchDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_PD_ArchitectureBreakdownDetailsInfo

'Description			 :	Function Used to add additional information of  Architecture Breakdown

'Parameters			   :   1.StrAction: Action Name
'										2.dicArchitectureBreakdownInfo: Architecture Breakdown additional information
'
'Return Value		   : 	true or False

'Pre-requisite			:	New Architecture Breakdown Dialog should be exist

'Examples				:   IMP Note : Dictionary details should declare in test case
'
'										Dim dicArchitectureBreakdownInfo
'										Set dicArchitectureBreakdownInfo = CreateObject( "Scripting.Dictionary" )
'										dicArchitectureBreakdownInfo("SharedNVEs")="on"
'										dicArchitectureBreakdownInfo("PreventOverlapping")="on"
'										dicArchitectureBreakdownInfo("FullBreakdown")="on"
'										dicArchitectureBreakdownInfo("EnforceHierarchicalVariability")="on"
'										dicArchitectureBreakdownInfo("ButtonName")="Next"
'										bReturn=Fn_PD_ArchitectureBreakdownDetailsInfo("ArchitectureBreakdownProperties",dicArchitectureBreakdownInfo)
'
'										Dim dicArchitectureBreakdownInfo
'										Set dicArchitectureBreakdownInfo = CreateObject( "Scripting.Dictionary" )
'										dicArchitectureBreakdownInfo("Represents")="Parameter Group"
'										dicArchitectureBreakdownInfo("ButtonName")="Next"
'										bReturn= Fn_PD_ArchitectureBreakdownDetailsInfo("AdditionalItemInfo",dicArchitectureBreakdownInfo)
'
'										Dim dicArchitectureBreakdownInfo
'										Set dicArchitectureBreakdownInfo = CreateObject( "Scripting.Dictionary" )
'										dicArchitectureBreakdownInfo("Comment")="Comment1"
'										dicArchitectureBreakdownInfo("ControlEngineer")="Analyst1"
'										dicArchitectureBreakdownInfo("ParameterGroupDescriptor")="Descriptor1"
'										dicArchitectureBreakdownInfo("Specialist")="Analyst1"
'										dicArchitectureBreakdownInfo("ButtonName")="Finish"
'										bReturn=Fn_PD_ArchitectureBreakdownDetailsInfo("AdditionalItemRevisionInfo",dicArchitectureBreakdownInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												08-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_PD_ArchitectureBreakdownDetailsInfo(StrAction,dicArchitectureBreakdownInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_PD_ArchitectureBreakdownDetailsInfo"
 	'Variable declaration
	Dim objArchDialog
	Dim bFlag,objStaticText,objChild
	Dim objTable, iCounter
	Fn_PD_ArchitectureBreakdownDetailsInfo=false
 	'Checking existance of [ New Architecture Breakdown ] dialog
 	If Fn_SISW_UI_Object_Operations("Fn_PD_ArchitectureBreakdownDetailsInfo","Exist",Window("PDWindow").JavaDialog("NewArchitectureBreakdown"),SISW_MIN_TIMEOUT) = False AND Fn_SISW_UI_Object_Operations("Fn_PD_ArchitectureBreakdownDetailsInfo","Exist",JavaDialog("NewArchitectureBreakdown"),SISW_MIN_TIMEOUT) = False Then
'	If not Window("PDWindow").JavaDialog("NewArchitectureBreakdown").Exist(6) AND  not JavaDialog("NewArchitectureBreakdown").Exist(6)Then
		'calling menu : File=>New=>Architecture Breakdown... : to invoke [ New Architecture Breakdown ] dialog
		bFlag = Fn_MenuOperation("Select","File:New:Architecture Breakdown...")
		Call Fn_ReadyStatusSync(1)
		If bFlag=false Then
			exit function
		End If
	End If
	'creating object of [ New Architecture Breakdown ] dialog
	If Fn_SISW_UI_Object_Operations("Fn_PD_ArchitectureBreakdownDetailsInfo","Exist",Window("PDWindow").JavaDialog("NewArchitectureBreakdown"),SISW_MIN_TIMEOUT)= True Then 
'	If  Window("PDWindow").JavaDialog("NewArchitectureBreakdown").Exist(5)Then
		Set objArchDialog=Window("PDWindow").JavaDialog("NewArchitectureBreakdown")
	Else
		Set objArchDialog=JavaDialog("NewArchitectureBreakdown")

	End If
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to click on specific Button
		Case "ClickButton"
			Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,dicArchitectureBreakdownInfo("ButtonName"))
			If Err.Number < 0 Then
				Fn_PD_ArchitectureBreakdownDetailsInfo=False
			Else
				Fn_PD_ArchitectureBreakdownDetailsInfo=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Set Architecture Breakdown Properties		
		Case "ArchitectureBreakdownProperties"
			'setting Shared NVEs option
			If dicArchitectureBreakdownInfo("SharedNVEs")<>"" Then
				Call Fn_CheckBox_Set("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, "SharedNVEs",dicArchitectureBreakdownInfo("SharedNVEs"))
			End If
			'setting Prevent Overlapping NVEs on Breakdown Elements option
			If dicArchitectureBreakdownInfo("PreventOverlapping")<>"" Then
				Call Fn_CheckBox_Set("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, "PreventOverlapping",dicArchitectureBreakdownInfo("PreventOverlapping"))
			End If
			'setting Full Breakdown option
			If dicArchitectureBreakdownInfo("FullBreakdown")<>"" Then
				Call Fn_CheckBox_Set("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, "FullBreakdown",dicArchitectureBreakdownInfo("FullBreakdown"))
			End If
			'setting Enforce Hierarchical Variability option
			If dicArchitectureBreakdownInfo("EnforceHierarchicalVariability")<>"" Then
				Call Fn_CheckBox_Set("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, "EnforceHierarchicalVariability",dicArchitectureBreakdownInfo("EnforceHierarchicalVariability"))
			End If
			'clicking on Button
			If dicArchitectureBreakdownInfo("ButtonName")<>"" Then
				Call  Fn_ReadyStatusSync(1)
				Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,dicArchitectureBreakdownInfo("ButtonName"))
				Call  Fn_ReadyStatusSync(1)
				If lcase(dicArchitectureBreakdownInfo("ButtonName"))="finish" Then
					Call  Fn_ReadyStatusSync(1)
					Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,"Close")
				End If
			End If
			If Err.Number < 0 Then
				Fn_PD_ArchitectureBreakdownDetailsInfo=False
			Else
				Fn_PD_ArchitectureBreakdownDetailsInfo=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Set additional Architecture Breakdown information		
		Case "AdditionalItemInfo"
			'Setting Represents
			If dicArchitectureBreakdownInfo("Represents")<>"" Then
                objArchDialog.JavaStaticText("ParaDefGroup_text").SetTOProperty "label","Represents:"
				Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, "ParaDefGroup_DropDown")
				wait 2

				Set objTable=Description.Create()
				objTable("Class Name").value="JavaTable"
				objTable("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
'				objTable("tagname").value="LOVTreeTable"
				Set objChild=objArchDialog.ChildObjects(objTable)
				For iCounter=0 to objChild(0).GetROProperty("rows")
				
					If trim(dicArchitectureBreakdownInfo("Represents"))=trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
						objChild(0).DoubleClickCell iCounter,0
						Exit for
					End If
				Next
				wait 2
				Set objTable=Nothing
				Set objChild=Nothing

'				Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"Represents",dicArchitectureBreakdownInfo("Represents"))
			
			End If
			'clicking on Button
			If dicArchitectureBreakdownInfo("ButtonName")<>"" Then
				Call  Fn_ReadyStatusSync(1)
				Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,dicArchitectureBreakdownInfo("ButtonName"))
				Call  Fn_ReadyStatusSync(1)
				wait 1
				If lcase(dicArchitectureBreakdownInfo("ButtonName"))="finish" Then
					Call  Fn_ReadyStatusSync(2)
					Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,"Close")
					Call  Fn_ReadyStatusSync(2)
				End If
			End If
			If Err.Number < 0 Then
				Fn_PD_ArchitectureBreakdownDetailsInfo=False
			Else
				Fn_PD_ArchitectureBreakdownDetailsInfo=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Set additional Architecture Breakdown revision information
		Case "AdditionalItemRevisionInfo"
			'setting Comment
			If dicArchitectureBreakdownInfo("Comment")<>"" Then
				Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"Comment",dicArchitectureBreakdownInfo("Comment"))
			End If
			'selecting ControlEngineer
			If dicArchitectureBreakdownInfo("ControlEngineer")<>"" Then
'				objArchDialog.JavaStaticText("ParaDefGroup_text").SetTOProperty "label","Control Engineer:"
'				Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, "ParaDefGroup_DropDown")
'				wait 2
'				Set objStaticText=Description.Create
'				objStaticText("Class Name").value="JavaStaticText"
'				objStaticText("label").value=dicArchitectureBreakdownInfo("ControlEngineer")
'				Set objChild=objArchDialog.ChildObjects(objStaticText)
'				objChild(0).Click 1,1
'				wait 2
'				Set objStaticText=nothing
'				Set objChild=nothing

                Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"ControlEngineer",dicArchitectureBreakdownInfo("ControlEngineer"))
			End If
			'setting Parameter Group Descriptor
			If dicArchitectureBreakdownInfo("ParameterGroupDescriptor")<>"" Then
				Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"ParameterGroupDescriptor",dicArchitectureBreakdownInfo("ParameterGroupDescriptor"))
			End If

			'selecting Specialist
			If dicArchitectureBreakdownInfo("Specialist")<>"" Then
'				objArchDialog.JavaStaticText("ParaDefGroup_text").SetTOProperty "label","Specialist:"
'				Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog, "ParaDefGroup_DropDown")
'				wait 2
'				Set objStaticText=Description.Create
'				objStaticText("Class Name").value="JavaStaticText"
'				objStaticText("label").value=dicArchitectureBreakdownInfo("Specialist")
'				Set objChild=objArchDialog.ChildObjects(objStaticText)
'				objChild(0).Click 1,1
'				wait 2
'				Set objStaticText=nothing
'				Set objChild=nothing

				Call Fn_Edit_Box("Fn_PD_ArchitectureBreakdownBasicCreate",objArchDialog,"Specialist",dicArchitectureBreakdownInfo("Specialist"))

			End If
			'clicking on Button
			If dicArchitectureBreakdownInfo("ButtonName")<>"" Then
				Call  Fn_ReadyStatusSync(1)
				Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,dicArchitectureBreakdownInfo("ButtonName"))
				Call  Fn_ReadyStatusSync(2)
				If lcase(dicArchitectureBreakdownInfo("ButtonName"))="finish" Then
					Call  Fn_ReadyStatusSync(1)
					objArchDialog.highlight
					Call Fn_Button_Click("Fn_PD_ArchitectureBreakdownBasicCreate", objArchDialog,"Close")
					Call  Fn_ReadyStatusSync(2)
				End If
			End If
			If Err.Number < 0 Then
				Fn_PD_ArchitectureBreakdownDetailsInfo=False
			Else
				Fn_PD_ArchitectureBreakdownDetailsInfo=True
			End If
	End Select
	'Releasing object of [ New Architecture Breakdown ] dialog
	Set objArchDialog=nothing
End function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_PD_TableTabOperations

'Description			 :	Function Used to Perform operation on Table tabs appear right side . Eg : Architecture Breakdown , Details

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrTabName: Tab Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	

'Examples				:   bReturn=Fn_SISW_PD_TableTabOperations("Activate","Architecture Breakdown")
'										bReturn=Fn_SISW_PD_TableTabOperations("Activate","Details")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												17-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_PD_TableTabOperations(StrAction,StrTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PD_TableTabOperations"
   Dim objPDtab 
   Fn_SISW_PD_TableTabOperations=False
   Set objPDtab = Fn_SISW_PD_GetObject("JTabbedPane")
   If objPDtab.Exist(6) = False Then
      JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1
  	  Set objPDtab = Fn_SISW_PD_GetObject("JTabbedPane")
   End If
   If objPDtab.Exist(6) Then
	   Select Case StrAction
	 		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - -- - - - - - -- - - - - - - - -
			Case "Activate" 'Case to activate Inner tabs
				objPDtab.Select StrTabName
				If Err.Number < 0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Tab [" + StrTabName + "] not exist")
					Fn_SISW_PD_TableTabOperations=False
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully activated Tab [" + StrTabName + "]")
					Fn_SISW_PD_TableTabOperations=True
				End If
		End Select
   End If
   JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0
   Set objPDtab = Nothing
End function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_PD_PasteOperations

'Description			 :	Function Used to perform operations on [ Paste Operations ] Dialog

'Parameters			   :   1.StrAction: Action Name
'										2.bProcessParents: Process Parents option
'										3.bProcessChildren: Process Children option
'
'Return Value		   : 	true/false

'Pre-requisite			:	Object should be copied

'Examples				:   bReturn=Fn_SISW_PD_PasteOperations("Paste","","on")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												17-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_PD_PasteOperations(StrAction,bProcessParents,bProcessChildren)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_PD_PasteOperations"
 	'Variable declaration
	Dim objPasteOprDialog
	Fn_SISW_PD_PasteOperations=false
	'Checking existance of [ PasteOptions ] dialog
	If Fn_SISW_UI_Object_Operations("Fn_SISW_PD_PasteOperations","Exist",JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").JavaDialog("PasteOptions"),SISW_MIN_TIMEOUT)= False Then 
'	If not JavaWindow("PlatformDesigner").JavaApplet("JApplet").JavaDialog("PasteOptions").Exist(6) Then
		'Calling menu [ Edit = > Paste ]
		JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,1
		If Fn_SISW_UI_Object_Operations("Fn_SISW_PD_PasteOperations","Exist",JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").JavaDialog("PasteOptions"),SISW_MIN_TIMEOUT)= False Then
		   If Fn_MenuOperation("Select", "Edit:Paste") Then
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"successfully called menu [ Edit = > Paste ]")
		   else
			  Exit function
		   End If
		End If
	End If
	'creating object of [ PasteOptions ] dialog
	set objPasteOprDialog=JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").JavaDialog("PasteOptions")
	
	If Fn_SISW_UI_Object_Operations("Fn_SISW_PD_PasteOperations","Exist",objPasteOprDialog,SISW_MIN_TIMEOUT)= False Then
		JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0
		set objPasteOprDialog=JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").JavaDialog("PasteOptions")
	End If	
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Paste"
			'Selecting Process Parents Option
			If bProcessParents<>"" Then
				If lcase(bProcessParents)="on" Then
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_PD_PasteOperations",objPasteOprDialog,"ProcessParents")
				else
					Call Fn_UI_JavaRadioButtont_setOff("Fn_SISW_PD_PasteOperations",objPasteOprDialog,"ProcessParents")
				End If
			End If
			'Selecting Process Children Option
			If bProcessChildren<>"" Then
				If lcase(bProcessChildren)="on" Then
					Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_PD_PasteOperations",objPasteOprDialog,"ProcessChildren")
				else
					Call Fn_UI_JavaRadioButtont_setOff("Fn_SISW_PD_PasteOperations",objPasteOprDialog,"ProcessChildren")
				End If
			End If
			'Clicking on [ OK ] button
			Call Fn_Button_Click("Fn_SISW_PD_PasteOperations", objPasteOprDialog, "OK")	
			wait 3
			Fn_SISW_PD_PasteOperations=true
	End Select
	'Releasing object of [ PasteOptions ] dialog
	JavaWindow("PlatformDesigner").JavaWindow("TcDefaultApplet").SetTOProperty "Index" ,0
	set objPasteOprDialog=nothing
End Function
