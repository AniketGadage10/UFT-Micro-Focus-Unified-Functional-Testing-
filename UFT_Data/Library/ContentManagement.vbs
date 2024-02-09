Option Explicit

' Function List
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'0. Fn_SISW_ContentM_GetObject
'1.  Fn_ContentM_NavTreeNodeOperation
'2.  Fn_ContentM_FolderCreate
'3.  Fn_ContentM_CreatePublicationType
'4.  Fn_ContentM_CreateTopicType
'5.  Fn_ContentM_CreateS1000DDataModule
'6.  Fn_ContentM_CreateS1000DPublicationModule4
'7.  Fn_ContentM_CreateDITAObject
'8.  Fn_ContentM_PublicationStructureTableTabOperations
'9.  Fn_ContentM_TableRowIndex
'10.Fn_ContentM_PublicationStructureTableOperations
'11.Fn_ContentM_CreateSchema
'12.Fn_ContentM_ImportGraphic
'13.Fn_ContentM_CreateS1000DDataDispatchNote4
'14.Fn_ContentM_ImportDocumentsFromFile
'15.Fn_ContentM_SummaryPropertyVerify
'16.Fn_ContentM_TextPadOperations
'17.Fn_ContentM_WindowPreferencesOperations
'18.Fn_ContentM_CreateS1000DDataModuleList4
'19.Fn_ContentM_CreateS1000DDataModule4
'20.Fn_ContentM_CreateS1000DCommentary4
'21.Fn_ContentM_CreateTranslationOrder
'22.Fn_ContentM_LanguageTableGetCellData
'23.Fn_ContentM_LanguageTableOperation
'24.Fn_ContentM_CreateTranslationOffice
'25.Fn_ContentM_CreateXMLSchema
'26.Fn_ContentM_SpecifySearchDetailsAndInvoke
'27.Fn_ContentM_CreateStylesheet
'28.Fn_ContentM_CreateXMLAttributeMapping
'29.Fn_ContentM_CreateStyleType
'30.Fn_ContentM_CreateEditingTool
'31.Fn_ContentM_ExportDocument
'32.Fn_ContentM_Import_DITA_Map
'33.Fn_ContentM_XMLAttributeMapTableEntry
'34.Fn_ContentM_XMLAttributeMapTableGetCellData
'35.Fn_ContentM_XMLAttributeMapTableOperations
'36.Fn_ContentM_CreateProcedure
'37.Fn_ContentM_CreateGraphicAttributeMapping
'38.Fn_ContentM_CreatePublication
'39.Fn_ContentM_PublishContent
'40.Fn_ContentM_PreviewTabOperations
'41.Fn_ContentM_CreateTopic
'42.Fn_ContentM_CreateTransformationPolicy
'43.Fn_SISW_ContentM_SummaryPolicyOperations()
'= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = 
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 			Function to get Object hierarchy  		- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''Function Name		:	Fn_SISW_ContentM_GetObject
'
''Description		  	 :  	Function to get Object hierarchy

''Parameters		   :	1. sObjectName : Object Handle name
								
''Return Value		   :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_ContentM_GetObject("NewAdministrativeClass")

'History:                
'								Developer Name							Date				Rev. No.		Reviewer		Changes Done	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'									Amol	 				06-Mar-2013				1.0					Sandeep N
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_ContentM_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\ContentManagement.xml"
	Set Fn_SISW_ContentM_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_NavTreeNodeOperation

'Description			 :	Function Used to Perform operation on nav tree in Content Management perspective

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrNodeName: tree Node path
'										 3.StrMenu: Popup menu
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:   Fn_ContentM_NavTreeNodeOperation("Exist","Home:Newstuff:Folder2","")
'										Fn_ContentM_NavTreeNodeOperation("Expand","Home:Newstuff:Folder2","")
'										Fn_ContentM_NavTreeNodeOperation("Select","Home:Newstuff:Folder2","")
'										Fn_ContentM_NavTreeNodeOperation("PopupMenuSelect","Home:Newstuff:Folder2","Edit Properties...")
'										Fn_ContentM_NavTreeNodeOperation("DoubleClick","Home:Newstuff:Folder2","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												13-Mar-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_NavTreeNodeOperation(StrAction,StrNodeName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_NavTreeNodeOperation"
    'Variable declaration
	Dim iPath,aMenuList,intCount,aNodePath,oCurrentNode,iCnt
	'Actions to perform different operations 
	Select Case StrAction
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Select"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_ContentM_NavTreeNodeOperation", JavaWindow("ContentManagement").JavaTree("NavTree"), StrNodeName , ":", "~")
			If iPath=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] is not exist in NavTree")
					Fn_ContentM_NavTreeNodeOperation = False
				Else
					JavaWindow("ContentManagement").JavaTree("NavTree").Select iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] from NavTree")
					Fn_ContentM_NavTreeNodeOperation = True
				End If
       '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Expand"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_ContentM_NavTreeNodeOperation", JavaWindow("ContentManagement").JavaTree("NavTree"), StrNodeName , ":", "~")
			If iPath=False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to Expand Node [" + StrNodeName + "] of NavTree")
				  Fn_ContentM_NavTreeNodeOperation = False
			Else
				JavaWindow("ContentManagement").JavaTree("NavTree").Expand iPath
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Expanded Node [" + StrNodeName + "] of NavTree")
				Fn_ContentM_NavTreeNodeOperation = True
			End If

			  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Deselect"	
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_ContentM_NavTreeNodeOperation", JavaWindow("ContentManagement").JavaTree("NavTree"), StrNodeName , ":", "@")
				If iPath=False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to DeSelect Node [" + StrNodeName + "] of NavTree")
					  Fn_ContentM_NavTreeNodeOperation = False
				Else
					JavaWindow("ContentManagement").JavaTree("NavTree").Deselect iPath
					Call Fn_ReadyStatusSync(1)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DeSelected Node [" + StrNodeName + "] of NavTree")
					Fn_ContentM_NavTreeNodeOperation = True
				End If
		 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Multiselect"
				arrNode=Split(StrNodeName,",")
				For iCnt=0 To UBound(arrNode)
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_ContentM_NavTreeNodeOperation", JavaWindow("ContentManagement").JavaTree("NavTree"), arrNode(iCnt) , ":", "~")
					If iPath=False Then
						 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Multi Select Node [" + StrNodeName + "] of NavTree")
						  Fn_ContentM_NavTreeNodeOperation = False
						  Exit Function
					Else
						JavaWindow("ContentManagement").JavaTree("NavTree").ExtendSelect iPath
						Call Fn_ReadyStatusSync(1)
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Multi Selected Node [" + StrNodeName + "] of NavTree")
						Fn_ContentM_NavTreeNodeOperation = True
					End If
				Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "PopupMenuSelect"
			'Build the Popup menu to be selected
			aMenuList = split(StrMenu, ":",-1,1)
			intCount = Ubound(aMenuList)
			'Select node
            iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_ContentM_NavTreeNodeOperation", JavaWindow("ContentManagement").JavaTree("NavTree"), StrNodeName , ":", "~")
			If iPath=False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Select Node [" + StrNodeName + "] of NavTree")
				  Fn_ContentM_NavTreeNodeOperation = False
				  Exit Function
			Else
				JavaWindow("ContentManagement").JavaTree("NavTree").Select iPath
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node [" + StrNodeName + "] of NavTree")
				Fn_ContentM_NavTreeNodeOperation = True
			End If
			'Open context menu
			Call Fn_UI_JavaTree_OpenContextMenu("Fn_ContentM_NavTreeNodeOperation",JavaWindow("ContentManagement"),"NavTree",iPath)
					'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
						Case "1"
							StrMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						Case "2"
							StrMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
						Case Else
							Fn_ContentM_NavTreeNodeOperation = False
							Exit Function
					End Select
					JavaWindow("ContentManagement").WinMenu("ContextMenu").Select StrMenu
					If Err.number < 0 Then
						Fn_ContentM_NavTreeNodeOperation = False
					Else
						Fn_ContentM_NavTreeNodeOperation = True
					End If
      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "DoubleClick"
			iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_ContentM_NavTreeNodeOperation", JavaWindow("ContentManagement").JavaTree("NavTree"), StrNodeName , ":", "~")
			If iPath=False Then
				 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to DoubleClick Node [" + StrNodeName + "] of NavTree")
				  Fn_ContentM_NavTreeNodeOperation = False
			Else
				JavaWindow("ContentManagement").JavaTree("NavTree").Activate iPath
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DoubleClick Node [" + StrNodeName + "] of NavTree")
				Fn_ContentM_NavTreeNodeOperation = True
			End If

	Case "PopupMenuExist"
        			aMenuList = split(StrMenu, ":",-1,1)
					intCount = Ubound(aMenuList)
					Set objJavaTreeNav = JavaWindow("ContentManagement").JavaTree("NavTree")
					'Open context menu
					iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_ContentM_NavTreeNodeOperation", JavaWindow("ContentManagement").JavaTree("NavTree"), StrNodeName , ":", "@")
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_ContentM_NavTreeNodeOperation",JavaWindow("ContentManagement"),"NavTree",iPath)
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
                        Case "1"
							StrMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
                        Case "2"
							StrMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
                        Case Else
							Fn_ContentM_NavTreeNodeOperation = False
                        Exit Function
					End Select
					If JavaWindow("ContentManagement").WinMenu("ContextMenu").GetItemProperty (StrMenu,"Exists") = True Then
						Fn_ContentM_NavTreeNodeOperation = True
					Else
						Fn_ContentM_NavTreeNodeOperation = False
					End If
					Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
		 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Exist"
				iPath = Fn_UI_JavaTreeGetItemPathExt("Fn_ContentM_NavTreeNodeOperation", JavaWindow("ContentManagement").JavaTree("NavTree"), StrNodeName , ":", "~")
				If iPath=False Then
					 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Not exist in NavTree")
					  Fn_ContentM_NavTreeNodeOperation = False
				Else
					aNodePath = split(replace(iPath,"#",""),":")
					Fn_ContentM_NavTreeNodeOperation = true
					Set oCurrentNode = JavaWindow("ContentManagement").JavaTree("NavTree").Object
					For iCnt = 0 to UBound(aNodePath) -1
						Set oCurrentNode = oCurrentNode.GetItem(aNodePath(iCnt))
						If cBool(oCurrentNode.getExpanded()) = False Then
							Fn_ContentM_NavTreeNodeOperation = false
							Exit for
						End If
					Next
					If Fn_ContentM_NavTreeNodeOperation Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Node [" + StrNodeName + "] Exist in NavTree")
					End If
				End If
				Set oCurrentNode = Nothing
			
		Case "GetChildrenList"
				sReturn=""
				Set objCMTcNavTree = JavaWindow("ContentManagement").JavaTree("NavTree")
				If Fn_ContentM_NavTreeNodeOperation("Expand",StrNodeName,"")=True Then
					arrStrNode = Split (StrNodeName, ":")
					If UBound(arrStrNode)=0 And  lCase(arrStrNode(0))="home" Then
							Set oCurrentNode = objCMTcNavTree.Object.getItem(0)
							intNodeCount = oCurrentNode.getItemCount()
							For iCount=0 To intNodeCount-1
								If iCount=0 Then
									sReturn=oCurrentNode.getItem(iCount).getData().toString()
								Else
									sReturn=sReturn+","+oCurrentNode.getItem(iCount).getData().toString()
								End If
							Next
							arr = Split(sReturn,",")
							Fn_ContentM_NavTreeNodeOperation = arr
							Set oCurrentNode=Nothing
							Exit Function
					Else
							Set objCMTcNavTree = JavaWindow("ContentManagement").JavaTree("NavTree")
							Set oCurrentNode = objCMTcNavTree.Object.getItem(0)
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
							Fn_ContentM_NavTreeNodeOperation = arr
							Set oCurrentNode=Nothing
					End If
				Else
					Fn_ContentM_NavTreeNodeOperation = False
				End If
		'- - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -		
		Case Else
				Fn_ContentM_NavTreeNodeOperation = False
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_FolderCreate

'Description			 :	Function Used to create folder in Content Management perspective

'Parameters			   :   '1.StrType: Folder type
'										 2.StrName: Folder name
'										 3.StrDescription: Folder description
'										 4.bOpenOnCreate: Folder open on create option
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:   Fn_ContentM_FolderCreate("Folder","Test","folder for creating objects","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-Mar-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Jan-2013								1.1																						 Sunny R
'1. Modified new folder object hierarchy : 
'Old : JavaWindow("ContentManagement").JavaWindow("TcApplet").JavaDialog("NewFolder")
'New : JavaWindow("ContentManagement").JavaWindow("NewFolder")
'2. Change call to select folder type as per 10.1 design change
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_FolderCreate(StrType,StrName,StrDescription,bOpenOnCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_FolderCreate"
 	'Declaring variables
    Dim objFolderDialog,WshShell,StrMenu
    Dim iItemCount,iCount,strItem,bFlag
	Dim sChar,i
	'Creating object of [ NewFolder ] dialog
	Set objFolderDialog=JavaWindow("ContentManagement").JavaWindow("NewFolder")
	'Checking existance of [ NewFolder ] dialog
	If Not objFolderDialog.Exist(2) Then
	   'Select menu [file -> New -> Folder...	Ctrl+F ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewFolder")
	   Call Fn_MenuOperation("WinMenuSelect",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If

	  If StrType <> "" And StrName <> "" Then
			bFlag=False
			 'Selecting folder type
			iItemCount=Fn_UI_Object_GetROProperty("Fn_ContentM_FolderCreate",objFolderDialog.JavaTree("FolderType"), "items count")
			For iCount=0 To iItemCount-1
				strItem=objFolderDialog.JavaTree("FolderType").GetItem(iCount)
				If Trim(strItem)="Most Recently Used:"+Trim(StrType) Then
					bFlag=True
					Exit For
				ElseIf Trim(strItem)="Complete List" Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Call Fn_JavaTree_Select("Fn_ContentM_FolderCreate", objFolderDialog, "FolderType","Most Recently Used")
				Call Fn_JavaTree_Select("Fn_ContentM_FolderCreate", objFolderDialog, "FolderType","Most Recently Used:"+StrType)
			Else
				Call Fn_UI_JavaTree_Expand("Fn_ContentM_FolderCreate", objFolderDialog, "FolderType","Complete List")
				Call Fn_JavaTree_Select("Fn_ContentM_FolderCreate", objFolderDialog, "FolderType","Complete List")
				Call Fn_JavaTree_Select("Fn_ContentM_FolderCreate", objFolderDialog, "FolderType","Complete List:"+StrType)	
			End If
			objFolderDialog.JavaButton("Next").WaitProperty "enabled", 1, 60000
			Call Fn_Button_Click("Fn_ContentM_FolderCreate", objFolderDialog, "Next")
			'Setting folder name
'			objFolderDialog.JavaEdit("FolderName").Type StrName
            For i = 1 to Len(StrName)
				sChar = mid(StrName, i, 1)
				If Asc(sChar) = 95 Then
					objFolderDialog.JavaEdit("FolderName").PressKey "_", micShift
				Else
					objFolderDialog.JavaEdit("FolderName").Type Chr(Asc(sChar))
				End If
			Next
			'Following piece of loop is written for typing "0" in the EditBox
			If trim(objFolderDialog.JavaEdit("FolderName").GetROProperty("value")) <> TRIM(StrName) Then
				objFolderDialog.JavaEdit("FolderName").Object.setText StrName
			End If
			wait(2)
			'Work-around to invoke event on Name edit box
'			Set WshShell = CreateObject("WScript.Shell")
'			WshShell.SendKeys "a"
'			wait(1)
'			WshShell.SendKeys "{BKSP}"
'			wait(1)
'			Set WshShell = nothing
			'Setting Folder description
			objFolderDialog.JavaEdit("Description").Set StrDescription
			wait(2)
			If Err.Number < 0 Then
				Fn_ContentM_FolderCreate =False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail to Set Folder Name as [" & StrName & "]")
				objFolderDialog.Close
				Set objFolderDialog = Nothing 
				Exit Function
			End If
			'Set open on create option
			If bOpenOnCreate<>"" Then
				Call Fn_CheckBox_Set("Fn_ContentM_FolderCreate", objFolderDialog,"OpenOnCreate",bOpenOnCreate)
			Else
				Call Fn_CheckBox_Set("Fn_ContentM_FolderCreate", objFolderDialog,"OpenOnCreate","OFF")
			End If
			If Cint(objFolderDialog.JavaButton("Finish").GetROProperty("enabled")) <> 1 Then		'Modified comparison Logic By Ketan on 25-May-2011.
				objFolderDialog.JavaButton("Finish").Object.setEnabled True
			End If
			Call Fn_Button_Click("Fn_ContentM_FolderCreate", objFolderDialog, "Finish")
			
			If Err.Number < 0 Then
				Fn_ContentM_FolderCreate = False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail to Click [ Finish ] Button")
				objFolderDialog.Close
				Set objFolderDialog = Nothing 
				Exit Function
			End If
		End If
		Call Fn_ReadyStatusSync(1)
		'clicking on Cancel button
		If objFolderDialog.JavaButton("Cancel").Exist(3) Then
			Call Fn_Button_Click("Fn_ContentM_FolderCreate", objFolderDialog, "Cancel")
		End If
    	Fn_ContentM_FolderCreate = True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully created folder [" &  StrName & "]")
		Set objFolderDialog = Nothing 
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreatePublicationType

'Description			 :	Function Used to create Publication Type

'Parameters			   :   '1.dicPublicationTypeInfo: Publication Type information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicPublicationTypeInfo("Name")="PT123456"
'										dicPublicationTypeInfo("Local Tag Name")="LocalTag1"
'										dicPublicationTypeInfo("System Usage")="User"
'										dicPublicationTypeInfo("Validate Incoming On Parse")="True"
'										dicPublicationTypeInfo("Validate Outgoing On Parse")="False"
'										dicPublicationTypeInfo("Validate Example Content On Parse")="False"
'										dicPublicationTypeInfo("Transfer Mode")="CRF_ECO_Signoff_Details"
'										dicPublicationTypeInfo("File Extension")=".sgm"
'										dicPublicationTypeInfo("Apply Classname")="Publication"
'                       				Fn_ContentM_CreatePublicationType(dicPublicationTypeInfo)
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-Mar-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreatePublicationType(dicPublicationTypeInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreatePublicationType"
 	'Declaring variables
    Dim objCMDialog,WshShell,StrMenu
	Dim bFlag,objTable,objChild,iRow,iCounter,iLastItem

	Fn_ContentM_CreatePublicationType=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking existance of [ NewFolder ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting Publication Type Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreatePublicationType",objCMDialog, "ClassTree","Complete List:Publication Type")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreatePublicationType",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreatePublicationType",objCMDialog)
	wait(3)
	'Set Publication Type Name
	If dicPublicationTypeInfo("Name")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Name:")
		Call Fn_Edit_Box("Fn_ContentM_CreatePublicationType",objCMDialog,"Name", dicPublicationTypeInfo("Name"))
	End If
	'Set Publication Type Local Tag Name
	If dicPublicationTypeInfo("Local Tag Name")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Root Element Name:")
		Call Fn_Edit_Box("Fn_ContentM_CreatePublicationType",objCMDialog,"LocalTagName", dicPublicationTypeInfo("Local Tag Name"))
	End If
	'Set Publication Type System Usage
	If dicPublicationTypeInfo("System Usage")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Usage:")
'		Call Fn_Edit_Box("Fn_ContentM_CreatePublicationType",objCMDialog,"SystemUsage", dicPublicationTypeInfo("System Usage"))
		Call Fn_Button_Click("Fn_ContentM_CreatePublicationType",objCMDialog,"SystemUsage")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicPublicationTypeInfo("System Usage")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set WshShell = Nothing
			Set objCMDialog=Nothing
			Exit Function
		End If
	End If
	'Set Publication Type Validate Incoming On Parse option
	If dicPublicationTypeInfo("Validate Incoming On Parse")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Validate Incoming:")
		objCMDialog.JavaRadioButton("ValidateIncomingOnParse").SetTOProperty "attached text",dicPublicationTypeInfo("Validate Incoming On Parse")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreatePublicationType",objCMDialog, "ValidateIncomingOnParse")
	End If
	'Set Publication Type Validate Outgoing On Parse
	'If dicPublicationTypeInfo("Validate Incoming On Parse")<>"" Then
	If dicPublicationTypeInfo("Validate Outgoing On Parse")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Validate Outgoing:")
		objCMDialog.JavaRadioButton("ValidateOutgoingOnParse").SetTOProperty "attached text",dicPublicationTypeInfo("Validate Outgoing On Parse")
		'objCMDialog.JavaRadioButton("ValidateIncomingOnParse").SetTOProperty "attached text",dicPublicationTypeInfo("Validate Outgoing On Parse")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreatePublicationType",objCMDialog, "ValidateOutgoingOnParse")
	End If
	'Set Publication Type Validate Example Content On Parse
	If dicPublicationTypeInfo("Validate Example Content On Parse")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Validate Example Content:")
		'objCMDialog.JavaRadioButton("ValidateIncomingOnParse").SetTOProperty "attached text",dicPublicationTypeInfo("Validate Example Content On Parse")
		objCMDialog.JavaRadioButton("ValidateExampleContentOnParse").SetTOProperty "attached text",dicPublicationTypeInfo("Validate Example Content On Parse")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreatePublicationType",objCMDialog, "ValidateExampleContentOnParse")
	End If
	'Set Publication Type Transfer Mode
	If dicPublicationTypeInfo("Transfer Mode")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Transfer Mode:")
        Call Fn_Button_Click("Fn_ContentM_CreatePublicationType",objCMDialog,"TransferMode")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 2
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			For iCounter=0 to 2
				iLastItem=objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").GetROProperty("items count")
				iLastItem=iLastItem-1
				objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Select "#"&Cstr(iLastItem)
				wait 1
			Next
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicPublicationTypeInfo("Transfer Mode")
			wait 1
			bFlag=True
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=false
			End If
		Else
			bFlag=false
		End if
        If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell = Nothing
			Exit Function
		End If
	End If
	 'Set Publication Type File Extension
	If dicPublicationTypeInfo("File Extension")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","File Extension:")
'		Call Fn_Edit_Box("Fn_ContentM_CreatePublicationType",objCMDialog,"FileExtension", dicPublicationTypeInfo("File Extension"))
		Call Fn_Button_Click("Fn_ContentM_CreatePublicationType",objCMDialog,"FileExtension")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicPublicationTypeInfo("File Extension")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell = Nothing
			Exit Function
		End If
	End If
	 'Set Publication Type Apply Classname
	If dicPublicationTypeInfo("Apply Classname")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Class Name Applied:")
		Call Fn_Button_Click("Fn_ContentM_CreatePublicationType",objCMDialog,"ApplyClassname")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicPublicationTypeInfo("Apply Classname")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell = Nothing
			Exit Function
		End If
	End If
	 'Set Publication Type Namespace URI
	If dicPublicationTypeInfo("Namespace URI")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Namespace URI:")
		Call Fn_Edit_Box("Fn_ContentM_CreatePublicationType",objCMDialog,"NamespaceURI", dicPublicationTypeInfo("Namespace URI"))
	End If
	 'Set Publication Type Default Namespace Prefix
	If dicPublicationTypeInfo("Default Namespace Prefix")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreatePublicationType",objCMDialog.JavaStaticText("FieldName"),"label","Default Namespace Prefix:")
		Call Fn_Edit_Box("Fn_ContentM_CreatePublicationType",objCMDialog,"DefaultNamespacePrefix", dicPublicationTypeInfo("Default Namespace Prefix"))
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreatePublicationType",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreatePublicationType",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreatePublicationType=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Publication Type of name [" + dicPublicationTypeInfo("Name") + "]")
	Set objCMDialog=Nothing
	Set WshShell = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateTopicType

'Description			 :	Function Used to create Topic Type

'Parameters			   :   '1.dicTopicTypeInfo: Topic Type information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicTopicTypeInfo("Name")="PT123456"
'										dicTopicTypeInfo("Local Tag Name")="LocalTag1"
'										dicTopicTypeInfo("System Usage")="User"
'										dicTopicTypeInfo("Validate Incoming On Parse")="True"
'										dicTopicTypeInfo("Validate Outgoing On Parse")="False"
'										dicTopicTypeInfo("Validate Example Content On Parse")="False"
'										dicTopicTypeInfo("Transfer Mode")="CRF_ECO_Signoff_Details"
'										dicTopicTypeInfo("File Extension")=".sgm"
'										dicTopicTypeInfo("Apply Classname")="Topic"
'                       				Fn_ContentM_CreateTopicType(dicTopicTypeInfo)
'
'
'									dicTopicTypeInfo("TopicType")="Reference Topic Type"
'									dicTopicTypeInfo("Name")="RTP123456"
'									dicTopicTypeInfo("Local Tag Name")="RLocalTag1"
'									dicTopicTypeInfo("System Usage")="User"
'									dicTopicTypeInfo("Validate Incoming On Parse")="True"
'									dicTopicTypeInfo("Validate Outgoing On Parse")="False"
'									dicTopicTypeInfo("Validate Example Content On Parse")="False"
'									dicTopicTypeInfo("Reference Type")="COMPOSABLE_TOPIC_REFERENCE"
'									dicTopicTypeInfo("Fragment Tag Names")="FTN1~FTN2~FTN3"
'									Fn_ContentM_CreateTopicType(dicTopicTypeInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												14-Mar-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateTopicType(dicTopicTypeInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateTopicType"
 	'Declaring variables
    Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter,arrTags,iLastItem

	Fn_ContentM_CreateTopicType=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	'creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking existance of [ NewFolder ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	If dicTopicTypeInfo("TopicType")="" Then
		'Selecting Topic Type Option from list
		Call Fn_JavaTree_Select("Fn_ContentM_CreateTopicType",objCMDialog, "ClassTree","Complete List:Topic Type")
	Else
		'Selecting Topic Type Option from list
		Call Fn_JavaTree_Select("Fn_ContentM_CreateTopicType",objCMDialog, "ClassTree","Complete List:"+dicTopicTypeInfo("TopicType"))
	End If
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateTopicType",objCMDialog)
	wait 3
	'Set Topic Type Name
	If dicTopicTypeInfo("Name")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Name:")
		Call Fn_Edit_Box("Fn_ContentM_CreateTopicType",objCMDialog,"Name", dicTopicTypeInfo("Name"))
	End If
	'Set Topic Type Local Tag Name
	If dicTopicTypeInfo("Local Tag Name")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Root Element Name:")
		Call Fn_Edit_Box("Fn_ContentM_CreateTopicType",objCMDialog,"LocalTagName", dicTopicTypeInfo("Local Tag Name"))
	End If
	'Set Topic Type System Usage
	If dicTopicTypeInfo("System Usage")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Usage:")
'		Call Fn_Edit_Box("Fn_ContentM_CreateTopicType",objCMDialog,"SystemUsage", dicTopicTypeInfo("System Usage"))
		Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"SystemUsage")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicTopicTypeInfo("System Usage")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set WshShell = Nothing
			Set objCMDialog=Nothing
			Exit Function
		End If
	End If
	'Set Topic Type Validate Incoming On Parse option
	If dicTopicTypeInfo("Validate Incoming On Parse")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Validate Incoming:")
		objCMDialog.JavaRadioButton("ValidateIncomingOnParse").SetTOProperty "attached text",dicTopicTypeInfo("Validate Incoming On Parse")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateTopicType",objCMDialog, "ValidateIncomingOnParse")
	End If
	'Set Topic Type Validate Outgoing On Parse
	If dicTopicTypeInfo("Validate Incoming On Parse")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Validate Outgoing:")
		objCMDialog.JavaRadioButton("ValidateIncomingOnParse").SetTOProperty "attached text",dicTopicTypeInfo("Validate Outgoing On Parse")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateTopicType",objCMDialog, "ValidateOutgoingOnParse")
	End If
	'Set Topic Type Validate Example Content On Parse
	If dicTopicTypeInfo("Validate Example Content On Parse")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Validate Example Content:")
		objCMDialog.JavaRadioButton("ValidateIncomingOnParse").SetTOProperty "attached text",dicTopicTypeInfo("Validate Example Content On Parse")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateTopicType",objCMDialog, "ValidateExampleContentOnParse")
	End If
	'Set Topic Type Transfer Mode
	If dicTopicTypeInfo("Transfer Mode")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Transfer Mode:")
		Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"TransferMode")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 2
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			For iCounter=0 to 2
				iLastItem=objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").GetROProperty("items count")
				iLastItem=iLastItem-1
				objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Select "#"&Cstr(iLastItem)
				wait 1
			Next
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicTopicTypeInfo("Transfer Mode")
			wait 1
			bFlag=True
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=false
			End If
		Else
			bFlag=false
		End if
        If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell = Nothing
			Exit Function
		End If
	End If
	 'Set Topic Type File Extension
	If dicTopicTypeInfo("File Extension")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","File Extension:")
'		Call Fn_Edit_Box("Fn_ContentM_CreateTopicType",objCMDialog,"FileExtension", dicTopicTypeInfo("File Extension"))
		Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"FileExtension")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicTopicTypeInfo("File Extension")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell = Nothing
			Exit Function
		End If
	End If
	 'Set Topic Type Apply Classname
	If dicTopicTypeInfo("Apply Classname")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Class Name Applied:")
		Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"ApplyClassname")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicTopicTypeInfo("Apply Classname")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell = Nothing
			Exit Function
		End If
	End If
	 'Set Topic Type Namespace URI
	If dicTopicTypeInfo("Namespace URI")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Namespace URI:")
		Call Fn_Edit_Box("Fn_ContentM_CreateTopicType",objCMDialog,"NamespaceURI", dicTopicTypeInfo("Namespace URI"))
	End If
	 'Set Topic Type Default Namespace Prefix
	If dicTopicTypeInfo("Default Namespace Prefix")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Default Namespace Prefix:")
		Call Fn_Edit_Box("Fn_ContentM_CreateTopicType",objCMDialog,"DefaultNamespacePrefix", dicTopicTypeInfo("Default Namespace Prefix"))
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Adding this code for '' Reference Topic Type "
	 'Set Reference Type
	If dicTopicTypeInfo("Reference Type")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Reference Type:")
		Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicTopicTypeInfo("Reference Type")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell = Nothing
			Exit Function
		End If
	End If
	'Set variant
	If dicTopicTypeInfo("Variant")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTopicType",objCMDialog.JavaStaticText("FieldName"),"label","Variant:")
		Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicTopicTypeInfo("Variant")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell = Nothing
			Exit Function
		End If
	End If
	'Setting Fragment Tag Names
	If dicTopicTypeInfo("Fragment Tag Names")<>"" Then
		arrTags=Split(dicTopicTypeInfo("Fragment Tag Names"),"~")
		Call Fn_CheckBox_Set("Fn_ContentM_CreateTopicType", objCMDialog.JavaApplet("FragmentTagNamesApplet"),"FragmentTagNames","on")
		For iCounter=0 to ubound(arrTags)
            Call Fn_Edit_Box("Fn_ContentM_CreateTopicType",objCMDialog.JavaApplet("FragmentTagNamesApplet"),"FragmentTagNames",arrTags(iCounter))
			Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog.JavaApplet("FragmentTagNamesApplet"),"add_FragmentTagNames")
		Next
		Call Fn_CheckBox_Set("Fn_ContentM_CreateTopicType", objCMDialog.JavaApplet("FragmentTagNamesApplet"),"FragmentTagNames","off")
	End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateTopicType",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateTopicType=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Topic Type of name [" + dicTopicTypeInfo("Name") + "]")
	Set objCMDialog=Nothing
	Set WshShell = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateS1000DDataModule

'Description			 :	Function Used to create S1000D Data Module

'Parameters			   :   '1.dicS1000DDataModuleInfo: S1000D Data Module information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicS1000DDataModuleInfo("TopicType")="Procedural-2-3"
'										dicS1000DDataModuleInfo("Name")="Procedural-2-3_124"
'										dicS1000DDataModuleInfo("MasterLanguageReference")="English US"
'										dicS1000DDataModuleInfo("DocumentTitle")="Procedural-2-3_Title"
'										dicS1000DDataModuleInfo("ModelIdentificationCode")="C3DO"
'										dicS1000DDataModuleInfo("SystemDifferenceCode")="A"
'										dicS1000DDataModuleInfo("ChapterNumber")="61"
'										dicS1000DDataModuleInfo("SectionNumber")="1"
'										dicS1000DDataModuleInfo("Subsection")="1"
'										dicS1000DDataModuleInfo("Subject")="11"
'										dicS1000DDataModuleInfo("DisassemblyCode")="11"
'										dicS1000DDataModuleInfo("DisassemblyCodeVariant")="A"
'										dicS1000DDataModuleInfo("InformationCode")="126"
'										dicS1000DDataModuleInfo("InformationCodeVariant")="A"
'										dicS1000DDataModuleInfo("ItemLocationCode")="B"
'										dicS1000DDataModuleInfo("TechnicalName")="TechName"
'										dicS1000DDataModuleInfo("InformationName")="InfoName"
'										dicS1000DDataModuleInfo("IssueNumber")="111"
'										dicS1000DDataModuleInfo("IssueType")="new"
'										dicS1000DDataModuleInfo("IssueDay")="25"
'										dicS1000DDataModuleInfo("IssueMonth")="11"
'										dicS1000DDataModuleInfo("IssueYear")="2011"
'										dicS1000DDataModuleInfo("SecurityClass")="12"
'										dicS1000DDataModuleInfo("ResponsiblePartnerCompany")="CRTN3"
'										dicS1000DDataModuleInfo("Originator")="CORT3"
'										dicS1000DDataModuleInfo("ApplicabilityOfTheMaterial")="C2D2"
'										dicS1000DDataModuleInfo("QualityAssurance")="tabtop"
'										dicS1000DDataModuleInfo("SystemBreakdownCode")="SBC10"
'										dicS1000DDataModuleInfo("Skill")="sk51"
'										dicS1000DDataModuleInfo("Remarks")="Not good for use"
'										bReturn=Fn_ContentM_CreateS1000DDataModule(dicS1000DDataModuleInfo)
'										
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												27-Mar-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreateS1000DDataModule(dicS1000DDataModuleInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateS1000DDataModule"
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter

	Fn_ContentM_CreateS1000DDataModule=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False
	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting S1000D Data Module Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DDataModule",objCMDialog, "ClassTree","Complete List:S1000D Data Module")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateS1000DDataModule",objCMDialog)
	wait 3
	'Selecting topic type
	If dicS1000DDataModuleInfo("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataModule", objCMDialog, "SelectTopicType",dicS1000DDataModuleInfo("TopicType"))
	End If
	'Setting Revision
	If dicS1000DDataModuleInfo("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Revision", dicS1000DDataModuleInfo("Revision"))
	End If
	'Setting Revision
	If dicS1000DDataModuleInfo("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Name", dicS1000DDataModuleInfo("Name"))
	End If
	'Setting Master Language Reference
	If dicS1000DDataModuleInfo("MasterLanguageReference")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
		bFlag = Fn_List_Select("Fn_ContentM_CreateS1000DDataModule", objCMDialog, "SelectTopicType",dicS1000DDataModuleInfo("MasterLanguageReference"))
		
		
'		
'		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"MasterLanguageReference")
'		wait 1
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicS1000DDataModuleInfo("MasterLanguageReference")
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting Document Title
	If dicS1000DDataModuleInfo("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"DocumentTitle", dicS1000DDataModuleInfo("DocumentTitle"))
	End If
	'Setting Model Identification Code
	If dicS1000DDataModuleInfo("ModelIdentificationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Model Identification Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ModelIdentificationCode", dicS1000DDataModuleInfo("ModelIdentificationCode"))
	End If
	'Setting SystemDifferenceCode
	If dicS1000DDataModuleInfo("SystemDifferenceCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","System Difference Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"SystemDifferenceCode", dicS1000DDataModuleInfo("SystemDifferenceCode"))
	End If
	'Setting Chapter Number
	If dicS1000DDataModuleInfo("ChapterNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Chapter Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ChapterNumber", dicS1000DDataModuleInfo("ChapterNumber"))
	End If
	'Setting Section Number
	If dicS1000DDataModuleInfo("SectionNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Section Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"SectionNumber", dicS1000DDataModuleInfo("SectionNumber"))
	End If
	'Setting Subsection
	If dicS1000DDataModuleInfo("Subsection")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Subsection:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Subsection", dicS1000DDataModuleInfo("Subsection"))
	End If
	'Setting Subsection
	If dicS1000DDataModuleInfo("Subject")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Subject:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Subject", dicS1000DDataModuleInfo("Subject"))
	End If
	'Setting DisassemblyCode
	If dicS1000DDataModuleInfo("DisassemblyCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Disassembly Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"DisassemblyCode", dicS1000DDataModuleInfo("DisassemblyCode"))
	End If
	'Setting DisassemblyCodeVariant
	If dicS1000DDataModuleInfo("DisassemblyCodeVariant")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Disassembly Code Variant:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"DisassemblyCodeVariant", dicS1000DDataModuleInfo("DisassemblyCodeVariant"))
	End If
	'Setting InformationCode
	If dicS1000DDataModuleInfo("InformationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Information Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"InformationCode", dicS1000DDataModuleInfo("InformationCode"))
	End If
	'Setting InformationCodeVariant
	If dicS1000DDataModuleInfo("InformationCodeVariant")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Information Code Variant:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"InformationCodeVariant", dicS1000DDataModuleInfo("InformationCodeVariant"))
	End If
	'Setting ItemLocationCode
	If dicS1000DDataModuleInfo("ItemLocationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Item Location Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ItemLocationCode", dicS1000DDataModuleInfo("ItemLocationCode"))
	End If
	'Setting SupportEquipmentVariantCode
	If dicS1000DDataModuleInfo("SupportEquipmentVariantCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Support Equipment Variant Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"SupportEquipmentVariantCode", dicS1000DDataModuleInfo("SupportEquipmentVariantCode"))
	End If
	'Setting EquipmentCategoryAndSub-CategoryCode
	If dicS1000DDataModuleInfo("EquipmentCategoryAndSubCategoryCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Equipment Category and Sub-Category Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"EquipmentCategoryAndSub-CategoryCode", dicS1000DDataModuleInfo("EquipmentCategoryAndSubCategoryCode"))
	End If
	'Setting EquipmentIdentificationCode
	If dicS1000DDataModuleInfo("EquipmentIdentificationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Equipment Identification Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"EquipmentIdentificationCode", dicS1000DDataModuleInfo("EquipmentIdentificationCode"))
	End If
	'Setting ComponentIdentificationCode
	If dicS1000DDataModuleInfo("ComponentIdentificationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Component Identification Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ComponentIdentificationCode", dicS1000DDataModuleInfo("ComponentIdentificationCode"))
	End If
	'Setting ExtensionCode
	If dicS1000DDataModuleInfo("ExtensionCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Extension Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ExtensionCode", dicS1000DDataModuleInfo("ExtensionCode"))
	End If
	'Setting ExtensionProducer
	If dicS1000DDataModuleInfo("ExtensionProducer")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Extension Producer:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ExtensionProducer", dicS1000DDataModuleInfo("ExtensionProducer"))
	End If
	'Setting ExportFileName
	If dicS1000DDataModuleInfo("ExportFileName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Export File Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ExportFileName", dicS1000DDataModuleInfo("ExportFileName"))
	End If
	'Set Is This a Template option
	If dicS1000DDataModuleInfo("IsThisATemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicS1000DDataModuleInfo("IsThisATemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataModule",objCMDialog, "IsThisATemplate")
	End If
	'Set Reference Only option
	If dicS1000DDataModuleInfo("ReferenceOnly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference Only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicS1000DDataModuleInfo("ReferenceOnly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataModule",objCMDialog, "ReferenceOnly")
	End If
	'Setting TechnicalName
	If dicS1000DDataModuleInfo("TechnicalName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Technical Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"TechnicalName", dicS1000DDataModuleInfo("TechnicalName"))
	End If
	'Setting InformationName
	If dicS1000DDataModuleInfo("InformationName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Information Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"InformationName", dicS1000DDataModuleInfo("InformationName"))
	End If
	'Setting IssueNumber
	If dicS1000DDataModuleInfo("IssueNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"IssueNumber", dicS1000DDataModuleInfo("IssueNumber"))
	End If
	'Setting IssueType
	If dicS1000DDataModuleInfo("IssueType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"IssueType", dicS1000DDataModuleInfo("IssueType"))
	End If
	'Setting IssueDay
	If dicS1000DDataModuleInfo("IssueDay")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Day:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"IssueDay", dicS1000DDataModuleInfo("IssueDay"))
	End If
	'Setting IssueMonth
	If dicS1000DDataModuleInfo("IssueMonth")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Month:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"IssueMonth", dicS1000DDataModuleInfo("IssueMonth"))
	End If
	'Setting IssueYear
	If dicS1000DDataModuleInfo("IssueYear")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Year:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"IssueYear", dicS1000DDataModuleInfo("IssueYear"))
	End If
	'Setting SecurityClass
	If dicS1000DDataModuleInfo("SecurityClass")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Security Class:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"SecurityClass", dicS1000DDataModuleInfo("SecurityClass"))
	End If
	'Setting ResponsiblePartnerCompany
	If dicS1000DDataModuleInfo("ResponsiblePartnerCompany")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Responsible Partner Company:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ResponsiblePartnerCompany", dicS1000DDataModuleInfo("ResponsiblePartnerCompany"))
	End If
	'Setting Originator
	If dicS1000DDataModuleInfo("Originator")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Originator", dicS1000DDataModuleInfo("Originator"))
	End If
	'Setting ApplicabilityOfTheMaterial
	If dicS1000DDataModuleInfo("ApplicabilityOfTheMaterial")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Applicability of the Material:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ApplicabilityOfTheMaterial", dicS1000DDataModuleInfo("ApplicabilityOfTheMaterial"))
	End If
	'Setting ApplicabilityType
	If dicS1000DDataModuleInfo("ApplicabilityType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Applicability Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ApplicabilityType", dicS1000DDataModuleInfo("ApplicabilityType"))
	End If
	'Setting QualityAssurance
	If dicS1000DDataModuleInfo("QualityAssurance")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Quality Assurance:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"QualityAssurance", dicS1000DDataModuleInfo("QualityAssurance"))
	End If
	'Setting SystemBreakdownCode
	If dicS1000DDataModuleInfo("SystemBreakdownCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","System Breakdown Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"SystemBreakdownCode", dicS1000DDataModuleInfo("SystemBreakdownCode"))
	End If
	'Setting Skill
	If dicS1000DDataModuleInfo("Skill")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Skill:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Skill", dicS1000DDataModuleInfo("Skill"))
	End If
	'Setting ReasonForUpdate
	If dicS1000DDataModuleInfo("ReasonForUpdate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reason for Update:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"ReasonForUpdate", dicS1000DDataModuleInfo("ReasonForUpdate"))
	End If
	'Setting Remarks
	If dicS1000DDataModuleInfo("Remarks")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Remarks:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Remarks", dicS1000DDataModuleInfo("Remarks"))
	End If
	'Setting Level
	If dicS1000DDataModuleInfo("Level")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Level:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Remarks", dicS1000DDataModuleInfo("Level"))
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateS1000DDataModule=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateS1000DPublicationModule4

'Description			 :	Function Used to create S1000D Publication Module 4.0

'Parameters			   :   '1.dicS1000DPublicationModule4Info: S1000D Publication Module 4.0 information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicS1000DPublicationModule4Info("TopicType")="PM-4-0"
'										dicS1000DPublicationModule4Info("Name")="PM-4-0_136"
'										dicS1000DPublicationModule4Info("MasterLanguageReference")="English US"
'										dicS1000DPublicationModule4Info("DocumentTitle")="PM-4-0_Title"
'										dicS1000DPublicationModule4Info("ModelIdentificationCode")="C3DO"
'										dicS1000DPublicationModule4Info("IssuingAuthority")="CRTN3_456"
'										dicS1000DPublicationModule4Info("Number")="00001"
'										dicS1000DPublicationModule4Info("VolumeOfPublication")="00"
'										dicS1000DPublicationModule4Info("IssueNumber")="000"
'										dicS1000DPublicationModule4Info("IssueType")="New"
'										dicS1000DPublicationModule4Info("IssueDay")="30"
'										dicS1000DPublicationModule4Info("IssueMonth")="11"
'										dicS1000DPublicationModule4Info("IssueYear")="2010"
'										dicS1000DPublicationModule4Info("SecurityClass")="01"
'										dicS1000DPublicationModule4Info("ResponsiblePartnerCompany")="CRTN3"
'										dicS1000DPublicationModule4Info("Originator")="CORT3"
'										dicS1000DPublicationModule4Info("QualityAssurance")="PV"
'										dicS1000DPublicationModule4Info("SystemBreakdownCode")="SBC10"
'										dicS1000DPublicationModule4Info("Effectivity")="01"
'										dicS1000DPublicationModule4Info("Media")="M1"
'										dicS1000DPublicationModule4Info("MediaType")="M1"
'										dicS1000DPublicationModule4Info("MediaCode")="C1"
'										dicS1000DPublicationModule4Info("InWorkNumber")="00"
'										bReturn=Fn_ContentM_CreateS1000DPublicationModule4(dicS1000DPublicationModule4Info)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												27-Mar-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Shweta Rathod											08-May-2015								1.2					Modified function as per 11.2 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreateS1000DPublicationModule4(dicS1000DPublicationModule4Info)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateS1000DPublicationModule4"
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateS1000DPublicationModule4=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False
	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	If dicS1000DPublicationModule4Info("AuthorClass")="" Then
		'Selecting S1000D Publication Module 4.0 Option from list
		Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog, "ClassTree","Complete List:S1000D Publication Module 4.0/4.1/4.2")
	Else
		Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog, "ClassTree","Complete List:"+dicS1000DPublicationModule4Info("AuthorClass"))
	End If
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog)
	wait 3
	'Selecting topic type
	If dicS1000DPublicationModule4Info("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DPublicationModule4", objCMDialog, "SelectTopicType",dicS1000DPublicationModule4Info("TopicType"))
	End If
	'Setting Revision
	If dicS1000DPublicationModule4Info("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Revision", dicS1000DPublicationModule4Info("Revision"))
	End If
	'Setting Revision
	If dicS1000DPublicationModule4Info("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Name", dicS1000DPublicationModule4Info("Name"))
	End If
	'Setting Master Language Reference
	If dicS1000DPublicationModule4Info("MasterLanguageReference")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataDispatchNote4", objCMDialog, "SelectTopicType",dicS1000DDataDispatchNote4Info("MasterLanguageReference"))
'		Call Fn_Button_Click("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"MasterLanguageReference")
'		wait 1
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicS1000DPublicationModule4Info("MasterLanguageReference")
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
'		If bFlag=False Then
'			Set objCMDialog=Nothing
'			Set WshShell=Nothing
'			Exit Function
'		End If
	End If
	'Setting Document Title
	If dicS1000DPublicationModule4Info("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"DocumentTitle", dicS1000DPublicationModule4Info("DocumentTitle"))
	End If
	'Setting ModelIdentificationCode
	If dicS1000DPublicationModule4Info("ModelIdentificationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Model Identification Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"ModelIdentificationCode", dicS1000DPublicationModule4Info("ModelIdentificationCode"))
	End If
	'Setting IssuingAuthority
	If dicS1000DPublicationModule4Info("IssuingAuthority")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issuing Authority:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"IssuingAuthority", dicS1000DPublicationModule4Info("IssuingAuthority"))
	End If
	'Setting Number
	If dicS1000DPublicationModule4Info("Number")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Number", dicS1000DPublicationModule4Info("Number"))
	End If
	'Setting Volume of the Publication
	If dicS1000DPublicationModule4Info("VolumeOfPublication")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Volume of the Publication:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"VolumeOfThePublication", dicS1000DPublicationModule4Info("VolumeOfPublication"))
	End If
	'Setting ExtensionCode
	If dicS1000DPublicationModule4Info("ExtensionCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Extension Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"ExtensionCode", dicS1000DPublicationModule4Info("ExtensionCode"))
	End If
	'Setting ExtensionProducer
	If dicS1000DPublicationModule4Info("ExtensionProducer")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Extension Producer:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"ExtensionProducer", dicS1000DPublicationModule4Info("ExtensionProducer"))
	End If
	'Setting IssueNumber
	If dicS1000DPublicationModule4Info("IssueNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"IssueNumber", dicS1000DPublicationModule4Info("IssueNumber"))
	End If
	'Setting IssueType
	If dicS1000DPublicationModule4Info("IssueType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"IssueType", dicS1000DPublicationModule4Info("IssueType"))
	End If
	'Setting IssueDay
	If dicS1000DPublicationModule4Info("IssueDay")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Day:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"IssueDay", dicS1000DPublicationModule4Info("IssueDay"))
	End If
	'Setting IssueMonth
	If dicS1000DPublicationModule4Info("IssueMonth")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Month:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"IssueMonth", dicS1000DPublicationModule4Info("IssueMonth"))
	End If
	'Setting IssueYear
	If dicS1000DPublicationModule4Info("IssueYear")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Year:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"IssueYear", dicS1000DPublicationModule4Info("IssueYear"))
	End If
	'Setting SecurityClass
	If dicS1000DPublicationModule4Info("SecurityClass")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Security Class:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"SecurityClass", dicS1000DPublicationModule4Info("SecurityClass"))
	End If
	'Setting ResponsiblePartnerCompany
	If dicS1000DPublicationModule4Info("ResponsiblePartnerCompany")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Responsible Partner Company Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"ResponsiblePartnerCompany", dicS1000DPublicationModule4Info("ResponsiblePartnerCompany"))
	End If
	'Setting Originator
	If dicS1000DPublicationModule4Info("Originator")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Originator", dicS1000DPublicationModule4Info("Originator"))
	End If
	'Setting QualityAssurance
	If dicS1000DPublicationModule4Info("QualityAssurance")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Quality Assurance:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"QualityAssurance", dicS1000DPublicationModule4Info("QualityAssurance"))
	End If
	'Setting SystemBreakdownCode
	If dicS1000DPublicationModule4Info("SystemBreakdownCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","System Breakdown Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"SystemBreakdownCode", dicS1000DPublicationModule4Info("SystemBreakdownCode"))
	End If
	'Setting Remarks
	If dicS1000DPublicationModule4Info("Remarks")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Remarks:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Remarks", dicS1000DPublicationModule4Info("Remarks"))
	End If
    'Setting Effectivity
	If dicS1000DPublicationModule4Info("Effectivity")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Effectivity:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Effectivity", dicS1000DPublicationModule4Info("Effectivity"))
	End If
	'Setting Media
	If dicS1000DPublicationModule4Info("Media")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Media:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Media", dicS1000DPublicationModule4Info("Media"))
	End If
	'Setting MediaType
	If dicS1000DPublicationModule4Info("MediaType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Media Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"MediaType", dicS1000DPublicationModule4Info("MediaType"))
	End If
	'Setting MediaCode
	If dicS1000DPublicationModule4Info("MediaCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Media Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"MediaCode", dicS1000DPublicationModule4Info("MediaCode"))
	End If
	'Setting FunctionalItemCode
	If dicS1000DPublicationModule4Info("FunctionalItemCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Functional Item Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"FunctionalItemCode", dicS1000DPublicationModule4Info("FunctionalItemCode"))
	End If
	'Setting ReasonForUpdate
	If dicS1000DPublicationModule4Info("ReasonForUpdate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reason for Update:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"ReasonForUpdate", dicS1000DPublicationModule4Info("ReasonForUpdate"))
	End If
	'Setting InWorkNumber
	If dicS1000DPublicationModule4Info("InWorkNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","In Work Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"InWorkNumber", dicS1000DPublicationModule4Info("InWorkNumber"))
	End If
	'Setting ExportFileName
	If dicS1000DPublicationModule4Info("ExportFileName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Export File Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"ExportFileName", dicS1000DPublicationModule4Info("ExportFileName"))
	End If
	'Set Is This a Template option
	If dicS1000DPublicationModule4Info("IsThisATemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicS1000DPublicationModule4Info("IsThisATemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog, "IsThisATemplate")
	End If
	'Set Reference Only option
	If dicS1000DPublicationModule4Info("ReferenceOnly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference Only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicS1000DPublicationModule4Info("ReferenceOnly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog, "ReferenceOnly")
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateS1000DPublicationModule4=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateDITAObject

'Description			 :	Function Used to create DITA Objects : Eg:- DITA Base , DITA Concept Topic, DITA Dynamic Map,DITA Reference Topic, DITA Static Map,DITA Task Topic,DITA Topic

'Parameters			   :   '1.StrDITAType: DITA Type
'										 2.dicDITAObjectInfo : DITA Object information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicDITAObjectInfo("TopicType")="DITA Composite"
'										dicDITAObjectInfo("ID")="X000086"
'										dicDITAObjectInfo("Revision")="A"
'										dicDITAObjectInfo("Name")="DITA Base1"
'										dicDITAObjectInfo("DocumentTitle")="Base 1"
'										dicDITAObjectInfo("MasterLanguageReference")="English US"
'										bReturn=Fn_ContentM_CreateDITAObject("DITA Base",dicDITAObjectInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												27-Mar-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreateDITAObject(StrDITAType,dicDITAObjectInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateDITAObject"
	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateDITAObject=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False
	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting DITA Object Type
	Call Fn_JavaTree_Select("Fn_ContentM_CreateDITAObject",objCMDialog, "ClassTree","Complete List:"+StrDITAType)
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateDITAObject",objCMDialog)
	wait 3
	'Selecting topic type
	If dicDITAObjectInfo("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateDITAObject", objCMDialog, "SelectTopicType",dicDITAObjectInfo("TopicType"))
	End If
	'Setting ID
	If dicDITAObjectInfo("ID")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateDITAObject",objCMDialog,"ID",dicDITAObjectInfo("ID"))
	End If
	'Setting Revision
	If dicDITAObjectInfo("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateDITAObject",objCMDialog,"ID",dicDITAObjectInfo("Revision"))
	End If
	'Setting Name
	If dicDITAObjectInfo("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateDITAObject",objCMDialog,"Name",dicDITAObjectInfo("Name"))
	End If
	'Setting Document Title
	If dicDITAObjectInfo("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateDITAObject",objCMDialog,"DocumentTitle",dicDITAObjectInfo("DocumentTitle"))
	End If
	'Setting Master Language Reference
	If dicDITAObjectInfo("MasterLanguageReference")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
		'Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"MasterLanguageReference")
		wait 1,500
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicDITAObjectInfo("MasterLanguageReference")
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
		Call Fn_SISW_UI_JavaList_Operations("Fn_ContentM_CreateDITAObject", "Select", objCMDialog, "JavaList", dicDITAObjectInfo("MasterLanguageReference"), "", "")
		wait 1,500
		If Err.number < 0 Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Set Is This a Template option
	If dicDITAObjectInfo("IsThisATemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicDITAObjectInfo("IsThisATemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateDITAObject",objCMDialog, "IsThisATemplate")
	End If
	'Set Reference Only option
	If dicDITAObjectInfo("ReferenceOnly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicDITAObjectInfo("ReferenceOnly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateDITAObject",objCMDialog, "ReferenceOnly")
	End If
	'Setting DITA Audience
	If dicDITAObjectInfo("DITAAudience")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","DITA Audience:"
		Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicDITAObjectInfo("DITAAudience")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting DITA Importance
	If dicDITAObjectInfo("DITAImportance")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","DITA Importance:"
		Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicDITAObjectInfo("DITAImportance")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting DITA Other Properties
	If dicDITAObjectInfo("DITAOtherProperties")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","DITA Other Properties:"
		Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicDITAObjectInfo("DITAOtherProperties")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting DITA Platform
	If dicDITAObjectInfo("DITAPlatform")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","DITA Platform:"
		Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicDITAObjectInfo("DITAPlatform")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting DITA Product
	If dicDITAObjectInfo("DITAProduct")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","DITA Product:"
		Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicDITAObjectInfo("DITAProduct")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Relation Selection:"
	If objCMDialog.JavaList("RelationSelection").Exist(2) Then
		If dicDITAObjectInfo("RelationSelection")<>"" Then
			objCMDialog.JavaList("RelationSelection").Select dicDITAObjectInfo("RelationSelection")
		Else
			objCMDialog.JavaList("RelationSelection").Select "DC_ComposableReferenceR"
		End If
	End If


	objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference Topic Type:"
	If objCMDialog.JavaList("ReferenceTopicType").Exist(2) Then
		If dicDITAObjectInfo("ReferenceTopicType")<>"" Then
			'objCMDialog.JavaList("ReferenceTopicType").Select dicDITAObjectInfo("ReferenceTopicType")
				Call Fn_List_Select("Fn_ContentM_CreateDITAObject", objCMDialog, "ReferenceTopicType",dicDITAObjectInfo("ReferenceTopicType"))
		Else
			objCMDialog.JavaList("ReferenceTopicType").Select "DITA Topicref Dyn Map to Dyn Map"
		End If
	End If

	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateDITAObject",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateDITAObject=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_PublicationStructureTableTabOperations

'Description			 :	Function Used to perform operations on Publication Structure Table Tabs

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrTabName: Tab name
'										 3.StrMenu: RMB Menu
'
'Return Value		   : 	True or False

'Pre-requisite			:	PublicationStructureTable Should activated

'Examples				:   Fn_ContentM_PublicationStructureTableTabOperations("Activate","C4DO-CRTN3-00001-00-PM-4-0_2","")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												28-Mar-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_PublicationStructureTableTabOperations(StrAction,StrTabName,StrMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_PublicationStructureTableTabOperations"
 	'Variable declaration
	Dim objTab
	'Function returns false
	Fn_ContentM_PublicationStructureTableTabOperations=False
	'Creating object of [ PublicationStructuresTab ] 
	Set objTab=JavaWindow("ContentManagement").JavaWindow("JApplet").JavaStaticText("PublicationStructuresTab")
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to activate Tab
		Case "Activate"
			objTab.SetTOProperty "label","("+StrTabName+")"
			If objTab.Exist(6) Then
				objTab.Click 1,1,"LEFT"
				Fn_ContentM_PublicationStructureTableTabOperations=True
			End If
	End Select
	'Releasing object of [ PublicationStructuresTab ] 
	Set objTab=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_TableRowIndex

'Description			 :	Function Used to retrive row number from table

'Parameters			   :   '1.objTable: Table Object
'										 2.sNodeName: Node name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Table Should activated

'Examples				:   Fn_ContentM_TableRowIndex(JavaWindow("ContentManagement").JavaApplet("JApplet").JavaTable("PublicationStructures"), "X000001/A;1-Base1 (View):X000003/A;1-Topic1")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												28-Mar-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_TableRowIndex(objTable, sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_TableRowIndex"
	Dim nodeArr, aRowNode, iColIndex, aPath
	Dim iRowCounter, sNode, iInstance, iNodeCounter, iPathCounter, bFound 
	Dim iRows, sNodePath, sPath, StrNodePath
	sPath = ""

	If Fn_UI_ObjectExist("Fn_ContentM_TableRowIndex", objTable) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ContentM_TableRowIndex ] Table does not exist.")	
		Fn_ContentM_TableRowIndex = -1
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
				aRowNode = split(trim((nodeArr(iNodeCounter))),"~")
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
			aRowNode = split(trim((nodeArr(iNodeCounter))),"~")
			iInstance = 0
			bFound = False
			do While iRowCounter < iRows
				If uBound(aRowNode) > 0 Then
							' instance number exist in name
							' initialize instance num
							' ith row matches with aRowNode(0) then
							sNodePath = objTable.object.getValueAt(iRowCounter, iColIndex).toString()
							If instr(1,lcase(err.Description),"object requiered",1 )  > 0then
								Exit Do
							End if
							If trim(sNodePath) = trim(aRowNode(0)) then
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
					If instr(1,lcase(err.Description),"object requiered",1 )  > 0then
						Exit Do
					End if
					If trim(sNodePath) = trim(aRowNode(0)) then
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
				Fn_ContentM_TableRowIndex = iRowCounter
	Else
				Fn_ContentM_TableRowIndex = -1
	End If
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ContentM_TableRowIndex ] executed successfully.")
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_TableRowIndex

'Description			 :	Function Used to retrive row number from table

'Parameters			   :   '1.objTable: Table Object
'										 2.sNodeName: Node name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Table Should activated

'Examples				:   Fn_ContentM_PublicationStructureTableOperations("Select","X000001/A;1-Base1 (View):X000003/A;1-Topic1", "", "", "")
'										Fn_ContentM_PublicationStructureTableOperations("Deselect","X000001/A;1-Base1 (View):X000003/A;1-Topic1", "", "", "")
'										Fn_ContentM_PublicationStructureTableOperations("MultiSelect","X000001/A;1-Base1 (View):X000003/A;1-Topic1^X000001/A;1-Base1 (View)", "", "", "")
'										Fn_ContentM_PublicationStructureTableOperations("Exist","X000001/A;1-Base1 (View):X000003/A;1-Topic1", "", "", "")
'										Fn_ContentM_PublicationStructureTableOperations("Expand","X000001/A;1-Base1 (View):X000003/A;1-Topic1", "", "", "")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												28-Mar-2012								1.0																						Sunny R
'													Sandeep N												03-Apr-2012								1.1							Added Case : Expand					Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_PublicationStructureTableOperations(sAction, sNodeName, sColName, sValue, sPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_PublicationStructureTableOperations"
	Dim iRowCounter, objTable, aNodeNames
	Dim strMenu, aMenu, iCounter
	Dim iRows, sSelectedNodes, iBOMLineColIndex
	Dim iSubMenuCount, objContextMenu, bFound
	Dim iCount, objNodeForRow, sColour, sColourCode
	dim StrNodePath, aPath, iCnt
	Dim objEditQuan

	If JavaWindow("ContentManagement").JavaWindow("JApplet").JavaTable("PublicationStructures").Exist = True Then
		Set objTable = JavaWindow("ContentManagement").JavaWindow("JApplet").JavaTable("PublicationStructures")   
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_ContentM_PublicationStructureTableOperations] BOM Table does not exists.")
		Set objTable = nothing
		Fn_ContentM_PublicationStructureTableOperations = False
		Exit function
	End if

	Select Case sAction
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Select"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
                    objTable.Object.clearSelection  
					objTable.SelectRow iRowCounter 
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False					
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Expand"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
					objTable.SelectRow iRowCounter
					objTable.ClickCell iRowCounter ,"Publication Line", "RIGHT","NONE"
					strMenu=JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath("Expand")
					JavaWindow("ContentManagement").WinMenu("ContextMenu").Select strMenu
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False					
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Deselect"
			Fn_ContentM_PublicationStructureTableOperations = False
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
                    objTable.DeselectRow iRowCounter 
					Fn_ContentM_PublicationStructureTableOperations = True
				End If
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "ActivateHiddenSiblings"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
					objTable.Object.getNodeForRow( iRowCounter + 1 ).stateIconClicked()
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False					
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "MultiSelect"
			aNodeNames = split(sNodeName , "^")
			'Clear the already selected Nodes
			objTable.Object.clearSelection
			For iCounter = 0 to UBound(aNodeNames)
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,trim(aNodeNames(iCounter)))
				If iRowCounter <> -1 Then
					objTable.ExtendRow iRowCounter 
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
					objTable.Object.clearSelection
					Exit for
				End If
			Next
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "SelectAll"
			'Clear the already selected Nodes
			objTable.Object.clearSelection
			iRows = cInt(objTable.GetROProperty ("rows"))
			For iCounter = 0 to iRows - 1
                objTable.ExtendRow iCounter 
			Next
			Fn_ContentM_PublicationStructureTableOperations = True
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "Exist", "Exists"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "CellEdit"
			Fn_ContentM_PublicationStructureTableOperations = False
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName)		
				If iRowCounter <> -1 Then				
					objTable.SelectRow iRowCounter
					objTable.ClickCell iRowCounter,sColName, "LEFT" 
					wait 1
					If JavaWindow("ContentManagement").JavaWindow("JApplet").JavaEdit("EditBox").exist(5) Then
						JavaWindow("ContentManagement").JavaWindow("JApplet").JavaEdit("EditBox").Set sValue
						JavaWindow("ContentManagement").JavaWindow("JApplet").JavaEdit("EditBox").Activate
						'objTable.SetCellData iRowCounter,sColName, sValue
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Cell Edited of Publication Structure Table Node [" + sNodeName + "]")
						Fn_ContentM_PublicationStructureTableOperations = True
                    End If
					End If
				End If

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "CellEditList"
			If sNodeName <> "" Then
				Dim objDrpDwn, intNoOfObjects
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 		
				If iRowCounter <> -1 Then
					objTable.SelectRow iRowCounter 
					'objTable.SetCellData iRowCounter,sColName, sValue
					objTable.ClickCell iRowCounter,sColName, "LEFT"
					JavaWindow("ContentManagement").JavaWindow("JApplet").JavaButton("DropDown").Click micLeftBtn
					wait 2
					Set objDrpDwn=description.Create()
					 objDrpDwn("Class Name").value = "JavaStaticText"
					 objDrpDwn("label").value = sValue
					 Set  intNoOfObjects = JavaWindow("ContentManagement").JavaWindow("JApplet").ChildObjects(objDrpDwn)
					For iCounter = 0 to intNoOfObjects.count - 1
						If intNoOfObjects(iCounter).toString()  = "[ " & sValue & "(st) ] text label" then
							intNoOfObjects(iCounter).Click 1,1
							JavaWindow("ContentManagement").JavaWindow("JApplet").JavaEdit("EditBox").Activate
							Exit for
						Else
							Fn_ContentM_PublicationStructureTableOperations = False
							Exit for
						End If
					Next
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Cell Edited of Publication Structure Table Node [" + sNodeName + "]")
					Fn_ContentM_PublicationStructureTableOperations = True
					Set intNoOfObjects = nothing
					Set objDrpDwn = nothing
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "CellVerify"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
					'objTable.SelectRow iRowCounter 
					bFound = Trim(cstr(objTable.GetCellData( iRowCounter,sColName)))
					If bFound = Trim(cstr(sValue)) Then
						Fn_ContentM_PublicationStructureTableOperations = True
					Else
						Fn_ContentM_PublicationStructureTableOperations = False
						' workaround if UOM template is deployed
						' by Koustubh
						If isNumeric(bFound) Then
							 bFound = Abs(bFound)
							 If cstr(bFound) = Trim(cstr(sValue)) Then
								 Fn_ContentM_PublicationStructureTableOperations = True
							end  If
						End If
					End If
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Cell verified of Publication Structure Table Node [" + sNodeName + "]")
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
	'.---------------------------------------This case is used to get the Cell Value For BOM Table Node cell.----------------------------------------------
		Case "GetCellData"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
					'objTable.SelectRow iRowCounter 
					Fn_ContentM_PublicationStructureTableOperations = Trim(cstr(objTable.GetCellData( iRowCounter,sColName)))
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Cell verified of Publication Structure Table Node [" + sNodeName + "]")
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "CellListVerify"
					Fn_ContentM_PublicationStructureTableOperations = False
					If sNodeName <> "" Then
						iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName)
						If iRowCounter <> -1 Then
							objTable.SelectRow iRowCounter 
							'objTable.SetCellData iRowCounter,sColName, sValue
							objTable.ClickCell iRowCounter,sColName, "LEFT"
							JavaWindow("ContentManagement").JavaWindow("JApplet").JavaButton("DropDown").Click micLeftBtn
							wait(2)
							Set objDrpDwn=description.Create()
							 objDrpDwn("Class Name").value = "JavaStaticText"
							 objDrpDwn("label").value = sValue
							 Set  intNoOfObjects = JavaWindow("ContentManagement").JavaWindow("JApplet").ChildObjects(objDrpDwn)
							For iCounter = 0 to intNoOfObjects.count - 1
								If intNoOfObjects(iCounter).toString()  = "[ " & sValue & "(st) ] text label" then
									Fn_ContentM_PublicationStructureTableOperations = True
									Exit for
								End If
							Next
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Cell Edited of Publication Structure Table Node [" + sNodeName + "]")
							Set intNoOfObjects = nothing
							Set objDrpDwn = nothing
						End If
					End If
					
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "CellDoubleClick"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName) 
				If iRowCounter <> -1 Then
					objTable.SelectRow iRowCounter 
					objTable.DoubleClickCell iRowCounter,sColName, "LEFT", "NONE" 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Cell DoubleClicked of Publication Structure Table Node [" + sNodeName + "]")
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "PopupSelect"
			objTable.Object.clearSelection  
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName)
				If iRowCounter <> -1 Then
					'Split Context menu to Build Path Accordingly
					aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ActivateRow iRowCounter 
						objTable.ClickCell iRowCounter ,"Publication Line"
						wait 1
						objTable.ClickCell iRowCounter ,"Publication Line", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCounter ,sColName
						wait 1
						objTable.ClickCell iRowCounter ,sColName, "RIGHT","NONE"
					End If
					wait 1
					Select Case Ubound(aMenu)
						Case "0"
							strMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
							JavaWindow("ContentManagement").WinMenu("ContextMenu").Select strMenu
						Case "1"
							strMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
							JavaWindow("ContentManagement").WinMenu("ContextMenu").Select strMenu
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [ Fn_ContentM_PublicationStructureTableOperations ] Context Menu Case NOT Exists for Supplied Menu [" + StrPopupMenu + "]")
							Fn_ContentM_PublicationStructureTableOperations = False
					End Select
					Wait 30
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "MultiSelectPopupMenuSelect"
			If sNodeName <> "" Then
				aNodeNames = split(sNodeName , "^")
				'Clear the already selected Nodes
				objTable.Object.clearSelection
				For iCounter = 0 to UBound(aNodeNames)
					iRowCounter = Fn_ContentM_TableRowIndex(objTable,trim(aNodeNames(iCounter)))
					If iRowCounter <> -1 Then
						objTable.ExtendRow iRowCounter 
						Fn_ContentM_PublicationStructureTableOperations = True
					Else
						Fn_ContentM_PublicationStructureTableOperations = False
						objTable.Object.clearSelection
						Exit for
					End If
				Next
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,trim(aNodeNames(UBound(aNodeNames))))
				If iRowCounter <> -1 Then
					'Split Context menu to Build Path Accordingly
					aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowCounter ,"BOM Line", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCounter ,sColName, "RIGHT","NONE"
					End If
					wait 1
					Select Case Ubound(aMenu)
						Case "0"
							strMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenu(0))
							JavaWindow("ContentManagement").WinMenu("ContextMenu").Select strMenu
						Case "1"
							strMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu").BuildMenuPath(aMenu(0),aMenu(1))
							JavaWindow("ContentManagement").WinMenu("ContextMenu").Select strMenu
						Case Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Fail: Function [ Fn_ContentM_PublicationStructureTableOperations ] Context Menu Case NOT Exists for Supplied Menu [" + StrPopupMenu + "]")
							Fn_ContentM_PublicationStructureTableOperations = False
					End Select
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "PopupMenuExists"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName)
				If iRowCounter <> -1 Then
					'Split Context menu to Build Path Accordingly
					aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowCounter ,"BOM Line", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCounter ,sColName, "RIGHT","NONE"
					End If
					Select Case cInt(Ubound(aMenu))
						Case 0
							Set objContextMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu")
							objTable.ClickCell iRowCounter,"BOM Line", "RIGHT","NONE"
							Wait(2)
							Fn_ContentM_PublicationStructureTableOperations = objContextMenu.CheckItemProperty (sPopupMenu, "Exists",true,10)
							objTable.ClickCell iRowCounter,"BOM Line", "LEFT","NONE"
							Set objContextMenu = nothing
					End Select
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "PopupMenuEnabled"
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName)
				If iRowCounter <> -1 Then
					'Split Context menu to Build Path Accordingly
					aMenu = split(sPopupMenu,":",-1,1)
					If sColName = "" Then
						objTable.ClickCell iRowCounter ,"BOM Line", "RIGHT","NONE"
					Else
						objTable.ClickCell iRowCounter ,sColName, "RIGHT","NONE"
					End If
					Select Case cInt(Ubound(aMenu))
						Case 0
							Set objContextMenu = JavaWindow("ContentManagement").WinMenu("ContextMenu")
							objTable.ClickCell iRowCounter,"BOM Line", "RIGHT","NONE"
							Wait(2)
							If objContextMenu.CheckItemProperty (sPopupMenu, "Exists",true,10) Then
								Fn_ContentM_PublicationStructureTableOperations = objContextMenu.CheckItemProperty (sPopupMenu, "Enabled",true,10)
							End IF
							objTable.ClickCell iRowCounter,"BOM Line", "LEFT","NONE"
							Set objContextMenu = nothing
					End Select
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Popup Menu ["+ sPopupMenu +"] Selected Sucessfully")
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
				End If
			Else
				Fn_ContentM_PublicationStructureTableOperations = False
			End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case "VerifyForegroundColour", "VerifyBackgroundColour"
			Fn_ContentM_PublicationStructureTableOperations = False
			If sNodeName <> "" Then
				iRowCounter = Fn_ContentM_TableRowIndex(objTable,sNodeName)
				If cint(iRowCounter) = -1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [Fn_ContentM_PublicationStructureTableOperations] Couldnt find  Publication Structure Table Node [" + StrNodeName + "]")
					Exit function
				End If
				iRows = iRowCounter +1
				iCount = iRowCounter
			Else
				iRows = JavaWindow("ContentManagement").JavaWindow("JApplet").JavaTable("PublicationStructures").GetROProperty("rows")
				iCount = 0
			End If

			Do While cint(iCount) < cint(iRows)
				Set  objNodeForRow =  objTable.Object.getNodeForRow(cint(iCount))
				' if background colour
				If sAction = "VerifyBackgroundColour" Then
					sColour = objTable.Object.getBackground(objNodeForRow,False).toString()
				Else
				' if foreground colour
					sColour = objTable.Object.getForeground(objNodeForRow,False).toString()
				End If

				sColour =  mid(sColour ,instr(sColour ,"[")  ,instr(sColour ,"]") )
				' comparing colour codes RGB
				Select Case cstr(sValue)
					Case "BLACK"
						sColourCode = "[r=0,g=0,b=0]"
					Case "WHITE"
						sColourCode =  "[r=255,g=255,b=255]"
					Case "GRAY"
						sColourCode = "[r=178,g=180,b=191]" 
					Case "DARKGRAY"
						sColourCode = "[r=128,g=128,b=128]"
					Case "DARKBLUE"
						sColourCode = "[r=0,g=0,b=255]" 
					Case "GREEN"
						sColourCode = "[r=80,g=176,b=128]"
					Case "DARKGREEN"
						sColourCode = "[r=0,g=255,b=0]"
					Case "ORANGE"
						sColourCode = "[r=255,g=200,b=0]"
					Case "RED"
						sColourCode = "[r=255,g=0,b=0]" 
					Case "YELLOW"
						sColourCode = "[r=255,g=255,b=0]"
					Case Else
						Exit function
				End Select
				if sColour = sColourCode  Then
					Fn_ContentM_PublicationStructureTableOperations = True
				Else
					Fn_ContentM_PublicationStructureTableOperations = False
					Exit function
				End If
				iCount = iCount +1
			loop
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [Fn_ContentM_PublicationStructureTableOperations] Successfully verified colour [ " & StrValue & " ] for case [" & StrAction & "]")
			Set objNodeForRow = nothing
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Case Else
			Fn_ContentM_PublicationStructureTableOperations = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_ContentM_PublicationStructureTableOperations ] Invalid Action [ " & sAction & " ].")
			Set objTable = nothing
			exit function
			
	End Select
	If Fn_ContentM_PublicationStructureTableOperations <>FALSE then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Function [ Fn_ContentM_PublicationStructureTableOperations ] executed successfully with Action [ " & sAction & " ].")	
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Failed to execute Function [ Fn_ContentM_PublicationStructureTableOperations ] with Action [ " & sAction & " ].")
	End if
	Set objTable = nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateSchema

'Description			 :	Function Used to create NewAdministrativeClass Schema

'Parameters			   :   '1.StrID: Schema ID
'										 2.StrRevision: Schema Revision
'										 3.StrName: Schema Name
'										 4.StrPublicIdentifier: Schema Public Identifier
'										 5.StrSchemaType: Schema Type
'										 6.StrContentFilePath: Schema Content File Path
'
'Return Value		   : 	True or False

'Pre-requisite			:	

'Examples				:   Fn_ContentM_CreateSchema("","","Schema1","PI2","DTD","C:\Documents and Settings\x_navgha\Desktop\Schema1.DTD")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												02-Apr-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreateSchema(StrID,StrRevision,StrName,StrPublicIdentifier,StrSchemaType,StrContentFilePath)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateSchema"
'	Declaring variables
    Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateSchema=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	'creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking existance of [ New Administrative Class ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting Schema Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateSchema",objCMDialog, "ClassTree","Complete List:Schema")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateSchema",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateSchema",objCMDialog)
	wait 3
	'Set Schema ID
	If StrID<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateSchema",objCMDialog,"Edit",StrID)
	End If
	'Set Schema Revision
	If StrRevision<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateSchema",objCMDialog,"Edit",StrRevision)
	End If
	'Set Schema Name
	If StrName<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateSchema",objCMDialog,"Edit",StrName)
	End If
	'Set Schema Public Identifier
	If StrPublicIdentifier<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Public ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateSchema",objCMDialog,"Edit",StrPublicIdentifier)
	End If
	'Set Schema Type
	If StrSchemaType<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Schema Type:"
		Call Fn_Button_Click("Fn_ContentM_CreateSchema",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrSchemaType
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			objCMDialog.Close
			Set WshShell = Nothing
			Set objCMDialog=Nothing
			Exit Function
		End If
	End If
	If StrContentFilePath<>"" Then
		Call Fn_Button_Click("Fn_ContentM_CreateSchema",objCMDialog,"Browse")
		If objCMDialog.Dialog("Browse").Exist(10) Then
			objCMDialog.Dialog("Browse").WinEdit("FileName").Set StrContentFilePath
			wait 2 
			objCMDialog.Dialog("Browse").WinButton("Open").Click
			wait 2
		Else
			Set objCMDialog=Nothing
			Exit FUnction
		End If
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateSchema",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateSchema",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateSchema=True
	Set objCMDialog=Nothing
	Set WshShell = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_ImportGraphic

'Description			 :	Function Used to Import Graphic Options

'Parameters			   :   '1.dicImportGraphicOptionsInfo: Import Graphic Options Info
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicImportGraphicOptionsInfo("FromDirectory")="C:\Documents and Settings\x_navgha\Desktop\illustrations"
'										dicImportGraphicOptionsInfo("FileNames")="SelectAll"
'										dicImportGraphicOptionsInfo("GraphicUsage")="PDF~SOURCE"
'										dicImportGraphicOptionsInfo("GraphicAttributeMapping")="Default Graphic Attribute Mapping"
'										dicImportGraphicOptionsInfo("GraphicClassname")="Graphic"
'										dicImportGraphicOptionsInfo("Language")="English US"
'										dicImportGraphicOptionsInfo("OverwriteMode")="Overwrite existing"
'										bReturn=Fn_ContentM_ImportGraphic(dicImportGraphicOptionsInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												03-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_ImportGraphic(dicImportGraphicOptionsInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_ImportGraphic"
 	'Variable declaration
	Dim objCMDialog,WshShell
	Dim StrMenu,aFileNames,iRowCount,iCounter,iCount,bFlag,aGraphicUsage,cFileName,cGraphicUsage,cLanguage
	Fn_ContentM_ImportGraphic=False
	'Creating Object of [ ImportGraphicOptions ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("ImportGraphicOptions")
	'Checking existance of [ ImportGraphicOptions ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ Tools->Import->Graphic... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "ImportGraphic")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Setting From Directory
'	Call Fn_Button_Click("Fn_ContentM_ImportGraphic",objCMDialog,"Browse")
	JavaWindow("ContentManagement").JavaWindow("ImportGraphicOptions").JavaButton("Browse").Click micLeftBtn
	wait 2
	objCMDialog.Dialog("BrowseForFolder").WinEdit("FolderPath").Click
	wait 10
	set WshShell = CreateObject("WScript.Shell")
	WshShell.SendKeys "^a"
	wait 1
	WshShell.SendKeys "{DELETE}"
	wait 1
	set WshShell =Nothing
	objCMDialog.Dialog("BrowseForFolder").WinEdit("FolderPath").Type dicImportGraphicOptionsInfo("FromDirectory")
'	objCMDialog.Dialog("BrowseForFolder").WinEdit("FolderPath").Set dicImportGraphicOptionsInfo("FromDirectory")
	wait 1
	objCMDialog.Dialog("BrowseForFolder").WinButton("OK").Click
	wait 2
	Call Fn_ReadyStatusSync(1)
	'Select file names
	If LCase(dicImportGraphicOptionsInfo("FileNames"))="selectall" Then
		Call Fn_Button_Click("Fn_ContentM_ImportGraphic",objCMDialog,"SelectAll")
	Else
		Call Fn_Button_Click("Fn_ContentM_ImportGraphic",objCMDialog,"DeselectAll")
		aFileNames=Split(dicImportGraphicOptionsInfo("FileNames"),"~")
		iRowCount=objCMDialog.JavaTable("FileNames").GetROProperty("rows")
		For iCounter=0 to UBound(aFileNames) 
			bFlag=False
			For iCount=0 to iRowCount-1
				cFileName=objCMDialog.JavaTable("FileNames").GetCellData(iCount,0)
				If trim(cFileName)=trim(aFileNames(iCounter)) Then
					objCMDialog.JavaTable("FileNames").SelectCell iCount,0
					objCMDialog.JavaTable("FileNames").PressKey " "
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Set objCMDialog=Nothing
				Exit Function
			End If
		Next
	End If
	'Select Graphic Usage
	If LCase(dicImportGraphicOptionsInfo("GraphicUsage"))="use graphic usages from graphics mapping" or LCase(dicImportGraphicOptionsInfo("GraphicUsage"))="usegraphicusagesfromgraphicsmapping" Then
		Call Fn_CheckBox_Set("Fn_ContentM_ImportGraphic", objCMDialog,"UseGraphicUsagesFromGraphicsMapping","on")
	Else
		aGraphicUsage=Split(dicImportGraphicOptionsInfo("GraphicUsage"),"~")
		iRowCount=objCMDialog.JavaTable("GraphicUsage").GetROProperty("rows")
		For iCounter=0 to UBound(aGraphicUsage) 
			bFlag=False
			For iCount=0 to iRowCount-1
				cGraphicUsage=objCMDialog.JavaTable("GraphicUsage").GetCellData(iCount,0)
				If trim(cGraphicUsage)=trim(aGraphicUsage(iCounter)) Then
					objCMDialog.JavaTable("GraphicUsage").SelectCell iCount,0
					objCMDialog.JavaTable("GraphicUsage").PressKey " "
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Set objCMDialog=Nothing
				Exit Function
			End If
		Next	
	End If
	'Setting Graphic Attribute Mapping
	If dicImportGraphicOptionsInfo("GraphicAttributeMapping")<>"" Then
		Call Fn_List_Select("Fn_ContentM_ImportGraphic",objCMDialog,"GraphicAttributeMapping",dicImportGraphicOptionsInfo("GraphicAttributeMapping"))
	End If
	'Setting Graphic Classname
	If dicImportGraphicOptionsInfo("GraphicClassname")<>"" Then
		Call Fn_List_Select("Fn_ContentM_ImportGraphic",objCMDialog,"GraphicClassname",dicImportGraphicOptionsInfo("GraphicClassname"))
	End If
	'Setting Language
	If dicImportGraphicOptionsInfo("Language")<>"" Then
			bFlag=False
			iRowCount=objCMDialog.JavaTable("Language").GetROProperty("rows")
			For iCount=0 to iRowCount-1
				cLanguage=objCMDialog.JavaTable("Language").GetCellData(iCount,0)
				If trim(cLanguage)=trim(dicImportGraphicOptionsInfo("Language")) Then
					objCMDialog.JavaTable("Language").SelectCell iCount,0
					objCMDialog.JavaTable("Language").PressKey " "
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Set objCMDialog=Nothing
				Exit Function
			End If
	End If
	'Selecting Overwrite mode option
	If dicImportGraphicOptionsInfo("OverwriteMode")<>"" Then
		objCMDialog.JavaRadioButton("OverwriteMode").SetTOProperty "attached text",dicImportGraphicOptionsInfo("OverwriteMode")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_ImportGraphic",objCMDialog, "OverwriteMode")
	End If
	'Selecting Usage option
	If dicImportGraphicOptionsInfo("Usages")<>"" Then
		objCMDialog.JavaRadioButton("Usages").SetTOProperty "attached text",dicImportGraphicOptionsInfo("Usages")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_ImportGraphic",objCMDialog, "Usages")
	End If
	Call Fn_Button_Click("Fn_ContentM_ImportGraphic",objCMDialog,"Finish")

	bFlag=False
	For iCount=1 to 10
		JavaWindow("Shell").SetTOProperty "index",iCount
		For iCounter=0 To 10
			If JavaWindow("Shell").JavaWindow("PleaseWait").Exist(1) Then
				wait 5
				bFlag=True
			Else
				Exit for
			End If
		Next
		If bFlag=True Then
			Exit for
		End If
	Next
	'Added code to handle SNS not selected dialog while importing graphic 
	If Fn_UI_ObjectExist("Fn_ContentM_ImportGraphic", JavaDialog("SNSDialog")) = True Then
		Call Fn_Button_Click("Fn_ContentM_ImportGraphic",JavaDialog("SNSDialog"),"Yes")
		wait 1
	End If
	
	If Fn_UI_ObjectExist("Fn_ContentM_ImportGraphic", objCMDialog.JavaWindow("Warning")) = True Then
		Call Fn_Button_Click("Fn_ContentM_ImportGraphic",objCMDialog.JavaWindow("Warning"),"OK")
		wait 1
	End If

	Fn_ContentM_ImportGraphic=True
	Set objCMDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateS1000DDataDispatchNote4

'Description			 :	Function Used to Create S1000D Data Dispatch Note 4.0 and S1000D Data Dispatch Note

'Parameters			   :   '1.dicS1000DDataDispatchNote4Info: S1000D Data Dispatch Note 4.0 Info
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicS1000DDataDispatchNote4Info("TopicType")="DDN-4-0"
'										dicS1000DDataDispatchNote4Info("Name")="DDN24"
'										dicS1000DDataDispatchNote4Info("MasterLanguageReference")="English US"
'										dicS1000DDataDispatchNote4Info("DocumentTitle")="DDN Title 1"
'										dicS1000DDataDispatchNote4Info("ModelIdentificationCode")="MI1"
'										dicS1000DDataDispatchNote4Info("Originator")="Org1"
'										dicS1000DDataDispatchNote4Info("ReceiverIdentification")="RI1"
'										dicS1000DDataDispatchNote4Info("YearOfDispatch")="2011"
'										dicS1000DDataDispatchNote4Info("SequenceNumber")="11"
'										dicS1000DDataDispatchNote4Info("IssueNumber")="101"
'										dicS1000DDataDispatchNote4Info("IssueType")="Type1"
'										dicS1000DDataDispatchNote4Info("IssuedDay")="10"
'										dicS1000DDataDispatchNote4Info("IssuedMonth")="7"
'										dicS1000DDataDispatchNote4Info("IssuedYear")="2010"
'										dicS1000DDataDispatchNote4Info("SecurityClass")="SC1"
'										dicS1000DDataDispatchNote4Info("AuthorizationIdentification")="AI1"
'										dicS1000DDataDispatchNote4Info("MediaIdentification")="MDI1"
'										dicS1000DDataDispatchNote4Info("Remarks")="Useless DDN"
'										dicS1000DDataDispatchNote4Info("InWorkNumber")="001"
'										dicS1000DDataDispatchNote4Info("ExportFileName")="File1"
'										dicS1000DDataDispatchNote4Info("DispatchToEnterpriseName")="EN1"
'										dicS1000DDataDispatchNote4Info("DispatchToCity")="City1"
'										dicS1000DDataDispatchNote4Info("DispatchToCountry")="Country1"
'										dicS1000DDataDispatchNote4Info("DispatchFromCompanyName")="Company1"
'										dicS1000DDataDispatchNote4Info("DispatchFromCity")="City2"
'										dicS1000DDataDispatchNote4Info("DispatchFromCountry")="Country2"
'										bReturn=Fn_ContentM_CreateS1000DDataDispatchNote4(dicS1000DDataDispatchNote4Info)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												04-Apr-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												17-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreateS1000DDataDispatchNote4(dicS1000DDataDispatchNote4Info)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateS1000DDataDispatchNote4"
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateS1000DDataDispatchNote4=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	If dicS1000DDataDispatchNote4Info("AuthorClass")="" Then
		'Selecting S1000D Data Dispatch Note 4.0 Option from list
		Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog, "ClassTree","Complete List:S1000D Data Dispatch Note 4.0/4.1/4.2")
	Else
		Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog, "ClassTree","Complete List:"+dicS1000DDataDispatchNote4Info("AuthorClass"))
	End If
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog)
	wait 3
	'Selecting topic type
	If dicS1000DDataDispatchNote4Info("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataDispatchNote4", objCMDialog, "SelectTopicType",dicS1000DDataDispatchNote4Info("TopicType"))
	End If
	'Set revision
	If dicS1000DDataDispatchNote4Info("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"Revision",dicS1000DDataDispatchNote4Info("Revision"))
	End If
	'Set Name
	If dicS1000DDataDispatchNote4Info("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"Name",dicS1000DDataDispatchNote4Info("Name"))
	End If
	'Setting Master Language Reference
	If dicS1000DDataDispatchNote4Info("MasterLanguageReference")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicS1000DDataDispatchNote4Info("MasterLanguageReference")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting Document Title
	If dicS1000DDataDispatchNote4Info("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"DocumentTitle",dicS1000DDataDispatchNote4Info("DocumentTitle"))
	End If
	'Setting Model Identification Code
	If dicS1000DDataDispatchNote4Info("ModelIdentificationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Model Identification Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"ModelIdentificationCode",dicS1000DDataDispatchNote4Info("ModelIdentificationCode"))
	End If
	'Setting Originator
	If dicS1000DDataDispatchNote4Info("Originator")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"Originator",dicS1000DDataDispatchNote4Info("Originator"))
	End If
	'Setting reciever Identification
	If dicS1000DDataDispatchNote4Info("ReceiverIdentification")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Receiver Identification:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"ReceiverIdentification",dicS1000DDataDispatchNote4Info("ReceiverIdentification"))
	End If
	'Setting Year of Dispatch
	If dicS1000DDataDispatchNote4Info("YearOfDispatch")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Year of Dispatch:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"YearofDispatch",dicS1000DDataDispatchNote4Info("YearOfDispatch"))
	End If
	'Setting Sequence Number
	If dicS1000DDataDispatchNote4Info("SequenceNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Sequence Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"SequenceNumber",dicS1000DDataDispatchNote4Info("SequenceNumber"))
	End If
	'Setting Issue Number
	If dicS1000DDataDispatchNote4Info("IssueNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"IssueNumber",dicS1000DDataDispatchNote4Info("IssueNumber"))
	End If
	'Setting Issue Type
	If dicS1000DDataDispatchNote4Info("IssueType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"IssueType",dicS1000DDataDispatchNote4Info("IssueType"))
	End If
	'Setting Issued Day
	If dicS1000DDataDispatchNote4Info("IssuedDay")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Day:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"IssueDay",dicS1000DDataDispatchNote4Info("IssuedDay"))
	End If
	'Setting Issued Month
	If dicS1000DDataDispatchNote4Info("IssuedMonth")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Month:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"IssueMonth",dicS1000DDataDispatchNote4Info("IssuedMonth"))
	End If
	'Setting Issued Year
	If dicS1000DDataDispatchNote4Info("IssuedYear")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Year:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"IssueYear",dicS1000DDataDispatchNote4Info("IssuedYear"))
	End If
	'Setting Security Class
	If dicS1000DDataDispatchNote4Info("SecurityClass")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Security Class:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"SecurityClass",dicS1000DDataDispatchNote4Info("SecurityClass"))
	End If
	'Setting Authorization Identification
	If dicS1000DDataDispatchNote4Info("AuthorizationIdentification")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Authorization Identification:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"AuthorizationIdentification",dicS1000DDataDispatchNote4Info("AuthorizationIdentification"))
	End If
	'Setting Media Identification
	If dicS1000DDataDispatchNote4Info("MediaIdentification")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Media Identification:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"MediaIdentification",dicS1000DDataDispatchNote4Info("MediaIdentification"))
	End If
	'Setting Remarks
	If dicS1000DDataDispatchNote4Info("Remarks")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Remarks:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"Remarks",dicS1000DDataDispatchNote4Info("Remarks"))
	End If
	'Setting In Work Number
	If dicS1000DDataDispatchNote4Info("InWorkNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","In Work Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"InWorkNumber",dicS1000DDataDispatchNote4Info("InWorkNumber"))
	End If
	'Setting Export File Name
	If dicS1000DDataDispatchNote4Info("ExportFileName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Export File Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"ExportFileName",dicS1000DDataDispatchNote4Info("ExportFileName"))
	End If
	'Set Is This A Template option
	If dicS1000DDataDispatchNote4Info("IsThisATemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicS1000DDataDispatchNote4Info("IsThisATemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog, "IsThisATemplate")
	End If
	'Set Reference Only option
	If dicS1000DDataDispatchNote4Info("ReferenceOnly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference Only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicS1000DDataDispatchNote4Info("ReferenceOnly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog, "ReferenceOnly")
	End If
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"Next")
	'Setting Dispatch To Enterprise Name
	If dicS1000DDataDispatchNote4Info("DispatchToEnterpriseName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch To Enterprise Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"DispatchToEnterpriseName",dicS1000DDataDispatchNote4Info("DispatchToEnterpriseName"))
	End If
	'Setting Dispatch To City
	If dicS1000DDataDispatchNote4Info("DispatchToCity")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch To City:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"DispatchToCity",dicS1000DDataDispatchNote4Info("DispatchToCity"))
	End If
	'Setting Dispatch To Country
	If dicS1000DDataDispatchNote4Info("DispatchToCountry")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch To Country:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"DispatchToCountry",dicS1000DDataDispatchNote4Info("DispatchToCountry"))
	End If
	'Setting Dispatch From Company Name
	If dicS1000DDataDispatchNote4Info("DispatchFromCompanyName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch From Company Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"DispatchFromCompanyName",dicS1000DDataDispatchNote4Info("DispatchFromCompanyName"))
	End If
	'Setting Dispatch From City
	If dicS1000DDataDispatchNote4Info("DispatchFromCity")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch From City:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"DispatchFromCity",dicS1000DDataDispatchNote4Info("DispatchFromCity"))
	End If
	'Setting Dispatch From Country
	If dicS1000DDataDispatchNote4Info("DispatchFromCountry")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch From Country:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"DispatchFromCountry",dicS1000DDataDispatchNote4Info("DispatchFromCountry"))
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataDispatchNote4",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateS1000DDataDispatchNote4=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_ImportDocumentsFromFile

'Description			 :	Function Used to Import Documents From File

'Parameters			   :   '1.StrDirectoryPath: Document Directory Path
'										StrFileNames = File name for selection
'										bShowAllFiles = Show all Files option
'										StrStylesheet = Stylesheet type
'										StrGraphicAttributeMapping = Graphic Attribute Mapping type
'										bGraphicMode = Graphic mode option
'										bFindByXMLNumber = Find By XML Number option
'										bOverwriteExisting =  Overwrite Existing option
'										bFindByContent = Find By Content option
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	Fn_ContentM_ImportDocumentsFromFile("C:\Documents and Settings\x_navgha\My Documents\Downloads\c9_Import_DMwithDME","DME-SF518-CE0701-S1000DBIKE-AAA-D00-00-00-00AA-131A-A_007-00_EN-US.xml~DME-SF518-MT0701-S1000DBIKE-AAA-D00-00-00-00AA-131A-A_007-00_EN-US.xml","","","S1000D v 4.0","XML Number","on","on","on")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												04-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_ImportDocumentsFromFile(StrDirectoryPath,StrFileNames,bShowAllFiles,StrStylesheet,StrGraphicAttributeMapping,bGraphicMode,bFindByXMLNumber,bOverwriteExisting,bFindByContent)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_ImportDocumentsFromFile"
   'Variable declaration
	Dim objCMDialog,StrMenu
	Dim aFileNames,iRowCount,iCounter,iCount,bFlag,cFileName
	Fn_ContentM_ImportDocumentsFromFile=False
	'Creating Object of [ ImportDocumentsFromFile ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("ImportDocumentsFromFile")
	'Checking existance of [ ImportGraphicOptions ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ Tools->Import->Graphic... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "ImportDocument")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(5)
	End If
'	Setting From Directory
	Call Fn_Button_Click("Fn_ContentM_ImportDocumentsFromFile",objCMDialog,"Browse")
	wait(3)
	objCMDialog.Dialog("BrowseForFolder").WinEdit("FolderPath").Set StrDirectoryPath
	wait 1
	objCMDialog.Dialog("BrowseForFolder").WinButton("OK").Click
	wait 2
	Call Fn_ReadyStatusSync(1)
	If StrFileNames<>"" Then
		'Select file names
		If LCase(StrFileNames)="selectall" Then
			Call Fn_Button_Click("Fn_ContentM_ImportDocumentsFromFile",objCMDialog,"SelectAll")
			wait(3)
		Else
			Call Fn_Button_Click("Fn_ContentM_ImportDocumentsFromFile",objCMDialog,"DeselectAll")
			aFileNames=Split(StrFileNames,"~")
			iRowCount=objCMDialog.JavaTable("FileNames").GetROProperty("rows")
			For iCounter=0 to UBound(aFileNames) 
				bFlag=False
				For iCount=0 to iRowCount-1
					cFileName=objCMDialog.JavaTable("FileNames").GetCellData(iCount,"File name")
					If trim(cFileName)=trim(aFileNames(iCounter)) Then
						objCMDialog.JavaTable("FileNames").SelectCell iCount,0
						objCMDialog.JavaTable("FileNames").PressKey " "
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Set objCMDialog=Nothing
					Exit Function
				End If
			Next
		End If
	End If
	'Selecting Show All Files option
	If bShowAllFiles<>"" Then
		wait(3)
		Call Fn_CheckBox_Set("Fn_ContentM_ImportDocumentsFromFile", objCMDialog,"ShowAllFiles",bShowAllFiles)
	End If
	'Selecting Stylesheet type
	If StrStylesheet<>"" Then
		wait(3)
		Call Fn_List_Select("Fn_ContentM_ImportDocumentsFromFile",objCMDialog,"Stylesheet",StrStylesheet)
	End If
	'Selecting Graphic Attribute Mapping type
	If StrGraphicAttributeMapping<>"" Then
		wait(3)
		Call Fn_List_Select("Fn_ContentM_ImportDocumentsFromFile",objCMDialog,"GraphicAttributeMapping",StrGraphicAttributeMapping)
	End If
	'Selecting Graphic mode option
	If bGraphicMode<>"" Then
		objCMDialog.JavaRadioButton("GraphicMode").SetTOProperty "attached text",bGraphicMode
			wait(3)
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_ImportDocumentsFromFile",objCMDialog, "GraphicMode")	
	End If
	'Selecting Find By XML Number option
	If bFindByXMLNumber<>"" Then
		Call Fn_CheckBox_Set("Fn_ContentM_ImportDocumentsFromFile", objCMDialog,"FindByXMLNumber",bFindByXMLNumber)
			wait(3)
	End If
	'Selecting Overwrite Existing option
	If bOverwriteExisting<>"" Then
		wait(3)
		Call Fn_CheckBox_Set("Fn_ContentM_ImportDocumentsFromFile", objCMDialog,"OverwriteExisting",bOverwriteExisting)
	End If
	'Selecting Find By Content option
	If bFindByContent<>"" Then
		wait(3)
		Call Fn_CheckBox_Set("Fn_ContentM_ImportDocumentsFromFile", objCMDialog,"FindByContent",bFindByContent)
	End If
	Call Fn_Button_Click("Fn_ContentM_ImportGraphic",objCMDialog,"Finish")

	bFlag=False
	For iCount=1 to 10
		JavaWindow("Shell").SetTOProperty "index",iCount
		For iCounter=0 To 10
			If JavaWindow("Shell").JavaWindow("PleaseWait").Exist(1) Then
				wait 5
				bFlag=True
			Else
				Exit for
			End If
		Next
		If bFlag=True Then
			Exit for
		End If
	Next
	'Added code to handle SNS not selected dialog while importing document 
	If Fn_UI_ObjectExist("Fn_ContentM_ImportGraphic", JavaDialog("SNSDialog")) = True Then
		Call Fn_Button_Click("Fn_ContentM_ImportGraphic",JavaDialog("SNSDialog"),"Yes")
		wait 1
	End If
	
	Fn_ContentM_ImportDocumentsFromFile=True
	Set objCMDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_SummaryPropertyVerify

'Description			 :	Function Used verify property and its value from Summary page

'Parameters			   :   '1.StrProperty: Property name and value ( Name:Newstuff;ID:X00009)
'
'Return Value		   : 	True or False

'Pre-requisite			:	Summary page should activated in Content Management perspective

'Examples				:  	Fn_ContentM_SummaryPropertyVerify("Extension Code:DMEC777")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												04-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_SummaryPropertyVerify(StrProperty)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_SummaryPropertyVerify"
	Dim objCMDialog,aProperty,iCounter,bCheck,iProperty,strValue,aPropertyValue,aPropertyName

	Fn_ContentM_SummaryPropertyVerify = False
	'Creating object of [ ContentManagement ] window
	Set objCMDialog = JavaWindow("ContentManagement")
	'Spliting Multiple properties
	aProperty = Split(StrProperty,";")
	For iCounter = 0 to UBound(aProperty)
			bCheck=False
			'Spliting property name and value
			iProperty = Split(aProperty(iCounter),":")
			aPropertyName= iProperty(0)
			aPropertyValue= iProperty(1)		
			Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_SummaryPropertyVerify",objCMDialog.JavaStaticText("Summary_Text"),"label",aPropertyName+":")
			If objCMDialog.JavaEdit("Summary_Text").Exist(6) Then
				strValue=Fn_UI_Object_GetROProperty("Fn_ContentM_SummaryPropertyVerify",objCMDialog.JavaEdit("Summary_Text"),"value")
				If Trim(strValue)=Trim(aPropertyValue) Then
					bCheck=True
				End If
			ElseIf objCMDialog.JavaObject("Summary_Object").Exist(6) Then
				strValue=Fn_UI_Object_GetROProperty("Fn_ContentM_SummaryPropertyVerify",objCMDialog.JavaObject("Summary_Object"), "text")
				If Trim(strValue)=Trim(aPropertyValue) Then
					bCheck=True
				End If
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Property [ "+aPropertyName+" ] does not exist on Summary Page")
				Set objCMDialog =Nothing
				Exit Function
			End If
			If bCheck=False Then
				Set objCMDialog =Nothing
				Exit Function
			End If
	Next
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Function Fn_ContentM_SummaryPropertyVerify Completed Successfully")
	Fn_ContentM_SummaryPropertyVerify = true
	objCMDialog.JavaStaticText("Summary_Text").SetTOProperty "label",""
	'Releasing object of [ ContentManagement ] window
	Set objCMDialog =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_TextPadOperations

'Description			 :	Function Used perform operations on Text Pad

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrFilePath : File path
'										 3.StrContent : Content to verify
'
'Return Value		   : 	True or False

'Pre-requisite			:	

'Examples				:  	Fn_ContentM_TextPadOperations("TextVerify","C:\Documents and Settings\x_navgha\Teamcenter\contmgmt\9000.0.0\edit\Notepad\C3DO-A-3-33-3-23-33-333\C3DO-A-3-33-3-23-33-333.xml","dmodule xmlns:dc=""http://www.purl.org/dc/elements/1.1/""")
'										Fn_ContentM_TextPadOperations("Close","","")
'										Fn_ContentM_TextPadOperations("AppendText","C:\Documents and Settings\x_navgha\Teamcenter\contmgmt\9000.0.0\edit\Notepad\C3DO-A-3-33-3-23-33-333\C3DO-A-3-33-3-23-33-333.xml","<Name>Test1</Name>")
'	
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												09-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_TextPadOperations(StrAction,StrFilePath,StrContent)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_TextPadOperations"
   'Declaring variables
   Dim fso,MyFile,sTextLine
   Fn_ContentM_TextPadOperations=False
   Select Case StrAction
	 	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	 	Case "TextVerify"
			'Checking Existance of TextPad
			If Window("TextPad").Exist Then
				'Creating Object of File System
				Set fso =CreateObject("Scripting.FileSystemObject")
				'Opening File
				If  fso.FileExists(StrFilePath)Then
					Set MyFile = fso.OpenTextFile(StrFilePath,1,True)
					Do While MyFile.AtEndOfStream <>True
						'Reading Data line by line for the file
						sTextLine =sTextLine+ MyFile.ReadLine
					Loop
					'Verifing Data with file data
					If InStr(1,LCase(Trim(sTextLine)),LCase(Trim(StrContent)))>0 Then
						Fn_ContentM_TextPadOperations=True
						MyFile.Close
					Else
						MyFile.Close
						Window("TextPad").Close
					End If
				End If
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "AppendText"
				'Creating Object of File System
				Set fso =CreateObject("Scripting.FileSystemObject")
				'Opening File
				If  fso.FileExists(StrFilePath)Then
					Set MyFile = fso.OpenTextFile(StrFilePath,1,True)
						
					Do While MyFile.AtEndOfStream <>True
						'Reading Data line by line for the file
						sTextLine =sTextLine+ MyFile.ReadLine
					Loop
				
					Set MyFile = fso.OpenTextFile(StrFilePath,2,True)
					MyFile.Write sTextLine+" "+StrContent
		
					'Verifing Data with file data
					Fn_ContentM_TextPadOperations=True
				End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Close"
			'Case to close text pad
			If Window("TextPad").Exist Then
				Window("TextPad").Close
				Fn_ContentM_TextPadOperations=True
			End If
   End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_WindowPreferencesOperations

'Description			 :	Function Used perform operations on Window Preferences

'Parameters			   :   '1.StrAction: Action Name
'										 2.dicWindowPreferencesInfo : Window preferences information
'
'Return Value		   : 	True or False

'Pre-requisite			:	

'Examples				:  	dicWindowPreferencesInfo("Editor")="TextPad"
'										Fn_ContentM_WindowPreferencesOperations("ReloadTools",dicWindowPreferencesInfo)
'	
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												09-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_WindowPreferencesOperations(StrAction,dicWindowPreferencesInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_WindowPreferencesOperations"
	Dim objPrefDialog,StrMenu,bFlag
	bFlag=False
	Fn_ContentM_WindowPreferencesOperations=False
   Set objPrefDialog=JavaWindow("ContentManagement").JavaWindow("Preferences")
   If Not objPrefDialog.Exist(6) Then
	   'Select menu [ Window->Preferences ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "WindowPreferences")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
   End If
   Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	 	Case "ReloadTools"
			objPrefDialog.JavaTree("Tree").Select "Teamcenter:Content Management:Edit"
			bFlag=Fn_UI_ListItemExist("Fn_ContentM_WindowPreferencesOperations",objPrefDialog,"Editor",dicWindowPreferencesInfo("Editor"))
			If bFlag=True Then
				Call Fn_List_Select("Fn_ContentM_WindowPreferencesOperations",objPrefDialog,"Editor",dicWindowPreferencesInfo("Editor"))
				Call Fn_Button_Click("Fn_ContentM_WindowPreferencesOperations",objPrefDialog, "ReloadTools")
				If objPrefDialog.JavaWindow("ReloadTools").Exist(6) Then
					Call Fn_Button_Click("Fn_ContentM_WindowPreferencesOperations",objPrefDialog.JavaWindow("ReloadTools"), "OK")
				End If
				wait 1
				'To retsart Tool message server 
				Call Fn_Button_Click("Fn_ContentM_WindowPreferencesOperations",objPrefDialog, "RestartToolMessageServer")
				If objPrefDialog.JavaWindow("ToolMessageServerRestarted").Exist(6) Then
					Call Fn_Button_Click("Fn_ContentM_WindowPreferencesOperations",objPrefDialog.JavaWindow("ToolMessageServerRestarted"), "OK")
				End If				
				wait 1
			        Call Fn_Button_Click("Fn_ContentM_WindowPreferencesOperations",objPrefDialog, "Apply and Close")
			        wait 1
'				Call Fn_Button_Click("Fn_ContentM_WindowPreferencesOperations",objPrefDialog, "OK")
				Fn_ContentM_WindowPreferencesOperations=True
			End If
   End Select
   Set objPrefDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateS1000DDataModuleList4

'Description			 :	Function Used to Create S1000D Data Module List 4.0

'Parameters			   :   '1.dicS1000DDataModuleList4Info: S1000D Data Module List 4.0 Info
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicS1000DDataModuleList4Info("TopicType")="DML-4-0"
'										dicS1000DDataModuleList4Info("Revision")="A"
'										dicS1000DDataModuleList4Info("Name")="DML2486"
'										dicS1000DDataModuleList4Info("MasterLanguageReference")="English US"
'										dicS1000DDataModuleList4Info("DocumentTitle")="DML2486 Title"
'										dicS1000DDataModuleList4Info("ModelIdentificationCode")="MID1"
'										dicS1000DDataModuleList4Info("Originator")="zzz"
'										dicS1000DDataModuleList4Info("TypeOfDataModuleList")="Type1"
'										dicS1000DDataModuleList4Info("YearOfDispatch")="2011"
'										dicS1000DDataModuleList4Info("SequenceNumber")="1"
'										dicS1000DDataModuleList4Info("IssueNumber")="6"
'										dicS1000DDataModuleList4Info("IssueType")="IT1"
'										dicS1000DDataModuleList4Info("IssuedDay")="11"
'										dicS1000DDataModuleList4Info("IssuedMonth")="11"
'										dicS1000DDataModuleList4Info("IssuedYear")="2010"
'										dicS1000DDataModuleList4Info("Remarks")="OK"
'										dicS1000DDataModuleList4Info("InWorkNumber")="W1"
'										dicS1000DDataModuleList4Info("SecurityClass")="SC6"
'										dicS1000DDataModuleList4Info("ExportFileName")="File1"
'										dicS1000DDataModuleList4Info("IsThisATemplate")="False"
'										dicS1000DDataModuleList4Info("ReferenceOnly")="False"
'										bReturn=Fn_ContentM_CreateS1000DDataModuleList4(dicS1000DDataModuleList4Info)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												10-Apr-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												17-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Shweta Rathod											08-May-2015								1.2					Modified function as per 10.1.4 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Function Fn_ContentM_CreateS1000DDataModuleList4(dicS1000DDataModuleList4Info)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateS1000DDataModuleList4"
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateS1000DDataModuleList4=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	If dicS1000DDataModuleList4Info("AuthorClass")="" Then
		'Selecting S1000D Data Module List 4.0 Option from list
		Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog, "ClassTree","Complete List:S1000D Data Module List 4.0/4.1/4.2")
	Else
		Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog, "ClassTree","Complete List:"+dicS1000DDataModuleList4Info("AuthorClass"))
	End If
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog)
	wait 3
	'Selecting topic type
	If dicS1000DDataModuleList4Info("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataModuleList4", objCMDialog, "SelectTopicType",dicS1000DDataModuleList4Info("TopicType"))
	End If
	'Set revision
	If dicS1000DDataModuleList4Info("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Revision",dicS1000DDataModuleList4Info("Revision"))
	End If
	'Set Name
	If dicS1000DDataModuleList4Info("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Name",dicS1000DDataModuleList4Info("Name"))
	End If
	'Setting Master Language Reference
	If dicS1000DDataModuleList4Info("MasterLanguageReference")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
'		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"DropDownButton")
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataModuleList4", objCMDialog, "SelectTopicType",dicS1000DDataModuleList4Info("MasterLanguageReference"))
'		wait 1
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicS1000DDataModuleList4Info("MasterLanguageReference")
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
'		If bFlag=False Then
'			Set objCMDialog=Nothing
'			Set WshShell=Nothing
'			Exit Function
'		End If
	End If
	'Setting Document Title
	If dicS1000DDataModuleList4Info("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"DocumentTitle",dicS1000DDataModuleList4Info("DocumentTitle"))
	End If
	'Setting Model Identification Code
	If dicS1000DDataModuleList4Info("ModelIdentificationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Model Identification Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"ModelIdentificationCode",dicS1000DDataModuleList4Info("ModelIdentificationCode"))
	End If
	'Setting Originator
	If dicS1000DDataModuleList4Info("Originator")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Originator",dicS1000DDataModuleList4Info("Originator"))
	End If
	'Setting Type of Data Module List
	If dicS1000DDataModuleList4Info("TypeOfDataModuleList")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Type of Data Module List:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"TypeofDataModuleList",dicS1000DDataModuleList4Info("TypeOfDataModuleList"))
	End If
	'Setting Year of Dispatch
	If dicS1000DDataModuleList4Info("YearOfDispatch")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Year of Dispatch:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"YearofDispatch",dicS1000DDataModuleList4Info("YearOfDispatch"))
	End If
	'Setting Sequence Number
	If dicS1000DDataModuleList4Info("SequenceNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Sequence Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"SequenceNumber",dicS1000DDataModuleList4Info("SequenceNumber"))
	End If
	
'	Select Case dicS1000DDataModuleList4Info("AuthorClass") 
'		Case "S1000D Data Module List"	
'		Setting Issue Number
			If dicS1000DDataModuleList4Info("IssueNumber")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Number:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueNumber",dicS1000DDataModuleList4Info("IssueNumber"))
			End If
			'Setting Issue Type
			If dicS1000DDataModuleList4Info("IssueType")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Type:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueType",dicS1000DDataModuleList4Info("IssueType"))
			End If
			'Setting Issued Day
			If dicS1000DDataModuleList4Info("IssuedDay")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Day:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueDay",dicS1000DDataModuleList4Info("IssuedDay"))
			End If
			'Setting Issued Month
			If dicS1000DDataModuleList4Info("IssuedMonth")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Month:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueMonth",dicS1000DDataModuleList4Info("IssuedMonth"))
			End If
			'Setting Issued Year
			If dicS1000DDataModuleList4Info("IssuedYear")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Year:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueYear",dicS1000DDataModuleList4Info("IssuedYear"))
			End If
'	End Select

	'Setting Remarks
	If dicS1000DDataModuleList4Info("Remarks")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Remarks:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Remarks",dicS1000DDataModuleList4Info("Remarks"))
	End If
	'Setting In Work Number
	If dicS1000DDataModuleList4Info("InWorkNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","In Work Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"InWorkNumber",dicS1000DDataModuleList4Info("InWorkNumber"))
	End If
	'Setting Security Class
	If dicS1000DDataModuleList4Info("SecurityClass")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Security Class:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"SecurityClass",dicS1000DDataModuleList4Info("SecurityClass"))
	End If
    'Setting Export File Name
'	If dicS1000DDataModuleList4Info("ExportFileName")<>"" Then
'		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Export File Name:"
'		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"ExportFileName",dicS1000DDataModuleList4Info("ExportFileName"))
'	End If
'	'Set Is This A Template option
	If dicS1000DDataModuleList4Info("IsThisATemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicS1000DDataModuleList4Info("IsThisATemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog, "IsThisATemplate")
	End If
'	'Set Reference Only option
	If dicS1000DDataModuleList4Info("ReferenceOnly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicS1000DDataModuleList4Info("ReferenceOnly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog, "ReferenceOnly")
	End If
    'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateS1000DDataModuleList4=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateS1000DDataModule4

'Description			 :	Function Used to create S1000D Data Module 4.0

'Parameters			   :   '1.dicS1000DDataModule4Info: S1000D Data Module 4.0 information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicS1000DDataModule4Info("TopicType")="Description-4-0"
'										dicS1000DDataModule4Info("Revision")="B"
'										dicS1000DDataModule4Info("Name")="DM42486"
'										dicS1000DDataModule4Info("MasterLanguageReference")="English US"
'										dicS1000DDataModule4Info("DocumentTitle")="Doc1"
'										dicS1000DDataModule4Info("ModelIdentifier")="MIC1"
'										dicS1000DDataModule4Info("SystemDifferenceCode")="SD1"
'										dicS1000DDataModule4Info("SystemCode")="SC1"
'										dicS1000DDataModule4Info("SubSystemCode")="SSC1"
'										dicS1000DDataModule4Info("SubSubSystemCode")="SSSC1"
'										dicS1000DDataModule4Info("AssemblyCode")="AC1"
'										dicS1000DDataModule4Info("DisassemblyCode")="DC6"
'										dicS1000DDataModule4Info("DisassemblyCodeVariant")="DVC6"
'										dicS1000DDataModule4Info("InformationCode")="IC6"
'										dicS1000DDataModule4Info("InformationCodeVariant")="ICV3"
'										dicS1000DDataModule4Info("ItemLocationCode")="IL4"
'										dicS1000DDataModule4Info("ExtensionCode")="EC24"
'										dicS1000DDataModule4Info("ExtensionProducer")="Proc1"
'										dicS1000DDataModule4Info("InWorkNumber")="IW1"
'										dicS1000DDataModule4Info("ExportFileName")="EFile1"
'										dicS1000DDataModule4Info("TechnicalName")="Tech"
'										dicS1000DDataModule4Info("InformationName")="Info1"
'										dicS1000DDataModule4Info("IssueNumber")="46"
'										dicS1000DDataModule4Info("IssueType")="Type2"
'										dicS1000DDataModule4Info("IssueDay")="11"
'										dicS1000DDataModule4Info("IssueMonth")="10"
'										dicS1000DDataModule4Info("IssueYear")="2010"
'										dicS1000DDataModule4Info("SecurityClass")="SC1"
'										dicS1000DDataModule4Info("ResponsiblePartnerCompany")="XXXX"
'										dicS1000DDataModule4Info("ResponsiblePartnerCompanyEnterpriseCode")="XX6"
'										dicS1000DDataModule4Info("OriginatorName")="ZZZ"
'										dicS1000DDataModule4Info("OriginatorEnterpriseCode")="YYY"
'										dicS1000DDataModule4Info("QualityAssurance")="QA"
'										dicS1000DDataModule4Info("SystemBreakdownCode")="BC32"
'										dicS1000DDataModule4Info("SkillLevel")="8"
'										dicS1000DDataModule4Info("ReasonForUpdate")="Not working"
'										dicS1000DDataModule4Info("Remarks")="Try"
'										bReturn=Fn_ContentM_CreateS1000DDataModule4(dicS1000DDataModule4Info)
'										
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												10-Apr-2012								1.0																						Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												17-Jan-2013								1.1					Modified function as per 10.1 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Shweta Rathod											11-May-2015								1.2					Modified function as per 11.2 design changes																	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreateS1000DDataModule4(dicS1000DDataModule4Info)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateS1000DDataModule4"
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateS1000DDataModule4=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting S1000D Data Module Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DDataModule4",objCMDialog, "ClassTree","Complete List:S1000D Data Module 4.0/4.1/4.2")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateS1000DDataModule4",objCMDialog)
	wait 3
	'Selecting topic type
	If dicS1000DDataModule4Info("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataModule4", objCMDialog, "SelectTopicType",dicS1000DDataModule4Info("TopicType"))
	End If
	'Setting Revision
	If dicS1000DDataModule4Info("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Revision", dicS1000DDataModule4Info("Revision"))
	End If
	'Setting Revision
	If dicS1000DDataModule4Info("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Name", dicS1000DDataModule4Info("Name"))
	End If
	'Setting Master Language Reference
	If dicS1000DDataModule4Info("MasterLanguageReference")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataDispatchNote4", objCMDialog, "SelectTopicType",dicS1000DDataModule4Info("MasterLanguageReference"))
'		Call Fn_Button_Click("Fn_ContentM_CreateS1000DPublicationModule4",objCMDialog,"MasterLanguageReference")
'		wait 1
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicS1000DPublicationModule4Info("MasterLanguageReference")
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
'		If bFlag=False Then
'			Set objCMDialog=Nothing
'			Set WshShell=Nothing
'			Exit Function
'		End If
	End If
	'Setting Document Title
	If dicS1000DDataModule4Info("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"DocumentTitle", dicS1000DDataModule4Info("DocumentTitle"))
	End If
	'Setting Model Identification Code
	If dicS1000DDataModule4Info("ModelIdentifier")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Model Identification:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"ModelIdentificationCode", dicS1000DDataModule4Info("ModelIdentifier"))
	End If
	'Setting SystemDifferenceCode
	If dicS1000DDataModule4Info("SystemDifferenceCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","System Difference Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"SystemDifferenceCode", dicS1000DDataModule4Info("SystemDifferenceCode"))
	End If
	'Setting System Code
	If dicS1000DDataModule4Info("SystemCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","System Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"SystemCode", dicS1000DDataModule4Info("SystemCode"))
	End If
	'Setting Sub System Code
	If dicS1000DDataModule4Info("SubSystemCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Sub-System Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"SubSystemCode", dicS1000DDataModule4Info("SubSystemCode"))
	End If
	'Setting Sub Sub-System Code
	If dicS1000DDataModule4Info("SubSubSystemCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Sub-subsystem Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"SubSubSystemCode", dicS1000DDataModule4Info("SubSubSystemCode"))
	End If
	'Setting Assembly Code
	If dicS1000DDataModule4Info("AssemblyCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Assembly Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"AssemblyCode", dicS1000DDataModule4Info("AssemblyCode"))
	End If
	'Setting DisassemblyCode
	If dicS1000DDataModule4Info("DisassemblyCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Disassembly Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"DisassemblyCode", dicS1000DDataModule4Info("DisassemblyCode"))
	End If
	'Setting DisassemblyCodeVariant
	If dicS1000DDataModule4Info("DisassemblyCodeVariant")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Disassembly Code Variant:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"DisassemblyCodeVariant", dicS1000DDataModule4Info("DisassemblyCodeVariant"))
	End If
	'Setting InformationCode
	If dicS1000DDataModule4Info("InformationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Information Code:"
'		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"InformationCode", dicS1000DDataModule4Info("InformationCode"))
		objCMDialog.JavaList("SelectTopicType").Type dicS1000DDataModule4Info("InformationCode")
	End If
	'Setting InformationCodeVariant
	If dicS1000DDataModule4Info("InformationCodeVariant")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Information Code Variant:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"InformationCodeVariant", dicS1000DDataModule4Info("InformationCodeVariant"))
	End If
	'Setting ItemLocationCode
	If dicS1000DDataModule4Info("ItemLocationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Item Location Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"ItemLocationCode", dicS1000DDataModule4Info("ItemLocationCode"))
	End If
	'Setting SupportEquipmentVariantCode
	If dicS1000DDataModule4Info("SupportEquipmentVariantCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Support Equipment Variant Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"SupportEquipmentVariantCode", dicS1000DDataModule4Info("SupportEquipmentVariantCode"))
	End If
	'Setting ExtensionCode
	If dicS1000DDataModule4Info("ExtensionCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Extension Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"ExtensionCode", dicS1000DDataModule4Info("ExtensionCode"))
	End If
	'Setting ExtensionProducer
	If dicS1000DDataModule4Info("ExtensionProducer")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Extension Producer:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"ExtensionProducer", dicS1000DDataModule4Info("ExtensionProducer"))
	End If
	'Setting In work number
	If dicS1000DDataModule4Info("InWorkNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","In Work Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"InWorkNumber", dicS1000DDataModule4Info("InWorkNumber"))
	End If
	'Setting ExportFileName
	If dicS1000DDataModule4Info("ExportFileName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Export File Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"ExportFileName", dicS1000DDataModule4Info("ExportFileName"))
	End If
	'Set Is This a Template option
	If dicS1000DDataModule4Info("IsThisATemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicS1000DDataModule4Info("IsThisATemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataModule4",objCMDialog, "IsThisATemplate")
	End If
	'Set Reference Only option
	If dicS1000DDataModule4Info("ReferenceOnly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicS1000DDataModule4Info("ReferenceOnly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataModule4",objCMDialog, "ReferenceOnly")
	End If
	'Setting TechnicalName
	If dicS1000DDataModule4Info("TechnicalName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Technical Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"TechnicalName", dicS1000DDataModule4Info("TechnicalName"))
	End If
	'Setting InformationName
	If dicS1000DDataModule4Info("InformationName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Information Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"InformationName", dicS1000DDataModule4Info("InformationName"))
	End If
	'Setting IssueNumber
	If dicS1000DDataModule4Info("IssueNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"IssueNumber", dicS1000DDataModule4Info("IssueNumber"))
	End If
	'Setting IssueType
	If dicS1000DDataModule4Info("IssueType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"IssueType", dicS1000DDataModule4Info("IssueType"))
	End If
	'Setting IssueDay
	If dicS1000DDataModule4Info("IssueDay")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Day:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"IssueDay", dicS1000DDataModule4Info("IssueDay"))
	End If
	'Setting IssueMonth
	If dicS1000DDataModule4Info("IssueMonth")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Month:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"IssueMonth", dicS1000DDataModule4Info("IssueMonth"))
	End If
	'Setting IssueYear
	If dicS1000DDataModule4Info("IssueYear")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Year:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"IssueYear", dicS1000DDataModule4Info("IssueYear"))
	End If
	'Setting SecurityClass
	If dicS1000DDataModule4Info("SecurityClass")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Security Class:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"SecurityClass", dicS1000DDataModule4Info("SecurityClass"))
	End If
	'Setting ResponsiblePartnerCompany
	If dicS1000DDataModule4Info("ResponsiblePartnerCompany")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Responsible Partner Company Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"ResponsiblePartnerCompany", dicS1000DDataModule4Info("ResponsiblePartnerCompany"))
	End If
	'Setting ResponsiblePartnerCompanyEnterpriseCode
	If dicS1000DDataModule4Info("ResponsiblePartnerCompanyEnterpriseCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Responsible Partner Company Enterprise Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"ResponsiblePartnerCompanyEnterpriseCode", dicS1000DDataModule4Info("ResponsiblePartnerCompanyEnterpriseCode"))
	End If
	'Setting OriginatorName
	If dicS1000DDataModule4Info("OriginatorName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Originator", dicS1000DDataModule4Info("OriginatorName"))
	End If
	'Setting OriginatorEnterpriseCode
	If dicS1000DDataModule4Info("OriginatorEnterpriseCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Enterprise Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Originator", dicS1000DDataModule4Info("OriginatorEnterpriseCode"))
	End If
	'Setting Caveat
	If dicS1000DDataModule4Info("Caveat")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Caveat:"
		objCMDialog.JavaButton("DropDownButton").Highlight
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule4", objCMDialog,"DropDownButton")
		Wait 2
		Call Fn_UI_ClickJavaTreeCell("Fn_ContentM_CreateS1000DDataModule4", objCMDialog,"ClassTree",dicS1000DDataModule4Info("Caveat"),"Value","RIGHT")
	End If
	'Setting CommercialClassification
	If dicS1000DDataModule4Info("CommercialClassification")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Commercial Classification:"
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule4", objCMDialog,"DropDownButton")
		Wait 2
		Call Fn_UI_ClickJavaTreeCell("Fn_ContentM_CreateS1000DDataModule4", objCMDialog,"ClassTree",dicS1000DDataModule4Info("CommercialClassification"),"Value","RIGHT")
	End If
    'Setting QualityAssurance
	If dicS1000DDataModule4Info("QualityAssurance")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Quality Assurance:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"QualityAssurance", dicS1000DDataModule4Info("QualityAssurance"))
	End If
	'Setting SystemBreakdownCode
	If dicS1000DDataModule4Info("SystemBreakdownCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","System Breakdown Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"SystemBreakdownCode", dicS1000DDataModule4Info("SystemBreakdownCode"))
	End If
	'Setting Skill
	If dicS1000DDataModule4Info("SkillLevel")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Skill Level:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Skill", dicS1000DDataModule4Info("SkillLevel"))
	End If
	'Setting ReasonForUpdate
	If dicS1000DDataModule4Info("ReasonForUpdate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reason For Update:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"ReasonForUpdate", dicS1000DDataModule4Info("ReasonForUpdate"))
	End If
	'Setting Remarks
	If dicS1000DDataModule4Info("Remarks")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Remarks:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Remarks", dicS1000DDataModule4Info("Remarks"))
	End If
    If dicS1000DDataModule4Info("LearnCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Learn Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"DocumentTitle", dicS1000DDataModule4Info("LearnCode"))
	End If
	If dicS1000DDataModule4Info("LearnEventCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Learn Event Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"DocumentTitle", dicS1000DDataModule4Info("LearnEventCode"))
	End If
    'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Finish")
	wait 2
	If dicS1000DDataModule4Info("ReferenceTopic")<>"" Then
		JavaWindow("ContentManagement").JavaWindow("Reference Topic Type Selection").JavaList("ReferenceTopicType").Select dicS1000DDataModule4Info("ReferenceTopic")
		JavaWindow("ContentManagement").JavaWindow("Reference Topic Type Selection").JavaButton("Finish").Click
	End If
	
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule4",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateS1000DDataModule4=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateS1000DCommentary4

'Description			 :	Function Used to create S1000D Commentary 4.0

'Parameters			   :   '1.dicCommentary4info: S1000D Commentary 4.0 information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicCommentary4info("TopicType")="Comment-4-0"
'										dicCommentary4info("Revision")="C"
'										dicCommentary4info("Name")="Commentary24"
'										dicCommentary4info("MasterLanguageReference")="English US"
'										dicCommentary4info("DocumentTitle")="Commentary 1"
'										dicCommentary4info("ModelIdentifier")="MI1"
'										dicCommentary4info("SenderIdentificationCode")="SIC1"
'										dicCommentary4info("YearOfDataIssue")="2011"
'										dicCommentary4info("SequentialNumber")="6"
'										dicCommentary4info("CommentType")="CType2"
'										dicCommentary4info("IssueType")="Type1"
'										dicCommentary4info("IssueDay")="11"
'										dicCommentary4info("IssueMonth")="11"
'										dicCommentary4info("IssueYear")="2010"
'										dicCommentary4info("IssueNumber")="6"
'										dicCommentary4info("CommentPriorityCode")="CPC6"
'										dicCommentary4info("Remarks")="new remark"
'										dicCommentary4info("InWorkNumber")="86"
'										dicCommentary4info("CommentResponceType")="RST5"
'										dicCommentary4info("SecurityClass")="SC1"
'										dicCommentary4info("ExportFileName")="File4"
'										dicCommentary4info("DispatchPersonFirstName")="aaa"
'										dicCommentary4info("DispatchPersonSurname")="zzz"
'										dicCommentary4info("DispatchPersonJobTitle")="Job1"
'										dicCommentary4info("OriginatorEmailAddress")="Address1"
'										dicCommentary4info("OriginatorDispatchAddressPhone")="111111"
'										dicCommentary4info("OriginatorDispatchAddressFax")="FXXXX"
'										dicCommentary4info("OriginatorInternetAddress")="abc@xyz.com"
'										dicCommentary4info("OriginatorDispatchAddressDepartment")="Dept1"
'										dicCommentary4info("OriginatorDispatchAddressBuilding")="B1"
'										dicCommentary4info("OriginatorDispatchAddressRoom")="R1"
'										dicCommentary4info("OriginatorDispatchAddressStreet")="S1"
'										dicCommentary4info("OriginatorDispatchAddressPostOfficeBox")="PO23"
'										dicCommentary4info("OriginatorDispatchAddressCity")="City1"
'										dicCommentary4info("OriginatorDispatchAddressState")="State1"
'										dicCommentary4info("OriginatorDispatchAddressZipCode")="411"
'										dicCommentary4info("OriginatorDispatchAddressProvince")="P3"
'										dicCommentary4info("OriginatorDispatchAddressPostCode")="567"
'										dicCommentary4info("OriginatorDispatchAddressCountry")="Country1"
'										bReturn=Fn_ContentM_CreateS1000DCommentary4(dicCommentary4info)
'										
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												11-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreateS1000DCommentary4(dicCommentary4info)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateS1000DCommentary4"
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateS1000DCommentary4=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting S1000D Data Module Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DCommentary4",objCMDialog, "ClassTree","Complete List:S1000D Commentary 4.0/4.1/4.2")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateS1000DCommentary4",objCMDialog)
	wait 3
	'Selecting topic type
	If dicCommentary4info("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DCommentary4", objCMDialog, "SelectTopicType",dicCommentary4info("TopicType"))
	End If
	'Setting Revision
	If dicCommentary4info("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"Revision", dicCommentary4info("Revision"))
	End If
	'Setting Revision
	If dicCommentary4info("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"Name", dicCommentary4info("Name"))
	End If
	'Setting Master Language Reference
	If dicCommentary4info("MasterLanguageReference")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DCommentary4", objCMDialog, "SelectTopicType",dicCommentary4info("MasterLanguageReference"))
'		Call Fn_Button_Click("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"MasterLanguageReference")
'		wait 1
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicCommentary4info("MasterLanguageReference")
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
'		If bFlag=False Then
'			Set objCMDialog=Nothing
'			Set WshShell=Nothing
'			Exit Function
'		End If
	End If
	'Setting Document Title
	If dicCommentary4info("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"DocumentTitle", dicCommentary4info("DocumentTitle"))
	End If
	'Setting Model Identification Code
	If dicCommentary4info("ModelIdentifier")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Model Identifier:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"ModelIdentificationCode", dicCommentary4info("ModelIdentifier"))
	End If
	'Setting SenderIdentificationCode
	If dicCommentary4info("SenderIdentificationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Sender Identification Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"SenderIdentificationCode", dicCommentary4info("SenderIdentificationCode"))
	End If
	'Setting YearOfDataIssue
	If dicCommentary4info("YearOfDataIssue")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Year Of Data Issue:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"YearOfDataIssue", dicCommentary4info("YearOfDataIssue"))
	End If
	'Setting SequentialNumber
	If dicCommentary4info("SequentialNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Sequential Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"SequentialNumber", dicCommentary4info("SequentialNumber"))
	End If
	'Setting CommentType
	If dicCommentary4info("CommentType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Comment Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"CommentType", dicCommentary4info("CommentType"))
	End If
	'Setting IssueType
	If dicCommentary4info("IssueType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"IssueType", dicCommentary4info("IssueType"))
	End If
	'Setting IssueDay
	If dicCommentary4info("IssueDay")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Day:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"IssueDay", dicCommentary4info("IssueDay"))
	End If
	'Setting IssueMonth
	If dicCommentary4info("IssueMonth")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Month:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"IssueMonth", dicCommentary4info("IssueMonth"))
	End If
	'Setting IssueYear
	If dicCommentary4info("IssueYear")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Year:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"IssueYear", dicCommentary4info("IssueYear"))
	End If
	'Setting IssueNumber
	If dicCommentary4info("IssueNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"IssueNumber", dicCommentary4info("IssueNumber"))
	End If
	'Setting CommentPriorityCode
	If dicCommentary4info("CommentPriorityCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Comment Priority Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"CommentPriorityCode", dicCommentary4info("CommentPriorityCode"))
	End If
	'Setting Remarks
	If dicCommentary4info("CommentPriorityCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Remarks:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"Remarks", dicCommentary4info("Remarks"))
	End If
	'Setting In work number
	If dicCommentary4info("InWorkNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","In Work Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"InWorkNumber", dicCommentary4info("InWorkNumber"))
	End If
	'Setting CommentResponceType
	If dicCommentary4info("CommentResponceType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Comment Response Type:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"CommentType", dicCommentary4info("CommentResponceType"))
	End If
	'Setting SecurityClass
	If dicCommentary4info("SecurityClass")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Security Class:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"SecurityClass", dicCommentary4info("SecurityClass"))
	End If
	'Setting ExportFileName
	If dicCommentary4info("ExportFileName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Export File Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"ExportFileName", dicCommentary4info("ExportFileName"))
	End If
	'Set Is This a Template option
	If dicCommentary4info("Isthisatemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicCommentary4info("Isthisatemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DCommentary4",objCMDialog, "IsThisATemplate")
	End If
	'Set Reference Only option
	If dicCommentary4info("Referenceonly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicCommentary4info("Referenceonly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DCommentary4",objCMDialog, "ReferenceOnly")
	End If
	'Setting DispatchPersonFirstName
	If dicCommentary4info("DispatchPersonFirstName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch Person First Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("DispatchPersonFirstName"))
	End If
	'Setting DispatchPersonSurname
	If dicCommentary4info("DispatchPersonSurname")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch Person Surname:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("DispatchPersonSurname"))
	End If
	'Setting DispatchPersonJobTitle
	If dicCommentary4info("DispatchPersonJobTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Dispatch Person Job Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("DispatchPersonJobTitle"))
	End If
	'Setting OriginatorEmailAddress
	If dicCommentary4info("OriginatorEmailAddress")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Email Address:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorEmailAddress"))
	End If
	'Setting OriginatorDispatchAddressPhone
	If dicCommentary4info("OriginatorDispatchAddressPhone")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Phone:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressPhone"))
	End If
	'Setting OriginatorDispatchAddressFax
	If dicCommentary4info("OriginatorDispatchAddressFax")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Fax:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressFax"))
	End If
	'Setting OriginatorInternetAddress
	If dicCommentary4info("OriginatorInternetAddress")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Internet Address:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorInternetAddress"))
	End If
	'Setting OriginatorDispatchAddressDepartment
	If dicCommentary4info("OriginatorDispatchAddressDepartment")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Department:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressDepartment"))
	End If
	'Setting OriginatorDispatchAddressBuilding
	If dicCommentary4info("OriginatorDispatchAddressBuilding")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Building:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressBuilding"))
	End If
	'Setting OriginatorDispatchAddressRoom
	If dicCommentary4info("OriginatorDispatchAddressRoom")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Room:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressRoom"))
	End If
	'Setting OriginatorDispatchAddressStreet
	If dicCommentary4info("OriginatorDispatchAddressStreet")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Street:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressStreet"))
	End If
	'Setting OriginatorDispatchAddressPostOfficeBox
	If dicCommentary4info("OriginatorDispatchAddressPostOfficeBox")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Post Office Box:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressPostOfficeBox"))
	End If
	'Setting OriginatorDispatchAddressCity
	If dicCommentary4info("OriginatorDispatchAddressCity")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address City:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressCity"))
	End If
	'Setting OriginatorInternetAddressState
	If dicCommentary4info("OriginatorDispatchAddressState")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address State:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressState"))
	End If
	'Setting OriginatorDispatchAddressZipCode
	If dicCommentary4info("OriginatorDispatchAddressZipCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Zip Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressZipCode"))
	End If
	'Setting OriginatorDispatchAddressProvince
	If dicCommentary4info("OriginatorDispatchAddressProvince")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Province:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressProvince"))
	End If
	'Setting OriginatorDispatchAddressPostCode
	If dicCommentary4info("OriginatorDispatchAddressPostCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Post Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressPostCode"))
	End If
	'Setting OriginatorDispatchAddressCountry
	If dicCommentary4info("OriginatorDispatchAddressCountry")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator Dispatch Address Country:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"AuthorEdit", dicCommentary4info("OriginatorDispatchAddressCountry"))
	End If
    'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DCommentary4",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateS1000DCommentary4=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateTranslationOrder

'Description			 :	Function Used to create Translation Order

'Parameters			   :   1.StrID: Translation Order ID
'										2.StrRevision: Revision
'										3.StrName: Name
'										4.StrOrderTitle: Order Title
'										5.StrOrderDescription: Order Description
'										6.StrTranslationOfficeReference : Translation Office Reference
'										7.RequestDeliveryDate: Request Delivery Date
'
'Return Value		   : 	True or False

'Pre-requisite			:	DITA Topic revision should be selected

'Examples				:  	bReturn=Fn_ContentM_CreateTranslationOrder("000999","A","TOrder24","New Order","New orderDescription","TOffice","11-Apr-2012")

'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												11-Apr-2012								1.0																						Sunny R
'												Shwetambari Rathod											03-Jul-2014													added code to handle date control object as per design change.														
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateTranslationOrder(StrID,StrRevision,StrName,StrOrderTitle,StrOrderDescription,StrTranslationOfficeReference,RequestDeliveryDate)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateTranslationOrder"
	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter,arrDate
	Fn_ContentM_CreateTranslationOrder=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting TranslationOrder Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateTranslationOrder",objCMDialog, "ClassTree","Complete List:Translation Order")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateTranslationOrder",objCMDialog,"Next")
	'Maximizing [ NewAuthorClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateTranslationOrder",objCMDialog)
	wait 3
	If StrID<>"" Then
		'Setting ID for Translation Order
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOrder",objCMDialog,"ID",StrID)
	End If
	If StrRevision<>"" Then
		'Setting Revision for Translation Order
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOrder",objCMDialog,"Revision",StrRevision)
	End If
    	If StrName<>"" Then
		'Setting Name for Translation Order
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOrder",objCMDialog,"Name",StrName)
	End If
	If StrOrderTitle<>"" Then
		'Setting Order Title for Translation Order
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Order Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOrder",objCMDialog,"AuthorEdit",StrOrderTitle)
	End If
	If StrOrderDescription<>"" Then
		'Setting Order Description for Translation Order
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Order Description:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOrder",objCMDialog,"AuthorEdit",StrOrderDescription)
	End If
	'Setting Translation Office Reference
	If StrTranslationOfficeReference<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Translation Office Reference:"
		Call Fn_List_Select("Fn_ContentM_CreateTranslationOrder", objCMDialog, "JavaList",StrTranslationOfficeReference)
	End If
	
	'' settting Request Delivery Date:
	If RequestDeliveryDate<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Request Delivery Date:"
		arrDate=Split(RequestDeliveryDate," ")
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOrder",objCMDialog,"ID",arrDate(0))
		Call Fn_KeyBoardOperation("SendKey", "{TAB}")
	End If

	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateTranslationOrder",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateTranslationOrder",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateTranslationOrder=True
	Set objCMDialog=Nothing	
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_LanguageTableGetCellData

'Description			 :	Function Used to Get Cell data from Language Table on summary tab

'Parameters			   :   '1.objTable: Table Object
'										 2.iRow : row number
'										 3.iCol: column number or name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Language Table Should activated

'Examples				:  	bReturn=Fn_ContentM_LanguageTableGetCellData(JavaWindow("ContentManagement").JavaTable("LanguageTable"),1,"Review Ordered")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												12-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_LanguageTableGetCellData(objTable,iRow,iCol)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_LanguageTableGetCellData"
   Dim sColName,sPropName
   Fn_ContentM_LanguageTableGetCellData=False
	If IsNumeric(iCol) Then
		sColName = objTable.Object.getColumn(iCol).getText()
	Else
		sColName = iCol
	End If
	
	Select Case Trim(sColName)
		Case "Language Reference"
			sPropName="fnd0LanguageTagref"
		Case "Review Ordered"
			sPropName="reviewOrdered"
		Case Else
			Exit function
	End Select
	Fn_ContentM_LanguageTableGetCellData=objTable.Object.getItem(iRow).getData().getComponent().getProperty(sPropName)
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_LanguageTableOperation

'Description			 :	Function Used to perform operations on Language Table on summary tab

'Parameters			   :   1.StrAction: Action name
'										2.StrLanguageReference : Language Reference
'										3.bReviewOrdered: Review Ordered
'										4.StrColName: Column name
'										5.StrValue: cell value
'										6.StrPopupmenu: popup menu
'
'Return Value		   : 	True or False

'Pre-requisite			:	Language Table Should activated

'Examples				:  	bReturn=Fn_ContentM_LanguageTableOperation("Select","Czech","","","","")
'										bReturn=Fn_ContentM_LanguageTableOperation("CellVerify","English UK","","Review Ordered","Y","")
'										bReturn=Fn_ContentM_LanguageTableOperation("CellVerify","English UK","","Language Reference","English UK","")
'										bReturn=Fn_ContentM_LanguageTableOperation("Delete","Czech","","","","")
'										bReturn=Fn_ContentM_LanguageTableOperation("Add","English UK","True","","","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												12-Apr-2012								1.0																						Sunny R
'													Sandeep N												31-Aug-2012								1.1																						Pallavi J
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_LanguageTableOperation(StrAction,StrLanguageReference,bReviewOrdered,StrColName,StrValue,StrPopupmenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_LanguageTableOperation"
	'Declaring Variables
	Dim objLanguageTable,bFlag,iRowCount,iCounter
	Dim objTable,objChild,iRow,WshShell

	Fn_ContentM_LanguageTableOperation=False
	Set objLanguageTable=JavaWindow("ContentManagement").JavaTable("LanguageTable")
	If objLanguageTable.Exist(6) Then
	   Select Case StrAction
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to Add Language Reference from Language Table
			Case "Add"
				Set WshShell = CreateObject("WScript.Shell")

				bFlag=True
				If Not JavaWindow("ContentManagement").JavaWindow("LanguagesTable").Exist(6) Then
'					bFlag=Fn_ToolBarOperation("Click", "Add...", "" )
					bFlag = Fn_Button_Click("Fn_ContentM_LanguageTableOperation",JavaWindow("ContentManagement"), "Add")
				End If
				If bFlag=True Then
					'Selecting Language Reference
					Call Fn_Button_Click("Fn_ContentM_LanguageTableOperation",JavaWindow("ContentManagement").JavaWindow("LanguagesTable"),"LanguageReference")
					wait 2
					bFlag=False
					wait 1
					WshShell.SendKeys "{TAB}"
					wait 1
					WshShell.SendKeys "{DOWN}"
					wait 1
					If JavaWindow("ContentManagement").JavaWindow("LanguagesTable").JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
						JavaWindow("ContentManagement").JavaWindow("LanguagesTable").JavaWindow("TreeShell").JavaTree("Tree").Activate StrLanguageReference
						wait 2
						bFlag=true
						If JavaWindow("ContentManagement").JavaWindow("LanguagesTable").JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
							bFlag=False
						End If
					Else
						bFlag=False
					End If
					If bFlag=False Then
						Set WshShell = Nothing
						Set objLanguageTable=Nothing
						Exit Function
					End If
					'Setting Review Ordered option
					' Added by shweta on 18-Jun-15 as per the discussion with akshay make default value to false if bReviewOrdered is balnk 
					If bReviewOrdered = "" Then
						bReviewOrdered = "False"
					End if	
					If  bReviewOrdered<>"" Then
						JavaWindow("ContentManagement").JavaWindow("LanguagesTable").JavaRadioButton("ReviewOrdered").SetTOProperty "attached text",bReviewOrdered
						Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_LanguageTableOperation",JavaWindow("ContentManagement").JavaWindow("LanguagesTable"), "ReviewOrdered")
					End If
					Call Fn_Button_Click("Fn_ContentM_LanguageTableOperation",JavaWindow("ContentManagement").JavaWindow("LanguagesTable"),"Finish")
					wait 2
					If JavaWindow("ContentManagement").JavaWindow("LanguagesTable").Exist(6) Then
						Call Fn_Button_Click("Fn_ContentM_LanguageTableOperation",JavaWindow("ContentManagement").JavaWindow("LanguagesTable"),"Cancel")
					End If
				End If
				Fn_ContentM_LanguageTableOperation=True
				Set WshShell = Nothing
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to delete Language Reference from Language Table
			Case "Delete"
'				bFlag=Fn_ContentM_LanguageTableOperation("Select",StrLanguageReference,"","","","")
				bFlag=False
				iRowCount=objLanguageTable.GetROProperty("rows")
				For iCounter=0 to iRowCount-1
					If trim(StrLanguageReference)=trim(objLanguageTable.Object.getItem(iCounter).getData().getComponent().getProperty("fnd0LanguageTagref")) Then
						wait 2
						objLanguageTable.ClickCell iCounter,0
						wait 1
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=True Then
                  '  JavaWindow("ContentManagement").JavaToolbar("LanguageTableAddAndDelete").Press "Delete"
                  Call Fn_Button_Click("Fn_ContentM_LanguageTableOperation",JavaWindow("ContentManagement"), "Delete")
					wait 1
					If Dialog("Warning").Exist(6) Then
						wait 1
						Dialog("Warning").WinButton("Yes").Click
					End If
					Call  Fn_ReadyStatusSync(1)
					Fn_ContentM_LanguageTableOperation=True
				End If
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			'Case to select row from Language Table
			Case "Select"
				iRowCount=objLanguageTable.GetROProperty("rows")
				For iCounter=0 to iRowCount-1
					If trim(StrLanguageReference)=trim(objLanguageTable.Object.getItem(iCounter).getData().getComponent().getProperty("fnd0LanguageTagref")) Then
						objLanguageTable.SelectRow iCounter
						wait 1
						Fn_ContentM_LanguageTableOperation=True
						Exit For
					End If
				Next
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "CellVerify"
				iRowCount=objLanguageTable.GetROProperty("rows")
				For iCounter=0 to iRowCount-1
					If trim(StrLanguageReference)=trim(objLanguageTable.Object.getItem(iCounter).getData().getComponent().getProperty("fnd0LanguageTagref")) Then
						If Fn_ContentM_LanguageTableGetCellData(objLanguageTable,iCounter,StrColName)=StrValue Then
							Fn_ContentM_LanguageTableOperation=True
							Exit For
						End If
					End If
				Next
	   End Select
   End If
   Set objLanguageTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateTranslationOffice

'Description			 :	Function Used to create Translation Office

'Parameters			   :   '1.dicTranslationOfficeInfo: Translation Office information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicTranslationOfficeInfo("ID")="000999"
'										dicTranslationOfficeInfo("Revision")="S"
'										dicTranslationOfficeInfo("Name")="TOffice6"
'										dicTranslationOfficeInfo("TranslationOfficeTitle")="Title1"
'										dicTranslationOfficeInfo("Address")="xxxxxxxxxxx"	
'										dicTranslationOfficeInfo("ContactName")="Name1"	
'										dicTranslationOfficeInfo("Phone")="1111111111"
'										dicTranslationOfficeInfo("Website")="abc@xyz.com"
'										dicTranslationOfficeInfo("EmailInbox")="Inbox1"
'										dicTranslationOfficeInfo("DeliverComposedContent")="True"
'										dicTranslationOfficeInfo("DeliverDecomposedContent")="True"
'										dicTranslationOfficeInfo("IncludeSupportingData")="True"
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												12-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateTranslationOffice(dicTranslationOfficeInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateTranslationOffice"
   'Declaring variables
    Dim objCMDialog,StrMenu
	Fn_ContentM_CreateTranslationOffice=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	'Checking existance of [ NewFolder ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting Translation Office Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateTranslationOffice",objCMDialog, "ClassTree","Complete List:Translation Office")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateTranslationOffice",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateTranslationOffice",objCMDialog)
	wait 3
	'Setting ID
	If dicTranslationOfficeInfo("ID")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("ID"))
	End If
	'Setting Revision
	If dicTranslationOfficeInfo("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("Revision"))
	End If
	'Setting Name
	If dicTranslationOfficeInfo("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("Name"))
	End If
	'Setting Translation Office Title
	If dicTranslationOfficeInfo("TranslationOfficeTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Translation Office Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("TranslationOfficeTitle"))
	End If
	'Setting Address
	If dicTranslationOfficeInfo("Address")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Address:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("Address"))
	End If
	'Setting ContactName
	If dicTranslationOfficeInfo("ContactName")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Contact Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("ContactName"))
	End If
	'Setting Phone
	If dicTranslationOfficeInfo("Phone")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Phone Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("Phone"))
	End If
	'Setting Website
	If dicTranslationOfficeInfo("Website")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Website:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("Website"))
	End If
	'Setting Email Inbox
	If dicTranslationOfficeInfo("EmailInbox")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Email Inbox:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTranslationOffice",objCMDialog,"AdministrativeEdit",dicTranslationOfficeInfo("EmailInbox"))
	End If
	'Selecting Deliver Composed Content
	If dicTranslationOfficeInfo("DeliverComposedContent")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Deliver Composed Content:"
		objCMDialog.JavaRadioButton("AdministrativeRadioButton").SetTOProperty "attached text",dicTranslationOfficeInfo("DeliverComposedContent")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateTranslationOffice",objCMDialog, "AdministrativeRadioButton")
	End If
	'Selecting Deliver Decomposed Content
	If dicTranslationOfficeInfo("DeliverDecomposedContent")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Deliver Decomposed Content:"
		objCMDialog.JavaRadioButton("AdministrativeRadioButton").SetTOProperty "attached text",dicTranslationOfficeInfo("DeliverDecomposedContent")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateTranslationOffice",objCMDialog, "AdministrativeRadioButton")
	End If
	'Selecting Include Supporting Data
	If dicTranslationOfficeInfo("IncludeSupportingData")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Include Supporting Data:"
		objCMDialog.JavaRadioButton("AdministrativeRadioButton").SetTOProperty "attached text",dicTranslationOfficeInfo("IncludeSupportingData")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateTranslationOffice",objCMDialog, "AdministrativeRadioButton")
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateTranslationOffice",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateTranslationOffice",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateTranslationOffice=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Translation Office of name [" + dicTranslationOfficeInfo("Name") + "]")
	Set objCMDialog=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateXMLSchema

'Description			 :	Function Used to create XML Schema

'Parameters			   :   1.StrID: XML Schema ID
'										2.StrRevision : XML Schema Revision
'										3.StrName: Name
'										4.StrPublicIdentifier: Public Identifier
'										5.StrDefaultPrefix: Default Prefix
'										6.StrSchemaLocation: Schema Location
'										7.StrContentPath: Content Path
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be present in Content Management perspective

'Examples				:  	bReturn=Fn_ContentM_CreateXMLSchema("X2486","X","XMLSchema1","PI6","D","","C:\tc91\TestData\ContentTestData\ContentXMLScemaData.txt")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateXMLSchema(StrID,StrRevision,StrName,StrPublicIdentifier,StrDefaultPrefix,StrSchemaLocation,StrContentPath)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateXMLSchema"
    'Declaring variables
    Dim objCMDialog,StrMenu
	Fn_ContentM_CreateXMLSchema=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	'Checking existance of [ NewAdministrativeClass ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting XML Schema Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateXMLSchema",objCMDialog, "ClassTree","Complete List:XML Schema")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateXMLSchema",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateXMLSchema",objCMDialog)
	wait 3
	'Setting ID
	If StrID<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateXMLSchema",objCMDialog,"AdministrativeEdit",StrID)
	End If
	'Setting Revision
	If StrRevision<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateXMLSchema",objCMDialog,"AdministrativeEdit",StrRevision)
	End If
	'Setting Name
	If StrName<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateXMLSchema",objCMDialog,"AdministrativeEdit",StrName)
	End If
	'Setting Public Identifier
	If StrRevision<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Public ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateXMLSchema",objCMDialog,"AdministrativeEdit",StrPublicIdentifier)
	End If
	'Setting Default Prefix
	If StrDefaultPrefix<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Default Prefix:"
		Call Fn_Edit_Box("Fn_ContentM_CreateXMLSchema",objCMDialog,"AdministrativeEdit",StrDefaultPrefix)
	End If
	'Setting Schema Location
	If StrSchemaLocation<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Schema Location:"
		Call Fn_Edit_Box("Fn_ContentM_CreateXMLSchema",objCMDialog,"AdministrativeEdit",StrSchemaLocation)
	End If
	'Taking content from file
	If StrContentPath<>"" Then
		'Press Browse button
		Call Fn_Button_Click("Fn_ContentM_CreateXMLSchema",objCMDialog,"Browse")
		If objCMDialog.Dialog("Browse").Exist(6) Then
			objCMDialog.Dialog("Browse").WinEdit("FileName").Set StrContentPath
			wait 2
			objCMDialog.Dialog("Browse").WinButton("Open").Click
			wait 1
		Else
			Set objCMDialog=Nothing
			Exit Function
		End If
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateXMLSchema",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateXMLSchema",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateXMLSchema=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created XML Schema of name [" + StrName + "]")
	Set objCMDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_SpecifySearchDetailsAndInvoke

'Description			 :	Function Used to search object by using criteria

'Parameters			   :   1.dicContentSearchDetails: Search criteria details
'
'
'Return Value		   : 	True or False

'Pre-requisite			:	Need to declare and define dictionary object [ dicContentSearchDetails ] in test case only
'										Example of how to declare dictionary
'
'										Dim dicContentSearchDetails
'										Set dicContentSearchDetails=CreateObject( "Scripting.Dictionary" )
'										With dicContentSearchDetails
'											.Add "Public ID",""
'										End With
'										
'Examples				:  	dicContentSearchDetails("Public ID")="ICN-CORT3-00777-04-1*"
'										bReturn=Fn_ContentM_SpecifySearchDetailsAndInvoke(dicContentSearchDetails)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Apr-2012								1.0																						Sunny R
'													Sandeep N												17-Apr-2012								1.1				Added case : "Schema Type","Apply Class Name","Name"	Swati K
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_SpecifySearchDetailsAndInvoke(dicContentSearchDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_SpecifySearchDetailsAndInvoke"
   'Variable Declaration
	Dim objCMDialog,bFlag,DictItems,DictKeys,iCounter,StrAction
	Dim objSelectType,objIntNoOfObjects
	Dim sValChar,iAsc,iValAsc,iLastItem

	Fn_ContentM_SpecifySearchDetailsAndInvoke=False
	'Creating object of [ ContentManagement ] window
	Set objCMDialog=JavaWindow("ContentManagement")
	'Checking existance of [ ContentManagement ] window
	If objCMDialog.Exist Then
		'Clearing all search fields
		bFlag= Fn_ToolbatButtonClick("Clear all search fields")
		If bFlag=False Then
			Exit Function
		End If
		'taking the keys & items count from data dictionary
		DictItems = dicContentSearchDetails.Items
		DictKeys = dicContentSearchDetails.Keys

		For iCounter=0 to dicContentSearchDetails.count-1
			If DictItems(iCounter)<>"" Then
				StrAction=DictKeys(iCounter)
				Select Case StrAction
					 '- - - - - - - - - - - - - -  Edit Box
					Case "Public ID","Extension Producer","Name","Style Sheet Type","Style Sheet Main File","Topic Type Reference","Model Identification Code"

						objCMDialog.JavaStaticText("srch_Type").SetTOProperty "label", StrAction+":"
						If objCMDialog.JavaEdit("srch_EditBox").Exist(2) Then
							Call Fn_Edit_Box("Fn_ContentM_SpecifySearchDetailsAndInvoke", objCMDialog,"srch_EditBox",DictItems(iCounter))
							wait(3)
						Else
							Set objCMDialog=Nothing
							Exit function
						End If
					'- - - - - - - - - - - - - -  Drop Down list
					Case "Schema Type","Apply Class Name"
						objCMDialog.JavaStaticText("srch_Type").SetTOProperty "label", StrAction+":"
						Call Fn_Button_Click( "Fn_ContentM_SpecifySearchDetailsAndInvoke", objCMDialog,"srch_MultipleDropDown" )
						wait 1
						sValChar=Mid(DictItems(iCounter),1,1)
						iAsc=Asc(sValChar)
						iValAsc=0
						Do While iValAsc<iAsc
							iLastItem=objCMDialog.JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").GetROProperty("items count")
							iLastItem=iLastItem-1
							sValue=objCMDialog.JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").GetItem(iLastItem)
							sChar=Mid(sValue,1,1)
							iValAsc=Asc(sChar)	
							objCMDialog.JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").Select "#"&Cstr(iLastItem)
							wait 1
						Loop
						objCMDialog.JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").Activate  DictItems(iCounter)
						wait 1
						If not JavaWindow("MyTeamcenter").JavaWindow("SearchCriteriaTreeShell").JavaTree("Tree").Exist(3) Then
							bFlag = True
						End If

					Case Else
						Set objCMDialog=Nothing
						Exit function
				End Select
			End If
		Next

	End If
	bFlag= Fn_ToolbatButtonClick("Executes the search and displays the results in search result view")
	If bFlag=True Then
		Fn_ContentM_SpecifySearchDetailsAndInvoke = True
	End If
	Set objCMDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateStylesheet

'Description			 :	Function Used to create Stylesheet

'Parameters			   :   1.StrID: Stylesheet ID
'										2.StrRevision : Stylesheet Revision
'										3.StrName: Name
'										4.StrPublicIdentifier: Public Identifier
'										5.StrStylesheetType: Stylesheet Type
'										6.StrStylesheetMainFile: Stylesheet Main File ( for zipped )
'										7.StrStylesheetResultingContentType: Stylesheet Resulting Content Type
'										8.StrANTTarget: ANT Target
'										9.StrContentPath: Content Path
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be present in Content Management perspective

'Examples				:  	bReturn=Fn_ContentM_CreateStylesheet("000998","A","Stylesheet1","PI24","EDITOR_VIEW","","XHTML","","C:\tc91\TestData\ContentTestData\ContentXMLScemaData.txt")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												16-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateStylesheet(StrID,StrRevision,StrName,StrPublicIdentifier,StrStylesheetType,StrStylesheetMainFile,StrStylesheetResultingContentType,StrANTTarget,StrContentPath)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateStylesheet"
    'Declaring variables
    Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateStylesheet=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking existance of [ NewAdministrativeClass ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting Stylesheet Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateStylesheet",objCMDialog, "ClassTree","Complete List:Stylesheet")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateStylesheet",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateStylesheet",objCMDialog)
	wait 3
	'Setting ID
	If StrID<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateStylesheet",objCMDialog,"AdministrativeEdit",StrID)
	End If
	'Setting Revision
	If StrRevision<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateStylesheet",objCMDialog,"AdministrativeEdit",StrRevision)
	End If
	'Setting Name
	If StrName<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateStylesheet",objCMDialog,"AdministrativeEdit",StrName)
	End If
	'Setting Public Identifier
	If StrRevision<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Public ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateStylesheet",objCMDialog,"AdministrativeEdit",StrPublicIdentifier)
	End If
	'Selecting Stylesheet Type
	If StrStylesheetType<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Stylesheet Type:"
		Call Fn_Button_Click("Fn_ContentM_CreateStylesheet",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrStylesheetType
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting Stylesheet Main File ( for zipped )
 	If StrStylesheetMainFile<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Stylesheet Main File (for Zipped):"
		Call Fn_Edit_Box("Fn_ContentM_CreateStylesheet",objCMDialog,"AdministrativeEdit",StrStylesheetMainFile)
	End If
	'Selecting Stylesheet Resulting Content Type
	If StrStylesheetResultingContentType<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Stylesheet Resulting Content Type:"
		Call Fn_Button_Click("Fn_ContentM_CreateStylesheet",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrStylesheetResultingContentType
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting ANT Target
 	If StrANTTarget<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ANT Build Target:"
		Call Fn_Edit_Box("Fn_ContentM_CreateStylesheet",objCMDialog,"AdministrativeEdit",StrANTTarget)
	End If
	'Taking content from file
	If StrContentPath<>"" Then
		'Press Browse button
		Call Fn_Button_Click("Fn_ContentM_CreateStylesheet",objCMDialog,"Browse")
		If objCMDialog.Dialog("Browse").Exist(6) Then
			objCMDialog.Dialog("Browse").WinEdit("FileName").Set StrContentPath
			wait 2
			objCMDialog.Dialog("Browse").WinButton("Open").Click
			wait 1
		Else
			Set objCMDialog=Nothing
			Exit Function
		End If
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateStylesheet",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateStylesheet",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateStylesheet=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Stylesheet of name [" + StrName + "]")
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateXMLAttributeMapping

'Description			 :	Function Used to create XML Attribute Mapping

'Parameters			   :   1.StrName: XML Attribute Mapping name
'										2.StrAdminComment : Admin Comment
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be present in Content Management perspective

'Examples				:  	bReturn=Fn_ContentM_CreateXMLAttributeMapping("XMLAttributeMapping24","Admin mapping")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												18-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateXMLAttributeMapping(StrName,StrAdminComment)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateXMLAttributeMapping"
    'Declaring variables
    Dim objCMDialog,StrMenu
	Fn_ContentM_CreateXMLAttributeMapping=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	'Checking existance of [ NewAdministrativeClass ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting XML Attribute Mapping Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateXMLAttributeMapping",objCMDialog, "ClassTree","Complete List:XML Attribute Mapping")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateXMLAttributeMapping",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateXMLAttributeMapping",objCMDialog)
	'Setting Name
	If StrName<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateXMLAttributeMapping",objCMDialog,"AdministrativeEdit",StrName)
	End If
	'Setting Admin Comment
	If StrAdminComment<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Administrator Comment:"
		Call Fn_Edit_Box("Fn_ContentM_CreateXMLAttributeMapping",objCMDialog,"AdministrativeEdit",StrAdminComment)
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateXMLAttributeMapping",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateXMLAttributeMapping",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateXMLAttributeMapping=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created XML Attribute Mapping of name [" + StrName + "]")
	Set objCMDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateStyleType

'Description			 :	Function Used to create Style Type

'Parameters			   :   1.StrName: Style Type name
'										2.StrSystemUsage : System Usage type
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be present in Content Management perspective

'Examples				:  	bReturn=Fn_ContentM_CreateStyleType("StyleType24","User")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												18-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateStyleType(StrName,StrSystemUsage)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateStyleType"
   'Declaring variables
    Dim objStyleTypeDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter

	Fn_ContentM_CreateStyleType=False
	'Creating object of [ New Administrative Class ] dialog
	Set objStyleTypeDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	 'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking existance of [ NewAdministrativeClass ] dialog
	If Not objStyleTypeDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting Style Type Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateStyleType",objStyleTypeDialog, "ClassTree","Complete List:Style Type")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateStyleType",objStyleTypeDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateStyleType",objStyleTypeDialog)
	'Setting Name
	If StrName<>"" Then
		objStyleTypeDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateStyleType",objStyleTypeDialog,"AdministrativeEdit",StrName)
	End If
	'Setting Admin Comment
	If StrSystemUsage<>"" Then
		objStyleTypeDialog.JavaStaticText("FieldName").SetTOProperty "label","System Usage:"
		Call Fn_Button_Click("Fn_ContentM_CreateStyleType",objStyleTypeDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objStyleTypeDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objStyleTypeDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrSystemUsage
			wait 2
			bFlag=true
			If objStyleTypeDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objStyleTypeDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateStyleType",objStyleTypeDialog,"Finish")
	If objStyleTypeDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateStyleType",objStyleTypeDialog,"Cancel")
	End If
	Fn_ContentM_CreateStyleType=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Style Type of name [" + StrName + "]")
	Set objStyleTypeDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateEditingTool

'Description			 :	Function Used to create Editing Tool

'Parameters			   :   1.StrName: Style Type name
'										2.StrToolActivation : Tool Activation
'										2.StrToolCommand : Tool Command
'										2.StrToolPath : Tool Path
'										2.StrTableReference : Graphic Priority Table Reference
'										2.bDownloadgraphics : Download graphics option
'										2.bDownloadschema : Download schema option
'										2.bDownloadstylesheet : Download stylesheet option
'										2.StrProcessingInstruction : Process ingInstruction
'										2.StrSystemUsage : System Usage
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be present in Content Management perspective

'Examples				:  	bReturn=Fn_ContentM_CreateEditingTool("Tool6","SIMPLE_TEXT_EDITOR","text.exe","path1","EDIT","True","True","True","","User")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												18-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateEditingTool(StrName,StrToolActivation,StrToolCommand,StrToolPath,StrTableReference,bDownloadgraphics,bDownloadschema,bDownloadstylesheet,StrProcessingInstruction,StrSystemUsage)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateEditingTool"
   'Declaring variables
    Dim obEditingToolDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter

	Fn_ContentM_CreateEditingTool=False
	'Creating object of [ New Administrative Class ] dialog
	Set obEditingToolDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	 'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking existance of [ NewAdministrativeClass ] dialog
	If Not obEditingToolDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting Editing Tool Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateEditingTool",obEditingToolDialog, "ClassTree","Complete List:Editing Tool")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateEditingTool",obEditingToolDialog)
	wait 3
	'Setting Name
	If StrName<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"AdministrativeEdit",StrName)
	End If
	'Setting Tool Activation
	If StrToolActivation<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Tool Activation:"
		Call Fn_Button_Click("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrToolActivation
			wait 2
			bFlag=true
			If obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set obEditingToolDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Setting Tool Command
	If StrToolCommand<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Tool Command:"
		Call Fn_Edit_Box("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"AdministrativeEdit",StrToolCommand)
	End If
	'Setting Tool Path
	If StrToolPath<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Tool Path:"
		Call Fn_Edit_Box("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"AdministrativeEdit",StrToolPath)
	End If
	'Setting Graphic Priority Table Reference
	If StrTableReference<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Graphic Priority Table Reference:"
		Call Fn_Button_Click("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrTableReference
			wait 2
			bFlag=true
			If obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set obEditingToolDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Set Download graphics
	If bDownloadgraphics<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Download graphics:"
		obEditingToolDialog.JavaRadioButton("AdministrativeRadioButton").SetTOProperty "attached text",bDownloadgraphics
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateEditingTool",obEditingToolDialog, "AdministrativeRadioButton")
	End If
	'Set Download schema
	If bDownloadschema<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Download schema:"
		obEditingToolDialog.JavaRadioButton("AdministrativeRadioButton").SetTOProperty "attached text",bDownloadschema
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateEditingTool",obEditingToolDialog, "AdministrativeRadioButton")
	End If
	'Set Download stylesheet
	If bDownloadstylesheet<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Download stylesheet:"
		obEditingToolDialog.JavaRadioButton("AdministrativeRadioButton").SetTOProperty "attached text",bDownloadstylesheet
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateEditingTool",obEditingToolDialog, "AdministrativeRadioButton")
	End If
	'Setting Processing Instruction
	If StrProcessingInstruction<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","Processing instruction:"
		Call Fn_Edit_Box("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"AdministrativeEdit",StrProcessingInstruction)
	End If
	'Setting System Usage
	If StrSystemUsage<>"" Then
		obEditingToolDialog.JavaStaticText("FieldName").SetTOProperty "label","System Usage:"
		Call Fn_Button_Click("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrSystemUsage
			wait 2
			bFlag=true
			If obEditingToolDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set obEditingToolDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"Finish")
	If obEditingToolDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateEditingTool",obEditingToolDialog,"Cancel")
	End If
	Fn_ContentM_CreateEditingTool=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Style Type of name [" + StrName + "]")
	Set obEditingToolDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_ExportDocument

'Description			 :	Function Used to Export Document

'Parameters			   :   1.StrLanguage: Language
'										2.bGraphicMode : Graphic mode option
'										3.bIncludeMainContent : Include main content option
'										4.bIncludeSupportingData : Include Supporting data option
'										5.bIncludeGraphicData : Include Graphic data option
'										6.bIncludeContentRef : Include Content reference option
'										7.bIncludeComposeRef : Include Compose reference option
'										8.StrZipFileSaveAsPath : Zip file save location
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be present in Content Management perspective

'Examples				:  	bReturn=Fn_ContentM_ExportDocument("English US","Import Original Name","EDIT","on","off","","","","D:\tc91\TestData\ContentManagementConfig\Test.zip")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												19-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_ExportDocument(StrLanguage,bGraphicMode,StrGraphicPriority,bIncludeMainContent,bIncludeSupportingData,bIncludeGraphicData,bIncludeContentRef,bIncludeComposeRef,StrZipFileSaveAsPath)
	'Variable declaration
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_ExportDocument"
	Dim objCMDialog
	Dim StrMenu,ZipName,bFlag,iCounter

	Fn_ContentM_ExportDocument=false
	'Creating Object of [ ExportDocument ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("ExportComposition")
	'Checking existance of [ ExportDocument ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ Tools->Import->Graphic... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "ExportDocument")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	   wait 5
	End If
	'Selecting language
	If StrLanguage<>"" Then
		Call Fn_List_Select("Fn_ContentM_ExportDocument", objCMDialog, "Language",StrLanguage)
	End If
	wait 1
	'Setting Graphic Mode option
	If bGraphicMode<>"" Then
		objCMDialog.JavaRadioButton("GraphicMode").SetTOProperty "attached text",bGraphicMode
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_ExportDocument",objCMDialog, "GraphicMode")
	End If
	wait 1

	'Selecting Graphic Priority
	If StrGraphicPriority<>"" Then
		Call Fn_List_Select("Fn_ContentM_ExportDocument", objCMDialog, "GraphicPriority",StrGraphicPriority)
	End If
	wait 1

	'Selcting Include Main Content option
	If bIncludeMainContent<>"" Then
		Call Fn_CheckBox_Set("Fn_ContentM_ExportDocument", objCMDialog, "IncludeMainContent", bIncludeMainContent)
	End If
	wait 1
	'Selcting Include Supporting Data option
	If bIncludeSupportingData<>"" Then
		Call Fn_CheckBox_Set("Fn_ContentM_ExportDocument", objCMDialog, "IncludeSupportingData", bIncludeSupportingData)
	End If
	wait 1
	'Selcting Include Graphic Data option
	If bIncludeGraphicData<>"" Then
		Call Fn_CheckBox_Set("Fn_ContentM_ExportDocument", objCMDialog, "IncludeGraphicData", bIncludeGraphicData)
	End If
	wait 1
	'Selcting Include Content References option
	If bIncludeContentRef<>"" Then
		Call Fn_CheckBox_Set("Fn_ContentM_ExportDocument", objCMDialog, "IncludeContentReferences", bIncludeContentRef)
	End If
	wait 1
	'Selcting Include Compose References option
	If bIncludeComposeRef<>"" Then
		Call Fn_CheckBox_Set("Fn_ContentM_ExportDocument", objCMDialog, "IncludeComposeReferences", bIncludeComposeRef)
	End If
	wait 1
	'clicking on Finish button
	Call Fn_Button_Click("Fn_ContentM_ExportDocument",objCMDialog,"Finish")
	'Checking existance of Save As dialog
	bFlag=false
	For iCounter=0 to 9
		If Dialog("Save As").Exist(5) Then
			wait(2)
			bFlag=true
			Exit for
		End If
	Next
	If bFlag=false Then
		Set objCMDialog=nothing
		Exit function
	End If
	'Save zip file at specific location
	If StrZipFileSaveAsPath<>"" Then
		If InStr(1,LCase(StrZipFileSaveAsPath),".zip") Then
			Dialog("Save As").WinEdit("FileName").Type StrZipFileSaveAsPath
			wait 2
		else
			ZipName=Dialog("Save As").WinEdit("FileName").GetROProperty("text")
			Dialog("Save As").WinEdit("FileName").Type StrZipFileSaveAsPath+"\"+ZipName+".Zip"
			wait 2
		End If
	End If
	Dialog("Save As").WinButton("Save").Click 1,1
	wait 1
	Fn_ContentM_ExportDocument=True
	'releasing object of [ ExportDocument ] dialog
	Set objCMDialog=nothing
End Function


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_Import_DITA_Map

'Description			 :	Function Used to Import Graphic Options

'Parameters			   :   '1.dicImportDITAMapInfo:  Import DITA Map Info
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	
'										dicImportDITAMapInfo("FileNames")="D:\tc91\testdata\000763.ditamap"
'										dicImportDITAMapInfo("TopicTypeName")="DITA Dynamic Map"
'										dicImportDITAMapInfo("GraphicAttributeMapping")="Default Graphic Attribute Mapping"
'										dicImportDITAMapInfo("GraphicMode")="Import Original Name"
'										dicImportDITAMapInfo("ReuseExistingTopic")="Overwrite Existing"

'										bReturn=Fn_ContentM_Import_DITA_Map(dicImportDITAMapInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Pallavi J												25 - Apr - 2012								1.0																						Sandeep N
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_Import_DITA_Map(dicImportDITAMapInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_Import_DITA_Map"
 	'Variable declaration
	Dim objCMDialog
	Dim StrMenu,iCounter,iCount,bFlag
	Fn_ContentM_Import_DITA_Map=False
	'Creating Object of [ Fn_ContentM_Import_DITA_Map ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("ImportDITAMapfromFile")
	'Checking existance of [ Fn_ContentM_Import_DITA_Map ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ Tools->Import->DITA Map... ]
	   StrMenu=Fn_GetXMLNodeValue( Environment.Value("sPath")+"\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "ImportDITAMap")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Setting FileName
'	objCMDialog.JavaEdit("FileName").Set dicImportDITAMapInfo("FileNames")
	wait 2
	
	'Setting From Directory
	JavaWindow("ContentManagement").JavaWindow("ImportDITAMapfromFile").JavaButton("Browse").Click micLeftBtn
	wait 2
	objCMDialog.Dialog("BrowseForFolder").WinEdit("FolderPath").Click
	wait 10
	set WshShell = CreateObject("WScript.Shell")
	WshShell.SendKeys "^a"
	wait 1
	WshShell.SendKeys "{DELETE}"
	wait 1
	set WshShell =Nothing
	objCMDialog.Dialog("BrowseForFolder").WinEdit("FolderPath").Type dicImportDITAMapInfo("FromDirectory")
	wait 1
	objCMDialog.Dialog("BrowseForFolder").WinButton("OK").Click
	wait 2
	Call Fn_ReadyStatusSync(1)
	
	'Select file names
	Call Fn_Button_Click("Fn_ContentM_Import_DITA_Map",objCMDialog,"SelectAll")

	
	'Setting Topic Type Name
	If dicImportDITAMapInfo("TopicTypeName")<>"" Then
		Call Fn_List_Select("Fn_ContentM_Import_DITA_Map",objCMDialog,"TopicTypeName",dicImportDITAMapInfo("TopicTypeName"))
	End If
	'Setting Graphic Attribute Mapping
	If dicImportDITAMapInfo("GraphicAttributeMapping")<>"" Then
		Call Fn_List_Select("Fn_ContentM_Import_DITA_Map",objCMDialog,"GraphicAttributeMapping",dicImportDITAMapInfo("GraphicAttributeMapping"))
	End If
	'Selecting Graphic Mode
	If dicImportDITAMapInfo("GraphicMode")<>"" Then
		objCMDialog.JavaRadioButton("GraphicMode").SetTOProperty "attached text",dicImportDITAMapInfo("GraphicMode")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_Import_DITA_Map",objCMDialog, "GraphicMode")
	End If
	'Selecting Reuse Existing Topic
	If dicImportDITAMapInfo("ReuseExistingTopic")<>"" Then
		If dicImportDITAMapInfo("ReuseExistingTopic") = "Find by XML Number" Then
			Call Fn_CheckBox_Set("Fn_ContentM_Import_DITA_Map", objCMDialog,"FindbyXMLNumber","on")
		Else
			Call Fn_CheckBox_Set("Fn_ContentM_Import_DITA_Map", objCMDialog,"FindbyXMLNumber","on")
			Call Fn_CheckBox_Set("Fn_ContentM_Import_DITA_Map", objCMDialog,"OverwriteExisting","on")
		End If	
	End If
	Call Fn_Button_Click("Fn_ContentM_Import_DITA_Map",objCMDialog,"Finish")

	bFlag=False
	For iCount=1 to 10
		JavaWindow("Shell").SetTOProperty "index",iCount
		For iCounter=0 To 10
			If JavaWindow("Shell").JavaWindow("PleaseWait").Exist(1) Then
				wait 5
				bFlag=True
			Else
				Exit for
			End If
		Next
		If bFlag=True Then
			Exit for
		End If
	Next

	Fn_ContentM_Import_DITA_Map=True
	Set objCMDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_XMLAttributeMapTableEntry

'Description			 :	Function Used to enter Entries in XML Attribute Map Table

'Parameters			   :   '1.dicXMLAttributeMapInfo: Table Entry Details
'
'Return Value		   : 	True or False

'Pre-requisite			:	XML Attribute Map Table Entry should appeared

'Examples				:  	dicXMLAttributeMapInfo("AttributeName")= "XMLAttr1"
'										dicXMLAttributeMapInfo("ConstantValue")="3"
'										dicXMLAttributeMapInfo("FieldSeparator")=":"
'										dicXMLAttributeMapInfo("FixFieldLength")="24"
'										dicXMLAttributeMapInfo("Function")="Clone"
'										dicXMLAttributeMapInfo("OmitEmptyAttribute")="True"
'										dicXMLAttributeMapInfo("Path")=	"Path1"
'										dicXMLAttributeMapInfo("XMLProcedure")="Procedure_5566"
'										bReturn=Fn_ContentM_XMLAttributeMapTableEntry(dicXMLAttributeMapInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												04-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_XMLAttributeMapTableEntry(dicXMLAttributeMapInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_XMLAttributeMapTableEntry"
	'variable declaration
	Dim objCMDialog,objTable,objChild,WshShell
	Dim bFlag,iRow,iCounter

	Set WshShell = CreateObject("WScript.Shell")
	Fn_ContentM_XMLAttributeMapTableEntry=False
	'Checking existance of [ XML Attribute Map Table Entry ] dialog
	If not JavaWindow("ContentManagement").JavaWindow("XMLAttributeMap").Exist(6) Then
		Exit function
	End If
	'Creating object of [ XML Attribute Map Table Entry ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("XMLAttributeMap")
	'Maximizing [ XML Attribute Map Table Entry ] dialog
	Call Fn_Window_Maximize("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog)
	wait 3
	'Setting [ Attribute Name ]
	If dicXMLAttributeMapInfo("AttributeName")<>"" Then
		Call Fn_Edit_Box("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"AttributeName",dicXMLAttributeMapInfo("AttributeName"))
	End If
	'Setting [ Constant Value ]
	If dicXMLAttributeMapInfo("ConstantValue")<>"" Then
		Call Fn_Edit_Box("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"ConstantValue",dicXMLAttributeMapInfo("ConstantValue"))
	End If
	'Setting [ Field Separator ]
	If dicXMLAttributeMapInfo("FieldSeparator")<>"" Then
		Call Fn_Edit_Box("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"FieldSeparator",dicXMLAttributeMapInfo("FieldSeparator"))
	End If
	'Setting [ Fix Field Length ]
	If dicXMLAttributeMapInfo("FixFieldLength")<>"" Then
		Call Fn_Edit_Box("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"FixedFieldLength",dicXMLAttributeMapInfo("FixFieldLength"))
	End If
	'Selecting [ Function ]
	If dicXMLAttributeMapInfo("Function")<>"" Then
		objCMDialog.JavaStaticText("FiedLable").SetTOProperty "label","Function:"
		Call Fn_Button_Click("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"ListButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicXMLAttributeMapInfo("Function")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set WshShell = Nothing
			Set objCMDialog=Nothing
			Exit Function
		End if
	End If
	'Setting [ Omit Empty Attribute ] option
	If dicXMLAttributeMapInfo("OmitEmptyAttribute")<>"" Then
		objCMDialog.JavaStaticText("FiedLable").SetTOProperty "label","Omit Empty Attribute:"
		objCMDialog.JavaRadioButton("OmitEmptyAttribute").SetTOProperty "attached text",dicXMLAttributeMapInfo("OmitEmptyAttribute")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog, "OmitEmptyAttribute")
	End If
	'Setting [ Path ]
	If dicXMLAttributeMapInfo("Path")<>"" Then
		objCMDialog.JavaStaticText("FiedLable").SetTOProperty "label","Path:"
		Call Fn_Edit_Box("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"Path",dicXMLAttributeMapInfo("Path"))
	End If
	'Selecting [ XMLProcedure ]
	If dicXMLAttributeMapInfo("XMLProcedure")<>"" Then
		objCMDialog.JavaStaticText("FiedLable").SetTOProperty "label","XML Procedure:"
		Call Fn_Button_Click("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"ListButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicXMLAttributeMapInfo("XMLProcedure")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set WshShell = Nothing
			Set objCMDialog=Nothing
			Exit Function
		End if
	End If
	 'Press finish button
	Call Fn_Button_Click("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_XMLAttributeMapTableEntry",objCMDialog,"Cancel")
	End If
	Fn_ContentM_XMLAttributeMapTableEntry=true
	'Releasing object of [ XML Attribute Map Table Entry ] dialog
	Set objCMDialog=Nothing
	Set WshShell = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_XMLAttributeMapTableGetCellData

'Description			 :	Function Used to Get Cell data from XML Attribute Map Table on summary tab

'Parameters			   :   '1.objTable: Table Object
'										 2.iRow : row number
'										 3.iCol: column number or name
'
'Return Value		   : 	True or False

'Pre-requisite			:	XML Attribute Map Table Should activated

'Examples				:  	bReturn=Fn_ContentM_XMLAttributeMapTableGetCellData(JavaWindow("ContentManagement").JavaTable("XMLAttributeMapTable"),1,"Attribute Name")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												04-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_XMLAttributeMapTableGetCellData(objTable,iRow,iCol)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_XMLAttributeMapTableGetCellData"
   Dim sColName,sPropName
   Fn_ContentM_XMLAttributeMapTableGetCellData=False
    If IsNumeric(iCol) Then
        sColName = objTable.Object.getColumn(iCol).getText()
    Else
        sColName = iCol
    End If
   
    Select Case Trim(sColName)
        Case "Attribute Name"
            sPropName="xamAttributeName"
        Case "Function"
            sPropName="xamFunction"
        Case "Field Separator"
            sPropName="xamFieldSeparator"
        Case "Fixed Field Length"
            sPropName="xamFixedFieldLength"
        Case "Constant Value"
            sPropName="xamConstantValue"
        Case "XML Procedure"
            sPropName="cmt0ProcedureTargef"
        Case "Omit Empty Attribute"
            sPropName="xamOmitEmptyAttribute"
        Case Else
            Exit function
    End Select
    Fn_ContentM_XMLAttributeMapTableGetCellData=objTable.Object.getItem(iRow).getData().getComponent().getProperty(sPropName)
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_XMLAttributeMapTableOperations

'Description			 :	Function Used to perform operations on XML Attribute Map Table on summary tab

'Parameters			   :   1.StrAction: Action name
'										2.StrAttributeName : Attribute Name
'										3.StrColName: Column name
'										4.StrValue: cell value
'										5.dicXMLAttributeMapInfo: XML Attribute Map Table entry details
'
'Return Value		   : 	True or False

'Pre-requisite			:	XML Attribute Map Table Should activated

'Examples				:  	bReturn= Fn_ContentM_XMLAttributeMapTableOperations("CellVerify","skdissno","Function","Bidirectional","")
'
'										dicXMLAttributeMapInfo("AttributeName")= "XMLAttr2"
'										dicXMLAttributeMapInfo("ConstantValue")="3"
'										dicXMLAttributeMapInfo("FieldSeparator")=":"
'										dicXMLAttributeMapInfo("FixFieldLength")="24"
'										dicXMLAttributeMapInfo("Function")="Clone"
'										dicXMLAttributeMapInfo("OmitEmptyAttribute")="True"
'										dicXMLAttributeMapInfo("Path")=	"Path1"
'										dicXMLAttributeMapInfo("XMLProcedure")="Procedure_5566"
'										bReturn=Fn_ContentM_XMLAttributeMapTableOperations("Add","","","",dicXMLAttributeMapInfo)
'
'										bReturn=Fn_ContentM_XMLAttributeMapTableOperations("Delete","skdissdate_month","","","")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												04-May-2012								1.0																						Sunny R
'													Sandeep N												31-Aug-2012								1.0					Modified Case : Delete						Pallavi J
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_XMLAttributeMapTableOperations(StrAction,StrAttributeName,StrColName,StrValue,dicXMLAttributeMapInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_XMLAttributeMapTableOperations"
     'Variable declaration
    Dim objXMLAttrMapTable
    Dim iRowCount,iCounter,bFlag
    Fn_ContentM_XMLAttributeMapTableOperations=false
    'creating object of [ XMLAttributeMapTable ]
    Set objXMLAttrMapTable=JavaWindow("ContentManagement").JavaTable("XMLAttributeMapTable")
   
    Select Case StrAction
        'Case to Add XML Attribute Map Entry
       '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        Case "Add"
            bFlag=Fn_TabFolder_Operation("Select", "Summary","")
			If bFlag=false Then
                JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").SetTOProperty "Index", "0"
				Call Fn_TabFolder_Operation("DoubleClickTab", "Summary","")
            else
				Call Fn_TabFolder_Operation("DoubleClickTab", "Summary","")	
			End If

            'bFlag=Fn_ToolbarButtonClick_Ext("2","Add...")
            bFlag = Fn_Button_Click( "Fn_ContentM_XMLAttributeMapTableOperations", JavaWindow("ContentManagement"),"XMLAttributeMapTableAdd" )
            If bFlag=true Then
                bFlag=Fn_ContentM_XMLAttributeMapTableEntry(dicXMLAttributeMapInfo)
                If bFlag=true Then
                    Fn_ContentM_XMLAttributeMapTableOperations=true
                End If
            End If
			Call Fn_TabFolder_Operation("DoubleClickTab", "Summary","")	
			JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").SetTOProperty "Index", "1"
       'Case to verify Cell value against specific Attrubite Name
       '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        Case "CellVerify"
            iRowCount=objXMLAttrMapTable.GetROProperty("rows")
            For iCounter=0 to iRowCount-1
                If trim(StrAttributeName)=trim(objXMLAttrMapTable.Object.getItem(iCounter).getData().toString()) Then
                    If Fn_ContentM_XMLAttributeMapTableGetCellData(objXMLAttrMapTable,iCounter,StrColName)=StrValue Then
                        Fn_ContentM_XMLAttributeMapTableOperations=True
                        Exit For
                    End If
                End If
            Next
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        'Case to Delete Entry from XML Attribute MapTable
        Case "Delete"
			 bFlag=Fn_TabFolder_Operation("Select", "Summary","")
			If bFlag=false Then
                JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").SetTOProperty "Index", "0"
				Call Fn_TabFolder_Operation("DoubleClickTab", "Summary","")
            else
				Call Fn_TabFolder_Operation("DoubleClickTab", "Summary","")	
			End If

            bFlag=false
            iRowCount=objXMLAttrMapTable.GetROProperty("rows")
            For iCounter=0 to iRowCount-1
			If trim(StrAttributeName)=trim(objXMLAttrMapTable.Object.getItem(iCounter).getData().getComponent().getProperty("xamAttributeName")) Then
				' objXMLAttrMapTable.ActivateRow iCounter
				objXMLAttrMapTable.SelectRow iCounter
				wait 1
				objXMLAttrMapTable.SelectCell iCounter,0
				'JavaWindow("ContentManagement").JavaToolbar("XMLAttributeMapAddAndDelete").Press "Delete"					
				bFlag = Fn_Button_Click( "Fn_ContentM_XMLAttributeMapTableOperations", JavaWindow("ContentManagement"),"XMLAttributeMapTableDelete" )
				wait 1
				' bFlag=Fn_ToolbarButtonClick_Ext("2","Delete")
				 If bFlag=true Then
						bFlag=false
						'                        If JavaDialog("DeletingSystemObject").Exist(3) then
						'                            JavaDialog("DeletingSystemObject").JavaButton("Yes").Click
						'							If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Delete").Exist(5) Then
						'								JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Delete").JavaButton("Yes").Click
						'								wait 2
						'								bFlag=true
						'								Exit For
						'							End If
						'                        end if
						If Dialog("Warning").Exist(3) then
							Dialog("Warning").WinButton("Yes").Click
							wait 1
							bFlag=true
							Exit For
						End if
				Else
					bFlag =  False
					Exit For
				End If
                	End If
            Next
            If bFlag=true Then
                Fn_ContentM_XMLAttributeMapTableOperations=true
            End If
			Call Fn_TabFolder_Operation("DoubleClickTab", "Summary","")	
			JavaWindow("DefaultWindow").JavaObject("RACTabFolderWidget").SetTOProperty "Index", "1"
   End Select
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateProcedure

'Description			 :	Function Used to create NewAdministrativeClass Schema

'Parameters			   :   '1.StrID: Procedure ID
'										 2.StrRevision: Procedure Revision
'										 3.StrName: Procedure Name
'                                        4.StrProcedureUsage: ProcedureUsage Type
'										 5.StrContentFilePath: StrProcedure Content File Path
'
'Return Value		   : 	True or False

'Pre-requisite			:	

'Examples				:   Fn_ContentM_CreateProcedure("","","Procedure","DC_ARCHIVE_TRANSFORMATION","C:\ProcDoc.txt")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Avinash Jagdale												04-May-2012								1.0																			
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreateProcedure(StrID,StrRevision,StrName,StrProcedureUsage,StrContentFilePath)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateProcedure"
'	Declaring variables
    Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateProcedure=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
    'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking existance of [ New Administrative Class ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting Procedure Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateProcedure",objCMDialog, "ClassTree","Complete List:Procedure")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateProcedure",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateProcedure",objCMDialog)
	wait 3
	'Set Procedure  ID
	If StrID<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateProcedure",objCMDialog,"Edit",StrID)
	End If
	'Set Procedure Revision
	If StrRevision<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateProcedure",objCMDialog,"Edit",StrRevision)
	End If
	'Set Procedure Name
	If StrName<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateProcedure",objCMDialog,"Edit",StrName)
	End If


	'Set Procedure Usage
	If StrProcedureUsage<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Procedure Usage:"
		Call Fn_Button_Click("Fn_ContentM_CreateProcedure",objCMDialog,"DropDownButton")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrProcedureUsage
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set objCMDialog=Nothing
			Set WshShell=Nothing
			Exit Function
		End If
	End If
	If StrContentFilePath<>"" Then
		Call Fn_Button_Click("Fn_ContentM_CreateProcedure",objCMDialog,"Browse")
		If objCMDialog.Dialog("Browse").Exist(10) Then
			objCMDialog.Dialog("Browse").WinEdit("FileName").Set StrContentFilePath
			wait 2 
			objCMDialog.Dialog("Browse").WinButton("Open").Click
			wait 2
		Else
			Set objCMDialog=Nothing
			Exit FUnction
		End If
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateProcedure",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateProcedure",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateProcedure=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateGraphicAttributeMapping

'Description			 :	Function Used to create Graphic Attribute Mapping

'Parameters			   :   1.StrName: Graphic Attribute Mapping name
'										2.StrAdminComment : Admin Comment
'
'Return Value		   : 	True or False

'Pre-requisite			:	Should be present in Content Management perspective

'Examples				:  	bReturn=Fn_ContentM_CreateGraphicAttributeMapping("GraphicAttributeMapping24","Admin mapping")
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Amol Lanke												07-Apr-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateGraphicAttributeMapping(StrName,StrAdminComment)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateGraphicAttributeMapping"
    'Declaring variables
    Dim objCMDialog,StrMenu
	Fn_ContentM_CreateGraphicAttributeMapping=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	'Checking existance of [ NewAdministrativeClass ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting XML Attribute Mapping Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateGraphicAttributeMapping",objCMDialog, "ClassTree","Complete List:Graphic Attribute Mapping")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateGraphicAttributeMapping",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateGraphicAttributeMapping",objCMDialog)
	wait 3
	'Setting Name
	If StrName<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateGraphicAttributeMapping",objCMDialog,"AdministrativeEdit",StrName)
	End If
	'Setting Admin Comment
	If StrAdminComment<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Administrator Comment:"
		Call Fn_Edit_Box("Fn_ContentM_CreateGraphicAttributeMapping",objCMDialog,"AdministrativeEdit",StrAdminComment)
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateGraphicAttributeMapping",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateGraphicAttributeMapping",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateGraphicAttributeMapping=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created XML Attribute Mapping of name [" + StrName + "]")
	Set objCMDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreatePublication

'Description			 :	Function Used to create NewAuthorClass Publication

'Parameters			   :   1.StrPublicationType : Publication Type
'										 2.StrID: Publication ID
'										 3.StrRevision: Publication Revision
'										 4.StrName: Publication Name
'                                        5.StrDocumentTitle: Document Title
'										 6.StrMasterLangRef: Master Language Reference
'										 7.bIsThisATemplate: Is this a Template
'
'Return Value		   : 	True or False

'Pre-requisite			:	

'Examples				:   bReturn=Fn_ContentM_CreatePublication("PubType04696","","A","Publication1","Title1","English US","False")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												09-May-2012								1.0																					Sunny R
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_CreatePublication(StrPublicationType,StrID,StrRevision,StrName,StrDocumentTitle,StrMasterLangRef,bIsThisATemplate)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreatePublication"
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreatePublication=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
    'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting S1000D Data Module Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreatePublication",objCMDialog, "ClassTree","Complete List:Publication")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreatePublication",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreatePublication",objCMDialog)
	wait 3
	'Selecting publication type
	If StrPublicationType<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreatePublication", objCMDialog,"JavaList",StrPublicationType)
	End If
	'Set publication ID
	If StrID<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreatePublication",objCMDialog,"ID",StrID)
	End If
	'Set publication Revision
	If StrRevision<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreatePublication",objCMDialog,"Revision",StrRevision)
	End If
	'Set publication Name
	If StrName<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreatePublication",objCMDialog,"Name",StrName)
	End If
	'Set publication Document Title
	If StrDocumentTitle<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreatePublication",objCMDialog,"DocumentTitle",StrDocumentTitle)
	End If
	'Setting Master Language Reference
	If StrMasterLangRef<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
		'Call Fn_Button_Click("Fn_ContentM_CreatePublication",objCMDialog,"DropDownButton")
'		JavaWindow("ContentManagement").JavaWindow("NewAuthorClass").JavaList("MasterLanguageReference").Select StrMasterLangRef
		If Fn_SISW_UI_JavaList_Operations("Fn_ContentM_CreatePublication", "Select", objCMDialog, "MasterLanguageReference", StrMasterLangRef, "", "") = false Then
		Fn_ContentM_CreatePublication = false
		Exit function
		End If
'		wait 1
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  StrMasterLangRef
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
'		If bFlag=False Then
'			Set objCMDialog=Nothing
'			Set WshShell=Nothing
'			Exit Function
'		End If
	End If
	'Set Is This a Template option
	If bIsThisATemplate<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is This A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",bIsThisATemplate
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreatePublication",objCMDialog, "IsThisATemplate")
	End If
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreatePublication",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreatePublication",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreatePublication=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_PublishContent

'Description			 :	Function Used to publish content

'Parameters			   :   '1.dicPublishInfo: Content publish information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Publish Dialog should be appear on screen

'Examples				:  	dicPublishInfo("Tool")="FOP"
'										dicPublishInfo("StyleType")="DITA"
'										dicPublishInfo("Language")="English US"
'										dicPublishInfo("TranslationVersionSelection")="Match Topic"
'										dicPublishInfo("RegisterResult")="Composed Document"
'										dicPublishInfo("DitaFilterValue")="DITAFilter1"
'										bReturn=Fn_ContentM_PublishContent(dicPublishInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												10-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_PublishContent(dicPublishInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_PublishContent"
 	'Declaring variables
	Dim objPublishDialog
	Dim aDitaFilterValue,iCounter
	Fn_ContentM_PublishContent=false
 	'Checking existance of [ Publish ] dialog
	If not JavaWindow("ContentManagement").JavaWindow("Publish").Exist(6) Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile")," [ Publish ] dialog is not exist")
		Exit function
	else
		'Creating object of [ Publish ] Dialog
		Set objPublishDialog=JavaWindow("ContentManagement").JavaWindow("Publish")
	End If
	'Selecting Publish Tool
	If dicPublishInfo("Tool")<>"" Then
		objPublishDialog.JavaStaticText("Field_Name").SetTOProperty "label","Tool"
		Call Fn_List_Select("Fn_ContentM_PublishContent", objPublishDialog,"Combo",dicPublishInfo("Tool"))
	End If
	'Selecting Publish Style Type
	If dicPublishInfo("StyleType")<>"" Then
		objPublishDialog.JavaStaticText("Field_Name").SetTOProperty "label","Style Type"
		Call Fn_List_Select("Fn_ContentM_PublishContent", objPublishDialog,"Combo",dicPublishInfo("StyleType"))
	End If
	'Selecting Publish Language
	If dicPublishInfo("Language")<>"" Then
		objPublishDialog.JavaStaticText("Field_Name").SetTOProperty "label","Language"
		Call Fn_List_Select("Fn_ContentM_PublishContent", objPublishDialog,"Combo",dicPublishInfo("Language"))
	End If
	'Selecting Publish Compose Version Selection
	If dicPublishInfo("ComposeVersionSelection")<>"" Then
		objPublishDialog.JavaStaticText("Field_Name").SetTOProperty "label","Compose Version Selection"
		Call Fn_List_Select("Fn_ContentM_PublishContent", objPublishDialog,"Combo",dicPublishInfo("ComposeVersionSelection"))
	End If
	'Selecting Publish Translation Version Selection
	If dicPublishInfo("TranslationVersionSelection")<>"" Then
		objPublishDialog.JavaStaticText("Field_Name").SetTOProperty "label","Translation Version Selection"
		Call Fn_List_Select("Fn_ContentM_PublishContent", objPublishDialog,"Combo",dicPublishInfo("TranslationVersionSelection"))
	End If
	'Setting Publish Resulting File Folder
	If dicPublishInfo("ResultingFileFolder")<>"" Then
		objPublishDialog.JavaStaticText("Field_Name").SetTOProperty "label","Resulting File Folder"
		Call Fn_Edit_Box("Fn_ContentM_PublishContent",objPublishDialog,"Text_Field",dicPublishInfo("ResultingFileFolder"))
	End If
	'Setting Publish Resulting File Name
	If dicPublishInfo("ResultingFileName")<>"" Then
		objPublishDialog.JavaStaticText("Field_Name").SetTOProperty "label","Resulting File Name"
		Call Fn_Edit_Box("Fn_ContentM_PublishContent",objPublishDialog,"Text_Field",dicPublishInfo("ResultingFileName"))
	End If
	'Selecting Publish Register Result
	If dicPublishInfo("RegisterResult")<>"" Then
		objPublishDialog.JavaStaticText("Field_Name").SetTOProperty "label","Register Result"
		Call Fn_List_Select("Fn_ContentM_PublishContent", objPublishDialog,"Combo",dicPublishInfo("RegisterResult"))
	End If
	'Selecting Dita Filter Value
	If dicPublishInfo("DitaFilterValue")<>"" Then
		aDitaFilterValue=Split(dicPublishInfo("DitaFilterValue"),"~")
		For iCounter=0 to uBound(aDitaFilterValue)
			If Fn_UI_ListItemExist("Fn_ContentM_PublishContent",objPublishDialog,"DitaFilterValue",aDitaFilterValue(iCounter)) Then
				Call Fn_List_Select("Fn_ContentM_PublishContent",objPublishDialog,"DitaFilterValue",aDitaFilterValue(iCounter))
			else
				Set objPublishDialog=nothing
				Exit function
			End If
		Next
	End If
	'Selecting [ View ] option
	If dicPublishInfo("View")<>"" Then
		Call Fn_CheckBox_Set("Fn_ContentM_PublishContent", objPublishDialog,"View",dicPublishInfo("View"))
	End If
	'Press finish button
	wait 2
	Call Fn_Button_Click("Fn_ContentM_PublishContent",objPublishDialog,"Finish")
	Fn_ContentM_PublishContent=true
	'Releasing object of [ Publish ] dialog
	Set objPublishDialog=nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_PreviewTabOperations

'Description			 :	Function Used to Perform operation on Preview Tab

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrText: Text
'
'Return Value		   : 	True or False

'Pre-requisite			:	Preview Tab Should be activated

'Examples				:   bReturn=Fn_ContentM_PreviewTabOperations("VerifyText","Topic TypeTitle")
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												10-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_PreviewTabOperations(StrAction,StrText)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_PreviewTabOperations"
	Fn_ContentM_PreviewTabOperations=false
	Select Case StrAction
		Case "VerifyText"
			If InStr(1,JavaWindow("ContentManagement").JavaObject("PreviewTabObject").Object.getText(),StrText) Then
				Fn_ContentM_PreviewTabOperations=true
			End If
	End Select
End Function

'------------------------------------------------------------------------------------------------------------------------------
'Function Name		:	Fn_ContentM_CreateTopic
'Description			 :	Function To Create Topic
'Parameters			   :   '1.dicS1000DDataModuleList4Info
'
'Return Value		   : 	True or False
'Pre-requisite			:	Node is selected in Home
'														dicS1000DDataModuleList4Info("TopicType")="content"
'														dicS1000DDataModuleList4Info("Revision")="A"
'														dicS1000DDataModuleList4Info("ID")="55696"
'														dicS1000DDataModuleList4Info("Name")="DML66"
'														dicS1000DDataModuleList4Info("MasterLanguageReference")="English US"
'														dicS1000DDataModuleList4Info("DocumentTitle")="DML2486 Title"
'Examples				:   bReturn=Fn_ContentM_CreateTopic(dicS1000DDataModuleList4Info)
'                       
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'												Amol Lanke											10-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 
Public Function Fn_ContentM_CreateTopic(dicS1000DDataModuleList4Info)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateTopic"
 	'variable declaration
	Dim objCMDialog,StrMenu
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateTopic=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
    'Creating shell object
	'Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If

		'Selecting S1000D Data Module List 4.0 Option from list
	Call Fn_JavaTree_Select("Fn_ContentM_CreateTopic",objCMDialog, "ClassTree","Complete List:Topic")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateTopic",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateTopic",objCMDialog)
	wait 3
	'Selecting topic type
	If dicS1000DDataModuleList4Info("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateTopic", objCMDialog, "SelectTopicType",dicS1000DDataModuleList4Info("TopicType"))
	End If

	If dicS1000DDataModuleList4Info("ID")<>"" Then
		'Setting ID for Translation Order
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","ID:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTopic",objCMDialog,"ID",dicS1000DDataModuleList4Info("ID"))
	End If

	'Set revision
	If dicS1000DDataModuleList4Info("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTopic",objCMDialog,"Revision",dicS1000DDataModuleList4Info("Revision"))
	End If
	'Set Name
	If dicS1000DDataModuleList4Info("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTopic",objCMDialog,"Name",dicS1000DDataModuleList4Info("Name"))
	End If
	'Setting Master Language Reference
	If dicS1000DDataModuleList4Info("MasterLanguageReference")<>"" Then	
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
		Call Fn_List_Select("Fn_ContentM_CreateTopic", objCMDialog, "SelectTopicType",dicS1000DDataModuleList4Info("MasterLanguageReference"))
'		Call Fn_Button_Click("Fn_ContentM_CreateTopic",objCMDialog,"DropDownButton")
'		wait 1
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicS1000DDataModuleList4Info("MasterLanguageReference")
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
'		If bFlag=False Then
'			Set objCMDialog=Nothing
'			Set WshShell=Nothing
'			Exit Function
'		End If
	End If
	'Setting Document Title
	If dicS1000DDataModuleList4Info("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateTopic",objCMDialog,"DocumentTitle",dicS1000DDataModuleList4Info("DocumentTitle"))
	End If
    
	'Set Is This A Template option
	If dicS1000DDataModuleList4Info("IsThisATemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicS1000DDataModuleList4Info("IsThisATemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateTopic",objCMDialog, "IsThisATemplate")
	End If
	'Set Reference Only option
	If dicS1000DDataModuleList4Info("ReferenceOnly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicS1000DDataModuleList4Info("ReferenceOnly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateTopic",objCMDialog, "ReferenceOnly")
	End If
    'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateTopic",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateTopic",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateTopic=True
	Set objCMDialog=Nothing
'	Set WshShell=Nothing
End Function


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateTransformationPolicy

'Description			 :	Function Used to CreateTransformationPolicy

'Parameters			   :   '1.CreateTransformationPolicy
'
'Return Value		   : 	True or False

'Pre-requisite			:	Node is selected in Home

'Examples				:	Fn_ContentM_CreateTransformationPolicy("Policy","AdminComment")
'                   						
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Amol Lanke											18-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateTransformationPolicy(StrName,StrAdminComment)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateTransformationPolicy"
 	'Declaring variables
    Dim objCMDialog,StrMenu
	Fn_ContentM_CreateTransformationPolicy=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	'Checking existance of [ NewFolder ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting Publication Type Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_CreateTransformationPolicy",objCMDialog, "ClassTree","Complete List:Transformation Policy")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateTransformationPolicy",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateTransformationPolicy",objCMDialog)
	wait 3
	'Set Publication Type Name
	If StrName<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateTransformationPolicy",objCMDialog.JavaStaticText("FieldName"),"label","Name:")
		Call Fn_Edit_Box("Fn_ContentM_CreateTransformationPolicy",objCMDialog,"Name", StrName)
	End If
	'Set Publication Type Local Tag Name
	If StrAdminComment<>"" Then
    		Call Fn_Edit_Box("Fn_ContentM_CreateTransformationPolicy",objCMDialog,"AdminComment", StrAdminComment)
	End If

	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateTransformationPolicy",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
		Call Fn_Button_Click("Fn_ContentM_CreateTransformationPolicy",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateTransformationPolicy=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Transformation Policy of name ["+StrName+"]")
	Set objCMDialog=Nothing
End Function


''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :     Fn_SISW_ContentM_SummaryPolicyOperations(sAction,bSummaryTab,sUserAction,sXMLProcedure,sVerifyVal,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will perform Add,Delete & Verify Operations on the Policy in Summary Tab
''''/$$$$ 
''''/$$$$   PRE-REQUISITES        :  Relevent Policy Should be Present
''''/$$$$
''''/$$$$  PARAMETERS   : 		sAction :Valid Action Name
''''/$$$$										bSummaryTab : Boolean Parameter to Activate the Summary Tab
''''/$$$$										sUserAction : Value to be set in the "User Action Field"
''''/$$$$										sXMLProcedure : Value to be set in the "sXMLProcedure Field"
''''/$$$$										sVerifyVal : Value to be Verified in the Policy Table
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2: For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_MenuOperation(), Fn_KeyBoardOperation
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS         21/05/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			 21/05/2012          1.0
''''/$$$$
''''/$$$$		How To Use :    bReturn=Fn_SISW_ContentM_SummaryPolicyOperations("Delete","true","Export","","","","")
''''/$$$$									 bReturn=Fn_SISW_ContentM_SummaryPolicyOperations("Add","true","Export","Proc_123","","","")
''''/$$$$									 bReturn=Fn_SISW_ContentM_SummaryPolicyOperations("Verify","true","","",""Export,"","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SISW_ContentM_SummaryPolicyOperations(sAction,bSummaryTab,sUserAction,sXMLProcedure,sVerifyVal,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_ContentM_SummaryPolicyOperations"
   Dim bReturn,objContent,sValue,iCount,sRows,sColumns,objTransformPolicy,jCount,bFlag,WshShell
	Fn_SISW_ContentM_SummaryPolicyOperations=false

	Set WshShell = CreateObject("WScript.Shell")
	Set objContent=JavaWindow("ContentManagement")
	Set objTransformPolicy=JavaWindow("ContentManagement").JavaWindow("TransformationPolicy")

	'Activate the Summary Tab If Required

	If cBool(bSummaryTab)=true Then
		bReturn=Fn_TabFolder_Operation("Select", "Summary","")
		If bReturn=True Then
			wait 3
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Activated the Summary Tab")

		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Activate the Summary Tab")
			Fn_SISW_ContentM_SummaryPolicyOperations=False
			Exit Function
		End If
	End If

	Select Case sAction

				Case "Verify"

					sRows=objContent.JavaTable("Policies").GetROProperty ("rows")
					  For iCount=0 to sRows-1
								sValue= objContent.JavaTable("Policies").Object.getItem(iCount).getData().getComponent().toString()
									If lCAse(sValue)=lCAse(sVerifyVal) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified the Value ["+sValue+"]")
										bFlag=true
										Exit For
									End If
					  Next


					If bFlag=True Then
						Fn_SISW_ContentM_SummaryPolicyOperations=true
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sValue+"]")
						Fn_SISW_ContentM_SummaryPolicyOperations=False
						Exit Function
					End If


					Case "Add"


							'Press the Add Button on the Toolbar
							'objContent.JavaToolbar("Policies").Press "Add..."
							objContent.JavaButton("Add").Click
							If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Press the Add Button")
									Fn_SISW_ContentM_SummaryPolicyOperations=False
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Pressed the Add Button")
									Fn_SISW_ContentM_SummaryPolicyOperations=true
							End If
							Wait 2

							If sUserAction<>"" Then
									objTransformPolicy.JavaButton("Button").Click
									wait 1
									WshShell.SendKeys "{TAB}"
									wait 1
									WshShell.SendKeys "{DOWN}"
									wait 1
									If objTransformPolicy.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
										objTransformPolicy.JavaWindow("TreeShell").JavaTree("Tree").Activate  sUserAction
										wait 2
										bFlag=true
										If objTransformPolicy.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
											bFlag=False
										End If
									Else
										bFlag=False
									End If
									If bFlag=False Then
										Set WshShell = Nothing
										Set objTransformPolicy=Nothing
										Exit Function
									End if
							End If

							If sXMLProcedure<>"" Then
								objTransformPolicy.JavaList("XMLProcedure").Select sXMLProcedure
							End If

							'Click on the Finish Button
							objTransformPolicy.JavaButton("Finish").Click micLeftBtn
							If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Press the Finish Button")
									Fn_SISW_ContentM_SummaryPolicyOperations=False
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Pressed the Finish Button")
									Fn_SISW_ContentM_SummaryPolicyOperations=true
							End If
							Wait 2

							'Click on the Cancel Button
							objTransformPolicy.JavaButton("Cancel").Click micLeftBtn
							If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Press the Cancel Button")
									Fn_SISW_ContentM_SummaryPolicyOperations=False
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Pressed the Cancel Button")
									Fn_SISW_ContentM_SummaryPolicyOperations=true
							End If

				Case "Delete"

					'Select the Row You Want to Delete
					bFlag=False
					sRows=objContent.JavaTable("Policies").GetROProperty ("rows")
					  For iCount=0 to sRows-1
								sValue= objContent.JavaTable("Policies").Object.getItem(iCount).getData().getComponent().getProperty("tpUserAction")
									If lCAse(sValue)=lCAse(sUserAction) Then
										 objContent.JavaTable("Policies").SelectRow(iCount)
                                         wait 1
										  objContent.JavaTable("Policies").SelectCell iCount,0
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Row At index ["+cstr(iCount)+"]")
										bFlag=true
										Exit For
									End If
					  Next


							'Press the Delete Button on the Toolbar
							objContent.JavaToolbar("Policies").Press "Delete"
							If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Press the Delete Button")
									Fn_SISW_ContentM_SummaryPolicyOperations=False
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Pressed the Delete Button")
									Fn_SISW_ContentM_SummaryPolicyOperations=true
							End If
							Wait 2

							'Click on Yes of the Deleting System Object Dialog
							JavaDialog("DeletingSystemObject").JavaButton("Yes").Click micLeftBtn
							If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Press the Delete Button on the Deleting System Object Dialog")
									Fn_SISW_ContentM_SummaryPolicyOperations=False
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Pressed the Delete Button on the Deleting System Object Dialog")
									Fn_SISW_ContentM_SummaryPolicyOperations=true
							End If

							'Click on Yes on the Delete Dialog

							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Delete").JavaButton("Yes").Click micLeftBtn

							If err.number<0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Press the Delete Button on the Delete Dialog")
									Fn_SISW_ContentM_SummaryPolicyOperations=False
									Exit Function
							Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Pressed the Delete Button on the Delete Dialog")
									Fn_SISW_ContentM_SummaryPolicyOperations=true
							End If

	End Select
	Set WshShell = Nothing
	Set objContent=Nothing
	Set objTransformPolicy=Nothing
End Function 


Public Function Fn_ContentM_CreateS1000DDataModuleList(dicS1000DDataModuleList4Info)
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_CreateS1000DDataModuleList=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	If dicS1000DDataModuleList4Info("AuthorClass")="" Then
		'Selecting S1000D Data Module List 4.0 Option from list
		Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog, "ClassTree","Complete List:S1000D Data Module List")
	Else
		Call Fn_JavaTree_Select("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog, "ClassTree","Complete List:"+dicS1000DDataModuleList4Info("AuthorClass"))
	End If
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog)
	wait 3
	'Selecting topic type
	If dicS1000DDataModuleList4Info("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataModuleList4", objCMDialog, "SelectTopicType",dicS1000DDataModuleList4Info("TopicType"))
	End If
	'Set revision
	If dicS1000DDataModuleList4Info("Revision")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Revision:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Revision",dicS1000DDataModuleList4Info("Revision"))
	End If
	'Set Name
	If dicS1000DDataModuleList4Info("Name")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Name:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Name",dicS1000DDataModuleList4Info("Name"))
	End If
	'Setting Master Language Reference
	If dicS1000DDataModuleList4Info("MasterLanguageReference")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Master Language Reference:"
'		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModule",objCMDialog,"DropDownButton")
		Call Fn_List_Select("Fn_ContentM_CreateS1000DDataModuleList4", objCMDialog, "SelectTopicType",dicS1000DDataModuleList4Info("MasterLanguageReference"))
'		wait 1
'		WshShell.SendKeys "{TAB}"
'		wait 1
'		WshShell.SendKeys "{DOWN}"
'		wait 1
'        If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate  dicS1000DDataModuleList4Info("MasterLanguageReference")
'			wait 2
'			bFlag=true
'			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
'				bFlag=False
'			End If
'		Else
'			bFlag=False
'		End If
'		If bFlag=False Then
'			Set objCMDialog=Nothing
'			Set WshShell=Nothing
'			Exit Function
'		End If
	End If
	'Setting Document Title
	If dicS1000DDataModuleList4Info("DocumentTitle")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Document Title:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"DocumentTitle",dicS1000DDataModuleList4Info("DocumentTitle"))
	End If
	'Setting Model Identification Code
	If dicS1000DDataModuleList4Info("ModelIdentificationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Model Identification Code:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"ModelIdentificationCode",dicS1000DDataModuleList4Info("ModelIdentificationCode"))
	End If
	'Setting Originator
	If dicS1000DDataModuleList4Info("Originator")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Originator:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Originator",dicS1000DDataModuleList4Info("Originator"))
	End If
	'Setting Type of Data Module List
	If dicS1000DDataModuleList4Info("TypeOfDataModuleList")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Type of Data Module List:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"TypeofDataModuleList",dicS1000DDataModuleList4Info("TypeOfDataModuleList"))
	End If
	'Setting Year of Dispatch
	If dicS1000DDataModuleList4Info("YearOfDispatch")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Year of Dispatch:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"YearofDispatch",dicS1000DDataModuleList4Info("YearOfDispatch"))
	End If
	'Setting Sequence Number
	If dicS1000DDataModuleList4Info("SequenceNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Sequence Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"SequenceNumber",dicS1000DDataModuleList4Info("SequenceNumber"))
	End If
	
'	Select Case dicS1000DDataModuleList4Info("AuthorClass") 
'		Case "S1000D Data Module List"	
'		Setting Issue Number
			If dicS1000DDataModuleList4Info("IssueNumber")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Number:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueNumber",dicS1000DDataModuleList4Info("IssueNumber"))
			End If
			'Setting Issue Type
			If dicS1000DDataModuleList4Info("IssueType")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issue Type:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueType",dicS1000DDataModuleList4Info("IssueType"))
			End If
			'Setting Issued Day
			If dicS1000DDataModuleList4Info("IssuedDay")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Day:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueDay",dicS1000DDataModuleList4Info("IssuedDay"))
			End If
			'Setting Issued Month
			If dicS1000DDataModuleList4Info("IssuedMonth")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Month:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueMonth",dicS1000DDataModuleList4Info("IssuedMonth"))
			End If
			'Setting Issued Year
			If dicS1000DDataModuleList4Info("IssuedYear")<>"" Then
				objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Issued Year:"
				Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"IssueYear",dicS1000DDataModuleList4Info("IssuedYear"))
			End If
'	End Select

	'Setting Remarks
	If dicS1000DDataModuleList4Info("Remarks")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Remarks:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Remarks",dicS1000DDataModuleList4Info("Remarks"))
	End If
	'Setting In Work Number
	If dicS1000DDataModuleList4Info("InWorkNumber")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","In Work Number:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"InWorkNumber",dicS1000DDataModuleList4Info("InWorkNumber"))
	End If
	'Setting Security Class
	If dicS1000DDataModuleList4Info("SecurityClass")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Security Class:"
		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"SecurityClass",dicS1000DDataModuleList4Info("SecurityClass"))
	End If
    'Setting Export File Name
'	If dicS1000DDataModuleList4Info("ExportFileName")<>"" Then
'		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Export File Name:"
'		Call Fn_Edit_Box("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"ExportFileName",dicS1000DDataModuleList4Info("ExportFileName"))
'	End If
'	'Set Is This A Template option
	If dicS1000DDataModuleList4Info("IsThisATemplate")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Is A Template:"
		objCMDialog.JavaRadioButton("IsThisATemplate").SetTOProperty "attached text",dicS1000DDataModuleList4Info("IsThisATemplate")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog, "IsThisATemplate")
	End If
'	'Set Reference Only option
	If dicS1000DDataModuleList4Info("ReferenceOnly")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Reference only:"
		objCMDialog.JavaRadioButton("ReferenceOnly").SetTOProperty "attached text",dicS1000DDataModuleList4Info("ReferenceOnly")
		Call Fn_UI_JavaRadioButton_SetON("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog, "ReferenceOnly")
	End If
    'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Finish")
	wait 2
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_CreateS1000DDataModuleList4",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateS1000DDataModuleList=True
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_ContentM_VariantExpressionEditor_Operations
'@@
'@@    Description			:	Function Used to Perform operation on Variant Expression Editor
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. StrTabName	: Tab Name
'@@							:	3. dicDetails	: Dictionary object
'@@							:	4. Popupmenu	: Popup menu name
'@@							:	3. sTabClose	: Flag to close Tab
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Contemt Management Perspective should be opened 
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@									dicDetails("ColumnName") = "0004"
'@@									dicDetails("Option") = "a1~b1"
'@@									dicDetails("OptionFlag") = "Check~Check"
'@@									dicDetails("ApplicabilityOption") = "m1~a1"
'@@									dicDetails("ApplicabilityOptionFlag") = "Check~Check"
'@@							bReturn = Fn_ContentM_VariantExpressionEditor_Operations("SetVariantExpEditorOptionAndSave","000202-PRoduct (Variant Expression Editor)",dicDetails,"","Yes")
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Amruta Patil			09-Feb-2022				1.0		  		 Created		
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_ContentM_VariantExpressionEditor_Operations(sAction,StrTabName,dicDetails,Popupmenu,sTabClose)

	GBL_FAILED_FUNCTION_NAME = "Fn_ContentM_VariantExpressionEditor_Operations"
	
	Dim objCCWindow,arrOptionFlag,arrOptions,iSubJectInx,iApplicabilityInx,sMenu,bFlag
	Dim iCnt,StrBounds,iX,iY,iCnt1,iRowIndex,strRow,iRowsCount,iColIndex,strColumn,iColCount
	Dim myDeviceReplay,sAppMsg,iColPosition,aOption,iInstance,iCounter
	
	Fn_ContentM_VariantExpressionEditor_Operations = False
	On Error Resume Next
	
	Set objCCWindow = JavaWindow("ContentManagement")
	If Fn_UI_ObjectExist("Fn_ContentM_VariantExpressionEditor_Operations",objCCWindow.JavaObject("VariantExpressionEditor")) = False Then
		Fn_ContentM_VariantExpressionEditor_Operations = False
		Set objCCWindow = Nothing
		Exit function
	End If
	' Maximize tab
	 If StrTabName <> "" Then
	 	Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
		Call Fn_ReadyStatusSync(1)
	 End If 
	
	Select Case sAction
		Case "ExpColumnExist"
			iColIndex = 0
			If dicDetails("ColumnName") <> "" Then
				columnArr = Split(dicDetails("ColumnName"),"~")
			Else
				Fn_ContentM_VariantExpressionEditor_Operations = False
				Exit Function
			End If
			For iCount = 0 To UBound(columnArr) - 1
		 		iColCount = objCCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
		 		For iCnt = 0 To iColCount - 1
		 			iColPosition = objCCWindow.JavaObject("VariantExpressionEditor").Object.getstartXOfcolumnPosition(iCnt+3)
					strColumn =  CStr(objCCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
					If trim(dicDetails("ColumnName")) = trim(strColumn) Then
							iColIndex = iColIndex + 1
					End If
				Next
			Next
			If iColIndex = UBound(columnArr) - 1 Then
				Fn_ContentM_VariantExpressionEditor_Operations = True
				Exit Function
			Else
				Fn_ContentM_VariantExpressionEditor_Operations = True
				Exit Function
			End If
		
		Case "SetVariantExpEditorOptionAndSave","MultiColumnAndSetVarVariantExpEditorAndSave"
		
				 If dicDetails("ColumnName") <> "" Then  'Get column index
			 		iColCount = objCCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
			 		For iCnt = 0 To iColCount - 1
						strColumn =  CStr(objCCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If trim(dicDetails("ColumnName")) = trim(strColumn) Then
								iColIndex = iCnt + 1
								Exit for
						End If
					Next
				 Else
					iColIndex = 2							 
				 End If
				 If dicDetails("GridNumber") <> "" Then
					iColIndex = dicDetails("GridNumber")
				 End If
				iRowsCount = objCCWindow.JavaObject("VariantExpressionEditor").Object.getRowCount()  'Get Index for Subject & Applicability sections
				 If dicDetails("Option") <> "" AND dicDetails("OptionFlag") <> "" Then
					arrOptions = Split(dicDetails("Option"),"~")
					arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
					For iCnt = 0 To UBound(arrOptions)
						'--------------- for multiple instance ------------------------------
				 		If instr(arrOptions(iCnt),"@") Then
				 			aOption = Split(arrOptions(iCnt),"@")
				 			arrOptions(iCnt) = aOption(0)
				 			iInstance = aOption(1)
				 		Else
							iInstance = 1				 		
				 		End If
				 		'--------------------------------------------------------------
				 		iCounter = 1
						For iCnt1 = 1 To iRowsCount
							If instr(CStr(objCCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
								iRowIndex = iCnt1
								'--------------------------------------------
								If cint(iCounter) = cint(iInstance) Then
					 				Exit for
					 			Else
									iCounter = iCounter + 1							 			
					 			End If
					 			'-------------------------------------------
							End If
						Next
						
							StrBounds = objCCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
				 			StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
							If IsEmpty(iX) or IsEmpty(iY) Then
					 			Fn_ContentM_VariantExpressionEditor_Operations = False
								Set objCCWindow = Nothing
					 			Exit Function
				 			End If
							iX = iX + 4
							iY = iY + 5	
							If sAction = "MultiColumnAndSetVarVariantExpEditorAndSave" Then
						 	  iX = iX + 29
							End  If
						Select Case arrOptionFlag(iCnt)
				 			Case "Check" 'For "=" or "=Any" condition
				 				objCCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				 				Wait 1
				 			Case "None" 'For "!=" or or "=NONE" condition
				 				objCCWindow.JavaObject("VariantExpressionEditor").dblClick iX,iY,"LEFT"
				 				Wait 1
				 			Case "Blank" 'to set blank
				 			   ' For future use
				 			Case else
				 				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Invalid case [ "&arrOptionFlag(iCnt)&" ]" )
								Fn_ContentM_VariantExpressionEditor_Operations = False
								Set objCCWindow = Nothing
								Exit function
				 		End Select		
					Next					
				 End If
				 
				 Fn_ContentM_VariantExpressionEditor_Operations = True	
				 
				 sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Toolbar"),"Savethecurrentcontents") 'Save the Variant Condition
				 Call Fn_ToolBarOperation("Click",sMenu,"")
				 Call Fn_ReadyStatusSync(1)

Set objCCWindow = Nothing
End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_ReferenceTopicTypeSelection

'Description			 :	Function Used to select Reference Topic Type 

'Parameters			   :   '1.dicTopicTypeInfo
'
'Return Value		   : 	True or False

'Pre-requisite			:	Reference Topic Type Selection window should be open

'Examples				:	Fn_ContentM_CreateTransformationPolicy(dicTopicTypeInfo,"Finish")
'                   						
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Vaishali D											18-May-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_ReferenceTopicTypeSelection(dicTopicTypeInfo,sButton)
	Dim objRefTopicDlg
	set objRefTopicDlg = JavaWindow("ContentManagement").JavaWindow("Reference Topic Type Selection")
	
	If objRefTopicDlg.Exist(2) Then
		Call Fn_SISW_UI_JavaList_Operations("Fn_ContentM_ReferenceTopicTypeSelection", "Select", objRefTopicDlg, "ReferenceTopicType", dicTopicTypeInfo("ReferenceTopicType"), "", "")
		Call Fn_Button_Click("Fn_ContentM_ReferenceTopicTypeSelection",objRefTopicDlg,sButton)
		Fn_ContentM_ReferenceTopicTypeSelection = True
		Exit function
	End If
	set objRefTopicDlg = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_ContentM_VariantExpressionEditor_Operations
'@@
'@@    Description			:	Function Used to Perform operation on Variant Expression Editor
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. StrTabName	: Tab Name
'@@							:	3. dicDetails	: Dictionary object
'@@							:	4. Popupmenu	: Popup menu name
'@@							:	3. sTabClose	: Flag to close Tab
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Contemt Management Perspective should be opened 
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@									dicDetails("ColumnName") = "0004"
'@@									dicDetails("Option") = "a1~b1"
'@@									dicDetails("OptionFlag") = "Check~Check"
'@@									dicDetails("ApplicabilityOption") = "m1~a1"
'@@									dicDetails("ApplicabilityOptionFlag") = "Check~Check"
'@@							bReturn = Fn_ContentM_VariantExpressionEditor_Operations("SetVariantExpEditorOptionAndSave","000202-PRoduct (Variant Expression Editor)",dicDetails,"","Yes")
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Amruta Patil			09-Feb-2022				1.0		  		 Created		
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_ContentM_VariantExpressionEditor_Operations(sAction,StrTabName,dicDetails,Popupmenu,sTabClose)

	GBL_FAILED_FUNCTION_NAME = "Fn_ContentM_VariantExpressionEditor_Operations"
	
	Dim objCCWindow,arrOptionFlag,arrOptions,iSubJectInx,iApplicabilityInx,sMenu,bFlag
	Dim iCnt,StrBounds,iX,iY,iCnt1,iRowIndex,strRow,iRowsCount,iColIndex,strColumn,iColCount
	Dim myDeviceReplay,sAppMsg,iColPosition,aOption,iInstance,iCounter
	
	Fn_ContentM_VariantExpressionEditor_Operations = False
	On Error Resume Next
	
	Set objCCWindow = JavaWindow("ContentManagement")
	If Fn_UI_ObjectExist("Fn_ContentM_VariantExpressionEditor_Operations",objCCWindow.JavaObject("VariantExpressionEditor")) = False Then
		Fn_ContentM_VariantExpressionEditor_Operations = False
		Set objCCWindow = Nothing
		Exit function
	End If
	' Maximize tab
	 If StrTabName <> "" Then
	 	Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
		Call Fn_ReadyStatusSync(1)
	 End If 
	
	Select Case sAction
		Case "ExpColumnExist"
			Call Fn_CM_SetRuleDate("SetNoRuleDate","","","No Date") 'Clear date rule
			Call Fn_ReadyStatusSync(1)
			iColIndex = 0
			If dicDetails("ColumnName") <> "" Then
				columnArr = Split(dicDetails("ColumnName"),"~")
			Else
				Fn_ContentM_VariantExpressionEditor_Operations = False
				Exit Function
			End If
			For iCount = 0 To UBound(columnArr) - 1
		 		iColCount = objCCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
		 		For iCnt = 0 To iColCount - 1
		 			iColPosition = objCCWindow.JavaObject("VariantExpressionEditor").Object.getstartXOfcolumnPosition(iCnt+3)
					strColumn =  CStr(objCCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
					If trim(dicDetails("ColumnName")) = trim(strColumn) Then
							iColIndex = iColIndex + 1
					End If
				Next
			Next
			If iColIndex = UBound(columnArr) - 1 Then
				Fn_ContentM_VariantExpressionEditor_Operations = True
				Exit Function
			Else
				Fn_ContentM_VariantExpressionEditor_Operations = True
				Exit Function
			End If
		
		Case "SetVariantExpEditorOptionAndSave","MultiColumnAndSetVarVariantExpEditorAndSave"
					Call Fn_CM_SetRuleDate("SetNoRuleDate","","","No Date") 'Clear date rule
					Call Fn_ReadyStatusSync(1)
					
				 If dicDetails("ColumnName") <> "" Then  'Get column index
			 		iColCount = objCCWindow.JavaObject("VariantExpressionEditor").Object.getColumnCount()
			 		For iCnt = 0 To iColCount - 1
						strColumn =  CStr(objCCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(iCnt,0).getlayer.getColumnHeaderlayer.getDatavalueByPosition(iCnt,0).tostring)
						If trim(dicDetails("ColumnName")) = trim(strColumn) Then
								iColIndex = iCnt + 1
								Exit for
						End If
					Next
				 Else
					iColIndex = 2							 
				 End If
				 If dicDetails("GridNumber") <> "" Then
					iColIndex = dicDetails("GridNumber")
				 End If
				iRowsCount = objCCWindow.JavaObject("VariantExpressionEditor").Object.getRowCount()  'Get Index for Subject & Applicability sections
				 If dicDetails("Option") <> "" AND dicDetails("OptionFlag") <> "" Then
					arrOptions = Split(dicDetails("Option"),"~")
					arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
					For iCnt = 0 To UBound(arrOptions)
						'--------------- for multiple instance ------------------------------
				 		If instr(arrOptions(iCnt),"@") Then
				 			aOption = Split(arrOptions(iCnt),"@")
				 			arrOptions(iCnt) = aOption(0)
				 			iInstance = aOption(1)
				 		Else
							iInstance = 1				 		
				 		End If
				 		'--------------------------------------------------------------
				 		iCounter = 1
						For iCnt1 = 1 To iRowsCount
							If instr(CStr(objCCWindow.JavaObject("VariantExpressionEditor").Object.getCellByPosition(0,iCnt1).getDataValue().getData().toString()),arrOptions(iCnt)) Then
								iRowIndex = iCnt1
								'--------------------------------------------
								If cint(iCounter) = cint(iInstance) Then
					 				Exit for
					 			Else
									iCounter = iCounter + 1							 			
					 			End If
					 			'-------------------------------------------
							End If
						Next
						
							StrBounds = objCCWindow.JavaObject("VariantExpressionEditor").Object.getBoundsByPosition(iColIndex,iRowIndex).tostring
				 			StrBounds = Split(Replace(Replace(StrBounds,"Rectangle {",""),"}",""),",")
				 			iX = cint(StrBounds(2))
				 			iY = cint(StrBounds(1))
							If IsEmpty(iX) or IsEmpty(iY) Then
					 			Fn_ContentM_VariantExpressionEditor_Operations = False
								Set objCCWindow = Nothing
					 			Exit Function
				 			End If
							iX = iX + 4
							iY = iY + 5	
							If sAction = "MultiColumnAndSetVarVariantExpEditorAndSave" Then
						 	  iX = iX + 29
							End  If
						Select Case arrOptionFlag(iCnt)
				 			Case "Check" 'For "=" or "=Any" condition
				 				objCCWindow.JavaObject("VariantExpressionEditor").Click iX,iY,"LEFT"
				 				Wait 1
				 			Case "None" 'For "!=" or or "=NONE" condition
				 				objCCWindow.JavaObject("VariantExpressionEditor").dblClick iX,iY,"LEFT"
				 				Wait 1
				 			Case "Blank" 'to set blank
				 			   ' For future use
				 			Case else
				 				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Invalid case [ "&arrOptionFlag(iCnt)&" ]" )
								Fn_ContentM_VariantExpressionEditor_Operations = False
								Set objCCWindow = Nothing
								Exit function
				 		End Select		
					Next					
				 End If
				 
				 Fn_ContentM_VariantExpressionEditor_Operations = True	
				 
				 sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Toolbar"),"Savethecurrentcontents") 'Save the Variant Condition
				 Call Fn_ToolBarOperation("Click",sMenu,"")
				 Call Fn_ReadyStatusSync(1)

Set objCCWindow = Nothing
End Select
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_CC_VariantConfigurationView_Operation
'@@
'@@    Description			:	Function Used to Perform operation on Variant Configuration View
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. StrTabName	: Tab Name
'@@							:	2. dicDetails	: Dictionary object
'@@							:	2. Popupmenu	: Popup menu name
'@@							:	3. sTabClose	: Flag to close Tab
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@
'@@    Examples				:	Set dicDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicDetails("Options") = "1200~1600~Arial~std"
'@@ 								dicDetails("OptionFlag") = "Check~Check~None~Check"
'@@									dicDetails("ToolBarButton") = "Expand"
'@@ 							bReturn = Fn_CC_VariantConfigurationView_Operation("SetVarOptionValue","Variant Configuration",dicDetails,"","Yes")
'@@ 								
'@@ 					
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done	
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Amruta Patil			18-Feb-2022				1.0		  		 Created		 
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_CC_VariantConfigurationView_Operation(sAction,StrTabName,dicDetails,Popupmenu,sTabClose)

	GBL_FAILED_FUNCTION_NAME = "Fn_CC_VariantConfigurationView_Operation"
	Dim objCCWindow,arrOptionFlag,arrOptions
	Dim sMenu,bFlag,iCnt,iCnt1,iRowIndex,iRowsCount,iColIndex,iColCount,iCnt2,sOption,strFlag
	Dim sAppMsg,iX,iY,iHeight,iItemHeight,aOption,iInstance,iCounter
	
	Fn_CC_VariantConfigurationView_Operation = False
	Set objCCWindow = JavaWindow("ContentManagement")
	
	Select Case sAction
		Case "SetVarOptionValue"
		       		Call Fn_CM_SetRuleDate("SetNoRuleDate","","","No Date") 'Clear date rule
			   		Call Fn_ReadyStatusSync(1)
				If Fn_UI_ObjectExist("Fn_CC_VariantConfigurationView_Operation",objCCWindow.JavaTable("VariantConfigTable")) = True Then
						 ' Maximize tab
						 If StrTabName <> "" Then
						 		Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
								Call Fn_ReadyStatusSync(1)
						 End If 
						 
						 arrOptions = Split(dicDetails("Options"),"~") ' Set Options as On or OFF
						 arrOptionFlag = Split(dicDetails("OptionFlag"),"~")
						 iRowsCount = objCCWindow.JavaTable("VariantConfigTable").GetROProperty("rows")
					     iColCount = objCCWindow.JavaTable("VariantConfigTable").GetROProperty("cols")			 
						 For iCnt = 0 To UBound(arrOptions)
						 		'--------------------- for multiple instance ----------------------------
						 		If instr(arrOptions(iCnt),"@") Then
						 			aOption = Split(arrOptions(iCnt),"@")
						 			arrOptions(iCnt) = aOption(0)
						 			iInstance = aOption(1)
						 		Else
									iInstance = 1				 		
						 		End If
						 		'--------------------- ---------------------- ----------------------------
							 	iCounter = 1
							 	For iCnt1 = 0 To iRowsCount - 1 'Get Row & col index as per option name	
							 		bFlag = False
							 		For iCnt2 = 0 To iColCount - 1
							 			sOption = objCCWindow.JavaTable("VariantConfigTable").GetCellData(iCnt1,iCnt2)
							 			If trim(arrOptions(iCnt)) = trim(sOption) Then
							 				'--------------------- ---------------------- ----------------------------
							 				If cint(iCounter) = cint(iInstance) Then
							 					bFlag = True
							 					Exit for
								 			Else
												iCounter = iCounter + 1							 			
								 			End If
							 				'--------------------- ---------------------- ----------------------------
							 			End If
							 		Next
							 		If bFlag = True Then
							 			iRowIndex = iCnt1
							 			If dicDetails("GridNumber") <> "" Then
							 				iColIndex = dicDetails("GridNumber")
							 			Else
							 				iColIndex = iCnt2-1
							 			End If
							 			Exit for
							 		End If	
							 	Next
							 	If sAction = "SetVarOptionValue_Ext" Then  'For Boolean Value to check  for same value
							 	    	iRowIndex=iRowIndex+1
							 	    	iColIndex=iColIndex+3
							 	       Call Fn_ReadyStatusSync(2)
							 	    End If
								Select Case arrOptionFlag(iCnt)
						 			Case "Check" 'For "=" or "=Any" condition
						 				objCCWindow.JavaTable("VariantConfigTable").SelectCell iRowIndex,iColIndex
						 				Wait 1
						 			Case "None" 'For "!=" or or "=NONE" condition
						 				objCCWindow.JavaTable("VariantConfigTable").SelectCell iRowIndex,iColIndex
						 				Wait 1
						 				objCCWindow.JavaTable("VariantConfigTable").SelectCell iRowIndex,iColIndex
						 				Wait 1
						 			Case Else
						 				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to invalid case [ "&arrOptionFlag(iCnt)&" ]" )
										Fn_CC_VariantConfigurationView_Operation = False
										Set objCCWindow = Nothing
										Exit function
						 		End Select	
								Fn_CC_VariantConfigurationView_Operation = True	
						 Next
						 
						 If dicDetails("ToolBarButton") <> "" Then    ' Click toolbar button in Configuration view
						 		sOption = Split(dicDetails("ToolBarButton"),"~")
						 		For iCnt1 = 0 To UBound(sOption)
						 			Call Fn_ToolBarOperation("Click",sOption(iCnt1),"")
						 			Call Fn_ReadyStatusSync(1)
						 		Next
						 End If
						 
						 If StrTabName <> "" Then ' Minimize Tab
						 		Call Fn_TabFolder_Operation("DoubleClickTab", StrTabName,"")
								Call Fn_ReadyStatusSync(1)
								If sTabClose = "Yes" Then ' Close Tab
									Call Fn_TabFolder_Operation("Close", StrTabName,"")
									Call Fn_ReadyStatusSync(1)
								End If	
						 End If
				End If
	End  Select	
	
	Set objCCWindow = Nothing
	
End Function
'======================================================================================================================================================================
'@@    Function Name		:	Fn_CC_VariantRule_Operations
'@@
'@@    Description			:	Function Used to Perform operation on Save Product Configuration  dialog
'@@
'@@    Parameters			:	1. sAction		: Action to be performed
'@@							:	2. dicDetails	: Dictionary object
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@
'@@    Examples				:	Set dicSaveDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicSaveDetails("Name") = "SVR01"
'@@									dicSaveDetails("Description") = "SVR01 description"
'@@									dicSaveDetails("Button") = "OK"
'@@ 							bReturn = Fn_CC_VariantRule_Operations("Save",dicSaveDetails)
'@@ 								
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done		
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Amruta Patil			13-Feb-2022				1.0		  		 Created		  
'=========================================================================================================================================================================
Public Function Fn_CC_VariantRule_Operations(sAction,dicSaveDetails)

	GBL_FAILED_FUNCTION_NAME = "Fn_CC_VariantRule_Operations"
	Fn_CC_VariantRule_Operations = False
	Set objSaveSVR = Fn_SISW_ContentM_GetObject("SaveVariantRule")
	'Check Existence of Save As window
	If Fn_UI_ObjectExist("Fn_CC_VariantRule_Operations",objSaveSVR) = False Then
		 Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Save Variant Rule ] window." )
		 Fn_CC_VariantRule_Operations = False
		 Set objSaveSVR = Nothing
		 Exit Function
	End If
	
	Select Case sAction
		Case "Save"
				If dicSaveDetails("VRScope") <> "" Then  ' Enter Name
					 objSaveSVR.JavaRadioButton("VariantRuleScope").SetTOProperty "attached text",dicSaveDetails("VRScope")
					 Call Fn_SISW_UI_JavaRadioButton_Operations("Fn_CC_VariantRule_Operations", "Set", objSaveSVR, "VariantRuleScope", "ON")
					 Call Fn_ReadyStatusSync(1)
				End If
				If dicSaveDetails("Name") <> "" Then  ' Enter Name
					 Call Fn_SISW_UI_JavaEdit_Operations("Fn_CC_VariantRule_Operations", "Type", objSaveSVR, "Name", dicSaveDetails("Name"))
					 Call Fn_ReadyStatusSync(1)
				End If
				
				If dicSaveDetails("Description") <> "" Then 'Enter Description
					 Call Fn_Edit_Box("Fn_CC_VariantRule_Operations",objSaveSVR,"Description",dicSaveDetails("Description"))
					 Call Fn_ReadyStatusSync(1)
				End If
				
				If dicSaveDetails("Button") <> "" Then  'Handle button OK or Cancel
					Call Fn_Button_Click("Fn_CC_VariantRule_Operations",objSaveSVR,dicSaveDetails("Button"))
					Call Fn_ReadyStatusSync(1)
				End If
	End Select
	Fn_CC_VariantRule_Operations = True
	Set objSaveSVR = Nothing
	
End Function
'@@=====================================================================================================================================================================
'@@
'@@    Function Name		:	Fn_CC_LoadVariantRule_Operations
'@@
'@@    Description			:	Function Used to Perform operation on Load Variant Rule dialog
'@@
'@@    Parameters			:	1. sAction				: Action to be performed
'@@							:	2. dicLoadVarDetails	: Dictionary object
'@@
'@@    Return Value		   	: 	True Or False
'@@
'@@    Pre-requisite		:	Load Variant Rule dialog should be opened 
'@@
'@@    Examples				:	Set dicLoadVarDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicLoadVarDetails("VariantRuleNames") = "SVR_CC_Car~SVR_IR_Car~SVR_CC_Utility_Vehicle"
'@@									dicLoadVarDetails("Button") = "Cancel"
'@@ 							bReturn = Fn_CC_LoadVariantRule_Operations("VerifyVariantRules",dicLoadVarDetails)	
'@@
'@@								Set dicLoadVarDetails = CreateObject( "Scripting.Dictionary")
'@@ 								dicLoadVarDetails("SearchCriteria") = "Name:SVR01~ID:0001~Description:TestDescription"
'@@ 							bReturn = Fn_CC_LoadVariantRule_Operations("SearchVariantRules",dicLoadVarDetails)	
'@@
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   History				:	Developer Name				Date	  			Rev. No.		Changes Done	
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@	   Created By 			:	Amruta Patil			13-Feb-2022				1.0		  		 Created		 
'@@=======================================================================================================================================================================
Public Function Fn_CC_LoadVariantRule_Operations(sAction,dicLoadVarDetails)

	GBL_FAILED_FUNCTION_NAME = "Fn_CC_LoadVariantRule_Operations"
	Dim objLoadVarRuledialog,arrVarRules,iRowCnt,iCount,bFlag,iCount1
	Dim sVarRule,sAppMsg,arrSearchCriteria,aSearchValues,sCheckStatus
	
	Fn_CC_LoadVariantRule_Operations = False
	Set objLoadVarRuledialog = Fn_SISW_ContentM_GetObject("LoadVariantRule")
	
	'Check Existence of Load Variant Rule dialog
	If Fn_UI_ObjectExist("Fn_CC_LoadVariantRule_Operations",objLoadVarRuledialog) = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Load Variant Rule ] dialog." )
		Fn_CC_LoadVariantRule_Operations = False
		Set objLoadVarRuledialog = Nothing
		Exit Function
	End If
	
	Select Case sAction
		'==============================================================================================================================		
		Case "SearchVariantRules" 'Verify variant Rule Names
				If dicLoadVarDetails("SearchCriteria") <> "" Then
					arrSearchCriteria = Split(dicLoadVarDetails("SearchCriteria"),"~")
					 For iCount = 0 To UBound(arrSearchCriteria)
							aSearchValues = Split(arrSearchCriteria(iCount),":")
							objLoadVarRuledialog.JavaEdit("SearchCriteria").SetTOProperty "attached text",aSearchValues(0)+":"
							bFlag = Fn_SISW_UI_JavaEdit_Operations("Fn_CC_LoadVariantRule_Operations", "Set", objLoadVarRuledialog, "SearchCriteria",aSearchValues(1))
							If bFlag = False Then
					  			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to check existence of [ Search field = "&aSearchValues(0)&" ].")
					  			Fn_CC_LoadVariantRule_Operations = False
					  			Exit For
					  		End If
					 Next
					Fn_CC_LoadVariantRule_Operations = Fn_Button_Click("Fn_CC_LoadVariantRule_Operations",objLoadVarRuledialog,"Search")
					Call Fn_ReadyStatusSync(1)
			  End If
	'==============================================================================================================================		
	 Case "SelectVariantRules"
	 	  If dicLoadVarDetails("VariantRuleNames") <> "" Then
			  arrVarRules = Split(dicLoadVarDetails("VariantRuleNames"),"~")
			  iRowCnt = objLoadVarRuledialog.JavaTable("VariantRules").GetROProperty("rows")
			  For iCount = 0 To UBound(arrVarRules)
			  		bFlag = False
			  		For iCount1 = 0 To iRowCnt - 1
			  			sVarRule = Fn_UI_JavaTable_GetCellData("Fn_CC_LoadVariantRule_Operations",objLoadVarRuledialog,"VariantRules",iCount1,"Name")
			  			If trim(sVarRule) = trim(arrVarRules(iCount)) Then
			  				bFlag = Fn_UI_JavaTable_ClickCell("Fn_CC_LoadVariantRule_Operations",objLoadVarRuledialog,"VariantRules",iCount1,"Select")
			  				Wait 1
			  				Exit For
			  			End If
			  		Next
			  		If bFlag = False Then
			  			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to select [ Variant Rule = "&arrVarRules(iCount)&" ].")
			  			Fn_CC_LoadVariantRule_Operations = False
			  			Exit For
			  		Else
			  			Fn_CC_LoadVariantRule_Operations = bFlag
			  		End If
			  Next  
		   End If	
	 '==============================================================================================================================	
	 Case "ButtonClick"
		 Fn_CC_LoadVariantRule_Operations = Fn_Button_Click("Fn_CC_LoadVariantRule_Operations",objLoadVarRuledialog,dicLoadVarDetails("Button"))		 
		 Call Fn_ReadyStatusSync(1)	
		 dicLoadVarDetails("Button") = ""
	'==============================================================================================================================		
	End Select
	
	If dicLoadVarDetails("Button") <> "" Then  'Click on Buttons
		 Call Fn_Button_Click("Fn_CC_LoadVariantRule_Operations",objLoadVarRuledialog,dicLoadVarDetails("Button"))	
		 Call Fn_ReadyStatusSync(1)	
	End If 
	
	Set objLoadVarRuledialog = Nothing
	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_CreateStandardNumberingSM

'Description			 :	Function Used to create S1000D Standard Numbering System Root Node  Type

'Parameters			   :   '1.dicNumberingSystemInfo: Publication Type information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicNumberingSystemInfo("Name")="PT123456"
'										dicNumberingSystemInfo("ModelIdentifier")="L10NTEST"
'										dicNumberingSystemInfo("SystemDifferenceCode")="AAA"
'										dicNumberingSystemInfo("DisassemblyCodeVariant")="AAA"
'										dicNumberingSystemInfo("SecurityClassification")="44"
'										dicNumberingSystemInfo("DefaultLanguage")="English US"
'History					 :			
'													Developer Name											Date								
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Amruta Patil										 17-Mar-2022								
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_ContentM_CreateStandardNumberingSM(dicNumberingSystemInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_CreateStandardNumberingSM"
 	'Declaring variables
    Dim objCMDialog,WshShell,StrMenu
	Dim bFlag

	Fn_ContentM_CreateStandardNumberingSM=False
	'Creating object of [ New Administrative Class ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAdministrativeClass")
	Set ErrDialog = JavaWindow("WEmbeddedFrame").JavaDialog("Paste").JavaDialog("Error")
	Set PasteDialog = JavaWindow("WEmbeddedFrame").JavaDialog("Paste")
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking existance of [ NewFolder ] dialog
	If Not objCMDialog.Exist(6) Then
	   'Select menu [ File->New->New Administrative Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAdministrativeClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting S1000D Standard Numbering System Root Node Type Option from list
	If dicNumberingSystemInfo("ItemType")<> "" Then
    
   	Call Fn_JavaTree_Select("Fn_ContentM_CreateStandardNumberingSM",objCMDialog, "ClassTree","Complete List:"&dicNumberingSystemInfo("ItemType"))
	End  IF
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_CreateStandardNumberingSM",objCMDialog)
	wait(3)
	If dicNumberingSystemInfo("TopicType")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Select Topic Type:")
		bFlag = Fn_List_Select("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"TopicList",dicNumberingSystemInfo("TopicType"))
		
		If bFlag=False Then
			Set WshShell = Nothing
			Set objCMDialog=Nothing
			Exit Function
		End If
	End If
	If dicNumberingSystemInfo("ID")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","ID:")
		Call Fn_Edit_Box("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Name", dicNumberingSystemInfo("ID"))
	End If
	If dicNumberingSystemInfo("Name")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Name:")
		Call Fn_Edit_Box("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Name", dicNumberingSystemInfo("Name"))
	End If
	If dicNumberingSystemInfo("Value")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Value:")
		Call Fn_Edit_Box("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Name", dicNumberingSystemInfo("Value"))
	End If
	If dicNumberingSystemInfo("Description")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Description:")
		Call Fn_Edit_Box("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Name", dicNumberingSystemInfo("Description"))
	End If
	
	If dicNumberingSystemInfo("ModelIdentifier")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Model Identifier:")
		Call Fn_Edit_Box("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Name", dicNumberingSystemInfo("ModelIdentifier"))
	End If
	If dicNumberingSystemInfo("SystemDifferenceCode")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","System Difference Code:")
		Call Fn_Edit_Box("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Name", dicNumberingSystemInfo("SystemDifferenceCode"))
	End If

	If dicNumberingSystemInfo("DisassemblyCodeVariant")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Disassembly Code Variant:")
		Call Fn_Edit_Box("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Name", dicNumberingSystemInfo("DisassemblyCodeVariant"))
	End If

	If dicNumberingSystemInfo("SecurityClassification")<>"" Then
        Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Security Classification:")
		Call Fn_Edit_Box("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Name", dicNumberingSystemInfo("SecurityClassification"))
	End If
	If dicNumberingSystemInfo("DefaultLanguage")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Default Language:")
		Call Fn_Button_Click("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"SystemUsage")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate dicNumberingSystemInfo("DefaultLanguage")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set WshShell = Nothing
			Set objCMDialog=Nothing
			Exit Function
		End If
	End If
	
	
	If dicNumberingSystemInfo("Type")<>"" Then
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ContentM_CreateStandardNumberingSM",objCMDialog.JavaStaticText("FieldName"),"label","Type:")
		Call Fn_Button_Click("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"SystemUsage")
		wait 1
		WshShell.SendKeys "{TAB}"
		wait 1
		WshShell.SendKeys "{DOWN}"
		wait 1
		If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
			objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Activate dicNumberingSystemInfo("Type")
			wait 2
			bFlag=true
			If objCMDialog.JavaWindow("TreeShell").JavaTree("Tree").Exist(1) Then
				bFlag=False
			End If
		Else
			bFlag=False
		End If
		If bFlag=False Then
			Set WshShell = Nothing
			Set objCMDialog=Nothing
			Exit Function
		End If
	End If
	
	'Press finish button
	Call Fn_Button_Click("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Finish")
	If objCMDialog.Exist(5) Then
		wait 2
'		Added Call to handle error dialog as per the discussion with Vishal Patil And Adapala Darvin
		For iCnt = 1 To 20
			JavaWindow("WEmbeddedFrame").SetTOProperty "index",iCnt
			If PasteDialog.Exist(5) Then
				Exit For
			End If
		Next
		If dicNumberingSystemInfo("ErrorHandle") <> "" Then
			If ErrDialog.Exist(5) Then
					ErrDialog.JavaButton("OK").Click
					wait 5
					If PasteDialog.Exist(5) Then
						PasteDialog.JavaButton("OK").Click
					End If
			End If
			Wait 2
		End If
		
		Call Fn_Button_Click("Fn_ContentM_CreateStandardNumberingSM",objCMDialog,"Cancel")
	End If
	Fn_ContentM_CreateStandardNumberingSM=True
    Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Publication Type of name [" + dicNumberingSystemInfo("Name") + "]")
	Set objCMDialog=Nothing
	Set WshShell = Nothing
End Function


Public Function Fn_CM_SetRuleDate(sAction,sDate,sTime,sButton)

	GBL_FAILED_FUNCTION_NAME = "Fn_CM_SetRuleDate"
	Dim objSetRuleDate,objPCWindow
	
	Fn_CM_SetRuleDate = False
	
	Set objPCWindow = JavaWindow("ContentManagement")
	Set objSetRuleDate = objPCWindow.JavaWindow("Set Rule Date")
	
	Select Case sAction
		
		Case "SetNoRuleDate"
					If Fn_UI_ObjectExist("Fn_CM_SetRuleDate",objPCWindow.JavaObject("DateTimeImageHyperlink")) = True Then
						Call Fn_UI_JavaObject_Click("Fn_CM_SetRuleDate",objPCWindow,"DateTimeImageHyperlink",5,5,"LEFT")
						Call Fn_ReadyStatusSync(1)
						'objPCWindow.WinMenu("ContextMenu").Select "Set Rule Date"
						objPCWindow.JavaMenu("Label:=Set Rule Date").Select
						Wait(1)
						
						'Click on button OK / Cancel / No Date
						If sButton <> "" Then
							objSetRuleDate.JavaButton("Button").SetTOProperty "label",sButton
							Fn_CM_SetRuleDate = Fn_Button_Click("Fn_CM_SetRuleDate",objSetRuleDate,"Button")
							Call Fn_ReadyStatusSync(1)
						End If
					End If
			End Select
	
	Set objPCWindow = Nothing
	Set objSetRuleDate = Nothing

End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_ContentM_VerifyS1000DDataModule4Field

'Description			 :	Function Used to Verify S1000D Data Module 4.0 fields

'Parameters			   :   '1.dicS1000DDataModule4Info: S1000D Data Module 4.0 information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Content Management perspective Should activated

'Examples				:  	dicS1000DDataModule4Info("TopicType")="Description-4-0"
'										dicS1000DDataModule4Info("ModelIdentifier")="MIC1"
'										dicS1000DDataModule4Info("SystemDifferenceCode")="SD1"
'										dicS1000DDataModule4Info("DisassemblyCode")="DC6"
'										dicS1000DDataModule4Info("DisassemblyCodeVariant")="DVC6"
'										dicS1000DDataModule4Info("InformationCode")="IC6"
'										dicS1000DDataModule4Info("InformationCodeVariant")="ICV3"
'										bReturn=Fn_ContentM_VerifyS1000DDataModule4Field(dicS1000DDataModule4Info)
'										
'History					 :			
'													Developer Name											Date								
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'													Amruta Patil										 30-March-2022
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_ContentM_VerifyS1000DDataModule4Field(dicS1000DDataModule4Info)
	GBL_FAILED_FUNCTION_NAME="Fn_ContentM_VerifyS1000DDataModule4Field"
 	'variable declaration
	Dim objCMDialog,StrMenu,WshShell
	Dim bFlag,objTable,objChild,iRow,iCounter
	Fn_ContentM_VerifyS1000DDataModule4Field=False
	'Creating object of [ NewAuthorClass ] dialog
	Set objCMDialog=JavaWindow("ContentManagement").JavaWindow("NewAuthorClass")
	'Creating shell object
	Set WshShell = CreateObject("WScript.Shell")
	bFlag=False

	'Checking Existance of [ NewAuthorClass ] dialog
	If Not objCMDialog.Exist(6) Then
		'Select menu [ File->New->New Author Class... ]
	   StrMenu=Fn_GetXMLNodeValue(Environment.Value("sPath") + "\TestData\AutomationXML\MenuXML\ContentManagement_Menu.xml", "NewAuthorClass")
	   Call Fn_MenuOperation("Select",StrMenu)
       Call Fn_ReadyStatusSync(1)
	End If
	'Selecting S1000D Data Module Option from list
    Call Fn_JavaTree_Select("Fn_ContentM_VerifyS1000DDataModule4Field",objCMDialog, "ClassTree","Complete List:S1000D Data Module 4.0/4.1/4.2")
	'Press Next button
	Call Fn_Button_Click("Fn_ContentM_VerifyS1000DDataModule4Field",objCMDialog,"Next")
	'Maximizing [ NewAdministrativeClass ] Dialog
	Call Fn_Window_Maximize("Fn_ContentM_VerifyS1000DDataModule4Field",objCMDialog)
	wait 3
	'Selecting topic type
	If dicS1000DDataModule4Info("TopicType")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Select Topic Type:"
		Call Fn_List_Select("Fn_ContentM_VerifyS1000DDataModule4Field", objCMDialog, "SelectTopicType",dicS1000DDataModule4Info("TopicType"))
	End If
	'Verify Model Identification Code
	If dicS1000DDataModule4Info("ModelIdentifier")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Model Identification:"
		
		value = objCMDialog.JavaEdit("DocumentTitle").GetROProperty("text")
		If value = dicS1000DDataModule4Info("ModelIdentifier") Then
			Fn_ContentM_VerifyS1000DDataModule4Field = True
		Else
			Fn_ContentM_VerifyS1000DDataModule4Field = False
			Exit Function
		End If
	End If
	
	If dicS1000DDataModule4Info("SystemDifferenceCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","System Difference Code:"
		value = objCMDialog.JavaEdit("DocumentTitle").GetROProperty("text")
		If value = dicS1000DDataModule4Info("SystemDifferenceCode") Then
			Fn_ContentM_VerifyS1000DDataModule4Field = True
		Else
			Fn_ContentM_VerifyS1000DDataModule4Field = False
			Exit Function
		End If
	End If
	If dicS1000DDataModule4Info("SystemCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","System Code:"
		value = objCMDialog.JavaEdit("DocumentTitle").GetROProperty("text")
		If value = dicS1000DDataModule4Info("SystemCode") Then
			Fn_ContentM_VerifyS1000DDataModule4Field = True
		Else
			Fn_ContentM_VerifyS1000DDataModule4Field = False
			Exit Function
		End If
	End If
	If dicS1000DDataModule4Info("DisassemblyCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Disassembly Code:"
		value = objCMDialog.JavaEdit("DocumentTitle").GetROProperty("text")
		If value = dicS1000DDataModule4Info("DisassemblyCode") Then
			Fn_ContentM_VerifyS1000DDataModule4Field = True
		Else
			Fn_ContentM_VerifyS1000DDataModule4Field = False
			Exit Function
		End If
	End If
	If dicS1000DDataModule4Info("InformationCode")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Information Code:"
		value = objCMDialog.JavaList("JavaList").GetItem("#0")
		If value = dicS1000DDataModule4Info("InformationCode") Then
			Fn_ContentM_VerifyS1000DDataModule4Field = True
		Else
			Fn_ContentM_VerifyS1000DDataModule4Field = False
			Exit Function
		End If
	End If	
	If dicS1000DDataModule4Info("InformationCodeVariant")<>"" Then
		objCMDialog.JavaStaticText("FieldName").SetTOProperty "label","Information Code Variant:"
		value = objCMDialog.JavaEdit("DocumentTitle").GetROProperty("text")
		If value = dicS1000DDataModule4Info("InformationCodeVariant") Then
			Fn_ContentM_VerifyS1000DDataModule4Field = True
		Else
			Fn_ContentM_VerifyS1000DDataModule4Field = False
			Exit Function
		End If
	End If
	
	If objCMDialog.Exist(5) Then
		Call Fn_Button_Click("Fn_ContentM_VerifyS1000DDataModule4Field",objCMDialog,"Cancel")
	End If
	If Fn_ContentM_VerifyS1000DDataModule4Field = True Then
		Fn_ContentM_VerifyS1000DDataModule4Field = True
		Exit Function
	End If
	Set objCMDialog=Nothing
	Set WshShell=Nothing
End Function
