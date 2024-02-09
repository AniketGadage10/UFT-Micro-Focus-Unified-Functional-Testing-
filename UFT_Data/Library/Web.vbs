Option Explicit
																								'Function List
'************************************************************************************************************************************************************************************************************
'0. Fn_SISW_Web_GetObject()
'1. Fn_Web_MenuOperation()
'2. Fn_Web_ItemBasicCreate()
'3. Fn_Web_NavTreeOperation()
'4. Fn_Web_FolderCreate()
'5. Fn_Web_Logout()
'6. Fn_Web_Login()
'7. Fn_Web_KillProcess()
'8. Fn_Web_ReadyStatusSync()
'9. Fn_Web_LoginErrorMsgVerify()
'10. Fn_Web_ChangePassword()
'11. Fn_Web_SetPerspective()
'12. Fn_Web_CreateNewForm()
'13. Fn_Web_QuickSearch()
'14. Fn_Web_SidePannelLinkOperations()
'15. Fn_Web_ErrorMsgVerify()
'16. Fn_Web_CreateDataset()
'17. Fn_Web_CreateWebLink()
'18. Fn_Web_VerifyProperties()
'19. Fn_Web_UserSettingsOperations()
'20. Fn_Web_CheckOutObject()
'21. Fn_Web_EditProperties()
'22. Fn_Web_SearchTypeSelect()
'23. Fn_Web_SearchResultOperations()
'24. Fn_Web_SearchOperation()
'25. Fn_Web_PasteAs()
'26. Fn_Web_BrowserOperations()
'27. Fn_Web_QuickSearchReasultOperations()
'28. Fn_Web_QuickLinksMenuOperations()
'29. Fn_Web_DeleteObject()
'30. Fn_Web_ColumnManagement()
'31. Fn_Web_DetailsTableOperations()
'32. Fn_Web_ItemDetailsCreate()
'33. Fn_Web_Tab_Login()
'34. Fn_Web_OverviewTabOperation()
'35. Fn_Web_SaveMySearches()
'36. Fn_Web_MySavedSearchesLinkOperation()
'37. Fn_Web_MyWorklistOperations()
'38. Fn_Web_ReviseItem()
'39. Fn_Web_CreateChange()
'40. Fn_Web_InformationVerify()
'41. Fn_Web_WorkflowProcessAssign()
'42. Fn_Web_SetPreference()
'43. Fn_Web_AssignParticipantsOperations()
'44. Fn_Web_IDDisplayRulesOperation()
'45. Fn_Web_SummaryTabOperations()
'46. Fn_Web_FileDownLoadOperations()
'47. Fn_Web_UploadFile()
'48. Fn_Web_CommonTableOperations()
'49. Fn_Web_TabOperations()
'50. Fn_Web_SaveAsObject()
'51. Fn_Web_ChangeOwnership()
'52. Fn_Web_ChangeManagerTreeOperation()
'53. Fn_Web_ImpactAnalysisListOperations()
'54. Fn_Web_ImpactAnalysisTreeOperations()
'55. Fn_Web_CreateNewParagraph()
'56. Fn_Web_CreateClassicChange()
'57. Fn_Web_ExportToExcel()
'58. Fn_Web_AssignProjectsOperations()
'59. Fn_Web_RemoveProjectsOperations()
'60. Fn_Web_BusinessObjectOperations()
'61. Fn_Web_PSConnectionCreate()
'62. Fn_BOMLineSearchResultOperations()
'63. Fn_Web_CompanyOperations()
'64. Fn_ExportObjects()
'65. Fn_SaveAs_XML()
'66. Fn_SISW_Web_AuditLogOperations()
'67. Fn_SISW_Web_MassUpdateOperations
'68. Fn_SISW_Web_MassUpdateSearchOperations
'69. Fn_SISW_Web_MassUpdateResultOperations
'70. Fn_SISW_Web_WhereUsedParentAssembliesTableOperations
'71. Fn_SISW_Web_AssignFinishOperation
'72. Fn_SISW_Web_ProjectDataTabOperation
'73. Fn_SISW_Web_MakeFromOperations
'74. Fn_SISW_Web_NewItemAssignToProgramOperations
'75. Fn_SISW_Web_PropertiesOnRelation
'76. Fn_SISW_Web_EditStandardNoteOperations
'77. Fn_SISW_Web_FinishesTabOperation
'78. Fn_SISW_Web_BOMComapreOperations
'79. Fn_SISIW_Web_BOMCompareResultTableOperation
'80. Fn_Web_CommonModifiableProperties
'81. Fn_Web_Login_WithoutInvoke_Browser
'************************************************************************************************************************************************************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_Web_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_GetObject("NewBusinessObject")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Priyanka Bhave		 29-Oct-2012		1.0			Pranav S.				
'	Pooja B.
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_Web_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Web.xml"
	Set Fn_SISW_Web_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'************************************************************************************************************************************************************************************************************
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_MenuOperation
'@@
'@@    Description				 :	Function Used to Call Web Menu's
'@@
'@@    Parameters			   :	1.strMenu: Menu Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_MenuOperation("Select","Tools:Check-In/Out:Check-Out...")
'@@												Call Fn_Web_MenuOperation("Select","View:Properties")
'@@												
'@@												Call Fn_Web_MenuOperation("MenuVerify","Edit~Properties:Cut:Copy:Paste:Paste As...:Delete")
'@@												Call Fn_Web_MenuOperation("MenuVerify","New~ID^Alternate ID... : ID^Alias ID... : Item...")
'@@												Call Fn_Web_MenuOperation("IsChecked","View:Pack All")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@												Pranav Ingle				  						8-Apr-2011						1.0							Added Case "MenuVerify"			    	Sunny Ruparel
'@@												Koustubh Watwe				  					   12-May-2011						1.0							Added Case "IsChecked"
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_MenuOperation(strAction,strMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_MenuOperation"
   Dim arrMenu,iCounter,iCount,arrMenuItems,arrMenuInnerItems,bReturn
	Fn_Web_MenuOperation=False
	
	Select Case strAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "SelectExt"
			arrMenu=Split(strMenu,":")
			For iCounter=0 To Ubound(arrMenu)
				If iCounter>0 Then
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"text",arrMenu(iCounter))
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"index",1)
					If Browser("TeamcenterWeb").Link("MenuLink").Exist(WEB_MICROLESS_TIMEOUT) Then
					Else
						Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"index",0)
					End If
				End If
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"text",arrMenu(iCounter))
				Call Fn_Web_UI_Link_Click("Fn_Web_MenuOperation", Browser("TeamcenterWeb"), "MenuLink", "","","")
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"index",0)
			Next
			Fn_Web_MenuOperation=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Select"
			arrMenu=Split(strMenu,":")	
			If UBound(arrMenu)=2 Then
				Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",arrMenu(0)
				Browser("TeamcenterWeb").Link("MenuLink").Click
				wait(WEB_MICRO_TIMEOUT)
				Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",arrMenu(1)
				Browser("TeamcenterWeb").Link("MenuLink").Click
				wait(WEB_MICRO_TIMEOUT)
				Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",arrMenu(2)
                			If lcase(arrMenu(2))="standard note" or  lcase(arrMenu(2))="custom note" Then
					Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "index",1
				End If
				Browser("TeamcenterWeb").Link("MenuLink").Click
               			Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "index",0
				Fn_Web_MenuOperation=True
				Exit Function
			End If

			For iCounter=0 To Ubound(arrMenu)
				If  Instr(arrMenu(0),"View") > 0 Then
					If Instr(arrMenu(iCounter),"Propert") > 0 Then
						 Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"index",1)
						 arrMenu(iCounter) = "Properties(\.\.\.)?" 
					ElseIf Instr(arrMenu(iCounter),"Requirements") > 0 Then
						Call Fn_Web_UI_Link_Click("Fn_Web_MenuOperation", Browser("TeamcenterWeb"), "Requirements", "","","")
						Exit For
					End If
				End If
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"text",arrMenu(iCounter))
				Call Fn_Web_UI_Link_Click("Fn_Web_MenuOperation", Browser("TeamcenterWeb"), "MenuLink", "","","")
				wait WEB_MICRO_TIMEOUT
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"index",0)
			Next			
			Fn_Web_MenuOperation=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "MenuVerify"
         			arrMenu=Split(strMenu,"~")
			Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",trim(arrMenu(0))
			Browser("TeamcenterWeb").Link("MenuLink").Click 1,1,micLeftBtn

			arrMenuItems = Split(arrMenu(1),":")
			For iCount = 0 to UBound(arrMenuItems)
				arrMenuInnerItems=Split(arrMenuItems(iCount),"^")
				If UBound(arrMenuInnerItems)=1 Then
						Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",trim(arrMenuInnerItems(0))
						Browser("TeamcenterWeb").Link("MenuLink").Click 1,1,micLeftBtn 
						Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",trim(arrMenuInnerItems(1))
						bReturn = Fn_Web_UI_ObjectExist("Fn_Web_MenuOperation", Browser("TeamcenterWeb").Link("MenuLink"))
						If bReturn=True Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Found  Menu Item " + arrMenuInnerItems(1) + " in Menu Bar "+arrMenu(0))	
						Else
							Exit Function
						End If
				Else
						Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",trim(arrMenuInnerItems(0))
						If  Instr(arrMenu(0),"View") > 0 Then
							If Instr(trim(arrMenuInnerItems(0)),"Propert") > 0 Then
								 Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"index",1)
							End If
						End If
						bReturn = Fn_Web_UI_ObjectExist("Fn_Web_MenuOperation", Browser("TeamcenterWeb").Link("MenuLink"))
						 Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_MenuOperation",Browser("TeamcenterWeb").Link("MenuLink"),"index",0)

						If bReturn=True Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Found  Menu Item " + arrMenuInnerItems(0) + " in Menu Bar "+arrMenu(0))	
						Else
							Exit Function
						End If
				End If
			Next
			Fn_Web_MenuOperation=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "IsChecked"
	                  arrMenu = Split(strMenu,":")
	                  For iCounter=0 To Ubound(arrMenu)
				Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",arrMenu(iCounter)
				If Browser("TeamcenterWeb").Link("MenuLink").Exist(5) Then
					If iCounter = Ubound(arrMenu) Then
						If instr(Browser("TeamcenterWeb").Link("MenuLink").GetROProperty("class"),"yuimenuitemlabel-checked-selected") <> 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_MenuOperation : Menu [" & strMenu & "] is [ Unchecked ].")
							Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",arrMenu(0)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_MenuOperation : Closing menu [ " & arrMenu(0) & " ].")
							Fn_Web_MenuOperation = False
						ElseIf instr(Browser("TeamcenterWeb").Link("MenuLink").GetROProperty("class"),"yuimenuitemlabel-checked") <> 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_MenuOperation : Menu [" & strMenu & "] is [ Checked ].")
							Browser("TeamcenterWeb").Link("MenuLink").SetTOProperty "text",arrMenu(0)
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_MenuOperation : Closing menu [ " & arrMenu(0) & " ].")
							Fn_Web_MenuOperation = True
						End If
					End If
					'Browser("TeamcenterWeb").Link("MenuLink").Click
                              		Call Fn_Web_UI_Link_Click(sFunctionName, Browser("TeamcenterWeb"), "MenuLink",  "", "", "")
					wait WEB_MICRO_TIMEOUT
				Else
					For iCount = 0 to iCounter
						If sMenu = "" Then
							sMenu = arrMenu(iCount)
						Else
							sMenu = sMenu & ":" & arrMenu(iCount)
						End If
					Next
		                        	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_MenuOperation : Menu [" & sMenu & "] does not exist.")
				      	Exit function
				End If
			Next
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_MenuOperation : Invalid case [ " & strAction & " ].")
	End Select
	If Fn_Web_MenuOperation then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_MenuOperation : Executed successfully with case [ " & strAction & " ].")
	End If
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ItemBasicCreate
'@@
'@@    Description				 :	Function Used to Create Basic Item
'@@
'@@    Parameters			   :	1.strType : Item Type
'@@												  2.strID : Item ID
'@@												  3.strRev : Item Revision														
'@@												  4.strName : Item Name
'@@												  5.strDesc : Item Description
'@@												  6.UOM : Unit Of Measure
'@@												  7.AltIDOpt : Create Alternate ID Option
'@@												  8.CheckOutOpt : Check-Out Item Revision on Create Option
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_ItemBasicCreate("Item","","S","FirstItem","Function Test","","","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ItemBasicCreate(strType,strID,strRev,strName,strDesc,UOM,AltIDOpt,CheckOutOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ItemBasicCreate"
	Dim ObjItem,objobjMDR,strMenu,crrType,iCounter
	Dim objElement, intIndex
	Dim objMyTcPage

   	Fn_Web_ItemBasicCreate=False

	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	Set ObjItem = Fn_SISW_Web_GetObject("NewItem")
	'Vallari [14Jun11] - Get the number of intances of New Item dialog
    	Set objElement = Description.Create()
	objElement("micclass").Value = "WebElement"
	objElement("innertext").Value = "New Item"
	objElement("html tag").Value = "SPAN"
'	intIndex =  Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objElement).count
'	Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem").SetTOProperty "index", cstr(intIndex)
	
	intIndex =  ObjMyTcPage.ChildObjects(objElement).count
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ItemBasicCreate",ObjItem,"index",cstr(intIndex))
	
	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "NewItem")
	
	If  Not ObjItem.Exist(SISW_MIN_TIMEOUT) Then
	'If New Item does not exist, Do menu operation
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(2)
'		'Vallari [14Jun11] - Get the number of intances of New Item dialog and set the index for WebTable in OR accordingly
'		intIndex =  Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objElement).count
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem").SetTOProperty "index", cstr(cint(intIndex)-1)
		intIndex =  ObjMyTcPage.ChildObjects(objElement).count
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ItemBasicCreate",ObjItem,"index",cstr(cint(intIndex)-1))
	End If

	Set objElement = Nothing
	'Set ObjItem=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem")

	'Code added as RequirmentSpec item type changed to Requirement Specification
	If strType="RequirementSpec" Then
		strType="Requirement Specification"
	End If

	If strType<>"" Then
		crrType=ObjItem.WebEdit("ItemTypeEdit").GetROProperty("value")
		wait(WEB_MICRO_TIMEOUT)
		If Trim(crrType)<>Trim(strType) Then
			'Setting Item Type
			Call Fn_Web_UI_Button_Click("Fn_Web_ItemBasicCreate",ObjItem,"ItemType")
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ItemBasicCreate",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",strType)
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ItemBasicCreate",ObjMyTcPage.WebElement("FormType"),"innertext",strType)
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
			ObjMyTcPage.WebElement("FormType").Click 1,1,micLeftBtn
			wait(WEB_MICRO_TIMEOUT)
		End If
	End If
	'creating object of Mercury device replay
	Set objobjMDR = CreateObject("Mercury.DeviceReplay")

'	Call Fn_Web_UI_Button_Click("Fn_Web_ItemBasicCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	Call Fn_Web_UI_Button_Click("Fn_Web_ItemBasicCreate", ObjMyTcPage.WebElement("ButtunPanel"), "Next")
	wait(WEB_MINLESS_TIMEOUT)
	If strID <>"" Then
'		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemBasicCreate", ObjItem.WebTable("ItemInfo"), "ID", strID)
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem").WebTable("ItemInfo").WebEdit("ID").Object.focus
		ObjItem.WebTable("ItemInfo").WebEdit("ID").Object.focus
		objobjMDR.SendString strID
	End If
	If strRev<>"" Then
		wait(WEB_MICRO_TIMEOUT)
'		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemBasicCreate", ObjItem.WebTable("ItemInfo"), "Revision", strRev)
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem").WebTable("ItemInfo").WebEdit("Revision").Object.focus
		ObjItem.WebTable("ItemInfo").WebEdit("Revision").Object.focus
		objobjMDR.SendString strRev
	End If
	If strName<>"" Then
		wait(WEB_MICRO_TIMEOUT)
'		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemBasicCreate", ObjItem.WebTable("ItemInfo"), "Name", strName)
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem").WebTable("ItemInfo").WebEdit("Name").Object.focus
		ObjItem.WebTable("ItemInfo").WebEdit("Name").set ""
		ObjItem.WebTable("ItemInfo").WebEdit("Name").Object.focus
		objobjMDR.SendString strName
	End If
	If strDesc<>"" Then
		wait(WEB_MICRO_TIMEOUT)
'		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemBasicCreate", ObjItem.WebTable("ItemInfo"), "Description", strDesc)
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem").WebTable("ItemInfo").WebEdit("Description").Object.focus
		ObjItem.WebTable("ItemInfo").WebEdit("Description").Object.focus
		objobjMDR.SendString strDesc
	End If
	If UOM<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemBasicCreate", ObjItem.WebTable("ItemInfo"), "UOM", UOM)
	End If
	If AltIDOpt<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_Web_ItemBasicCreate", ObjItem.WebTable("ItemInfo"), "CreateALTID", AltIDOpt)
	End If
	If CheckOutOpt<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_Web_ItemBasicCreate", ObjItem.WebTable("ItemInfo"), "CheckOutRevision", CheckOutOpt)
	End If
'	Call Fn_Web_UI_Button_Click("Fn_Web_ItemBasicCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	Call Fn_Web_UI_Button_Click("Fn_Web_ItemBasicCreate", ObjMyTcPage.WebElement("ButtunPanel"), "Finish")
	Call Fn_Web_ReadyStatusSync(2)
	For iCounter=0 To 2
		If ObjItem.Exist(1) Then
			wait(WEB_MICRO_TIMEOUT)
		Else
			Exit For
		End If
	Next
	Fn_Web_ItemBasicCreate=True
	Set objMyTcPage = Nothing
	Set objobjMDR =Nothing
	Set ObjItem=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_NavTreeOperation
'@@
'@@    Description				 :	Function Used to Perform Operation On Nav Tree
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@											:	 2: strNode Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_NavTreeOperation("Expand","Home:Newstuff")
'@@												Call Fn_Web_NavTreeOperation("Collapse","Home:Newstuff")
'@@												Call Fn_Web_NavTreeOperation("Select","Home:Newstuff")
'@@												Call Fn_Web_NavTreeOperation("MultiSelect","Home:Newstuff~Home:Mailbox")
'@@												Call Fn_Web_NavTreeOperation("Exist","Home:Newstuff")
'@@												Call Call Fn_Web_NavTreeOperation("ClickLink","Home:TestLink Go")
'@@												Call Fn_Web_NavTreeOperation("DoubleClick","Home:Newstuff")
'@@												Call Fn_Web_NavTreeOperation("GetNode","Home:Newstuff")
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									15-Apr-2011						1.1								Case "MultiSelect"							Sunny Ruparel
'@@												Sandeep Navghane									20-Apr-2011						1.2								Case "ClickLink"							Sunny Ruparel
'@@												Sandeep Navghane									27-Sep-2011						1.3								Case "DoubleClick"							Sunny Ruparel
'@@												Ashok kakade												08-Nov-12						1.3								Case "Select 	"	modified for second instanance 	Rupali Palhade
'@@												Ashok kakade												08-Nov-12						1.3								Case "GetNode	"							Rupali Palhade
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_Web_NavTreeOperation(strAction,strNode)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_NavTreeOperation"
   	Dim rowid,obj,iCounter,arrNode,iLength,objSelectType,objSelectType1,intNoOfObjects,iniWTCount,intNoOfObjects1,finWTCount,iHeight,i,arrNode1, arrNode2
   	Dim iNum,arrItem,objWebtable,objChild,StrNode1,strNode2,iCnt1,icnt,bFlag,iInstance,arrInstance,iTempInstance
	Dim nodeName, nodeName1,iCount
	Dim sABS_X, sABS_Y, sHeight
	Dim objMyTcPage, sParentPath
   	Fn_Web_NavTreeOperation=False
   	Set objMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	Select Case strAction
'		Case "Expand","Collapse"
'			Set objSelectType=description.Create()
'			objSelectType("micClass").value = "WebTable"
'			objSelectType("height").value = 0
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType)
'			Set objSelectType1=description.Create()
'			objSelectType1("micClass").value = "WebTable"
'			Set  intNoOfObjects1 = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType1)
'			iniWTCount =  intNoOfObjects1.Count - intNoOfObjects.Count
'			arrNode=Split(strNode,":")
'			iLength=UBound(arrNode)
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").SetTOProperty "text",arrNode(iLength)
'			arrNode2 = Split(arrNode(iLength)," ")
'			'-------------- 04-May-2012---------------- Sandeep Navghane
'			If instr(1,arrNode2(0),".*") Then
'				arrNode2(0)=replace(arrNode2(0),".*","")
'				arrNode(iLength)=arrNode2(0)
'			End If
'
'			'-------------- 04-May-2012---------------- Sandeep Navghane
'			bFlag=False
'			If UBound(arrNode2) > 0 Then
'				rowid = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").GetRowWithCellText(arrNode2(0))
'  			Else
'				rowid = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").GetRowWithCellText(arrNode(iLength))
'			End If
'            Set obj = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").ChildItem(rowid,iLength+1,"WebElement",0)
'			'Temp workaround implemented by : sandeep navghane
'			If lcase(typename(obj))="cowebelement" Then
'				obj.click
'				wait(5)
'				Fn_Web_NavTreeOperation=True
'			Else
'				Call Fn_Web_NavTreeOperation("Select","Home")
'				wait(3)
'				Set obj = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").ChildItem(rowid,iLength+1,"WebElement",0)
'				obj.click
'				wait(5)
'				bFlag=True
'				Call Fn_Web_NavTreeOperation("Select",strNode)
'				Fn_Web_NavTreeOperation=True
'			End If
'			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Set objSelectType=description.Create()
'			objSelectType("micClass").value = "WebTable"
'			objSelectType("height").value = 0
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType)
'			Set objSelectType1=description.Create()
'			objSelectType1("micClass").value = "WebTable"
'			Set  intNoOfObjects1 = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType1)
'			 finWTCount =  intNoOfObjects1.Count - intNoOfObjects.Count
'			 If strAction = "Expand" Then
''				 If bFlag=False Then
'					 If finWTCount < iniWTCount Then
'						 obj.click
'						 wait(2)
'						 Fn_Web_NavTreeOperation=True
'					 End If
''				 End If
'			Elseif strAction = "ExpandExt"  then
'				 If bFlag=False Then
'					 If finWTCount < iniWTCount Then
'						 obj.click
'						wait(2)
'						Fn_Web_NavTreeOperation=True
'					 End If
'				 End If
'			 Else
'				 If finWTCount > iniWTCount Then
'					 obj.click
'					wait(2)
'					Fn_Web_NavTreeOperation=True
'				 End If
'			 End If
'			 Set objSelectType = Nothing
'			 Set objSelectType1 = Nothing
'			 Set obj = Nothing
'			 
'		Case "Select"
'			arrNode=Split(strNode,":")
'			iLength=UBound(arrNode)
'			Set objSelectType=description.Create()
'			Set objSelectType1=description.Create()
'			objSelectType("micClass").value = "WebTable"
'			objSelectType("innertext").value =arrNode(iLength)
'			objSelectType1("micClass").value = "WebElement"
'			objSelectType1("innertext").value = arrNode(iLength)
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType)
'			For i=0 to intNoOfObjects.Count-1
'					iHeight = intNoOfObjects(i).getroproperty("height")
'					If iHeight > 0 Then
'						Set intNoOfObjects1 =  intNoOfObjects(i).ChildObjects(objSelectType1)
'						intNoOfObjects1(2).Click 1,1
'						Fn_Web_NavTreeOperation=True
'					End If
'			Next

		Case "MultiSelect"
			Dim myDeviceReplay, objWebEle
			'Browser("TeamcenterWeb").Page("MyTeamCenter").Object.focus
			objMyTcPage.Object.focus
			Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
			arrNode=Split(strNode,"~")
			arrNode1=Split(arrNode(0),":")
			iLength=UBound(arrNode1)
			iHeight = 10
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_NavTreeOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName"),"text",arrNode1(iLength))
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_NavTreeOperation",objMyTcPage.WebTable("NavTreeNodeName"),"text",arrNode1(iLength))
'			Set objWebEle = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").WebElement("innertext:="&arrNode1(iLength), "index:=0")
			Set objWebEle = objMyTcPage.WebTable("NavTreeNodeName").WebElement("innertext:="&arrNode1(iLength), "index:=0")
			
			'[TC1123-20161031-11_11_2016-VivekA-HCMaintenance] ---------------------------
			sABS_X = objWebEle.GetROProperty("abs_x")
			sABS_Y = cInt(objWebEle.GetROProperty("abs_y"))
			sHeight = objWebEle.GetROProperty("height")
			'------------------------------------------------------------------------------------------------------
			myDeviceReplay.MouseMove  cInt(sABS_X) + 150,  cInt(sABS_Y) + (cInt(sHeight)/2)
			wait(WEB_MIN_TIMEOUT)
			myDeviceReplay.MouseClick  cInt(sABS_X) + 150,  cInt(sABS_Y) + (cInt(sHeight)/2), 0
			myDeviceReplay.KeyDown 29
			wait(WEB_MICRO_TIMEOUT)
			For iCounter=1 To UBound(arrNode)
				arrNode1=Split(arrNode(iCounter),":")
				iLength=UBound(arrNode1)
'				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_NavTreeOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName"),"text",arrNode1(iLength))
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_NavTreeOperation",objMyTcPage.WebTable("NavTreeNodeName"),"text",arrNode1(iLength))
'				Set objWebEle = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").WebElement("innertext:="&arrNode1(iLength), "index:=0")			
				Set objWebEle = objMyTcPage.WebTable("NavTreeNodeName").WebElement("innertext:="&arrNode1(iLength), "index:=0")			
				'[TC1123-20161031-11_11_2016-VivekA-HCMaintenance] ---------------------------
				sABS_X = objWebEle.GetROProperty("abs_x")
				sABS_Y = cInt(objWebEle.GetROProperty("abs_y"))
				sHeight = objWebEle.GetROProperty("height")
				'------------------------------------------------------------------------------------------------------
				myDeviceReplay.MouseMove  cInt(sABS_X) + 150,  cInt(sABS_Y) + (cInt(sHeight)/2)
				wait(WEB_MIN_TIMEOUT)
				myDeviceReplay.MouseClick  cInt(sABS_X) + 150,  cInt(sABS_Y) + (cInt(sHeight)/2), 0
			Next
			myDeviceReplay.KeyUp 29
			Fn_Web_NavTreeOperation=True
'		Case "Exist"
'			arrNode=Split(strNode,":")
'			iLength=UBound(arrNode)
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_NavTreeOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName"),"text",arrNode(iLength))
'			If Fn_Web_UI_ObjectExist("Fn_Web_NavTreeOperation", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName"))=True Then
'				If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").GetROProperty("height") > 0 Then
'					Fn_Web_NavTreeOperation=True
'				Else
'					Call Fn_Web_NavTreeOperation("Select","Home")
'					wait 2
'					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_NavTreeOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName"),"text",arrNode(iLength))
'					wait 2
'					If Fn_Web_UI_ObjectExist("Fn_Web_NavTreeOperation", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName"))=True Then
'						If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").GetROProperty("height") > 0 Then
'							Fn_Web_NavTreeOperation=True
'							Call Fn_Web_NavTreeOperation("Select",strNode)
'						End if
'					End if
'				End If
'			End If
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		Case "ClickLink"
			iInstance=1
			iTempInstance=1
			arrInstance=Split(strNode,"~")
			If ubound(arrInstance)=1 Then
				iInstance=arrInstance(1)
			End If
			arrNode=Split(arrInstance(0),":")

			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
			'Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType)
			Set  intNoOfObjects = objMyTcPage.WebElement("NavTreePanel").ChildObjects(objSelectType)
			iLength=0
			For iCounter=0 to ubound(arrNode)
				bFlag=False
				For i=iLength to intNoOfObjects.Count-1
					nodeName = Split( trim(intNoOfObjects(i).getroproperty("innertext")) ," [ ")
					If instr( trim(arrNode(iCounter)) ,".*") Then
						nodeName1 = Split( trim(arrNode(iCounter)) ,".*")
						arrNode(iCounter) = nodeName1(0)
					ElseIf instr( trim(arrNode(iCounter)) ," [ ") Then
						nodeName1 = Split( trim(arrNode(iCounter)) ," [ ")
							arrNode(iCounter) = nodeName1(0)
					End If
					If trim(arrNode(iCounter))=trim(nodeName(0)) Then
						iHeight = intNoOfObjects(i).getroproperty("height")
						If iHeight > 0 Then
							If iCounter=cint(ubound(arrNode)) Then
                               					 If Cint(iTempInstance)=Cint(iInstance) Then
									Set  intNoOfObjects1 = intNoOfObjects(i).ChildItem(1,cint(ubound(arrNode))+2,"Link",0) 
									If typename(intNoOfObjects1)="Nothing" Then
										Set objMyTcPage = Nothing
										Set  intNoOfObjects1 = Nothing
										Set  objSelectType1 = Nothing
										Set objSelectType=Nothing
										Exit function
									End If
									intNoOfObjects1.Click
									wait WEB_MICROLESS_TIMEOUT
									Set  intNoOfObjects1 = Nothing
									Set  objSelectType1 = Nothing
								End if
							End If
							iLength=i+1
							If iCounter=cint(ubound(arrNode)) Then
								If Cint(iTempInstance)=Cint(iInstance) Then
									bFlag=True
									Exit for
								End If
								iTempInstance=iTempInstance+1
							Else
								bFlag=True
								Exit for
							End if
						Else
							bFlag=False
							Exit for
						End If
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_Web_NavTreeOperation=True
			End If
			Set objSelectType=Nothing
'			arrNode=Split(strNode,":")
'			iLength=UBound(arrNode)
'           		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").SetTOProperty "text",arrNode(iLength)+".*"
'			rowid = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").GetRowWithCellText(arrNode(iLength))
'			Set obj = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").ChildItem(rowid,iLength+2,"Link",0) 
'			wait(1)
'			obj.click
'			wait(1)
'			Fn_Web_NavTreeOperation=True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "DoubleClick"
			arrNode=Split(strNode,":")
			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
			'Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType)
			Set  intNoOfObjects = objMyTcPage.WebElement("NavTreePanel").ChildObjects(objSelectType)
			iLength=0
			For iCounter=0 to ubound(arrNode)
				bFlag=False
				For i=iLength to intNoOfObjects.Count-1
					If trim(arrNode(iCounter))=trim(intNoOfObjects(i).getroproperty("innertext")) Then
						iHeight = intNoOfObjects(i).getroproperty("height")
						If iHeight > 0 Then
							If iCounter=cint(ubound(arrNode)) Then
								Set objSelectType1=description.Create()
								objSelectType1("micClass").value = "WebElement"
                               					 If  instr(1,arrNode(iCounter),"[ View | Edit ]" ) Then
									arrNode(iCounter)=Replace(arrNode(iCounter)," [ View | Edit ]",".*")
									objSelectType1("innertext").RegularExpression = True
									objSelectType1("innertext").value = arrNode(iCounter)
								Else
									objSelectType1("innertext").value = arrNode(iCounter)
								End If
								Set  intNoOfObjects1 = intNoOfObjects(i).ChildObjects(objSelectType1)
								intNoOfObjects1(2).Click 1,1
								wait WEB_MICROLESS_TIMEOUT
								Set objMDR = CreateObject("Mercury.DeviceReplay")
								iX=intNoOfObjects1(2).GetROProperty("abs_x")
								iY=intNoOfObjects1(2).GetROProperty("abs_y")
								iH=intNoOfObjects1(2).GetROProperty("height")'Calculate Height
								iW=intNoOfObjects1(2).GetROProperty("width")'Calculate Width
								objMDR.MouseDblClick Cint(iX+Cint(iW/2)),Cint(iY+Cint(iH/2)+10),LEFT_MOUSE_BUTTON  
								wait WEB_MICROLESS_TIMEOUT
								Set objMDR = Nothing
								Set  intNoOfObjects1 = Nothing
								Set  objSelectType1 = Nothing
							End If
							iLength=i+1
							bFlag=True
							Exit for
						Else
							bFlag=False
							Exit for
						End If
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_Web_NavTreeOperation=True
			End If
			Set objSelectType=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	    Case "GetNode"				
			iNum = 0
			Fn_Web_NavTreeOperation = ""
			arrItem = Split(strNode,":",-1,1)
			iCounter = ubound(arrItem) + 1
			Set objWebtable=Description.Create()
			objWebtable("html tag").value="TD"
'			Set objChild=Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objWebtable)    	
			Set objChild=objMyTcPage.ChildObjects(objWebtable) 
		
			For iCnt1=3 to objChild.count-1		
				StrNode1 =objChild(iCnt1).GetROProperty("innertext")
				If trim(arrItem(iNum)) = trim(StrNode1) Then			
					If iNum=0 Then
						strNode2 = arrItem(iNum)		
						iNum =iNum +1
					Else
						strNode2 = strNode2 +":"+ arrItem(iNum)		 				
						iNum =iNum +1
							If iNum = iCounter Then
								iNum =iNum - 1
							End If
					End If       
					If trim(strNode) = trim(strNode2) Then							
						For icnt =iCnt1+1 to objChild.count-1						
							StrNode1 =objChild(icnt).GetROProperty("innertext")							
							If  trim(StrNode1) = "" Then
									''continue
							Else
								strNode2 = strNode2 +":"+StrNode1
								Fn_Web_NavTreeOperation =  StrNode1
								iCnt1 =  objChild.count-1
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Found Node Name ["+StrNode1+"] ")
								Exit For	
								Exit Function									
							End If							
						Next
					End If
				End If
			Next
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		Case "Expand_Mass_Update", "Expand_Mass_Collapse"
			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
			objSelectType("height").value = 0
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType)
			Set  intNoOfObjects = objMyTcPage.ChildObjects(objSelectType)
			Set objSelectType1=description.Create()
			objSelectType1("micClass").value = "WebTable"
'			Set  intNoOfObjects1 = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType1)
			Set  intNoOfObjects1 = objMyTcPage.ChildObjects(objSelectType1)
			iniWTCount =  intNoOfObjects1.Count - intNoOfObjects.Count
			arrNode=Split(strNode,"~")
			iLength=UBound(arrNode)
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").SetTOProperty "text",arrNode(iLength)
			objMyTcPage.WebTable("NavTreeNodeName").SetTOProperty "text",arrNode(iLength)
			arrNode2 = Split(arrNode(iLength)," ")
			If instr(1,arrNode2(0),".*") Then
				arrNode2(0)=replace(arrNode2(0),".*","")
				arrNode(iLength)=arrNode2(0)
			End If
			If UBound(arrNode2) > 0 Then
'				rowid = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").GetRowWithCellText(arrNode2(0))
				rowid = objMyTcPage.WebTable("NavTreeNodeName").GetRowWithCellText(arrNode2(0))
  			Else
'				rowid = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").GetRowWithCellText(arrNode(iLength))
				rowid = objMyTcPage.WebTable("NavTreeNodeName").GetRowWithCellText(arrNode(iLength))
			End If
'            		Set obj = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").ChildItem(rowid,iLength+1,"WebElement",0)
            		Set obj = objMyTcPage.WebTable("NavTreeNodeName").ChildItem(rowid,iLength+1,"WebElement",0)
			obj.click
			wait(WEB_MIN_TIMEOUT)
			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
			objSelectType("height").value = 0
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType)
			Set  intNoOfObjects = objMyTcPage.ChildObjects(objSelectType)
			Set objSelectType1=description.Create()
			objSelectType1("micClass").value = "WebTable"
'			Set  intNoOfObjects1 = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType1)
			Set  intNoOfObjects1 = objMyTcPage.ChildObjects(objSelectType1)
			 finWTCount =  intNoOfObjects1.Count - intNoOfObjects.Count
			 If strAction = "Expand_Mass_Update" Then
				 If finWTCount < iniWTCount Then
					 obj.click
					wait(WEB_MICROLESS_TIMEOUT)
				 End If
			 Else
				 If finWTCount > iniWTCount Then
					 obj.click
					wait(WEB_MICROLESS_TIMEOUT)
				 End If
			 End If
			 Set objSelectType = Nothing
			 Set objSelectType1 = Nothing
			 Set obj = Nothing
			 Fn_Web_NavTreeOperation=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		Case "Select_Mass_Update"
          		iCounter = 1
		  	iCnt =1
		   	 If instr(strNode, "@") > 0 Then
				iNum = Split(strNode,"@") 
				iCnt =iNum(1)
				strNode = iNum(0)
			End If
			arrNode=Split(strNode,"~")
			iLength=UBound(arrNode)
			Set objSelectType=description.Create()
			Set objSelectType1=description.Create()
			objSelectType("micClass").value = "WebTable"
			objSelectType("innertext").value =arrNode(iLength)
			objSelectType1("micClass").value = "WebElement"
			objSelectType1("innertext").value = arrNode(iLength)
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType)
			Set  intNoOfObjects = objMyTcPage.ChildObjects(objSelectType)

			For i=0 to intNoOfObjects.Count-1
				iHeight = intNoOfObjects(i).getroproperty("height")
				If iHeight > 0 Then					
					 If cInt(iCnt) = iCounter then
						Set intNoOfObjects1 =  intNoOfObjects(i).ChildObjects(objSelectType1)
						intNoOfObjects1(2).Click 1,1
					 End If
					 iCounter = iCounter + 1
				End If
			Next
			Fn_Web_NavTreeOperation=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		Case "Exist_Mass_Update"
			arrNode=Split(strNode,"~")
			iLength=UBound(arrNode)
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_NavTreeOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName"),"text",arrNode(iLength))
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_NavTreeOperation",objMyTcPage.WebTable("NavTreeNodeName"),"text",arrNode(iLength))
'			If Fn_Web_UI_ObjectExist("Fn_Web_NavTreeOperation", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName"))=True Then
'				If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NavTreeNodeName").GetROProperty("height") > 0 Then
'					Fn_Web_NavTreeOperation=True
'				End If
'			End If
			If Fn_Web_UI_ObjectExist("Fn_Web_NavTreeOperation", objMyTcPage.WebTable("NavTreeNodeName"))=True Then
				If objMyTcPage.WebTable("NavTreeNodeName").GetROProperty("height") > 0 Then
					Fn_Web_NavTreeOperation=True
				End If
			End If
			Set objChild = nothing
			Set objWebtable = nothing
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		Case "SelectExt","Select"
			iInstance=1
			iTempInstance=1
			arrInstance=Split(strNode,"~")
			If ubound(arrInstance)=1 Then
				iInstance=arrInstance(1)
			End If
			arrNode=Split(arrInstance(0),":")

			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType)
			Set  intNoOfObjects = objMyTcPage.WebElement("NavTreePanel").ChildObjects(objSelectType)
			iLength=0
			For iCounter=0 to ubound(arrNode)
				bFlag=False
				For i=iLength to intNoOfObjects.Count-1
					nodeName = Split( trim(intNoOfObjects(i).getroproperty("innertext")) ," [ ")
					If instr( trim(arrNode(iCounter)) ,".*") Then
						nodeName1 = Split( trim(arrNode(iCounter)) ,".*")
						arrNode(iCounter) = nodeName1(0)
					ElseIf instr( trim(arrNode(iCounter)) ," [ ") Then
						nodeName1 = Split( trim(arrNode(iCounter)) ," [ ")
						arrNode(iCounter) = nodeName1(0)
					End If
					If trim(arrNode(iCounter))=trim(nodeName(0)) Then
						iHeight = intNoOfObjects(i).getroproperty("height")
						If iHeight > 0 Then
							If iCounter=cint(ubound(arrNode)) Then
                               					 If Cint(iTempInstance)=Cint(iInstance) Then
									Set objSelectType1=description.Create()
									objSelectType1("micClass").value = "WebElement"
									 If  instr(1,arrNode(iCounter),"[ View | Edit ]" ) Then
										arrNode(iCounter)=Replace(arrNode(iCounter)," [ View | Edit ]",".*")
										objSelectType1("innertext").RegularExpression = True
										objSelectType1("innertext").value = arrNode(iCounter)
									ELSEIf  instr(1,arrNode(iCounter),"[ View ]" ) Then
										arrNode(iCounter)=Replace(arrNode(iCounter)," [ View ]",".*")
										objSelectType1("innertext").RegularExpression = True
										objSelectType1("innertext").value = arrNode(iCounter)
									ELSEIf  instr(1,arrNode(iCounter),"[ Go ]" ) Then
										arrNode(iCounter)=Replace(arrNode(iCounter)," [ Go ]",".*")
										objSelectType1("innertext").RegularExpression = True
										objSelectType1("innertext").value = arrNode(iCounter)
									ELSEIf  instr(1,arrNode(iCounter),".*" ) Then
										objSelectType1("innertext").RegularExpression = True
										objSelectType1("innertext").value = arrNode(iCounter)
									Else
										objSelectType1("innertext").value = arrNode(iCounter)&".*"
									End If
									Set  intNoOfObjects1 = intNoOfObjects(i).ChildObjects(objSelectType1)
									intNoOfObjects1(2).Click 1,1
									wait WEB_MICROLESS_TIMEOUT
									Set  intNoOfObjects1 = Nothing
									Set  objSelectType1 = Nothing
								End if
							End If
							iLength=i+1
							If iCounter=cint(ubound(arrNode)) Then
								If Cint(iTempInstance)=Cint(iInstance) Then
									bFlag=True
									Exit for
								End If
								iTempInstance=iTempInstance+1
							Else
								bFlag=True
								Exit for
							End if
						Else
							bFlag=False
							Exit for
						End If
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_Web_NavTreeOperation=True
			End If
			Set objSelectType=Nothing
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		Case "ExistExt","Exist"
			arrNode=Split(strNode,":")
'			iLength=UBound(arrNode)
			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType)
			Set  intNoOfObjects = objMyTcPage.WebElement("NavTreePanel").ChildObjects(objSelectType)
			iLength=0
			For iCounter=0 to ubound(arrNode)
				bFlag=False
				For i=iLength to intNoOfObjects.Count-1
					If trim(arrNode(iCounter))=trim(intNoOfObjects(i).getroproperty("innertext")) Then
						iHeight = intNoOfObjects(i).getroproperty("height")
						If iHeight > 0 Then
							iLength=i+1
							bFlag=True
							Exit for
						Else
							bFlag=False
							Exit for
						End If
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_Web_NavTreeOperation=True
			End If
			Set objSelectType=Nothing
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		Case "ExpandExt","CollapseExt","Expand","Collapse"
			arrNode=Split(strNode,":")
'			iLength=UBound(arrNode)
			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
'			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("NavTreePanel").ChildObjects(objSelectType)
			Set  intNoOfObjects = objMyTcPage.WebElement("NavTreePanel").ChildObjects(objSelectType)
			iLength=0
			For iCounter=0 to ubound(arrNode)
				bFlag=False
				For i=iLength to intNoOfObjects.Count-1
					If trim(arrNode(iCounter))=trim(intNoOfObjects(i).getroproperty("innertext")) Then
						iHeight = intNoOfObjects(i).getroproperty("height")
						If iHeight > 0 Then
							If iCounter=cint(ubound(arrNode)) Then
								If StrAction="CollapseExt" Then
                                    						StrNode1="collapsed"
								Else
									StrNode1="expanded"
								End If
								If instr(1,lcase(intNoOfObjects(i).getroproperty("class")),StrNode1) Then
									'do nothing
									Fn_Web_NavTreeOperation=True
								Else
									rowid = intNoOfObjects(i).GetRowWithCellText(arrNode(iCounter))
									Set obj = intNoOfObjects(i).ChildItem(rowid, cint(ubound(arrNode))+1,"WebElement",0)
									If typename(obj)<>"Nothing" Then
										obj.click
										wait WEB_MINLESS_TIMEOUT
										Fn_Web_NavTreeOperation=True
									End If
								End If
							End if
							iLength=i+1
							bFlag=True
							Exit for
						Else
							bFlag=False
							Exit for
						End If
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			Set objSelectType=Nothing
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		'To Expand nodes upto Last child's parent node in StrNodeName hierarchy and then select the last node
		Case "ExpandAndSelect"
			'Initial Item Path
			arrNode=Split(strNode,":")
			For iCount = 0 to UBound(arrNode)-1
				If sParentPath = "" Then
					sParentPath  = arrNode(iCount)
				Else
					sParentPath  = sParentPath + ":" + arrNode(iCount)
				End If
				Call Fn_Web_NavTreeOperation("Expand", sParentPath)
				If arrNode(iCount) = "AutomatedTests" Then
					Call Fn_Web_ReadyStatusSync(WEB_MICROLESS_TIMEOUT)
				End If
				Call Fn_Web_ReadyStatusSync(WEB_MICRO_TIMEOUT)
			Next
			bFlag = Fn_Web_NavTreeOperation("Select", strNode)
			If bFlag = True Then
				Fn_Web_NavTreeOperation = True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 		
		'To expand all the nodes present in StrNodeName
		Case "ExpandAll"
			'Initial Item Path
			arrNode=Split(strNode,":")
			For iCount = 0 to UBound(arrNode)-1
				If sParentPath = "" Then
					sParentPath  = arrNode(iCount)
				Else
					sParentPath  = sParentPath + ":" + arrNode(iCount)
				End If
				Call Fn_Web_NavTreeOperation("Expand", sParentPath)
				If arrNode(iCount) = "AutomatedTests" Then
					Call Fn_Web_ReadyStatusSync(WEB_MICROLESS_TIMEOUT)
				End If
				Call Fn_Web_ReadyStatusSync(WEB_MICRO_TIMEOUT)
			Next
			bFlag = Fn_Web_NavTreeOperation("Expand", strNode)
			If bFlag = True Then
				Fn_Web_NavTreeOperation = True
			End If
	End Select
	Set objMyTcPage = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_FolderCreate
'@@
'@@    Description				 :	Function Used to Create Folder
'@@
'@@    Parameters			   :	1.strName : Folder Name
'@@											:	 2: strDescription : Folder Description
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_FolderCreate("FunctionTest","Folder to take function Demo")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_FolderCreate(strName,strDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_FolderCreate"
	Fn_Web_FolderCreate=False
	Dim ObjFolder,strMenu
	Dim objButtonPanel
	
'	Set ObjFolder=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewFolder")
	Set ObjFolder = Fn_SISW_Web_GetObject("NewFolder")
	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "NewFolder")
	If Not ObjFolder.Exist(SISW_MIN_TIMEOUT) Then
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(1)
	End If
	
	''-------------------------------------
	If strName<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_FolderCreate", ObjFolder.WebTable("FolderInfo"), "Name", strName)
	End If
	If strDescription<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_FolderCreate", ObjFolder.WebTable("FolderInfo"), "Description", strDescription)
	End If
	Set objButtonPanel = Fn_SISW_Web_GetObject("ButtunPanel")
	Call Fn_Web_UI_Button_Click("Fn_Web_MyTc_FolderCreate", objButtonPanel, "Finish")
	'Call Fn_Web_UI_Button_Click("Fn_Web_FolderCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	Call Fn_Web_ReadyStatusSync(3)
	Fn_Web_FolderCreate=True
	Set ObjFolder = Nothing
	Set objButtonPanel = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_Logout
'@@
'@@    Description				 :	Function Used to Log Out From Web Client
'@@
'@@    Parameters			   :	NA
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_Logout()
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_Logout()
	GBL_FAILED_FUNCTION_NAME="Fn_Web_Logout"
   Dim strBrowser, WshShell
	Fn_Web_Logout=False
	strBrowser=Environment.Value("WebBrowserName")
	
	If Browser("TeamcenterWeb").Link("Logout").Exist(5) Then
		Call Fn_Web_ReadyStatusSync(1)
		Call Fn_Web_UI_Link_Click("Fn_Web_Logout", Browser("TeamcenterWeb"), "Logout", "","","")
		If Browser("TeamcenterWeb").Dialog("LogoutDialog").Exist(8) Then
			If InStr(1,strBrowser,"IE")>0 Then
				Browser("TeamcenterWeb").Dialog("LogoutDialog").WinButton("OK").Click 1,1,micLeftBtn
				
			'Swapnil: IF unable to click on OK button handled through key strokes.
			
				wait 1
				
				If  Browser("TeamcenterWeb").Dialog("LogoutDialog").Exist(2)Then
						Browser("TeamcenterWeb").Dialog("LogoutDialog").Activate
						Set WshShell = CreateObject("WScript.Shell")
						WshShell.SendKeys "{ENTER}"
						Set WshShell = Nothing
				End If
			Else
				Call Fn_Web_UI_Button_Click("Fn_Web_Logout", Browser("TeamcenterWeb").Dialog("LogoutDialog").Page("FFLogoutPage"), "OK")
			End If
			wait(10)
			If Browser("TeamcenterLogin").Page("Logout").WebButton("LoginAgain").Exist(5) Then
				Browser("TeamcenterLogin").Page("Logout").WebButton("LoginAgain").Click
			ElseIf Browser("TeamcenterLogin").Page("Logout").WebTable("Logout").WebButton("LoginAgain").Exist(5) Then
				Browser("TeamcenterLogin").Page("Logout").WebTable("Logout").WebButton("LoginAgain").Click
			End If
			Fn_Web_Logout=True
		End If
	End If

End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_Login
'@@
'@@    Description				 :	Function Used to Log In Web Client
'@@
'@@    Parameters			   :	1.strUserName : User Name
'@@											 	 2.strPassword : Password
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Environment XML should be properly filled							
'@@
'@@    Examples					:	Call Fn_Web_Login("AutoTest1","AutoTest1")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_Login(strUserName,strPassword)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_Login"
	Dim strBrowser, strServerPath, ObjBrowser, handlewin, bFlag
	Dim iIterate

	Fn_Web_Login = False
	strBrowser = Environment.Value("WebBrowserName")
	If Browser("TeamcenterWeb").Link("Logout").Exist(5) Then
		Call Fn_Web_Logout()
		Call Fn_Web_KillProcess("")
	End If
	''//Killing Intenet Explorer forcibly
	SystemUtil.CloseProcessByName("iexplore.exe")
	strServerPath = Environment.Value("TcWebServer")
	
	If InStr(1,strBrowser,"IE")>0 Then
		Set ObjBrowser = CreateObject("InternetExplorer.Application")
		ObjBrowser.Visible = True
		ObjBrowser.Navigate strServerPath
		handlewin = ObjBrowser.HWND
		wait WEB_DEFAULT_TIMEOUT
		Window("hwnd:="+CStr(handlewin)).Maximize
		
		'[TC1017-2016101100-25_10_2016-VivekA-Maintenance] - As per new Hierarchy - Made generalised function for every script
		bFlag = Fn_Web_Login_WithoutInvoke_Browser("Fn_Web_Login", strUserName, strPassword)
'		If bFlag = True Then
'			If Browser("TeamcenterWeb").Link("Logout").Exist(20) Then
'				Call Fn_Web_ReadyStatusSync(2)
'				Fn_Web_Login = True
'			Else
'				Fn_Web_Login = False
'			End If
'		End If
		
		'----------------------------------------------------
	Else
		If Not Browser("version:=.*Firefox.*").Exist(5) Then
			SystemUtil.Run "firefox.exe"
		End If
		Browser("version:=.*Firefox.*").Navigate strServerPath
		If Browser("TeamcenterLogin").Page("Login").Exist(10) Then
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_Login", Browser("TeamcenterLogin").Page("Login").WebTable("Login"), "Username", strUserName)
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_Login", Browser("TeamcenterLogin").Page("Login").WebTable("Login"), "Password", strPassword)
			Call Fn_Web_UI_Button_Click("Fn_Web_Login", Browser("TeamcenterLogin").Page("Login").WebTable("Login"), "Login")
			bFlag = True
'			If  Browser("TeamcenterWeb").Link("Logout").Exist(20) Then
'				Call Fn_Web_ReadyStatusSync(2)
'				Fn_Web_Login = True
'			End If
		End If
	End If
	If bFlag = True Then
'			If Browser("TeamcenterWeb").Link("Logout").Exist(20) Then
'				Call Fn_Web_ReadyStatusSync(2)
'				Fn_Web_Login = True
'			Else
'				Fn_Web_Login = False
'			End If
		For iIterate = 0 to 20
			If Fn_Web_UI_ObjectExist("Fn_Web_Login", Browser("TeamcenterWeb").Link("Logout")) = True Then
				Fn_Web_Login = True
				Exit For
			End If
		Next
	End If
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_KillProcess
'@@
'@@    Description				 :	Function Used to Kill the Processes
'@@
'@@    Parameters			   :	1.strProcess : Process Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	NA							
'@@
'@@    Examples					:	Call Fn_Web_KillProcess("")
'@@												Call Fn_Web_KillProcess("iexplore.exe")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_KillProcess(strProcess)
   Fn_Web_KillProcess=False
	'Restore Image of the Application before killing Application.
	sImgPath = Environment.Value("BatchFldName") +"\"   + Environment.Value("TestName") + ".png"
	Desktop.CaptureBitmap sImgPath,True
	'To close IE browser which is sometimes considered as Window insted of browser
	Call Fn_Web_ErrorMsgVerify("","OK")
	Call Fn_Web_Logout()
	If Browser("TeamcenterLogin").Exist(5) Then
		Browser("TeamcenterLogin").Close()
		wait(2)
		Fn_Web_KillProcess=True
	End If
	SystemUtil.CloseProcessByName("iexplore.exe")
	SystemUtil.CloseProcessByName("firefox.exe")
	Fn_Web_KillProcess=True
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ReadyStatusSync
'@@
'@@    Description				 :	Function Used to Perform Syncronisation
'@@
'@@    Parameters			   :	1.iTeration : Number of Iterations
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	NA							
'@@
'@@    Examples					:	Call Fn_Web_ReadyStatusSync(1)
'@@												Call Fn_Web_ReadyStatusSync(4)
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ReadyStatusSync(iTeration)
   Dim iCounter
   Browser("TeamcenterWeb").Page("MyTeamCenter").Sync
   For iCounter=1 To iTeration
	If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Overview").Exist(20) Then
		wait WEB_MICROLESS_TIMEOUT
		Exit For
	End If
   Next
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_LoginErrorMsgVerify
'@@
'@@    Description				 :	Function Used to Verify Login Error which occured by Invalid Inputs
'@@
'@@    Parameters			   :	1.strErrMsg : Error Message
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	NA							
'@@
'@@    Examples					:	Call Fn_Web_LoginErrorMsgVerify("The login attempt failed: either the user ID or the password is invalid")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									8-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_LoginErrorMsgVerify(strErrMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_LoginErrorMsgVerify"
	GBL_EXPECTED_MESSAGE=strErrMsg
	Fn_Web_LoginErrorMsgVerify=False
	strCurrError=Browser("TeamcenterLogin").Page("Login").WebElement("LoginErrorMsg").GetROProperty("innertext")
	If InStr(1,strCurrError,strErrMsg)>=1 Then
		Fn_Web_LoginErrorMsgVerify=True
	Else
		GBL_ACTUAL_MESSAGE=strCurrError
	End If
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ChangePassword
'@@
'@@    Description				 :	Function Used to Change Password of Current User
'@@
'@@    Parameters			   :	1.strCurrentPassword : Current Password
'@@												 2.strNewPassword : New Password To Set
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log in Web Client							
'@@
'@@    Examples					:	Call Fn_Web_ChangePassword("webuser02","Password123")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									12-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ChangePassword(strCurrentPassword,strNewPassword)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ChangePassword"
	'Initially Function Rerturns False
	Fn_Web_ChangePassword=False
	'Variable Declaration
	Dim ObjPwdDialog,strWEBMenuPath,strMenu
	'Creating Object of "ChangePassword" Table
	Set ObjPwdDialog=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangePassword")
	'Checking Existance of "ChangePassword" Table
	If Not ObjPwdDialog.Exist(7) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "EditChangePassword")
		'Calling Edit->Change Password... Menu Option
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	'Entering Current Password
	If strCurrentPassword<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ChangePassword",ObjPwdDialog, "CurrentPassword", strCurrentPassword)
	End If
	'Setting New Password
	If strNewPassword<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ChangePassword",ObjPwdDialog, "NewPassword", strNewPassword)
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ChangePassword",ObjPwdDialog, "ConfirmPassword", strNewPassword)
	End If
	'Clicking On OK button
	Call Fn_Web_UI_Button_Click("Fn_Web_ChangePassword", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
	'Checking Confirm Password Change Dialog appear or Not
	If Fn_Web_UI_ObjectExist("Fn_Web_ChangePassword", Browser("TeamcenterWeb").Dialog("Dialog"))=True Then
		If InStr(1,Environment.Value("WebBrowserName"),"IE")>0 Then
			Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").Click
		Else
			Browser("TeamcenterWeb").Dialog("Dialog").Page("FFPage").WebButton("OK").Click
		End If
		'Function Returns True
		Fn_Web_ChangePassword=True	
	End If
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SetPerspective
'@@
'@@    Description				 :	Function Used to Set TeamCenter Perspective
'@@
'@@    Parameters			   :	1.strPerspectiveName : Perspective Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log in Web Client							
'@@
'@@    Examples					:	Call Fn_Web_SetPerspective("Structure Manager")
'@@												Call Fn_Web_SetPerspective("My Teamcenter")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									13-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_SetPerspective(strPerspectiveName)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SetPerspective"
	Dim bFlag, iCounter
	'Seeting False Return value To Function
	Fn_Web_SetPerspective=False
	'Setting Name Of Perspective which Need to Set
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SetPerspective",Browser("TeamcenterWeb").Link("PerspectiveLink"),"text",strPerspectiveName)
	If strPerspectiveName = "Design Context" Then
		For iCounter = 0 To 2		
'			Browser("TeamcenterWeb").Link("PerspectiveLink").SetTOProperty "Index", iCounter	
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SetPerspective",Browser("TeamcenterWeb").Link("PerspectiveLink"),"Index",iCounter)
			If Fn_Web_UI_ObjectExist("Fn_Web_SetPerspective", Browser("TeamcenterWeb").Link("PerspectiveLink")) = true then
				Exit For
			End If
		Next
	End If
	bFlag=Fn_Web_UI_Link_Click("Fn_Web_SetPerspective", Browser("TeamcenterWeb"), "PerspectiveLink", "","","")
	If bFlag=True Then
		'Function Returns True
		Fn_Web_SetPerspective=True
	End If
	Wait WEB_MICROLESS_TIMEOUT
'    Browser("TeamcenterWeb").Link("PerspectiveLink").SetTOProperty "Index", 0
    Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SetPerspective",Browser("TeamcenterWeb").Link("PerspectiveLink"),"Index",0)
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CreateNewForm
'@@
'@@    Description				 :	Function Used to Create Forms
'@@
'@@    Parameters			   :	1.strFormType : Form Type
'@@												 2.strFormName : Form Name
'@@												 3.strFormDesc : Form Description
'@@												 4.strURL : Website URL
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log in Web Client							
'@@
'@@    Examples					:	Call Fn_Web_CreateNewForm("ArcWeld Master","Form1","Demo Form","")
'@@												Call Fn_Web_CreateNewForm("AutoForm","Form3","Demo Form","www.google.com")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									13-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_CreateNewForm(strFormType,strFormName,strFormDesc,strURL)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CreateNewForm"
   'Variable Declaration
   Dim ObjForm,ObjFormInfo,ObjMyTcPage
   Dim strMenu,crrType,i,bFlag
	Fn_Web_CreateNewForm=False
	'Creating Objects Of "Form" And "FormInfo" Tables
'	Set ObjForm=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewForm")
'	Set ObjFormInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewForm").WebTable("FormInfo")
	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	Set ObjForm=Fn_SISW_Web_GetObject("NewForm")
	Set ObjFormInfo=ObjForm.WebTable("FormInfo")
	
	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "NewForm")
	'Checking Existance Of "New Form" Dialog
	If Not ObjForm.Exist(SISW_MIN_TIMEOUT) Then
		'Calling New->Form... Menu Option
		Call Fn_Web_MenuOperation("Select",strMenu)
       	Call Fn_Web_ReadyStatusSync(1)
	End If

	bFlag=False
	
	For i=0 to 2
'		If ObjForm.Exist(5) Then
		If Fn_Web_UI_ObjectExist("Fn_Web_CreateNewForm",ObjForm)=True Then
			bFlag=True
			Exit for
		Else
			wait WEB_MICROLESS_TIMEOUT
		End If
	Next
	If bFlag=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "New Form dialog is not exist")
		Exit function
	End If

	If strFormType<>"" Then
'		crrType=ObjForm.WebEdit("FormType").GetROProperty("value")
		crrType=Fn_WEB_UI_Object_GetROProperty("Fn_Web_CreateNewForm",ObjForm.WebEdit("FormType"),"value")
		wait(WEB_MICRO_TIMEOUT)
		If Trim(crrType)<>Trim(strFormType) Then
			'Setting Form Type
'			Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewForm",Browser("TeamcenterWeb").Page("MyTeamCenter"),"FormTypeButton")
			Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewForm",ObjMyTcPage,"FormTypeButton")
			wait(WEB_MICRO_TIMEOUT)
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateNewForm",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",strFormType)
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateNewForm",ObjMyTcPage.WebElement("FormType"),"innertext",strFormType)
			'If Not Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Exist(3) Then
			Wait WEB_MICRO_TIMEOUT
			If Not Fn_Web_UI_ObjectExist("Fn_Web_CreateNewForm",ObjMyTcPage.WebElement("FormType")) Then
'				Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewForm",Browser("TeamcenterWeb").Page("MyTeamCenter"),"FormTypeButton")
				Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewForm",ObjMyTcPage,"FormTypeButton")
				wait(WEB_MICRO_TIMEOUT)
'				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateNewForm",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",strFormType)
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateNewForm",ObjMyTcPage.WebElement("FormType"),"innertext",strFormType)
			End If
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
			Call Fn_Web_UI_WebElement_Click("Fn_Web_CreateNewForm", ObjMyTcPage, "FormType", 1,1,micLeftBtn)
			wait(WEB_MICROLESS_TIMEOUT)
		End If
	End If
   	'Clicking On Next Button
'	Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewForm", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewForm", ObjMyTcPage.WebElement("ButtunPanel"), "Next")
	If strFormName<>"" Then
		'Setting "Form Name"
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateNewForm", ObjFormInfo, "FormName", strFormName)
	End If
	If strFormDesc<>"" Then
		'Setting "Form Dewscription"
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateNewForm", ObjFormInfo, "FormDesc", strFormDesc)
	End If
	If strURL<>"" Then
		'Setting "Website URL"
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateNewForm", ObjFormInfo, "WebsiteURL", strURL)
	End If
	'Clicking Finish Button To Create New Form
'	Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewForm", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewForm", ObjMyTcPage.WebElement("ButtunPanel"), "Finish")
'   	For i=0 to 2
'		If ObjFormInfo.Exist(5) and ObjFormInfo.GetROProperty("height")>0 Then
'			wait 2
'			Fn_Web_CreateNewForm=False
'		Else
'			Fn_Web_CreateNewForm=True
'			Exit for
'		End If
'	Next
	For i=0 to 2
		If Fn_Web_UI_ObjectExist("Fn_Web_CreateNewForm",ObjFormInfo) and Fn_WEB_UI_Object_GetROProperty("Fn_Web_CreateNewForm",ObjFormInfo,"height") > 0 Then
			wait WEB_MICROLESS_TIMEOUT
			Fn_Web_CreateNewForm=False
		Else
			Fn_Web_CreateNewForm=True
			Exit for
		End If
	Next
	'Releasing Object Of "Form" And "FormInfo" Table
	Set ObjMyTcPage = Nothing
	Set ObjFormInfo=Nothing
	Set ObjForm=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_QuickSearch
'@@
'@@    Description				 :	Function Used to Perform Quick Search For Item,Dataset ect.
'@@
'@@    Parameters			   :	1.strSearchType : Search Type ( eg. Item ID)
'@@												 2.strCriteria : Criteria ( Item Name "ABC")
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log in Web Client							
'@@
'@@    Examples					:	Call Fn_Web_QuickSearch("Item Name","Test")
'@@												Call Fn_Web_QuickSearch("Dataset Name","DemoDataset")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												    Sandeep Navghane									13-Apr-2011						1.0																				Sunny Ruparel
'@@                                                 Pritam Shikare									    4-Oct-2013						1.1						modified the code to select the search type				Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_QuickSearch(strSearchType,strCriteria)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_QuickSearch"
	'Variable Declaration
	Dim ObjMyTC,WshShell, iCount, sItem, iItemsCount
	Fn_Web_QuickSearch=False
	'Creating Object Of "MyTeamCenter" Page
	Set ObjMyTC=Browser("TeamcenterWeb").Page("MyTeamCenter")
	If strSearchType<>"" Then
		'Selecting Search Type eg. Item ID,Dataset Name
		wait(5)
		Call Fn_Web_UI_List_Select("Fn_Web_QuickSearch", ObjMyTC, "QuickSearchList",strSearchType)
		ObjMyTC.WebList("QuickSearchList").Object.setActive
		'Added By Pritam
		set WshShell = CreateObject("WScript.shell")
		ObjMyTC.WebList("QuickSearchList").Click 5,5,micLeftBtn
		wait 1
		sItem = ObjMyTC.WebList("QuickSearchList").GetROProperty("selection")
		If sItem <> strCriteria Then
			WshShell.SendKeys("{HOME}")
			wait 1
		End If
		iItemsCount = ObjMyTC.WebList("QuickSearchList").GetROProperty("items count")
		For iCount = 1 to iItemsCount
			sItem = Browser("TeamcenterWeb").Page("MyTeamCenter").WebList("QuickSearchList").GetROProperty("selection")
			If sItem <> strSearchType Then
				WshShell.SendKeys("{DOWN}")
				wait 1
			End If
		Next
		WshShell.SendKeys("{ENTER}")
		Wait 1
	End If
	If strCriteria<>"" Then
		'Setting Search Criteria eg. Dataset Name ( TestDataset )
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_QuickSearch", ObjMyTC, "QuickSearchCriteria", strCriteria)
		Wait 1
		Browser("TeamcenterWeb").Page("MyTeamCenter").Image("QuickSearch").Click
		wait 5
		If Fn_Web_UI_ObjectExist("Fn_Web_QuickSearch", ObjMyTC.WebTable("QuickSearchTable"))=True Then
			'Function Returns True
			Fn_Web_QuickSearch=True
		End If
	End If
	'Releasing Object Of "MyTeamCenter" Page
	Set ObjMyTC=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SidePannelLinkOperations
'@@
'@@    Description				 :	Function Used to Perform Operations On Side Pannel Links ( eg. Quick Links )
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@												 2.strLinkName : Link Name ( eg. Home , Create an Item...)
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log in Web Client							
'@@
'@@    Examples					:	Call Fn_Web_SidePannelLinkOperations("Select","Home")
'@@												Call Fn_Web_SidePannelLinkOperations("Select","Create an Item...")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									13-Apr-2011						1.0																								Sunny Ruparel
'@@            									Sandeep Navghane         							10-Oct-2011      				 1.1     								Replace [ Link ] With [ WebButton ]       Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_SidePannelLinkOperations(strAction,strLinkName)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SidePannelLinkOperations"
	'Variable Declaration
	Dim ObjMyTc
	'Creating Object Of "MyTeamCenter" Web Page
'	Set ObjMyTc=Browser("TeamcenterWeb").Page("MyTeamCenter")
	Set ObjMyTc=Fn_SISW_Web_GetObject("MyTeamCenter")
	Fn_Web_SidePannelLinkOperations=False
	 Select Case strAction
		Case "Select" 'Case To Select Or Click Link 
			If strLinkName<>"" Then
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SidePannelLinkOperations",ObjMyTc.WebButton("SidePanelButtons"),"name",strLinkName)
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SidePannelLinkOperations",ObjMyTc.Link("SidePannelLinks"),"text",strLinkName)
				If ObjMyTc.WebButton("SidePanelButtons").Exist(5) Then
					  Call Fn_Web_UI_Button_Click("Fn_Web_SidePannelLinkOperations",ObjMyTc,"SidePanelButtons")
					  'Function Returns True 
					  Fn_Web_SidePannelLinkOperations=True
				ElseIf ObjMyTc.Link("SidePannelLinks").Exist(1) Then
					  Call Fn_Web_UI_Link_Click("Fn_Web_SidePannelLinkOperations", ObjMyTc,"SidePannelLinks","","","")
					  'Function Returns True 
					   Fn_Web_SidePannelLinkOperations=True 
				End If
				Wait WEB_MICRO_TIMEOUT
			End IF
	End Select
 'Releasing Object Of "MyTeamCenter" Web Page
   Set ObjMyTc=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ErrorMsgVerify
'@@
'@@    Description				 :	Function Used to Handle Error Dialogs And To Verify Error Mesaages
'@@
'@@    Parameters			   :	1.strErrorMsg : Error Message
'@@												  2.strButton : Button Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Error Message Should be Appear on Screen						
'@@
'@@    Examples					:	Call Fn_Web_ErrorMsgVerify("The object 001168-Test","OK")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									13-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ErrorMsgVerify(strErrorMsg,strButton)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ErrorMsgVerify"
	GBL_EXPECTED_MESSAGE=strErrorMsg
 	Dim bFlag,strBrowser,CurrErr,WshShell
	Fn_Web_ErrorMsgVerify=False
	bFlag=False
	strBrowser=Environment.Value("WebBrowserName")

	If Fn_Web_UI_ObjectExist("Fn_Web_ErrorMsgVerify", Browser("TeamcenterWeb").Dialog("Dialog"))=True Then
		If InStr(1,strBrowser,"IE")>0 Then
			If strErrorMsg<>"" Then
				CurrErr=Browser("TeamcenterWeb").Dialog("Dialog").Static("ErrorMsg").GetROProperty("text")
				wait(1)
				If InStr(1,CurrErr,strErrorMsg)>0 Then
					bFlag=True
				Else 
					GBL_ACTUAL_MESSAGE=CurrErr
				End If
			End If
			Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").SetTOProperty "text",strButton
			wait(1)
			Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").Click
			
		
			'Temp Code to Handle Stop Script Error From line 851 to 855
			If Browser("TeamcenterWeb").Dialog("Dialog").Exist(6)  Then
				Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").SetTOProperty "text","&Yes"
			wait(1)
			Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").Click
			End If
		Else
			If strErrorMsg<>"" Then
				CurrErr=Browser("TeamcenterWeb").Dialog("Dialog").Page("FFPage").WebElement("ErrorMsg").GetROProperty("innertext")
				If InStr(1,CurrErr,strErrorMsg)>0 Then
					bFlag=True
				Else
					GBL_ACTUAL_MESSAGE=CurrErr
				End If
			End If
			Browser("TeamcenterWeb").Dialog("Dialog").Page("FFPage").WebButton("OK").SetTOProperty "name",strButton
			wait(1)
			Browser("TeamcenterWeb").Dialog("Dialog").Page("FFPage").WebButton("OK").Click
		End If
	End If
	If bFlag=True Then
		Fn_Web_ErrorMsgVerify=True
	End If
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CreateDataset
'@@
'@@    Description				 :	Function Used to Create Dataset
'@@
'@@    Parameters			   :	1.strDatasetType : Dataset Type
'@@												 2.strDatasetName : Dataset Name
'@@												 3.strDatasetDesc : Dataset Description
'@@												 4.strFilePath : External File Path (Full Path including File name and Extension)
'@@												 5.strReference : Reference
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log in Web Client							
'@@
'@@    Examples					:	Call Fn_Web_CreateDataset("Briefcase","Dataset1","Demo Dataset","","")
'@@												Call Fn_Web_CreateDataset("AutoText","Dataset2","Demo Dataset","","")
'@@												Call Fn_Web_CreateDataset("MSWord","WordDataset","Dataset","C:\mainline\TestData\UserAccessScenario007\Dataset_001.doc","word")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									14-Apr-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									03-Nov-2011						1.1																							Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_CreateDataset(strDatasetType,strDatasetName,strDatasetDesc,strFilePath,strReference)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CreateDataset"
   'Variable Declaration
   Dim ObjForm,ObjFormInfo
   Dim objNewDataset, ObjMyTcPage
   Dim strMenu,crrType,currRef
	Fn_Web_CreateDataset=False
	'Creating Objects Of "Form" And "FormInfo" Tables
'	Set ObjForm=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewForm")
'	Set ObjFormInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewForm").WebTable("FormInfo")
	
	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	Set ObjForm = Fn_SISW_Web_GetObject("NewForm")
	Set ObjFormInfo = Fn_SISW_Web_GetObject("NewForm").WebTable("FormInfo")
	Set objNewDataset = Fn_SISW_Web_GetObject("NewDataset")
	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "NewDataset")
	
	'Checking Existance Of "New Form" Dialog
	If Not ObjForm.Exist(SISW_MIN_TIMEOUT) Then
		'Calling New->Form... Menu Option
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(2)
	End If

	If strDatasetType<>"" Then
'		crrType=ObjForm.WebEdit("DatasetType").GetROProperty("value")
		crrType=Fn_WEB_UI_Object_GetROProperty("Fn_Web_CreateDataset",ObjForm.WebEdit("DatasetType"),"value")
		wait(WEB_MICRO_TIMEOUT)
		If Trim(crrType)<>Trim(strDatasetType) Then
			'Setting Form Type
'			Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset",Browser("TeamcenterWeb").Page("MyTeamCenter"),"FormTypeButton")
'			Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewDataset"),"DatasetTypeButton")
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateDataset",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",strDatasetType)
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
'			wait(2)
			Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset", objNewDataset,"DatasetTypeButton")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateDataset",ObjMyTcPage.WebElement("FormType"),"innertext",strDatasetType)
			Call Fn_Web_UI_WebElement_Click("Fn_Web_CreateDataset", ObjMyTcPage, "FormType", 1,1,micLeftBtn)
			wait(WEB_MICROLESS_TIMEOUT)
		End IF
	End If
	'Clicking On Next Button
'	Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset", ObjMyTcPage.WebElement("ButtunPanel"), "Next")
	If strDatasetName<>"" Then
		'Setting "Form Name"
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateDataset", ObjFormInfo, "FormName", strDatasetName)
	End If
	If strDatasetDesc<>"" Then
		'Setting "Form Dewscription"
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateDataset", ObjFormInfo, "FormDesc", strDatasetDesc)
	End If
	If strFilePath<>"" Or strReference<>"" Then
		'Clicking On Next Button
'		Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
		Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset", ObjMyTcPage.WebElement("ButtunPanel"), "Next")
		wait(WEB_MICROLESS_TIMEOUT)
		If strFilePath<>"" Then
'			Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewDataset").WebTable("References"), "Browse")
			Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset",objNewDataset.WebTable("References"), "Browse")
			wait(WEB_DEFAULT_TIMEOUT)
			If  JavaDialog("UploadFile").Exist(10) Then
				JavaDialog("UploadFile").JavaEdit("FileName").Set strFilePath
				wait(WEB_MICRO_TIMEOUT)
				JavaDialog("UploadFile").JavaButton("Open").Click micLeftBtn
				wait(WEB_MICRO_TIMEOUT)
			Else
				Fn_Web_CreateDataset = false
				Exit function
			End If
		End If
		If strReference<>"" Then
'			currRef=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewDataset").WebTable("References").WebEdit("ReferenceType").GetROProperty("value")
			currRef=Fn_WEB_UI_Object_GetROProperty("Fn_Web_CreateDataset",objNewDataset.WebTable("References").WebEdit("ReferenceType"),"value")
			If Trim(currRef)<>Trim(strReference) Then
'				Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset",Browser("TeamcenterWeb").Page("MyTeamCenter"),"DatasetReference")
'				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateDataset",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",strReference)
'				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
				Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset",ObjMyTcPage,"DatasetReference")
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateDataset",ObjMyTcPage.WebElement("FormType"),"innertext",strReference)
				Call Fn_Web_UI_WebElement_Click("Fn_Web_CreateDataset", ObjMyTcPage, "FormType", 1,1,micLeftBtn)
				wait(WEB_MICROLESS_TIMEOUT)
			End If
		End If
	End If
	'Clicking Finish Button To Create New Form
'	Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	Call Fn_Web_UI_Button_Click("Fn_Web_CreateDataset", ObjMyTcPage.WebElement("ButtunPanel"), "Finish")
'	If Browser("TeamcenterWeb").Dialog("Dialog").Exist(5) Then
	If Fn_Web_UI_ObjectExist("Fn_Web_CreateDataset",Browser("TeamcenterWeb").Dialog("Dialog"))=True Then
		Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").Click
		wait WEB_MICROLESS_TIMEOUT
	End If
	'Function Returns True
	Fn_Web_CreateDataset=True
	'Releasing Object Of "Form" And "FormInfo" Table
	Set ObjFormInfo=Nothing
	Set ObjForm=Nothing
	Set ObjMyTcPage = Nothing
	Set objNewDataset = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CreateWebLink
'@@
'@@    Description				 :	Function Used to Create Web URL
'@@
'@@    Parameters			   :	1.strURLName : URL ( Link ) Name
'@@												 2.strURLDesc : URL Description
'@@												 3.strURL : URL ( Link ) Eg. = "www.google.com"
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log in Web Client							
'@@
'@@    Examples					:	Call Fn_Web_CreateWebLink("Google","Google Home Page URL","www.google.com")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									14-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_CreateWebLink(strURLName,strURLDesc,strURL)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CreateWebLink"
   'Variable Declaration
   Dim ObjURLInfo,strWEBMenuPath,strMenu
   Fn_Web_CreateWebLink=False
   'Creating Object Of "WebLinkInfo" Table
	Set ObjURLInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebLinkInfo")
	If Not ObjURLInfo.Exist(5) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewWebLink")
		'Calling New->Form... Menu Option
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	If strURLName<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateWebLink", ObjURLInfo, "URLName", strURLName)
	End If
	If strURLDesc<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateWebLink", ObjURLInfo, "URLDesc", strURLDesc)
	End If
	If strURL<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateWebLink", ObjURLInfo, "URL", strURL)
	End If
	Call Fn_Web_UI_Button_Click("Fn_Web_CreateWebLink",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	Fn_Web_CreateWebLink=True
	Set ObjURLInfo=Nothing
End Function 
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_VerifyProperties
'@@
'@@    Description				 :	Function Used to Verify Object Properties
'@@
'@@    Parameters			   :	1.strPropertyValuePair: Property Name And Property Value Separeted by [ : ]
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Object Should be Selected 
'@@
'@@    Examples					:	Call Fn_Web_VerifyProperties("Current ID:000011")
'@@												Call Fn_Web_VerifyProperties("Current ID:000011,Current Name:Test,Date Created:15-Apr-2011")
'@@												Call Fn_Web_VerifyProperties("Current Name:Cmp,Form Definition File:n/a,Last Modified Date:Y")
'@@
'@@	   History					 	:	
'@@													Developer Name									Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane								18-Apr-2011						1.0																								Sunny Ruparel
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Koustubh Watwe									29-Nov-2011						1.0					Added codeto handle WebEditbox
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'Function Fn_Web_VerifyProperties(strPropertyValuePair)
'   'Variable Declaration
'   Dim ObjTable,ObjInfoTbd,objPropTable
'   Dim strWEBMenuPath,strMenu,crrPropValues,arrPair,iCounter,arrProp,strPropVal
'   Dim iRwCount,iCount,sPropName,sPropVal,bFlag,arrPropName,iPropertyCounts
'   Dim objWebEle
'   Fn_Web_VerifyProperties=False
'   'Creating Object Ob "Object" Table
'   Set ObjTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object")
'   Set ObjInfoTbd=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectInfo")
'
'   bFlag=False
'   strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
'	strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "ViewProperties")
'   'Checking Existance Of Properties Dialog
'   If Not ObjTable.Exist(5) Then
'		Call Fn_Web_MenuOperation("Select",strMenu)
'   End If
'   'Clicking On "All" Tab
'   Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext","All")
'   If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Exist(4) Then
'	   Call Fn_Web_UI_WebElement_Click("Fn_Web_VerifyProperties", Browser("TeamcenterWeb").Page("MyTeamCenter"), "FormType", "","","")
'		Set ObjTable=nothing
''		'Added Code to handle properties of All tab
'		Set ObjTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AllProperties")
'	   bFlag=True
'   End If
'    iPropertyCounts=Split(strPropertyValuePair,":")
'   If UBound(iPropertyCounts)=1 Then
'			arrPair=Split(strPropertyValuePair,"~")
'	Else
'		If InStr(1,strPropertyValuePair,"~")>0 Then
'			arrPair=Split(strPropertyValuePair,"~")
'		Else
'			arrPair=Split(strPropertyValuePair,",")
'		End If
'	End If
'   If bFlag=True Then
'	   'Taking All Data From Table
'	   'crrPropValues=ObjTable.GetCellData(1,1)
'	   crrPropValues=ObjTable.GetROProperty("innertext")
'		For iCounter=0 To UBound(arrPair)
'			arrProp=Split(arrPair(iCounter),":")
'			If arrProp(1)<>"" Then
'				strPropVal=arrProp(0)+": "+arrProp(1)
'			Else
'				strPropVal=arrProp(0)
'			End If
'			'Checking Property Value Pair Exist Or Not
'			If InStr(1,crrPropValues,strPropVal)>0 Then
'				Fn_Web_VerifyProperties=True
'			Else
'				Fn_Web_VerifyProperties=False
'				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Properti-Value ["+crrPropValues+":"+strPropVal+"] Pair Not Found Properties Table")
'				Exit For
'			End If
'		Next
'	Else
'		If ObjInfoTbd.Exist(7) Then
'			Set objPropTable=ObjInfoTbd
'		Else
'			Set objPropTable=ObjTable
'		End If
'		iRwCount=objPropTable.RowCount()
'		For iCount=0 To UBound(arrPair)
'			arrProp=Split(arrPair(iCount),":")
'			For iCounter=0 To iRwCount
'				sPropName=objPropTable.GetCellData(iCounter,1)
'				arrPropName=Split(sPropName,":")
'				If arrPropName(0)=arrProp(0) Then
'					sPropVal=objPropTable.GetCellData(iCounter,2)
'					Set objWebEle = objPropTable.ChildItem(iCounter, 2, "WebEdit", 0)
'					If TypeName(objWebEle) <> "Nothing" Then
'							If trim(objWebEle.GetROProperty("value")) = Trim(arrProp(1)) then 
'								bFlag = True
'							Else
'								bFlag=False
'								Exit For
'							end if
'					ElseIf sPropVal=arrProp(1) Then
'						bFlag=True
'						Exit For
'					Else
'						bFlag=False
'						Exit For
'					End If
'				End If
'			Next
'			If bFlag=False Then
'				Exit For
'			End If
'		Next
'	If bFlag=True Then
'		Fn_Web_VerifyProperties=True
'	End If
'	End If
'		'Clicking "Cancel" Button
'		If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Cancel").Exist(5) Then
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Cancel").Click
'		End If
'		If 	Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Close").Exist(5) Then
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Close").Click
'		End If
'End Function

'Modified by Sandeep Navghane : 29-Aug-2012
' - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -
'bReturn=Fn_Web_VerifyProperties("Description:Item Created,Displayable Revisions:02776/A;1-Item")
'bReturn=Fn_Web_VerifyProperties("Name:Form34328,Group ID:Engineering$Test$2")
' - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -- - - - - - - - - - - - -
Function Fn_Web_VerifyProperties(strPropertyValuePair)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_VerifyProperties"
   'Variable Declaration
   Dim ObjTable,objPropTable, arrPropValue, Iterator
   Dim strWEBMenuPath,strMenu,arrPair,iCounter,arrProp,strPropVal
   Dim iRwCount,iCount,sPropName,sPropVal,bFlag,arrPropName,iPropertyCounts
   Dim objWebEle,aPropTabNameAndIndex,i,ObjMyTcPage

	Fn_Web_VerifyProperties=False
	'Creating Object Ob "Object" Table
'	Set ObjTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object")
'	Set objPropTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AllProperties")
	
	Set ObjTable = Fn_SISW_Web_GetObject("Object")
	Set objPropTable = Fn_SISW_Web_GetObject("AllProperties")
	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")

	aPropTabNameAndIndex=split(strPropertyValuePair,"$")
	If ubound(aPropTabNameAndIndex)=0 Then
		'Clicking On "All" Tab
'		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext","All")
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebElement("FormType"),"innertext","All")
	Else
		'Clicking On user required Tab
'		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",aPropTabNameAndIndex(1))
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebElement("FormType"),"innertext",aPropTabNameAndIndex(1))
		If aPropTabNameAndIndex(1)="General" Then
			aPropTabNameAndIndex(1)=""
		End If
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").SetTOProperty "text",".*General.*"+aPropTabNameAndIndex(1)+".*"
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").SetTOProperty "outertext",".*General.*"+aPropTabNameAndIndex(1)+".*"
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").SetTOProperty "innertext",".*General.*"+aPropTabNameAndIndex(1)+".*"
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebTable("Object"),"text",".*General.*"+aPropTabNameAndIndex(1)+".*")
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebTable("Object"),"outertext",".*General.*"+aPropTabNameAndIndex(1)+".*")
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebTable("Object"),"innertext",".*General.*"+aPropTabNameAndIndex(1)+".*")
	End If
'   	bFlag=False

   	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "ViewProperties")
'	Checking Existance Of Properties Dialog
	   If Not ObjTable.Exist(SISW_MIN_TIMEOUT) and not objPropTable.Exist(SISW_MICRO_TIMEOUT)Then
		'strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "ViewProperties")
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(2)
	   End If

	bFlag=False
	For i=0 to 2
		'If ObjTable.Exist(5) or objPropTable.Exist(5) Then
		If Fn_Web_UI_ObjectExist("Fn_Web_VerifyProperties",ObjTable)=True Or Fn_Web_UI_ObjectExist("Fn_Web_VerifyProperties",objPropTable)=True Then
			bFlag=True
			Exit for
'		Else
'			wait 2
		End If
	Next
	If bFlag=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Properties dialog is not exist")
		Exit function
	End If
	bFlag=False
   
'	If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Exist(4) Then
	If Fn_Web_UI_ObjectExist("Fn_Web_VerifyProperties",ObjMyTcPage.WebElement("FormType"))=True Then
'		Call Fn_Web_UI_WebElement_Click("Fn_Web_VerifyProperties", Browser("TeamcenterWeb").Page("MyTeamCenter"), "FormType", "","","")
		Call Fn_Web_UI_WebElement_Click("Fn_Web_VerifyProperties", ObjMyTcPage, "FormType", "","","")
		
		'Added Code to handle properties of All tab
		If ubound(aPropTabNameAndIndex)=2 Then
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectProperties").SetTOProperty "index",Cint(aPropTabNameAndIndex(2))-1
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjTable.WebTable("ObjectProperties"),"index",Cint(aPropTabNameAndIndex(2))-1)
			wait WEB_MICRO_TIMEOUT
		End If
'		Set ObjTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectProperties")
		Set ObjTable=nothing
		Set ObjTable=Fn_SISW_Web_GetObject("Object").WebTable("ObjectProperties")
	Else
		Set ObjTable=objPropTable
	End If

	iPropertyCounts=Split(aPropTabNameAndIndex(0),":")
	If UBound(iPropertyCounts)=1 Then
			arrPair=Split(aPropTabNameAndIndex(0),"~")
	Else
		If InStr(1,aPropTabNameAndIndex(0),"~")>0 Then
			arrPair=Split(aPropTabNameAndIndex(0),"~")
		Else
			arrPair=Split(aPropTabNameAndIndex(0),",")
		End If
	End If
	iRwCount=ObjTable.RowCount()
	For iCount=0 To UBound(arrPair)
		arrProp=Split(arrPair(iCount),":")
		For iCounter=0 To iRwCount
			sPropName=ObjTable.GetCellData(iCounter,1)
			arrPropName=Split(sPropName,":")
			If arrPropName(0)=arrProp(0) Then
				If instr(arrPair(iCount), "?") = 0 Then
				    	sPropVal=ObjTable.GetCellData(iCounter,2)
				    	Set objWebEle = ObjTable.ChildItem(iCounter, 2, "WebEdit", 0)
					    If TypeName(objWebEle) <> "Nothing" Then
						If trim(objWebEle.GetROProperty("value")) = Trim(arrProp(1)) then 
							bFlag = True
						Else
							bFlag=False
							Exit For
						End If
					    ElseIf sPropVal=arrProp(1) Then
						  bFlag=True
						  Exit For
					   Else
						  bFlag=False
						  Exit For
					   End If
				Else		'' added code to verify multiple property values whose sequence they appears are changing 
					arrPropValue = Split(arrProp(1), "?")
					For Iterator = 0 To uBound(arrPropValue)
						sPropVal=ObjTable.GetCellData(iCounter,2)
			    	   		 Set objWebEle = ObjTable.ChildItem(iCounter, 2, "WebEdit", 0)
				       		 If TypeName(objWebEle) <> "Nothing" Then
							If instr(trim(objWebEle.GetROProperty("value")), Trim(arrPropValue(Iterator)))>0 then 
								bFlag = True
							Else
							  	bFlag=False
							  	Exit For
							End if
						ElseIf instr(sPropVal, arrPropValue(Iterator))> 0 Then
					      		bFlag=True
						Else
							  bFlag=False
							  Exit For
			          		End If
				     	Next
				End If
			End If
		Next
		If bFlag=False Then
			Exit For
		End If
	Next
	If bFlag=True Then
		Fn_Web_VerifyProperties=True
	End If
	'Clicking "Cancel" Button
'	If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Cancel").Exist(5) Then
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Cancel").Click
'	End If
	If Fn_Web_UI_ObjectExist("Fn_Web_VerifyProperties",ObjMyTcPage.WebElement("ButtunPanel").WebButton("Cancel"))=True Then
		Call Fn_Web_UI_Button_Click("Fn_Web_VerifyProperties",ObjMyTcPage.WebElement("ButtunPanel"),"Cancel")
	End If
	
'	If ObjTable.Exist(5) Then
'		If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Close").Exist(5) and ObjTable.getROProperty("height")>0Then
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Close").Click
'		End If
'	End If
	If Fn_Web_UI_ObjectExist("Fn_Web_VerifyProperties",ObjTable)=True Then
		If Fn_Web_UI_ObjectExist("Fn_Web_VerifyProperties",ObjMyTcPage.WebElement("ButtunPanel").WebButton("Close"))=True And Fn_WEB_UI_Object_GetROProperty("Fn_Web_VerifyProperties",ObjTable,"height") > 0 Then
			Call Fn_Web_UI_Button_Click("Fn_Web_VerifyProperties",ObjMyTcPage.WebElement("ButtunPanel"),"Close")
		End If
	End If
	
	
	
	If ubound(aPropTabNameAndIndex)=2 or ubound(aPropTabNameAndIndex)=1 Then
		'- - - - - - - - - - - - - - - - - Reverting All changes
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").SetTOProperty "text",".*General.*All.*"
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").SetTOProperty "outertext",".*General.*All.*"
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").SetTOProperty "innertext",".*General.*All.*"
'		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectProperties").SetTOProperty "index",3
		
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebTable("Object"),"text",".*General.*All.*")
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebTable("Object"),"outertext",".*General.*All.*")
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebTable("Object"),"innertext",".*General.*All.*")
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_VerifyProperties",ObjMyTcPage.WebTable("Object").WebTable("ObjectProperties"),"index",3)
		'- - - - - - - - - - - - - - - - - - - - 
	End If
	Set ObjTable=Nothing
	Set objPropTable=Nothing
	Set ObjMyTcPage = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@     FUNCTION NAME   :   Fn_Web_UserSettingsOperations()
'@@
'@@    DESCRIPTION     :   1. This function Modifies values in user Settings Dialog Box
'@@											2.This Function Verifies Item values Changes By user
'@@ 
'@@   PARAMETERS      :   sAction  			   - String value for Action to be performed   
'@@											 ''strLabels = 		- labels Seperated by "~"	
''@@   										strValues = 		- Values Seperated by  "~"
'@@   											userLink	 			-  String value   
'@@
'@@   Return Value  :   True/False  
'@@
'@@   EXAMPLE : 1. To modify  User Settings
''@@ 							 strLabels = "Group / Role~Revision Rule~Engineering~Default Group"
''@@ 							 strValues = "Engineering/Designer~Latest Working~Designer~Engineering"
'@@ 							 
'@@ 							 Fn_Web_UserSettingsOperations("Edit",strLabels,strValues,"")
'@@
'@@							2.To verify  Changed user Settings
'@@							Fn_Web_UserSettingsOperations("Verify",strLabels,strValues,"Engineering / Designer - Latest Working")
'@@							Fn_Web_UserSettingsOperations("Verify","","","( AutoTest1 ( autotest1 ) - Engineering / Designer - Latest Working")
'@@							
'@@							3. To verify Combobox Containts of any Field
'@@							strLabels = "dba~dba~Engineering~Engineering"
'@@							strValues = "Designer~CostDBA~Designer~Tc_QALead"
'@@                            Fn_Web_UserSettingsOperations("ListVerify",strLabels,strValues,"")
'@@							
'@@							4. To verify Labels on Dialog
'@@							strLabels = "dba~dba~Engineering~Engineering"
'@@                            Fn_Web_UserSettingsOperations("LabelVerify",strLabels,"","")
'@@  History					 :		
'@@												Developer Name												Date							Version						Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Pranav Ingle									   			18-Apr-2011			              1.0									Created														Sunny	
'@@												Pranav Ingle									   			18-Apr-2011			              1.1									Added Case "ListVerify"						Sunny	
'@@												Pranav Ingle									   			05-May-2011			              1.2									Added Case "LabelVerify"				Sunny	
'@@												Vallari												   			28-Jul-2011			              1.2									Added Code to click on Tab2					
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_UserSettingsOperations(sAction,strLabels,strValues,userLink)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_UserSettingsOperations"
	'variable Declaration       
	Dim objMyTc, ObjButton,objWebChild,arrLabels,arrValues,arrUserLink,iCount,iCounter,strCellData,iRowCount,strEditValue,bFlag,intNoOfObjects,objSelectType
	Dim strMenu,Row1
	Dim aAction,objUserSettings,objMyTcPage
	Fn_Web_UserSettingsOperations =false
	Set objMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")

	If instr(sAction, ":") > 0 Then
		aAction = split(sAction, ":", -1,1)
		sAction = aAction(1)
	End If

    Set objUserSettings = Fn_SISW_Web_GetObject("UserSettings")
'    Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings").WebTable("GroupRole").SetTOProperty "index", "0"
     Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_UserSettingsOperations",objUserSettings.WebTable("GroupRole"),"index","0")

'    Set ObjButton = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel")
    Set ObjButton = Fn_SISW_Web_GetObject("ButtunPanel")
        
	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "EditUserSettings")        
	' Check Existance Of User Settings Dialog
	If Fn_Web_UI_ObjectExist("Fn_Web_UserSettingsOperations", ObjButton)= False  OR Fn_Web_UI_ObjectVisible("Fn_Web_UserSettingsOperations", ObjButton) = False Then
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(1)
	End If

	If isarray(aAction) Then
		If lcase(aAction(0)) = "tab2" Then
			Set objTAB=description.Create()
			objTAB("micClass").value = "WebElement"
			objTAB("html tag").value = "A"
			objTAB("innertext").value = "Default Role Settings"                                        
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement(objTAB).Click 5,5,micLeftBtn
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings").WebTable("GroupRole").SetTOProperty "index", "1"
			objMyTcPage.WebElement(objTAB).Click 5,5,micLeftBtn
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_UserSettingsOperations",objUserSettings.WebTable("GroupRole"),"index","1")
			Set objTAB = nothing
			' Create Object of MyTc                
'			Set objMyTc = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings").WebTable("GroupRole")
			Set objMyTc = objUserSettings.WebTable("GroupRole")
		Else
'	        		Set objMyTc =Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings").WebTable("LoginSetting")
	        		Set objMyTc =objUserSettings.WebTable("LoginSetting")
		End If
	Else
'		Set objMyTc =Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings").WebTable("LoginSetting")
		Set objMyTc =objUserSettings.WebTable("LoginSetting")
	End If
   
     Select Case sAction
		Case "Edit"
			If strLabels<> "" Then
				arrLabels = Split(strLabels,"~")
				arrValues= Split(strValues,"~")
			Else
			   	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : strLabels  can not be Empty to Edit values")
				Exit Function
			End If

			For iCount = 0 to UBound(arrLabels)
'				iRowCount = objMyTc.GetROProperty("rows")
				iRowCount = Fn_WEB_UI_Object_GetROProperty("Fn_Web_UserSettingsOperations",objMyTc,"rows")
				bFlag=False
				For iCounter= 1 to iRowCount-1
					strCellData = objMyTc.GetCellData(iCounter,1)
					If Instr(1,strCellData, arrLabels(iCount)) > 0 Then
						Set objWebChild = objMyTc.ChildItem(iCounter,2,"WebButton",0)
						objWebChild.Click 1,1,micleftBtn
						Call Fn_Web_ReadyStatusSync(1)

						Set objSelectType=description.Create()
						objSelectType("micClass").value = "WebElement"
						objSelectType("innertext").value = arrValues(iCount)
						Set  intNoOfObjects = objMyTc.ChildObjects(objSelectType)
						If  intNoOfObjects.Count > 0 Then
							bFlag=true
							intNoOfObjects(0).Click 1,1
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set  "+arrValues(iCount)+" For Group "+arrLabels(iCount))
						End If
						Set objWebChild=Nothing
						Set  intNoOfObjects=Nothing
						Set objSelectType=Nothing
					End If
				Next
				If bFlag=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed Label"+arrLabels(iCount)+ "Not Found")
					Exit Function
				End If
			Next
			'Set the Value in the Location Text Textbox
'			If  Fn_Web_UI_ObjectExist("Fn_Web_UserSettingsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings").WebEdit("LocationCode"))= true then
			If Fn_Web_UI_ObjectExist("Fn_Web_UserSettingsOperations", objUserSettings.WebEdit("LocationCode"))= true then
				If userLink<>"" Then
					'Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings").WebEdit("LocationCode").set userLink
'					Call Fn_Web_UI_WebEdit_SetExt("Fn_Web_UserSettingsOperations", "SendString",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings"), "LocationCode", userLink)	
					Call Fn_Web_UI_WebEdit_SetExt("Fn_Web_UserSettingsOperations", "SendString",objUserSettings, "LocationCode", userLink)	
					Wait WEB_MICROLESS_TIMEOUT
					'Verify if the Value is set in the Edit Box
'					strCellData=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("UserSettings").WebEdit("LocationCode").GetROProperty ("value")
					strCellData = Fn_WEB_UI_Object_GetROProperty("Fn_Web_UserSettingsOperations",objUserSettings.WebEdit("LocationCode"),"value")
					If lCase(strCellData)=lCase(userLink) Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set  "+userLink+" For Location Code")
						Fn_Web_UserSettingsOperations =True
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed To Set  "+userLink+" For Location Code")
						Fn_Web_UserSettingsOperations =false
						Exit Function
					End If
				End If
			End If
			'Click on  OK Button 
			Call Fn_Web_UI_Button_Click("Fn_Web_UserSettingsOperations",ObjButton,"OK")   
			Wait WEB_MINLESS_TIMEOUT			
		Case "Verify"
			If strLabels<> "" Then
				arrLabels = Split(strLabels,"~")
				arrValues= Split(strValues,"~")
				For iCount = 0 to UBound(arrLabels)
					bFlag=False
'					iRowCount = objMyTc.GetROProperty("rows")
					iRowCount = Fn_WEB_UI_Object_GetROProperty("Fn_Web_UserSettingsOperations",objMyTc,"rows")
					For iCounter= 1 to iRowCount-1
						strCellData = objMyTc.GetCellData(iCounter,1)
						If Instr(1,strCellData, arrLabels(iCount)) > 0 Then
							Set objWebChild = objMyTc.ChildItem(iCounter,2,"WebEdit",0)
							strEditValue = objWebChild.getRoProperty("value")
							if strEditValue = arrValues(iCount) then 
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : "+arrLabels(iCount)+" Verified Successfully")
								bFlag=True
								Exit For
							Else        
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify "+arrLabels(iCount))
								Exit Function
							End If
							Set objWebChild=Nothing
						End If
					Next
					wait WEB_MICRO_TIMEOUT
					If bFlag=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed Label"+arrLabels(iCount)+ "Not Found")
						Exit Function
					End If
				Next
			End If
			If userLink <> "" Then
'				arrUserLink = Browser("TeamcenterWeb").Page("MyTeamCenter").Link("userLink").GetROProperty("innertext")     
				arrUserLink = Fn_WEB_UI_Object_GetROProperty("Fn_Web_UserSettingsOperations",objMyTcPage.Link("userLink"),"innertext")
				If  Instr(1,arrUserLink,userLink) >= 0 Then 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Changes  appeared at topright corner under logout Successfully")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Changes  Not appeared at topright corner under logout")    
					Exit Function                      
				End If
			End If
			'Click on  Cancel Button 
			Call Fn_Web_UI_Button_Click("Fn_Web_UserSettingsOperations",ObjButton,"Cancel")
		Case "ListVerify"
			If strLabels<> "" Then
				arrLabels = Split(strLabels,"~")
				arrValues= Split(strValues,"~")

				For iCount = 0 to UBound(arrLabels)
'					iRowCount = objMyTc.GetROProperty("rows")
					iRowCount = Fn_WEB_UI_Object_GetROProperty("Fn_Web_UserSettingsOperations",objMyTc,"rows")
					bFlag=False
					For iCounter= 1 to iRowCount-1
						strCellData = objMyTc.GetCellData(iCounter,1)
						If Instr(1,strCellData, arrLabels(iCount)) > 0 Then
							Set objWebChild = objMyTc.ChildItem(iCounter,2,"WebButton",0)
							objWebChild.Click 1,1,micleftBtn
							Call Fn_Web_ReadyStatusSync(1)

							Set objSelectType=description.Create()
							objSelectType("micClass").value = "WebElement"
							objSelectType("innertext").value = arrValues(iCount)
							Set  intNoOfObjects = objMyTc.ChildObjects(objSelectType)
							If  intNoOfObjects.Count >= 0 Then
								bFlag=true
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified  "+arrValues(iCount)+" For Group "+arrLabels(iCount))
							End If
							Set objWebChild=Nothing
							Set  intNoOfObjects=Nothing
							Set objSelectType=Nothing
						End If
					Next
					If bFlag=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed Label"+arrLabels(iCount)+ "Not Found")
						Exit Function
					End If
				Next
			End If
			'Click on  Cancel Button 
			Call Fn_Web_UI_Button_Click("Fn_Web_UserSettingsOperations",ObjButton,"Cancel")
		Case "LabelVerify"
			If strLabels<> "" Then
				arrLabels = Split(strLabels,"~")
				For iCount = 0 to UBound(arrLabels)
'					iRowCount = objMyTc.GetROProperty("rows")
					iRowCount = Fn_WEB_UI_Object_GetROProperty("Fn_Web_UserSettingsOperations",objMyTc,"rows")
					bFlag=False
					For iCounter= 1 to iRowCount
						strCellData = objMyTc.GetCellData(iCounter,1)
						If Instr(1,strCellData, arrLabels(iCount)) > 0 Then
							bFlag=true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified  Label "+arrLabels(iCount))
						End If
					Next
					If bFlag=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed Label"+arrLabels(iCount)+ "Not Found")
						Exit Function
					End If
				Next
			End If
			'Click on  Cancel Button 
			Call Fn_Web_UI_Button_Click("Fn_Web_UserSettingsOperations",ObjButton,"Cancel")
	        Case "ListOrderVerify" 'to verify the Order pass the values without commas and space in Alpabetical manner
			If strLabels<> "" Then
				arrLabels = Split(strLabels,"~")
				arrValues= Split(strValues,"~")

				For iCount = 0 to UBound(arrLabels)
'					iRowCount = objMyTc.GetROProperty("rows")
					iRowCount = Fn_WEB_UI_Object_GetROProperty("Fn_Web_UserSettingsOperations",objMyTc,"rows")
					bFlag=False
					For iCounter= 1 to iRowCount-1
						strCellData = objMyTc.GetCellData(iCounter,1)
						If Instr(1,strCellData, arrLabels(iCount)) > 0 Then
							Set objWebChild = objMyTc.ChildItem(iCounter,2,"WebButton",0)
							objWebChild.Click 1,1,micleftBtn
							Call Fn_Web_ReadyStatusSync(1)

							Set objSelectType=description.Create()
							objSelectType("micClass").value = "WebElement"
							objSelectType("innertext").value = arrValues(iCount)
							Set  intNoOfObjects = objMyTc.ChildObjects(objSelectType)
							If  intNoOfObjects.Count > 0 Then
								bFlag=true
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified  "+arrValues(iCount)+" For Group "+arrLabels(iCount))
							End If
							Set objWebChild=Nothing
							Set  intNoOfObjects=Nothing
							Set objSelectType=Nothing
						End If
					Next
					If bFlag=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed Label"+arrLabels(iCount)+ "Not Found")
						Exit Function
					End If
				Next
			End If
			'Click on  Cancel Button 
			Call Fn_Web_UI_Button_Click("Fn_Web_UserSettingsOperations",ObjButton,"Cancel")
      End Select
                
	'Return Function Successful Log
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Function Successfully Completed")
	' Return True Value
	Fn_Web_UserSettingsOperations =true
	
	' Setting created objects to nothing
	Set objMyTcPage = Nothing
	Set objMyTc = Nothing
	Set ObjButton = Nothing
End Function


'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_EditProperties
'@@
'@@    Description				 :	Function Used to Edit Objects Name And Description Property
'@@
'@@    Parameters			   :	1.dicProperties: Properties Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Object Should be Selected
'@@
'@@    Examples					:	dicProperties("EditBox:Description")="Modified Description"
'@@												dicProperties("EditBox:Name")="Modified Name"
'@@												dicProperties("CheckOut)="Yes"
'@@												Call Fn_Web_EditProperties(dicProperties)
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									26-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'Function Fn_Web_EditProperties(dicProperties)
'   Dim ObjInfo,strWEBMenuPath,strMenu,ObjEdit,ObjChkbx
'   Dim dicItems,dicKeys,iCounter,arrKeys,iRowCount,iCount,crrCellData,bFlag:bFlag=False
'	Fn_Web_EditProperties=False
'
'	If Not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectProperties").Exist(10) Then
'		If Not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").Exist(5) Then
'			strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
'			strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "EditProperties")
'			'If Properties Dialog Not Appearing On Screen The Calling Menu Option "Edit:Properties"
'			Call Fn_Web_MenuOperation("Select",strMenu)
'		End If
'	End If
'	If dicProperties("CheckOut")<>"Yes" Then
'		Call Fn_Web_CheckOutObject("","")
'	End If
'
'	wait(2)
'	If Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectProperties").Exist(5) Then
'		Set ObjInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectProperties")
'	ElseIf Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").Exist(5) Then
'		Set ObjInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object")
'	Else
'		Exit Function
'	End If
'	wait(2)
'	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebElement("PropertyTabs"),"innertext","All")
'	If Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebElement("PropertyTabs").Exist(4) Then
'		Call Fn_Web_UI_WebElement_Click("Fn_Web_EditProperties", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object"), "PropertyTabs", "","","")
'	End If
'
'	dicItems=dicProperties.Items
'	dicKeys=dicProperties.Keys
'	For iCounter=0 To dicProperties.Count-1
'			If dicItems(iCounter)<>"" Then
'				bFlag=False
'			Else
'				bFlag=True
'			End If
'			'Checking Key is Exist in Dectionary or Not
'			If IsNull(dicKeys(iCounter))=False Then
'				'Checking For Key value is Associated or not
'				If dicItems(iCounter)<>"" Then
'					arrKeys=Split(dicKeys(iCounter),":")
'					iRowCount=ObjInfo.RowCount
'					For iCount=0 To iRowCount
'						crrCellData=ObjInfo.GetCellData(iCount,1)
'						If Trim(arrKeys(0))<>"CheckOut" Then
'							If Trim(crrCellData)=Trim(arrKeys(1))+":" Then
'								Select Case arrKeys(0)
'										Case "EditBox"
'												Set ObjEdit=ObjInfo.ChildItem(iCount,2,"WebEdit",0)
'												If TypeName(ObjEdit)<>"Nothing" Then
'													ObjEdit.Set dicItems(iCounter)
'												End If
'												Set ObjEdit=Nothing
'												bFlag=True
'												Exit For
'											Case  "CheckBox"
'												Set ObjChkbx=ObjInfo.ChildItem(iCount,2,"WebCheckBox",0)
'												If TypeName(ObjEdit)<>"Nothing" Then
'													ObjChkbx.Set dicItems(iCounter)
'												End If
'												Set ObjEdit=Nothing
'												bFlag=True
'												Exit For
'								End Select
'							End If
'						Else
'							bFlag=True
'							Exit For
'						End If
'					Next	
'					If bFlag=False Then
'						If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Close").Exist Then
'							Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
'							Fn_Web_EditProperties=False
'							Exit For
'						Else
'							Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")
'							Fn_Web_EditProperties=False
'							Exit For		
'						End If						
'					End If
'				End If
'			End If		
'	Next
'	If bFlag=True Then
'		If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("SaveAndCheckIn").Exist Then
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("SaveAndCheckIn").Click
'		Fn_Web_EditProperties=True
'	Else
'		Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
'		Fn_Web_EditProperties=True
'	End If
'	End If
'	Set ObjInfo=Nothing
'End Function

'Modified Functio By : Sandeep Navghane : 29-Aug-2012
Function Fn_Web_EditProperties(dicProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_EditProperties"
   Dim ObjTable,strMenu,ObjEdit,ObjChkbx,objPropTable
   Dim dicItems,dicKeys,iCounter,arrKeys,iRowCount,iCount,crrCellData,bFlag:bFlag=False
   Dim ObjCheckOut,objWshell,i,sChar
   Dim arrValues, iCnt,ObjMyTcPage
	Fn_Web_EditProperties=False

'	Set ObjTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object")
'	Set objPropTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AllProperties")
'	Set ObjCheckOut=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CheckOut")
	
	Set ObjTable = Fn_SISW_Web_GetObject("Object")
	Set objPropTable = Fn_SISW_Web_GetObject("AllProperties")
	Set ObjCheckOut = Fn_SISW_Web_GetObject("CheckOut")
	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
    
	 'Checking Existance Of Properties Dialog and Checkout Dialog
' 	  If Not ObjTable.Exist(5) and not objPropTable.Exist(5)Then
	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "EditProperties")
	If Not ObjTable.Exist(SISW_MIN_TIMEOUT) and not objPropTable.Exist(SISW_MICRO_TIMEOUT) and not ObjCheckOut.Exist(SISW_MICRO_TIMEOUT) Then
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(2)
   	End If
   	
	If ObjCheckOut.Exist(2) Then
		If dicProperties("CheckOut")<>"Yes" Then
			Call Fn_Web_CheckOutObject("","")
		End If
	End If

'	If dicProperties("CheckOut")<>"Yes" Then
'		Call Fn_Web_CheckOutObject("","")
'	End If
	Set objWshell=CreateObject("WScript.Shell")
'	wait(WEB_MICROLESS_TIMEOUT)
    
	 'Clicking On "All" Tab
	'   Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext","All")
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",ObjMyTcPage.WebElement("FormType"),"innertext","All")
	'   If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Exist(4) Then
	If ObjMyTcPage.WebElement("FormType").Exist(4) Then
'	   	Call Fn_Web_UI_WebElement_Click("Fn_Web_EditProperties", Browser("TeamcenterWeb").Page("MyTeamCenter"), "FormType", "","","")
	   	Call Fn_Web_UI_WebElement_Click("Fn_Web_EditProperties", ObjMyTcPage, "FormType", "","","")
		Set ObjTable=nothing
		'Added Code to handle properties of All tab
'		Set ObjTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectProperties")
		Set ObjTable=Fn_SISW_Web_GetObject("Object").WebTable("ObjectProperties")
	Else
		Set ObjTable=objPropTable
	End If

	dicItems=dicProperties.Items
	dicKeys=dicProperties.Keys
	For iCounter=0 To dicProperties.Count-1
		If dicItems(iCounter)<>"" Then
			bFlag=False
		Else
			bFlag=True
		End IF
		'Checking Key is Exist in Dectionary or Not
		If IsNull(dicKeys(iCounter))=False Then
			'Checking For Key value is Associated or not
			If dicItems(iCounter)<>"" Then
				arrKeys=Split(dicKeys(iCounter),":")
				iRowCount=ObjTable.RowCount
				For iCount=0 To iRowCount
					crrCellData=ObjTable.GetCellData(iCount,1)
					 If Trim(arrKeys(0))<>"CheckOut" Then
						If Trim(crrCellData)=Trim(arrKeys(1))+":" Then
							Select Case arrKeys(0)
									' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									Case "ListVerify"
'											ObjTable.WebElement("Property_Label").SetTOProperty "innertext", Trim(arrKeys(1))+":"
											Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",ObjTable.WebElement("Property_Label"),"innertext",Trim(arrKeys(1))+":")
											Call Fn_Web_UI_Button_Click("Fn_Web_EditProperties",ObjTable,"WebButton")
                                           							arrValues = Split(dicItems(iCounter),"~")
											For iCnt = 0 to UBound(arrValues)
												bFlag = False
'												ObjTable.WebElement("PropertyValue").SetTOProperty "innertext", arrValues(iCnt)
												Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",ObjTable.WebElement("PropertyValue"),"innertext",arrValues(iCnt))
												If ObjTable.WebElement("PropertyValue").Exist(4) Then
													bFlag=True
												End If
											Next
											objWshell.SendKeys "{ESC}"
											Exit For
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									Case "DropdownSelect"
'											ObjTable.WebElement("Property_Label").SetTOProperty "innertext", Trim(arrKeys(1))+":"
											Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",ObjTable.WebElement("Property_Label"),"innertext",Trim(arrKeys(1))+":")
											Call Fn_Web_UI_Button_Click("Fn_Web_EditProperties",ObjTable,"WebButton")
'											ObjTable.WebElement("PropertyValue").SetTOProperty "innertext", dicItems(iCounter)
											Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",ObjTable.WebElement("PropertyValue"),"innertext",dicItems(iCounter))
											If ObjTable.WebElement("PropertyValue").Exist(4) Then
												ObjTable.WebElement("PropertyValue").Click 1,1
											End If
											bFlag=True
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									Case "LinkVerify"
'											ObjTable.WebElement("Property_Label").SetTOProperty "innertext", Trim(arrKeys(1))+":"
											Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",ObjTable.WebElement("Property_Label"),"innertext",Trim(arrKeys(1))+":")
											If Instr(1, ObjTable.Link("PropertyLink").GetROProperty("innertext"), dicItems(iCounter)) > 0 Then
												bFlag=True
											End If
											Exit For
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									Case "EditBox"
											Set ObjEdit=ObjTable.ChildItem(iCount,2,"WebEdit",0)
											If TypeName(ObjEdit)<>"Nothing" Then
                                            								ObjEdit.Click 
												Wait WEB_MICRO_TIMEOUT
												ObjEdit.object.Focus
												wait WEB_MICROLESS_TIMEOUT
												objWshell.SendKeys "{HOME}"&"+"&"{END}"
												Wait WEB_MICRO_TIMEOUT
                                              								objWshell.SendKeys "^a"
												Wait WEB_MICRO_TIMEOUT
'												objWshell.SendKeys dicItems(iCounter)
'												objWshell.SendKeys "{ENTER}"
'												ObjEdit.Set dicItems(iCounter)
												Set WshShell = CreateObject("WScript.Shell")
												For itr=1 to Len(dicItems(iCounter))
													WshShell.SendKeys Mid(dicItems(iCounter),itr,1)
												Next
												wait 1
'												ObjTable.WebElement("PropertyValue").SetTOProperty "innertext", dicItems(iCounter)
												Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",ObjTable.WebElement("PropertyValue"),"innertext",dicItems(iCounter))
												If ObjTable.WebElement("PropertyValue").Exist(1) Then
													ObjTable.WebElement("PropertyValue").Click
												End If
'												ObjTable.WebElement("PropertyValue").SetTOProperty "innertext", ""
												Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_EditProperties",ObjTable.WebElement("PropertyValue"),"innertext","")
'												wait WEB_MINLESS_TIMEOUT
												Set WshShell = nothing				
											End If
											Set ObjEdit=Nothing
											bFlag=True
											Exit For
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									Case "EditBox_WithSpecialCharacters"
											Set ObjEdit=ObjTable.ChildItem(iCount,2,"WebEdit",0)
											If TypeName(ObjEdit)<>"Nothing" Then
												ObjEdit.Set ""
'												Wait WEB_MICRO_TIMEOUT
                                               							ObjEdit.Click 
												Wait WEB_MICRO_TIMEOUT
												ObjEdit.object.Focus
												wait WEB_MICRO_TIMEOUT
												For i=1 to len(dicItems(iCounter))
													sChar = mid(dicItems(iCounter), i, 1)
													If Asc(sChar) = 37 Then
														objWshell.SendKeys "+{%}"	
													Elseif  Asc(sChar) = 43 then
														objWshell.SendKeys "+{+}"
													Else
														objWshell.SendKeys Chr(Asc(sChar))
														wait WEB_MICRO_TIMEOUT
													End if
												Next
												'objWshell.SendKeys "{ENTER}" commented by shweta on 18-sep-2014
											End If
											Set ObjEdit=Nothing
											bFlag=True
											Exit For
										' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
									Case  "CheckBox"
											Set ObjChkbx=ObjTable.ChildItem(iCount,2,"WebCheckBox",0)
											If TypeName(ObjEdit)<>"Nothing" Then
												ObjChkbx.Set dicItems(iCounter)
											End If
											bFlag=True
											Exit For
							End Select
						End If
					Else
						bFlag=True
						Exit For
					End If
				Next	
				If bFlag=False Then
'					If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Close").Exist Then
					If ObjMyTcPage.WebElement("ButtunPanel").WebButton("Close").Exist Then
'						Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
						Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", ObjMyTcPage.WebElement("ButtunPanel"), "Close")
						Fn_Web_EditProperties=False
						Exit For
					Else
'						Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")
						Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", ObjMyTcPage.WebElement("ButtunPanel"), "Cancel")
						Fn_Web_EditProperties=False
						Exit For		
					End If						
				End If
			End If
		End If		
	Next
	If bFlag=True Then
'		wait WEB_MINLESS_TIMEOUT
'		If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("SaveAndCheckIn").Exist Then
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("SaveAndCheckIn").Click
		If ObjMyTcPage.WebElement("ButtunPanel").WebButton("SaveAndCheckIn").Exist Then
			ObjMyTcPage.WebElement("ButtunPanel").WebButton("SaveAndCheckIn").Click
			Fn_Web_EditProperties=True
		Else
			Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", ObjMyTcPage.WebElement("ButtunPanel"), "OK")
			'Call  Fn_Web_UI_Button_Click("Fn_Web_EditProperties", ObjMyTcPage.WebElement("ButtunPanel"), "OK")
			Fn_Web_EditProperties=True
		End If
	End If
	wait WEB_MICRO_TIMEOUT
	Set ObjTable=Nothing
	Set ObjEdit=Nothing
	Set objWshell=Nothing
	Set ObjMyTcPage = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CheckOutObject
'@@
'@@    Description				 :	Function Used to Check Out  Object
'@@
'@@    Parameters			   :	1.strChangeID: New Change ID
'@@												 2.strReason: Check Out Reason
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Object Should be Selected
'@@
'@@    Examples					:	Call Fn_Web_CheckOutObject("1001","Function Test")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									18-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_CheckOutObject(strChangeID,strReason)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CheckOutObject"
   'Variable Declaration
   Dim ObjCheckOut,strWEBMenuPath,strMenu,iCounter,bFlag
   Dim ObjMyTcPage
   
   Fn_Web_CheckOutObject=False
   'Creating Object "Check-Out" Table
'	Set ObjCheckOut=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CheckOut")
	Set ObjCheckOut = Fn_SISW_Web_GetObject("CheckOut")
	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	'Checking Existance Of "Check Out" Dialog
	If Not ObjCheckOut.Exist(3) Then
'	If Fn_Web_UI_ObjectExist("Fn_Web_CheckOutObject",ObjCheckOut)=False Then
'		If Not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Object").WebTable("ObjectInfo").Exist(1) Then
		If Fn_Web_UI_ObjectExist("Fn_Web_CheckOutObject",ObjMyTcPage.WebTable("Object").WebTable("ObjectInfo"))=False Then
			strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
			strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "ToolsCheckOut")
			'If Properties Dialog Not Appearing On Screen The Calling Menu Option "Tools:CheckOut"
			Call Fn_Web_MenuOperation("Select",strMenu)
			Call Fn_Web_ReadyStatusSync(2)
			bFlag=False
			For iCounter=0 to 2
'				If ObjCheckOut.Exist(5) Then
				If Fn_Web_UI_ObjectExist("Fn_Web_CheckOutObject",ObjCheckOut)=True Then
					bFlag=True
					Exit for
				End If
			Next
			If bFlag=False Then
				Fn_Web_CheckOutObject=True
				Exit Function
			End If
		Else
			Fn_Web_CheckOutObject=True
			Exit Function
		End If
	End If
	If strChangeID<>"" Then
		'Setting Change ID
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CheckOutObject", ObjCheckOut, "ChangeID", strChangeID)
	End If
	If strReason<>"" Then
		'Setting Reason For Object Check Out
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CheckOutObject", ObjCheckOut, "Reason", strReason)
	End If
	'Clicking "OK" Button
'	Call Fn_Web_UI_Button_Click("Fn_Web_CheckOutObject", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
	Call Fn_Web_UI_Button_Click("Fn_Web_CheckOutObject", ObjMyTcPage.WebElement("ButtunPanel"), "OK")
	For iCounter=0 To 2
		If ObjCheckOut.Exist(2) Then
'		If Fn_Web_UI_ObjectExist("Fn_Web_CheckOutObject",ObjCheckOut)=True Then
			Wait(WEB_MICROLESS_TIMEOUT)
		Else
			Exit For
		End If
	Next
	'Function Returns True
	Fn_Web_CheckOutObject=True
	'Releasing Object of "Check-Out" Table
	Set ObjCheckOut=Nothing
	Set ObjMyTcPage=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SearchTypeSelect
'@@
'@@    Description				 :	Function Used to Select Search Type
'@@
'@@    Parameters			   :	1.strSearchType : Search Type
'@@
'@@    Return Value		   	   : 	True Or False Or RowIndex
'@@
'@@    Pre-requisite			:	Search be Log In Thin Client						
'@@
'@@    Examples					:	Call Fn_Web_SearchTypeSelect("Item...")
'@@												Call Fn_Web_SearchTypeSelect("Item Revision...")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									19-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_Web_SearchTypeSelect(strSearchType)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SearchTypeSelect"
	'Variable Declaration
	Dim ObjSearchType, objMyTcPage
	'Function Returns Flase
	Fn_Web_SearchTypeSelect=False
	Set objMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	'Creating Object "SearchTypes" Web Table
'	Set ObjSearchType=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SavedSearches").WebTable("SearchTypes")
	Set ObjSearchType=objMyTcPage.WebTable("SavedSearches").WebTable("SearchTypes")
	'Checking Existance Of "Change Search" Dialog
	If Not ObjSearchType.Exist(WEB_MIN_TIMEOUT) Then
		'If "Change Search" Dialog Not Exist Then Opening Dialog
		Call Fn_Web_UI_WebElement_Click("Fn_Web_SearchTypeSelect", Browser("TeamcenterWeb"), "AdvanceSearch", "","","")
		wait WEB_MICROLESS_TIMEOUT
'		Call Fn_Web_UI_Link_Click("Fn_Web_SearchTypeSelect",Browser("TeamcenterWeb"), "More","","","")
		Browser("TeamcenterWeb").Link("More").Click 1,1,micLeftBtn
		wait WEB_MIN_TIMEOUT
'		wait WEB_DEFAULT_TIMEOUT
'		Browser("TeamcenterWeb").Link("More").highlight
'		Dim objMDR,iX,iY
'		Set objMDR = CreateObject("Mercury.DeviceReplay")
'		iX=Browser("TeamcenterWeb").Link("More").GetROProperty("abs_x")
'		iY=Browser("TeamcenterWeb").Link("More").GetROProperty("abs_y")
'		Set objMDR = CreateObject("Mercury.DeviceReplay")
'		objMDR.MouseClick iX+25,iY,0
	End If
	'Selecting "System Define Searches" Tab
'    	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SearchTypeSelect",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext","System Defined Searches")
'	Call Fn_Web_UI_WebElement_Click("Fn_Web_SearchTypeSelect",Browser("TeamcenterWeb").Page("MyTeamCenter"), "FormType", "","","")
	
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SearchTypeSelect",objMyTcPage.WebElement("FormType"),"innertext","System Defined Searches")
	Call Fn_Web_UI_WebElement_Click("Fn_Web_SearchTypeSelect",objMyTcPage, "FormType", "","","")
	If strSearchType<>"" Then
		'Selecting Search Type
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SearchTypeSelect",ObjSearchType.Link("SearchTypeName"),"innertext",strSearchType)
		If ObjSearchType.Link("SearchTypeName").Exist(5) Then
			Call Fn_Web_UI_Link_Click("Fn_Web_SearchTypeSelect",ObjSearchType, "SearchTypeName","","","")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Search Type [ "+strSearchType+" ]")
			Fn_Web_SearchTypeSelect=True
			wait(WEB_MICROLESS_TIMEOUT)
		Else
'			Call Fn_Web_UI_Button_Click("Fn_Web_SearchTypeSelect",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Close")
			Call Fn_Web_UI_Button_Click("Fn_Web_SearchTypeSelect",objMyTcPage.WebElement("ButtunPanel"),"Close")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Invalid Search Type [ "+strSearchType+" ]")
		End If
	Else
'		Call Fn_Web_UI_Button_Click("Fn_Web_SearchTypeSelect",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Close")
		Call Fn_Web_UI_Button_Click("Fn_Web_SearchTypeSelect",objMyTcPage.WebElement("ButtunPanel"),"Close")
	End If
	If ObjSearchType.Link("SearchTypeName").Exist(1) Then
'		Call Fn_Web_UI_Button_Click("Fn_Web_SearchTypeSelect",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Close")
		'Call Fn_Web_UI_Button_Click("Fn_Web_SearchTypeSelect",objMyTcPage.WebElement("ButtunPanel"),"Close")
		objMyTcPage.WebElement("ButtunPanel").WebButton("Close").Click
	End if
	'Releasing Object of "SearchTypes" Web Table
	Set objMyTcPage = Nothing
	Set ObjSearchType=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SearchResultOperations
'@@
'@@    Description				 :	Function Used to Perform Operations On Search Results
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@												  2.strItem : Item Name
'@@
'@@    Return Value		   	   : 	True Or False Or RowIndex
'@@
'@@    Pre-requisite			:	Search Results should be Display						
'@@
'@@    Examples					:	Call Fn_Web_SearchResultOperations("Select","000021-Item1")
'@@												Call Fn_Web_SearchResultOperations("Verify","000021-Item1")
'@@												Call Fn_Web_SearchResultOperations("MultiSelect","000021-Item1:000022-Item2")
'@@												Call Fn_Web_SearchResultOperations("GetAllColumnNames","")
'@@												Call Fn_Web_SearchResultOperations("ClickLink","000018-Unix")
'@@												Call Fn_Web_SearchResultOperations("DeSelect","000021-Item1")
'@@												Call Fn_Web_SearchResultOperations("GetChildrenList","")
'@@												Call Fn_Web_SearchResultOperations("ExpandNode","Training:Preferred Items:000070-I5")
'@@												Call Fn_Web_SearchResultOperations("SelectNode","Training:Preferred Items:000070-I5:000070/A;1-I5")
'@@												Call Fn_Web_SearchResultOperations("VerifyNode","Training:Preferred Items:000070-I5:000070/A;1-I5")
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done												Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									19-Apr-2011						1.0																												Sunny Ruparel
'@@												Sandeep Navghane									24-Apr-2011						1.0							Added Cases "GetAllColumnNames"			  Sunny Ruparel
'@@												Sandeep Navghane									24-Apr-2011						1.0							Added Cases "ClickLink"			  							Sunny Ruparel
'@@												Deepak kumar											 26-Apr-2011						1.0							Added Case "Delete"
'@@												Sandeep Navghane									09-May-2011						1.0							Added Cases "DeSelect"			  							Sunny Ruparel
'@@												Nilesh Gadekar											25-Sep-2012					1.0							Added Cases "GetChildrenList"			  							Sandeep Navghane
'@@												Sandeep Navghane									07-May-2013					1.0							Added Cases "VerifyNode"			  							Sneha C
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_SearchResultOperations(strAction,strItem)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SearchResultOperations"
	'Variable Declaration
	Dim ObjMTcPage,ObjResultTb
	Dim iRwCount,iCounter,currCellValue,iRowNum,ObjChk,arrItem,iCount,iColCount,ColName,arrColName
	Dim iniWTCount,iniWTCount1
	Dim objSelectType,objSelectType1,intNoOfObjects,intNoOfObjects1,obj,rowid
	Dim iX,iY,objMDR,iH,iW
	Dim iLength,i,iHeight

	'Initially Function Returns False
	Fn_Web_SearchResultOperations=False
	
	'Creating Object Of "MyTeamCenter" Page And "SearchResult" WebTable
'	Set ObjMTcPage=Browser("TeamcenterWeb").Page("MyTeamCenter")
'	Set ObjResultTb=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchResult")
	Set ObjMTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	Set ObjResultTb = Fn_SISW_Web_GetObject("SearchResult@1")
'	ObjMTcPage.WebButton("LoadAll").SetTOProperty "name","Load All"
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ItemBasicCreate",ObjMTcPage.WebButton("LoadAll"),"name","Load All")
	'Checking Existance Of "LoadAll" Button
	If ObjMTcPage.WebButton("LoadAll").Exist(2) Then
	'If Fn_Web_UI_ObjectExist("Fn_Web_SearchResultOperations",ObjMTcPage.WebButton("LoadAll"))=True Then
		ObjMTcPage.WebButton("LoadAll").Click 1,1,micLeftBtn
		If Browser("TeamcenterWeb").Dialog("Dialog").Exist(2) Then
		'If Fn_Web_UI_ObjectExist("Fn_Web_SearchResultOperations",Browser("TeamcenterWeb").Dialog("Dialog"))=True Then
			Call Fn_Web_ErrorMsgVerify("","OK")
		End If
	End If
	'Cases For Table Operations	
   	Select Case strAction
		 	'Case to Retrieve Row Index
		 	Case "GetRowIndex"  'Fn_Web_SearchResultOperations("GetRowIndex","000021-Item1")
				iRwCount=ObjResultTb.RowCount()
				For iCounter=1 To iRwCount
					currCellValue=ObjResultTb.GetCellData(iCounter,2)
					If Trim(currCellValue)=Trim(strItem) Then
						Fn_Web_SearchResultOperations=iCounter
						Exit For
					End If
				Next
			'------------------------------------------------------------------------------------------------------------------------
			'Case To Select Item
		 	Case "Select" 'Fn_Web_SearchResultOperations("Select","000021-Item1")
				iRowNum=Fn_Web_SearchResultOperations("GetRowIndex",strItem)
				If iRowNum<>"" Then
					Set ObjChk=ObjResultTb.ChildItem(CInt(iRowNum), 1,"WebCheckBox", 0)
					If TypeName(ObjChk) <> "Nothing" Then
						If ObjChk.GetROProperty("checked") = "0" Then
							ObjChk.Click 1, 1, micLeftBtn
							Fn_Web_SearchResultOperations=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Item [ "+strItem+" ]")
						ElseIf ObjChk.GetROProperty("checked") = "1" Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Item [ "+strItem+" ]")
							Fn_Web_SearchResultOperations=True
						End If
					End If
				End If
				Wait WEB_MICRO_TIMEOUT
			'------------------------------------------------------------------------------------------------------------------------
			'Case To Select Item
		 	Case "DeSelect" 'Fn_Web_SearchResultOperations("DeSelect","000021-Item1")
				iRowNum=Fn_Web_SearchResultOperations("GetRowIndex",strItem)
				If iRowNum<>"" Then
					Set ObjChk=ObjResultTb.ChildItem(CInt(iRowNum), 1,"WebCheckBox", 0)
					If TypeName(ObjChk) <> "Nothing" Then
						If ObjChk.GetROProperty("checked") = "1" Then
							ObjChk.Click 1, 1, micLeftBtn
							Fn_Web_SearchResultOperations=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully DeSelected Item [ "+strItem+" ]")
						ElseIf ObjChk.GetROProperty("checked") = "1" Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully DeSelected Item [ "+strItem+" ]")
							Fn_Web_SearchResultOperations=True
						End If
					End If
				End If
			'------------------------------------------------------------------------------------------------------------------------
			'Case to Verify Item is Exist In Table
			Case "Verify" 'Fn_Web_SearchResultOperations("Verify","000021-Item1")
				iRowNum=Fn_Web_SearchResultOperations("GetRowIndex",strItem)
				If CBool(iRowNum)=True Then
					Fn_Web_SearchResultOperations=True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Found Item [ "+strItem+" ]")
				End If
			'------------------------------------------------------------------------------------------------------------------------
			'Case to MultiSelect Items
			Case "MultiSelect" 'Fn_Web_SearchResultOperations("MultiSelect","000021-Item1:000022-Item2")
				arrItem=Split(strItem,":")
				For iCount=0 To UBound(arrItem)
					iRowNum=Fn_Web_SearchResultOperations("GetRowIndex",arrItem(iCount))
					If iRowNum<>"" Then
						Set ObjChk=ObjResultTb.ChildItem(CInt(iRowNum), 1,"WebCheckBox", 0)
						If TypeName(ObjChk) <> "Nothing" Then
							If ObjChk.GetROProperty("checked") = "0" Then
								ObjChk.Click 1, 1, micLeftBtn
								Fn_Web_SearchResultOperations=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully MultiSelected Item [ "+strItem+" ]")
							ElseIf ObjChk.GetROProperty("checked") = "1" Then
								Fn_Web_SearchResultOperations=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully MultiSelected Item [ "+strItem+" ]")
							End If
						End If
					End If	
				Next
				Wait WEB_MICRO_TIMEOUT
			'------------------------------------------------------------------------------------------------------------------------
        			Case "ClickLink"
				iRowNum=Fn_Web_SearchResultOperations("GetRowIndex",strItem)
				If iRowNum<>"" Then
					Set ObjChk=ObjResultTb.ChildItem(CInt(iRowNum),2,"Link", 0)
					If TypeName(ObjChk) <> "Nothing" Then
						ObjChk.Click 1, 1, micLeftBtn
						wait(WEB_MICROLESS_TIMEOUT)
						Fn_Web_SearchResultOperations=True
					End If
				End If
			'------------------------------------------------------------------------------------------------------------------------
      			Case "GetAllColumnNames"
'				iColCount=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchResult").ColumnCount(1)
				iColCount=ObjResultTb.ColumnCount(1)
				For iCounter=2 To iColCount
'					ColName=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchResult").GetCellData(1,iCounter)
					ColName=ObjResultTb.GetCellData(1,iCounter)
					If iCounter=2 Then
						arrColName=ColName
					Else
						arrColName=arrColName+":"+ColName
					End If
				Next
				Fn_Web_SearchResultOperations=arrColName
			'------------------------------------------------------------------------------------------------------------------------
			Case "Delete" 'Case returns True After Deletion Item
				Call Fn_Web_SearchResultOperations("Select", strItem)
'				Call Fn_Web_ReadyStatusSync(1)
				Call Fn_Web_MenuOperation("Select","Edit:Delete")
'				Call Fn_Web_ReadyStatusSync(1)
				Wait WEB_MICRO_TIMEOUT
'				If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("OK").Exist Then
				If ObjMTcPage.WebElement("ButtunPanel").WebButton("OK").Exist(2) Then
'					Call Fn_Web_UI_Button_Click("Fn_Web_SearchResultOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
					Call Fn_Web_UI_Button_Click("Fn_Web_SearchResultOperations", ObjMTcPage.WebElement("ButtunPanel"), "OK")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Deleted Item [ "+strItem+" ]")
					Fn_Web_SearchResultOperations = True
				Else 
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Failed to  Delete Item [ "+strItem+" ]")
					Fn_Web_SearchResultOperations= False
				End If
			'------------------------------------------------------------------------------------------------------------------------
       			Case "GetChildrenList"
				sString=""
				iRwCount=ObjResultTb.RowCount()
				'For iCounter=1 To iRwCount-1  'Modified by Nilesh on 1-April-2013
				For iCounter=1 To iRwCount
					currCellValue=ObjResultTb.GetCellData(iCounter,2)
					sString=sString+currCellValue
					'If iCounter<>iRwCount-1 And currCellValue<>"" Then
					If iCounter<>iRwCount And currCellValue<>"" Then
						sString=sString+"~"
					End If
				Next
				If sString<>"" Then
					Fn_Web_SearchResultOperations =sString
				Else
					Fn_Web_SearchResultOperations = False
				End If
			'------------------------------------------------------------------------------------------------------------------------
           		 Case "DoubleClickProject"
'				Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject").SetTOProperty "innertext",strItem
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ItemBasicCreate",Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject"),"innertext",strItem)
				If Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject").Exist(5)  then
					If Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject").GetROProperty("height")>0 then
					        Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject").Click
					        Wait WEB_MICROLESS_TIMEOUT
					        Set objMDR = CreateObject("Mercury.DeviceReplay")
						iX=Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject").GetROProperty("abs_x")
						iY=Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject").GetROProperty("abs_y")
						'*Added by Nilesh on 3-June-2013 for IE9 Support
				       		iH=Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject").GetROProperty("height")'Calculate Height
						iW=Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultProject").GetROProperty("width")'Calculate Width
						objMDR.MouseDblClick Cint(iX+Cint(iW/2)),Cint(iY+Cint(iH/2)),LEFT_MOUSE_BUTTON  
				'		objMDR.MouseDblClick iX+30,iY+10,LEFT_MOUSE_BUTTON
						'*End
						Set objMDR = Nothing
						Wait WEB_MIN_TIMEOUT
				       		If Err.number<0 Then
							Fn_Web_SearchResultOperations=False
						Else
							Fn_Web_SearchResultOperations=True
						End If
					End If
				End If
	        		'------------------------------------------------------------------------------------------------------------------------
			Case "SelectNode"
				arrItem=Split(strItem,":")
				Set objSelectType=description.Create()
				objSelectType("micClass").value = "WebTable"
				Set  intNoOfObjects = Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultsTreePanel").ChildObjects(objSelectType)
				iLength=0
				For iCounter=0 to ubound(arrItem)
					bFlag=False
					For i=iLength to intNoOfObjects.Count-1
						If trim(arrItem(iCounter))=trim(intNoOfObjects(i).getroproperty("innertext")) Then
							iHeight = intNoOfObjects(i).getroproperty("height")
							If iHeight > 0 Then
								If iCounter=cint(ubound(arrItem)) Then
									Set objSelectType1=description.Create()
									objSelectType1("micClass").value = "WebElement"
	    								objSelectType1("innertext").value = arrItem(iCounter)
									Set  intNoOfObjects1 = intNoOfObjects(i).ChildObjects(objSelectType1)
									intNoOfObjects1(2).Click 1,1
									wait WEB_MIN_TIMEOUT
									Set  intNoOfObjects1 = Nothing
									Set  objSelectType1 = Nothing
								End If
								iLength=i+1
								bFlag=True
								Exit for
							Else
								bFlag=False
								Exit for
							End If
						End If
					Next
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Web_SearchResultOperations=True
				End If
				Set objSelectType=Nothing
	      		'------------------------------------------------------------------------------------------------------------------------
			Case "VerifyNode"
				arrItem=Split(strItem,":")
	'			iLength=UBound(arrNode)
				Set objSelectType=description.Create()
				objSelectType("micClass").value = "WebTable"
				Set  intNoOfObjects = Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultsTreePanel").ChildObjects(objSelectType)
				iLength=0
				For iCounter=0 to ubound(arrItem)
					bFlag=False
					For i=iLength to intNoOfObjects.Count-1
						If trim(arrItem(iCounter))=trim(intNoOfObjects(i).getroproperty("innertext")) Then
							iHeight = intNoOfObjects(i).getroproperty("height")
							If iHeight > 0 Then
								iLength=i+1
								bFlag=True
								Exit for
							Else
								bFlag=False
								Exit for
							End If
						End If
					Next
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_Web_SearchResultOperations=True
				End If
				Set objSelectType=Nothing
       		'------------------------------------------------------------------------------------------------------------------------
		Case "ExpandNode"
			arrItem=Split(strItem,":")
			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("SearchPage").WebElement("SearchResultsTreePanel").ChildObjects(objSelectType)
			iLength=0
			For iCounter=0 to ubound(arrItem)
				bFlag=False
				For i=iLength to intNoOfObjects.Count-1
					If trim(arrItem(iCounter))=trim(intNoOfObjects(i).getroproperty("innertext")) Then
						iHeight = intNoOfObjects(i).getroproperty("height")
						If iHeight > 0 Then
							If iCounter=cint(ubound(arrItem)) Then
								If instr(1,lcase(intNoOfObjects(i).getroproperty("class")),"expanded") Then
									'do nothing
									Fn_Web_SearchResultOperations=True
								Else
									rowid = intNoOfObjects(i).GetRowWithCellText(arrItem(iCounter))
									Set obj = intNoOfObjects(i).ChildItem(rowid, cint(ubound(arrItem))+1,"WebElement",0)
									If typename(obj)<>"Nothing" Then
										obj.click
										wait WEB_MIN_TIMEOUT
										Fn_Web_SearchResultOperations=True
									End If
								End If
							End if
							iLength=i+1
							bFlag=True
							Exit for
						Else
							bFlag=False
							Exit for
						End If
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			Set objSelectType=Nothing
   End Select
	'Releasing Object Of "MyTeamCenter" Page And "SearchResult" WebTable
	Set ObjMTcPage=Nothing
	Set ObjResultTb=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SearchOperation
'@@
'@@    Description				 :	Function Used to Search Objects
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@												  2.dicWebSearch : Dectionary [ Key - Value Pair ]
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Search be Log In Thin Client						
'@@
'@@    Examples					:	dicWebSearch("EditBox:Name")="Item1"
'@@												dicWebSearch("EditBox:Description")="Demo Item"
'@@												dicWebSearch("EditBox:Item ID")="000024"
'@@												Call Fn_Web_SearchOperation("Search",dicWebSearch)
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									19-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_SearchOperation(strAction,dicWebSearch)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SearchOperation"
	'Variable Declaration
	Dim ObjSrchCriteria,ObjEdit,ObjButton,objMyTcPage
	Dim dicItems,dicKeys,iCounter,bFlag,arrKeys,iRwCount,iCount,crrCellValue
	
	Set objMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	'Creating Object "SearchCriteria" Web Table
'	Set ObjSrchCriteria=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchCriteria")
	Set ObjSrchCriteria=objMyTcPage.WebTable("SearchCriteria")
	Select Case strAction
		'Case to Search Object As per User Criteria
		Case "Search"
			'Clicking On Clear button To make All Fields Empty
'			Call  Fn_Web_UI_Button_Click("Fn_Web_SearchOperation", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchButtonTable"), "Clear")
			Call  Fn_Web_UI_Button_Click("Fn_Web_SearchOperation", objMyTcPage.WebTable("SearchButtonTable"), "Clear")
			bFlag=True
			'Taking All Items And Keys From Dectionary Object [ To Support multiple criteia in One function Call ]
			dicItems=dicWebSearch.Items
			dicKeys=dicWebSearch.Keys			
			For iCounter=0 To dicWebSearch.Count-1
				'Checking Key is Exist in Dectionary or Not
				If IsNull(dicKeys(iCounter))=False Then
					'Checking For Key value is Associated or not
					If dicItems(iCounter)<>"" Then
						bFlag=False
						'Splitting Key to Get Control Type And Control Name
						arrKeys=Split(dicKeys(iCounter),":")
						'Taking Row Count From "SearchCriteria" Table
						iRwCount=ObjSrchCriteria.RowCount
						For iCount=0 To iRwCount
							'Retriving Current Cell Value
							crrCellValue=ObjSrchCriteria.GetCellData(iCount,1)
							If Trim(crrCellValue)=Trim(arrKeys(1)+":") Or Trim(crrCellValue)=Trim(arrKeys(1)) Or Trim(crrCellValue)=Trim(arrKeys(1)+".:") Or Trim(crrCellValue)=Trim(arrKeys(1)+".") Then
								Select Case arrKeys(0)
									'Case to Enter Criteria in Edit Box
									Case "EditBox"
										Set ObjEdit=ObjSrchCriteria.ChildItem(iCount,2,"WebEdit",0)
										If TypeName(ObjEdit)<>"Nothing" Then
											ObjEdit.Set dicItems(iCounter)
										End If
										Set ObjEdit=Nothing
										bFlag=True
										Exit For
									'Case to Select Criteria from List
									Case "Button"
										Set ObjButton=ObjSrchCriteria.ChildItem(iCount,2,"WebButton",0)
										If TypeName(ObjButton)<>"Nothing" Then
                                            							ObjButton.Click 1,1
											wait(WEB_MICROLESS_TIMEOUT)
'											Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SearchOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchCriteria").WebElement("SearchCriteria"),"innertext",dicItems(iCounter))
'											Call Fn_Web_UI_WebElement_Click("Fn_Web_SearchOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchCriteria"),"SearchCriteria","","","")
											Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SearchOperation",objMyTcPage.WebTable("SearchCriteria").WebElement("SearchCriteria"),"innertext",dicItems(iCounter))
											Call Fn_Web_UI_WebElement_Click("Fn_Web_SearchOperation",objMyTcPage.WebTable("SearchCriteria"),"SearchCriteria","","","")
										End If
										Set ObjButton=Nothing
										bFlag=True
										Exit For
								End Select
							End If
						Next
					End If
				End If
				If bFlag=False Then
					Fn_Web_SearchOperation=False
					Exit For
				End If
			Next
			If bFlag=True Then
				'Clicking On Finish button to Execute Query
'				Call  Fn_Web_UI_Button_Click("Fn_Web_SearchOperation", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchButtonTable"), "Find")
				Call  Fn_Web_UI_Button_Click("Fn_Web_SearchOperation", objMyTcPage.WebTable("SearchButtonTable"), "Find")
				Wait WEB_MICRO_TIMEOUT
				Fn_Web_SearchOperation=True
			End If
	End Select
	Set ObjSrchCriteria=Nothing
	Set objMyTcPage = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_PasteAs
'@@
'@@    Description				 :	Function Used to Paste As Type
'@@
'@@    Parameters			   :	1.strType: type
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_PasteAs("Alias IDs")
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									20-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_PasteAs(strType)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_PasteAs"
   Dim ObjPasteAs,strWEBMenuPath,strMenu,crrType

	Set ObjPasteAs=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("PasteAs")
	Fn_Web_PasteAs=False
	If Not ObjPasteAs.Exist(5) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "EditPasteAs")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	If strType<>"" Then
		crrType=ObjPasteAs.WebEdit("PasteAsType").GetROProperty("value")
		If Trim(crrType)<>Trim(strType) Then
			Call Fn_Web_UI_Button_Click("Fn_Web_PasteAs",ObjPasteAs,"PasteAsButton")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_PasteAs",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",strType)
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
			wait(2)
		End If
		Fn_Web_PasteAs=True
	End If
	Call Fn_Web_UI_Button_Click("Fn_Web_PasteAs", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
	Set ObjPasteAs=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_BrowserOperations
'@@
'@@    Description				 :	Function Used to Perform Operations On Web Browsers
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@
'@@    Return Value		   	   : 	True Or False Or Browser Title
'@@
'@@    Pre-requisite			:	Environment XML should be properly filled							
'@@
'@@    Examples					:	Call Fn_Web_BrowserOperations("GetTitle")
'@@												Call Fn_Web_BrowserOperations("Back")
'@@												Call Fn_Web_BrowserOperations("Forward")
'@@												Call Fn_Web_BrowserOperations("Refresh")
'@@												Call Fn_Web_BrowserOperations("Exist~Google")
'@@												Call Fn_Web_BrowserOperations("LaunchAndSetURL~http://www.google.co.in/")
'@@												Call Fn_Web_BrowserOperations("CloseTab~Google")
'@@
'@@	   History					 	:	
'@@													Developer Name							Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane						20-Apr-2011						1.0																								Sunny Ruparel
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Koustubh Watwe						    11-May-2011					   1.0							Added case Refresh
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane						11-Jan-2011					   1.2							Added case LaunchAndSetURL
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Ashwini Kumar						    11-Jan-2011					   1.2							Replaced "WScript.Shell" with "Mercury.DeviceReplay" and changed function accordingly.
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_BrowserOperations(strAction)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_BrowserOperations"
    Dim strBrowser,arrAction,ObjBrowser,handlewin
    Dim WshShell,StrTitle,ObjIE

'	Set WshShell = CreateObject("WScript.Shell")
    Set DeviceReplay = CreateObject("Mercury.DeviceReplay")

    arrAction=Split(strAction,"~")
    Fn_Web_BrowserOperations = False
    strBrowser=Environment.Value("WebBrowserName")
    Select Case arrAction(0)
         Case "Back"
            If InStr(1,strBrowser,"IE")>0 Then
'                Browser("application version:=.*internet explorer.*").Back
                wait 1
                DeviceReplay.KeyDown 56				'Alt Down
				wait 1
				DeviceReplay.PressKey 203			'Left Arrow pressed
				DeviceReplay.KeyUp 56					'Alt Up
                wait 1
                DeviceReplay.PressKey 63			'F5 pressed
                wait 1
            ElseIf InStr(1,strBrowser,"FF")>0 Then
                Browser("version:=.*firefox.*").Back
            End If
            Fn_Web_BrowserOperations=True
        Case "Forward"
            If InStr(1,strBrowser,"IE")>0 Then
'                Browser("application version:=.*internet explorer.*").Forward
                wait 1
                DeviceReplay.KeyDown 56
				wait 1
				DeviceReplay.PressKey 205
				DeviceReplay.KeyUp 56
                wait 1
                DeviceReplay.PressKey 63
                wait 1
            ElseIf InStr(1,strBrowser,"FF")>0 Then
                Browser("version:=.*firefox.*").Forward
            End If
            Fn_Web_BrowserOperations=True
        Case "Refresh"
            If InStr(1,strBrowser,"IE")>0 Then
'                Browser("application version:=.*internet explorer.*").Refresh
'                wait 1
'                WshShell.SendKeys "{F5}"
'                wait 1
'				Set objBrowser=Description.Create()
'				objBrowser("micClass").Value="Browser"
'				objBrowser("application version").RegularExpression=True
'				objBrowser("application version").Value=".*internet explorer .*"
'				Set ObjIE=Desktop.ChildObjects(objBrowser)
'				'Print ObjIE.Count
'				If  ObjIE.Count<> 0 Then
'					StrTitle=ObjIE(0).GetROProperty("title")		
'					WshShell.AppActivate StrTitle,5
'					WshShell.SendKeys "{F5}"
'					Fn_Web_BrowserOperations=True
'				Else
'					Fn_Web_BrowserOperations=False
'				End If
'				   wait 1
'				Set objBrowser=Nothing
'				Set ObjIE=Nothing

                DeviceReplay.PressKey 63
				wait 5
            ElseIf InStr(1,strBrowser,"FF")>0 Then
                Browser("version:=.*firefox.*").Refresh
            End If
            Fn_Web_BrowserOperations=True
        Case "GetTitle"
            If InStr(1,strBrowser,"IE")>0 Then
                Fn_Web_BrowserOperations=Browser("application version:=.*internet explorer.*").GetROProperty("name")
            ElseIf InStr(1,strBrowser,"FF")>0 Then
                Fn_Web_BrowserOperations=Browser("version:=.*firefox.*").GetROProperty("name")
            End If
        Case "LaunchAndSetURL"

            If InStr(1,strBrowser,"IE")>0 Then
                Set ObjBrowser = CreateObject("InternetExplorer.Application")
                ObjBrowser.Visible = True
                ObjBrowser.Navigate arrAction(1)
                handlewin = ObjBrowser.HWND
                Window("hwnd:="+CStr(handlewin)).Maximize
                Fn_Web_BrowserOperations=True
            ElseIf InStr(1,strBrowser,"FF")>0 Then
                SystemUtil.Run "firefox.exe"
                Browser("version:=.*Firefox.*").Navigate arrAction(1)
                Fn_Web_BrowserOperations=True
            End If

		Case "PasteURL"				
            If InStr(1,strBrowser,"IE")>0 Then
                Set ObjBrowser = CreateObject("InternetExplorer.Application")
				ObjBrowser.Visible = True
				ObjBrowser.GoHome
				handlewin = ObjBrowser.HWND
                Window("hwnd:="+CStr(handlewin)).Maximize
				wait 2
			If Browser("TeamcenterWeb").Exist Then 
				Browser("TeamcenterWeb").WinEdit("URLEdit").Set ""
			ElseIf Browser("CreationTime:=0").Exist(1) Then
				Browser("CreationTime:=0").WinEdit("Location:=0").Set ""
			End If

				wait 1
                DeviceReplay.KeyDown 29				'Ctrl Down
				wait 1
				DeviceReplay.PressKey 47			'V pressed
				DeviceReplay.KeyUp 29					'Ctrl Up
                wait 1
                DeviceReplay.PressKey 28			'ENTER pressed
                wait 1

'                wait 1
'                WshShell.SendKeys "^(V)"
'                wait 1
'                WshShell.SendKeys "^{ENTER}"
'                wait 1
				Fn_Web_BrowserOperations=True
            ElseIf InStr(1,strBrowser,"FF")>0 Then
                'not yet implemented code
            End If

	    Case "Exist" ' this case need browser title 
			 If InStr(1,strBrowser,"IE")>0 Then
				Fn_Web_BrowserOperations=Browser("title:=.*"&arrAction(1)&".*").Exist(5)
            End If

		Case "CloseTab" 'this case need browser title 
			 If InStr(1,strBrowser,"IE")>0 Then
				If Browser("title:=.*"&arrAction(1)&".*").Exist(5) then
					Browser("title:=.*"&arrAction(1)&".*").Close
					wait 2
					If not Browser("title:=.*"&arrAction(1)&".*").Exist(5) then
						Fn_Web_BrowserOperations=True
					End if
				End if
            End If
   End Select
  If Browser("ExportComplete").Dialog("WindowsInternetExplorer").Exist(5) Then
		Browser("ExportComplete").Dialog("WindowsInternetExplorer").WinButton("Retry").Click 1,1,micLeftBtn
		wait 5
   End If
   If Window("WindowsInternetExplorer").Dialog("WindowsInternetExplorer9").Exist(5) Then
		Window("WindowsInternetExplorer").Dialog("WindowsInternetExplorer9").WinButton("Retry").Click 1,1,micLeftBtn
		wait 5
   End If
'   Set WshShell =nothing
   Set DeviceReplay=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_QuickSearchReasultOperations
'@@
'@@    Description				 :	Function Used to Perform Operations On Quick Search Results
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@												  2.strNode : Object Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Quick Search Results should be Display						
'@@
'@@    Examples					:	Call Fn_Web_QuickSearchReasultOperations("Verify","000009-Item1")
'@@												Call Fn_Web_QuickSearchReasultOperations("Select","000009-Item1")
'@@												Call Fn_Web_QuickSearchReasultOperations("MultiSelect","000009-Item1~000024-newitem")
'@@												Call Fn_Web_QuickSearchReasultOperations("FirstElement","")
'@@												Call Fn_Web_QuickSearchReasultOperations("ClickLink","ECN-000003-test2")
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									21-Apr-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									02-Apr-2011						1.0									Case "ClickLink"							Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_QuickSearchReasultOperations(strAction,strNode)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_QuickSearchReasultOperations"
   'Variable Declaration
   Dim ObjQuickResult,ObjChk, objLink, iColPos, iRwCount2
   Dim iRwCount,iCounter,crrNode,iCount,bFlag,arrNode
   'Creating Object Of "QuickSearchResult" Table
	Set ObjQuickResult=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("QuickSearchResult")
	Fn_Web_QuickSearchReasultOperations=False
	'Checking Existance Of "QuickSearchResult" Table
	If ObjQuickResult.Exist(7) Then
		'Clicking On Load All Button To Load All Search Results
		If Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("LoadAll").Exist(5) Then
			Call Fn_Web_UI_Button_Click("Fn_Web_QuickSearchReasultOperations",Browser("TeamcenterWeb").Page("MyTeamCenter"),"LoadAll")
			Call Fn_Web_ErrorMsgVerify("","OK")
			wait(3)
		End If
		
		Select Case strAction
				Case "Verify" 'Case to Verify Specific Object Exist in Result Table Or Not
					iRwCount=ObjQuickResult.RowCount
					For iCounter=0 To iRwCount
						crrNode=ObjQuickResult.GetCellData(iCounter,2)
						If Trim(crrNode)=Trim(strNode) Then
							Fn_Web_QuickSearchReasultOperations=True
							Exit For
						End If
					Next
				Case "Select"  'Case to Select Specific Object From Result Table
					iRwCount=ObjQuickResult.RowCount
					For iCounter=0 To iRwCount
						crrNode=ObjQuickResult.GetCellData(iCounter,2)
						If Trim(crrNode)=Trim(strNode) Then
							Set ObjChk=ObjQuickResult.ChildItem(iCounter,1,"WebCheckBox",0)
							If TypeName(ObjChk) <> "Nothing" Then
									If ObjChk.GetROProperty("checked") = "0" Then
										ObjChk.Click 1, 1, micLeftBtn
										Fn_Web_QuickSearchReasultOperations=True
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected item [ "+strNode+" ]")
									ElseIf ObjChk.GetROProperty("checked") = "1" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Item [ "+strNode+" ]")
										Fn_Web_QuickSearchReasultOperations=True
									End If
							End If
							Fn_Web_QuickSearchReasultOperations=True
							Set ObjChk=Nothing
							Exit For
						End If
					Next
				Case "MultiSelect"  'Case to MultiSelect Specific Object From Result Table
					arrNode=Split(strNode,"~")
					iRwCount=ObjQuickResult.RowCount
					For iCount=0 To UBound(arrNode)
						bFlag=False
						For iCounter=0 To iRwCount
							crrNode=ObjQuickResult.GetCellData(iCounter,2)
							If Trim(crrNode)=Trim(arrNode(iCount)) Then
								Set ObjChk=ObjQuickResult.ChildItem(iCounter,1,"WebCheckBox",0)
								If TypeName(ObjChk) <> "Nothing" Then
										If ObjChk.GetROProperty("checked") = "0" Then
											ObjChk.Click 1, 1, micLeftBtn
											Fn_Web_QuickSearchReasultOperations=True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected item [ "+strNode+" ]")
										ElseIf ObjChk.GetROProperty("checked") = "1" Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Item [ "+strNode+" ]")
											Fn_Web_QuickSearchReasultOperations=True
										End If
								End If
								bFlag=True
								Fn_Web_QuickSearchReasultOperations=True
								Set ObjChk=Nothing
								Exit For
							End If
						Next
						If bFlag=False Then
							Fn_Web_QuickSearchReasultOperations=False
							Exit For
						End If
					Next
			Case "FirstElement"
						' For Getting the Column Header Position
						iRwCount = ObjQuickResult.GetROProperty("rows")
						For iCounter = 1  To iRwCount
							If ObjQuickResult.ColumnCount(iCounter) > 0 Then
								Set objLink = ObjQuickResult.ChildItem(iCounter, 2, "Link", 0)
								If TypeName(objLink) <> "Nothing" Then
											If Trim(objLink.GetROProperty("text")) = "Name" Then
													iColPos = iCounter ' For Column Number
													Exit For
											End If
								End If
								Set objLink = Nothing
							End If
						Next
						iOuterCnt = 0
						iRwCount = ObjQuickResult.GetROProperty("rows")
						iCounter = iColPos+1
						iRwCount2 = ObjQuickResult.GetCellData(iCounter, 2)
						If iRwCount2 <> "" Then
								Fn_Web_QuickSearchReasultOperations = iRwCount2
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: First Element ["+CStr(iRwCount2)+"] Present in the Find Table ")
						Else
								Fn_Web_QuickSearchReasultOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: No First Element Found in Find Table ")
						End If
                Case "GetRowIndex"  'Fn_Web_QuickSearchReasultOperations("GetRowIndex","000021-Item1")
						iRwCount=ObjQuickResult.RowCount()
						For iCounter=1 To iRwCount
							currCellValue=ObjQuickResult.GetCellData(iCounter,2)
							If Trim(currCellValue)=Trim(strNode) Then
								Fn_Web_QuickSearchReasultOperations=iCounter
								Exit For
							End If
						Next
				Case "ClickLink"
						iRowNum=Fn_Web_QuickSearchReasultOperations("GetRowIndex",strNode)
						If iRowNum<>"" Then
								Set ObjChk=ObjQuickResult.ChildItem(CInt(iRowNum),2,"Link", 0)
								If TypeName(ObjChk) <> "Nothing" Then
									ObjChk.Click 1, 1, micLeftBtn
									wait(2)
									Fn_Web_QuickSearchReasultOperations=True
								End If
						End If
		End Select
	End If
	Set ObjQuickResult=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_QuickLinksMenuOperations
'@@
'@@    Description				 :	Function Used to Perform Operations On Quick Links of Ride Side Pannel
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@												 2.StrLinkName : Link Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_QuickLinksMenuOperations("Click","Paste")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									22-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_QuickLinksMenuOperations(StrAction,StrLinkName)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_QuickLinksMenuOperations"
	 Dim ObjLinks
	 Fn_Web_QuickLinksMenuOperations=False
	 If Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("QuickLinks").Exist(2) Then
		Set ObjLinks=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("QuickLinks")
	Else
		Set ObjLinks=Browser("TeamcenterWeb").Page("MyTeamCenter")
	 End If
    Select Case StrAction
		Case "Click"
            ObjLinks.WebButton("QuickButton").SetTOProperty "Name",StrLinkName
			wait 3
			If ObjLinks.WebButton("QuickButton").Exist(6) Then
				ObjLinks.WebButton("QuickButton").Click
				Fn_Web_QuickLinksMenuOperations=True
			End If
			
    End Select
	Set ObjLinks=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_DeleteObject
'@@
'@@    Description				 :	Function Used To Delete Objects
'@@
'@@   Parameter		   	   : 	bDeleteAllSeq : Delete All Sequence Option
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_DeleteObject("")
'@@												Call Fn_Web_DeleteObject("On")
'@@												
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									24-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_DeleteObject(bDeleteAllSeq)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_DeleteObject"
   Dim ObjDelete
   Dim strWEBMenuPath,strMenu,bFlag,iCounter
   bFlag=False
   Fn_Web_DeleteObject=False
	Set ObjDelete=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Delete")
	If Not ObjDelete.Exist(10) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "EditDelete")
		Call Fn_Web_MenuOperation("Select",strMenu)
		wait 2
	End If
	For iCounter=0 to 2
		If ObjDelete.Exist(4) Then
			wait 2
			bFlag=True
			Exit for
		End if
	Next
	If bFlag=False Then
		Set ObjDelete=Nothing
		Exit function
	End If

	If bDeleteAllSeq<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_Web_DeleteObject", ObjDelete, "DeleteAllSeq", bDeleteAllSeq)
	End If
	'Clicking On OK Button
'	bFlag=Fn_Web_UI_Button_Click("Fn_Web_DeleteObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"OK")
	bFlag=Fn_Web_UI_Button_Click("Fn_Web_DeleteObject",Browser("TeamcenterWeb").Page("MyTeamCenter"),"OK")
	If bFlag=True Then
		Fn_Web_DeleteObject=True
	End If

	For iCounter=0 To 2
		If ObjDelete.Exist(5) Then
			wait(5)
		Else
			Exit For
		End If
	Next

	Set ObjDelete=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ColumnManagement
'@@
'@@    Description				 :	Function Used to Manage Object Columns
'@@
'@@    Parameters			   :	1.strAction: Action Name
'@@												 2.strObjectType : Object Type
'@@												 3.strType : Type
'@@												 4.strShowColName : Show Column Names
'@@												 5.strHideColName : Hide Column Names
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_ColumnManagement("Add","Item Columns","Item","","Alias IDs:Contacts")
'@@												Call Fn_Web_ColumnManagement("Remove","Item Columns","Item","Alias IDs:Contacts","")
'@@												Call Fn_Web_ColumnManagement("MoveUpFromShowList","Folder Columns","Envelope","Type:1","")
'@@												Call Fn_Web_ColumnManagement("MoveDownFromShowList","Folder Columns","Envelope","Type:2","")
'@@												Call Fn_Web_ColumnManagement("MoveUpFromHideList","Structure Manager Columns","","","APN UID:2")
'@@												Call Fn_Web_ColumnManagement("MoveDownFromHideList","Structure Manager Columns","","","APN UID:3")
'@@												Call Fn_Web_ColumnManagement("VerifyItemsFromShowList","Folder Columns","","Name:Type","")
'@@												Call Fn_Web_ColumnManagement("VerifyItemsFromHideList","Folder Columns","","","Name:Type")
'																																			"TabName~Link Name"
'@@												Call Fn_Web_ColumnManagement("Add","Display~Shown Item Relations","Item","","Alias IDs:Contacts")
'@@												Call Fn_Web_ColumnManagement("VerifyFromShowListAndAdd","Item Columns","Item","","MFK Key1:MFK Key2")
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									24-Apr-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									26-Apr-2011						1.1						Case "VerifyItemsFromShowList"					""
'@@												Sandeep Navghane									26-Apr-2011						1.1						Case "VerifyItemsFromHideList"					  "'
'@@												Sandeep Navghane									13-Oct-2011						1.3						Added Code to handle Column from other Tabs					  "'
'@@												Sandeep Navghane									10-Feb-2012						1.4						Added Case "VerifyFromShowListAndAdd"		Swati K
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ColumnManagement(strAction,strObjectType,strType,strShowColName,strHideColName)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ColumnManagement"
	'Variable Declaration
	Dim ObjOptions,ObjColMngmnt,ObjColPref,arrTab
	Dim strMenu,CurrType,arrHideCols,arrShowCols,iCounter
	Dim iItemsCount,iCount,bFlag,Curritem,objMyTcPage
	Fn_Web_ColumnManagement=False
	
	Set objMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	'Creating Object Of "Options" , "ColumnManagement" , "ColumnsPreference" Table
'	Set ObjOptions=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Options")
'	Set ObjColMngmnt=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Options").WebTable("ColumnManagement")
'	Set ObjColPref=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ColumnsPreference")
	
	Set ObjOptions=objMyTcPage.WebTable("Options")
	Set ObjColMngmnt=objMyTcPage.WebTable("Options").WebTable("ColumnManagement")
	Set ObjColPref=objMyTcPage.WebTable("ColumnsPreference")
	
	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "EditOptions")
	'Checking Existance of Options Dialog
	If Not ObjOptions.Exist(SISW_MIN_TIMEOUT) Then
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(1)
	End If
	
	arrTab=Split(strObjectType,"~")
	If UBound(arrTab)=1 Then
		'Clicking On "Column Management" Tab
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",ObjOptions.WebElement("TabName"),"innertext",arrTab(0))
		Call Fn_Web_UI_WebElement_Click("Fn_Web_ColumnManagement", ObjOptions,"TabName", "","","")
		'Selecting Object Type
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",ObjOptions.Link("ObjectType"),"text",arrTab(1))
		Call Fn_Web_UI_Link_Click("Fn_Web_ColumnManagement",ObjOptions, "ObjectType","","","")
	Else
		'Clicking On "Column Management" Tab
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",ObjOptions.WebElement("TabName"),"innertext","Column Management")
		Call Fn_Web_UI_WebElement_Click("Fn_Web_ColumnManagement", ObjOptions,"TabName", "","","")
		'Selecting Object Type
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",ObjColMngmnt.Link("ObjectType"),"text",strObjectType)
		Call Fn_Web_UI_Link_Click("Fn_Web_ColumnManagement",ObjColMngmnt, "ObjectType","","","")
	End If
	Wait WEB_MIN_TIMEOUT

'	If Not ObjColPref.Exist(15) Then
	If Not Fn_Web_UI_ObjectExist("Fn_Web_ColumnManagement",ObjColPref) Then
		Set ObjOptions=Nothing
		Set ObjColMngmnt=Nothing
		Set ObjColPref=Nothing
		Exit Function
	End If
	If strType<>"" Then
'		CurrType=ObjColPref.WebEdit("Type").GetROProperty("value")
		CurrType=Fn_WEB_UI_Object_GetROProperty("Fn_Web_ColumnManagement",ObjColPref.WebEdit("Type"),"value")
		If Trim(CurrType)<>Trim(strType) Then
			Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", ObjColPref, "Type")
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",strType)
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",objMyTcPage.WebElement("FormType"),"innertext",strType)
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
			Call Fn_Web_UI_WebElement_Click("Fn_Web_ColumnManagement", objMyTcPage, "FormType", 1,1,micLeftBtn)
			Wait(WEB_MICROLESS_TIMEOUT)
		End If
	End If
	Select Case strAction
		Case "Add"	'Case to Add Columns From Hide List
			arrHideCols=Split(strHideColName,":")
			For iCounter=0 To UBound(arrHideCols)
				Call Fn_Web_UI_List_Select("Fn_Web_ColumnManagement",ObjColPref, "Hide",arrHideCols(iCounter))
				Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", ObjColPref, "Add")
			Next
			Fn_Web_ColumnManagement=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyFromShowListAndAdd"	'Case to Verify whether properties already exists in Show list and then Add them
            		arrShowCols=Split(strHideColName,":")
			For iCounter=0 To UBound(arrShowCols)
				bFlag=False
'				iItemsCount=ObjColPref.WebList("Show").GetROProperty("items count")
				iItemsCount=Fn_WEB_UI_Object_GetROProperty("Fn_Web_ColumnManagement",ObjColPref.WebList("Show"),"items count")
				For iCount=1 To iItemsCount
					Curritem=ObjColPref.WebList("Show").GetItem(iCount)
					If Trim(Curritem)=arrShowCols(iCounter) Then
						bFlag=True
						Exit For
					End If
				 Next
				If bFlag=False Then
					Call Fn_Web_UI_List_Select("Fn_Web_ColumnManagement",ObjColPref, "Hide",arrShowCols(iCounter))
					Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", ObjColPref, "Add")
					bFlag=True
				End If
			Next
			If bFlag=True Then
				Fn_Web_ColumnManagement=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "Remove"	'Case to Remove Columns From Show List
			arrShowCols=Split(strShowColName,":")
			For iCounter=0 To UBound(arrShowCols)
				Call Fn_Web_UI_List_Select("Fn_Web_ColumnManagement",ObjColPref, "Show",arrShowCols(iCounter))
				Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", ObjColPref, "Remove")
			Next
			Fn_Web_ColumnManagement=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "MoveUpFromShowList"	'To Move Up Column From Show List
			arrShowCols=Split(strShowColName,":")
			If Ubound(arrShowCols)=0 Then
				iCount=1
			Else
				iCount=arrShowCols(1)
			End If
			For iCounter=1 To iCount
				Call Fn_Web_UI_List_Select("Fn_Web_ColumnManagement",ObjColPref, "Show",arrShowCols(0))
				Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", ObjColPref, "MoveUp")
			Next
			Fn_Web_ColumnManagement=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "MoveDownFromShowList"		'To Move Down Column From Show List
			arrShowCols=Split(strShowColName,":")
			If Ubound(arrShowCols)=0 Then
				iCount=1
			Else
				iCount=arrShowCols(1)
			End If
			For iCounter=1 To iCount
				Call Fn_Web_UI_List_Select("Fn_Web_ColumnManagement",ObjColPref, "Show",arrShowCols(0))
				Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", ObjColPref, "MoveDown")
			Next
			Fn_Web_ColumnManagement=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "MoveUpFromHideList"	'To Move Up Column From Hide List
			arrHideCols=Split(strHideColName,":")
			If Ubound(arrHideCols)=0 Then
				iCount=1
			Else
				iCount=arrHideCols(1)
			End If
			For iCounter=1 To iCount
				Call Fn_Web_UI_List_Select("Fn_Web_ColumnManagement",ObjColPref, "Hide",arrHideCols(0))
				Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", ObjColPref, "MoveUp")
			Next
			Fn_Web_ColumnManagement=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "MoveDownFromHideList"		'To Move Down Column From Hide List
			arrHideCols=Split(strHideColName,":")
			If Ubound(arrHideCols)=0 Then
				iCount=1
			Else
				iCount=arrHideCols(1)
			End If
			For iCounter=1 To iCount
				Call Fn_Web_UI_List_Select("Fn_Web_ColumnManagement",ObjColPref, "Hide",arrHideCols(0))
				Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", ObjColPref, "MoveDown")
			Next
			Fn_Web_ColumnManagement=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyItemsFromShowList"
			arrShowCols=Split(strShowColName,":")
'			iItemsCount=ObjColPref.WebList("Show").GetROProperty("items count")
			iItemsCount=Fn_WEB_UI_Object_GetROProperty("Fn_Web_ColumnManagement",ObjColPref.WebList("Show"),"items count")
			For iCount=0 To UBound(arrShowCols)
				bFlag=False
				For iCounter=1 To iItemsCount
					Curritem=ObjColPref.WebList("Show").GetItem(iCounter)
					If Trim(Curritem)=arrShowCols(iCount) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_Web_ColumnManagement=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "VerifyItemsFromHideList"
			arrHideCols=Split(strHideColName,":")
'			iItemsCount=ObjColPref.WebList("Hide").GetROProperty("items count")
			iItemsCount=Fn_WEB_UI_Object_GetROProperty("Fn_Web_ColumnManagement",ObjColPref.WebList("Hide"),"items count")
			For iCount=0 To UBound(arrHideCols)
				bFlag=False
				For iCounter=1 To iItemsCount
					Curritem=ObjColPref.WebList("Hide").GetItem(iCounter)
					If Trim(Curritem)=arrHideCols(iCount) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_Web_ColumnManagement=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "RestoreDefaults"
		  'Case to Restore Default Column Values
'		   Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("LoadAll"),"name","Restore Defaults")
'		   Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", Browser("TeamcenterWeb").Page("MyTeamCenter"), "LoadAll")
'		   Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
		   
		   Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",objMyTcPage.WebButton("LoadAll"),"name","Restore Defaults")
		   Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", objMyTcPage, "LoadAll")
		   Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", objMyTcPage.WebElement("ButtunPanel"), "Close")
		   Fn_Web_ColumnManagement=True
	End Select

	 If strAction <> "RestoreDefaults" Then
'		If Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("OK").Exist(5) Then
'			Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("OK").Click 1,1,micLeftBtn
'			Wait(5)
'			Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
'		Else
		If objMyTcPage.WebButton("OK").Exist(2) Then
			objMyTcPage.WebButton("OK").Click 1,1,micLeftBtn
			Wait(WEB_MICROLESS_TIMEOUT)
			Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", objMyTcPage.WebElement("ButtunPanel"), "Close")
		Else
'			  Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("LoadAll"),"name","OK")
'			  Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", Browser("TeamcenterWeb").Page("MyTeamCenter"), "LoadAll")
'			  Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
			   Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ColumnManagement",objMyTcPage.WebButton("LoadAll"),"name","OK")
			  Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", objMyTcPage, "LoadAll")
			  Call Fn_Web_UI_Button_Click("Fn_Web_ColumnManagement", objMyTcPage.WebElement("ButtunPanel"), "Close")
		End If
	 End If

	Set objMyTcPage = Nothing
	Set ObjOptions=Nothing
	Set ObjColMngmnt=Nothing
	Set ObjColPref=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_DetailsTableOperations
'@@
'@@    Description				 :	Function Used to Perform Operations On Details Table
'@@
'@@    Parameters			   :	1.strAction: Action Name
'@@												 2.strObjectName : Object name
'@@												 3.strColName : Column Name
'@@												 4.strValue : Cell Value
'@@
'@@    Return Value		   	   : 	True Or False Or Column Names
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_DetailsTableOperations("GetAllColumnNames","","","")
'@@												Call Fn_Web_DetailsTableOperations("Select","000017","","")
'@@												Call Fn_Web_DetailsTableOperations("GetCellData","REQ-123458-Req3","Checked-Out","")
'@@
'@@	   History					 	:	
'@@			Developer Name			Date		Rev. No.	Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Sandeep Navghane	25-Apr-2011		  1.0													Sunny Ruparel
'@@			Sandeep Navghane	17-Oct-2011		  1.1	
'@@			Pritam Shikare		24-May-2013       1.2       Added case "VerifyRowData"  				Sunny Ruparel
'@@			Vivek Ahirrao		27-Oct-2015       1.2       Added case "GetAllColumnNamesExt"  			[TC1121-2015101200-27_10_2015-VivekA-Maintenance]	
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_DetailsTableOperations(strAction,strObjectName,strColName,strValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_DetailsTableOperations"
	'Variable Declaration
	Dim ObjDetailsTB,ObjChk,bFlag, aColumns, aValues, sCellData
	Dim iColCount,iCounter,ColName,arrColName,iRwCount,iRowNum,currCellValue
    'Creating Object Of Details Table
'	Set ObjDetailsTB=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("DetailsTabTable").WebTable("DetailsTable")
	Set ObjDetailsTB=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("DetailsTable")
	'Clicking On Details Tab
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_DetailsTableOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Overview"),"innertext","Details")
	Call Fn_Web_UI_WebElement_Click("Fn_Web_DetailsTableOperations",Browser("TeamcenterWeb").Page("MyTeamCenter"),"Overview", "","","")
	Select Case strAction
		Case "GetAllColumnNames" 'Case returns All Column Name of Table
				iColCount=ObjDetailsTB.ColumnCount(1)
				For iCounter=2 To iColCount
						ColName=ObjDetailsTB.GetCellData(1,iCounter)
						If iCounter=2 Then
							arrColName=ColName
						Else
							arrColName=arrColName+":"+ColName
						End If
				Next
				Fn_Web_DetailsTableOperations=arrColName
		'IE11 does not support above table method ColumnCount(1). So added new case [TC11.2.1-2015101200-21_10_2015-VivekA-Maintenance]
		Case "GetAllColumnNamesExt" 'Case returns All Column Name of Table
				If ObjDetailsTB.Exist(1) Then
					iColCount = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("DetailsTableHeader").ColumnCount(1)
					For iCounter=2 To iColCount
						ColName=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("DetailsTableHeader").GetCellData(1,iCounter)
						If iCounter=2 Then
							arrColName=ColName
						Else
							arrColName=arrColName+":"+ColName
						End If
					Next
					Fn_Web_DetailsTableOperations=arrColName
				Else
					Fn_Web_DetailsTableOperations = False
					Set ObjDetailsTB = Nothing
					Exit Function
				End If
		'Case to Retrieve Column Index
		 Case "GetColIndex" 
				'------------------------------------------------------------------------------
				'------------------------------------------------------------------------------
				Fn_Web_DetailsTableOperations = -1
				iColCount = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("DetailsTableHeader").ColumnCount(1)
				For iCounter=1 To iColCount
					currCellValue=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("DetailsTableHeader").GetCellData(1,iCounter)
					If Trim(currCellValue)=Trim(strColName) Then
						Fn_Web_DetailsTableOperations = iCounter
						Exit for
					End If
				Next
				Set objDetailsTB = Nothing
		'Case to Retrieve Row Index
		 Case "GetRowIndex" 
				iRwCount=ObjDetailsTB.RowCount()
				For iCounter=1 To iRwCount
					currCellValue=ObjDetailsTB.GetCellData(iCounter,2)
					If Trim(currCellValue)=Trim(strObjectName) Then
						Fn_Web_DetailsTableOperations=iCounter
						Exit For
					End If
				Next
		'Case To Select Item
		 	Case "Select"
				iRowNum=Fn_Web_DetailsTableOperations("GetRowIndex",strObjectName,"","")
				If iRowNum<>"" Then
						Set ObjChk=ObjDetailsTB.ChildItem(CInt(iRowNum), 1,"WebElement", 0)
						If TypeName(ObjChk) <> "Nothing" Then
								ObjChk.Click 1, 1, micLeftBtn
								Fn_Web_DetailsTableOperations=True
						End If
						Set ObjChk=Nothing
				End If
			
			'Case to get Cell data
			Case "GetCellData"
				iColCount= Fn_Web_DetailsTableOperations("GetColIndex","",strColName,"")
					If iColCount <> -1 Then
						iRwCount=ObjDetailsTB.RowCount()
						For iCounter=1 To iRwCount
							currCellValue=ObjDetailsTB.GetCellData(iCounter,2)
							If Trim(currCellValue)=Trim(strObjectName) Then
								Fn_Web_DetailsTableOperations = ObjDetailsTB.GetCellData(iCounter,iColCount)
								Exit For
							End If
						Next
					End If
			
'					bFlag=False
'					iColCount=ObjDetailsTB.ColumnCount(1)
'					For iCounter=1 To iColCount
'						currCellValue=ObjDetailsTB.GetCellData(1,iCounter)
'						If Trim(currCellValue)=Trim(strColName) Then
'							iColCount=iCounter
'							bFlag=True
'							Exit For
'						End If
'					Next
'					If bFlag=True Then
'						iRwCount=ObjDetailsTB.RowCount()
'						For iCounter=1 To iRwCount
'							currCellValue=ObjDetailsTB.GetCellData(iCounter,2)
'							If Trim(currCellValue)=Trim(strObjectName) Then
'								Fn_Web_DetailsTableOperations= ObjDetailsTB.GetCellData(iCounter,iColCount)
'								Exit For
'							End If
'						Next
'					End If
				'Case to get Cell data
				Case "VerifyRowData"
					bFlag = True
					If strColName = "" Then
						strColName = "Object"
						strValue = strObjectName
					End If

					aColumns = Split(strColName,"~",-1,1)
					aValues = Split(strValue,"~",-1,1)

					For iCounter = 0 to Ubound(aColumns)
						sCellData =  Fn_Web_DetailsTableOperations("GetCellData",strObjectName,aColumns(iCounter),"")
						If  sCellData <> aValues(iCounter) Then
							bFlag = False
							Exit For
						End If
					Next
					If bFlag = True Then
						Fn_Web_DetailsTableOperations = True
					Else
						Fn_Web_DetailsTableOperations = False
					End If

	End Select
	'Releasing Object Of Details Table
	Set ObjDetailsTB=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ItemDetailsCreate
'@@
'@@    Description				 :	Function Used to Create Item Detail Information
'@@
'@@    Parameters			   :	1.dicItemDetailsCreate : Item Detail Information Dic Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	
'@@												dicItemDetailsCreate("Type")="Item"
'@@												dicItemDetailsCreate("Revision")="A"
'@@												dicItemDetailsCreate("Name")="TestItem"
'@@												dicItemDetailsCreate("Description")="Tes Details tItem"
'@@												dicItemDetailsCreate("CreateAlternateID")="Off"
'@@												dicItemDetailsCreate("CheckOnRevision") = "Off"
'@@												dicItemDetailsCreate("ProjectID")="12345"
'@@												dicItemDetailsCreate("PreviousID")="22"
'@@												dicItemDetailsCreate("SerialNumber")="9999"
'@@												dicItemDetailsCreate("ItemComment")="DetailedItem"
'@@												dicItemDetailsCreate("WA11int")="12345"
'@@												Call Fn_Web_ItemDetailsCreate(dicItemDetailsCreate)
'@@												dicItemDetailsCreate("MFKkey1")="12345"		
'@@												dicItemDetailsCreate("MFKkey2")="12345"		
'@@	
'@@	   History					 	:	
'@@				Developer Name									Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Sandeep Navghane								8-Apr-2011						1.0																								Sunny Ruparel
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Koustubh Watwe									29-Dec-2011						1.0
'@@				Avinash Jagdale								    18-jan-2012					1.0
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@				Swati Kuntullu									07-Feb-2012						1.0						Modifeid code to set values					Koustubh Watwe		
'@@			   Sukhada Bakshi							 22 -Oct-2012						1.0						Added Code for  "OriginalCageCode"				
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ItemDetailsCreate(dicItemDetailsCreate)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ItemDetailsCreate"
   'Varibale Declaration
   Dim ObjItem,strWEBMenuPath,strMenu,crrType, WshShell
   Dim dicItems, sID, sRev
   Fn_Web_ItemDetailsCreate=False
   Set WshShell = CreateObject("WScript.Shell")
   'Creating Object Of "NewItem" Table
   Set ObjItem=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem")
	'Checking Existance Of "New Item" Dialog
	If  Not ObjItem.Exist(7) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewItem")
		Call Fn_Web_MenuOperation("Select",strMenu)
		wait 2
	End If
	'Taking All Values From Dictionary
	dicItems=dicItemDetailsCreate.Items
	'Selecting Item Type
	If dicItems(0)<>"" Then
		crrType=ObjItem.WebEdit("ItemTypeEdit").GetROProperty("value")
		wait(1)
		If Trim(crrType)<>Trim(dicItems(0)) Then
				'Setting Item Type
				wait 2
				Call Fn_Web_UI_Button_Click("Fn_Web_ItemDetailsCreate",ObjItem,"ItemType")
				wait 2
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ItemDetailsCreate",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",dicItems(0))
				wait 1
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
				wait(2)
		End If
	End If

	If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Next").GetROProperty("disabled") = "0" Then
		Call Fn_Web_UI_Button_Click("Fn_Web_ItemDetailsCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
		wait 1
	End If
    If dicItemDetailsCreate("IDNamingRule")<>"" then
		ObjItem.WebTable("ItemInfo").WebList("IDNamingRule").dicItemDetailsCreate("IDNamingRule")
		Wait 1
		If trim(ObjItem.WebTable("ItemInfo").WebList("IDNamingRule").GetROProperty("value"))<>trim(dicItemDetailsCreate("IDNamingRule")) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail to select Naming rule " & dicItemDetailsCreate("IDNamingRule") &" for ID")
			Exit function
		End If
	End if
	If dicItemDetailsCreate("ID") <>"" Then
		If  dicItemDetailsCreate("ID") = "ASSIGN" Then
			Call Fn_Web_UI_Button_Click("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "AssignID")
			Wait 4
			sID = Fn_Web_UI_WebEdit_GetValue ("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "ID")
		Else
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "ID", dicItemDetailsCreate("ID"))
			wait 2
		End If
	End If
	If dicItemDetailsCreate("Revision")<>"" Then
		If  dicItemDetailsCreate("Revision") = "ASSIGN" Then
			Call Fn_Web_UI_Button_Click("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "AssignRevID")
			Wait 4
			sRev = Fn_Web_UI_WebEdit_GetValue ("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "Revision")
		Else
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "Revision", dicItemDetailsCreate("Revision"))
			wait 1
		End If
	End If
	If dicItemDetailsCreate("Name")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "Name", dicItemDetailsCreate("Name"))
		wait 2
	End If
	If dicItemDetailsCreate("Description")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "Description", dicItemDetailsCreate("Description"))
		wait 2
	End If
	If dicItemDetailsCreate("UOM")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "UOM", dicItemDetailsCreate("UOM"))
		wait 1
	End If
	If dicItemDetailsCreate("CreateAlternateID") <>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "CreateALTID",dicItemDetailsCreate("CreateAlternateID"))
		wait 1
	End If
	If dicItemDetailsCreate("CheckOnRevision") <> "" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemInfo"), "CheckOutRevision", dicItemDetailsCreate("CheckOnRevision"))
		wait 1
	End If
	If dicItemDetailsCreate("FinishItems") <> "" Then
		ObjItem.WebTable("ItemInfo").Image("Paste").Click
		Wait 1
	End If

	If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Next").GetROProperty("disabled") = "0" Then
		Call Fn_Web_UI_Button_Click("Fn_Web_ItemDetailsCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
		wait 1
	End If

	If dicItemDetailsCreate("WA11int")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfoWA11Int"), "WA11int", dicItemDetailsCreate("WA11int"))
	End If
	If dicItemDetailsCreate("MFKkey1")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "MFKkey1", dicItemDetailsCreate("MFKkey1"))
	End If
	If dicItemDetailsCreate("MFKkey2")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "MFKkey2", dicItemDetailsCreate("MFKkey2"))
	End If
	If dicItemDetailsCreate("ProjectID") <>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "ProjectID", dicItemDetailsCreate("ProjectID"))
	End If
	If dicItemDetailsCreate("PreviousID") <> "" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "PreviousID", dicItemDetailsCreate("PreviousID"))
	End If
	If dicItemDetailsCreate("SerialNumber")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "SerialNumber", dicItemDetailsCreate("SerialNumber"))
	End If
	If dicItemDetailsCreate("ItemComment") <>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "ItemComment", dicItemDetailsCreate("ItemComment"))
	End If
	If dicItemDetailsCreate("UserData1") <>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "UserData1",dicItemDetailsCreate("UserData1"))
	End If
	If dicItemDetailsCreate("UserData2")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "UserData2", dicItemDetailsCreate("UserData2"))
	End If
	If dicItemDetailsCreate("UserData3")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "UserData3", dicItemDetailsCreate("UserData3"))
	End If
     If dicItemDetailsCreate("OriginalCageCode")<>"" Then
		wait 1
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "OriginalCageCode", dicItemDetailsCreate("OriginalCageCode"))
	End If
    If dicItemDetailsCreate("NoteCategory")<>"" Then
		wait 1
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "NoteCategory", dicItemDetailsCreate("NoteCategory"))
	End if
	
	If dicItemDetailsCreate("Category")<>"" Then
		ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Category*:"
		If ObjItem.WebTable("AddtionalPartTechDocInfo").WebButton("ItemType").Exist(5) Then
			ObjItem.WebTable("AddtionalPartTechDocInfo").WebButton("ItemType").Click
			wait(2)
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").SetTOProperty "innertext",dicItemDetailsCreate("Category")
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click
			wait(2)
		Else
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("Category"))
			wait(2)
			WshShell.SendKeys "{ENTER}"
			wait(2)
		End If
	End If

	If dicItemDetailsCreate("TecDocCategory")<>"" Then
		ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Technical Document Category:"
		If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait 2
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("TecDocCategory"))
		Else
			ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Technical Document Category*:"
			If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait 2
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("TecDocCategory"))
			End If
		End If
	End If

	If dicItemDetailsCreate("PartCategory")<>"" Then
		ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Part Category:"
		If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait 2
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("PartCategory"))
		else
			ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Part Category*:"
			If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
				wait 2
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("PartCategory"))
			End If
		End If
	End If

	If dicItemDetailsCreate("SrcDocID")<>"" Then
		ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Source Document ID:"
		If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait 2
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("SrcDocID"))
		else
			ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Source Document ID*:"
			If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
				wait 2
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("SrcDocID"))
			End If
		End If
	End If
	
	If dicItemDetailsCreate("ContractCategory")<>"" Then
		ObjItem.WebTable("AdditionalItemInfo").WebElement("Label").SetTOProperty "innertext","Contract Category:"
		If ObjItem.WebTable("AdditionalItemInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait(2)
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "EditAddnInfo", dicItemDetailsCreate("ContractCategory"))
		else
			ObjItem.WebTable("AdditionalItemInfo").WebElement("Label").SetTOProperty "innertext","Contract Category*:"
			If ObjItem.WebTable("AdditionalItemInfo").WebEdit("EditAddnInfo").Exist(5) Then
				wait(2)
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "EditAddnInfo", dicItemDetailsCreate("ContractCategory"))
			End If
		End If
		WshShell.SendKeys "{ENTER}"
		wait(2)
	End If

	If dicItemDetailsCreate("DesignCategory")<>"" Then
		ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Design Category:"
		If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait(2)
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("DesignCategory"))
		else
			ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Design Category*:"
			If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
				wait(2)
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("DesignCategory"))
			End If
		End If
		if ObjItem.WebTable("AddtionalPartTechDocInfo").WebButton("AddnInfo_Button").Exist(1) then
			WshShell.SendKeys "{ENTER}"
			wait(2)
		End if
	End If

	If dicItemDetailsCreate("WorkPkgComplexity")<>"" Then
		ObjItem.WebTable("AdditionalItemInfo").WebElement("Label").SetTOProperty "innertext","Work Package Complexity:"
		If ObjItem.WebTable("AdditionalItemInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait(2)
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "EditAddnInfo", dicItemDetailsCreate("WorkPkgComplexity"))
		else
			ObjItem.WebTable("AdditionalItemInfo").WebElement("Label").SetTOProperty "innertext","Work Package Complexity*:"
			If ObjItem.WebTable("AdditionalItemInfo").WebEdit("EditAddnInfo").Exist(5) Then
				wait(2)
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "EditAddnInfo", dicItemDetailsCreate("WorkPkgComplexity"))
			End If
		End If
	    if ObjItem.WebTable("AddtionalPartTechDocInfo").WebButton("AddnInfo_Button").Exist(1) then
			WshShell.SendKeys "{ENTER}"
			wait(2)
		End if
	End If

	If dicItemDetailsCreate("WorkPkgSecurity")<>"" Then
		ObjItem.WebTable("AdditionalItemInfo").WebElement("Label").SetTOProperty "innertext","Work Package Security:"
		If ObjItem.WebTable("AdditionalItemInfo").WebButton("ItemType").Exist(5) Then
			ObjItem.WebTable("AdditionalItemInfo").WebButton("ItemType").Click
			wait(2)
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").SetTOProperty "innertext",dicItemDetailsCreate("WorkPkgSecurity")
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click
			wait(2)
		Else
			wait(2)
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "EditAddnInfo", dicItemDetailsCreate("WorkPkgSecurity"))
		End If
		if ObjItem.WebTable("AddtionalPartTechDocInfo").WebButton("AddnInfo_Button").Exist(1) then
			WshShell.SendKeys "{ENTER}"
			wait(2)
		End if
	End If

	If dicItemDetailsCreate("WorkPkgType")<>"" Then
		ObjItem.WebTable("AdditionalItemInfo").WebElement("Label").SetTOProperty "innertext","Work Package Type:"
		If ObjItem.WebTable("AdditionalItemInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait 2
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "EditAddnInfo", dicItemDetailsCreate("WorkPkgType"))
		else
			ObjItem.WebTable("AdditionalItemInfo").WebElement("Label").SetTOProperty "innertext","Work Package Type*:"
			If ObjItem.WebTable("AdditionalItemInfo").WebEdit("EditAddnInfo").Exist(5) Then
				wait 2
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AdditionalItemInfo"), "EditAddnInfo", dicItemDetailsCreate("WorkPkgType"))
			End If
		End If
		WshShell.SendKeys "{ENTER}"
	End If
	
	If dicItemDetailsCreate("SrcDocCategory")<>"" Then
		ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Source Document Category:"
		If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait(2)
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("SrcDocCategory"))
		Else
			ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Source Document Category*:"
			If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
				wait(2)
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("SrcDocCategory"))
			End If
		End If
		ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("ListAddnInfo").SetTOProperty "innertext",dicItemDetailsCreate("SrcDocCategory")
		If ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("ListAddnInfo").Exist(5) Then
			call Fn_Web_UI_WebElement_Click("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "ListAddnInfo", "","","")
		End If
'        if ObjItem.WebTable("AddtionalPartTechDocInfo").WebButton("AddnInfo_Button").Exist(1) then
'			WshShell.SendKeys "{ENTER}"
'		End if
	End If
    Wait 5
	
	If dicItemDetailsCreate("SrcTechDocCategory")<>"" Then
		ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Source Technical Document Category:"
		If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
			wait(2)
			Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("SrcTechDocCategory"))
		Else
			ObjItem.WebTable("AddtionalPartTechDocInfo").WebElement("Label").SetTOProperty "innertext","Source Technical Document Category*:"
			If ObjItem.WebTable("AddtionalPartTechDocInfo").WebEdit("EditAddnInfo").Exist(5) Then
				wait(2)
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("AddtionalPartTechDocInfo"), "EditAddnInfo", dicItemDetailsCreate("SrcTechDocCategory"))
			End If
		End If
	End If
	
	If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Next").GetROProperty("disabled") = "0" Then
		Call Fn_Web_UI_Button_Click("Fn_Web_ItemDetailsCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	End If

	If dicItemDetailsCreate("RevProjectID")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemRevInfo"), "RevProjectID", dicItemDetailsCreate("RevProjectID"))
	End If
	If dicItemDetailsCreate("PreviousVersionID")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemRevInfo"), "PreviousVersionID", dicItemDetailsCreate("PreviousVersionID"))
	End If
	If dicItemDetailsCreate("RevSerialNumber")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemRevInfo"), "RevSerialNumber", dicItemDetailsCreate("RevSerialNumber"))
	End If
	If dicItemDetailsCreate("RevItemComment")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemRevInfo"), "RevItemComment", dicItemDetailsCreate("RevItemComment"))
	End If
	If dicItemDetailsCreate("RevUserData1")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemRevInfo"), "RevUserData1", dicItemDetailsCreate("RevUserData1"))
	End If
	If dicItemDetailsCreate("RevUserData2")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemRevInfo"), "RevUserData2", dicItemDetailsCreate("RevUserData2"))
	End If
	If dicItemDetailsCreate("RevUserData3")<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_ItemDetailsCreate", ObjItem.WebTable("ItemRevInfo"), "RevUserData3", dicItemDetailsCreate("RevUserData3"))
	End If
	
	If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Next").GetROProperty("disabled") = "0" Then
		Call Fn_Web_UI_Button_Click("Fn_Web_ItemDetailsCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	End If

    If dicItemDetailsCreate("AvailableProjects")<>"" Then
        bReturn =  Fn_SISW_Web_NewItemAssignToProgramOperations("Add",dicItemDetailsCreate,"")
        If bReturn = False Then
            Fn_Web_ItemDetailsCreate = False
            Exit Function
        End If
     End If
	
	Call Fn_Web_UI_Button_Click("Fn_Web_ItemDetailsCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	wait(2)
	If dicItemDetailsCreate("ID") = "ASSIGN" AND dicItemDetailsCreate("Revision") = "ASSIGN" Then
		Fn_Web_ItemDetailsCreate =  sID+"-"+sRev
	ElseIf dicItemDetailsCreate("ID") = "ASSIGN" Then
		Fn_Web_ItemDetailsCreate =  sID
	ElseIf dicItemDetailsCreate("Revision") = "ASSIGN" Then
		Fn_Web_ItemDetailsCreate =  sRev
	Else
		Fn_Web_ItemDetailsCreate=True
	End If

	Set ObjItem=Nothing
	Set WshShell = Nothing	
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_Tab_Login
'@@
'@@    Description				 :	Function Used to Log In Web Client using Tab
'@@
'@@    Parameters			   :	1.strUserName : User Name
'@@											 	 2.strPassword : Password
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Environment XML should be properly filled							
'@@
'@@    Examples					:	Call Fn_Web_Tab_Login("AutoTest1","AutoTest1")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@													Sunny Ruparel												26-Apr-2011						1.0																								
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_Tab_Login(strUserName,strPassword)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_Tab_Login"
   Dim strBrowser,strServerPath,ObjBrowser,handlewin,i,WshShell,strIpUName,strIpPass

   Fn_Web_Tab_Login=False
	strBrowser=Environment.Value("WebBrowserName")
	If Browser("TeamcenterWeb").Link("Logout").Exist(5) Then
		Call Fn_Web_Logout()
		Call Fn_Web_KillProcess("")
	End If
	strServerPath = Environment.Value("TcWebServer")

	If InStr(1,strBrowser,"IE")>0 Then
		Set ObjBrowser = CreateObject("InternetExplorer.Application")
		ObjBrowser.Visible = True
		ObjBrowser.Navigate strServerPath
		handlewin = ObjBrowser.HWND
		Window("hwnd:="+CStr(handlewin)).Maximize

		If Browser("TeamcenterLogin").Page("Login").Exist(10) Then
			Browser("TeamcenterLogin").Page("Login").Object.focus()
			For i=1 to Len(strUserName)
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys Mid(strUserName,i,1)
							Set WshShell = nothing
			Next
			strIpUName = Browser("TeamcenterLogin").Page("Login").WebTable("Login").WebEdit("Username").GetROProperty("value")
			If strIpUName <> strUserName Then
				Exit Function
			End If
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "{TAB}"
			Set WshShell = nothing
			Browser("TeamcenterLogin").Page("Login").Object.focus()
			For i=1 to Len(strPassword)
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys Mid(strPassword,i,1)
							Set WshShell = nothing
			Next
			strIpPass = Browser("TeamcenterLogin").Page("Login").WebTable("Login").WebEdit("Password").GetROProperty("value")
			If strIpPass <> strPassword Then
				Exit Function
			End If
			Call Fn_Web_UI_Button_Click("Fn_Web_Tab_Login", Browser("TeamcenterLogin").Page("Login").WebTable("Login"), "Login")
			Call Fn_Web_ReadyStatusSync(1)
			If  Browser("TeamcenterWeb").Link("Logout").Exist(20) Then
				Call Fn_Web_ReadyStatusSync(2)
				Fn_Web_Tab_Login=True
			End If
		End If

	Else
		If Not Browser("version:=.*Firefox.*").Exist(5) Then
			SystemUtil.Run "firefox.exe"
		End If
		Browser("version:=.*Firefox.*").Navigate strServerPath
		If Browser("TeamcenterLogin").Page("Login").Exist(10) Then
			Browser("TeamcenterLogin").Page("Login").WebTable("Login").Click
			For i=1 to Len(strUserName)
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys Mid(strUserName,i,1)
							Set WshShell = nothing
			Next
			strIpUName = Browser("TeamcenterLogin").Page("Login").WebTable("Login").WebEdit("Username").GetROProperty("value")
			If strIpUName <> strUserName Then
				Exit Function
			End If
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "{TAB}"
			Set WshShell = nothing
			Browser("TeamcenterLogin").Page("Login").WebTable("Login").Click
			For i=1 to Len(strPassword)
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys Mid(strPassword,i,1)
							Set WshShell = nothing
			Next
			strIpPass = Browser("TeamcenterLogin").Page("Login").WebTable("Login").WebEdit("Password").GetROProperty("value")
			If strIpPass <> strPassword Then
				Exit Function
			End If
			Call Fn_Web_UI_Button_Click("Fn_Web_Tab_Login", Browser("TeamcenterLogin").Page("Login").WebTable("Login"), "Login")
			If  Browser("TeamcenterWeb").Link("Logout").Exist(20) Then
				Call Fn_Web_ReadyStatusSync(2)
				Fn_Web_Tab_Login=True
			End If
		End If

	End If
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_OverviewTabOperation
'@@
'@@    Description				 :	Function Used to Perform Operation On Overview Tab
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@												  2.dicOverviewTabContent : Dectionary [ Key - Value Pair ]
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Object Search be Selected
'@@
'@@    Examples					:	dicOverviewTabContent("ObjectName")="NewForm11548"
'@@												dicOverviewTabContent("ObjectDescription")="Modified Desc"
'@@												Call Fn_Web_OverviewTabOperation("ModifyEditBox",dicOverviewTabContent)
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									26-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_OverviewTabOperation(StrAction,dicOverviewTabContent)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_OverviewTabOperation"
	Dim dicItems,dicKeys,iCounter,bFlag:bFlag=False
	Dim objMDR, iRowIndex,ObjMyTcPage
	Set objMDR = CreateObject("Mercury.DeviceReplay")
	Fn_Web_OverviewTabOperation=False
	
	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
'	Call Fn_Web_UI_WebElement_Click("Fn_Web_OverviewTabOperation",Browser("TeamcenterWeb").Page("MyTeamCenter"),"Overview","","","")
	Call Fn_Web_UI_WebElement_Click("Fn_Web_OverviewTabOperation",ObjMyTcPage,"Overview","","","")
	Wait WEB_MICROLESS_TIMEOUT
	Select Case StrAction
		Case "ModifyEditBox"
'			If Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("CheckOutAndEdit").Exist(10) Then
'				Call Fn_Web_UI_Button_Click("Fn_Web_OverviewTabOperation", Browser("TeamcenterWeb").Page("MyTeamCenter"),"CheckOutAndEdit")
'				wait(3)
'			End If
			If Fn_Web_UI_ObjectExist("Fn_Web_OverviewTabOperation",ObjMyTcPage.WebButton("CheckOutAndEdit"))=True Then
				Call Fn_Web_UI_Button_Click("Fn_Web_OverviewTabOperation", ObjMyTcPage,"CheckOutAndEdit")
				wait(WEB_MICROLESS_TIMEOUT)
			End If
'			If Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CheckOut").Exist(5) Then
'				Call Fn_Web_CheckOutObject("","")
'			End If
			If Fn_Web_UI_ObjectExist("Fn_Web_OverviewTabOperation",ObjMyTcPage.WebTable("CheckOut"))=True Then
				Call Fn_Web_CheckOutObject("","")
			End If
			dicItems=dicOverviewTabContent.Items
			dicKeys=dicOverviewTabContent.Keys
			For iCounter=0 To dicOverviewTabContent.Count-1
				'Checking Key is Exist in Dectionary or Not
				If IsNull(dicKeys(iCounter))=False Then
					'Checking For Key value is Associated or not
					If dicItems(iCounter)<>"" Then
'						Call Fn_Web_UI_WebEdit_Set("Fn_Web_OverviewTabOperation",Browser("TeamcenterWeb").Page("MyTeamCenter"),dicKeys(iCounter),dicItems(iCounter))
						Call Fn_Web_UI_WebEdit_Set("Fn_Web_OverviewTabOperation",ObjMyTcPage,dicKeys(iCounter),dicItems(iCounter))
'						Browser("TeamcenterWeb").Page("MyTeamCenter").WebEdit(dicKeys(iCounter)).Object.focus
'						objMDR.SendString dicItems(iCounter)
					End If
				End If
			Next
'			bFlag=Fn_Web_UI_Button_Click("Fn_Web_OverviewTabOperation", Browser("TeamcenterWeb").Page("MyTeamCenter"),"SaveAndCheckIn")
			bFlag=Fn_Web_UI_Button_Click("Fn_Web_OverviewTabOperation", ObjMyTcPage,"SaveAndCheckIn")
			If bFlag=True Then
				Wait(WEB_MICROLESS_TIMEOUT)
				Fn_Web_OverviewTabOperation=True
			End If		
		Case "StandardNotesLists"
'			iRowIndex = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("StandardNotesListsPanel").WebTable("StandardNotes").GetRowWithCellText(dicOverviewTabContent("ObjectName"))
			iRowIndex = ObjMyTcPage.WebElement("StandardNotesListsPanel").WebTable("StandardNotes").GetRowWithCellText(dicOverviewTabContent("ObjectName"))
			Select Case dicOverviewTabContent("Action")
				Case "Verify"
					If iRowIndex > 0 Then
						Fn_Web_OverviewTabOperation = True
					End If
				Case "Attach", "Replace", "Cut"
'					Set obj= Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("StandardNotesListsPanel").WebTable("StandardNotes").ChildItem(iRowIndex,1,"WebElement","0")
					Set obj= ObjMyTcPage.WebElement("StandardNotesListsPanel").WebTable("StandardNotes").ChildItem(iRowIndex,1,"WebElement","0")
'					If obj.GetROProperty("innertext") = dicOverviewTabContent("ObjectName") Then
'						obj.Click
'					End If
					If Fn_WEB_UI_Object_GetROProperty("Fn_Web_OverviewTabOperation",obj,"innertext") = dicOverviewTabContent("ObjectName") Then
						obj.Click
					End If
				End Select
		Case "CustomNotesLists"
'			iRowIndex = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("CustomNotesListsPanel").WebTable("CustomNotes").GetRowWithCellText(dicOverviewTabContent("ObjectName"))
			iRowIndex = ObjMyTcPage.WebElement("CustomNotesListsPanel").WebTable("CustomNotes").GetRowWithCellText(dicOverviewTabContent("ObjectName"))
			Select Case dicOverviewTabContent("Action")
				Case "Verify"
					If iRowIndex > 0 Then
						Fn_Web_OverviewTabOperation = True
					End If
				Case "Attach","Cut"
'					Set obj= Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("CustomNotesListsPanel").WebTable("CustomNotes").ChildItem(iRowIndex,1,"WebElement","0")
					Set obj= ObjMyTcPage.WebElement("CustomNotesListsPanel").WebTable("CustomNotes").ChildItem(iRowIndex,1,"WebElement","0")
'					If obj.GetROProperty("innertext") = dicOverviewTabContent("ObjectName") Then
'						obj.Click
'					End If
					If Fn_WEB_UI_Object_GetROProperty("Fn_Web_OverviewTabOperation",obj,"innertext") = dicOverviewTabContent("ObjectName") Then
						obj.Click
					End If
			End Select
			
	End Select
	Set objMDR =Nothing
	Set ObjMyTcPage = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SaveMySearches
'@@
'@@    Description				 :	Function Used to Save Searches
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@												  2.StrFolder : Save Folder Name
'@@												  3.StrName : New Search Name
'@@												  4.bShared : Shared Option
'@@												  5.StrNewFolderName : New Folder Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Object Search be Selected
'@@
'@@    Examples					:	Call Fn_Web_SaveMySearches("New","","TestSaveSearch1","Off","")
'@@												Call Fn_Web_SaveMySearches("New","","TestSaveSearch2","Off","TestNewFolder")
'@@												Call Fn_Web_SaveMySearches("New","MyFolder","TestSaveSearch5","Off","")
'@@												Call Fn_Web_SaveMySearches("Verify","","TestSaveSearch2","","")
'@@												Call Fn_Web_SaveMySearches("Modify","","TestSaveSearch2","Off","NewTestSaveSearch2")
'@@												Call Fn_Web_SaveMySearches("Delete","","NewTestSaveSearch2","","")
'@@												Call Fn_Web_SaveMySearches("ExecuteSaveSearch","","86193Test","","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									27-Apr-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									03-May-2011						1.1									Case "Delete"															Sunny Ruparel
'@@												Sandeep Navghane									17-Oct-2011						1.2									Case "ExecuteSaveSearch"															Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_SaveMySearches(StrAction,StrFolder,StrName,bShared,StrNewFolderName)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SaveMySearches"
	'Variable Declaration
	Dim ObjSvdSrchs,ObjSrchFolder,objEle,objImg,WshShell,objEle1
	Dim iRowCount,iCounter,iFolderRwCount,iImageCount,iCount1
	Dim iChldItmCount,iCount,bFlag,iChldItmCount1,iCount2
	Dim objBtn,objChld
	Dim crrSvdSrch,objRdo,bReturn,ObjDsBtn

	bFlag=False
	Fn_Web_SaveMySearches=False
   'Creating Object Of "MySavedSearches" And "CreateSavedSearchFolder" Tables
	Set ObjSearchType=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SavedSearches").WebTable("SearchTypes")
	Set ObjModSvdSrch=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SavedSearches").WebTable("ModifyMySavedSearches")
	Set ObjSvdSrchs=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MySavedSearches")
	Set ObjSrchFolder=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CreateSavedSearchFolder")

	
	Select Case StrAction
		Case "New"
			If Not ObjSvdSrchs.Exist(5) Then
				Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SearchButtonTable"), "SaveSearch")
				wait(2)
			End If
			iRowCount=ObjSvdSrchs.RowCount
			For iCounter=0 To iRowCount
				If Trim("Folder:")=Trim(ObjSvdSrchs.GetCellData(iCounter,1)) Then
					iFolderRwCount= iCounter
					Exit For
				End If
			Next
			If StrFolder<>"My Saved Searches" Or StrNewFolderName<>"" Then
				If StrFolder<>"" Or StrNewFolderName<>"" Then
					iImageCount=ObjSvdSrchs.ChildItemCount(iFolderRwCount,2,"Image")
					For iCount1=0 To iImageCount-1
						Set objImg=ObjSvdSrchs.ChildItem(iFolderRwCount,2,"Image",iCount1)
						If  TypeName(objImg) <> "Nothing" Then
							If objImg.GetROProperty("file name") = "plus.png" Then
								objImg.Click 1,1
							End If
						End If
					Next
				End If
			End If
						
			If StrNewFolderName<>"" Then
				iChldItmCount=ObjSvdSrchs.ChildItemCount(iFolderRwCount,2,"WebElement")
				For iCount=0 To iChldItmCount-1
					Set objEle=ObjSvdSrchs.ChildItem(iFolderRwCount,2,"WebElement",iCount)
					If  TypeName(objEle) <> "Nothing" Then
						If Trim(objEle.GetROProperty("innertext")) =Trim(StrNewFolderName) Then
							objEle.Click 1,1
							bFlag=True
							Exit For
						End If
					End If
				Next
				If bFlag=False Then
					Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",ObjSvdSrchs, "NewFolder")
					wait(2)
					Call Fn_Web_UI_WebEdit_Set("Fn_Web_SaveMySearches",ObjSrchFolder,"Name", StrNewFolderName)
					Set objBtn=Description.Create
					objBtn("html tag").value="INPUT"
					objBtn("class").value="roundbutton"
					objBtn("value").value="OK"
					Set objChld=Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objBtn)
					objChld(0).click 1,1
				End If
			End If
			If StrFolder<>"" Then
				iChldItmCount1=ObjSvdSrchs.ChildItemCount(iFolderRwCount,2,"WebElement")
				For iCount2=0 To iChldItmCount1-1
					Set objEle1=ObjSvdSrchs.ChildItem(iFolderRwCount,2,"WebElement",iCount2)
					If  TypeName(objEle1) <> "Nothing" Then
						If Trim(objEle1.GetROProperty("innertext")) =Trim(StrFolder) Then
							objEle1.Click 1,1
							Exit For
						End If
					End If
				Next
			End If
			If StrName<>"" Then
			Wait(2)
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_SaveMySearches",ObjSvdSrchs,"Name",StrName)
			End If
			If bShared<>"" Then
				Call Fn_Web_UI_CheckBox_Set("Fn_Web_SaveMySearches", ObjSvdSrchs, "IsShared", bShared)
			End If
			bFlag=False
			Wait(3)
			bFlag=Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
			Wait(3)
			If bFlag=True Then
				Fn_Web_SaveMySearches=True
			End If

		Case "Delete"
			'Checking Existance Of "Change Search" Dialog
			If Not ObjSearchType.Exist(5) Then
				'If "Change Search" Dialog Not Exist Then Opening Dialog
				Call Fn_Web_UI_WebElement_Click("Fn_Web_SaveMySearches", Browser("TeamcenterWeb"), "AdvanceSearch", "","","")
				Call Fn_Web_UI_Link_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb"), "More","","","")
			End If
			'Selecting "My Saved Searches" Tab
			Call Fn_Web_UI_WebElement_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SavedSearches"), "MySavedSearches", "","","")
		
				iRowCount=ObjModSvdSrch.RowCount
				iCount=0
				For iCounter=4 To iRowCount
					If iCounter=4 Then
						iCount=0
					Else
						iCount=iCount+1
					End If
					crrSvdSrch=ObjModSvdSrch.GetCellData(iCounter,2)
					If Trim(crrSvdSrch)=Trim(StrName) Then
						Set objRdo=ObjModSvdSrch.ChildItem(iCounter,1,"WebRadioGroup",0)
						If TypeName(objRdo)<>"Nothing" Then
								objRdo.Select "#"&Cstr(iCount)
							bFlag=True
							Exit For
						End If
					End If
				Next 
			
			If bFlag=True Then
				Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",ObjModSvdSrch, "Delete")
				Call Fn_Web_ErrorMsgVerify("","OK")
            	Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
				Fn_Web_SaveMySearches=True
			End If
	Case "Modify"
		'Checking Existance Of "Change Search" Dialog
		If Not ObjSearchType.Exist(5) Then
			'If "Change Search" Dialog Not Exist Then Opening Dialog
			Call Fn_Web_UI_WebElement_Click("Fn_Web_SaveMySearches", Browser("TeamcenterWeb"), "AdvanceSearch", "","","")
			Call Fn_Web_UI_Link_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb"), "More","","","")
		End If
		'Selecting "My Saved Searches" Tab
		Call Fn_Web_UI_WebElement_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SavedSearches"), "MySavedSearches", "","","")
        iRowCount=ObjModSvdSrch.RowCount
			iCount=0
			For iCounter=4 To iRowCount
				If iCounter=4 Then
					iCount=0
				Else
					iCount=iCount+1
				End If
				crrSvdSrch=ObjModSvdSrch.GetCellData(iCounter,2)
				If Trim(crrSvdSrch)=Trim(StrName) Then
					Set objRdo=ObjModSvdSrch.ChildItem(iCounter,1,"WebRadioGroup",0)
					If TypeName(objRdo)<>"Nothing" Then
						objRdo.Select "#"&Cstr(iCount)
						bFlag=True
						Exit For
					End If
				End If
			Next 
		If bFlag=True Then
			Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",ObjModSvdSrch, "Modify")
			If StrName<>""  Then
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("RenameMySavedSearches"),"NewName",StrNewFolderName)
			End If
			If bShared<>"" Then
				Call Fn_Web_UI_CheckBox_Set("Fn_Web_SaveMySearches", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("RenameMySavedSearches"), "IsShared", bShared)
			End If
		End If
		Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("LoadAll").SetTOProperty "Name","OK"
		Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches", Browser("TeamcenterWeb").Page("MyTeamCenter"), "LoadAll")
		wait(2)
	
		Set ObjDsBtn=Description.Create
		ObjDsBtn("html tag").value="INPUT"
		ObjDsBtn("name").value="Close"
		Set ObjChld=Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(ObjDsBtn)
		ObjChld(0).Click 1,1
		Fn_Web_SaveMySearches=True
	Case "Verify"
		'Checking Existance Of "Change Search" Dialog
		If Not ObjSearchType.Exist(5) Then
			'If "Change Search" Dialog Not Exist Then Opening Dialog
			Call Fn_Web_UI_WebElement_Click("Fn_Web_SaveMySearches", Browser("TeamcenterWeb"), "AdvanceSearch", "","","")
			Call Fn_Web_UI_Link_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb"), "More","","","")
		End If
		'Selecting "My Saved Searches" Tab
		Call Fn_Web_UI_WebElement_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SavedSearches"), "MySavedSearches", "","","")
		iRowCount=ObjModSvdSrch.RowCount
		For iCounter=0 To iRowCount
			crrSvdSrch=ObjModSvdSrch.GetCellData(iCounter,2)
			If Trim(crrSvdSrch)=Trim(StrName) Then
				Fn_Web_SaveMySearches=True
				Exit For
			End If
		Next
		Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")

	Case "ExecuteSaveSearch"
			'Checking Existance Of "Change Search" Dialog
			If Not ObjSearchType.Exist(5) Then
				'If "Change Search" Dialog Not Exist Then Opening Dialog
				Call Fn_Web_UI_WebElement_Click("Fn_Web_SaveMySearches", Browser("TeamcenterWeb"), "AdvanceSearch", "","","")
				Call Fn_Web_UI_Link_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb"), "More","","","")
			End If
			'Selecting "My Saved Searches" Tab
			Call Fn_Web_UI_WebElement_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SavedSearches"), "MySavedSearches", "","","")
		
				iRowCount=ObjModSvdSrch.RowCount
				iCount=0
				For iCounter=4 To iRowCount
					If iCounter=4 Then
						iCount=0
					Else
						iCount=iCount+1
					End If
					crrSvdSrch=ObjModSvdSrch.GetCellData(iCounter,2)
					If Trim(crrSvdSrch)=Trim(StrName) Then
						Set objRdo=ObjModSvdSrch.ChildItem(iCounter,2,"Link",0)
						If TypeName(objRdo)<>"Nothing" Then
								objRdo.Click 0,0
							bFlag=True
							Exit For
						End If
					End If
				Next 
			
			If bFlag=True Then
				Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
				Fn_Web_SaveMySearches=True
			Else
				Call Fn_Web_UI_Button_Click("Fn_Web_SaveMySearches",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
			End If

	End Select
	'Releasing Object Of "MySavedSearches" And "CreateSavedSearchFolder" Tables
	Set ObjSvdSrchs=Nothing
	Set ObjSrchFolder=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SaveMySearches
'@@
'@@    Description				 :	Function Used to open Save Searches from Quick Links
'@@
'@@    Parameters			   :	1.StrAction : Action Name
'@@												  2.StrSavedSearchName : Save Search Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client
'@@
'@@    Examples					:	Call Fn_Web_MySavedSearchesLinkOperation("Select","TestSaveSearch1")
'@@												Call Fn_Web_MySavedSearchesLinkOperation("Verify","TestSaveSearch1")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									27-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_MySavedSearchesLinkOperation(StrAction,StrSavedSearchName)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_MySavedSearchesLinkOperation"
	Dim ObjSrchImg,ObjSrchName,bFlag
	Set ObjSrchImg=Browser("TeamcenterWeb").Image("MySavedSearches")
	Set ObjSrchName=Browser("TeamcenterWeb").Link("More")
	Fn_Web_MySavedSearchesLinkOperation=False
	Select Case StrAction
	 	Case "Select"
			bFlag=Fn_Web_MySavedSearchesLinkOperation("Verify",StrSavedSearchName)
			If bFlag=True Then
				ObjSrchName.SetTOProperty "text",StrSavedSearchName
				Call Fn_Web_UI_Link_Click("Fn_Web_MySavedSearchesLinkOperation",Browser("TeamcenterWeb"),"More","","","")
				wait(2)
				Fn_Web_MySavedSearchesLinkOperation=True
			End If
		Case "Verify"
			ObjSrchImg.Click
			wait(1)
			ObjSrchName.SetTOProperty "text",StrSavedSearchName
			If ObjSrchName.Exist(10) Then
				Fn_Web_MySavedSearchesLinkOperation=True
			End If
	End Select
	Set ObjSrchImg=Nothing
	Set ObjSrchName=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_MyWorklistOperations
'@@
'@@    Description				 :	Function Used to perform Operation On My WorkList Table
'@@
'@@    Parameters			   :	1.StrAction : Action Name
'@@												  2.strNode : Node Or Item Name
'@@												  3.strColName : Column Name
'@@												  4.strValue : Expected Value
'@@
'@@    Return Value		   	   : 	True Or False Or Column Names
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client
'@@
'@@    Examples					:	Call Fn_Web_MyWorklistOperations("GetAllColumnNames","","","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									27-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_MyWorklistOperations(strAction,strNode,strColName,strValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_MyWorklistOperations"
	Dim ObjWrkList
	Dim iCols,iCounter,ColName,arrColName
	Fn_Web_MyWorklistOperations=False
	Set ObjWrkList=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MyWorkList")
	If ObjWrkList.Exist(5) Then
		Select Case strAction
			Case "GetAllColumnNames"
				iCols=ObjWrkList.ColumnCount(1)
				For iCounter=2 To iCols
					ColName=ObjWrkList.GetCellData(1,iCounter)
					If iCounter=2 Then
						arrColName=ColName
					Else
						arrColName=arrColName+":"+ColName
					End If
				Next
				Fn_Web_MyWorklistOperations=arrColName
		End Select
	End If
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@     FUNCTION NAME   :   Fn_Web_ReviseItem()
'@@
'@@    DESCRIPTION     :   1. This function Revise ItemRevision
'@@ 
'@@   PARAMETERS      :   ''strLabels = 		- labels Seperated by "~"	
''@@   											strValues = 		- Values Seperated by  "~"
'@@
'@@   Return Value  :   True/False  
'@@
'@@   EXAMPLE : 1. To Revise Item Revision
''@@ 							 strLabels = "Rev ID~Name~Description"
''@@ 							 strValues = "B~NewItem1~New Item Desc"
'@@ 							 
'@@ 							 Fn_Web_ReviseItem(strLabels,strValues,"")
'@@
'@@  History					 :		
'@@												Developer Name												Date							Version						Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Pranav Ingle									   			27-Apr-2011			              1.0									Created														Sunny	
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ReviseItem(strLabels,strValues)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ReviseItem"
	'variable Declaration	
	Dim objMyTc, ObjButton,objWebChild,arrLabels,arrValues,iCount,iCounter,strCellData,iRowCount,bFlag
	Dim strMenu,objMDR
	Fn_Web_ReviseItem =false

	' Create Object of MyTc	
'	Set objMyTc = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ReviseItemRevision").WebTable("ItemInfo")
'	Set ObjButton = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel")
	
	Set objMyTc = Fn_SISW_Web_GetObject("ReviseItemRevision").WebTable("ItemInfo")
	Set ObjButton = Fn_SISW_Web_GetObject("ButtunPanel")
	strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "EditRevise")
	
	If Not objMyTc.Exist(SISW_MIN_TIMEOUT) Then
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(1)
	End If

	''-------------------------------------
	If strLabels <> "" Then
		If Fn_Web_UI_ObjectExist("",objMyTc) = True Then
			arrLabels = Split(strLabels,"~")
			arrValues= Split(strValues,"~")
				For iCount = 0 to UBound(arrLabels)
'						iRowCount = objMyTc.GetROProperty("rows")
						iRowCount = Fn_WEB_UI_Object_GetROProperty("Fn_Web_ReviseItem",objMyTc,"rows")
						bFlag=False
						For iCounter= 1 to iRowCount-1
							 strCellData = objMyTc.GetCellData(iCounter,1)
							 If arrLabels(iCount)="Rev ID" Then
								arrLabels(iCount)="Revision:"
							 End If
							If Instr(1,strCellData, arrLabels(iCount)) > 0 Then
								Set objWebChild = objMyTc.ChildItem(iCounter,2,"WebEdit",0)
								'creating object of Mercury device replay
								If arrLabels(iCount)="Revision:" Then
									objWebChild.object.focus()
									Set objMDR = CreateObject("Mercury.DeviceReplay")
								    	objMDR.SendString arrValues(iCount)
								Else
									objWebChild.Set arrValues(iCount)
								End If
								bFlag = True
								Set objWebChild=Nothing
							End If
						Next
						If bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed Label"+arrLabels(iCount)+ "Not Found")
							Exit Function
						End If
				Next
		End If
	End If
	wait(WEB_MIN_TIMEOUT)
	'Click on  OK Button 
	Call Fn_Web_UI_Button_Click("Fn_Web_ReviseItem",ObjButton,"Finish")	
	wait(WEB_MINLESS_TIMEOUT)

	'Return Function Successful Log
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Function Successfully Completed")
	' Return True Value
	Fn_Web_ReviseItem =true
	
	' Setting created objects to nothing
	Set objMyTc = Nothing
	Set ObjButton = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CreateChange
'@@
'@@    Description				 :	Function Used To Create New Change
'@@
'@@    Parameters			   :	1.dicNewChange: Change Full Information Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:
'@@ 	Case "Create"	:		dicNewChange("Menu")="ChangeCreate"
'@@										dicNewChange("Action")="Create"
'@@										dicNewChange("Type")="Problem Report"
'@@										dicNewChange("ChangeID")="PR-000009"
'@@										dicNewChange("Revision")="A"
'@@										dicNewChange("Synopsis")="TestPR"
'@@										dicNewChange("Desc")="Testing"
'@@										Call Fn_Web_CreateChange(dicNewChange)
'@@ 	Case "VerifyTypes"	:		 dicNewChange("Menu")="ChangeCreate"
'@@												dicNewChange("Action")="VerifyTypes"
'@@												dicNewChange("Type")="Problem Report:Change Request:Change Notice"
'@@ 	Case "VerifyProblemItems"	:		 dicNewChange("Menu")="ChangeInContext"
'@@															dicNewChange("Action")="VerifyProblemItems"
'@@															dicNewChange("Type")="Change Notice"
'@@															dicNewChange("NodeName")="000220/A;1-Item123~000221/A;1-item123"
'@@												
'@@	   History					 	:	
'@@													Developer Name					Date						Rev. No.						Reviewer
'@@-------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@													Ketan Raje						27-Apr-2011						1.0							Sunny Ruparel
'@@													Sandeep N						21-Jun-2011						1.1							Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_CreateChange(dicNewChange)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CreateChange"
   'Variable Declaration
	Dim ObjChange, ObjChangeInfo, objobjMDR,aTypes, iCounter, iItems, iNodes
	Dim strWEBMenuPath, strMenu, intCnt, iCount, intCounter,crrType

	Fn_Web_CreateChange=False
	Set ObjChange=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewChangeItem")
	Set ObjChangeInfo=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewChangeItem").WebTable("ChangeInfo")
	Set ObjChangeType=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewChangeItem").WebTable("ChangeType")


	If Not ObjChange.Exist(5) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, dicNewChange("Menu"))	'Select appropriate sub menu for creating Change.
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If

	Select Case dicNewChange("Action")
	Case "Create","ErrorMsgVerify"
				'Setting Change Type
				If dicNewChange("Type")<>"" Then
					crrType=ObjChangeType.WebEdit("ChangeType").GetROProperty("value")
					wait(1)
					If Trim(crrType)<>Trim(dicNewChange("Type")) Then
						Call Fn_Web_UI_Button_Click("Fn_Web_CreateChange",ObjChangeType,"ChangeType")
						Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",dicNewChange("Type"))
						Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
						wait(2)
					End If
				End If
				'Clicking On Next Button
				Call Fn_Web_UI_Button_Click("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
				Set objobjMDR = CreateObject("Mercury.DeviceReplay")

				'Setting Change ID
				If dicNewChange("ChangeID")<>"" Then
'					Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateChange", ObjChangeInfo, "ID", dicNewChange("ChangeID"))
					ObjChangeInfo.WebEdit("ID").Object.focus
					objobjMDR.SendString dicNewChange("ChangeID")
					wait(1)
				End If
				'Setting Change Revision
				If dicNewChange("Revision")<>"" Then
'					Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateChange", ObjChangeInfo, "Revision", dicNewChange("Revision"))
					ObjChangeInfo.WebEdit("Revision").Object.focus
					objobjMDR.SendString dicNewChange("Revision")
					wait 1
				End If
				'Setting Change Synopsis
				If dicNewChange("Synopsis")<>"" Then
					Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateChange", ObjChangeInfo, "Synopsis", dicNewChange("Synopsis"))
					wait 1
				End If
				'Setting Change Description
				If dicNewChange("Desc")<>"" Then
					Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateChange", ObjChangeInfo, "Description", dicNewChange("Desc"))
					wait 1
				End If	
				'Setting Change Type
				If dicNewChange("Type")<>"" Then
					Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateChange", ObjChangeInfo, "ChangeType", dicNewChange("Type"))
					wait 1
				End If					
				'Clicking On Finish Button
				Call Fn_Web_UI_Button_Click("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Finish")
				wait(5)
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Cancel").SetTOProperty "name","Cancel"
				wait(1)
				If dicNewChange("Action") <> "ErrorMsgVerify" Then
					If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Cancel").Exist(6) Then
						If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Cancel").getroProperty("visible") = "True" Then
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("Cancel").Click
						End If
					End If
				End If
				For iCount=0 to 5
					If ObjChange.Exist(5) Then
						wait(5)
					End If
				Next
				Fn_Web_CreateChange = True
	Case "VerifyTypes"
				intCnt = 0
				iCounter = 0
				'Verify Change Types
				If dicNewChange("Type")<>"" Then
					aTypes = Split(dicNewChange("Type"),":",-1,1)
					Call Fn_Web_UI_Button_Click("Fn_Web_CreateChange",ObjChangeType,"ChangeType")
					For iCount = 0 to Ubound(aTypes)
						intCnt = intCnt + 1
						Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",aTypes(iCount))
						If Fn_Web_UI_ObjectExist("Fn_Web_CreateChange", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType")) = True Then
							iCounter = iCounter + 1
						End If
					Next
				End If				
				'Clicking On Cancel Button
				Call Fn_Web_UI_Button_Click("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Cancel")				
				If intCnt = iCounter Then
					Fn_Web_CreateChange = True
				End If
	Case "VerifyProblemItems"
				intCnt = 0
				intCounter = 0
				'Setting Change Type
				If dicNewChange("Type")<>"" Then
					Call Fn_Web_UI_Button_Click("Fn_Web_CreateChange",ObjChangeType,"ChangeType")
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",dicNewChange("Type"))
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
					wait(2)
				End If
				'Clicking On Next Button
				Call Fn_Web_UI_Button_Click("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
				'Verify all items in Problem Items List.
				iCount = ObjChangeInfo.WebList("ProblemItems").GetROProperty("items count")
				iItems = Split(dicNewChange("NodeName"),"~",-1,1)
				For iCounter = 0 to Ubound(iItems)
					intCnt = intCnt + 1
					For iNodes = 1 to iCount
						If Trim(Lcase(ObjChangeInfo.WebList("ProblemItems").GetItem(iNodes))) = Trim(Lcase(iItems(iCounter))) Then
							intCounter = intCounter + 1
							Exit For
						End If
					Next
				Next
				'Clicking On Cancel Button
				Call Fn_Web_UI_Button_Click("Fn_Web_CreateChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Cancel")				
				If intCnt = intCounter Then
					Fn_Web_CreateChange = True
				End If
	End Select
	
	'Releasing All Objects Of Tables
	Set ObjChange=Nothing
	Set ObjChangeInfo=Nothing
	Set ObjChangeType=Nothing
	Set objobjMDR = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_InformationVerify
'@@
'@@    Description				 :	Function Used to Verify Information and handle the Dialog.
'@@
'@@    Parameters			   :	1.strErrorMsg : Error Message
'@@												  2.strButton : Button Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Information Dialog Should be Appear on Screen						
'@@
'@@    Examples					:	Msgbox Fn_Web_InformationVerify("No relationship is defined for the selected object(s): Form","OK")
'@@
'@@	   History					 	:	
'@@													Developer Name					Date						Rev. No.							Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@													Ketan Raje						29-Apr-2011						1.0								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_InformationVerify(strErrorMsg,strButton)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_InformationVerify"
	GBL_EXPECTED_MESSAGE=strErrorMsg
    Dim iCount, objWebTab, objWebTabChld
	Fn_Web_InformationVerify=False

	Set objWebTab = Description.Create
	objWebTab("micclass").value = "WebTable"
'	objWebTab("innertext").value = strErrorMsg
	Set objWebTabChld = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objWebTab)
	For i=0 to objWebTabChld.Count-1
		If Instr(1,Trim(Lcase(objWebTabChld(i).getroproperty("innertext"))),Trim(Lcase(strErrorMsg))) > 0 Then
			Fn_Web_InformationVerify = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified Information.")
			Exit For
		Else
			GBL_ACTUAL_MESSAGE=objWebTabChld(i).getroproperty("innertext")
		End If
	Next
	'Click on the button.
	If Fn_Web_InformationVerify = True Then
		If strButton = "Close" Then
			Call Fn_Web_UI_Link_Click("Fn_Web_InformationVerify", Browser("TeamcenterWeb").Page("MyTeamCenter"), "InformationClose",  "","","")
		Else
			Call Fn_Web_UI_Button_ClickExt("Fn_Web_InformationVerify", "Click", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"), "OK")
		End If
	End If
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_WorkflowProcessAssign
'@@
'@@    Description				 :	Function Used to Assign New Process
'@@
'@@    Parameters			   :	1.StrProcName : Process Name
'@@												 2.StrDescription : Process Description
'@@												 3.StrProcessTemFilter : Process Template Filter														
'@@												 4.StrProcessTemplate : Process Template
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client and Object should be Selected on which process have to assign							
'@@
'@@    Examples					:	Call Fn_Web_WorkflowProcessAssign("000212/A;1-Process2","Unit Test","All","TCM Release Process")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									29-Apr-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_WorkflowProcessAssign(StrProcName,StrDescription,StrProcessTemFilter,StrProcessTemplate)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_WorkflowProcessAssign"
   'Variable Declaration
	Dim objProcess, ObjMyTcPage
	Dim strWEBMenuPath,strMenu,StrCrrFilter,StrCrrTemp,bFlag,iCounter
	'Creating Object "New Process" Dialog
'	Set objProcess=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewProcess")
	Set objProcess=Fn_SISW_Web_GetObject("NewProcess")
	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	Fn_Web_WorkflowProcessAssign=False
	bFlag=False
	
	strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
	strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewProcess")
	
'	Checking Existance Of "New Process" Dialog
	If objProcess.Exist(5) Then
'	If Fn_Web_UI_ObjectExist("Fn_Web_WorkflowProcessAssign",objProcess)=True Then
		'Calling "File->Workflow Process..." Menu Option
'		If objProcess.GetROProperty("height")=0 Then
		If Fn_WEB_UI_Object_GetROProperty("Fn_Web_WorkflowProcessAssign",objProcess,"height") = 0 Then
			Call Fn_Web_MenuOperation("Select",strMenu)
			Wait(WEB_MINLESS_TIMEOUT)
'			objProcess.SetTOProperty "index",1
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_WorkflowProcessAssign",objProcess,"index",1)
		End If
	Else
		Call Fn_Web_MenuOperation("Select",strMenu)
		Wait(WEB_MINLESS_TIMEOUT)
	End If

	'Setting Process Name
	If StrProcName<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_WorkflowProcessAssign",objProcess,"ProcessName",StrProcName)
	End If
	'Setting Process Description
	If StrDescription<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_WorkflowProcessAssign",objProcess,"Description",StrDescription)
	End If
	'Setting Process Template Filter
	If StrProcessTemFilter<>"" Then
		'Taking Current Process Template Filter
'		StrCrrFilter=objProcess.WebEdit("TemplateFilter").GetROProperty("value")
		StrCrrFilter=Fn_WEB_UI_Object_GetROProperty("Fn_Web_WorkflowProcessAssign",objProcess.WebEdit("TemplateFilter"),"value") 
		'Matching Current Template Filter with Users Template Filter
		If Trim(StrCrrFilter)<>Trim(StrProcessTemFilter) Then
			'If Fiters Not match Then selecting Filter
			Call Fn_Web_UI_Button_Click("Fn_Web_WorkflowProcessAssign",objProcess,"TemplateFilter")
			wait(1)
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_WorkflowProcessAssign",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",StrProcessTemFilter)
'			Call Fn_Web_UI_WebElement_Click("Fn_Web_WorkflowProcessAssign",Browser("TeamcenterWeb").Page("MyTeamCenter"),"FormType","","","")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_WorkflowProcessAssign",ObjMyTcPage.WebElement("FormType"),"innertext",StrProcessTemFilter)
			Call Fn_Web_UI_WebElement_Click("Fn_Web_WorkflowProcessAssign",ObjMyTcPage,"FormType","","","")
			wait(WEB_MICRO_TIMEOUT)
			wait(1)
		End If
	End If
	'Setting Process Template
	If StrProcessTemplate<>"" Then
		'Taking Current Process Template Filter
'		StrCrrTemp=objProcess.WebEdit("Template").GetROProperty("value")
		StrCrrTemp=Fn_WEB_UI_Object_GetROProperty("Fn_Web_WorkflowProcessAssign",objProcess.WebEdit("TemplateFilter"),"value") 
		wait(WEB_MICRO_TIMEOUT)
		'Matching Current Template Filter with Users Template Filter
		If Trim(StrCrrTemp)<>Trim(StrProcessTemplate) Then
			'If Fiters Not match Then selecting Filter
			Call Fn_Web_UI_Button_Click("Fn_Web_WorkflowProcessAssign",objProcess,"Template")
'			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_WorkflowProcessAssign",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",StrProcessTemplate)
'			Call Fn_Web_UI_WebElement_Click("Fn_Web_WorkflowProcessAssign",Browser("TeamcenterWeb").Page("MyTeamCenter"),"FormType","","","")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_WorkflowProcessAssign",ObjMyTcPage.WebElement("FormType"),"innertext",StrProcessTemplate)
			Call Fn_Web_UI_WebElement_Click("Fn_Web_WorkflowProcessAssign",ObjMyTcPage,"FormType","","","")
			wait(WEB_MICRO_TIMEOUT)
		End If
	End If
'	Call Fn_Web_UI_Button_Click("Fn_Web_WorkflowProcessAssign", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"OK")
	Call Fn_Web_UI_Button_Click("Fn_Web_WorkflowProcessAssign", ObjMyTcPage.WebElement("ButtunPanel"),"OK")
    	Fn_Web_WorkflowProcessAssign=True
	
	For iCounter=0 To 2
'		If objProcess.Exist(5) Then
		If Fn_Web_UI_ObjectExist("Fn_Web_WorkflowProcessAssign",objProcess)=True Then
'			Call Fn_Web_UI_Button_Click("Fn_Web_WorkflowProcessAssign", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Close")
			Call Fn_Web_UI_Button_Click("Fn_Web_WorkflowProcessAssign", ObjMyTcPage.WebElement("ButtunPanel"),"Close")
			wait(WEB_MICROLESS_TIMEOUT)
		Else
			Exit For
		End If
	Next
	'Releasing Object "New Process" Dialog
	Set objProcess=Nothing
	Set ObjMyTcPage = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SetPreference
'@@
'@@    Description				 :	Function Used to Set Preferences
'@@
'@@    Parameters			   :	1.StrTabName : Tab Name
'@@												 2.StrAction : Action Name
'@@												 3.StrPrefName : Preference Name														
'@@												 4.StrPrefValue : Preference value
'@@												 5.StrPrefValue : Item Type
'@@												 6.StrShowCols : Show Column Name
'@@												 7.StrHideCols : Hide Column Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client.
'@@												IMP NOTE : - BEFORE And AFTER Calling this function either Call Set Perspective or Refresh the Page
'@@										
'@@    Examples					:	Call Fn_Web_SetPreference("Display","SetValue","Show all revisions","yes","","","")
'@@												Call Fn_Web_SetPreference("Advanced","SetValue","Group name length display limit","3","","","")
'@@												Call Fn_Web_SetPreference("Product Structure","GetValue","Default Revision Rule","","","","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									02-May-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									02-May-2011						1.0							Added Case "GetValue"					Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_SetPreference(StrTabName,StrAction,StrPrefName,StrPrefValue,StrItemType,StrShowCols,StrHideCols)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SetPreference"
	   'Variable Declaration
	   Dim ObjOptions,ObjPrefTbl,obLnk,ObjPrefValTbl,objEdt,objBtn
	   Dim strWEBMenuPath,strMenu,iRwCount,iCounter,crrPrefName,iRowCnt,iColCount,bFlag,bReturn
	   Dim ObjDsBtn,ObjChld,ObjMyTcPage
	'Creating Object Of "Options" Table
'	Set ObjOptions=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Options")
	Set ObjOptions = Fn_SISW_Web_GetObject("Options")
	Set ObjMyTcPage = Fn_SISW_Web_GetObject("MyTeamCenter")
	
    	Select Case StrTabName
		Case "Display"
'			Set ObjPrefTbl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Options").WebTable("DisplayPreferance")
			Set ObjPrefTbl=ObjOptions.WebTable("DisplayPreferance")
		Case "Search"
'			Set ObjPrefTbl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Options").WebTable("SearchPreferance")
			Set ObjPrefTbl=ObjOptions.WebTable("SearchPreferance")
		Case "Product Structure"
'			Set ObjPrefTbl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Options").WebTable("ProductStructurePreferance")
			Set ObjPrefTbl=ObjOptions.WebTable("ProductStructurePreferance")
		Case "Advanced"
'			Set ObjPrefTbl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Options").WebTable("AdvancePreferance")
			Set ObjPrefTbl=ObjOptions.WebTable("AdvancePreferance")
		Case else
'			Set ObjPrefTbl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Options").WebTable("Text:=.*"&StrPrefName&".*")
			Set ObjPrefTbl=ObjOptions.WebTable("Text:=.*"&StrPrefName&".*")
	End Select

	Fn_Web_SetPreference=False
	Browser("TeamcenterWeb").RefreshObject

	If Not ObjOptions.Exist(SISW_MIN_TIMEOUT) Then
		strMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("Web_Menu"), "EditOptions")
		Call Fn_Web_MenuOperation("Select",strMenu)
		Call Fn_Web_ReadyStatusSync(2)
	End If
	
	bFlag=False
	If ObjPrefTbl.Exist(3) Then
		'do nothing
		bFlag=True
	Else
		For iCounter=0 to 5
'			ObjOptions.WebTable("PreferanceTable").SetTOProperty "index",iCounter
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SetPreference",ObjOptions.WebTable("PreferanceTable"),"index",iCounter)
'			If ObjOptions.WebTable("PreferanceTable").getROProperty("height")>0 Then
			If Fn_WEB_UI_Object_GetROProperty("Fn_Web_SetPreference",ObjOptions.WebTable("PreferanceTable"),"height") > 0 Then
'				If instr(1,ObjOptions.WebTable("PreferanceTable").GetROProperty("text"),StrPrefName) Then
				If instr(1,Fn_WEB_UI_Object_GetROProperty("Fn_Web_SetPreference",ObjOptions.WebTable("PreferanceTable"),"text"),StrPrefName) Then
					Set ObjPrefTbl=ObjOptions.WebTable("PreferanceTable")
					bFlag=True
					Exit for
				End If
			End If
		Next
	End If
	If bFlag=False Then
		Exit function
	End If
	'Clicking On "Apropriate" Tab
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SetPreference",ObjOptions.WebElement("TabName"),"innertext",StrTabName)
	Call Fn_Web_UI_WebElement_Click("Fn_Web_SetPreference", ObjOptions,"TabName", "","","")
	iRwCount=ObjPrefTbl.RowCount
	For iCounter=0 To iRwCount
		crrPrefName=ObjPrefTbl.GetCellData(iCounter,1)
		If Trim(crrPrefName)=Trim(StrPrefName) Then
			Set obLnk=ObjPrefTbl.ChildItem(iCounter,1,"Link",0)
			If TypeName(obLnk)<>"Nothing" Then
				obLnk.click 1,1
				wait(WEB_MICROLESS_TIMEOUT)
				Exit For
			End If
			Set obLnk=Nothing
		End If
	Next
	If StrAction="Add" Or StrAction="Remove" Or StrAction="MoveUp" Or StrAction="MoveDown" Then
'		Set ObjPrefValTbl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Text:=.*Item Type:.*")
		Set ObjPrefValTbl=ObjMyTcPage.WebTable("Text:=.*Item Type:.*")
		If Not ObjPrefValTbl.Exist(3) Then
'			Set ObjPrefValTbl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Text:=.*ShowHide.*")
			Set ObjPrefValTbl=ObjMyTcPage.WebTable("Text:=.*ShowHide.*")
			If Not ObjPrefValTbl.Exist(1) Then
				Exit Function
			End If
		End If
	ElseIf StrAction="SetValue" Or StrAction="GetValue" Or StrAction="CancelSetValue" Then
'		Set ObjPrefValTbl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Text:=.*"&StrPrefName&":.*")
		Set ObjPrefValTbl=ObjMyTcPage.WebTable("Text:=.*"&StrPrefName&":.*")
	End If
	Select Case StrAction
		Case "SetValue", "CancelSetValue"
			iRowCnt=ObjPrefValTbl.RowCount
			If iRowCnt=1 Then
				iColCount=ObjPrefValTbl.ColumnCount(1)
				bFlag=False
				Set objBtn=ObjPrefValTbl.ChildItem(1,2,"WebButton",0)
				If TypeName(objBtn)<>"Nothing" Then
					bFlag=True
				End If
			End If
			bReturn=False
			If bFlag=False Then
				Set objEdt=ObjPrefValTbl.ChildItem(1,2,"WebEdit",0)
				If TypeName(objEdt)<>"Nothing" Then
					bReturn=True
				End If
			End If
			If bFlag=True Then
				Set objEdt=ObjPrefValTbl.ChildItem(1,2,"WebEdit",0)
				If TypeName(objEdt)<>"Nothing" Then
					If Trim(objEdt.GetROProperty("value"))<>Trim(StrPrefValue) Then
						objBtn.click 1,1
'						Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").SetTOProperty "innertext",StrPrefValue
'						Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click
						Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SetPreference",ObjMyTcPage.WebElement("FormType"),"innertext",StrPrefValue)
						Call Fn_Web_UI_WebElement_Click("Fn_Web_SetPreference", ObjMyTcPage, "FormType", "","","")
					End If
					Fn_Web_SetPreference=True
				End If
			ElseIf bReturn=True Then
				objEdt.Set StrPrefValue
				Fn_Web_SetPreference=True
			End If
			If StrAction = "CancelSetValue" Then
'				Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("LoadAll").SetTOProperty "Name","Cancel"
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SetPreference",ObjMyTcPage.WebButton("LoadAll"),"Name","Cancel")
'				Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", Browser("TeamcenterWeb").Page("MyTeamCenter"), "LoadAll")
				Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", ObjMyTcPage, "LoadAll")
				wait(WEB_MICROLESS_TIMEOUT)
			Else
'				Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("LoadAll").SetTOProperty "Name","OK"
'				Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", Browser("TeamcenterWeb").Page("MyTeamCenter"), "LoadAll")
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SetPreference",ObjMyTcPage.WebButton("LoadAll"),"Name","OK")
				Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", ObjMyTcPage, "LoadAll")
				wait(WEB_MICROLESS_TIMEOUT)
			End If

'			Set ObjDsBtn=Description.Create
'			ObjDsBtn("html tag").value="INPUT"
'			ObjDsBtn("name").value="Close"		
''			Set ObjChld=Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(ObjDsBtn)
'			Set ObjChld=ObjMyTcPage.ChildObjects(ObjDsBtn)
'			If ObjChld.Count > 0 Then
'				ObjChld(0).Click 1,1
'			End If
			Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", ObjMyTcPage.WebElement("ButtunPanel"), "Close")
		'Case To Retrieve Current value of Preference
		Case "GetValue"
			iRowCnt=ObjPrefValTbl.RowCount
        			Set objEdt=ObjPrefValTbl.ChildItem(1,2,"WebEdit",0)
			If TypeName(objEdt)<>"Nothing" Then
				Fn_Web_SetPreference=objEdt.GetROProperty("value")
			End If
		        	wait(WEB_MICRO_TIMEOUT)
		        	Set ObjDsBtn=Description.Create
			ObjDsBtn("html tag").value="INPUT"
			ObjDsBtn("name").value="Cancel"
'			Set ObjChld=Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(ObjDsBtn)
			Set ObjChld=ObjMyTcPage.ChildObjects(ObjDsBtn)
			If ObjChld.Count > 0 Then
				If  ObjChld.Count=2 Then
					ObjChld(1).Click 1,1
				Else
					ObjChld(0).Click 1,1
				End If
			End If	
			Set ObjChld=Nothing
			Set ObjDsBtn=Nothing
			wait(WEB_MICROLESS_TIMEOUT)

			Set ObjDsBtn=Description.Create
			ObjDsBtn("html tag").value="INPUT"
			ObjDsBtn("name").value="Close"
'			Set ObjChld=Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(ObjDsBtn)
			Set ObjChld=ObjMyTcPage.ChildObjects(ObjDsBtn)
			If ObjChld.Count > 0 Then
				ObjChld(0).Click 1,1
			End If
	'-------------------------------------------------------------------------------------		
	'TC11.3_20170509d.00_DIPRO_NewDevelopment_PoonamC_16Oct2017 : Added Cases to Add & remove relations for ItemRev.
	Case "AddRelation"
		If StrShowCols <> "" Then
			 iRwCount = ObjMyTcPage.WebList("ShowList").GetROProperty("items count")
			 bFlag = False
			 For iCounter = 1 To iRwCount
				 	crrPrefName = ObjMyTcPage.WebList("ShowList").GetItem(iCounter)
				 	If trim(crrPrefName) = trim(StrShowCols) Then
				 		Fn_Web_SetPreference = True
				 		bFlag = True
				 		Exit For 
				 	End If
			 Next
			 If bFlag = False Then
				iRwCount = ObjMyTcPage.WebList("HideList").GetROProperty("items count")
				 For iCounter = 1 To iRwCount
					 	crrPrefName = ObjMyTcPage.WebList("HideList").GetItem(iCounter)
					 	If trim(crrPrefName) = trim(StrShowCols) Then
					 		ObjMyTcPage.WebList("HideList").Select(iCounter-1) 
					 		Call Fn_Web_ReadyStatusSync(1)
					 		'Click on Add button
					 		ObjMyTcPage.WebButton("AddRemove").SetTOProperty "name","<"
					 		Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", ObjMyTcPage,"AddRemove")
					 		Call Fn_Web_ReadyStatusSync(1)
					 		Fn_Web_SetPreference = True
					 		Exit For 
					 	End If
				 Next
			End If
		End If
		'Click OK to Close dialog
		Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", ObjMyTcPage,"OK")
		Call Fn_Web_ReadyStatusSync(1)
		'Click Close option dialog
		Set ObjDsBtn=Description.Create
		ObjDsBtn("html tag").value="INPUT"
		ObjDsBtn("name").value="Close"
		Set ObjChld=ObjMyTcPage.ChildObjects(ObjDsBtn)
		If ObjChld.Count > 0 Then
			ObjChld(0).Click 1,1
		End If

	Case "RemoveRelation"
			If StrHideCols <> "" Then
			 iRwCount = ObjMyTcPage.WebList("HideList").GetROProperty("items count")
			 bFlag = False
			 For iCounter = 1 To iRwCount
				 	crrPrefName = ObjMyTcPage.WebList("HideList").GetItem(iCounter)
				 	If trim(crrPrefName) = trim(StrShowCols) Then
				 		Fn_Web_SetPreference = True
				 		bFlag = True
				 		Exit For 
				 	End If
			 Next
			 If bFlag = False Then
				iRwCount = ObjMyTcPage.WebList("ShowList").GetROProperty("items count")
				 For iCounter = 1 To iRwCount
					 	crrPrefName = ObjMyTcPage.WebList("ShowList").GetItem(iCounter)
					 	If trim(crrPrefName) = trim(StrHideCols) Then
					 		ObjMyTcPage.WebList("ShowList").Select(iCounter-1) 
					 		Call Fn_Web_ReadyStatusSync(1)
					 		'Click on Add button
					 		ObjMyTcPage.WebButton("AddRemove").SetTOProperty "name",">"
					 		Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", ObjMyTcPage,"AddRemove")
					 		Call Fn_Web_ReadyStatusSync(1)
					 		Fn_Web_SetPreference = True
					 		Exit For 
					 	End If
				 Next
			End If
		End If
		
		'Click OK to Close dialog
		Call Fn_Web_UI_Button_Click("Fn_Web_SetPreference", ObjMyTcPage,"OK")
		Call Fn_Web_ReadyStatusSync(1)
		'Click Close option dialog
		Set ObjDsBtn=Description.Create
		ObjDsBtn("html tag").value="INPUT"
		ObjDsBtn("name").value="Close"
		Set ObjChld=ObjMyTcPage.ChildObjects(ObjDsBtn)
		If ObjChld.Count > 0 Then
			ObjChld(0).Click 1,1
		End If
	'------------------------------------------------------------------------------------
	End Select
	Set ObjOptions=Nothing
	Set ObjPrefTbl=Nothing
	Set objBtn=Nothing
	Set objEdt=Nothing
	Set ObjMyTcPage = Nothing
End Function
'-------------------------------------------------------------------Function Used to perform operations on { AssignParticipants } Dialog--------------------------------------------------------------
'Function Name		:	Fn_Web_AssignParticipantsOperations

'Description			 :	Function Used to perform operatons on  Trees which are present on { AssignParticipants } Dialog

'Return Value		   : 	True Or False

'Pre-requisite			:	Revision of Object should be selected to Assign Participants.

'Examples				:
								'Following are fields handled and the respective parameters
								'With dicAssignParticipants
								'	.Add "sAction","Add"
								'	.Add "sMenu","AssignParticipants"
								'	.Add "sResponsibleParty","ProjTeam:Group:Role:User|OFF"
								'	.Add "sReviewers","Add:ProjTeam:Group:Role:User~Add:ProjTeam:Group:Role:User|OFF"
								'	'.Add "sAnalyst","ProjTeam:Group:Role:User|OFF"
								'	'.Add "sChangeImplementation","Add:ProjTeam:Group:Role:User~Remove:ProjTeam:Group:Role:User|OFF"
								'	.Add "sChangeSpecialist1","ProjTeam:Group:Role:User|OFF"
								'	'.Add "sRequestor","Add:ProjTeam:Group:Role:User|OFF"
								'	.Add "sButton","OK"
								'End With

'	Presently the function is used to set any one field of Assign Participants at a time.
'											Dim dicAssignParticipants
'											Set dicAssignParticipants = CreateObject("Scripting.Dictionary")
'	Case for setting				With dicAssignParticipants
'	"ChangeSpecialist1"		  		 .Add "sAction","Add"
'												.Add "sMenu","AssignParticipants"
'												.Add "sChangeSpecialist1","None:dba:DBA:AutoTestDBA (dba/DBA/autotestdba)|OFF"
'												.Add "sButton","OK"
'											End With
'											Environment.Value("TestLogFile") = "D:\Log.txt"
'											Environment.Value("WebBrowserName") = "IE"
'											Msgbox Fn_Web_AssignParticipantsOperations(dicAssignParticipants)
'											
'											dicAssignParticipants.RemoveAll
'	Case for setting				With dicAssignParticipants
'	"Analyst"								.Add "sAction","Add"
'												.Add "sMenu","AssignParticipants"
'												.Add "sAnalyst","None:dba:DBA:AutoTestDBA (dba/DBA/autotestdba)|OFF"
'												.Add "sButton","OK"
'											End With
'											Environment.Value("TestLogFile") = "D:\Log.txt"
'											Environment.Value("WebBrowserName") = "IE"
'											Msgbox Fn_Web_AssignParticipantsOperations(dicAssignParticipants)										   
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje.										   			03/05/2011			           1.0																			Sunny R.
'	Modified By:					Swapna Ghatge											28-Dec-2011
'
'	Changes Done :-  The "Change Specialist 1" label has been changed to "Change Specialist |".
'
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Web_AssignParticipantsOperations(dicAssignParticipants)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_AssignParticipantsOperations"
   'Variable declaration
	Dim ObjAssgnParticpntDialog, strWEBMenuPath, sMenu, aResponsibleParty, aPartyDetails, strWEBErrorPath
	Dim aReviewers, aReviews, objImg, arrReviewer, aAnalyst, aAnalystDetails, iRows, iRowNum, bFlag, sError, sStringApp
	Fn_Web_AssignParticipantsOperations=False
	bFlag = True
	'verifying Existance of { AssignParticipants } Dialog
	If Fn_Web_UI_ObjectExist("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants"))=False Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		sMenu=Fn_GetXMLNodeValue(strWEBMenuPath, dicAssignParticipants("sMenu"))	'Select appropriate sub menu for assigning participants.
		Call Fn_Web_MenuOperation("Select",sMenu)
	End If
	'Creating Object of { AssignParticipants } Dialog
	Set ObjAssgnParticpntDialog = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants")
	Set ObjParticpntUsers = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ParticiapntsUser")
	strWEBErrorPath=Fn_LogUtil_GetXMLPath("Web_ErrorMsg")
	sError=Fn_GetXMLNodeValue(strWEBErrorPath, "AssignParticipantsError")	'Select appropriate error message.
	Select Case dicAssignParticipants("sAction")
	Case "Add" , "Modify"
			'Set values for Propesed Responsible Party.	
				If dicAssignParticipants("sResponsibleParty") <> "" Then
						'Extract the row number.
						iRows = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").RowCount
						For iRowNum = 1 to iRows
							If Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Proposed Responsible Party" or Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Proposed Responsible Party:" Then
								Exit For
							End If
						Next
						'Click on the image of Proposed Responsible Party.
						Set objImg=ObjAssgnParticpntDialog.ChildItem(iRowNum,3,"Image",0)
						objImg.Click 1,1
						'Handle Information msgbox
						bFlag = Fn_Web_InformationVerify(sError, "OK")
						If bFlag = True Then
							Fn_Web_AssignParticipantsOperations = False
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Assigning participants is not allowed.")	
							'Clicking On Cancel Button
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")					
							Exit Function
						End If
						aResponsibleParty = Split(dicAssignParticipants("sResponsibleParty"), "|", -1, 1)
						If Trim(Lcase(aResponsibleParty(1))) = "on" Then
							'Set Select Member from Project Team ON
							Call Fn_Web_UI_CheckBox_Set("Fn_Web_AssignParticipantsOperations", ObjParticpntUsers, "MemberOpt", "ON")
						End If
						aPartyDetails = Split(aResponsibleParty(0), ":", -1, 1)
						'Set the value for Project Team		'To be coded as required.
						'Set Group
						If Trim(Lcase(aPartyDetails(1))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Group")
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aPartyDetails(1)
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
							wait(1)
						End If
						'Set Role
						If Trim(Lcase(aPartyDetails(2))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Role")
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aPartyDetails(2)
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
							wait(1)
						End If
						'Set User
						If Trim(Lcase(aPartyDetails(3))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"User")
							wait(2)
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aPartyDetails(3)
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
							wait(2)
						End If
						'Clicking On OK Button
						Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"OK")
				End If
			'Set values for Proposed Reviewers.
				If dicAssignParticipants("sReviewers") <> "" Then
						'Extract the row number.
						iRows = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").RowCount
						For iRowNum = 1 to iRows
							sStringApp = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)
							'If Trim(sStringApp) = "Proposed Reviewers" or Trim(sStringApp) = "Proposed Reviewers:" Then
							If Trim(sStringApp) = "Proposed Reviewers" or Trim(sStringApp) = "Proposed Reviewer:" Then
								Exit For
							End If
						Next
						'Click on the image of Proposed Reviewers.
						Set objImg=ObjAssgnParticpntDialog.ChildItem(iRowNum,3,"Image",0)
						objImg.Click 1,1
						'Handle Information msgbox
						bFlag = Fn_Web_InformationVerify(sError, "OK")
						If bFlag = True Then
							Fn_Web_AssignParticipantsOperations = False
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Assigning participants is not allowed.")	
							'Clicking On Cancel Button
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")					
							Exit Function
						End If
						aReviews = Split(dicAssignParticipants("sReviewers"),"|",-1,1)
						If Trim(Lcase(aReviews(1))) = "on" Then
							'Set Select Member from Project Team ON
							Call Fn_Web_UI_CheckBox_Set("Fn_Web_AssignParticipantsOperations", ObjParticpntUsers, "MemberOpt", "ON")
						End If
						aReviewers = Split(aReviews(0),"~",-1,1)
						ReDim arrReviewer(Cint(Ubound(aReviewers))+1)
						For iCount = 0 to Ubound(aReviewers)
							arrReviewer(iCount) = Split(aReviewers(iCount),":",-1,1)
							'Set the value for Project Team		'To be coded as required.
							'Set Group
							If Trim(Lcase(arrReviewer(iCount)(2))) <> "none" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Group")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",arrReviewer(iCount)(2)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
							End If
							'Set Role
							If Trim(Lcase(arrReviewer(iCount)(3))) <> "none" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Role")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",arrReviewer(iCount)(3)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
							End If
							'Set User
							If Trim(Lcase(arrReviewer(iCount)(4))) <> "none" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"User")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",arrReviewer(iCount)(4)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
							End If
							'Click on Add OR Remove button.\		
							If Trim(Lcase(arrReviewer(iCount)(0))) = "add" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"Add")
							ElseIf Trim(Lcase(arrReviewer(iCount)(0))) = "remove" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"Remove")
							End If
						Next
						'Clicking On OK Button
						Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"OK")
				End If
			'Set values for Analyst.
				If dicAssignParticipants("sAnalyst") <> "" Then
						'Extract the row number.
						iRows = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").RowCount
						For iRowNum = 1 to iRows
							If Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Analyst" or Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Analyst:" Then
								Exit For
							End If
						Next
						'Click on the image of Analyst.
						Set objImg=ObjAssgnParticpntDialog.ChildItem(iRowNum,3,"Image",0)
						objImg.Click 1,1
						'Handle Information msgbox
						bFlag = Fn_Web_InformationVerify(sError, "OK")
						If bFlag = True Then
							Fn_Web_AssignParticipantsOperations = False
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Assigning participants is not allowed.")	
							'Clicking On Cancel Button
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")
							Exit Function
						End If
						aAnalyst = Split(dicAssignParticipants("sAnalyst"), "|", -1, 1)
						If Trim(Lcase(aAnalyst(1))) = "on" Then
							'Set Select Member from Project Team ON
							Call Fn_Web_UI_CheckBox_Set("Fn_Web_AssignParticipantsOperations", ObjParticpntUsers, "MemberOpt", "ON")
						End If
						aAnalystDetails = Split(aAnalyst(0), ":", -1, 1)
						'Set the value for Project Team		'To be coded as required.
						'Set Group
						If Trim(Lcase(aAnalystDetails(1))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Group")
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aAnalystDetails(1)
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
							wait(1)
						End If
						'Set Role
						If Trim(Lcase(aAnalystDetails(2))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Role")
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aAnalystDetails(2)
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
							wait(1)
						End If
						'Set User
						If Trim(Lcase(aAnalystDetails(3))) <> "none" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"User")
								wait(2)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aAnalystDetails(3)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(2)
						End If
						'Clicking On OK Button
						Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"OK")
				End If
			'Set values for Change Implementation Board.
				If dicAssignParticipants("sChangeImplementation") <> "" Then
						'Extract the row number.
						iRows = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").RowCount
						For iRowNum = 1 to iRows
							If Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Change Implementation Board" OR Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Change Review Board" Then
								Exit For
							Elseif Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Change Implementation Board:" OR Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Change Review Board:" Then
								Exit For
							End If
						Next
						'Click on the image of Change Implementation Board.
						Set objImg=ObjAssgnParticpntDialog.ChildItem(iRowNum,3,"Image",0)
						objImg.Click 1,1
						'Handle Information msgbox
						bFlag = Fn_Web_InformationVerify(sError, "OK")
						If bFlag = True Then
							Fn_Web_AssignParticipantsOperations = False
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Assigning participants is not allowed.")	
							'Clicking On Cancel Button
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")
							Exit Function
						End If
						aChangeImplementation = Split(dicAssignParticipants("sChangeImplementation"),"|",-1,1)
						If Trim(Lcase(aChangeImplementation(1))) = "on" Then
							'Set Select Member from Project Team ON
							Call Fn_Web_UI_CheckBox_Set("Fn_Web_AssignParticipantsOperations", ObjParticpntUsers, "MemberOpt", "ON")
						End If
						aChangeImple = Split(aChangeImplementation(0),"~",-1,1)
						ReDim arrChangeImple(Cint(Ubound(aChangeImple))+1)
						For iCount = 0 to Ubound(aChangeImple)
							arrChangeImple(iCount) = Split(aChangeImple(iCount),":",-1,1)
							'Set the value for Project Team		'To be coded as required.
							'Set Group
							If Trim(Lcase(arrChangeImple(iCount)(2))) <> "none" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Group")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",arrChangeImple(iCount)(2)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
							End If
							'Set Role
							If Trim(Lcase(arrChangeImple(iCount)(3))) <> "none" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Role")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",arrChangeImple(iCount)(3)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
							End If
							'Set User
							If Trim(Lcase(arrChangeImple(iCount)(4))) <> "none" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"User")								
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",arrChangeImple(iCount)(4)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
							End If
							'Click on Add OR Remove button.\		
							If Trim(Lcase(arrChangeImple(iCount)(0))) = "add" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"Add")
							ElseIf Trim(Lcase(arrChangeImple(iCount)(0))) = "remove" Then
								Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"Remove")
							End If
						Next
						'Clicking On OK Button
						Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"OK")
				End If
			'Set values for Change Specialist1
				If dicAssignParticipants("sChangeSpecialist1") <> "" Then
						'Extract the row number.
						iRows = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").RowCount
						For iRowNum = 1 to iRows
							If Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Change Specialist I" or Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Change Specialist I:" Then
								Exit For
							End If
						Next
						'Click on the image of Change Specialist1
						Set objImg=ObjAssgnParticpntDialog.ChildItem(iRowNum,3,"Image",0)
						objImg.Click 1,1
						'Handle Information msgbox
						bFlag = Fn_Web_InformationVerify(sError, "OK")
						If bFlag = True Then
							Fn_Web_AssignParticipantsOperations = False
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Assigning participants is not allowed.")	
							'Clicking On Cancel Button
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")
							Exit Function
						End If
						aChangeSpecialist1 = Split(dicAssignParticipants("sChangeSpecialist1"), "|", -1, 1)
						If Trim(Lcase(aChangeSpecialist1(1))) = "on" Then
							'Set Select Member from Project Team ON
							Call Fn_Web_UI_CheckBox_Set("Fn_Web_AssignParticipantsOperations", ObjParticpntUsers, "MemberOpt", "ON")
						End If
						aSpecialistDetails = Split(aChangeSpecialist1(0), ":", -1, 1)
						'Set the value for Project Team		'To be coded as required.
						'Set Group
						If Trim(Lcase(aSpecialistDetails(1))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Group")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aSpecialistDetails(1)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
						End If
						'Set Role
						If Trim(Lcase(aSpecialistDetails(2))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Role")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aSpecialistDetails(2)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
						End If
						'Set User
						If Trim(Lcase(aSpecialistDetails(3))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"User")
								wait(2)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aSpecialistDetails(3)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(2)
						End If
						'Clicking On OK Button
						Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"OK")
				End If
			'Set values for Requestor
				If dicAssignParticipants("sRequestor") <> "" Then
						'Extract the row number.
						iRows = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").RowCount
						For iRowNum = 1 to iRows
							If Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Requestor" or Trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignParticipants").GetCellData(iRowNum,1)) = "Requestor:" Then
								Exit For
							End If
						Next
						'Click on the image of Requestor
						Set objImg=ObjAssgnParticpntDialog.ChildItem(iRowNum,3,"Image",0)
						objImg.Click 1,1
						'Handle Information msgbox
						bFlag = Fn_Web_InformationVerify(sError, "OK")
						If bFlag = True Then
							Fn_Web_AssignParticipantsOperations = False
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Assigning participants is not allowed.")	
							'Clicking On Cancel Button
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")
							Exit Function
						End If
						aRequestor = Split(dicAssignParticipants("sRequestor"), "|", -1, 1)
						If Trim(Lcase(aRequestor(1))) = "on" Then
							'Set Select Member from Project Team ON
							Call Fn_Web_UI_CheckBox_Set("Fn_Web_AssignParticipantsOperations", ObjParticpntUsers, "MemberOpt", "ON")
						End If
						aRequestorDetails = Split(aRequestor(0), ":", -1, 1)
						'Set the value for Project Team		'To be coded as required.
						'Set Group
						If Trim(Lcase(aRequestorDetails(1))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Group")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aRequestorDetails(1)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
						End If
						'Set Role
						If Trim(Lcase(aRequestorDetails(2))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"Role")
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aRequestorDetails(2)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(1)
						End If
						'Set User
						If Trim(Lcase(aRequestorDetails(3))) <> "none" Then
							Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",ObjParticpntUsers,"User")
								wait(2)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "html tag","LI"
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").SetTOProperty "innertext",aRequestorDetails(3)
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Information").Click 1,1
								wait(2)
						End If
						'Clicking On OK Button
						Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WebButtons"),"OK")
				End If
	End Select
	'Click on "Cancel" OR "OK" button.
	Call Fn_Web_UI_Button_Click("Fn_Web_AssignParticipantsOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),dicAssignParticipants("sButton"))
	'Handle the Says Dialog box
	Call Fn_Web_ErrorMsgVerify("Participants are set successfully.","OK")
	Fn_Web_AssignParticipantsOperations = True
	Set ObjAssgnParticpntDialog=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_IDDisplayRulesOperation
'@@
'@@    Description				 :	Function Used To perform Operation On ID Display Rules
'@@
'@@    Parameters			   :	1.dicIDDisplayRules: Id Display Rules Full Information Dictionary Object
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	dicIDDisplayRules("RuleName")="TestRule1"
'@@												
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									03-May-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_IDDisplayRulesOperation(StrAction,dicIDDisplayRules)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_IDDisplayRulesOperation"
 	'Variable Declaration
	Dim objIDDispRl,objCrtIDRule
	Dim strWEBMenuPath,strMenu,iRuleCount,iCounter,crrRule,bFlag
	bFlag=False
	Fn_Web_IDDisplayRulesOperation=False
	'Creating Object "IDDisplayRules" & "CreateIDDisplayRule" Table
	Set objIDDispRl=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("IDDisplayRules")
	Set objCrtIDRule=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CreateIDDisplayRule")
	'Checking Existance Of "IDDisplayRules" Dialog
	If Not objIDDispRl.Exist(6) Then
		'Calling "Tools->Id Display Rule..." Menu
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "ToolsIdDisplayRules")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
   Select Case StrAction
		Case "Create"		'Case to Create New Display Rule
			iRuleCount=objIDDispRl.WebList("RuleList").GetROProperty("items count")
			For iCounter=1 To iRuleCount
				crrRule=objIDDispRl.WebList("RuleList").GetItem(iCounter)
				If Trim(crrRule)=Trim(dicIDDisplayRules("RuleName")) Then
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				'Clicking On Create Button To Open "Create Id Diaplay Rule" Dialog
				Call Fn_Web_UI_Button_Click("Fn_Web_IDDisplayRulesOperation",objIDDispRl,"Create")
				'Setting Rule Name
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_IDDisplayRulesOperation", objCrtIDRule, "RuleName",dicIDDisplayRules("RuleName"))
				'Clicking On 'OK' Button
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("LoadAll").SetTOProperty "Name","OK"
				Call Fn_Web_UI_Button_Click("Fn_Web_IDDisplayRulesOperation", Browser("TeamcenterWeb").Page("MyTeamCenter"), "LoadAll")
				Fn_Web_IDDisplayRulesOperation=True
			Else
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("LoadAll").SetTOProperty "Name","Cancel"
				Call Fn_Web_UI_Button_Click("Fn_Web_IDDisplayRulesOperation", Browser("TeamcenterWeb").Page("MyTeamCenter"), "LoadAll")
			End If
			Call Fn_Web_UI_Button_Click("Fn_Web_IDDisplayRulesOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Close")
   End Select
	Set objIDDispRl=Nothing
	Set objCrtIDRule=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SummaryTabOperations
'@@
'@@    Description				 :	Function Used To perform Operation On [ Summary ] Or [ Overview ]  Tab
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@												 2.StrPropValue : Property Value Pair ( Property Name : Value , Property Name : Value)
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client And Object Need to be Selected On Which have to perform Operations (Eg : Item)							
'@@
'@@    Examples					:	Call Fn_Web_SummaryTabOperations("Verify","Owner:AutoTest3 (autotest3),Name:tEST,Last Modifying User:AutoTest3")
'@@												Call Fn_Web_SummaryTabOperations("Verify","Checked-Out:Y")												
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									05-May-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_SummaryTabOperations(StrAction,StrPropValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SummaryTabOperations"
 	'Variable Declaration
	Dim ObjProperty,ObjSmryTab,objEdit
	Dim bFlag,iRowCount,arrProp,iCounter,arrValues,iCount,crrPropName,crrPropValue,bReturn
	Dim iCount1,crrPropName1,crrPropValue1
	bReturn=False
	bFlag=False
	Fn_Web_SummaryTabOperations=False
	'Creating Object Of [ SummaryTabProperties ] WebTable
	Set ObjProperty=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SummaryTabProperties")
	Set ObjSmryTab=Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Overview")
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SummaryTabOperations",ObjSmryTab,"innertext","Summary")
	'Chaecking Existence Of [ Summary ] Tab
	If ObjSmryTab.Exist(10) Then
		''Clicking on [ Summary ] Tab
		Call Fn_Web_UI_WebElement_Click("Fn_Web_SummaryTabOperations",Browser("TeamcenterWeb").Page("MyTeamCenter"),"Overview", "","","")
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Succesfully Click on Summary Tabs")
	Else
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SummaryTabOperations",ObjSmryTab,"innertext","Overview")
		'Chaecking Existence Of [ Overview ] Tab
		If ObjSmryTab.Exist(5) Then
			''Clicking on [ Overview ] Tab
			Call Fn_Web_UI_WebElement_Click("Fn_Web_SummaryTabOperations",Browser("TeamcenterWeb").Page("MyTeamCenter"),"Overview", "","","")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Succesfully Click on Overview Tabs")
		Else
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SummaryTabOperations",ObjSmryTab,"innertext","General")
			If ObjSmryTab.Exist(5) Then
				''Clicking on [ Overview ] Tab
				Call Fn_Web_UI_WebElement_Click("Fn_Web_SummaryTabOperations",Browser("TeamcenterWeb").Page("MyTeamCenter"),"Overview", "","","")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Succesfully Click on Overview Tabs")
			Else
				'If Both The Tabs Are not Exist Then Function Will Exit Returns False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Summary or General or Overview Tabs Are Not Exist currently")
				Set ObjProperty=Nothing
				Set ObjSmryTab=Nothing
				Exit Function
			End If
		End If
	End If
	'Taking Row Count Of Table
	iRowCount=ObjProperty.RowCount
	Select Case StrAction
		Case "Verify"	'Case to Verify Properties From Summary Tab
			For iCount1=0 To iRowCount
				crrPropName1=ObjProperty.GetCellData(iCount1,1)
				wait 1
				If Trim(crrPropName1)=Trim("Checked-Out:") Or Trim(crrPropName1)=Trim("Checked-Out") Then
					crrPropValue1=ObjProperty.GetCellData(iCount1,2)
					wait 1
					If Trim(crrPropValue1)=Trim("Y") Then
						bReturn=True
						Exit For
					End If
				End If
			Next
			If bReturn=False Then
				arrProp=Split(StrPropValue,",")
				For iCounter=0 To UBound(arrProp)
					bFlag=False
					arrValues=Split(arrProp(iCounter),":")
					For iCount=0 To iRowCount
						crrPropName=ObjProperty.GetCellData(iCount,1)
						wait 1
						If Trim(crrPropName)=Trim(arrValues(0)+":") Or Trim(crrPropName)=Trim(arrValues(0)) Then
							crrPropValue=ObjProperty.GetCellData(iCount,2)
							wait 1
							If Trim(crrPropValue)=Trim(arrValues(1)) Then
								bFlag=True
								Exit For
							End If
						End If
					Next
					If bFlag=False Then
						Exit For
					End If
				Next
				If bFlag=True Then
					Fn_Web_SummaryTabOperations=True
				End If
			Else
				arrProp=Split(StrPropValue,",")
				For iCounter=0 To UBound(arrProp)
					bFlag=False
					arrValues=Split(arrProp(iCounter),":")
					For iCount=0 To iRowCount

						crrPropName=ObjProperty.GetCellData(iCount,1)
						wait 1
						If Trim(crrPropName)=Trim(arrValues(0)+":") Or Trim(crrPropName)=Trim(arrValues(0)) Then
							Set objEdit=ObjProperty.ChildItem(iCount,2,"WebEdit",0)
							If TypeName(objEdit)<>"Nothing" Then
								crrPropValue=objEdit.GetROProperty("value")
								Set objEdit=Nothing
							Else
								crrPropValue=ObjProperty.GetCellData(iCount,2)
								wait 1
							End If
							If Trim(crrPropValue)=Trim(arrValues(1)) Then
								bFlag=True
								Exit For
							End If
						End If
					Next
					If bFlag=False Then
						Exit For
					End If
				Next
				If bFlag=True Then
					Fn_Web_SummaryTabOperations=True
				End If
			End If
	End Select
	'Releasing Objects
	Set ObjProperty=Nothing
	Set ObjSmryTab=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_FileDownLoadOperations
'@@
'@@    Description				 :	Function Used to Perform Operations on File DownLoad Dialog
'@@
'@@    Parameters			   :	1.StrAction : Action Name
'@@										 2.StrAskAgainOptn : Ask Me Again Option
'@@										 3.StrLocation : File Location Path to Save
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	File DownLoad Dialog Should be Open
'@@
'@@    Examples					:	Call Fn_Web_FileDownLoadOperations("Open","","")
'@@										Call Fn_Web_FileDownLoadOperations("Open","Off","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									6-May-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_FileDownLoadOperations(StrAction,StrAskAgainOptn,StrLocation)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_FileDownLoadOperations"
 	'Variable Declaration
	Dim ObjFlDwnLd
	Dim StrBrowser
	Fn_Web_FileDownLoadOperations=False
	StrBrowser=Environment.Value("WebBrowserName")
	'Creating Object Of "File Download" Dialog
	If InStr(1,StrBrowser,"FF")>0 Then
		Set ObjFlDwnLd=Browser("TeamcenterWeb").Dialog("FFFileDownload")
	Else
		Set ObjFlDwnLd=Dialog("IEFileDownload")
	End If
	'Checking Existence Of "File Download" Dialog
	If ObjFlDwnLd.Exist(10) Then
		Select Case StrAction
			Case "Open"
				If InStr(1,StrBrowser,"FF")>0 Then
					ObjFlDwnLd.Activate
					wait(2)
					If StrAskAgainOptn<>"" Then
						ObjFlDwnLd.Page("FFFileDownloadPage").WebCheckBox("AlwaysAskOptn").Set StrAskAgainOptn
					End If
					ObjFlDwnLd.Page("FFFileDownloadPage").WebButton("OK").Click 1,1,micLeftBtn
					wait(2)
				Else
					ObjFlDwnLd.Activate
					wait(2)
					If StrAskAgainOptn<>"" Then
						ObjFlDwnLd.WinCheckBox("AlwaysAskOptn").Set StrAskAgainOptn
					End If
					ObjFlDwnLd.WinButton("Open").Click 1,1,micLeftBtn
					wait(2)
				End If
				Fn_Web_FileDownLoadOperations=True
		End Select
	End If
	Set ObjFlDwnLd=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_UploadFile
'@@
'@@    Description				 :	Function Used To Upload External Files
'@@
'@@    Parameters			   :	1.StrFilePath: File Full Path including File name And Extension
'@@												 2.StrReference :Reference Type
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client
'@@
'@@    Examples					:	Call Fn_Web_UploadFile("C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\Siemens.jpg","")
'@@												Call Fn_Web_UploadFile("C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\Sunset.jpg","Image")										
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									11-May-2011						1.0																								Sunny Ruparel
'@@---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Shwetambari Rathod									11-Jun-2014										                 Added Code to handle open dialog box
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_UploadFile(StrFilePath,StrReference)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_UploadFile"
 	'Variable Declaration
	Dim objJApp,objMyTcPage
	Dim strBrowser,crrRef
	Fn_Web_UploadFile=False
	'Taking Working Browser Name
	strBrowser=Environment.Value("WebBrowserName")
	'Creating Object Of [ JApplet ] And [ MyTeamCenter ] Page
	Set objJApp=Browser("TeamcenterWeb").Page("MyTeamCenter").JavaApplet("JApplet")
	Set objMyTcPage=Browser("TeamcenterWeb").Page("MyTeamCenter")
	
	'Click on the 'Browse' button and select an file  added by shweta rathod
	If strFilePath<>"" Then
		Call Fn_Web_UI_Button_Click("Fn_Web_UploadFile",Browser("TeamcenterWeb").Page("MyTeamCenter"), "Browse")
		If  JavaDialog("UploadFile").Exist(3) Then
			If InStr(1,strBrowser,"IE")>0 Then
				JavaDialog("UploadFile").JavaEdit("FileName").Set strFilePath
			else
				JavaDialog("UploadFile").JavaEdit("FileName").Set ""
				JavaDialog("UploadFile").JavaEdit("FileName").Type StrFilePath
			End if
			JavaDialog("UploadFile").JavaButton("Open").Click micLeftBtn
			Wait 2
		End If
	End If
	'Choosing [ Reference ] Type
	If StrReference<>"" Then
		crrRef=objMyTcPage.WebList("Reference").GetROProperty("value")
		If Trim(crrRef)<>Trim(StrReference) Then
			objMyTcPage.WebList("Reference").Select StrReference
			wait 1
		End If
	End If
	'Clicking On [ UpLoad ] Button
	objMyTcPage.WebButton("Upload").Click
	wait 2
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Succesfully Uploaded File "&StrFilePath)
	Fn_Web_UploadFile=True

    If Browser("TeamcenterWeb").Dialog("Dialog").Exist(5) Then
		Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").Click
		wait 2
	End If
	'Releasing Object Of [ JApplet ] And [ MyTeamCenter ] Page
	Set objJApp=Nothing
	Set objMyTcPage=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CommonTableOperations
'@@
'@@    Description				 :	Function Used to perform Operation On BOM Table
'@@
'@@    Parameters			   :	1.sAction : Action Name
'@@												  2.sNodeName : Node Path Or Item Name
'@@												  3.sColumn : Column Name
'@@												  4.sCellValue : Expected Value
'@@
'@@    Return Value		   	   : 	True Or False Or Column Names Or Image Name
'@@
'@@    Pre-requisite			:	Should Be Log in Web Client And PSEperspective should be open
'@@
'@@    Examples					:	Call Fn_Web_CommonTableOperations("GetImage","","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@				cases		NodeSelect / Select 		- Call  Fn_Web_CommonTableOperations("NodeSelect","","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@														  Call  Fn_Web_CommonTableOperations("Select","","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View) @2:000057/A;1-asm","","")
'@@							NodeDeSelect / Deselect 	- Call Fn_Web_CommonTableOperations("NodeDeSelect","","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							NodeVerify / Exist /Exists 	- Call Fn_Web_CommonTableOperations("NodeVerify","","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View):000056/A;1-SubItem2","","")
'@@							CellVerify					- Call Fn_Web_CommonTableOperations("CellVerify","","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View):000056/A;1-SubItem2","Item Type","Item")
'@@							Collapse					- Call Fn_Web_CommonTableOperations("Collapse","","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							MultiSelect					- Call Fn_Web_CommonTableOperations("MultiSelect","","000054/A;1-TopItem (View)~000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							Expand						- Call Fn_Web_CommonTableOperations("Expand","","000054/A;1-TopItem (View):000055/A;1-SubItem1 (View)","","")
'@@							ExpandBelow					- Call Fn_Web_CommonTableOperations("ExpandBelow","","000054/A;1-TopItem (View)","","")
'@@							GetImage					- Call Fn_Web_CommonTableOperations("GetImage","","","","")
'@@							FirstElement				- Call Fn_Web_CommonTableOperations("FirstElement","","","","")
'@@							ColumnExist					- Call Fn_Web_CommonTableOperations("ColumnExist","","","Name~BOM Line~Find No.","")
'@@							ColumnClick					- Call Fn_Web_CommonTableOperations("ColumnClick","","","Name~BOM Line~Find No.","")
'@@							ClickLink					- Call Fn_Web_CommonTableOperations("ClickLink","","REQ-000001/A;1-vgg (View)","","Edit")
'@@							NodeClick					- Call Fn_Web_CommonTableOperations("ClickLink","","REQ-000001/A;1-vgg (View)","","Edit")
'@@							CellEdit					- Call Fn_Web_CommonTableOperations("CellEdit","","000015/A;1-top (View):000016/A;1-sub","Find No.","20") 
'@@
'@@	   History:				Developer Name										Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Sandeep Navghane									12-May-2011						1.0																		Sunny Ruparel
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe										12-May-2011						1.0					Added case CellEdit
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@							Koustubh Watwe										12-May-2011						1.0					Modify case CellVerify
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_Web_CommonTableOperations(sAction,objTable,sNodeName, sColumn, sCellValue)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CommonTableOperations"
		' Declaration of an Variable
		Dim objDialog, objImg, objWebChk, objLink, objButton,strImageName
        Dim aElements, aSubElement, iCounter, bFlag, jCounter, iRowCnt, iColPos, iOuterCnt, sText, iCounter2, iRowCnt2
		Dim objTRs, objTDs, objElements,iRow,iCol,i,j,sVal
		' Initialization of an Variable
		If TypeName(objTable)="String" Then
			Set objDialog =Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommonTable")
		Else
			Set objDialog =objTable
		End If

		bFlag = False
		Fn_Web_CommonTableOperations = False
		'  Operations Case
		Select Case sAction
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
        		Case "NodeSelect", "Select"
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
							Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
							If TypeName(objWebChk) <> "Nothing" Then
								If objWebChk.GetROProperty("checked") = "0" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = True
								elseIf objWebChk.GetROProperty("checked") = "1" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = True
								End If
							End If
							Set objWebChk = Nothing
					else
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_Web_CommonTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node ["+CStr(sNodeName)+"] Selected Successfully ")
					Else
						Fn_Web_CommonTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to Select the Nod ["+CStr(sNodeName)+"] . ")
					End If

			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "NodeVerify", "Exist", "Exists"
					' Write the Log of Success or Failure
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
						Fn_Web_CommonTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node  ["+CStr(sNodeName)+"] Verified Successfully. ")
					Else
						Fn_Web_CommonTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to Verify the Node  ["+CStr(sNodeName)+"] . ")
					End If

			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "CellVerify"
					' For Getting the Column Header Position
					If sColumn = "" Then
						iColPos = Fn_WebUI_TableColumnIndex(objDialog,"Name")
					else
						iColPos = Fn_WebUI_TableColumnIndex(objDialog, sColumn)
					End If
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iColPos <> -1 and iRowCnt <> -1 Then
						Set objLink = objDialog.ChildItem(iRowCnt, iColPos, "WebEdit", 0)
						If TypeName(objLink) <> "Nothing" Then
							If trim(objLink.GetROProperty("value")) = Trim(sCellValue) then 
								bFlag = True
							end if
						ElseIf Trim(objDialog.GetCellData(iRowCnt, iColPos)) = Trim(sCellValue) Then
								bFlag = True
						End If
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_Web_CommonTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations :  Value [ " & sCellValue & "]  is successfully verified for column [ " & sColumn & " ] at Node ["+CStr(sNodeName)+"]. ")
					Else
						Fn_Web_CommonTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations :  Failed to verify value [ " & sCellValue & "]  for column [ " & sColumn & " ] at Node ["+CStr(sNodeName)+"]. ")
					End If

			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -				
			Case "Expand"
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
					If iRowCnt <> -1 Then
							Set objImg = objDialog.ChildItem(iRowCnt, iColPos, "Image", 0)
							If TypeName(objImg) <> "Nothing" Then
									If objImg.GetROProperty("file name") = "plus.png" Then
												objImg.Click 1,1, micLeftBtn
												bFlag = True
									elseIf objImg.GetROProperty("file name") = "minus.png" Then
												bFlag = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations: node  ["+CStr(sNodeName)+"] was already expanded.")
									else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations: can not expand node  ["+CStr(sNodeName)+"].")
									End If
							End If						
							Set objImg = Nothing
					else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations: node  ["+CStr(sNodeName)+"] does not exist in BOM table.")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_Web_CommonTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node  ["+CStr(sNodeName)+"] expanded successfully. ")
					Else
						Fn_Web_CommonTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to expand node  ["+CStr(sNodeName)+"]")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -				
			Case "ExpandBelow"
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					If iRowCnt <> -1 Then
							Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
							If TypeName(objWebChk) <> "Nothing" Then
								    If objWebChk.GetROProperty("checked") = "0" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = Fn_Web_MenuOperation("Select",Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewExpandBelow"))
								elseIf objWebChk.GetROProperty("checked") = "1" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = Fn_Web_MenuOperation("Select",Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_PSE_Menu"), "ViewExpandBelow"))
								End If
							End If
							Set objWebChk = Nothing
					else
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_Web_CommonTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node  ["+CStr(sNodeName)+"] expanded successfully. ")
					Else
						Fn_Web_CommonTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to expand node  ["+CStr(sNodeName)+"]")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "GetImage"
'				objDialog.RefreshObject
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
					If iRowCnt <> -1 and iColPos <> -1 Then
							Set objImg = objDialog.ChildItem(iRowCnt, iColPos, "Image", 0)
							If TypeName(objImg) <> "Nothing" Then
									strImageName=Split(objImg.GetROProperty("file name"),".")
									Fn_Web_CommonTableOperations=strImageName(0)
									bFlag = True
							Else
									Fn_Web_CommonTableOperations=False
							End If						
							Set objImg = Nothing
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_Web_CommonTableOperations = strImageName(0)
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Image [ " & Fn_Web_CommonTableOperations & " ]  associated with node  ["+CStr(sNodeName)+"].")
					Else
									Fn_Web_CommonTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations :  Failed to retrieve image name of node  ["+CStr(sNodeName)+"].")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "Collapse"
'		
'					aSubElement = Split(sNodeName, ":", -1, 1)
'					jCounter = 0
'					iRowCnt = objDialog.RowCount
'					iCounter = 1
'					Do While 1 = 1
'							objDialog.RefreshObject
'							iRowCnt = objDialog.RowCount
'							' For Last Node of an Element
'							If jCounter=UBound(aSubElement) Then
'									If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
'												Set objImg = objDialog.ChildItem(iCounter, 2, "Image", 0)
'												If TypeName(objImg) <> "Nothing" Then
'														
'													If objImg.GetROProperty("file name") = "minus.png" Then
'																objImg.Click 1,1, micLeftBtn
'													End If
'														
'												End If
'												bFlag = True
'												jCounter = jCounter + 1
'												Set objImg = Nothing
'												Exit Do 
'									End If
'							Else
'									' For the Node Hierarchy of an Element
'									If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
'												Set objImg = objDialog.ChildItem(iCounter, 2, "Image", 0)
'												If objImg.GetROProperty("file name") = "plus.png" Then
'															objImg.Click 1,1, micLeftBtn
'												End If
'												jCounter = jCounter + 1
'												Set objImg = Nothing
'									End If
'							End If
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
					If iRowCnt <> -1 and iColPos <> -1 Then
							Set objImg = objDialog.ChildItem(iRowCnt, iColPos, "Image", 0)
							If TypeName(objImg) <> "Nothing" Then
								If objImg.GetROProperty("file name") = "minus.png" Then
											objImg.Click 1,1, micLeftBtn
											bFlag = True
								End If
							End If
					end if
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_Web_CommonTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node  ["+CStr(sNodeName)+"] collapsed Successfully. ")
					Else
									Fn_Web_CommonTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to collapse the Node  ["+CStr(sNodeName)+"] . ")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -		
			Case "MultiSelect"
						aElements = Split(sNodeName, "~", -1, 1)
						For iOuterCnt = 0 To UBound(aElements)
								iRowCnt = Fn_WebUI_TableRowIndex(objDialog, aElements(iOuterCnt), "")
								If iRowCnt <> -1 and iColPos <> -1 Then
										Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
										If TypeName(objWebChk) <> "Nothing" Then
												If objWebChk.GetROProperty("checked") = "0" Then
														objWebChk.Click 
														bFlag = True
												End If
										End If
										Set objWebChk = Nothing
								else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to find node ["+CStr(aElements(iOuterCnt))+"] . ")	
										Exit for
								End If
						Next
						If bFlag = True Then
								' For Success Log
								Fn_Web_CommonTableOperations = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Multiple Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] selected Successfully. ")
						Else
								' For Failure Log
								Fn_Web_CommonTableOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to Select the Multiple Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] . ")	
						End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "NodeDeSelect", "Deselect"
					aElements = Split(sNodeName, "~", -1, 1)
					jCounter = 0
					For iOuterCnt = 0 To UBound(aElements)
								iRowCnt = Fn_WebUI_TableRowIndex(objDialog, aElements(iOuterCnt), "")
								If iRowCnt <> -1 and iColPos <> -1 Then
										Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
										If TypeName(objWebChk) <> "Nothing" Then
												If objWebChk.GetROProperty("checked") = "1" Then
														objWebChk.Click 
														bFlag = True
												End If
										End If
										Set objWebChk = Nothing
								else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to find node ["+CStr(aElements(iOuterCnt))+"] . ")	
										Exit for
								End If
						Next
					' Write the Log of Success or Failure
					If bFlag = True Then
							Fn_Web_CommonTableOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Deselected  Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] successfully. ")
					Else
							' For Failure Log
							Fn_Web_CommonTableOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to deselect  the Nodes  ["+CStr(replace(sNodeName,"~",", "))+"] . ")	
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "NodeClick"
					' For Getting the Column Header Position
					jCounter = objDialog.ColumnCount(1)
					aSubElement = Split(sNodeName, ":", -1, 1)
					If Trim(sColumn) <> "" Then ' here
							iRowCnt = 1
					Else
							iRowCnt = 2
					End If
					For iCounter = iRowCnt To jCounter
						Set objLink = objDialog.ChildItem(1, iCounter, "Link", 0)
						If TypeName(objLink) <> "Nothing" Then
									If Trim(objLink.GetROProperty("text")) = Trim(sColumn) Then
											iColPos = iCounter
											Exit For
									End If
						Else
									sText = objDialog.GetCellData(1, iCounter)
									If Trim(sText) = Trim(sColumn) Then
											iColPos = iCounter
											Exit For
									End If
									sText = ""
						End If
						Set objLink = Nothing
					Next
					jCounter = 0
					iRowCnt = objDialog.RowCount
					iCounter = 1
					If Trim(sCellValue) = "" and Trim(sColumn) = "" Then
								iColPos = iColPos - 1
					End If
					Do While 1 = 1
							objDialog.RefreshObject
							iRowCnt = objDialog.RowCount
							' For Last Node of an Element
							If jCounter=UBound(aSubElement) Then
								If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
										iRowCnt2 = objDialog.ChildItemCount(iCounter, iColPos, "Link")
										 For iCounter2 = 0 To  iRowCnt2 -1
												Set objLink = objDialog.ChildItem(iCounter, iColPos, "Link", iCounter2)
												If TypeName(objLink) <> "Nothing" Then
															If Trim(objLink.GetROProperty("text")) = Trim(sCellValue) Then
																bFlag = True
																objLink.Click
																Exit For
															End If
												End If
										 Next
										 If Trim(sCellValue) = "" Then
											Set objLink = objDialog.ChildItem(iCounter, iColPos, "Link", 0)
											If TypeName(objLink) <> "Nothing" Then
												bFlag = True
												objLink.Click
											End If
										 End If
										jCounter = jCounter + 1
										Set objLink = Nothing
										Exit Do
								End If
							Else
								' For the Node Hierarchy of an Element
								If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
										jCounter = jCounter + 1
								End If
							End If
							' Exit from the Loop when Total Rows Finished without finding an Node
							If iRowCnt = iCounter  Then
									Exit Do
							Else
									' Increment the Counter for next level
									iCounter = iCounter + 1
							End If
					Loop
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_Web_CommonTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Node  ["+CStr(sNodeName)+"]  Clicked Successfully. ")
					Else
									Fn_Web_CommonTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click the Node  ["+CStr(sNodeName)+"] . ")
					End If

				' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "ClickLink"
					' For Getting the Column Header Position
					jCounter = objDialog.ColumnCount(1)
					aSubElement = Split(sNodeName, ":", -1, 1)
					If Trim(sColumn) <> "" Then ' here
							iRowCnt = 1
					Else
							iRowCnt = 2
					End If
					For iCounter = iRowCnt To jCounter
						Set objLink = objDialog.ChildItem(1, iCounter, "Link", 0)
						If TypeName(objLink) <> "Nothing" Then
									If Trim(objLink.GetROProperty("text")) = Trim(sColumn) Then
											iColPos = iCounter
											Exit For
									End If
						Else
									sText = objDialog.GetCellData(1, iCounter)
									If Trim(sText) = Trim(sColumn) Then
											iColPos = iCounter
											Exit For
									End If
									sText = ""
						End If
						Set objLink = Nothing
					Next
					jCounter = 0
					iRowCnt = objDialog.RowCount
					iCounter = 1
					If Trim(sCellValue) = "" and Trim(sColumn) = "" Then
								iColPos = iColPos - 1
					End If
					Do While 1 = 1
							objDialog.RefreshObject
							iRowCnt = objDialog.RowCount
							' For Last Node of an Element
							If jCounter=UBound(aSubElement) Then
								If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
										iRowCnt2 = objDialog.ChildItemCount(iCounter, iColPos, "Link")
										 For iCounter2 = 0 To  iRowCnt2 -1
												Set objLink = objDialog.ChildItem(iCounter, iColPos, "Link", iCounter2)
												If TypeName(objLink) <> "Nothing" Then
															If Trim(objLink.GetROProperty("text")) = Trim(sCellValue) Then
																bFlag = True
																objLink.Click
																Exit For
															End If
												End If
										 Next
										 If Trim(sCellValue) = "" Then
											Set objLink = objDialog.ChildItem(iCounter, iColPos, "Link", 0)
											If TypeName(objLink) <> "Nothing" Then
												bFlag = True
												objLink.Click
											End If
										 End If
										jCounter = jCounter + 1
										Set objLink = Nothing
										Exit Do
								End If
							Else
								' For the Node Hierarchy of an Element
								If Trim(objDialog.GetCellData(iCounter, 2)) = Trim(aSubElement(jCounter)) Then
										jCounter = jCounter + 1
								End If
							End If
							' Exit from the Loop when Total Rows Finished without finding an Node
							If iRowCnt = iCounter  Then
									Exit Do
							Else
									' Increment the Counter for next level
									iCounter = iCounter + 1
							End If
					Loop
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_Web_CommonTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Node  ["+CStr(sNodeName)+"]  Clicked Successfully. ")
					Else
									Fn_Web_CommonTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Failed to Click the Node  ["+CStr(sNodeName)+"] . ")
					End If
			
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "ColumnExist"
					If sColumn <> "" Then
						aElements = Split(sColumn, "~", -1, 1)
						For iCounter = 0 To UBound(aElements)
							iColPos = Fn_WebUI_TableColumnIndex(objDialog,aElements(iCounter))
							If iColPos <> -1  Then
								bFlag = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Column  [ " & aElements(iCounter) & " ]  exists in BOM table. ")
							else
								bFlag = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to check existence of column  [ " & aElements(iCounter) & " ]. ")
								Exit for
							End If
						Next
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
									Fn_Web_CommonTableOperations = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : All Columns  ["+CStr(replace(sColumn,"~", ", "))+"]  exists in BOM table. ")
					Else
									Fn_Web_CommonTableOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to check existence of column(s) ["+CStr(replace(sColumn,"~", ", "))+"]. ")
					End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "FirstElement"
						' For Getting the Column Header Position
						iRowCnt = objDialog.GetROProperty("rows")
						iColPos = Fn_WebUI_TableColumnIndex(objDialog, "Name")
						If iRowCnt > 1 Then
								iRowCnt2 = objDialog.GetCellData(2, iColPos)
						End If
						If iRowCnt2 <> "" Then
								Fn_Web_CommonTableOperations = iRowCnt2
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : First Element ["+CStr(iRowCnt2)+"] Present in the BOM Table ")
						Else
								Fn_Web_CommonTableOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : No First Element Found in BOM Table ")
						End If

			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "ColumnClick"
						' For Getting the Column Header Position
						If sColumn <> "" Then
							iColPos = Fn_WebUI_TableColumnIndex(objDialog, sColumn)' For Row Number
							iOuterCnt = 0
							If iColPos <> -1 Then
								iRowCnt = objDialog.ChildItemCount(1, iColPos, "Link")
								If iRowCnt > 0 Then
											Set objLink = objDialog.ChildItem(1, iColPos, "Link", 0)
											If TypeName(objLink) <> "Nothing" Then
													If Trim(objLink.GetROProperty("text")) = Trim(sColumn) Then
															objLink.Click 
															bFlag = True
													End If
											End If
											Set objLink = Nothing
								End If
							End If
						End If
						If bFlag = True Then
							' Write the Log of Success or Failure
							Fn_Web_CommonTableOperations = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Column Heading  ["+CStr(sColumn)+"]  Clicked Successfully. ")
						Else
							Fn_Web_CommonTableOperations = False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Column Heading  ["+CStr(sColumn)+"]  Not Found to Click. ")
						End If
			' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
        		Case "CellEdit"
					objDialog.RefreshObject
					iRowCnt = Fn_WebUI_TableRowIndex(objDialog, sNodeName, "")
					iColPos = Fn_WebUI_TableColumnIndex(objDialog, sColumn)' For Row Number
					If iRowCnt <> -1 Then
							Set objWebChk = objDialog.ChildItem(iRowCnt, 1, "WebCheckBox", 0)
							If TypeName(objWebChk) <> "Nothing" Then
								If objWebChk.GetROProperty("checked") = "0" Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node ["+CStr(sNodeName)+"] found.")
										objWebChk.Click 1, 1, micLeftBtn
										bFlag = True
								End If

								Set objElements = Description.Create
								objElements("html tag").value = "TR"
								Set objTRs =  objDialog.ChildObjects(objElements)
								objElements("html tag").value = "TD"
								Set objTDs = objTRs(iRowCnt - 1 ).ChildObjects(objElements)
								objTDs(iColPos -1).click

								Set objLink = objDialog.ChildItem(iRowCnt, iColPos, "WebEdit", 0)
								If TypeName(objLink) <> "Nothing" Then
									bFlag = True
									objLink.set sCellValue
								End If
							End If
'						End If
							Set objWebChk = Nothing
					else
						     Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to find Node ["+CStr(sNodeName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_Web_CommonTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Node ["+CStr(sNodeName)+"] Selected Successfully ")
					Else
						Fn_Web_CommonTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to Select the Nod ["+CStr(sNodeName)+"] . ")
					End If
				'----------------------------------------------------------------------------------------------------------------------------------------
				Case "GetCellData"
					iRow = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommonTable").RowCount()
					iCol=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommonTable").ColumnCount(3)
					
					For i = 1 To iRow-1
						For j = 1 To iCol-1
							sVal=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommonTable").GetCellData(i,j)
							If sVal=sNodeName Then
								bFlag = True
								Exit For 
							End If
						Next
					Next
					If bFlag = True Then
						Fn_Web_CommonTableOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_CommonTableOperations : Get Node ["+CStr(sNodeName)+"] Successfully ")
					Else
						Fn_Web_CommonTableOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_CommonTableOperations : Failed to Get the Node ["+CStr(sNodeName)+"] . ")
					End If
		End Select
		Set objDialog = Nothing
		Set objImg = Nothing
		Set objWebChk = Nothing
		Set objLink = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_TabOperations
'@@
'@@    Description				 :	Function Used To Perform Operations On Tabs
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@												 2.StrTab :Tab Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client
'@@
'@@    Examples					:	Call Fn_Web_TabOperations("Verify","Overview:Impact Analysis:Details")
'@@												Call Fn_Web_TabOperations("Activate","Details")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									11-May-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_TabOperations(StrAction,StrTab)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_TabOperations"
	Dim objTab
	Dim bFlag,arrTab,iCounter
	Fn_Web_TabOperations=False
   	bFlag=False
	Set objTab=Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Overview")
	Select Case StrAction
		Case "Activate","Click"
			bFlag=Fn_Web_TabOperations("Verify",StrTab)
			If bFlag=True Then
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_TabOperations",objTab,"innertext",StrTab)
				Call Fn_Web_UI_WebElement_Click("Fn_Web_TabOperations", Browser("TeamcenterWeb").Page("MyTeamCenter"), "Overview", "","","")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Fn_Web_TabOperations: Succesfully Activated Tab [ " & StrTab & " ]")
				Fn_Web_TabOperations=True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fn_Web_TabOperations: Failed to Activated Tab [ " & StrTab & " ] because its not Exist")
			End If
		Case "Verify","Exist"
			arrTab=Split(StrTab,":")
			For iCounter=0 To UBound(arrTab)
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_TabOperations",objTab,"innertext",arrTab(iCounter))
				bFlag=objTab.Exist(5)
				If bFlag=False Then
					Fn_Web_TabOperations=False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fn_Web_TabOperations: Tab [ " & arrTab(iCounter) & " ] is not Exist")
					Exit For
				End If
				Fn_Web_TabOperations=True
			Next
	End Select
	Set objTab=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_SaveAsObject
'@@
'@@    Description				 :	Function Used To Save As the Object
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@												 2.StrTab :Tab Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client
'@@
'@@    Examples					:	'dicSaveAsObject("ID")="16645"
'@@												dicSaveAsObject("Name")="New231"
'@@												dicSaveAsObject("Description")="NewD1escription"
'@@												Call Fn_Web_SaveAsObject(dicSaveAsObject)
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									02-June-2011						1.0																								Sunny Ruparel
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Changed By Prathama									23-Apr-2015						1.0						Added OR condition for Vendor Part Name as Field name changed																		Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_SaveAsObject(dicSaveAsObject)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_SaveAsObject"
	Dim objSaveAs,objSaveAsRev
	Dim strWEBMenuPath,strMenu,objTable
	Dim bFlag,objButton,objPageChld,bHgt,bFlag1
	Dim iRowCnt,iCounter,objEdit,crrRwData,iCount
	Dim StrTabName,i,crrTabName
	Dim objSaveAsObject,objMDR

	Set objSaveAsObject=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAs").WebTable("SaveAsObject")
	Set objSaveAs=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsObject").WebTable("SaveAs")
	Set objSaveAsRev=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsObject").WebTable("SaveAsItemInformation")

	Set objMDR = CreateObject("Mercury.DeviceReplay")
	Fn_Web_SaveAsObject=False
	strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
	strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "EditSaveAs")

	bFlag=False
	Set objButton=Description.Create
	objButton("micClass").value="WebButton"
	objButton("Name").value="Cancel"
	Set objPageChld= Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objButton)
	For iCounter=1 To objPageChld.Count-1
		bHgt=objPageChld(iCounter).GetROProperty("height")
		If bHgt>0 Then
			bFlag=True
		End If
	Next
	If bFlag=False Then
		Call Fn_Web_MenuOperation("Select",strMenu)	
		wait 3
	End If

	bFlag=False
	bFlag1=False
	If objSaveAs.Exist(5) Then
		If objSaveAs.GetROProperty("height")>0 Then
			Set objTable=objSaveAs
			bFlag=True
		ElseIf objSaveAsRev.GetROProperty("height")>0 Then
			Set objTable=objSaveAsRev
		End If
	ElseIf objSaveAsRev.Exist(5) Then
		If objSaveAsRev.GetROProperty("height")>0 Then
			Set objTable=objSaveAsRev
		ElseIf objSaveAsObject.GetROProperty("height")>0 Then
			Set objTable=objSaveAsObject
			bFlag1=True
		End If
	Elseif objSaveAsObject.Exist(5) then
		If objSaveAsObject.GetROProperty("height")>0 Then
			Set objTable=objSaveAsObject
			bFlag1=True
		End If
	End If

	iRowCnt=objTable.RowCount
	If bFlag1=True Then
		'Do nothing
	Else
		Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",".*Information")
		crrTabName=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").GetROProperty("innertext")
	End If

	If dicSaveAsObject("ID")<>"" Then
		If bFlag=False and bFlag1=False Then
			For i=1 to 10
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").SetTOProperty "index",i	
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",".*Information")
				StrTabName=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").GetROProperty("innertext")
				If instr(1,StrTabName,"Information") Then
					If instr(1,StrTabName,"Revision") or instr(1,StrTabName,"Master") or instr(1,StrTabName,"FullText") Then
						'Do nothing
					Else
						Exit for
					End If
				End If
			Next
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",StrTabName)
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").Click 1,1,micLeftBtn
			wait(2)
			Set objTable=objSaveAs
		End if
		For iCounter=1 To iRowCnt
			crrRwData=objTable.GetCellData(iCounter,1)
			If Trim(crrRwData)=Trim("ID") Or Trim(crrRwData)=Trim("ID:") or instr(1,crrRwData,"ID") Then
				
				Set objEdit=objTable.ChildItem(iCounter,2,"WebEdit",0)
				If TypeName(objEdit)<>"Nothing" Then
					wait 3
'					call Fn_Web_UI_WebEdit_SetExt("Fn_Web_SaveAsObject", "SendString",objEdit,"", dicSaveAsObject("ID"))
					objEdit.Set ""
					objEdit.Object.focus
					objMDR.SendString dicSaveAsObject("ID")
					wait 3
					Exit For
				End If
				Set objEdit=Nothing
			End If
		Next
		crrRwData=""
		If bFlag=False  and bFlag1=False Then
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",crrTabName)
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").Click 1,1,micLeftBtn
			wait(2)
			Set objTable=objSaveAsRev
		End If
	End If
	
	If dicSaveAsObject("RevID")<>"" Then
		If bFlag=True  and bFlag1=False Then
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",".*Revision Information")
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").Click 1,1,micLeftBtn
			wait(2)
			Set objTable=objSaveAsRev
		End If

		For iCounter=1 To iRowCnt
			crrRwData=objTable.GetCellData(iCounter,1)
			If Trim(crrRwData)=Trim("Revision") Or Trim(crrRwData)=Trim("Revision:") Then
				Set objEdit=objTable.ChildItem(iCounter,2,"WebEdit",0)
				If TypeName(objEdit)<>"Nothing" Then
					wait 3
					objEdit.Set ""
					objEdit.Object.focus
					objMDR.SendString dicSaveAsObject("RevID")

					wait 1
					Exit For
				End If
				Set objEdit=Nothing
			End If
		Next
		crrRwData=""
		If bFlag=True Then
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",crrTabName)
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").Click 1,1,micLeftBtn
			wait(2)
			Set objTable=objSaveAs
		End if
	End If
	
	If dicSaveAsObject("Name")<>"" Then
		For iCounter=1 To iRowCnt
			crrRwData=objTable.GetCellData(iCounter,1)
			If Trim(crrRwData)=Trim("Name") Or Trim(crrRwData)=Trim("Name:") Or Trim(crrRwData)=Trim("Name*") Or Trim(crrRwData)=Trim("Name*:") Or Trim(crrRwData)=Trim("Vendor Part Name") Or Trim(crrRwData)=Trim("Vendor Part Name:") Or Trim(crrRwData)=Trim("Vendor Part Name*") Then
				Set objEdit=objTable.ChildItem(iCounter,2,"WebEdit",0)
				If TypeName(objEdit)<>"Nothing" Then
					wait 1
					objEdit.Set ""
					wait 1
					objEdit.Object.focus
					'objEdit.Set dicSaveAsObject("Name")
                    objMDR.SendString dicSaveAsObject("Name")
                    wait 1
					Exit For
				End If
				Set objEdit=Nothing
			End If
		Next
		crrRwData=""
	End If
	If dicSaveAsObject("Description")<>"" Then
		For iCounter=1 To iRowCnt
			crrRwData=objTable.GetCellData(iCounter,1)
			If Trim(crrRwData)=Trim("Description") Or Trim(crrRwData)=Trim("Description:") Then
				Set objEdit=objTable.ChildItem(iCounter,2,"WebEdit",0)
				If TypeName(objEdit)<>"Nothing" Then
					wait 1
					objEdit.Set ""
					objEdit.Object.focus
                    objMDR.SendString dicSaveAsObject("Description")
					wait 1
					Exit For
				End If
				Set objEdit=Nothing
			End If
		Next
		crrRwData=""
	End If
	
	If isObject(dicSaveAsObject("DicRevProp")) Then
		Set dicSaveAsRev = dicSaveAsObject("DicRevProp")
		For iCount = 1 To 3
			Call Fn_Web_UI_Button_Click("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Next")
		Next

			bFlag=False
			bFlag1=False
			If objSaveAs.Exist(5) Then
				If objSaveAs.GetROProperty("height")>0 Then
					Set objTable=objSaveAs
					bFlag=True
				ElseIf objSaveAsRev.GetROProperty("height")>0 Then
					Set objTable=objSaveAsRev
				End If
			ElseIf objSaveAsRev.Exist(5) Then
				If objSaveAsRev.GetROProperty("height")>0 Then
					Set objTable=objSaveAsRev
				End If
			Elseif objSaveAsObject.Exist(5) then
				If objSaveAsObject.GetROProperty("height")>0 Then
					Set objTable=objSaveAsObject
					bFlag1=True
				End If
			End If

			iRowCnt=objTable.RowCount
			If bFlag1=True Then
				'Do nothing
			Else
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",".*Information")
				crrTabName=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").GetROProperty("innertext")
			End If
		
			If dicSaveAsRev("ID")<>"" Then
				If bFlag=False and bFlag1=False Then
					For i=1 to 10
						Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").SetTOProperty "index",i	
						Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",".*Information")
						StrTabName=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").GetROProperty("innertext")
						If instr(1,StrTabName,"Information") Then
							If instr(1,StrTabName,"Revision") or instr(1,StrTabName,"Master") or instr(1,StrTabName,"FullText") Then
								'Do nothing
							Else
								Exit for
							End If
						End If
					Next
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",StrTabName)
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").Click 1,1,micLeftBtn
					wait(2)
					Set objTable=objSaveAs
				End if
				For iCounter=1 To iRowCnt
					crrRwData=objTable.GetCellData(iCounter,1)
					If Trim(crrRwData)=Trim("ID") Or Trim(crrRwData)=Trim("ID:") or instr(1,crrRwData,"ID") Then
						
						Set objEdit=objTable.ChildItem(iCounter,2,"WebEdit",0)
						If TypeName(objEdit)<>"Nothing" Then
							wait 1
		'					objEdit.Set dicSaveAsRev("ID")
							objEdit.Set ""
							objEdit.Object.focus
							objMDR.SendString dicSaveAsRev("ID")
							wait 1
							Exit For
						End If
						Set objEdit=Nothing
					End If
				Next
				crrRwData=""
				If bFlag=False  and bFlag1=False Then
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",crrTabName)
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").Click 1,1,micLeftBtn
					wait(2)
					Set objTable=objSaveAsRev
				End If
			End If
			
			If dicSaveAsRev("RevID")<>"" Then
				If bFlag=True  and bFlag1=False Then
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",".*Revision Information")
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").Click 1,1,micLeftBtn
					wait(2)
					Set objTable=objSaveAsRev
				End If
		
				For iCounter=1 To iRowCnt
					crrRwData=objTable.GetCellData(iCounter,1)
					If Trim(crrRwData)=Trim("Revision") Or Trim(crrRwData)=Trim("Revision:") Then
						Set objEdit=objTable.ChildItem(iCounter,2,"WebEdit",0)
						If TypeName(objEdit)<>"Nothing" Then
							 wait 1
		'					objEdit.Set dicSaveAsRev("RevID")
							objEdit.Set ""
							objEdit.Object.focus
							objMDR.SendString dicSaveAsRev("RevID")
		
							wait 1
							Exit For
						End If
						Set objEdit=Nothing
					End If
				Next
				crrRwData=""
				If bFlag=True Then
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab"),"innertext",crrTabName)
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("SaveAsTabTable").WebElement("SaveAsTab").Click 1,1,micLeftBtn
					wait(2)
					Set objTable=objSaveAs
				End if
			End If
			
			If dicSaveAsRev("Name")<>"" Then
				For iCounter=1 To iRowCnt
					crrRwData=objTable.GetCellData(iCounter,1)
		'			If Trim(crrRwData)=Trim("Name") Or Trim(crrRwData)=Trim("Name:") Or Trim(crrRwData)=Trim("Name*") Or Trim(crrRwData)=Trim("Name*:") Then
					If Instr(1,crrRwData, Trim("Name")) > 0 Then
						Set objEdit=objTable.ChildItem(iCounter,2,"WebEdit",0)
						If TypeName(objEdit)<>"Nothing" Then
							 wait 1
							objEdit.Set dicSaveAsRev("Name")
							wait 1
							Exit For
						End If
						Set objEdit=Nothing
					End If
				Next
				crrRwData=""
			End If
			If dicSaveAsRev("Description")<>"" Then
				For iCounter=1 To iRowCnt
					crrRwData=objTable.GetCellData(iCounter,1)
					If Trim(crrRwData)=Trim("Description") Or Trim(crrRwData)=Trim("Description:") Then
						Set objEdit=objTable.ChildItem(iCounter,2,"WebEdit",0)
						If TypeName(objEdit)<>"Nothing" Then
							wait 1
							'objEdit.Set dicSaveAsRev("Description")
							'Replace By Shrikant
							objEdit.Set ""
							objEdit.Object.focus
							objMDR.SendString dicSaveAsRev("Description")
							wait 1
							Exit For
						End If
						Set objEdit=Nothing
					End If
				Next
				crrRwData=""
			End If
	End If
	
	Call Fn_Web_UI_Button_Click("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Finish")
'	Call Fn_Web_UI_Button_Click("Fn_Web_SaveAsObject",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"OK")

	Fn_Web_SaveAsObject=True
	Set objSaveAs=Nothing
	Set objSaveAsRev=Nothing
	Set objTable=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ChangeOwnership
'@@
'@@    Description				 :	Function Used to Change Ownership of object
'@@
'@@    Parameters			   :	1.strGroup : Select group
'@@												 2.strUser : User from selected group
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log in Web Client							
'@@
'@@    Examples					:	Call Fn_Web_ChangeOwnership("Select","dba","AutoTestDBA")
'@@      										Call Fn_Web_ChangeOwnership("Verify","dba","AutoTestDBA~AutoTest5~AutoTest7")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@													Pranav  Ingle												03-Jun-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ChangeOwnership(strAction,strGroup,strUser)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ChangeOwnership"
	'Initially Function Rerturns False
	Fn_Web_ChangeOwnership=False
	'Variable Declaration
	Dim ObjOwner,objWebChild,strWEBMenuPath,strMenu,iCounter,arrUsers
    
	'Creating Object of "ChangePassword" Table
	Set ObjOwner = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangeOwnership")
	'Checking Existance of "ChangePassword" Table
	If Not ObjOwner.Exist(7) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "EditChangeOwnnership")
		'Calling Edit->Change Password... Menu Option
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If

	Select Case strAction
		Case "Select"
			If strGroup<>"" Then
					iCounter=ObjOwner.GetRowWithCellText("Group")
					Call Fn_Web_UI_Button_Click("Fn_Web_ChangeOwnership", ObjOwner, "Group")
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ChangeOwnership",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeOwnershipGroup"),"innertext",strGroup)
					Call Fn_Web_UI_WebElement_Click("",Browser("TeamcenterWeb").Page("MyTeamCenter"),"ChangeOwnershipGroup","","","")
			End If
			wait(1)
			If strUser<>"" Then
					iCounter=ObjOwner.GetRowWithCellText("User")
					Call Fn_Web_UI_Button_Click("Fn_Web_ChangeOwnership", ObjOwner, "User")
					wait 2
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ChangeOwnership",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeOwnershipGroup"),"innertext",strUser)
					Call Fn_Web_UI_WebElement_Click("",Browser("TeamcenterWeb").Page("MyTeamCenter"),"ChangeOwnershipGroup","","","")
			End If
			'Clicking On OK button
			Call Fn_Web_UI_Button_Click("Fn_Web_ChangeOwnership", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
			Set objWebChild = Nothing

		Case "Verify"
			If strGroup<>"" Then
					iCounter=ObjOwner.GetRowWithCellText("Group")
					Call Fn_Web_UI_Button_Click("Fn_Web_ChangeOwnership", ObjOwner, "Group")
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ChangeOwnership",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeOwnershipGroup"),"innertext",strGroup)
					Call Fn_Web_UI_WebElement_Click("",Browser("TeamcenterWeb").Page("MyTeamCenter"),"ChangeOwnershipGroup","","","")
			End If
			wait(1)
			If strUser<>"" Then
					iCounter=ObjOwner.GetRowWithCellText("User")
					Call Fn_Web_UI_Button_Click("Fn_Web_ChangeOwnership", ObjOwner, "User")
					arrUsers = Split(strUser,"~")
					For iCounter = 0 To UBound(arrUsers)
						Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ChangeOwnership",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeOwnershipGroup"),"innertext",arrUsers(iCounter))
						If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeOwnershipGroup").Exist(5) Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS : Fn_Web_ChangeOwnership: User [ " & arrUsers(iCounter) & " ] is Exist")
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fn_Web_ChangeOwnership: User [ " & arrUsers(iCounter) & " ] is not Exist")
							Exit Function
						End If
					Next
			End If
			'Clicking On OK button
			Call Fn_Web_UI_Button_Click("Fn_Web_ChangeOwnership", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")
			Set objWebChild = Nothing
	End Select

	For iCounter=0 To 2
		If ObjOwner.Exist(5) Then
			wait(5)
		Else
			Exit For
		End If
	Next

	Fn_Web_ChangeOwnership=True	
	Set ObjOwner = Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ChangeManagerTreeOperation
'@@
'@@    Description				 :	Function Used to Perform Operation On Change Manager Tree
'@@
'@@    Parameters			   :	1.strAction : Action Name
'@@											:	 2: strNode Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client	And Change Manager Perspective should be appear						
'@@
'@@    Examples					:	   Call Fn_Web_ChangeManagerTreeOperation("Expand","ECN-000001-Sandeep")
'@@												Call Fn_Web_ChangeManagerTreeOperation("Collapse","ECN-000001-Sandeep:ECN-000001/A;1-Sandeep")
'@@												Call Fn_Web_ChangeManagerTreeOperation("Select","ECN-000001-Sandeep:ECN-000001/A;1-Sandeep:Impacted Items")
'@@												Call Fn_Web_ChangeManagerTreeOperation("MultiSelect","ECN-000001-Sandeep:ECN-000001/A;1-Sandeep:Problem Items~ECN-000001-Sandeep:ECN-000001/A;1-Sandeep:Impacted Items")
'@@												Call Fn_Web_ChangeManagerTreeOperation("Exist","ECN-000001-Sandeep:ECN-000001/A;1-Sandeep:ECN-000001/A")
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									21-Jun-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_Web_ChangeManagerTreeOperation(strAction,strNode)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ChangeManagerTreeOperation"
   Dim rowid,obj,iCounter,arrNode,iLength,objSelectType,objSelectType1,intNoOfObjects,iniWTCount,intNoOfObjects1,finWTCount,iHeight,i,arrNode1
   Fn_Web_ChangeManagerTreeOperation=False
	Select Case strAction
		Case "Expand","Collapse"
			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
			objSelectType("height").value = 0
			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType)
			Set objSelectType1=description.Create()
			objSelectType1("micClass").value = "WebTable"
			Set  intNoOfObjects1 = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType1)
			iniWTCount =  intNoOfObjects1.Count - intNoOfObjects.Count
			arrNode=Split(strNode,":")
			iLength=UBound(arrNode)
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangeTreeNodeName").SetTOProperty "text",arrNode(iLength)
			rowid = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangeTreeNodeName").GetRowWithCellText(arrNode(iLength))
			Set obj = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangeTreeNodeName").ChildItem(rowid,iLength+1,"WebElement",0)
			obj.click
			wait(5)
			Set objSelectType=description.Create()
			objSelectType("micClass").value = "WebTable"
			objSelectType("height").value = 0
			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType)
			Set objSelectType1=description.Create()
			objSelectType1("micClass").value = "WebTable"
			Set  intNoOfObjects1 = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType1)
			 finWTCount =  intNoOfObjects1.Count - intNoOfObjects.Count
			 If strAction = "Expand" Then
				 If finWTCount < iniWTCount Then
					 obj.click
					wait(2)
				 End If
			 Else
				 If finWTCount > iniWTCount Then
					 obj.click
					wait(2)
				 End If
			 End If
			 Set objSelectType = Nothing
			 Set objSelectType1 = Nothing
			 Set obj = Nothing
			 Fn_Web_ChangeManagerTreeOperation=True
		Case "Select"
			arrNode=Split(strNode,":")
			iLength=UBound(arrNode)
			Set objSelectType=description.Create()
			Set objSelectType1=description.Create()
			objSelectType("micClass").value = "WebTable"
			objSelectType("innertext").value =arrNode(iLength)
			objSelectType1("micClass").value = "WebElement"
			objSelectType1("innertext").value = arrNode(iLength)
			Set  intNoOfObjects = Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objSelectType)
			For i=0 to intNoOfObjects.Count-1
					iHeight = intNoOfObjects(i).getroproperty("height")
					If iHeight > 0 Then
						Set intNoOfObjects1 =  intNoOfObjects(i).ChildObjects(objSelectType1)
						intNoOfObjects1(2).Click 1,1
					End If
			Next
			Fn_Web_ChangeManagerTreeOperation=True
		Case "MultiSelect"
			arrNode=Split(strNode,"~")
			arrNode1=Split(arrNode(0),":")
			iLength=UBound(arrNode1)
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ChangeManagerTreeOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeManagerTree"),"innertext",arrNode1(iLength))
			Call Fn_Web_UI_WebElement_Click("Fn_Web_ChangeManagerTreeOperation", Browser("TeamcenterWeb").Page("MyTeamCenter"), "ChangeManagerTree", "","","")
			wait(5)
			For iCounter=1 To UBound(arrNode)
				arrNode1=Split(arrNode(iCounter),":")
				iLength=UBound(arrNode1)
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ChangeManagerTreeOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeManagerTree"),"innertext",arrNode1(iLength))
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeManagerTree").Drag 1,1,micLeftBtn,micCtrl
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeManagerTree").Drop 1,1
				wait(5)
			Next
			Fn_Web_ChangeManagerTreeOperation=True
		Case "Exist"
			arrNode=Split(strNode,":")
			iLength=UBound(arrNode)
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ChangeManagerTreeOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeManagerTree"),"innertext",arrNode(iLength))
			If Fn_Web_UI_ObjectExist("Fn_Web_ChangeManagerTreeOperation", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeManagerTree"))=True Then
				If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ChangeManagerTree").GetROProperty("height") > 0 Then
					Fn_Web_ChangeManagerTreeOperation=True
				End If
			End If
		Case "ClickLink"
			arrNode=Split(strNode,":")
			iLength=UBound(arrNode)
            Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangeTreeNodeName").SetTOProperty "text",arrNode(iLength)
			rowid = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangeTreeNodeName").GetRowWithCellText(arrNode(iLength))
			Set obj = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ChangeTreeNodeName").ChildItem(rowid,iLength+2,"Link",0) 
			wait(1)
			obj.click
			wait(1)
			Fn_Web_ChangeManagerTreeOperation=True

	End Select
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ImpactAnalysisListOperations
'@@
'@@    Description				 :	Function Used To perform Operation On Impact Analysis List
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@											  2.StrOption : Option Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client And Object Need to be Selected On Which have to perform Operations (Eg : Item)							
'@@
'@@    Examples					:	Call Fn_Web_ImpactAnalysisListOperations("Select","Defining Objects")
'@@											 Call Fn_Web_ImpactAnalysisListOperations("Verify","Defining Objects")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									25-Sep-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ImpactAnalysisListOperations(StrAction,StrOption)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ImpactAnalysisListOperations"
   '------- Variable Declaration ------------------------------------------------------------------------------------
	Dim ObjSmryTab,objImpAnalysisTbl
	'------- Creating Objects of Table ------------------------------------------------------------------------------------
	Set ObjSmryTab=Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Overview")
	Set objImpAnalysisTbl=Browser("TeamcenterWeb").Page("MyTeamCenter")

	Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ImpactAnalysisListOperations",ObjSmryTab,"innertext","Impact Analysis")
	Call Fn_Web_UI_WebElement_Click("Fn_Web_ImpactAnalysisListOperations",Browser("TeamcenterWeb").Page("MyTeamCenter"),"Overview", "","","")
	wait 2
	Select Case StrAction
		'------- Case to Select Impact Analysis Option---------------------------------------------------------------------------------
		Case "Select"
			If  Fn_Web_ImpactAnalysisListOperations("Verify",StrOption)=True Then
				Call Fn_Web_UI_Button_Click("Fn_Web_ImpactAnalysisListOperations", objImpAnalysisTbl, "ImpactAnalysis")
				wait 0,200
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ImpactAnalysisListOperations",objImpAnalysisTbl.WebElement("AnalysisElement"),"innertext",StrOption)
				Call Fn_Web_UI_WebElement_Click("Fn_Web_ImpactAnalysisListOperations",objImpAnalysisTbl,"AnalysisElement", "","","")
				Fn_Web_ImpactAnalysisListOperations=True
			Else
				Fn_Web_ImpactAnalysisListOperations=False
			End If
		'------- Case to Verify Impact Analysis Option---------------------------------------------------------------------------------
		Case "Verify"
			Call Fn_Web_UI_Button_Click("Fn_Web_ImpactAnalysisListOperations", objImpAnalysisTbl, "ImpactAnalysis")
			wait 1
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ImpactAnalysisListOperations",objImpAnalysisTbl.WebElement("AnalysisElement"),"innertext",StrOption)
			If Fn_Web_UI_ObjectExist("Fn_Web_ImpactAnalysisListOperations", objImpAnalysisTbl.WebElement("AnalysisElement"))=True Then
				Fn_Web_ImpactAnalysisListOperations=True
			Else
				Fn_Web_ImpactAnalysisListOperations=False
			End If
	End Select
	Set ObjSmryTab=Nothing
	Set objImpAnalysisTbl=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ImpactAnalysisTreeOperations
'@@
'@@    Description				 :	Function Used To perform Operation On Impact Analysis Tree
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@											  2.sNodeName : sNodeName Name
'@@											  3.sColName:Column Name [ First node is column name ]
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client						
'@@
'@@    Examples					:	Call Fn_Web_ImpactAnalysisTreeOperations("Expand","000116-Item1","000116-Item1")
'@@											 Call Fn_Web_ImpactAnalysisTreeOperations("Verify","000116-Item1:Sandeep23456","000116-Item1")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									25-Sep-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_ImpactAnalysisTreeOperations(StrAction,sNodeName,sColName)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ImpactAnalysisTreeOperations"

	Dim objDialog
	Dim iRowCnt,iColPos,objImg,iColCount,iCounter,arrNode

    Set objDialog=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ImpactAnalysisTreeTable")

	Select Case StrAction

		Case "Expand"
					iRowCnt = objDialog.RowCount
					iColCount = objDialog.ColumnCount(1)
					For iCounter = 1 to iColCount
							If  sColName = objDialog.GetCellData(1,iCounter) Then
										iColPos = iCounter
										Exit for
							End If
					Next
					If iRowCnt <> -1 Then
							Set objImg = objDialog.ChildItem(iRowCnt, iColPos, "Image", 0)
							If TypeName(objImg) <> "Nothing" Then
									If objImg.GetROProperty("file name") = "plus.png" Then
												objImg.Click 1,1, micLeftBtn
												bFlag = True
									elseIf objImg.GetROProperty("file name") = "minus.png" Then
												bFlag = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_ImpactAnalysisTreeOperations: node  ["+CStr(sNodeName)+"] was already expanded.")
									else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_ImpactAnalysisTreeOperations: can not expand node  ["+CStr(sNodeName)+"].")
									End If
							End If						
							Set objImg = Nothing
					else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_ImpactAnalysisTreeOperations: node  ["+CStr(sNodeName)+"] does not exist in BOM table.")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_Web_ImpactAnalysisTreeOperations = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_ImpactAnalysisTreeOperations : Node  ["+CStr(sNodeName)+"] expanded successfully. ")
					Else
						Fn_Web_ImpactAnalysisTreeOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_ImpactAnalysisTreeOperations : Failed to expand node  ["+CStr(sNodeName)+"]")
					End If
		Case "Verify"
					arrNode=Split(sNodeName,":")
                    objDialog.Link("TreeLinks").SetTOProperty "Text",arrNode(UBound(arrNode))
					wait 2
					If  objDialog.Link("TreeLinks").Exist(5) Then
						Fn_Web_ImpactAnalysisTreeOperations = True
					Else
						Fn_Web_ImpactAnalysisTreeOperations = False
					End If
	End Select
	Set objDialog=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CreateNewParagraph
'@@
'@@    Description				 :	Function Used to Create Basic Paragraph
'@@
'@@    Parameters			   :        1.strID : Paragraph ID
'@@												  2.strRev : Paragraph Revision														
'@@												  3.strName : Paragraph Name
'@@												  4.strDesc : Paragraph Description
'@@												  5.UOM : Unit Of Measure
'@@												  6.AltIDOpt : Create Alternate ID Option
'@@												  7.CheckOutOpt : Check-Out Item Revision on Create Option
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client							
'@@
'@@    Examples					:	Call Fn_Web_CreateNewParagraph("009218","A","Para1","New Para","","","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									27-Sep-2011						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_CreateNewParagraph(strID,strRev,strName,strDesc,UOM,AltIDOpt,CheckOutOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CreateNewParagraph"
   Dim ObjItem,strWEBMenuPath,strMenu,iCounter,crrType
	Dim objElement, intIndex

	Fn_Web_CreateNewParagraph=False

	Set objElement = Description.Create()
	objElement("micclass").Value = "WebElement"
	objElement("innertext").Value = "New Paragraph"
	objElement("html tag").Value = "SPAN"
	intIndex =  Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objElement).count
	Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewParagraph").SetTOProperty "index", cstr(intIndex)
	
	If  Not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewParagraph").Exist(7) Then
	'If New Item does not exist, Do menu operation
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewParagraph")
		Call Fn_Web_MenuOperation("Select",strMenu)

		'Vallari [14Jun11] - Get the number of intances of New Item dialog and set the index for WebTable in OR accordingly
		intIndex =  Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objElement).count
		Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewParagraph").SetTOProperty "index", cstr(cint(intIndex)-1)
	End If
	
	Set objElement = Nothing
	Set ObjItem=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewParagraph")

	If ObjItem.WebElement("ListName").Exist(5) Then
        crrType=ObjItem.WebEdit("ParagraphType").GetROProperty("value")
        wait(1)
		If Trim(crrType)<>Trim("Paragraph") Then
				'Setting Item Type
				Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewParagraph",ObjItem,"ParagraphType")
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ItemBasicCreate",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext","Paragraph")
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
		End If
		Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewParagraph", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	End If

	If strID<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateNewParagraph", ObjItem.WebTable("ParagraphInfo"), "ID", strID)
	End If
	If strRev<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateNewParagraph", ObjItem.WebTable("ParagraphInfo"), "Revision", strRev)
	End If
	If strName<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateNewParagraph", ObjItem.WebTable("ParagraphInfo"), "Name", strName)
	End If
	If strDesc<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateNewParagraph", ObjItem.WebTable("ParagraphInfo"), "Description", strDesc)
	End If
	If UOM<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateNewParagraph", ObjItem.WebTable("ParagraphInfo"), "UOM", UOM)
	End If
	If AltIDOpt<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_Web_CreateNewParagraph", ObjItem.WebTable("ParagraphInfo"), "CreateALTID", AltIDOpt)
	End If
	If CheckOutOpt<>"" Then
		Call Fn_Web_UI_CheckBox_Set("Fn_Web_CreateNewParagraph", ObjItem.WebTable("ParagraphInfo"), "CheckOutRevision", CheckOutOpt)
	End If
	Call Fn_Web_UI_Button_Click("Fn_Web_CreateNewParagraph", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	For iCounter=0 To 2
		If Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewParagraph").Exist(5) Then
			wait(5)
		Else
			Exit For
		End If
	Next
	
	Fn_Web_CreateNewParagraph=True
	Set ObjItem=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CreateClassicChange
'@@
'@@    Description				 :	Function Used to Create Classic Change
'@@
'@@    Parameters			   :	1.StrChangeType : Change Type
'@@												 2.StrChangeID : Change ID
'@@												 3.StrChangeRevID : Change Rev ID.
'@@												 4.StrChangeName : Change Name
'@@												 4.StrChangeDesc : Change Description
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Should be Log In Web Client.
'@@
'@@    Examples					:	Call Fn_Web_CreateClassicChange("CN","CN0079","A","Change1","ChangeTest")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sachin Joshi									          15-Nov-2011						      1.0																						Rupali Palhade
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_CreateClassicChange(StrChangeType,StrChangeID,StrChangeRevID,StrChangeName,StrChangeDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CreateClassicChange"
   'Variable Declaration
	Dim objChange
	Dim strWEBMenuPath,strMenu,StrCrrFilter,StrCrrTemp,bFlag,iCounter
	'Creating Object "New Classic Change" Dialog
	Set objChange=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("Create/Revise Change")
	Fn_Web_CreateClassicChange=False
	bFlag=False
	strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
	strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "ClassicChangeCreate")
	'Checking Existance Of "Classic Change" Dialog
	If objChange.Exist(5) Then
		'Calling "New--> Classic Change" Menu Option
		If objChange.GetROProperty("height")=0 Then
			Call Fn_Web_MenuOperation("Select",strMenu)
			wait(3)
			objChange.SetTOProperty "index",1
		End If
	Else
		Call Fn_Web_MenuOperation("Select",strMenu)
		wait(3)
	End If
	'Setting Classic Change Type
	If StrChangeType<>"" Then
			Call Fn_Web_UI_Button_Click("Fn_Web_CreateClassicChange",objChange,"ChangeType")
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CreateClassicChange",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",StrChangeType)
			Call Fn_Web_UI_WebElement_Click("Fn_Web_CreateClassicChange",Browser("TeamcenterWeb").Page("MyTeamCenter"),"FormType","","","")
			wait(1)
	End If
	'Setting Classic Change ID
	If StrChangeID<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateClassicChange",objChange,"ChangeID",StrChangeID)
	End If
	'Setting Classic Change Rev ID
	If StrChangeRevID<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateClassicChange",objChange,"ChangeRevID",StrChangeRevID)
	End If
	'Setting Classic Change Name
	If StrChangeName<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateClassicChange",objChange,"ChangeName",StrChangeName)
	End If
	'Setting Classic Change Description
	If StrChangeDesc<>"" Then
		Call Fn_Web_UI_WebEdit_Set("Fn_Web_CreateClassicChange",objChange,"ChangeDescription",StrChangeDesc)
	End If
	'Click on OK Button
    Call Fn_Web_UI_Button_Click("Fn_Web_CreateClassicChange", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"OK")
    Fn_Web_CreateClassicChange=True
	
	For iCounter=0 To 2
		If objChange.Exist(5) Then
			wait(5)
		Else
			Exit For
		End If
	Next
	'Releasing Object "objChange" Dialog
	Set objChange=Nothing
End Function
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_ExportToExcel
'@@
'@@    Description				 :	Function Export Object in Excel
'@@
'@@    Parameters			   :	1.StrExcelTemplate : Excel Tamplate Type
'@@												 2.StrOutputOption : Output Option [ Live , Static ]
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Object which have to export to Excel Should be Selected
'@@
'@@    Examples					:	Call Fn_Web_ExportToExcel("REQ_TraceLink_complying_template","Live")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									          17-Nov-2011						      1.0																						Rupali Palhade
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_Web_ExportToExcel(StrExcelTemplate,StrOutputOption)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_ExportToExcel"
   Dim strWEBMenuPath,strMenu,crrTemplateType,ObjTemplate
   Fn_Web_ExportToExcel=False
	Set ObjTemplate=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("ExcelTemplate")
	If Not ObjTemplate.Exist(6) Then
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "ExportToExcel")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	'Selecting Excel Template to Export
	If StrExcelTemplate<>"" Then
		crrTemplateType=ObjTemplate.WebEdit("TemplateTypeEdit").GetROProperty("value")
		If Trim(crrTemplateType)<>Trim(StrExcelTemplate) Then
			wait(1)
			Call Fn_Web_UI_Button_Click("Fn_Web_ExportToExcel",ObjTemplate,"ExcelTemplateType")
			wait(2)
			Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_ExportToExcel",ObjTemplate.WebElement("TemplateType"),"innertext",StrExcelTemplate)
			ObjTemplate.WebElement("TemplateType").Click 1,1,micLeftBtn
			wait(2)
		End If
	End If
	'Selecting Output Option
	If StrOutputOption<>"" Then
		'StrOutputOption : = Pass the value eg:- Live Or Static ect.....
		ObjTemplate.WebRadioGroup("OutputOption").Select StrOutputOption
	End If
	Call Fn_Web_UI_Button_Click("Fn_Web_ExportToExcel", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
	Fn_Web_ExportToExcel=True
	Set ObjTemplate=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Web_AssignProjectsOperations

'Description			 :	Function Used to Perform operations on Assign Projects

'Parameters			   :   '1.StrAction: Action Name
'										 2.dicProjectInfo: Assign Project Information
'
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be log in Thin Client and Object should be selected whom project has to assign

'Examples				:   dicProjectInfo("AvailableProjects")="Projects A~Projects B"
'										Msgbox Fn_Web_AssignProjectsOperations("Add",dicProjectInfo)
'										dicProjectInfo("SelectedProjetcs")="Projects A~Projects B"
'										Msgbox Fn_Web_AssignProjectsOperations("Remove",dicProjectInfo)
'										dicProjectInfo("AvailableProjects")="Projects A~Projects B"
'										Msgbox Fn_Web_AssignProjectsOperations("VerifyAvailableProjects",dicProjectInfo)
'										dicProjectInfo("SelectedProjetcs")="Projects A~Projects B"
'										Msgbox Fn_Web_AssignProjectsOperations("VerifySelectedProjects",dicProjectInfo)

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												24-Jan-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Web_AssignProjectsOperations(StrAction,dicProjectInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_AssignProjectsOperations"
 	'Variable Declaration
    Dim strWEBMenuPath,strMenu,ObjAssignProjects
	Dim arrProjects,iCounter,iCounter1,iCount,bFlag,crrProject
	'Function Returns False
    Fn_Web_AssignProjectsOperations=False
	'Creating Object of [ AssignProjects ] table
	Set ObjAssignProjects=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AssignProjects")
	'Checking Existance of [ AssignProjects ] table
	If Not ObjAssignProjects.Exist(6) Then
		'Calling menu : "Tools:Project:Assign"
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "AssignProjects")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	Select Case StrAction
		'Case To Assign project to Object
		Case "Add"
			arrProjects=Split(dicProjectInfo("AvailableProjects"),"~")
			For iCounter=0 To UBound(arrProjects)
				Call Fn_Web_UI_List_Select("Fn_Web_AssignProjectsOperations",ObjAssignProjects, "AvailabeProjects",arrProjects(iCounter))
				Call Fn_Web_UI_Button_Click("Fn_Web_AssignProjectsOperations", ObjAssignProjects, "Add")
			Next
			Call Fn_Web_UI_Button_Click("Fn_Web_AssignProjectsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
			Fn_Web_AssignProjectsOperations=True

		'Case To Remove assign projects from Object
		Case "Remove"
			arrProjects=Split(dicProjectInfo("SelectedProjetcs"),"~")
			For iCounter=0 To UBound(arrProjects)
				Call Fn_Web_UI_List_Select("Fn_Web_AssignProjectsOperations",ObjAssignProjects, "SelectedProjects",arrProjects(iCounter))
				Call Fn_Web_UI_Button_Click("Fn_Web_AssignProjectsOperations", ObjAssignProjects, "Remove")
			Next
			Call Fn_Web_UI_Button_Click("Fn_Web_AssignProjectsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
			Fn_Web_AssignProjectsOperations=True

		'Case to Verify Avaliable Projects
		Case "VerifyAvailableProjects"
			arrProjects=Split(dicProjectInfo("AvailableProjects"),"~")
			iCounter=ObjAssignProjects.WebList("AvailabeProjects").GetROProperty("items count")
			For iCounter1=0 To UBound(arrProjects)
				bFlag=False
				For iCount=1 To iCounter
					crrProject=ObjAssignProjects.WebList("AvailabeProjects").GetItem(iCount)
					If Trim(crrProject)=arrProjects(iCounter1) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_Web_AssignProjectsOperations=True
			End If
			Call Fn_Web_UI_Button_Click("Fn_Web_AssignProjectsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")

		'Case to Verify Selected Projects
		Case "VerifySelectedProjects"
			arrProjects=Split(dicProjectInfo("SelectedProjetcs"),"~")
			iCounter=ObjAssignProjects.WebList("SelectedProjects").GetROProperty("items count")
			For iCounter1=0 To UBound(arrProjects)
				bFlag=False
				For iCount=1 To iCounter
					crrProject=ObjAssignProjects.WebList("SelectedProjects").GetItem(iCount)
					If Trim(crrProject)=arrProjects(iCounter1) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_Web_AssignProjectsOperations=True
			End If
			Call Fn_Web_UI_Button_Click("Fn_Web_AssignProjectsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")

	End Select
	'Releasing Object of [ AssignProjects ] table
	Set ObjAssignProjects=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Web_RemoveProjectsOperations

'Description			 :	Function Used to Perform operations on Remove Projects

'Parameters			   :   '1.StrAction: Action Name
'										 2.dicRemoveProjectInfo: Assign Project Information
'
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be log in Thin Client and Object should be selected whom project has to remove

'Examples				:   dicRemoveProjectInfo("AvailableProjects")="PLSProj_2_12201214329~PLSProj_1_12201214131"
'										Msgbox Fn_Web_RemoveProjectsOperations("Add",dicRemoveProjectInfo)
'										dicRemoveProjectInfo("SelectedProjetcs")="PLSProj_2_12201214329~PLSProj_1_12201214131"
'										Msgbox Fn_Web_RemoveProjectsOperations("Remove",dicRemoveProjectInfo)
'										dicRemoveProjectInfo("AvailableProjects")="PLSProj_2_12201214329~PLSProj_1_12201214131"
'										Msgbox Fn_Web_RemoveProjectsOperations("VerifyAvailableProjects",dicRemoveProjectInfo)
'										dicRemoveProjectInfo("SelectedProjetcs")="PLSProj_2_12201214329~PLSProj_1_12201214131"
'										Msgbox Fn_Web_RemoveProjectsOperations("VerifySelectedProjects",dicRemoveProjectInfo)

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												01-Feb-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Web_RemoveProjectsOperations(StrAction,dicRemoveProjectInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_RemoveProjectsOperations"
 	'Variable Declaration
    Dim strWEBMenuPath,strMenu,ObjRemoveProject
	Dim arrProjects,iCounter,iCounter1,iCount,bFlag,crrProject
	'Function Returns False
    Fn_Web_RemoveProjectsOperations=False
	'Creating Object of [ RemoveProject ] table
	Set ObjRemoveProject=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("RemoveProject")
	'Checking Existance of [ RemoveProject ] table
	If Not ObjRemoveProject.Exist(6) Then
		'Calling menu : "Tools:Project:Assign"
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "RemoveProject")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	Select Case StrAction
		'Case To Assign project to Object
		Case "Add"
			arrProjects=Split(dicRemoveProjectInfo("AvailableProjects"),"~")
			For iCounter=0 To UBound(arrProjects)
				Call Fn_Web_UI_List_Select("Fn_Web_RemoveProjectsOperations",ObjRemoveProject, "AvailabeProjects",arrProjects(iCounter))
				Call Fn_Web_UI_Button_Click("Fn_Web_RemoveProjectsOperations", ObjRemoveProject, "Add")
			Next
			Call Fn_Web_UI_Button_Click("Fn_Web_RemoveProjectsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
			Fn_Web_RemoveProjectsOperations=True

		'Case To Remove assign projects from Object
		Case "Remove"
			arrProjects=Split(dicRemoveProjectInfo("SelectedProjetcs"),"~")
			For iCounter=0 To UBound(arrProjects)
				Call Fn_Web_UI_List_Select("Fn_Web_RemoveProjectsOperations",ObjRemoveProject, "SelectedProjects",arrProjects(iCounter))
				Call Fn_Web_UI_Button_Click("Fn_Web_RemoveProjectsOperations", ObjRemoveProject, "Remove")
			Next
			Call Fn_Web_UI_Button_Click("Fn_Web_RemoveProjectsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
			Fn_Web_RemoveProjectsOperations=True

		'Case to Verify Avaliable Projects
		Case "VerifyAvailableProjects"
			arrProjects=Split(dicRemoveProjectInfo("AvailableProjects"),"~")
			iCounter=ObjRemoveProject.WebList("AvailabeProjects").GetROProperty("items count")
			For iCounter1=0 To UBound(arrProjects)
				bFlag=False
				For iCount=1 To iCounter
					crrProject=ObjRemoveProject.WebList("AvailabeProjects").GetItem(iCount)
					If Trim(crrProject)=arrProjects(iCounter1) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_Web_RemoveProjectsOperations=True
			End If
			Call Fn_Web_UI_Button_Click("Fn_Web_RemoveProjectsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")

		'Case to Verify Selected Projects
		Case "VerifySelectedProjects"
			arrProjects=Split(dicRemoveProjectInfo("SelectedProjetcs"),"~")
			iCounter=ObjRemoveProject.WebList("SelectedProjects").GetROProperty("items count")
			For iCounter1=0 To UBound(arrProjects)
				bFlag=False
				For iCount=1 To iCounter
					crrProject=ObjRemoveProject.WebList("SelectedProjects").GetItem(iCount)
					If Trim(crrProject)=arrProjects(iCounter1) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_Web_RemoveProjectsOperations=True
			End If
			Call Fn_Web_UI_Button_Click("Fn_Web_RemoveProjectsOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Cancel")

	End Select
	'Releasing Object of [ RemoveProject ] table
	Set ObjRemoveProject=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Web_BusinessObjectOperations

'Description			 :	Function Used to Create Business Object in Thin client

'Parameters			   :   '1.StrObjectName: Business Object  Name
'										 2.dicWebBOInfo: Business Object information 
'
'Return Value		   : 	True Or False

'Pre-requisite			:	Select Unique Item Dialog Should be present

'Examples				:  	dicWebBOInfo("ID")="123345"
'										dicWebBOInfo("Revision")="B"
'										dicWebBOInfo("Name")="Test"
'										dicWebBOInfo("Description")="Test Desc~Next"
'										dicWebBOInfo("MFK Key1")="Key1"
'										dicWebBOInfo("MFK Key2")="Key2"
'										bReturn=Fn_Web_BusinessObjectOperations("MFK_Calibration",dicWebBOInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												06-Feb-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Web_BusinessObjectOperations(StrObjectName,dicWebBOInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_BusinessObjectOperations"
 	'Declaring Variables
	Dim strWEBMenuPath,strMenu,crrType,iTemp,ObjBOInfo,ObjBO,arrKeys,iCounter,arrItem,iRwCount,iCount,sFieldName,objWebEle,bFlag
	Dim objobjMDR
	Fn_Web_BusinessObjectOperations=False
	'Checking Existance of [ New Business Object ] dialog
	If  Not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewBusinessObject").Exist(7) Then
	'If New Business Object dialog does not exist, Do menu operation
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewBusinessObject")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If	
	'Creating Object of [ New Business Object ] dialog
	Set ObjBO=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewBusinessObject")
	'Selecting [ Business Object ] type
	If StrObjectName<>"" Then
		crrType=ObjBO.WebEdit("BOTypeEdit").GetROProperty("value")
		wait(1)
		If Trim(crrType)<>Trim(StrObjectName) Then
				'Setting Business Object Type
				Call Fn_Web_UI_Button_Click("Fn_Web_BusinessObjectOperations",ObjBO,"BOType")
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_BusinessObjectOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",StrObjectName)
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
				wait(2)
		End If
	End If
	'Clicking Next button to go to [Business Object ] page
	Call Fn_Web_UI_Button_Click("Fn_Web_BusinessObjectOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	'Declaring Temp counter
	iTemp=0
	arrKeys=dicWebBOInfo.Keys
	For iCounter=0 to dicWebBOInfo.Count-1
		'Creating Table object
		Select Case CInt(iTemp)
			Case 0
					Set ObjBOInfo=ObjBO.WebTable("BOInfo")
			Case 1
					Set ObjBOInfo=ObjBO.WebTable("AdditionalBOInfo")
			Case 2
					Set ObjBOInfo=ObjBO.WebTable("BORevInfo")
		End Select

		'creating object of Mercury device replay
		Set objobjMDR = CreateObject("Mercury.DeviceReplay")

		bFlag=True
		'Cases to fill the [ Business Object ] Fileds
		Select Case arrKeys(iCounter)
			'to enter "ID","Name","Revision","MFK Key1","MFK Key2","Description"
			Case "ID","Name","Revision","MFK Key1","MFK Key2","Description"
				
				If dicWebBOInfo(arrKeys(iCounter))<>"" Then
					arrItem=Split(dicWebBOInfo(arrKeys(iCounter)),"~")
					bFlag=False
					'Taking Field name from Table
					iRwCount=ObjBOInfo.RowCount()
					For iCount=0 To iRwCount
						sFieldName=ObjBOInfo.GetCellData(iCount,1)
						If Trim(sFieldName)=arrKeys(iCounter)+":" Then
							Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebEdit", 0)
							bFlag=True
							Exit For
						ElseIF Trim(sFieldName)=arrKeys(iCounter)+"*:" Then
							Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebEdit", 0)
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=False Then
						Set ObjBO=Nothing
						Set ObjBOInfo=Nothing
						Exit Function
					End If
					If TypeName(objWebEle) <> "Nothing" Then
						If arrKeys(iCounter) = "ID" Then
							objWebEle.Object.focus
							objobjMDR.SendString arrItem(0)
						Else	
							objWebEle.Object.focus
							objobjMDR.SendString arrItem(0)
'							objWebEle.Set arrItem(0)
						End If
					End If

                    If UBound(arrItem)=1 Then
						Call Fn_Web_UI_Button_Click("Fn_Web_BusinessObjectOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
						iTemp=iTemp+1
					End If
				End If
			'to select "CreateAlternateID","CheckOutItemRevisionOnCreate" options
			Case "CreateAlternateID","CheckOutItemRevisionOnCreate"
				If dicWebBOInfo(arrKeys(iCounter))<>"" Then
					arrItem=Split(dicWebBOInfo(arrKeys(iCounter)),"~")
					bFlag=False
					iRwCount=ObjBOInfo.RowCount()
					For iCount=0 To iRwCount
						sFieldName=ObjBOInfo.GetCellData(iCount,1)
						If Trim(sFieldName)=arrKeys(iCounter)+":" Then
							Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebEdit", 0)
							bFlag=True
							Exit For
						ElseIF Trim(sFieldName)=arrKeys(iCounter)+"*:" Then
							Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebCheckBox", 0)
							bFlag=True
							Exit For
						End If
					Next

					If bFlag=False Then
						Set ObjBO=Nothing
						Set ObjBOInfo=Nothing
						Exit Function
					End If

					If TypeName(objWebEle) <> "Nothing" Then
						If objWebEle.GetROProperty("checked") = "0" and LCase(arrItem(0))="on" Then
							objWebEle.Click 1, 1, micLeftBtn
						ElseIf objWebEle.GetROProperty("checked") = "1" and LCase(arrItem(0))="off" Then
							objWebEle.Click 1, 1, micLeftBtn
						End If
					End If
                    If UBound(arrItem)=1 Then
						Call Fn_Web_UI_Button_Click("Fn_Web_BusinessObjectOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
						iTemp=iTemp+1
					End If
				End If
		End Select
	Next
	'clicking on finish button to create New BO
	Call Fn_Web_UI_Button_Click("Fn_Web_BusinessObjectOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	For iCounter=0 To 2
		If ObjBO.Exist(5) Then
			wait(5)
		Else
			Exit For
		End If
	Next
	Fn_Web_BusinessObjectOperations=True
	Set ObjBO=Nothing
	Set ObjBOInfo=Nothing
	Set objWebEle =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name			:	Fn_Web_PSConnectionCreate

'Description			:	Function Used to Create New PS Connection on thin client

'Parameters			   	:   '1. dicWebBOInfo: Business Object information 
'
'Return Value		   	: 	True Or False

'Pre-requisite			:	Select Unique Item Dialog Should be present

'Examples				:  	dicWebBOInfo("ID")="123345"
'							dicWebBOInfo("Revision")="B"
'							dicWebBOInfo("Name")="Test"
'							dicWebBOInfo("Description")="Test Desc"
'							dicWebBOInfo("MFK Key1")="Key1"
'							dicWebBOInfo("MFK Key2")="Key2"
'							bReturn = Fn_Web_PSConnectionCreate(dicWebBOInfo)
'
'History					 :			
'				Developer Name			Date				Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Koustubh W				07-Feb-2012			1.0
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Web_PSConnectionCreate(dicWebBOInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_PSConnectionCreate"
	Dim objPSTable, objButtonPanel, objEdit, sMenu, bReturn, objobjMDR
	Set objPSTable = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewPSConnection")
	Set objButtonPanel = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel")
	Set objobjMDR = CreateObject("Mercury.DeviceReplay")
	Fn_Web_PSConnectionCreate = False

	If objPSTable.Exist(5) = False Then
	   ' perform menu operation
       'sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WEB_Menu"), "NewConnectionRevisable")
	   strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewConnectionRevisable")
       	bReturn = Fn_Web_MenuOperation("Select", strMenu)
		If bReturn = False OR objPSTable.Exist(15) = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_Web_PSConnectionCreate : Failed to perform menu operation [ "& sMenu & " ].")	
			Exit function
		End If
	End If
	' selecting Type
	'Set objPSTable =  Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewPSConnection").WebTable("Type")
	If dicWebBOInfo("Type") <> "" Then
		objPSTable.WebEdit("Type_Editbox").Set dicWebBOInfo("Type")
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ Type = " & dicWebBOInfo("Type") & " ].")
	End If
	call Fn_KeyBoardOperation("SendKey", "{TAB}")
	
	' clicking on Next
	objButtonPanel.WebButton("Next").Click 1,1,micLeftBtn
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully Clicked on [ Next ] button.")

	Set objPSTable =  Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewPSConnection").WebTable("MFK_ConnInfo")
	' setting ID
	If dicWebBOInfo("ID") <> "" Then
			Set objEdit = objPSTable.ChildItem(2, 2,"WebEdit", 0)
			If TypeName(objEdit) <> "Nothing" Then
					'objEdit.set dicWebBOInfo("ID")
					objEdit.Object.focus
				    objobjMDR.SendString dicWebBOInfo("ID")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ ID = " & dicWebBOInfo("ID") & " ].")
			End If
	End If
	If  dicWebBOInfo("Revision") <> "" Then
		Set objEdit = objPSTable.ChildItem(3, 2,"WebEdit", 0)
		If TypeName(objEdit) <> "Nothing" Then
				'objEdit.set dicWebBOInfo("Revision")
				objEdit.Object.focus
				objobjMDR.SendString dicWebBOInfo("Revision")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ Revision = " & dicWebBOInfo("Revision") & " ].")
		End If
	End If
	If  dicWebBOInfo("Name") <> "" Then
		Set objEdit = objPSTable.ChildItem(4, 2,"WebEdit", 0)
		If TypeName(objEdit) <> "Nothing" Then
				objEdit.set dicWebBOInfo("Name")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ Name = " & dicWebBOInfo("Name") & " ].")
		End If
	End If
	If  dicWebBOInfo("Description") <> "" Then
		Set objEdit = objPSTable.ChildItem( 5, 2, "WebEdit", 0)
		If TypeName(objEdit) <> "Nothing" Then
				objEdit.set dicWebBOInfo("Description")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ Description = " & dicWebBOInfo("Description") & " ].")
		End If
	End If
	If  dicWebBOInfo("UOM") <> "" Then
		Set objEdit = objPSTable.ChildItem( 6, 2, "WebEdit", 0)
		If TypeName(objEdit) <> "Nothing" Then
				objEdit.set dicWebBOInfo("UOM")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ UOM = " & dicWebBOInfo("UOM") & " ].")
		End If
	End If
	If  dicWebBOInfo("CreateAlternateID") <> "" Then
		Set objEdit = objPSTable.ChildItem( 7, 2, "WebCheckBox", 0)
		If TypeName(objEdit) <> "Nothing" Then
				objEdit.set dicWebBOInfo("CreateAlternateID")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ Create Alternate ID = " & dicWebBOInfo("CreateAlternateID") & " ].")
		End If
	End If
	If  dicWebBOInfo("CheckOutItemRevisionOnCreate") <> "" Then
		Set objEdit = objPSTable.ChildItem( 8, 2, "WebCheckBox", 0)
		If TypeName(objEdit) <> "Nothing" Then
				objEdit.set dicWebBOInfo("CheckOutItemRevisionOnCreate")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ Check-Out Item Revision On Create = " & dicWebBOInfo("CheckOutItemRevisionOnCreate") & " ].")
		End If
	End If
	objButtonPanel.WebButton("Next").Click 1,1,micLeftBtn
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully Clicked on [ Next ] button.")

	Set objPSTable =  Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewPSConnection").WebTable("Additional_MFK_ConnInfo")
	If  dicWebBOInfo("MFK Key1") <> "" Then
		Set objEdit = objPSTable.ChildItem(1, 2,"WebEdit", 0)
		If TypeName(objEdit) <> "Nothing" Then
				'objEdit.set dicWebBOInfo("MFK Key1")
				objEdit.Object.focus
				objobjMDR.SendString dicWebBOInfo("MFK Key1")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ MFK Key1 = " & dicWebBOInfo("MFK Key1") & " ].")
		End If
	End If
	If  dicWebBOInfo("MFK Key2") <> "" Then
		Set objEdit = objPSTable.ChildItem(2, 2,"WebEdit", 0)
		If TypeName(objEdit) <> "Nothing" Then
				'objEdit.set dicWebBOInfo("MFK Key2")
				objEdit.Object.focus
				objobjMDR.SendString dicWebBOInfo("MFK Key2")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully set [ MFK Key2 = " & dicWebBOInfo("MFK Key2") & " ].")
		End If
	End If
	' Clicking on Finish Button
	objButtonPanel.WebButton("Finish").Click 1,1,micLeftBtn
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : Successfully Clicked on [ Finish ] button.")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_Web_PSConnectionCreate : executed successfully.")
	Fn_Web_PSConnectionCreate = True
	Set objPSTable =  Nothing
	Set objButtonPanel =  Nothing
End Function 

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_BOMLineSearchResultOperations
'@@
'@@    Description				 :	Function Used to Perform Operation on BOM Line search results
'@@
'@@    Parameters			   :	1.StrAction : Action Name
'@@												  2.StrName : Search Name
'@@												  3.StrColName : Column Name
'@@												  4.StrValue : Column Value
'@@												  5.StrButton : Button Name
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	BOM Line Search Result Dialog should exist
'@@
'@@    Examples					:	Call Fn_BOMLineSearchResultOperations("Select","52737-Item4-85328-22773","","","Configure")
'@@												Call Fn_BOMLineSearchResultOperations("VerifyNames","06789-Item1-15272-07027~06789-Item3-28397-18255","","","Cancel")
'@@												Call Fn_BOMLineSearchResultOperations("GetAllColumnNames","","","","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done											Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									10-Fec-2012						1.0																												Amol Lanke
'@@												Swati K																  10-Fec-2012					1.1							Added Case "GetAllColumnNames"				Sandeep Navghane
'@@												Sandeep Navghane									10-Fec-2012					1.2								Added Case "VerifyNames"							Swati K
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_BOMLineSearchResultOperations(StrAction,StrName,StrColName,StrValue,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_BOMLineSearchResultOperations"
 	'Variable Declaration
	Dim objSeachTable,bFlag,objRdo,crrName,iRows,iCounter
	Dim iCount,ColName,arrColName,arrNames
	Fn_BOMLineSearchResultOperations=False
 	'Creating Object Of [ SearchResult ] table
	Set objSeachTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BOMLineSearchResult").WebTable("SearchResult")
	'Checking Existance Of [ SearchResult ] table
	If Not objSeachTable.Exist(6) Then
		Exit Function
	End If
	Select Case StrAction
		'Case to Select Result from the existing Results
		Case "Select"
			'Retriving number of results exist in [ SearchResult ] table
			iRows=objSeachTable.RowCount()
			For iCounter=1 To iRows
				'Taking Current name from  [ SearchResult ] table : Row by row
				crrName=objSeachTable.GetCellData(iCounter,2)
				'Checking Specific result matching or not
				If Trim(crrName)=Trim(StrName) Then
					Set objRdo=objSeachTable.ChildItem(iCounter,1,"WebRadioGroup",0)
						If TypeName(objRdo)<>"Nothing" Then
								'Selecting Specific result from SearchResult
								objRdo.Select "#"&Cstr(iCounter-3)
								bFlag=True
								Set objRdo=Nothing
								Exit For
						End If
				End If
			Next
			If bFlag=True Then
				If StrButton<>"" Then
					If StrButton="Configure" Then
						'Clicking on Configure button
						Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("Configure").Click 1,1,micLeftBtn
					ElseIf LCase(StrButton)="ok" Then
						Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("OKSearchResults").Click 1,1,micLeftBtn
					End If
				End If
				'Function Returns True
				Fn_BOMLineSearchResultOperations=True
			End If
			'----------------------------------------------------------------------------------------------------------------------------------------------------------------
			Case "GetAllColumnNames"
					iCount=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BOMLineSearchResult").WebTable("SearchResult").ColumnCount(1)
					For iCounter=2 To iCount
						ColName=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BOMLineSearchResult").WebTable("SearchResult").GetCellData(1,iCounter)
						If iCounter=2 Then
							arrColName=ColName
						Else
							arrColName=arrColName+"~"+ColName
						End If
					Next
					Fn_BOMLineSearchResultOperations=arrColName
			'----------------------------------------------------------------------------------------------------------------------------------------------------------------
			Case "VerifyNames"
				arrNames=Split(StrName,"~")
                For iCounter=0 to UBound(arrNames)
					bFlag=False
					iRows=objSeachTable.RowCount()
					For iCount=0 to  iRows
						crrName=objSeachTable.GetCellData(iCount,2)
						If Trim(crrName)=Trim(arrNames(iCounter)) Then
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=False Then
						Exit For
					End If
				Next
				If StrButton<>"" Then
					If LCase(StrButton)="cancel" Then
						Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("CancelBOMLineSearchResult").Click 1,1,micLeftBtn
					End If
				End If
				If bFlag=True Then
					Fn_BOMLineSearchResultOperations=True
				End If
	End Select
	Set objSeachTable=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Web_CompanyOperations

'Description			 :	Function Used to Create Business Object in Thin client

'Parameters			   :   '1.StrObjectName: Business Object  Name
'										 2.dicWebCompanyInfo: Business Object information 
'
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be log in Thin Client

'Examples				:  	dicWebCompanyInfo("Name")="Test"
'										dicWebCompanyInfo("Location Type")="CAGE, Commercial and Government Entity"
'										bReturn=Fn_Web_CompanyOperations("Company Location",dicWebCompanyInfo)
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												28-Feb-2012								1.0																						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_Web_CompanyOperations(StrObjectName,dicWebCompanyInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CompanyOperations"
 	'Declaring Variables
	Dim strWEBMenuPath,strMenu,crrType,iTemp,ObjBOInfo,ObjBO,arrKeys,iCounter,arrItem,iRwCount,iCount,sFieldName,objWebEle,bFlag
	Fn_Web_CompanyOperations=False
	'Checking Existance of [ New Business Object ] dialog
	If  Not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewBusinessObject").Exist(7) Then
	'If New Business Object dialog does not exist, Do menu operation
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewBusinessObject")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If	
	'Creating Object of [ New Business Object ] dialog
	Set ObjBO=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewBusinessObject")
	'Selecting [ Business Object ] type
	If StrObjectName<>"" Then
		crrType=ObjBO.WebEdit("BOTypeEdit").GetROProperty("value")
		wait(1)
		If Trim(crrType)<>Trim(StrObjectName) Then
				'Setting Business Object Type
				Call Fn_Web_UI_Button_Click("Fn_Web_CompanyOperations",ObjBO,"BOType")
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CompanyOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",StrObjectName)
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
				wait(2)
		End If
	End If
	'Clicking Next button to go to [Business Object ] page
	Call Fn_Web_UI_Button_Click("Fn_Web_CompanyOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	'Declaring Temp counter
	iTemp=0
	arrKeys=dicWebCompanyInfo.Keys
	For iCounter=0 to dicWebCompanyInfo.Count-1
		'Creating Table object
		Select Case CInt(iTemp)
			Case 0
					Set ObjBOInfo=ObjBO.WebTable("CompanyInfo")
		End Select
		bFlag=True
		'Cases to fill the [ Business Object ] Fileds
		Select Case arrKeys(iCounter)
			Case "Location Type"
			
				If dicWebCompanyInfo(arrKeys(iCounter))<>"" Then
    				arrItem=Split(dicWebCompanyInfo(arrKeys(iCounter)),"~")
					bFlag=False
					'Taking Field name from Table
					iRwCount=ObjBOInfo.RowCount()
					For iCount=0 To iRwCount
						sFieldName=ObjBOInfo.GetCellData(iCount,1)
						If Trim(sFieldName)=arrKeys(iCounter)+":" Then
							Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebEdit", 0)
							bFlag=True
							Exit For
						ElseIF Trim(sFieldName)=arrKeys(iCounter)+"*:" Then
							Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebEdit", 0)
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=True Then
							If arrItem(0)<>objWebEle.GetROProperty("value") Then
								Set objWebEle =Nothing
								Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebButton", 0)
								objWebEle.Click 1,1
								Call Fn_WEB_UI_Object_SetTOProperty("Fn_Web_CompanyOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",arrItem(0))
								Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
							End If
					End If
				End IF
'			'to enter "ID","Name","Revision","MFK Key1","MFK Key2","Description"
			Case "Name","Description","CAGE Code","Location Code"
				
				If dicWebCompanyInfo(arrKeys(iCounter))<>"" Then
		
					arrItem=Split(dicWebCompanyInfo(arrKeys(iCounter)),"~")
					bFlag=False
					'Taking Field name from Table
					iRwCount=ObjBOInfo.RowCount()
					For iCount=0 To iRwCount
						sFieldName=ObjBOInfo.GetCellData(iCount,1)
						If Trim(sFieldName)=arrKeys(iCounter)+":" Then
							wait 1
							Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebEdit", 0)
							bFlag=True
							Exit For
						ElseIF Trim(sFieldName)=arrKeys(iCounter)+"*:" Then
							wait 1
							Set objWebEle = ObjBOInfo.ChildItem(iCount, 2, "WebEdit", 0)
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=False Then
						Set ObjBO=Nothing
						Set ObjBOInfo=Nothing
						Exit Function
					End If
					If TypeName(objWebEle) <> "Nothing" Then
						Set objMDR = CreateObject("Mercury.DeviceReplay")
						objWebEle.Object.focus
						wait 1
                        objMDR.SendString arrItem(0)
						Set objMDR =Nothing
					End If

                    If UBound(arrItem)=1 Then
						Call Fn_Web_UI_Button_Click("Fn_Web_CompanyOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
						iTemp=iTemp+1
					End If
				End If
                If UBound(arrItem)=1 Then
						Call Fn_Web_UI_Button_Click("Fn_Web_CompanyOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
						iTemp=iTemp+1
				End If
		End Select
	Next
	'clicking on finish button to create New BO
	Call Fn_Web_UI_Button_Click("Fn_Web_CompanyOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Finish")
	For iCounter=0 To 2
		If ObjBO.Exist(5) Then
			wait(5)
		Else
			Exit For
		End If
	Next
	Fn_Web_CompanyOperations=True
	Set ObjBO=Nothing
	Set ObjBOInfo=Nothing
	Set objWebEle =Nothing
End Function



''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_ExportObjects(sTransferMode,sRevisionRule,sLanguage,sButton,sVerifyMessage,bClickLink,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will perform Export Operation on the Selected Node
''''/$$$$  
''''/$$$$   PRE-REQUISITES        :  Login to Teamcenter Web Should be Done
''''/$$$$
''''/$$$$  PARAMETERS   : 		sTransferMode : Valid TransferMode Name
''''/$$$$											sRevisionRule :Valid RevisionRule Name
''''/$$$$ 											sLanguage: Valid Language(s)
''''/$$$$										sButton : Button to be CLicked
''''/$$$$										sVerifyMessage : To verify Export Complete Message
''''/$$$$										 bClickLink : To click the exported xml Hyperlink			
''''/$$$$										sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_Web_MenuOperation(), Fn_Web_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          30/03/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			30/03/2012            1.0
''''/$$$$
''''/$$$$		How To Use :   bReturn=Fn_ExportObjects("ConfiguredDataExportDefault","Any Status; Working","","OK","Succesfully completed exporting to plmxml.plmxml_AutomatedTests.xml","yes","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Function Fn_ExportObjects(sTransferMode,sRevisionRule,sLanguage,sButton,sVerifyMessage,bClickLink,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_ExportObjects"
   Dim objBrowser,bReturn,sValue
   Set objBrowser=Browser("TeamcenterWeb").Page("MyTeamCenter")
	Fn_ExportObjects=false

			'Invoke the Export Dialog
			bReturn=Fn_Web_MenuOperation("Select","Tools:PLMXML Export...")
			If bReturn=true Then
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Selected the Menu [Tools:PLMXML Export...]")	
				Fn_ExportObjects=True
			Else
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to Select the Menu [Tools:PLMXML Export...]")	
					Fn_ExportObjects=False
					Exit Function
			End If

			'Wait for Synchronisation
			bReturn=Fn_Web_ReadyStatusSync(1)

			'Enter the Transfer Mode Name
			If  sTransferMode<>"" Then
				wait(1)
				objBrowser.WebEdit("TransferMode").Set sTransferMode
				wait 1

				'Check for Invalid Selection Message
				If Browser("TeamcenterWeb").Dialog("Dialog").Exist(5) then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Selection ["+sTransferMode+"] is Invalid")
						Fn_ExportObjects=False
						Exit Function
				End if

				'Check if the Transfermode is set
				sValue=objBrowser.WebEdit("TransferMode").GetROProperty ("value")
				If lCase(sValue)=lCase(sTransferMode) Then
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the TransferMode ["+sTransferMode+"] is Set")	
					Fn_ExportObjects=True
				Else
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the TransferMode ["+sTransferMode+"] is Set")
						Fn_ExportObjects=False
						Exit Function
				End If
			End If
			wait 2


			'Enter the Revision Rule
			If  sRevisionRule<>"" Then
				wait(1)
				objBrowser.WebEdit("RevisionRule").Set sRevisionRule
				wait 1

				'Check for Invalid Selection Message
				If Browser("TeamcenterWeb").Dialog("Dialog").Exist(5) then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : The Selection ["+sRevisionRule+"] is Invalid")
						Fn_ExportObjects=False
						Exit Function
				End if

				'Check if the Transfermode is set
				sValue=objBrowser.WebEdit("RevisionRule").GetROProperty ("value")
				If lCase(sValue)=lCase(sRevisionRule) Then
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the Revision Rule ["+sRevisionRule+"] is Set")	
					Fn_ExportObjects=True
				Else
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the Revision Rule ["+sRevisionRule+"] is Set")
						Fn_ExportObjects=False
						Exit Function
				End If
			End If
			wait 2

			'Set the Language

			If sLanguage<>"" Then
				'Will be Coded as required 
			End If

			If sButton<>"" Then
				objBrowser.WebButton(sButton).Click 5,5,micLeftBtn
				wait 5
			End If

				If sVerifyMessage<>"" Then
					If Browser("ExportComplete").Page("ExportComplete").Exist(5) then
						sValue=Browser("ExportComplete").Page("ExportComplete").WebElement("Message").GetROProperty("innertext")
						If sValue<>sVerifyMessage Then
									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify the Export Message ["+sVerifyMessage+"]")
									Fn_ExportObjects=False
									Exit Function
						Else
									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified the Export Message ["+sVerifyMessage+"]")
									Fn_ExportObjects=True	
						End If
						wait 3
				End If
			End if


		If bClickLink<>"" then
			If uCase(bClickLink)="YES" Then
				Browser("ExportComplete").Page("ExportComplete").Link("Link").Click 5,5,micLeftBtn
			End If
	  End if

End Function



''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''/$$$$
''''/$$$$   FUNCTION NAME   :  Fn_SaveAs_XML(sFilePath,bCloseBrowser,sInfo1,sInfo2)
''''/$$$$
''''/$$$$   DESCRIPTION        :  This function will save the XML that is opened in the Browser to the Desired Path
''''/$$$$  
''''/$$$$   PRE-REQUISITES        :  The XML Browser Window should be Present
''''/$$$$
''''/$$$$  PARAMETERS   : 		sFilePath : Valid filepath to save the XML File
''''/$$$$											bCloseBrowser :To close the Browser
''''/$$$$											sInfo1: For Future Use
''''/$$$$										sInfo2 : For Future Use
''''/$$$$	
''''/$$$$		Return Value : 				True or False
''''/$$$$
''''/$$$$    Function Calls       :   Fn_WriteLogFile(), Fn_Web_MenuOperation(), Fn_Web_ReadyStatusSync
''''/$$$$										
''''/$$$$
''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''/$$$$
''''/$$$$    CREATED BY     :   SHREYAS          30/03/2012         1.0
''''/$$$$
''''/$$$$    REVIWED BY     :  Shreyas			30/03/2012            1.0
''''/$$$$
''''/$$$$		How To Use :   bReturn=Fn_SaveAs_XML("D:\Shreyas.xml","yes","","")
''''/$$$$							
''''/$$$$	
''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function Fn_SaveAs_XML(sFilePath,bCloseBrowser,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SaveAs_XML"

   Dim objToolbar,objSave,objBrowser
   Dim objFSO
   Fn_SaveAs_XML=FAlse

   Set objToolbar=Browser("ExportComplete").WinToolbar("ToolBar")
   Set objSave=Browser("ExportComplete").Dialog("Save As")
   Set objBrowser=Browser("ExportComplete")


		   'Check if the Object Exists
		   If not objToolbar.Exist(10) Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the Toolbar Exists")
						Fn_SaveAs_XML=False
						Exit Function
			Else
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the Toolbar Exists")
						Fn_SaveAs_XML=True	
		   End If

		  'Click the File Option in the ToolBar

			'objToolbar.Press "&File"
			Browser("ExportComplete").WinToolbar("Toolbar").Press "&File"
			If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to Click on the [File] Menu on the Toolbar")
						Fn_SaveAs_XML=False
						Exit Function	
			Else
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Clicked on the [File] Menu on the Toolbar")
						Fn_SaveAs_XML=True				
			End If
			wait 2

			'Select the Save AS Operation in the ContextMenu

			objBrowser.WinMenu("ContextMenu").Select "Save as...	Ctrl+S"
'			objBrowser.WinMenu("ContextMenu").Select "Save As..."
			If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to Click on the [Save As...] Context Menu ")
						Fn_SaveAs_XML=False
						Exit Function	
			Else
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Clicked on the [Save As...] Context Menu ")
						Fn_SaveAs_XML=True					
			End If
			wait 1

			'Check if the Save AS Dialog Exists
		   If not objSave.Exist(10) Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verify that the [Save As] Dialog Exists")
						Fn_SaveAs_XML=False
						Exit Function
			Else
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully verified that the [Save As] Dialog Exists")
						Fn_SaveAs_XML=True	
		   End If
		   wait 1

		   'Set the File Path And Click on Save

		   'objSave.WinEdit("FilePath").Set sFilePath
		   'Added by Nilesh on 27-Mar-2013
           If Instr(1,sFilePath,"PIE")>1Then
			   Set objFSO = CreateObject("Scripting.FileSystemObject")
				If not objFSO.FolderExists(Environment.Value("BatchFldName")+"\PIE") then
					objFSO.CreateFolder(Environment.Value("BatchFldName")&"\PIE")
					wait 5
				End if
				Set objFSO=nothing
		   End If

		   'Browser("ExportComplete").Dialog("Save As").WinEdit("FilePath").Set sFilePath
			Wait 1
			Browser("ExportComplete").Dialog("Save As").WinEdit("FilePath").Type sFilePath
			If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to set the File Path ["+sFilePath+"]")
						Fn_SaveAs_XML=False
						Exit Function	
			Else
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully set the File Path ["+sFilePath+"]")
						Fn_SaveAs_XML=True					
			End If
			wait 1
			Err.Clear
			objSave.WinButton("Save").Click 5,5,micLeftBtn
			If Err.Number<0 Then
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to  click on the Save Button")
						Fn_SaveAs_XML=False
						Exit Function	
			Else
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully clicked on the Save Button")
						Fn_SaveAs_XML=True					
			End If
			wait 3

			If uCase(bCloseBrowser)="YES" Then
				objBrowser.Close
				If Err.Number<0 Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to  Close the Browser")
							Fn_SaveAs_XML=False
							Exit Function	
				Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Closed the Browser")
							Fn_SaveAs_XML=True					
				End If
			End If

End Function



''''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''''''/$$$$
''''''''/$$$$   FUNCTION NAME   :   Fn_SISW_Web_AuditLogOperations(sAction,sColNames,sColValues,sInfo1,sInfo2,sInfo3)
''''''''/$$$$
''''''''/$$$$   DESCRIPTION        :  This function Will  Perform  several desired operations in the AuditLog Table in Web Version of Teamcenter
''''''''/$$$$
''''''''/$$$$	PRE-REQUISITERS :  Summary Tab should Be Activated
''''''''/$$$$
''''''''/$$$$  PARAMETERS   : 		sAction : Valid Action Name
''''''''/$$$$										sColNames : Valid Column names for values to be verified
''''''''/$$$$										sColValues : Valid Column Values
''''''''/$$$$										sInfo1 : For Future Use
''''''''/$$$$										sInfo2:	For Future Use
''''''''/$$$$										sInfo3:	For Future Use
''''''''/$$$$	
''''''''/$$$$		Return Value : 				True or False
''''''''/$$$$
''''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''''''/$$$$
''''''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''''''/$$$$  
''''''''/$$$$    CREATED BY     :   SHREYAS          17/09/2012         1.0
''''''''/$$$$
''''''''/$$$$    REVIWED BY     :  Shreyas			17/09/2012            1.0
''''''''/$$$$
''''''''/$$$$		How To Use :      				Example #1
'''''''/$$$$																
'''''''/$$$$							bReturn=  Fn_SISW_Web_AuditLogOperations("VerifyData","Event Type Name:Object Type:Event Type Name","__Check_Out:Item:__Check_In","","","")
'''''''/$$$$
'''''''/$$$$													Example #2
'''''''/$$$$							
'''''''/$$$$							bReturn=  Fn_SISW_Web_AuditLogOperations("VerifyData","Object Type","Item","","","")
'''''''/$$$$
''''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$



Public function Fn_SISW_Web_AuditLogOperations(sAction,sColNames,sColValues,sInfo1,sInfo2,sInfo3)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_AuditLogOperations"
   Fn_SISW_Web_AuditLogOperations=false
   Dim aColumns,aColsData,iCols,iRows,bFlag
   Dim colcnt,aResult,j
   Dim aTable,i,aColumns1
   Dim iCounter,jCounter,mCounter,aColname
   Dim ColName,colIndex,nCounter,sCol,aCol,intCol,bCheck,objTable
   Dim tempVar
   Dim sColVal,objTableHdr,sMenu
	Set objTable=Browser("Browser").Page("AuditLogs").WebTable("AuditLogTable")
	bFlag=False
	bCheck=False
   Select Case sAction


 	Case "VerifyData"
			'Modified case by Nilesh Gadekar on 26-Sep-2012
			If sColNames<>"" and sColValues<>"" Then
						aColumns = Split(sColNames,":",-1,1)  'get column name in array to be verify
						aColumns1 = Split(sColNames,":",-1,1)
						aColsData = Split(sColValues,":",-1,1)' Get values in array to be verify
						iCols = objTable.GetROProperty("cols")    'column count
						iRows =objTable.GetROProperty("rows")  'row count
						'Get displayed column name in array
						ReDim aColname(iCols-1)
						For mCounter=1 To iCols
							aColname(mCounter-1)=	objTable.GetCellData( 1,mCounter)
						Next

						'Remove duplicate column names
						For i=0 to Ubound(aColumns1)
								For j=i+1 To Ubound(aColumns1)
									If aColumns1(i)=aColumns1(j) and aColumns1(i)<>""  Then
										aColumns1(j)=""
									End If
								Next
						Next

						'Get column names into string
						For i=0 To Ubound(aColumns1)
							If aColumns1(i)<>"" Then
								If sCol<>"" Then
									sCol=sCol+":"
								End If
								sCol=sCol+aColumns1(i)
							End If
						Next

							'Create array of colums
							aCol=Split(sCol,":",-1,1)
							intCol=Ubound(aCol)
							'Define Result Array 
							ReDim aResult(Ubound(aColsData))
							For j=0 to Ubound(aColsData)
								aResult(j)=False
							Next

						If intCol=0 Then
											'Take neccessary column values in array
											ReDim aTable (iRows-2)
											For j=1 to iRows-1
												For mCounter=0 To Ubound(aCol)
													For iCounter=0 To Ubound(aColname)
														If Trim(aCol(mCounter))=Trim(aColname(iCounter)) Then
															colcnt=iCounter+1
															bCheck=True
															Exit For
														End If
														
													Next
													If bCheck=True Then
															Exit For
														End If
												Next
													'Get values cell from table
													aTable(j-1)=objTable.GetCellData(j+1,colcnt)
											Next 

											'validating actual with expected
											For iCounter=0 To Ubound(aColsData)
												For nCounter=0 To iRows-2
														If Trim(aTable(nCounter))= Trim(aColsData(iCounter))Then
															aResult(iCounter)=True
															aTable(nCounter)=""
															Exit For
														End If
												Next
										Next

							Else
							'Exception case where only one column values to be validate
										ReDim aTable (intCol,(iRows-2))
										For j=0 to intCol
											For i=1 to iRows-1
												For iCounter=0 To Ubound(aColname)
													If Trim(aCol(j))=Trim(aColname(iCounter)) Then
														colcnt=iCounter+1
														Exit For
													End If
												Next
								'[TC1015-20151013A-04_11_2015-VivekA-Maintenance] - Added by Chandrakant T
								'Added code to handle extra string attached to getcell data text for 1st column of Audit Licence change table in web				
													
													If Instr(1,objTable.GetCellData( i+1,colcnt),"stringReplace('__") <> 0 Then
														tempVar=Split(Trim(objTable.GetCellData( i+1,colcnt)),")")
														aTable((j),(i-1)) = Trim(tempVar(UBound(tempVar)))
													Else
														aTable((j),(i-1))=objTable.GetCellData( i+1,colcnt)
													End If
								'------------------------------------------------------------------
											Next
										Next
										
										For iCounter=0 To Ubound(aColsData)
												ColName=aColumns(iCounter)
												For jCounter=0 To Ubound(aCol)
													If Trim(ColName)=Trim(aCol(jCounter)) Then
														colIndex=jCounter
													End If
												Next

												For nCounter=0 To iRows-2
														If Trim(aTable(colIndex,nCounter))= Trim(aColsData(iCounter))Then
															aResult(iCounter)=True
															aTable(colIndex,nCounter)=""  'Delete array value once its verified
															Exit For
														End If
												Next
										Next
						End If

	
		

						'Get final result 
						For j=0 to Ubound(aColsData)
								If aResult(j)=False Then
									bFlag=False
									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Failed to verfiy audit value  ["+aColsData(j)+"] from colum ["+aColumns1(0)+"]")
									Exit For
								Else
									bFlag=True
									Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Successfully Verified all audit values ")
								End If
						Next
			End If
			'Final Result
			If bFlag=false Then
				Fn_SISW_Web_AuditLogOperations=False
			Else
				Fn_SISW_Web_AuditLogOperations=True
			End if
        Case "GetData"
				If sColNames<>"" Then
						aColumns = Split(sColNames,":",-1,1)  'get column name in array to be verify
						aColumns1 = Split(sColNames,":",-1,1)
						iCols = objTable.GetROProperty("cols")    'column count
						iRows =objTable.GetROProperty("rows")  'row count
						'Get displayed column name in array
						ReDim aColname(iCols-1)
						For mCounter=1 To iCols
							aColname(mCounter-1)=	objTable.GetCellData( 1,mCounter)
						Next

						'Remove duplicate column names
						For i=0 to Ubound(aColumns1)
								For j=i+1 To Ubound(aColumns1)
									If aColumns1(i)=aColumns1(j) and aColumns1(i)<>""  Then
										aColumns1(j)=""
									End If
								Next
						Next

						'Get column names into string
						For i=0 To Ubound(aColumns1)
							If aColumns1(i)<>"" Then
								If sCol<>"" Then
									sCol=sCol+":"
								End If
								sCol=sCol+aColumns1(i)
							End If
						Next

							'Create array of colums
							aCol=Split(sCol,":",-1,1)
							intCol=Ubound(aCol)

						If intCol=0 Then
											'Take neccessary column values in array
											ReDim aTable (iRows-2)
											For j=1 to iRows-1
												For mCounter=0 To Ubound(aCol)
													For iCounter=0 To Ubound(aColname)
														If Trim(aCol(mCounter))=Trim(aColname(iCounter)) Then
															colcnt=iCounter+1
															bCheck=True
															Exit For
														End If
														
													Next
													If bCheck=True Then
															Exit For
														End If
												Next
													'Get values cell from table
													aTable(j-1)=objTable.GetCellData(j+1,colcnt)
											Next 

							Else
							'Exception case where only one column values to be validate
										ReDim aTable (intCol,(iRows-2))
										For j=0 to intCol
											For i=1 to iRows-1
												For iCounter=0 To Ubound(aColname)
													If Trim(aCol(j))=Trim(aColname(iCounter)) Then
														colcnt=iCounter+1
														Exit For
													End If
												Next
												aTable((j),(i-1))=objTable.GetCellData( i+1,colcnt)
											Next
										Next
						End If
			End If
			'Final Result
			Dim str
		If Ubound(aColumns)<>0 Then
					ReDim aResult( Ubound(aTable,1))
					
					For i=0 to Ubound(aTable,1)
						For j=0 to Ubound(aTable,2)
							If  j<>0  Then
								str=str+"~"+aTable(i,j)
							Else
								str=aTable(i,j)
							End If
						Next
						aResult(i)=str
					Next		
			Else
					ReDim aResult(Ubound(aTable))
					For i=0 To Ubound(aTable)
						aResult(i)=aTable(i)
					Next
					str=Join(aResult,"~")
					ReDim aResult(0)
					aResult(0)=str
			End If								
				If ISArray(aResult)=false Then
					Fn_SISW_Web_AuditLogOperations=False
				Else
					Fn_SISW_Web_AuditLogOperations=Join(aResult,"|")	
				End if
	
		 'TC11.3_20170509d.00_NewDevelopment_DIPRO_Admin_PoonamC_13Oct2017 : Added Case to verify Audit Log				
  		 '----------------------------------------------------------------------------------------------------------
  		 Case "VerifyTableData_AuditLogTab"
					If sColNames<>"" and sColValues<>"" Then
						'Select Audit Log Tab
						Call Fn_Web_TabOperations("Activate","Audit Logs")
						Call Fn_Web_ReadyStatusSync(1)
						
						Set objTable = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AuditLogTable")
						Set objTableHdr = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AuditLogTableHeader")
						
						'Click General Logs link
						If Fn_Web_UI_ObjectExist("Fn_SISW_Web_AuditLogOperations", objTable)  = False Then
							Browser("TeamcenterWeb").Page("MyTeamCenter").Link("GeneralLogLink").Click 1,1,micLeftBtn 
							Call Fn_Web_ReadyStatusSync(1)
						End If
						
						aColumns = Split(sColNames,":",-1,1)  'get column name in array to be verify
						aColsData = Split(sColValues,":",-1,1)' Get values in array to be verify
						iCols = objTableHdr.GetROProperty("column names")    'column count
						iRows =objTable.GetROProperty("rows")  
						aColumns1 = Split(iCols,";")
						
						For i=0 to Ubound(aColumns)
								colIndex = -1
								For j = 0 To UBound(aColumns1)
							    	If trim(aColumns(i)) = trim(aColumns1(j)) Then
										colIndex = j + 1
										Exit For
									End If
								Next
								
							   If colIndex <> -1 Then
							   		bFlag = False
							   		 For j = 2 To iRows
										  	sColVal = objTable.GetCellData(j,colIndex)
										  	If trim(sColVal) = trim(aColsData(i)) Then
										  		 bFlag = True
										  		 Exit For
										  	End If
									  Next 
							   End If 
							   
							   If bFlag = False Then
							   		Fn_SISW_Web_AuditLogOperations=False
							   		Exit Function
							   Else
									Fn_SISW_Web_AuditLogOperations=True				   	
							   End If
						Next	   
				   End if
		 Case "VerifyTableData_AuditLogMenu"
				If sColNames<>"" and sColValues<>"" Then
						'Call menu operation
						 Set objTable = Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("AuditLogTable2")
						 If Fn_Web_UI_ObjectExist("Fn_SISW_Web_AuditLogOperations", objTable)  = False Then
						 		sMenu = Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("WebMyTc_Menu"),"ViewAuditLogs") 
						 		Call Fn_Web_MenuOperation("Select",sMenu)
						 		Call Fn_Web_ReadyStatusSync(1)
						 End If
						 
						'click Serach button
						Call Fn_Web_UI_Button_Click("Fn_SISW_Web_AuditLogOperations", Browser("TeamcenterWeb").Page("MyTeamCenter"), "Search")
						Call Fn_Web_ReadyStatusSync(1)
						
						'Set Object
						aColumns = Split(sColNames,":",-1,1)  
						aColsData = Split(sColValues,":",-1,1)
						iRows = objTable.GetROProperty("rows") ' Get rows count
						iCols = objTable.ColumnCount(3)		   ' Get column count	
						
						For i=0 to Ubound(aColumns)
								colIndex = -1
								For j = 1 To iCols
							    	If trim(aColumns(i)) = trim(objTable.GetCellData(3,j)) Then
										colIndex = j
										Exit For
									End If
								Next
								
							   If colIndex <> -1 Then
							   		bFlag = False
							   		 For j = 4 To iRows
										  	sColVal = objTable.GetCellData(j,colIndex)
										  	If trim(sColVal) = trim(aColsData(i)) Then
										  		 bFlag = True
										  		 Exit For
										  	End If
									  Next 
							   End If 
							   
							   If bFlag = False Then
							   		Fn_SISW_Web_AuditLogOperations=False
							   		Exit Function
							   Else
									Fn_SISW_Web_AuditLogOperations=True				   	
							   End If
						Next	   
				   End if
			'----------------------------------------------------------------------------------------------------------	   
  End Select
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Web_MassUpdateOperations

'Description			 :	Function Used to Perform operation of Mass update

'Parameters			   :  1.StrAction : Action name
'									   2.dicWebMassUpdateInfo: Mass updation information
'
'Return Value		   : 	True or False

'Pre-requisite			:	Object should be selected on which have to perform mass updation

'Examples				:   Dim dicWebMassUpdateInfo
'										Set dicWebMassUpdateInfo = CreateObject( "Scripting.Dictionary")
'
'										dicWebMassUpdateInfo("ItemId")="000019"
'										dicWebMassUpdateInfo("Object")="000019/A;1-CommonTir"
'										dicWebMassUpdateInfo("SearchDialogButtonName")="OK"
'										bReturn= Fn_SISW_Web_MassUpdateOperations("Target",dicWebMassUpdateInfo)
'
'										dicWebMassUpdateInfo("ItemId")="000020"
'										dicWebMassUpdateInfo("Object")="000020/A;1-NewTire"
'										dicWebMassUpdateInfo("SearchDialogButtonName")="OK"
'										dicWebMassUpdateInfo("ButtonName")="Next"
'										bReturn= Fn_SISW_Web_MassUpdateOperations("Replacement",dicWebMassUpdateInfo)
'
'										dicWebMassUpdateInfo("Operation")="SelectAll"
'										dicWebMassUpdateInfo("ButtonName")="Next"
'										bReturn= Fn_SISW_Web_MassUpdateOperations("ImpactedPartsToUpdateTable",dicWebMassUpdateInfo)
'										
'										dicWebMassUpdateInfo("Operation")="Select"
'										dicWebMassUpdateInfo("Object")="000036_4816/A;1-WheelAssy2013~000027_4816/A;1-WheelAssy2012"
'										bReturn= Fn_SISW_Web_MassUpdateOperations("ImpactedPartsToUpdateTable",dicWebMassUpdateInfo)
'										
'										bReturn= Fn_SISW_Web_MassUpdateOperations("Execute",dicWebMassUpdateInfo)
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Sandeep N								16-Nov-2012					1.0																									 	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Web_MassUpdateOperations(StrAction,dicWebMassUpdateInfo)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_MassUpdateOperations"
	Fn_SISW_Web_MassUpdateOperations=False
	Dim StrWEBMenuPath,bFlag,aObject,iCounter,iCount, aValue, iCnt
	Dim ObjMassUpdateDialog,dicWebMassUpdateSearchInfo,ObjCheckBox
 	'checking existance of [ Mass Update ] dialog
 	If not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdate").Exist(6) Then
		StrWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		Call Fn_Web_MenuOperation("Select",Fn_GetXMLNodeValue(StrWEBMenuPath, "EditMassUpdate"))
	End If
 	'creating object of [ Mass Update ] dialog
	Set ObjMassUpdateDialog=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdate")
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'case to select Operation for Mass Update
		Case "SelectOperation"
			'selecting Operation from list
			Fn_SISW_Web_MassUpdateOperations=Fn_Web_UI_List_Select("Fn_SISW_Web_MassUpdateOperations", ObjMassUpdateDialog, "Operation",dicWebMassUpdateInfo("Operation"))
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'case to select target
		Case "Target"
			If not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateSearchCiteria").Exist(3) Then
				ObjMassUpdateDialog.Image("SearchTarget").Click 1,1,micLeftBtn
			else
				Call Fn_Web_UI_Button_Click("Fn_SISW_Web_MassUpdateSearchOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateSearchCiteria"), "Clear")
			End if
			wait 2
			
			Set dicWebMassUpdateSearchInfo = CreateObject( "Scripting.Dictionary")

			dicWebMassUpdateSearchInfo("Name")=dicWebMassUpdateInfo("Name")
			dicWebMassUpdateSearchInfo("ItemId")=dicWebMassUpdateInfo("ItemId")
			dicWebMassUpdateSearchInfo("Revision")=dicWebMassUpdateInfo("Revision")
			bFlag=Fn_SISW_Web_MassUpdateSearchOperations("EnterSearchCriteria",dicWebMassUpdateSearchInfo,"")
			If dicWebMassUpdateInfo("Object")<>"" and bFlag=True Then
				dicWebMassUpdateSearchInfo("Object")=dicWebMassUpdateInfo("Object")
				bFlag=Fn_SISW_Web_MassUpdateSearchOperations("Select",dicWebMassUpdateSearchInfo,dicWebMassUpdateInfo("SearchDialogButtonName"))
			End If
			'setting target description
			If dicWebMassUpdateInfo("Description")<>"" Then
				Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_MassUpdateOperations", ObjMassUpdateDialog, "TargetPartDescription", dicWebMassUpdateInfo("Description"))
			End If
			'selecting Add to newstuff Folder option
			If dicWebMassUpdateInfo("AddToNewstuffFolder")<>"" Then
				Call Fn_Web_UI_CheckBox_Set("Fn_SISW_Web_MassUpdateOperations", ObjMassUpdateDialog,"CheckTarget", dicWebMassUpdateInfo("AddToNewstuffFolder"))
			End If
			If bFlag=True Then
				Fn_SISW_Web_MassUpdateOperations=True
			End If
			dicWebMassUpdateSearchInfo.RemoveAll
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'case to Add or Replace
		Case "Add","Replacement"
			If not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateSearchCiteria").Exist(3) Then
				ObjMassUpdateDialog.Image("SearchAddReplacement").Click 1,1,micLeftBtn
			else
				Call Fn_Web_UI_Button_Click("Fn_SISW_Web_MassUpdateSearchOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateSearchCiteria"), "Clear")
			End if
			wait 2
			
			Set dicWebMassUpdateSearchInfo = CreateObject( "Scripting.Dictionary")

			dicWebMassUpdateSearchInfo("Name")=dicWebMassUpdateInfo("Name")
			dicWebMassUpdateSearchInfo("ItemId")=dicWebMassUpdateInfo("ItemId")
			dicWebMassUpdateSearchInfo("Revision")=dicWebMassUpdateInfo("Revision")
			bFlag=Fn_SISW_Web_MassUpdateSearchOperations("EnterSearchCriteria",dicWebMassUpdateSearchInfo,"")
			If dicWebMassUpdateInfo("Object")<>"" and bFlag=True Then
				dicWebMassUpdateSearchInfo("Object")=dicWebMassUpdateInfo("Object")
				bFlag=Fn_SISW_Web_MassUpdateSearchOperations("Select",dicWebMassUpdateSearchInfo,dicWebMassUpdateInfo("SearchDialogButtonName"))
			End If
			'setting target description
			If dicWebMassUpdateInfo("Description")<>"" Then
				Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_MassUpdateOperations", ObjMassUpdateDialog, "AddReplacementPartDescription", dicWebMassUpdateInfo("Description"))
			End If
			'selecting Add to newstuff Folder option
			If dicWebMassUpdateInfo("AddToNewstuffFolder")<>"" Then
				Call Fn_Web_UI_CheckBox_Set("Fn_SISW_Web_MassUpdateOperations", ObjMassUpdateDialog,"CheckAddReplacement", dicWebMassUpdateInfo("AddToNewstuffFolder"))
			End If
			If bFlag=True Then
				Fn_SISW_Web_MassUpdateOperations=True
			End If
			dicWebMassUpdateSearchInfo.RemoveAll
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'case to click on execute button
		Case "Execute"
			Fn_SISW_Web_MassUpdateOperations=Fn_Web_UI_Button_Click("Fn_SISW_Web_MassUpdateOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),"Execute")
			wait 3
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to perform operations on impacted part to update table
		Case "ImpactedPartsToUpdateTable"
			If ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").Exist(5) Then
					Select Case dicWebMassUpdateInfo("Operation")
						'case to select all parts
						Case "SelectAll"
							Set ObjCheckBox=ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").ChildItem(1,0,"WebCheckBox",0)
							If TypeName(ObjCheckBox)<>"Nothing" Then
								ObjCheckBox.Set "ON"
								Fn_SISW_Web_MassUpdateOperations=True
							End If
							Set ObjCheckBox=Nothing
						'case to select part from table
						Case "Select"
							aObject=Split(dicWebMassUpdateInfo("Object"),"~",-1,1)
							For iCounter=0 to ubound(aObject)
								bFlag=False
								For iCount=0 to ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").RowCount
									If trim(aObject(iCounter))=trim(ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").GetCellData(iCount,2)) Then
										Set ObjCheckBox=ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").ChildItem(iCount,0,"WebCheckBox",0)
										If TypeName(ObjCheckBox)<>"Nothing" Then
											ObjCheckBox.Set "ON"
											bFlag=True
											Exit for
										End If
										Set ObjCheckBox=Nothing
									End If
								Next
								If bFlag=False Then
									Exit for
								End If
							Next
							If bFlag=True Then
								Fn_SISW_Web_MassUpdateOperations=True
							End If
						'case to select part from table
						Case "Verify"
							bFlag=False
							aObject=Split(dicWebMassUpdateInfo("Object"),"~",-1,1)
							aValue=Split(dicWebMassUpdateInfo("Value"),"~",-1,1)	
							For iCnt=0 to ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").ColumnCount(1)
								If ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").GetCellData(1,iCnt)=dicWebMassUpdateInfo("ColumnName") Then
                                    bFlag=True
									Exit for
								End If
							Next
							If bFlag=True Then
								For iCounter=0 to ubound(aObject)
									bFlag=False
									For iCount=0 to ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").RowCount
										If trim(aObject(iCounter))=trim(ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").GetCellData(iCount,2)) Then
											If trim(aValue(iCounter))=trim(ObjMassUpdateDialog.WebTable("ImpactedPartsToUpdateTable").GetCellData(iCount,iCnt)) Then
												bFlag=True
												Exit For
											End If	
										End If
									Next
									If bFlag=False Then
										Exit for
									End If
								Next
							End If
							If bFlag=True Then
								Fn_SISW_Web_MassUpdateOperations=True
							End If
				End Select
			End If
	End Select

	If dicWebMassUpdateInfo("ButtonName")<>"" Then
		Call Fn_Web_UI_Button_Click("Fn_SISW_Web_MassUpdateOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), dicWebMassUpdateInfo("ButtonName"))
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Web_MassUpdateSearchOperations

'Description			 :	Function Used to Perform search for Mass update

'Parameters			   :  1.StrAction : Action name
'									   2.dicWebMassUpdateSearchInfo: Mass updation search information
'									   3.StrButtonName : button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Search dialog for Mass Update should be exist

'Examples				:   Set dicWebMassUpdateSearchInfo = CreateObject( "Scripting.Dictionary")

'										dicWebMassUpdateSearchInfo("Name")=dicWebMassUpdateInfo("Name")
'										dicWebMassUpdateSearchInfo("ItemId")=dicWebMassUpdateInfo("ItemId")
'										dicWebMassUpdateSearchInfo("Revision")=dicWebMassUpdateInfo("Revision")
'										bFlag=Fn_SISW_Web_MassUpdateSearchOperations("EnterSearchCriteria",dicWebMassUpdateSearchInfo,"")
'
'										dicWebMassUpdateSearchInfo("Object")=dicWebMassUpdateInfo("Object")
'										bFlag=Fn_SISW_Web_MassUpdateSearchOperations("Select",dicWebMassUpdateSearchInfo,dicWebMassUpdateInfo("SearchDialogButtonName"))
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Sandeep N								16-Nov-2012					1.0																									 	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Web_MassUpdateSearchOperations(StrAction,dicWebMassUpdateSearchInfo,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_MassUpdateSearchOperations"
   Fn_SISW_Web_MassUpdateSearchOperations=false
   Dim ObjMassUpdateSearchCiteriaDialog,ObjMassUpdateSearchResultTable,ObjRadioButton
   Dim dicKeys,dicItems,iCounter
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'case to enter search criteria
		Case "EnterSearchCriteria"
			'checking existance of [ MassUpdateSearchCiteria ] dialog
		   If not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateSearchCiteria").Exist(6) Then
			   Exit function
		   End If
		   'creating object of [ MassUpdateSearchCiteria ] dialog
			Set ObjMassUpdateSearchCiteriaDialog=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateSearchCiteria")
			dicKeys=dicWebMassUpdateSearchInfo.Keys
			dicItems=dicWebMassUpdateSearchInfo.Items
			For iCounter=0 to dicWebMassUpdateSearchInfo.Count-1
				If dicItems(iCounter)<>"" Then
					Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_MassUpdateSearchOperations", ObjMassUpdateSearchCiteriaDialog, dicKeys(iCounter), dicItems(iCounter))
				End If
			Next
			Fn_SISW_Web_MassUpdateSearchOperations=Fn_Web_UI_Button_Click("Fn_SISW_Web_MassUpdateSearchOperations", ObjMassUpdateSearchCiteriaDialog, "Find")
			wait 3
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'case to select entry from search results
		Case "Select"
		   If not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateSearchResult").Exist(6) Then
			   Exit function
		   End If
		   'creating object of [ MassUpdateSearchCiteria ] dialog
			Set ObjMassUpdateSearchResultTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateSearchResult")
			For iCounter=0 to ObjMassUpdateSearchResultTable.RowCount
				If trim(dicWebMassUpdateSearchInfo("Object"))=ObjMassUpdateSearchResultTable.GetCellData(iCounter,2) Then
					Set ObjRadioButton=ObjMassUpdateSearchResultTable.ChildItem(iCounter,1,"WebRadioGroup",0)
					If TypeName(ObjRadioButton)<>"Nothing" Then
						ObjRadioButton.Select "#"&Cstr(iCounter-3)
						Fn_SISW_Web_MassUpdateSearchOperations=True
						Exit for
					End If
				End If
			Next
			If StrButtonName<>"" Then
				If lcase(StrButtonName)="ok" Then
					Call Fn_Web_UI_Button_Click("Fn_SISW_Web_MassUpdateSearchOperations", Browser("TeamcenterWeb").Page("MyTeamCenter"), "OK")
				End If
			End If

	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Web_MassUpdateResultOperations

'Description			 :	Function Used to Perform operations on Mass update results

'Parameters			   :  1.StrAction : Action name
'									   2.dicWebMassUpdateResultInfo: Mass updation result information
'									   3.StrButtonName : button name
'
'Return Value		   : 	True or False

'Pre-requisite			:	Result dialog for Mass Update should be exist

'Examples				: 	Dim dicWebMassUpdateResultInfo
'										Set dicWebMassUpdateResultInfo = CreateObject( "Scripting.Dictionary")
'
'										dicWebMassUpdateResultInfo("Target")="000019/A;1-CommonTir"
'										dicWebMassUpdateResultInfo("Change Object")=""
'										bReturn= Fn_SISW_Web_MassUpdateResultOperations("VerifySummary",dicWebMassUpdateResultInfo,"")
'
'										dicWebMassUpdateResultInfo("ImpactedParts")="000045_4816/A;1-WheelAssy2014~000027_4816/A;1-WheelAssy2012"
'										dicWebMassUpdateResultInfo("ColumnName")="Status"
'										dicWebMassUpdateResultInfo("Value")="Success~Success"
'										bReturn= Fn_SISW_Web_MassUpdateResultOperations("VerifyImpactedParts",dicWebMassUpdateResultInfo,"Close")
'                       
'History					 :			
'										Developer Name							Date						Rev. No.				Changes Done											Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'										Sandeep N								16-Nov-2012					1.0																									 	Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Web_MassUpdateResultOperations(StrAction,dicWebMassUpdateResultInfo,StrButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_MassUpdateResultOperations"
	Dim dicItems,dicKeys,iCounter,iCount,bFlag,iColNumber,aImpactedParts,aValue

	Fn_SISW_Web_MassUpdateResultOperations=False
   Select Case StrAction
	 	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	 	'Case to verify summary of Mass Update result
	 	Case "VerifySummary"
			dicItems=dicWebMassUpdateResultInfo.Items
			dicKeys=dicWebMassUpdateResultInfo.Keys
			For iCounter=0 to dicWebMassUpdateResultInfo.Count-1
				bFlag=False
				For iCount=0 to Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateResultSummary").RowCount
					If trim(dicKeys(iCounter))=trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateResultSummary").GetCellData(iCount,1)) Then
						If trim(dicItems(iCounter))=trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateResultSummary").GetCellData(iCount,2)) Then
							bFlag=True
							Exit For
						End If	
					End If
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_Web_MassUpdateResultOperations=True
			End If
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
	 	'Case to verify summary of Mass Update result
	 	Case "VerifyImpactedParts"
			aImpactedParts=Split(dicWebMassUpdateResultInfo("ImpactedParts"),"~")
			If dicWebMassUpdateResultInfo("Value")<>"" Then
				aValue=Split(dicWebMassUpdateResultInfo("Value"),"~")
			End If
			For iCounter=0 to ubound(aImpactedParts)
				bFlag=False
				For iCount=0 to Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateResultImpactedParts").RowCount
					If trim(aImpactedParts(iCounter))=trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateResultImpactedParts").GetCellData(iCount,1)) Then
						If dicWebMassUpdateResultInfo("ColumnName")<>"" Then
							If dicWebMassUpdateResultInfo("ColumnName")="Status" Then
								iColNumber=2
							End If
							If trim(aValue(iCounter))=trim(Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MassUpdateResultImpactedParts").GetCellData(iCount,iColNumber)) Then
								bFlag=True
							End if
						Else
							bFlag=True
						End If
						Exit for
					End if
				Next
				If bFlag=False Then
					Exit for
				End If
			Next
			If bFlag=True Then
				Fn_SISW_Web_MassUpdateResultOperations=True
			End If
   End Select
   If StrButtonName<>"" Then
		Call Fn_Web_UI_Button_Click("Fn_SISW_Web_MassUpdateResultOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), StrButtonName)
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'    Function Name		:	Fn_SISW_Web_WhereUsedParentAssembliesTableOperations
'
'    Description				 :	Function Used to Perform Operations On Parent Assemblies table under where used tab
'
'    Parameters			   :	1.StrAction: Action Name
'												 2.StrObjectName : Object name
'												 3.StrColName : Column Name
'												 4.StrValue : Cell Value
'
'    Return Value		   	   : 	True Or False Or Column Number or Row number
'
'    Pre-requisite			:	Should be Log In Web Client							
'
'    Examples					:	bReturn=Fn_SISW_Web_WhereUsedParentAssembliesTableOperations("VerifyCellData","000036_4816/A;1-WheelAssy2013~000027_4816/A;1-WheelAssy2012","Type","Part Revision~Part Revision")
'												bReturn=Fn_SISW_Web_WhereUsedParentAssembliesTableOperations("VerifyCellData","000036_4816/A;1-WheelAssy2013~000027_4816/A;1-WheelAssy2012","","")
'												bReturn=Fn_SISW_Web_WhereUsedParentAssembliesTableOperations("VerifyCellData","000045_4816/A;1-WheelAssy2014","Owner","AutoTest1 (autotest1)")
'	   History					 	:	
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'												Sandeep Navghane									16-Nov-2012						1.0																								Sunny Ruparel
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Web_WhereUsedParentAssembliesTableOperations(StrAction,StrObjectName,StrColName,StrValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_WhereUsedParentAssembliesTableOperations"
	'Variable Declaration
	Dim ObjParentAssembliesTB,ObjChk,bFlag
	Dim iColCount,iCounter,ColName,arrColName,iRwCount,iRowNum,currCellValue,iCount,aObjectName,aValue
    'Creating Object Of Parent Assemblies table
	Fn_SISW_Web_WhereUsedParentAssembliesTableOperations=False
	Set ObjParentAssembliesTB=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("WhereUsedParentAssemblies")
	'Clicking On Where Used tab
	Call Fn_WEB_UI_Object_SetTOProperty("Fn_SISW_Web_WhereUsedParentAssembliesTableOperations",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Overview"),"innertext","Where Used *")
	Call Fn_Web_UI_WebElement_Click("Fn_SISW_Web_WhereUsedParentAssembliesTableOperations",Browser("TeamcenterWeb").Page("MyTeamCenter"),"Overview", "","","")
	Select Case StrAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Retrieve Row Index
		 Case "GetRowIndex" 
				iRwCount=ObjParentAssembliesTB.RowCount()
				For iCounter=1 To iRwCount
					currCellValue=ObjParentAssembliesTB.GetCellData(iCounter,2)
					If Trim(currCellValue)=Trim(StrObjectName) Then
						Fn_SISW_Web_WhereUsedParentAssembliesTableOperations=iCounter
						Exit For
					End If
				Next
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to get column index
			Case "GetColumnIndex"
					bFlag=False
					iColCount=ObjParentAssembliesTB.ColumnCount(1)
					For iCounter=1 To iColCount
						currCellValue=ObjParentAssembliesTB.GetCellData(1,iCounter)
						If Trim(currCellValue)=Trim(StrColName) Then
							iColCount=iCounter
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=True Then
						Fn_SISW_Web_WhereUsedParentAssembliesTableOperations=iColCount
					Else
						Fn_SISW_Web_WhereUsedParentAssembliesTableOperations=-1
					End If
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			'Case to veify Cell data
			Case "VerifyCellData"

				aObjectName=Split(StrObjectName,"~")
				If StrValue<>"" Then
					aValue=Split(StrValue,"~")
				End If
				For iCount=0 to ubound(aObjectName)
					bFlag=False
					iRwCount=ObjParentAssembliesTB.RowCount()
					For iCounter=1 To iRwCount
						currCellValue=ObjParentAssembliesTB.GetCellData(iCounter,1)
						If Trim(currCellValue)=Trim(aObjectName(iCount)) Then
							If StrValue<>"" Then
								iColCount=Fn_SISW_Web_WhereUsedParentAssembliesTableOperations("GetColumnIndex","",StrColName,"")
								If trim(ObjParentAssembliesTB.GetCellData(iCounter,iColCount))=trim(aValue(iCount)) Then
									bFlag=True
								End If
							Else
								bFlag=True
							End If
							Exit for
						End If
					Next
					If bFlag=False Then
						Exit for
					End If
				Next
				If bFlag=True Then
					Fn_SISW_Web_WhereUsedParentAssembliesTableOperations=True
				End If
	End Select
	'Releasing Object Of Parent Assemblies Table unde Where Used tab
	Set ObjParentAssembliesTB=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'    Function Name		:	Fn_SISW_Web_AssignFinishOperation
'
'    Description	       :	Function Used to Find and assign the Finish object
'
'    Parameters			   :	1.sAction: Action Name
'												 2.sObjName : Object name
'												 3.sItemID : 
'												 4.sDesc : 
'												 5.sType : 
'												 6.sOwnUser :
'												 7.sOwnGrp :
'
'    Return Value		   	   : 	True Or False
'
'    Pre-requisite			:	Should be Log In Web Client	and the Object to which the Assign Finish should be selected						
'
'    Examples					:	bReturn = Fn_SISW_Web_AssignFinishOperation("Find","*","","","","","","", "")
'                                        bReturn = Fn_SISW_Web_AssignFinishOperation("Find","Item1","","","","","<SKIP>","", "")

'	   History					 :	
'													Developer Name								Date						Rev. No.						Changes Done								Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'											     	Pritam Shikare								23-May-2013					1.0																						
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Web_AssignFinishOperation(sAction,sObjName,sItemID,sDesc,sType,sOwnUser,sOwnGrp,sRes1, sRes2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_AssignFinishOperation"
	Dim objAssgnFin,objSrchRes,objobjMDR,strWEBMenuPath,strMenu,iCounter
	Dim objElement, intIndex

   Fn_SISW_Web_AssignFinishOperation=False

   Set objAssgnFin = Fn_SISW_Web_GetObject("AssignFinish")
   Set objSrchRes = Fn_SISW_Web_GetObject("AssignFinishSrchRes")

	'Vallari [14Jun11] - Get the number of intances of New Item dialog
    Set objElement = Description.Create()
	objElement("micclass").Value = "WebElement"
	objElement("innertext").Value = "New Item"
	objElement("html tag").Value = "SPAN"
	intIndex =  Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objElement).count
	objAssgnFin.SetTOProperty "index", cstr(intIndex)
	Set objElement = Nothing


	If  Not objAssgnFin.Exist(7) Then
	'If New Item does not exist, Do menu operation
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "ToolsAssignFinish")
		Call Fn_Web_MenuOperation("Select",strMenu)
		Wait 5

		'Vallari [14Jun11] - Get the number of intances of New Item dialog and set the index for WebTable in OR accordingly
		intIndex =  Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objElement).count
		objAssgnFin.SetTOProperty "index", cstr(cint(intIndex)-1)
	End If

	'creating object of Mercury device replay
	Set objobjMDR = CreateObject("Mercury.DeviceReplay")

	Select Case (sAction)
		Case "Find"
			If sObjName<>"" Then
				objAssgnFin.WebEdit("Name").Object.focus
				objobjMDR.SendString sObjName
			End If

			If sItemID<>"" Then
				objAssgnFin.WebEdit("ItemID").Object.focus
				objobjMDR.SendString sItemID
			End If

			If sDesc<>"" Then
				objAssgnFin.WebEdit("Description").Object.focus
				objobjMDR.SendString sDesc
			End If

			If sType<>"<SKIP>" Then
				If sType<>"" Then
					Call Fn_Web_UI_Button_Click("Fn_SISW_Web_AssignFinishOperation",objAssgnFin,"Type")
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_SISW_Web_AssignFinishOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",sType)
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
				Else 
					objAssgnFin.WebEdit("Type").Set ""
				End If
			End If

			If sOwnUser<>"<SKIP>" Then
				objAssgnFin.WebEdit("OwningUser").Set sOwnUser
			End If

			If sOwnGrp<> "<SKIP>" Then
				If  sOwnGrp <> "" Then
					Call Fn_Web_UI_Button_Click("Fn_SISW_Web_AssignFinishOperation",objAssgnFin,"OwningGroup")
					Call Fn_WEB_UI_Object_SetTOProperty("Fn_SISW_Web_AssignFinishOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",sOwnGrp)
					Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
				Else
					objAssgnFin.WebEdit("OwningGroup").Set ""
				End If
			End If
			Call Fn_Web_UI_Button_Click("Fn_SISW_Web_AssignFinishOperation", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Search")
			Fn_SISW_Web_AssignFinishOperation = true
		'-----------------------------------------------------------------------------------------------------------------------------------------------------
		Case "SelectSearchResult"

			If objSrchRes.Exist Then
				objSrchRes.WebTable("Results").WebElement("ObjectName").SetTOProperty "innertext",sObjName
				If objSrchRes.WebTable("Results").WebElement("ObjectName").Exist(5) Then
					objSrchRes.WebTable("Results").WebCheckBox("ObjectChkBox").Set "ON"
				Else
					Fn_SISW_Web_AssignFinishOperation =FALSE
				End If
				Call Fn_Web_UI_Button_Click("Fn_SISW_Web_AssignFinishOperation", Browser("TeamcenterWeb").Page("MyTeamCenter"), "OKSearchResults")
				Fn_SISW_Web_AssignFinishOperation =True
			Else
				Fn_SISW_Web_AssignFinishOperation =FALSE
			End If
		'-----------------------------------------------------------------------------------------------------------------------------------------------------
	End Select
	Set objAssgnFin =Nothing
   Set objSrchRes = Nothing
   Set objobjMDR = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'    Function Name		:	Fn_SISW_Web_NewAuditCreate
'
'    Description		   :	Function Used to create the New Audit
'
'    Parameters			  :	   1.StrAction    : Action Name
'									 2.sAuditNum :  Audit Number    ( format CA-nnnnnn)
'									 3.sType         : Type 
'									 4.sRev          : Revision 
'									 5.sSynopsis   : Name given to the Audit
'									 6.sDesc         : Description
'									 7.sAuditType : Audit Type, chose from the dropdown
'								     8.sDate         : Date, Month/dd/YYYY   eg. November/12/2013,  February/28/2014 etc
'								     9.sComments : Comments
'								     10, 11. sRes1, sRes2 : Reserved for the future use
'								     12. sBtn : Button name

'    Return Value		  : 	True Or False 
'
'    Pre-requisite		   :	Should be Log In Web Client							
'
'    Examples			   :	bReturn=Fn_SISW_Web_NewAuditCreate("","Configuration Audit","CA-000123","A","CRAudit","asdf","","Today","comments","","","Finish")
'	
'	   History				 :	
'													Developer Name									Date						Rev. No.						Changes Done								Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'											     	Pritam Shikare									20-May-2013					1.0																								Sunny Ruparel
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_Web_NewAuditCreate(sAction,sType,sAuditNum,sRev,sSynopsis,sDesc,sAuditType,sDate,sComments,sRes1, sRes2, sBtn)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_NewAuditCreate"
	Dim ObjNewAudit,objobjMDR,strWEBMenuPath,strMenu,crrType,iCount, aBtn, aDate
	Dim objElement, intIndex

    Fn_SISW_Web_NewAuditCreate=False
	Set ObjNewAudit = Fn_SISW_Web_GetObject("NewAudit")
	
	'Get the number of intances of New Item dialog and set the index for WebTable in OR accordingly
	Set objElement = Description.Create()
	objElement("micclass").Value = "WebElement"
	objElement("innertext").Value = "New Item"
	objElement("html tag").Value = "SPAN"
	intIndex =  Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objElement).count
	ObjNewAudit.SetTOProperty "index", cstr(Cint(intIndex)-1)

	If  Not ObjNewAudit.Exist(7) Then
	'If New Item does not exist, Do menu operation
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewAudit")
		Call Fn_Web_MenuOperation("Select",strMenu)

		'Vallari [14Jun11] - Get the number of intances of New Item dialog and set the index for WebTable in OR accordingly
		intIndex =  Browser("TeamcenterWeb").Page("MyTeamCenter").ChildObjects(objElement).count
		ObjNewAudit.SetTOProperty "index", cstr(cint(intIndex)-1)
	End If

	Set objElement = Nothing

	'Select the Type
	If sType<>"" Then
		crrType=ObjNewAudit.WebEdit("TypeEdit").GetROProperty("value")
		wait(1)
		If Trim(crrType)<>Trim(sType) Then
				'Setting Item Type
				Call Fn_Web_UI_Button_Click("Fn_Web_ItemBasicCreate",ObjNewAudit,"Type")
				Call Fn_WEB_UI_Object_SetTOProperty("Fn_SISW_Web_NewAuditCreate",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType"),"innertext",sType)
				Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click 1,1,micLeftBtn
				wait(2)
		End If
	End If
	'creating object of Mercury device replay
	Set objobjMDR = CreateObject("Mercury.DeviceReplay")

	Call Fn_Web_UI_Button_Click("Fn_SISW_Web_NewAuditCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "Next")
	wait(1)
	'Fill in the Audit Number
	If sAuditNum <>"" Then
		ObjNewAudit.WebTable("AuditInfo").WebEdit("AuditNumber").Object.focus
		objobjMDR.SendString sAuditNum
	End If
		Wait 0, 300
	'Fill in the Revision
	If sRev<>"" Then
		ObjNewAudit.WebTable("AuditInfo").WebEdit("Revision").Object.focus
		objobjMDR.SendString sRev
	End If
		Wait 0, 300
	'Fill in the synopsis field
	If sSynopsis<>"" Then
		ObjNewAudit.WebTable("AuditInfo").WebEdit("Synopsis").Object.focus
		objobjMDR.SendString sSynopsis
	End If
		Wait 0, 300
	'Fill in the Description
	If sDesc<>"" Then
		ObjNewAudit.WebTable("AuditInfo").WebEdit("Description").Object.focus
		objobjMDR.SendString sDesc
	End If
'				Wait 0, 300
	'Select the Audit Type
	If sAuditType<>"" Then
		Wait 5
		ObjNewAudit.WebTable("AuditInfo").WebButton("Type").Click
		Wait 5
		Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").SetTOProperty "innertext",sAuditType
		Wait 1
		If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Exist Then
			Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FormType").Click
			wait 1
		End If
		Wait 1
	End If

	'Fill in the comments
	If sComments<>"" Then
		ObjNewAudit.WebTable("AuditInfo").WebEdit("Comments").Object.focus
		objobjMDR.SendString sComments
	End If

	'Split the Date and fill in the date field
	If sDate<>"" Then
		If sDate= "Today" Then
			aDate = Split(Cstr(Date),"/",-1,1)
			aDate(0) = Cstr(MonthName(aDate(0)))
		Else
			aDate = Split(sDate,"/",-1,1)
		End If
		ObjNewAudit.WebTable("AuditInfo").WebList("Month").Select aDate(0)
		Wait 1
		ObjNewAudit.WebTable("AuditInfo").WebList("Date").Select aDate(1)
		Wait 1
		ObjNewAudit.WebTable("AuditInfo").WebList("Year").Select aDate(2)
		Wait 1
	End If

	'Press the Buttons after the Fields are populated
	If sBtn<>"" Then
		aBtn = Split(sBtn,":",-1,1)
		For iCount=0 to Ubound(aBtn)
			Call Fn_Web_UI_Button_Click("Fn_SISW_Web_NewAuditCreate", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), aBtn(iCount))
			Wait 3
		Next
	End If

	'Wait for the Dialog to be close
	For iCount=0 To 2
		If ObjNewAudit.Exist(5) Then
			wait(5)
		Else
			Exit For
		End If
	Next

	Fn_SISW_Web_NewAuditCreate=True
	Set objobjMDR =Nothing
	Set ObjNewAudit=Nothing

End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'    Function Name		:	Fn_SISW_Web_ProjectDataTabOperation
'
'    Description		   :	Function Used to Perofrm the operations in ProjectData tab for the project
'
'    Parameters			  :	   1.StrAction    : Action Name
'									 2.dicProjectData : Dictionary object for the data parameters

'    Return Value		  : 	True Or False 
'
'    Pre-requisite		   :	Should be Log In Web Client							
'									 
'    Examples			   :	           Set dicProjectData = CreateObject("Scripting.Dictionary")
'												dicProjectData("Action") =  "Verify"
'												dicProjectData("ObjectName") =  "000033-Item1"
'												bReturn=Fn_SISW_Web_ProjectDataTabOperation("PreferredItems",dicProjectData)

'												dicProjectData("Action") =  "Paste"
'												bReturn=Fn_SISW_Web_ProjectDataTabOperation("ProgramData",dicProjectData)

'												dicProjectData("Action") =  "Cut"
'												dicProjectData("ObjectName") =  "000033-Item1"
'												bReturn=Fn_SISW_Web_ProjectDataTabOperation("PreferredItems",dicProjectData)

'												
'                                               dicItemDetailsCreate("ItemType") = "Part"
'												dicItemDetailsCreate("Name") = "TestPart"
'												dicItemDetailsCreate("ID") = "ASSIGN"
'												Set dicProjectData = CreateObject("Scripting.Dictionary")
'												dicProjectData("Action") =  "AddNew"
'												Set dicProjectData("ItemDetails") =  dicItemDetailsCreate
'												bReturn=Fn_SISW_Web_ProjectDataTabOperation("PreferredItems",dicProjectData)
'	
'	   History				 :	
'													Developer Name						Date				Rev. No.		Changes Done															Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'											     	Pritam Shikare					28-May-2013					1.0																		           Sandeep
'													Shailendra sahu					30-May-2013					1.1		Added Case "Cut"&"Paste" for case "PreferredItems"&"ProgramData"		   Pritam
'													Shailendra sahu					31-May-2013					1.1		Modified Case : "Cut"&"Paste" for case "PreferredItems"&"ProgramData"	   Pritam
'											     	Pritam Shikare					04-June-2013				1.2		Added Case : "AddNew"	for case "ProgramData"																																			
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Function Fn_SISW_Web_ProjectDataTabOperation(strAction,dicProjectData)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_ProjectDataTabOperation"
   Dim iRowindex,objProjectDataTab
'   Set objProjectDataTab=Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("Project Data")
   Fn_SISW_Web_ProjectDataTabOperation = False
   'checking existence of Project Data Tab and then clicking on that
	   Call Fn_Web_TabOperations("Activate","Project Data")
   Select Case strAction
	 	Case "PreferredItems"
	 			If Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("PreferredItemsPanel").WebTable("PrefferedItems").Exist Then
	 				iRowindex = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("PreferredItemsPanel").WebTable("PrefferedItems").GetRowWithCellText(dicProjectData("ObjectName"))
	 			Else
	 				iRowindex = -1
	 			End If
				Select Case dicProjectData("Action")
					Case "Verify"
						If iRowIndex > 0 Then
							Fn_SISW_Web_ProjectDataTabOperation = True
						End If
					Case "Cut"
						Set obj=Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("PreferredItemsPanel").WebTable("PrefferedItems").ChildItem(iRowindex,1,"WebElement",0)
						If obj.getROProperty("innertext")=dicProjectData("ObjectName") Then
							obj.click 1,1
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("PreferredItemsPanel").Link("Cut").Click
							Fn_SISW_Web_ProjectDataTabOperation = True
						End If
					Case "Paste"
							Call Fn_Web_UI_Button_Click("Fn_SISW_Web_ProjectDataTabOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("PreferredItemsPanel"),"Paste")
							Fn_SISW_Web_ProjectDataTabOperation = True
					Case "AddNew"
						'Developed soon
				End Select

		Case "ProgramData"
				iRowindex = Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ProgramDataPanel").WebTable("ProgramData").GetRowWithCellText(dicProjectData("ObjectName"))
				Select Case dicProjectData("Action")
					Case "Verify"
						If iRowIndex > 0 Then
							Fn_SISW_Web_ProjectDataTabOperation = True
						End If
					Case "Cut"
						Set obj=Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ProgramDataPanel").WebTable("ProgramData").ChildItem(iRowindex,1,"WebElement",0)
						If obj.getROProperty("innertext")=dicProjectData("ObjectName") Then
							obj.click 1,1
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ProgramDataPanel").Link("Cut").Click
							Fn_SISW_Web_ProjectDataTabOperation = True
						End If
					Case "Paste"
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ProgramDataPanel").Link("Paste").Click
							Fn_SISW_Web_ProjectDataTabOperation = True
					Case "AddNew"
						Call Fn_Web_UI_Button_Click("Fn_SISW_Web_ProjectDataTabOperation",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ProgramDataPanel"),"Add New...")
						If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem").Exist(5) Then
							Fn_SISW_Web_ProjectDataTabOperation = Fn_Web_ItemDetailsCreate(dicProjectData("ItemDetails"))
						End If
				End Select
   End Select
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'    Function Name		:	Fn_SISW_Web_MakeFromOperations
'
'    Description		   :	Function Used to Perofrm the Make From operation from Toold>> Makefrom on the object
'
'    Parameters			  :	   1.StrAction    : Action Name
'									 2.dicMakeFrom : Dictionary object for the data parameters

'    Return Value		  : 	True Or False 
'
'    Pre-requisite		   :	Should be Log In Web Client							
'									 
'    Examples			   :	           Set dicMakeFrom = CreateObject("Scripting.Dictionary")
'												 dicMakeFrom("MakeFrom") =  "Stock Material"
'                                                dicMakeFrom("SelStockMaterial") = "0123"            'use ID of the stock material
'												 bReturn=Fn_SISW_Web_ProjectDataTabOperation("PreferredItems",dicProjectData)

'     Note                   :       For fields use folowing names for the Dictionary parameters 
										'Make From =>  dicMakeFrom("MakeFrom")
										'Select Stock Material => dicMakeFrom("SelStockMaterial")
										'Dimensions Used =>  dicMakeFrom("DimensionsUsed")
										'Cut Length => dicMakeFrom("CutLength")
										'Cut Width => dicMakeFrom("CutWidth")
										'Cut Thickness => dicMakeFrom("CutThickness")
										'Stock Quantity => dicMakeFrom("StockQuantity")
										'Unit Of Measure => dicMakeFrom("UnitOfMeasure")
'	
'	   History				 :	
'													Developer Name									Date						Rev. No.						Changes Done								Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'											     	Pritam Shikare									28-May-2013					1.0																				
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_Web_MakeFromOperations(strAction,dicMakeFrom)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_MakeFromOperations"
	Dim strWEBMenuPath, strMenu
	Fn_SISW_Web_MakeFromOperations = False
	Set objMakeFrm = Fn_SISW_Web_GetObject("MakeFrom")


	If  Not objMakeFrm.Exist(7) Then
	'If New Item does not exist, Do menu operation
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "ToolsMakeFrom")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	
	Do
	Select Case strAction

		Case "Assign"
				akeys = dicMakeFrom.Keys
				aItems = dicMakeFrom.Items
				For iCounter = 0 to Ubound(aKeys)
					If  dicMakeFrom(akeys(iCounter)) <> "" Then
						If  akeys(iCounter) = "SelStockMaterial" Then
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MakeFrom").Image("SelStockMaterial").Click 1,1
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("FindItem").WebEdit("ItemId").Set aItems(iCounter)
							Browser("TeamcenterWeb").Page("MyTeamCenter").WebButton("Search").Click
							If Browser("TeamcenterWeb").Dialog("Dialog").Exist(2) Then
								wait 1
								Browser("TeamcenterWeb").Dialog("Dialog").WinButton("OK").Click
								Fn_SISW_Web_MakeFromOperations = False
								Exit Do
							End If
						Else						
						  Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("MakeFrom").WebEdit(aKeys(iCounter)).Set aItems(iCounter)
						End If
						
					End If
				Next
				Call Fn_Web_UI_Button_Click("",Browser("TeamcenterWeb").Page("MyTeamCenter"),"OK")
				Fn_SISW_Web_MakeFromOperations = True

		Case Else
				Fn_SISW_Web_MakeFromOperations = False

	End Select
	Exit Do: Loop

Set objMakeFrm = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Web_NewItemAssignToProgramOperations

'Description			 :	Function Used to Perform operations on Assign Projects

'Parameters			   :   '1.StrAction: Action Name
'										 2.dicNewItemProgramInfo: Assign to program Information
'										 3.StrButton: Button Name
'
'Return Value		   : 	True Or False

'Pre-requisite			:	Should be log in Thin Client

'Examples				:   dicNewItemProgramInfo("AvailableProjects")="Projects A~Projects B"
'										Msgbox Fn_SISW_Web_NewItemAssignToProgramOperations("Add",dicNewItemProgramInfo,"")
'										dicNewItemProgramInfo("SelectedProjetcs")="Projects A~Projects B"
'										Msgbox Fn_SISW_Web_NewItemAssignToProgramOperations("Remove",dicNewItemProgramInfo,"")
'										dicNewItemProgramInfo("AvailableProjects")="Projects A~Projects B"
'										Msgbox Fn_SISW_Web_NewItemAssignToProgramOperations("VerifyAvailableProjects",dicNewItemProgramInfo,"")
'										dicNewItemProgramInfo("SelectedProjetcs")="Projects A~Projects B"
'										Msgbox Fn_SISW_Web_NewItemAssignToProgramOperations("VerifySelectedProjects",dicNewItemProgramInfo,"")

'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												29-May-2013								1.0																					Rima P
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_Web_NewItemAssignToProgramOperations(StrAction,dicNewItemProgramInfo,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_NewItemAssignToProgramOperations"
 	'Variable Declaration
    Dim strWEBMenuPath,strMenu,ObjAssignProgram
	Dim arrProjects,iCounter,iCounter1,iCount,bFlag,crrProject
	'Function Returns False
    Fn_SISW_Web_NewItemAssignToProgramOperations=False
	'Creating Object of [ AssignToProgram ] table
	Set ObjAssignProgram=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("NewItem").WebTable("AssignToProgram")
	'Checking Existance of [ AssignToProgram ] table
	If Not ObjAssignProgram.Exist(6) Then
		'Calling menu : "Tools:Project:Assign"
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewItem")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If
	Select Case StrAction
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        'Case To click on button
        Case "ClickButton"
            Fn_SISW_Web_NewItemAssignToProgramOperations=Fn_Web_UI_Button_Click("Fn_Web_ItemBasicCreate", ObjAssignProgram, StrButton)
       '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case To Assign Program to Object
		Case "Add"
			arrProjects=Split(dicNewItemProgramInfo("AvailableProjects"),"~")
			For iCounter=0 To UBound(arrProjects)
				Call Fn_Web_UI_List_Select("Fn_SISW_Web_NewItemAssignToProgramOperations",ObjAssignProgram, "ProgramsForSelection",arrProjects(iCounter))
				Fn_SISW_Web_NewItemAssignToProgramOperations=Fn_Web_UI_Button_Click("Fn_SISW_Web_NewItemAssignToProgramOperations", ObjAssignProgram, "Add")
			Next
            if StrButton<>"" then
			    Call Fn_Web_UI_Button_Click("Fn_SISW_Web_NewItemAssignToProgramOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),StrButton)
            End if
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case To Remove assign projects from Object
		Case "Remove"
			arrProjects=Split(dicNewItemProgramInfo("SelectedPrograms"),"~")
			For iCounter=0 To UBound(arrProjects)
				Call Fn_Web_UI_List_Select("Fn_SISW_Web_NewItemAssignToProgramOperations",ObjAssignProgram, "SelectedPrograms",arrProjects(iCounter))
				Fn_SISW_Web_NewItemAssignToProgramOperations=Fn_Web_UI_Button_Click("Fn_SISW_Web_NewItemAssignToProgramOperations", ObjAssignProgram, "Remove")
			Next
			if StrButton<>"" then
			    Call Fn_Web_UI_Button_Click("Fn_SISW_Web_NewItemAssignToProgramOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),StrButton)
            End if
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Verify Avaliable Projects
		Case "VerifyAvailableProjects"
			arrProjects=Split(dicNewItemProgramInfo("AvailableProjects"),"~")
			iCounter=ObjAssignProgram.WebList("ProgramsForSelection").GetROProperty("items count")
			For iCounter1=0 To UBound(arrProjects)
				bFlag=False
				For iCount=1 To iCounter
					crrProject=ObjAssignProgram.WebList("ProgramsForSelection").GetItem(iCount)
					If Trim(crrProject)=arrProjects(iCounter1) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_SISW_Web_NewItemAssignToProgramOperations=True
			End If
			if StrButton<>"" then
			    Call Fn_Web_UI_Button_Click("Fn_SISW_Web_NewItemAssignToProgramOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),StrButton)
            End if
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to Verify Selected Projects
		Case "VerifySelectedProjects"
			arrProjects=Split(dicNewItemProgramInfo("SelectedPrograms"),"~")
			iCounter=ObjAssignProgram.WebList("SelectedPrograms").GetROProperty("items count")
			For iCounter1=0 To UBound(arrProjects)
				bFlag=False
				For iCount=1 To iCounter
					crrProject=ObjAssignProgram.WebList("SelectedPrograms").GetItem(iCount)
					If Trim(crrProject)=arrProjects(iCounter1) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_SISW_Web_NewItemAssignToProgramOperations=True
			End If
			if StrButton<>"" then
			    Call Fn_Web_UI_Button_Click("Fn_SISW_Web_NewItemAssignToProgramOperations", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"),StrButton)
            End if
	End Select
	'Releasing Object of [ AssignToProgram ] table
	Set ObjAssignProgram=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Web_PropertiesOnRelation()
'
'Description			 :	Function Used to perform operations like "Verify" on "Enter Properties on:Made From" table

'Parameters			   :  StrAction: valid Action Name (Mandatory)
'								  dicProperties: dictionary object(Mendatory)
'								  StrButton : Valid Button (Separated by : )
'
'Return Value		   : 	True or False

'Pre-requisite			:	 Login to Webclient and the Object on which the Properties On Realation are  to verified should be selected

'Examples				:  	 Set dicProperties=CreateObject("Scripting.Dictionary")
'									dicProperties("Fields")="Cut Length~Cut Thickness"
'									dicProperties("ExpectedValues")="1~1"
'	'								Fn_SISW_Web_PropertiesOnRelation("Verify",dicProperties,"OK")			

									'Set dicProperties=CreateObject("Scripting.Dictionary")
'									dicProperties("Fields")="Cut Length~Cut Thickness~Relation Type"
'									dicProperties("ExpectedValues")="12~1~Made From"
'									Fn_SISW_Web_PropertiesOnRelation("Verify",dicProperties,"OK")
'									Fn_SISW_Web_PropertiesOnRelation("Verify",dicProperties,"Apply:OK")
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done																				Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Shailendra Sahu						29-May-2013				1.0																																		Pritam
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


Function Fn_SISW_Web_PropertiesOnRelation(StrAction,dicProperties,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_PropertiesOnRelation"
   'declaring variable
   Dim objDialog,objTable,aFields,aExpectedValues,iFlag,iIndex,iCount,StrCurrValue
   Dim strMenu,strWEBMenuPath,aButton

    iFlag=0
    Fn_SISW_Web_PropertiesOnRelation=False

	'setting objects
	Set objDialog=Browser("TeamcenterWeb").Page("MyTeamCenter")
	Set objTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("PropertiesOnRelation")
	
	If  Not objTable.Exist(5) Then
	'If New Item does not exist, Do menu operation
		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "EditPropertiesOnRelation")
		Call Fn_Web_MenuOperation("Select",strMenu)
	End If

	'Fetching values from dictionary objects
	If  objTable.Exist(2) Then
		aFields=split(dicProperties("Fields"),"~")
		aExpectedValues=split(dicProperties("ExpectedValues"),"~")
		'selecting the Action
		Select Case StrAction
			Case "Verify"
				For iCount=0 to Ubound(aFields)
					'Fetching position of specified object (row number)
					iIndex=objTable.GetRowWithCellText(aFields(iCount)+":")
					'Checking number of child object
					If objTable.ChildItemCount(iIndex,2,"WebEdit")="1" Then
						'getting value of web edit box
						StrCurrValue=objTable.ChildItem(iIndex,2,"WebEdit",0).getROProperty("value")
							'verifing expected and actual values
							If StrCurrValue=Cstr(aExpectedValues(iCount)) Then
								iFlag=iFlag+1
							End If
					Else
						'getting cell data
						StrCurrValue=objTable.GetCellData(iIndex,2)
							'verifing expected and actual values
							If StrCurrValue=Cstr(aExpectedValues(iCount)) Then
								iFlag=iFlag+1
							End If
					End If
				Next

				If iFlag=Ubound(aExpectedValues)+1 Then
					Fn_SISW_Web_PropertiesOnRelation=true
				Else
					Fn_SISW_Web_PropertiesOnRelation=False
				End If
	'--------------------------------------------------------------------------------------
		End Select
	    'Clicking on button
		aButton=Split(StrButton,":")
		For iCount=0 to Ubound(aButton)
			Call Fn_Web_UI_Button_Click("Fn_SISW_Web_PropertiesOnRelation",objDialog.WebElement("ButtunPanel"),aButton(iCount))
			Wait 1
		Next
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Property Relation table doesn't exist")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'    Function Name		:	Fn_SISW_Web_EditStandardNoteOperations
'
'    Description				 :	Function Used to Perform Operations On Edit Parameters Note
'
'    Parameters			   :	1.StrAction: Action Name
'												 2.dicStandardNoteInfo : Edit Standard Note information
'												 3.StrButton : Button Name
'
'    Return Value		   	   : 	True Or False
'
'    Pre-requisite			:		Should be Log In Web Client							
'
'    Examples					:	Dim dicStandardNoteInfo
'												Set dicStandardNoteInfo=CreateObject("Scripting.Dictionary")
'
'												dicStandardNoteInfo("Temperature")="98.5"
'												dicStandardNoteInfo("humidity")="4.5"
'												bReturn=Fn_SISW_Web_EditStandardNoteOperations("SetParameters",dicStandardNoteInfo,"OK")
'												
'												dicStandardNoteInfo("Temperature")="98.5"
'												dicStandardNoteInfo("humidity")="4.50"
'												bReturn=Fn_SISW_Web_EditStandardNoteOperations("VerifyParameters",dicStandardNoteInfo,"")
'												
'												dicStandardNoteInfo("Note Text")="Temparature 98.5 and humidity 4.5"
'												bReturn=Fn_SISW_Web_EditStandardNoteOperations("Verify",dicStandardNoteInfo,"")
'	   History					 	:	
'													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'												Sandeep Navghane									31-May-2013						1.0																								Veena G
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Web_EditStandardNoteOperations(StrAction,dicStandardNoteInfo,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_EditStandardNoteOperations"
 	'Declaring variables
	Dim objNoteTable,objParaTable,objWebEdit
	Dim DictItems,DictKeys,iCounter,sAction,iCount,bFlag
	Fn_SISW_Web_EditStandardNoteOperations=False
	'Checking existance of [ Edit Standard Note ] dialog
	If Not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("EditStandardNote").Exist(10) Then
		'calling menu [ Edit = > Standard Note Parameters ]
		Call Fn_Web_MenuOperation("Select","Edit:Standard Note Parameters")
		wait 3
	End If
	'creating object of [ EditStandardNote ] dialog
	Set objNoteTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("EditStandardNote")
	DictItems = dicStandardNoteInfo.Items
	DictKeys = dicStandardNoteInfo.Keys
	For iCounter=0 to dicStandardNoteInfo.count-1
		If DictItems(iCounter)<>"" Then
			sAction=DictKeys(iCounter)
			Select Case StrAction
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "SetParameters"
					Set objParaTable=objNoteTable.WebTable("ParameterValueTable")
					Select Case sAction
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
						Case "Temperature","humidity"
							For iCount=0  to objParaTable.GetROProperty("rows")
								bFlag=False
								If DictKeys(iCounter)=objParaTable.GetCellData(iCount,1) or DictItems(iCounter)+":"=objParaTable.GetCellData(iCount,1) Then
									Set objWebEdit = objParaTable.ChildItem(iCount, 2, "WebEdit", 0)
									If TypeName(objWebEdit) <> "Nothing" Then
										objWebEdit.Set DictItems(iCounter)+vbLf
										wait 1
										bFlag=True
									Else
										bFlag=False
									End if
									Set objWebEdit =Nothing
									Exit for
								End If
							Next
							If bFlag=False Then
								Fn_SISW_Web_EditStandardNoteOperations=False
								Exit function
							End If
					End Select
					If bFlag=True Then
						Fn_SISW_Web_EditStandardNoteOperations=True
					End If
					Set objParaTable=Nothing
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "VerifyParameters"
					Set objParaTable=objNoteTable.WebTable("ParameterValueTable")
					Select Case sAction
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Temperature","humidity"
							For iCount=0  to objParaTable.GetROProperty("rows")
								bFlag=False
								If DictKeys(iCounter)=objParaTable.GetCellData(iCount,1) or DictItems(iCounter)+":"=objParaTable.GetCellData(iCount,1) Then
									Set objWebEdit = objParaTable.ChildItem(iCount, 2, "WebEdit", 0)
									If TypeName(objWebEdit) <> "Nothing" Then
										If DictItems(iCounter)="{BLANK}" Then
											DictItems(iCounter)=""
										End if
										If objWebEdit.GetROProperty("value")=DictItems(iCounter) then
											wait 1
											bFlag=True
										End if
									Else
										bFlag=False
									End if
									Set objWebEdit =Nothing
									Exit for
								End If
							Next
							If bFlag=False Then
								Fn_SISW_Web_EditStandardNoteOperations=False
								Exit function
							End If
					End Select
					If bFlag=True Then
						Fn_SISW_Web_EditStandardNoteOperations=True
					End If
					Set objParaTable=Nothing
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "Set"
					Select Case sAction
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Note Text"
							For iCount=0  to objNoteTable.GetROProperty("rows")-1
								bFlag=False
								If DictKeys(iCounter)=trim(objNoteTable.GetCellData(iCount,1)) or DictKeys(iCounter)+":"=trim(objNoteTable.GetCellData(iCount,1)) Then
									Set objWebEdit = objNoteTable.ChildItem(iCount, 2, "WebEdit", 0)
									If TypeName(objWebEdit) <> "Nothing" Then
										objWebEdit.Set DictItems(iCounter)+vbLf
										wait 1
										bFlag=True
									Else
										bFlag=False
									End if
									Set objWebEdit =Nothing
									Exit for
								End If
							Next
							If bFlag=False Then
								Fn_SISW_Web_EditStandardNoteOperations=False
								Exit function
							End If
					End Select
					If bFlag=True Then
						Fn_SISW_Web_EditStandardNoteOperations=True
					End If
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
				Case "Verify"
					Select Case sAction
						' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
						Case "Note Text"
							For iCount=0  to objNoteTable.GetROProperty("rows")-1
								bFlag=False
								If DictKeys(iCounter)=trim(objNoteTable.GetCellData(iCount,1)) or DictKeys(iCounter)+":"=trim(objNoteTable.GetCellData(iCount,1)) Then
									Set objWebEdit = objNoteTable.ChildItem(iCount, 2, "WebEdit", 0)
									If TypeName(objWebEdit) <> "Nothing" Then
										If DictItems(iCounter)="{BLANK}" Then
											DictItems(iCounter)=""
										End if
										If objWebEdit.GetROProperty("value")=DictItems(iCounter) then
											wait 1
											bFlag=True
										End if
									Else
										bFlag=False
									End if
									Set objWebEdit =Nothing
									Exit for
								End If
							Next
							If bFlag=False Then
								Fn_SISW_Web_EditStandardNoteOperations=False
								Exit function
							End If
					End Select
					If bFlag=True Then
						Fn_SISW_Web_EditStandardNoteOperations=True
					End If
			End Select
		End if
	Next
	If StrButton<>"" Then
		Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton(StrButton).Click
		wait 2
	End If
	Set objNoteTable=Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Web_FinishesTabOperation()
'
'Description			 :	Function Used to perform operations like "Verify", Add Finish, Remove Finish on "Finishes" tab on the Finish Group item 

'Parameters			   :    StrAction: valid Action Name     "Finishes" or "FinishSequence"
'								  dicFinishData: dictionary object 
'
'Return Value		   : 	True or False

'Pre-requisite			:	 Login to Webclient and the Object on which the function is to be performed should be selected

'Examples				:  	  Set dicFinishData=CreateObject("Scripting.Dictionary")
'									dicFinishData("Action")="Verify"
'									dicFinishData("ObjectName")="FIN-000026-FinItem1"
'	'								Fn_SISW_Web_FinishesTabOperation("Finishes",dicFinishData)			

'									Set dicFinishData=CreateObject("Scripting.Dictionary")
'									dicFinishData("Action")="RemoveFinish"
'									dicFinishData("ObjectName")="FIN-000026-FinItem1"
'	'								Fn_SISW_Web_FinishesTabOperation("Finishes",dicFinishData)	

'									Set dicFinishData=CreateObject("Scripting.Dictionary")
'									dicFinishData("Action")="AddFinish"
'									dicFinishData("FinishNameToAdd") = "FinItem1"
'									dicFinishData("ObjectName")="FIN-000026-FinItem1"
'	'								Fn_SISW_Web_FinishesTabOperation("Finishes",dicFinishData)	
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Pritam Shikare						06-June-2013			1.0																			  Veena
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function Fn_SISW_Web_FinishesTabOperation(strAction,dicFinishData)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_FinishesTabOperation"
   Dim iRowindex,obFinishesPanel

   Set obFinishesPanel=Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("FinishesPanel")
   Fn_SISW_Web_FinishesTabOperation = False
   'checking existence of Project Data Tab and then clicking on that
   Call Fn_Web_TabOperations("Activate","Finishes")

   Select Case strAction
	 	'------------------------------------------Case to Perform the Operations on the Finishes Panel' Add Finish, Remove Finish, Verify the Object in the table etc.---------
	 	Case "Finishes"
				iRowindex = obFinishesPanel.WebTable("FinishesTable").GetRowWithCellText(dicFinishData("ObjectName"))
				Select Case dicFinishData("Action")
					'---------Verify from the Table---------------------------
					Case "Verify"
						If Not(iRowIndex < 0) Then
							Set obj=obFinishesPanel.WebTable("FinishesTable").ChildItem(iRowindex,1,"WebElement",0)
							If obj.getROProperty("innertext")=dicFinishData("ObjectName") Then
								Fn_SISW_Web_FinishesTabOperation = True
							End If
						End If
						Set obj= Nothing
					'------------Remove Finish item from the Table---------
					Case "RemoveFinish"
						Set obj=obFinishesPanel.WebTable("FinishesTable").ChildItem(iRowindex,1,"WebElement",0)
						If obj.getROProperty("innertext")=dicFinishData("ObjectName") Then
							obj.click 1,1
							obFinishesPanel.Link("RemoveFinish").Click
							Fn_SISW_Web_FinishesTabOperation = True
						End If
						Set obj= Nothing
					'------------Add Finish item to the Table---------
					Case "AddFinish"
							obFinishesPanel.Link("AddFinish").Click
							'Develop further to add the Finish Items
							Fn_SISW_Web_FinishesTabOperation = True
							Call Fn_SISW_Web_AssignFinishOperation("Find",dicFinishData("FinishNameToAdd"),"","","Finish","","","","")
							Wait 2
							bReturn =  Fn_SISW_Web_AssignFinishOperation("SelectSearchResult",dicFinishData("ObjectName"),"","","","","","","")
							Wait 2
							Fn_SISW_Web_FinishesTabOperation = bReturn
					Case "Else"
							Exit Function
				End Select
		'----------------------------------------Case to perform the Operation on the Finish Sequence Panel----------------------------------------------------
		Case "FinishSequence"
				'Develop as per Requirement
   End Select

End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISW_Web_BOMComapreOperations()
'
'Description			 :	Function Used to perform operations of BOM Compare

'Parameters			   :    StrAction: Valid Action Name
'								  		 dicBOMCompareInfo: BOM Compare information
'								  		 StrButton: Button Name
'
'Return Value		   : 	True or False

'Pre-requisite			:	 Assembly should be selected

'Examples				:  	 Dim dicBOMCompareInfo
'										Set dicBOMCompareInfo=CreateObject("Scripting.Dictionary")
'										dicBOMCompareInfo("BOMLine1RevisionRule")="Latest Working"
'										dicBOMCompareInfo("BOMLine2RevisionRule")="Latest Working"
'										dicBOMCompareInfo("BOMCompareMode")="Single level (with substitutes)"
'										bReturn=Fn_SISW_Web_BOMComapreOperations("Compare",dicBOMCompareInfo,"OK")
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep Navghane				15-July-2013			1.0																			  		Gaurav S
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISW_Web_BOMComapreOperations(StrAction,dicBOMCompareInfo,StrButton)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Web_BOMComapreOperations"
   Dim ObjBOMCompareDialog,wshShell
   Fn_SISW_Web_BOMComapreOperations=False
   'checking existance of  [ BOMCompare ] dialog
   If not Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BOMCompare").Exist(6) Then
		Call Fn_Web_MenuOperation("Select","Actions:BOM Compare...")
		Call Fn_Web_ReadyStatusSync(1)
   End If
   'Creating object of [ BOMCompare ] dialog
	If Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BOMCompare").Exist(6) Then
		Set ObjBOMCompareDialog=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("BOMCompare")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : BOM Compare dialog not exist")
		Exit function
	End If
   Select Case StrAction
	 	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "Compare"
			'Setting BOM Line 1
			If dicBOMCompareInfo("BOMLine1")<>"" Then
				Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_BOMComapreOperations", ObjBOMCompareDialog, "BOMLine1", dicBOMCompareInfo("BOMLine1"))
			End If
			'Setting BOM Line 1 Revision Rule
			If dicBOMCompareInfo("BOMLine1RevisionRule")<>"" Then
				Call Fn_Web_UI_WebEdit_SetExt("Fn_SISW_Web_BOMComapreOperations", "Set",ObjBOMCompareDialog, "BOMLine1RevisionRule", dicBOMCompareInfo("BOMLine1RevisionRule"))
			End If
			'Setting BOM Line 2
			If dicBOMCompareInfo("BOMLine2")<>"" Then
				Call Fn_Web_UI_WebEdit_Set("Fn_SISW_Web_BOMComapreOperations", ObjBOMCompareDialog, "BOMLine2", dicBOMCompareInfo("BOMLine2"))
			End If
			'Setting BOM Line 2 Revision Rule
			If dicBOMCompareInfo("BOMLine2RevisionRule")<>"" Then
				Call Fn_Web_UI_WebEdit_SetExt("Fn_SISW_Web_BOMComapreOperations", "Set",ObjBOMCompareDialog, "BOMLine2RevisionRule", dicBOMCompareInfo("BOMLine2RevisionRule"))
			End If
			'Setting BOM Compare Mode
			If dicBOMCompareInfo("BOMCompareMode")<>"" Then
				Call Fn_Web_UI_WebEdit_SetExt("Fn_SISW_Web_BOMComapreOperations", "Set",ObjBOMCompareDialog, "BOMCompareMode", dicBOMCompareInfo("BOMCompareMode"))
			End If
			If StrButton<>"" Then
				Fn_SISW_Web_BOMComapreOperations=Fn_Web_UI_Button_ClickExt("Fn_SISW_Web_BOMComapreOperations", "Click",Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), StrButton)
			Else
				Fn_SISW_Web_BOMComapreOperations=True
			End If
			Call Fn_Web_ReadyStatusSync(1)
			Set wshShell=Nothing
   End Select
   'Releasing object of BOM Compare dialog
    Set ObjBOMCompareDialog=Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_SISIW_Web_BOMCompareResultTableOperation()
'
'Description			 :	Function Used to perform operations of BOM Compare Result tables

'Parameters			   :    StrAction: Valid Action Name
'								  		 StrBOMTable: BOM Line table 
'								  		 StrName: BOMLine Name path/ Node path
'								  		 StrColumn: Column name
'								  		 StrValue: value
'								  		 bCloseFlag: Page close value
'
'Return Value		   : 	True or False

'Pre-requisite			:	 BOM Compare Result tables should appear

'Examples				:  	bReturn=Fn_SISIW_Web_BOMCompareResultTableOperation("GetImage","BOMLine1","000024/A;1-TopItem1 (View):000025/A;1-Child1:000028/A;1-Item1","Name","2","True")
'										bReturn=Fn_SISIW_Web_BOMCompareResultTableOperation("VerifyCellBackgroundColour","BOMLine1","000024/A;1-TopItem1 (View):000025/A;1-Child1:000028/A;1-Item1","Name","Red","")
'History					 :			
'				Developer Name						Date					Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'				Sandeep Navghane				15-July-2013			1.0																			  		Gaurav S
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_SISIW_Web_BOMCompareResultTableOperation(StrAction,StrBOMTable,StrName,StrColumn,StrValue,bCloseFlag)
	GBL_FAILED_FUNCTION_NAME="Fn_SISIW_Web_BOMCompareResultTableOperation"
	Dim ObjBOMLineTable
	Dim iRowCnt,iColPos,objImg,strImageName,bFlag,objTDs,objTRs,Style,bgCol,Str,sColour
	Fn_SISIW_Web_BOMCompareResultTableOperation=False
	If not Browser("TeamcenterWeb-BOMCompare").Page("BOMCompare").Exist(5) Then
		Exit function
	End If
   Select Case StrBOMTable
	 	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "BOMLine1"
			Set ObjBOMLineTable=Browser("TeamcenterWeb-BOMCompare").Page("BOMCompare").WebTable("BOMLine1")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "BOMLine2"
			Set ObjBOMLineTable=Browser("TeamcenterWeb-BOMCompare").Page("BOMCompare").WebTable("BOMLine2")
   End Select

   Select Case StrAction
		' - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  -
			Case "GetImage"
					If StrValue = "" then 
						StrValue = 1
					Else 
						StrValue = cInt("" & StrValue)
					End If
					iRowCnt = Fn_WebUI_TableRowIndex(ObjBOMLineTable, StrName, "Name")
					iColPos = Fn_WebUI_TableColumnIndex(ObjBOMLineTable, "Name")
					If iRowCnt <> -1 and iColPos <> -1 Then
							Set objImg = ObjBOMLineTable.ChildItem(iRowCnt, iColPos, "Image",(StrValue - 1))
							If TypeName(objImg) <> "Nothing" Then
									strImageName=Split(objImg.GetROProperty("file name"),".")
									Fn_SISIW_Web_BOMCompareResultTableOperation=strImageName(0)
									bFlag = True
							Else
									Fn_SISIW_Web_BOMCompareResultTableOperation=False
							End If						
							Set objImg = Nothing
					End If
					If bFlag = True Then
							Fn_SISIW_Web_BOMCompareResultTableOperation=strImageName(0)
					Else
							Fn_SISIW_Web_BOMCompareResultTableOperation = False
					End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	 	Case "VerifyCellBackgroundColour"
					iRowCnt = Fn_WebUI_TableRowIndex(ObjBOMLineTable, StrName, "Name")
                    iColPos = Fn_WebUI_TableColumnIndex(ObjBOMLineTable, "Name")
					
					bFlag = False
					If iRowCnt <> -1 Then
						Set objTDs = ObjBOMLineTable.ChildItem(iRowCnt, iColPos,"WebElement","0")
						Set objTRs = objTDs.object.parentNode
						If InStr(1,Environment.Value("WebBrowserName"),"IE")>0 Then
							do while lcase(trim(objTRs.NodeName)) <> "tr"
									Set objTRs = objTRs.parentNode
							loop
							Style = objTRs.style.cssText
							Style = lcase(trim(Style))
						ElseIf InStr(1,Environment.Value("WebBrowserName"),"FF")>0 Then
							Style = objTRs.getAttribute("style")
						End If
						If instr(Style,"background-color:") > 0 OR objTRs.getAttribute("bgcolor") <> "" Then
							If objTRs.getAttribute("bgcolor") <> "" then
								' not yet implemented
							Else
								bgCol = instr(Style,"background-color:")
								If inStr(bgCol,Style,";") > 0 Then
									Str = trim(mid (Style, bgCol +  len("background-color:"),  instr(bgCol, Style,";") - bgCol -  len("background-color:")))
								Else
									Str = trim(mid (Style, bgCol +  len("background-color:"),  len(style) - bgCol +  len("background-color:")))
								End If
							End If
						End If
						Select Case lCase(Str)
							Case "red"
								sColour = "red"
						End Select
						If lCase(StrValue) = sColour then bFlag = True
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_SISIW_Web_BOMCompareResultTableOperation : Failed to find Node ["+CStr(StrName)+"] . ")
					End If
					' Write the Log of Success or Failure
					If bFlag = True Then
						Fn_SISIW_Web_BOMCompareResultTableOperation = True
					Else
						Fn_SISIW_Web_BOMCompareResultTableOperation = False
					End If					
   End Select
   If bCloseFlag<>"" Then
	   If lcase(bCloseFlag)="true" or lcase(bCloseFlag)="yes" Then
			Browser("TeamcenterWeb-BOMCompare").Close
			Call Fn_Web_ReadyStatusSync(1)
	   End If
   End If
   Set ObjBOMLineTable=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_Web_CommonModifiableProperties
'@@
'@@    Description				 :	Function Used to Modify properties
'@@
'@@    Parameters			   :	1.sAction: String value for Action to be performed
'@@												  2.sPropButton: Property name which will be modified
'@@												  3. sPropValue: New value of the property
'@@												  4.sButton:Button name
'@@												  5.strChangeID:New Change ID for check out
'@@												  6.strReason: Check Out Reason
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Multiple objects Should be Selected and Common Modifiable Property window should be open
'@@
'@@    Examples					:	Call Fn_Web_CommonModifiableProperties("EditProperty","Description","Test","SaveAndCheckIn","","")
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								Reviewer
'@@--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Avnessh Kumar									21-feb-2014						1.0																								Sunny Ruparel
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_CommonModifiableProperties(sAction,sPropButton,sPropValue,sButton,strChangeID,strReason)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_CommonModifiableProperties"
    Dim ObjTable,strWEBMenuPath,strMenu
	 Fn_Web_CommonModifiableProperties=False
	Set ObjTable=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CommonModifiableProperties")
	Set ObjCheckOut=Browser("TeamcenterWeb").Page("MyTeamCenter").WebTable("CheckOut")

   If  Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("CheckOutAndEdit").Exist(5) Then
		   Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton("CheckOutAndEdit").Click
		   wait 3
			If strChangeID<>"" Then
				'Setting Change ID
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_CheckOutObject", ObjCheckOut, "ChangeID", strChangeID)
			End If
			If strReason<>"" Then
				'Setting Reason For Object Check Out
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_CheckOutObject", ObjCheckOut, "Reason", strReason)
			End If
			'Clicking "OK" Button
			Call Fn_Web_UI_Button_Click("Fn_Web_CheckOutObject", Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel"), "OK")
			wait 2
   End If
		
	Select Case sAction
		Case "EditProperty"
			 ObjTable.WebButton("PropertyButton").SetTOProperty "name" , sPropButton
			 ObjTable.WebButton("PropertyButton").Click
			 ObjTable.SetTOProperty "text" ,sPropButton&".*Teamcenter.*"
			If sPropValue <>"" Then
				ObjTable.WebEdit("WebEdit").set sPropValue
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Edited Property Field [ "+sPropButton+" ]")
				Fn_Web_CommonModifiableProperties = True
			Else 
				Fn_Web_CommonModifiableProperties = False
			End If
			Select Case sButton
				Case  "SaveAndCheckIn" , "Save"
					ObjTable.WebButton(sButton).Click
				Case "CancelCheckOut", "Close"
                    Browser("TeamcenterWeb").Page("MyTeamCenter").WebElement("ButtunPanel").WebButton(sButton).Click
			End Select
		End Select
		 Set ObjTable=Nothing
		 Set ObjCheckOut=Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_Web_Login_WithoutInvoke_Browser
'@@
'@@    Description		:	Function Used to Log In Web Client without invoking Internet explorer, it sets username and password on already opened browser
'@@
'@@    Parameters		:	1. sAction		: Action Name
'@@						:	2. sUserName	: UserName
'@@						:	3. sPassword 	: Password
'@@
'@@    Return Value		: 	True Or False	(Browser should be opened with Login page)
'@@
'@@    Examples			:   bReturn = Fn_Web_Login_WithoutInvoke_Browser("Login", "AutoTest1", "AutoTest1")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.	Changes Done										Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vivek Ahirrao		25-Oct-2016		1.0		  	Created												[TC1017-2016101100-25_10_2016-VivekA-Maintenance]
'@@			
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Public Function Fn_Web_Login_WithoutInvoke_Browser(sAction, sUserName, sPassword)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_Login_WithoutInvoke_Browser"
	Dim objParent, bFlag
	Dim objErrPage
	
	Fn_Web_Login_WithoutInvoke_Browser = False
	'Set objErrPage = Fn_SISW_Web_GetObject("ErrorPage")
	
'	If Browser("TeamcenterLogin").Page("Login").Exist(5) Then
'		Set objParent = Browser("TeamcenterLogin").Page("Login").WebTable("Login")
'	ElseIf Browser("Browser").Page("ErrorPage").Exist(2) Then
'		Set objParent = Browser("Browser").Page("ErrorPage")
'	End If
	If Browser("TeamcenterLogin").Page("Login").Exist(5) Then
		Set objParent = Fn_SISW_Web_GetObject("WebLogin")
	ElseIf Fn_Web_UI_ObjectExist("Fn_Web_Login_WithoutInvoke_Browser", Browser("Browser").Page("ErrorPage")) Then
		Set objParent = Browser("Browser").Page("ErrorPage")
	End If
		
	bFlag = False
	Select Case sAction
		Case "Login", "Fn_Web_Login"
			'If objParent.Exist(30) Then
			If objParent.Exist(15) Then
				'Wait WEB_MICRO_TIMEOUT
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_Login_WithoutInvoke_Browser", objParent, "Username", sUserName)
				Wait WEB_MICRO_TIMEOUT
				Call Fn_Web_UI_WebEdit_Set("Fn_Web_Login_WithoutInvoke_Browser", objParent, "Password", sPassword)
				Wait WEB_MICRO_TIMEOUT
				Call Fn_Web_UI_Button_Click("Fn_Web_Login_WithoutInvoke_Browser", objParent, "Login")
				Call Fn_Web_ReadyStatusSync(2)
				Wait WEB_MICROLESS_TIMEOUT
				bFlag = True
			End If
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : [ Fn_Web_Login_WithoutInvoke_Browser ] Failed to Login to Web Client")
				Fn_Web_Login_WithoutInvoke_Browser = False
			Else
				Fn_Web_Login_WithoutInvoke_Browser = True
			End If
	End Select
	Set objParent = Nothing
End Function

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name	:	Fn_Web_MyTc_FolderCreate
'@@
'@@    Description		:	Function to create test case folder in My Teamcenter
'@@
'@@    Parameters		:	1. strName		: Folder Name
'@@					:	2. strDescription	: Folder Description
'@@
'@@    Return Value		: 	True Or False
'@@
'@@    Examples			:   bReturn = Fn_Web_MyTc_FolderCreate("CreateItem","foldercreate")
'@@    							
'@@	   History			:	
'@@			Developer Name		Date	  		Rev. No.			Changes Done			Reviewer
'@@-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@			Vrushali S		03-Jan-2017				1.0		  	Created					Prasad K
'@@			
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_Web_MyTc_FolderCreate(strName,strDescription)
	GBL_FAILED_FUNCTION_NAME="Fn_Web_MyTc_FolderCreate"
	Fn_Web_MyTc_FolderCreate=False
	
	Dim strMenu
	Dim sTestFolder
	Dim bFlag
	
	'Check Existence of "Home:AutomatedTests" Folder, if not exist then create
	bFlag = Fn_Web_NavTreeOperation("Exist","Home:AutomatedTests")
	If bFlag = False Then
		'Select Parent folder
		bFlag = Fn_Web_NavTreeOperation("Select","Home")
		If bFlag = False Then
			Call Fn_UpdateLogFiles(Cstr(now) + " - Action - FAIL | Failed to Select [ Home ].", "FAIL : Fail to Select [ Home ].")
			Exit Function
		End If
		Call Fn_Web_ReadyStatusSync(1)
		'Create folder
		bFlag = Fn_Web_FolderCreate("AutomatedTests","Automation Artifacts")
		If bFlag = False Then
			Call Fn_UpdateLogFiles(Cstr(now) + " - Action - FAIL | Failed to Create [ AutomatedTests ].", "FAIL : Fail to Create [ AutomatedTests ].")
			Exit Function
		End If
		Call Fn_Web_ReadyStatusSync(1)
	End If
	
	'Expand Parent Folder
	bFlag = Fn_Web_NavTreeOperation("Expand","Home:AutomatedTests")
	If bFlag = False Then
		Call Fn_UpdateLogFiles(Cstr(now) + " - Action - FAIL | Failed to Expand & Select Folder [ Home:AutomatedTests ].", "FAIL : Fail to Expand [ Home:AutomatedTests ].")
		Exit Function
	End If
	Call Fn_Web_ReadyStatusSync(1)
	
	'Select Parent Folder
	bFlag = Fn_Web_NavTreeOperation("Select","Home:AutomatedTests")
	If bFlag = False Then
		Call Fn_UpdateLogFiles(Cstr(now) + " - Action - FAIL | Failed to Select Folder [ Home:AutomatedTests ].", "FAIL : Fail to Select [ Home:AutomatedTests ].")
		Exit Function
	End If
	Call Fn_Web_ReadyStatusSync(1)
	
	'Create Testcase Folder under Parent folder
	sTestFolder = Environment.Value("TestName")
	sTestFolder=Replace(sTestFolder,".","")
	sTestFolder=Replace(sTestFolder," ","")
	sTestFolder=Replace(sTestFolder,"&","")
	sTestFolder=Replace(sTestFolder,"-","")
	
	If Len(sTestFolder)>25 Then
		sTestFolder = Mid(sTestFolder,1,25) & "_" & CStr(Fn_Setup_RandNoGenerate(5))
	Else
		sTestFolder = sTestFolder & "_" & CStr(Fn_Setup_RandNoGenerate(5))
	End If

'	If Not ObjFolder.Exist(7) Then
'		strWEBMenuPath=Fn_LogUtil_GetXMLPath("Web_Menu")
'		strMenu=Fn_GetXMLNodeValue(strWEBMenuPath, "NewFolder")
'		Call Fn_Web_MenuOperation("Select",strMenu)
'	End If

	bFlag = Fn_Web_FolderCreate(sTestFolder,sTestFolder & "-Test Case Folder")
	If bFlag = False Then
		Call Fn_UpdateLogFiles(Cstr(now) + " - Action - FAIL | Failed to Create folder [ "&sTestFolder&" ] under [ Home:AutomatedTests ].", "FAIL : Fail to Create [ "&sTestFolder&" ] under [ Home:AutomatedTests ].")
		Exit Function
	End If
	Call Fn_Web_ReadyStatusSync(1)
	
	'Select Parent Folder
	bFlag = Fn_Web_NavTreeOperation("Select","Home:AutomatedTests:" & sTestFolder)
	If bFlag = False Then
		Call Fn_UpdateLogFiles(Cstr(now) + " - Action - FAIL | Failed to Select Folder [ "&sTestFolder&" ].", "FAIL : Fail to Select [ "&sTestFolder&" ].")
		Exit Function
	End If
	Call Fn_Web_ReadyStatusSync(1)

	Fn_Web_MyTc_FolderCreate = sTestFolder
End Function
