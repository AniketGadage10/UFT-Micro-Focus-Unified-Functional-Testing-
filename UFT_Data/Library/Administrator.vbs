'------------------------'Global variables for Teamcenter Perspective Names--------------------------------------------------------------
Public GBL_PERSPECTIVE_ORGANIZATION
GBL_PERSPECTIVE_ORGANIZATION="Organization"
'------------------------'Global variables for Teamcenter Perspective Names--------------------------------------------------------------
'##############################################FUNCTION LIST#####################################################
'0.  Fn_SISW_Org_GetObject()
'1. Fn_Org_DialogMsgVerify()
'2. Fn_WebBro_AddressBarOperations()
'3. Fn_WebBro_ToolbarOperations()  
'4. Fn_Org_CategoryTreeOperations()
'5. Fn_Org_OrganizationTreeOperations()
'6. Fn_Org_GroupOperations()
'7. Fn_Org_RoleOperations()
'8. Fn_Org_SubGroupOperations()
'9. Fn_SetupWizard()
'10. Fn_Org_UserOperations()
'11. Fn_SetupWizard_MsgVerify()
'12. Fn_Org_PersonOperations()
'13. Fn_Org_VolumeOperations()
'14. Fn_Org_GroupMemberSettings()
'15. Fn_Reg_EditorOperation()
'16. Fn_RegEdit_ToolbatButtonClick()
'17. Fn_CommSup_MenuOperation()
'18. Fn_CommSup_SelectApp()
'19. Fn_CommSup_SelectOrg()
'20  Fn_NullPointerExceptionHandler()
'21. Fn_Auth_ShowHideApplication()
'22. Fn_Auth_OrganizationTreeOpration()
'23. Fn_Project_CheckButtonEnable()
'24. Fn_Project_TabOperation()
'25. Fn_Project_TabSet()
'26. Fn_Project_TreeOpeartion()
'27. Fn_Auth_ImportExportRule()
'28. Fn_CreateSetupWizardFile()
'29. Fn_Auth_ShowHideApplExt()
'30. Fn_Auth_UnsaveMsgVerify()  - Eliminated. Replaced by GeneralFunctions.vbs:: Fn_SISW_ErrorVerify By Sushma Pagare [7-Jun-13]
'31. Fn_Org_UserRate()
'32.Fn_Org_SiteOperations()
'33.Fn_Org_DisciplinesOperations()
'34.Fn_Org_CalendarOperations()
'35.Fn_Org_AddDisciplinesOperations()

'*********************************************************	Function List		***********************************************************************
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_Org_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_Org_GetObject("Remove")

'History:
'	Developer Name			Date						Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Snehal Salunkhe		     21-June-2012				1.0				Sandeep N.				
'-----------------------------------------------------------------------------------------------------------------------------------
'	Ashwini Kumar		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Org_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Administrator.xml"
	Set Fn_SISW_Org_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function
'#################################################################################################################
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_Org_DialogMsgVerify(sErrMsg,sDialogTitle)
'###
'###    DESCRIPTION     :   This function used to verify the Dialog messages
'###
'###    PARAMETERS      :   sErrMsg,sDialogTitle
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Harshal      		01/06/2010	  1.0
'###
'###    REVIWED BY      :   Harshal		   		01/06/2010	  1.0          
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Fn_Org_DialogMsgVerify("not an authorized user", "non-DBA Access")
'################################################################################################################
Function Fn_Org_DialogMsgVerify(sErrMsg,sDialogTitle)

	Dim dicErrorInfo
	Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	With dicErrorInfo 
	 .Add "Title", sDialogTitle
	 .Add "Message", sErrMsg
	 .Add "Button", "OK"
	End with
	Fn_Org_DialogMsgVerify = Fn_SISW_ErrorVerify(dicErrorInfo)

End Function
'*********************************************************		Function to Handle address bar operations		***********************************************************************
'Function Name		:				Fn_WebBro_AddressBarOperations() 

'Description			 :		 		 NavigateURL/VerifyURL

'Parameters			   :	 			1.sAction: Navigate/Verify
'													2.sURL:

'Return Value		   : 				True/False 

'Pre-requisite			:		 		WebBrowser Prespective is Open

'Examples				:				Call  Fn_WebBro_AddressBarOperations("Navigate","http://www.plm.automation.siemens.com/en_us/")
'													Call Fn_WebBro_AddressBarOperations("Verify","http://www.plm.automation.siemens.com/en_us/")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'															Pranav													   01/06/2010			          1.0										Created									Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''*********************************************************'*********************************************************'**********************************************************************
Public Function Fn_WebBro_AddressBarOperations(sAction,sURL)
	GBL_FAILED_FUNCTION_NAME="Fn_WebBro_AddressBarOperations"
		Dim sValue
		Set ObjDialog=Fn_UI_ObjectCreate("Fn_WebBro_AddressBarOperations",JavaWindow("Web Browser - Teamcenter").JavaEdit("AddressBar"))
		' Select the navigate Case 
		Select Case sAction				
			'Case to navigate the web browser
        
			Case "Navigate"
				' If Toolbar window exist
				If   Fn_UI_ObjectExist("Fn_WebBro_AddressBarOperations",ObjDialog) Then
					'To insert the URL into the address bar passed as a input parameter
					    Wait(2)
						Call  Fn_Edit_Box("Fn_WebBro_AddressBarOperations",JavaWindow("Web Browser - Teamcenter"),"AddressBar","")
						Call  Fn_Edit_Box("Fn_WebBro_AddressBarOperations",JavaWindow("Web Browser - Teamcenter"),"AddressBar",sURL)
						Wait(5)
						ObjDialog.Activate
						Call  Fn_KeyBoardOperation("SendKey", "{ENTER}")
                      ' Log the result
                		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: URL entered successfully ")
						Fn_WebBro_AddressBarOperations=True
				Else
					'If Toolbar Window does not exist  then Log the result
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Toolbar window does not exist ")
						Fn_WebBro_AddressBarOperations=False
                End If

             'Case to verify the URLfom the passed URL
		 	Case "Verify" 
						
						'If Address bar exists and Editable
						If   Fn_UI_ObjectExist("Fn_WebBro_AddressBarOperations",ObjDialog) Then
							'Get the URL from Address bar
							sValue=ObjDialog.GetROProperty ("value")
					While InStr(sValue, "Loading") <> 0
						wait 2
						sValue=ObjDialog.GetROProperty ("value")
					Wend
								If  instr(1,sValue,sURL)<>0 Then
								'Log the result
                                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Given URL and passed URL are same. ")
								Fn_WebBro_AddressBarOperations=True
							Else
								'Log the result
                                Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Given URL and passed URL are not same. ")
								Fn_WebBro_AddressBarOperations=False
							End If
						Else
								'If Address Bar does not exist, then Log the result
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Address Bar does not exist ")
								Fn_WebBro_AddressBarOperations=False
						End If

					Case "RMB_Paste"

								If   Fn_UI_ObjectExist("Fn_WebBro_AddressBarOperations",ObjDialog) Then
										JavaWindow("Web Browser - Teamcenter").JavaEdit("AddressBar").Set ""
										JavaWindow("Web Browser - Teamcenter").JavaEdit("AddressBar").Click 10, 10, "RIGHT"
										JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select "Paste"
										JavaWindow("Web Browser - Teamcenter").JavaEdit("AddressBar").Activate
										wait(1)
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: URL Pasted Successfully to Web-Browser")
										Fn_WebBro_AddressBarOperations=True
								Else
										'If Toolbar Window does not exist  then Log the result
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Web-Browser window does not exist ")
										Fn_WebBro_AddressBarOperations=False
								End If

            End Select
			Set ObjDialog=Nothing
End function
'*********************************************************		Function to Handle ToolBar Operations		***********************************************************************
'Function Name		:				Fn_WebBro_ToolbarOperations()  

'Description			 :		 		 Click the button in ToolBar

'Parameters			   :	 			1.sAction: Click
'													2.sButtonToolTip:

'Return Value		   : 				True/False 

'Pre-requisite			:		 		WebBrowser Prespective is Open

'Examples				:				'Call Fn_WebBro_ToolbarOperations("Click","Back to the previous page")
													'Call Fn_WebBro_ToolbarOperations("Click","Forward to the next page")
													'Call Fn_WebBro_ToolbarOperations("Click","Refresh the current page")
													'Call Fn_WebBro_ToolbarOperations("Click","Stop loading the current page")
													'Call Fn_WebBro_ToolbarOperations("Click","Print the current page")
													'Call Fn_WebBro_ToolbarOperations("Click","Home")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'															Pranav													   01/06/2010			          1.0										Created									Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''*********************************************************'*********************************************************'*****************************************************************

Public Function Fn_WebBro_ToolbarOperations(sAction,sButtonToolTip)
	GBL_FAILED_FUNCTION_NAME="Fn_WebBro_ToolbarOperations"
		Dim ToolBarObj
		Set ToolBarObj = Fn_UI_ObjectCreate("Fn_WebBro_ToolbarOperations",JavaWindow("Web Browser - Teamcenter").JavaToolbar("ToolBar"))
		' Select the navigate Case 
		Select Case sAction				
			'Case to Click on buttons on the Toolbar
        	Case "Click"
				' If Toolbar window exist
				If   Fn_UI_ObjectExist("Fn_WebBro_ToolbarOperations",ToolBarObj) Then
						'it will do the operation as per given parameter
                        ToolBarObj.Press sButtonToolTip
                        'Log the result
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"PASS: Successfully clicked on  ["+ sButtonToolTip +"] button")
						Fn_WebBro_ToolbarOperations=True
                Else
						'If Toolbar Window does not exist  then Log the result
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Toolbar window does not exist ")
						Fn_WebBro_ToolbarOperations=False
                End If

		End Select
		Set ToolBarObj=Nothing
End Function
''*********************************************************		Function to action perform on CategoryTree	***********************************************************************
'Function Name		:				Fn_Org_CategoryTreeOperations()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. Activate, Deactivate
'																	
'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Organization Prespective is Open.

'Examples				:				Case "Select" : Call Fn_Org_CategoryTreeOperations("Select","Users:Ketan Raje (x_raje)")
'													Case "Expand" : Call Fn_Org_CategoryTreeOperations("Expand","Users:Ketan Raje (x_raje)")
'													Case "Collapse" : Call Fn_Org_CategoryTreeOperations("Collapse","Users:Ketan Raje (x_raje)")
'													Case "Exist" : Call Fn_Org_CategoryTreeOperations("Exist","Groups:system")
'													Case "GetIndex" : Call Fn_Org_CategoryTreeOperations("GetIndex","Users:Ketan Raje (x_raje)")
'													Case "Activate","Deactivate" : 	Call Fn_Org_CategoryTreeOperations("Activate","Groups") / Call Fn_Org_CategoryTreeOperations("Deactivate","Groups")
'													Case "Getlist","GetChildCount" : Call Fn_Org_CategoryTreeOperations("Getlist","Groups") / Call Fn_Org_CategoryTreeOperations("GetChildCount","Groups")
'													Case "PopupMenuSelect" : Call Fn_Org_CategoryTreeOperations("PopupMenuSelect","Users:Ketan Raje (x_raje)~Rate")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			01/06/2010			              1.0										Created								Harshal
'													Prasanna B																		30/01/2012											 1.1							' Changes in the 'value' property of Category Tree from 9.1-0118 build 		
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_Org_CategoryTreeOperations(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_CategoryTreeOperations"
	Dim objJavaWindowCat, objJavaTreeCat, intNodeCount, intCount, sTreeItem, iLen, iCounter, iIndex, iTotal, sResult, arr
	Set objJavaWindowCat = Fn_UI_ObjectCreate( "Fn_Org_CategoryTreeOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
					sNodeName = "OrganizationListTree_ROOT:"+sNodeName        
                    Call Fn_JavaTree_Select("Fn_Org_CategoryTreeOperations", objJavaWindowCat, "CategoryTree",sNodeName)
					Fn_Org_CategoryTreeOperations = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
					sNodeName = "OrganizationListTree_ROOT:"+sNodeName        
                    Call Fn_UI_JavaTree_Expand("Fn_Org_CategoryTreeOperations",objJavaWindowCat,"CategoryTree",sNodeName)
					Fn_Org_CategoryTreeOperations = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
					sNodeName = "OrganizationListTree_ROOT:"+sNodeName        
                    Call Fn_UI_JavaTree_Collapse("Fn_Org_CategoryTreeOperations", objJavaWindowCat,"CategoryTree",sNodeName)
					Fn_Org_CategoryTreeOperations = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				'sNodeName = "Hi There:"+sNodeName
				sNodeName = "OrganizationListTree_ROOT:"+sNodeName        
				Set objJavaTreeCat = Fn_UI_ObjectCreate( "Fn_Org_CategoryTreeOperations", objJavaWindowCat.JavaTree("CategoryTree"))
					intNodeCount = objJavaTreeCat.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeCat.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_Org_CategoryTreeOperations = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_Org_CategoryTreeOperations = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				sNodeName = "OrganizationListTree_ROOT:"+sNodeName        
				bFlag = False
				For intCount=0 to objJavaWindowCat.JavaTree("CategoryTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowCat.JavaTree("CategoryTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_Org_CategoryTreeOperations = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_Org_CategoryTreeOperations = FALSE
					Exit Function
                 End If
		'----------------------------------------------------------------------- Activate/Deactivate-------------------------------------------------------------------------
		Case "Activate","Deactivate"
				sNodeName = "OrganizationListTree_ROOT:"+sNodeName        
				Call Fn_JavaTree_Node_Activate("Fn_Org_CategoryTreeOperations",objJavaWindowCat,"CategoryTree",sNodeName)	
				Fn_Org_CategoryTreeOperations = TRUE
		'----------------------------------------------------------------------- GetList of Child-nodes.-------------------------------------------------------------------------
		Case "Getlist","GetChildCount"
					sNodeName = "OrganizationListTree_ROOT:"+sNodeName        
					iCounter = 0
					iLen = Len(sNodeName)
					'Get the Index of parent node
					iIndex = Fn_JavaTree_NodeIndex("Fn_Org_CategoryTreeOperations",objJavaWindowCat,"CategoryTree",sNodeName)
					'Get the count of nodes in the tree.
					iTotal = Fn_UI_Object_GetROProperty("Fn_Org_CategoryTreeOperations",objJavaWindowCat.JavaTree("CategoryTree"), "items count")
					For iCount=(iIndex+1) to iTotal
						sReturn = objJavaWindowCat.JavaTree("CategoryTree").GetItem(iCount)
						iReturn = Len(sReturn)
						iReturn = iReturn - iLen
						If Instr(1, sReturn, sNodeName) > 0 Then
							If iCounter=0 Then
								sResult = mid(sReturn,(iLen+2),iReturn)+","
							Else
								sResult = sResult+mid(sReturn,(iLen+2),iReturn)+","
							End If
							'arr(iCounter) = mid(sReturn,(iLen+1),iReturn)
						Else
							Exit For
						End If
						iCounter = iCounter+1
					Next
					sResult = mid(sResult,1,(Len(sResult)-1)) 
					arr = Split(sResult,",")
					iCount = Ubound(arr)
					If sAction="Getlist" Then
						Fn_Org_CategoryTreeOperations = arr
					ElseIf sAction="GetChildCount" Then
						Fn_Org_CategoryTreeOperations = iCount+1
					End If		
		'----------------------------------------------------------------------- PopUp Menu Select.-------------------------------------------------------------------------		
		Case "PopupMenuSelect"
					Dim aMenuList,aNodeName,StrMenu,sName
					'Build the Popup menu to be selected
					aNodeName = split(sNodeName, "~",-1,1)
					aMenuList = split(aNodeName(1), ":",-1,1)
					intCount = Ubound(aMenuList)
					sNodeName = "OrganizationListTree_ROOT:"+aNodeName(0)        

					iIndex = Fn_JavaTree_NodeIndex("Fn_Org_CategoryTreeOperations",objJavaWindowCat,"CategoryTree",sNodeName)
'					iIndex = Cint(iIndex) + 1
					iIndex = Cint(iIndex)
					Call Fn_JavaTree_Select("Fn_Org_CategoryTreeOperations", objJavaWindowCat, "CategoryTree",sNodeName)
					Wait(5)
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_Org_CategoryTreeOperations", objJavaWindowCat, "CategoryTree",sNodeName)
					Wait(2)
					If not JavaWindow("Organization - Teamcenter").WinMenu("ContextMenu").Exist(3) Then
						sName = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaTree("CategoryTree").GetItem(iIndex)
						'Select node
						Call Fn_JavaTree_Select("Fn_Org_CategoryTreeOperations", objJavaWindowCat, "CategoryTree",sName)
				    	'Open context menu
						Call Fn_UI_JavaTree_OpenContextMenu("Fn_Org_CategoryTreeOperations", objJavaWindowCat, "CategoryTree",sName)
						Wait(2)
					End If
					If not JavaWindow("Organization - Teamcenter").WinMenu("ContextMenu").Exist(3) Then
						Call Fn_JavaTree_Select("Fn_Org_CategoryTreeOperations", objJavaWindowCat, "CategoryTree",sNodeName)
						'Open context menu
						Call Fn_UI_JavaTree_OpenContextMenu("Fn_Org_CategoryTreeOperations", objJavaWindowCat, "CategoryTree",sNodeName)
					End If 	
					'Select Menu action
					Select Case intCount
						Case "0"
							 StrMenu = JavaWindow("Organization - Teamcenter").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
							
						Case "1"
							StrMenu = JavaWindow("Organization - Teamcenter").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
						
						Case "2"
							StrMenu = JavaWindow("Organization - Teamcenter").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
							
						Case Else
							Fn_Org_CategoryTreeOperations = FALSE
							Exit Function
					End Select

					wait 3
				    If JavaWindow("Organization - Teamcenter").WinMenu("ContextMenu").GetItemProperty(StrMenu,"enabled") Then
						JavaWindow("Organization - Teamcenter").WinMenu("ContextMenu").Select StrMenu
					Else 
						Fn_Org_CategoryTreeOperations = FALSE	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "ContextMenu is disable.")
						Exit Function
					End If 

					Fn_Org_CategoryTreeOperations = TRUE				
		Case Else
						Fn_Org_CategoryTreeOperations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_CategoryTreeOperations function failed")
						Exit Function
End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_Org_CategoryTreeOperations")
	Set objJavaWindowCat = nothing
	Set objJavaTreeCat = nothing
End Function
''*********************************************************		Function to action perform on OrganisationTree	***********************************************************************
'Function Name		:				Fn_Org_OrganizationTreeOperations()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Organization Prespective is Open.

'Examples				:				Case "Select" : Call Fn_Org_OrganizationTreeOperations("Select","Organization:Engineering:Designer:Ketan Raje (x_raje)")
'													Case "Expand" : Call Fn_Org_OrganizationTreeOperations("Expand","Organization:GroupM:Checker")	
'													Case "Collapse" : Call Fn_Org_OrganizationTreeOperations("Collapse","Organization:GroupM:Checker")	
'													Case "Exist" : Call Fn_Org_OrganizationTreeOperations("Exist","Organization:GroupM:Checker")
'													Case "GetIndex" : Call Fn_Org_OrganizationTreeOperations("GetIndex","Organization:GroupM:Checker")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			01/06/2010			              1.0										Created				Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_Org_OrganizationTreeOperations(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_OrganizationTreeOperations"
	Dim objJavaWindowOrg, objJavaTreeOrg, intNodeCount, intCount, sTreeItem
	Dim iCounter, iLen,iIndex,iTotal,iCount,sReturn,iReturn,sResult,arr
	Set objJavaWindowOrg = Fn_UI_ObjectCreate( "Fn_Org_OrganizationTreeOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
                    Call Fn_JavaTree_Select("Fn_Org_OrganizationTreeOperations", objJavaWindowOrg, "OrganizationTree",sNodeName)
					Fn_Org_OrganizationTreeOperations = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_Org_OrganizationTreeOperations",objJavaWindowOrg,"OrganizationTree",sNodeName)
					Fn_Org_OrganizationTreeOperations = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_Org_OrganizationTreeOperations", objJavaWindowOrg,"OrganizationTree",sNodeName)
					Fn_Org_OrganizationTreeOperations = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeOrg = Fn_UI_ObjectCreate( "Fn_Org_OrganizationTreeOperations", objJavaWindowOrg.JavaTree("OrganizationTree"))
					intNodeCount = objJavaTreeOrg.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeOrg.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_Org_OrganizationTreeOperations = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_Org_OrganizationTreeOperations = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
	'----------------------------------------------------------------------- Get GetChildCount value of a particular node-------------------------------------------------------------------------		
		Case "Getlist","GetChildCount","GetChildCountExt"				'Added By Jotiba T - New Dev
'					sNodeName = sNodeName        
					iCounter = 0
					iLen = Len(sNodeName)
					'Get the Index of parent node
					iIndex = Fn_JavaTree_NodeIndex("Fn_Org_OrganizationTreeOperations",objJavaWindowOrg,"OrganizationTree",sNodeName)
					'Get the count of nodes in the tree.
					iTotal = Fn_UI_Object_GetROProperty("Fn_Org_OrganizationTreeOperations",objJavaWindowOrg.JavaTree("OrganizationTree"), "items count")
					For iCount=(iIndex+1) to iTotal
						sReturn = objJavaWindowOrg.JavaTree("OrganizationTree").GetItem(iCount)
						iReturn = Len(sReturn)
						iReturn = iReturn - iLen
						If Instr(1, sReturn, sNodeName) > 0 Then
							If iCounter=0 Then
								sResult = mid(sReturn,(iLen+2),iReturn)+","
							Else
								sResult = sResult+mid(sReturn,(iLen+2),iReturn)+","
							End If
							'arr(iCounter) = mid(sReturn,(iLen+1),iReturn)
						Else
							Exit For
						End If
						iCounter = iCounter+1
					Next
					
					If sAction="GetChildCountExt" Then ' Check child is empty
						If sResult=Empty Then
							Fn_Org_OrganizationTreeOperations=sResult
							Exit Function
						End If
					End If
					
					sResult = mid(sResult,1,(Len(sResult)-1)) 
					arr = Split(sResult,",")
					iCount = Ubound(arr)
					If sAction="Getlist" Then
						Fn_Org_OrganizationTreeOperations = arr
					ElseIf sAction="GetChildCount" Then
						Fn_Org_OrganizationTreeOperations = iCount+1
					End If		
					
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowOrg.JavaTree("OrganizationTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowOrg.JavaTree("OrganizationTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_Org_OrganizationTreeOperations = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_Org_OrganizationTreeOperations = FALSE
				End If

		Case Else
						Fn_Org_OrganizationTreeOperations = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_OrganizationTreeOperations function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_Org_OrganizationTreeOperations")
	Set objJavaWindowOrg = nothing
	Set objJavaTreeOrg = nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Org_GroupOperations()
'###
'###    DESCRIPTION        :   Create,Modify,Delete Groups
'###
'###    PARAMETERS      :   1. sAction: Create/Modify/Delete
'###											 2.	sGrpName:
'###											3.	sGrpDesc:
'###											4.	sGrpSecurity:
'###											5.	bDBAPrivilage:
'###											6.	sDefaultVol:
'###											7.	sDefaultLocalVol:
'###											8.	sRoles: ":"Seperated Values
'###											9.	sAttributes: ":" Seperated Values for blank Pass "None"
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           02/06/2010         1.0
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "Create" : Call Fn_Org_GroupOperations("Create" , "SQS" , "desccc" , "External" , "GroupM" , "ON" , "volume1" , "" , "DBA:Checker" , "SQS1:aaa1:None:None:None:11:None:None:A")
'###										 Case "Modify" : Call Fn_Org_GroupOperations("Modify" , "SQS" , "desc" , "Internal" , "GroupM" , "ON" , "volume1" , "" , "remove|DBA:remove|Checker:add|IP Admin:add|Designer" , "SQS1:aaa1:None:None:None:11:None:None:A")	
'###										 Case "Delete" : Call Fn_Org_GroupOperations("Delete" , "" , "" , "" , "" , "" , "" , "" , "" , "")	
'###										 Case "Verify" : Call Fn_Org_GroupOperations("Verify" , "AutoGroup1782010174228" , "GroupDesc" , "" , "" , "1" , "volume1" , "" , "" , "None:None:None:None:None:0:1111:None:None")	
'#############################################################################################################
Public Function Fn_Org_GroupOperations(sAction , sGrpName , sGrpDesc , sGrpSecurity , sToParent , bDBAPrivilage , sDefaultVol , sDefaultLocalVol , sRoles , sAttributes)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_GroupOperations"
	Fn_Org_GroupOperations = false
	Dim objGroup, objSelectType, intNoOfObjects, iCounter, bReturn, aColname, iCount, iRowData, aAttributes, iDefcount, iSelcount, aCols, bFlag
	Set objGroup = Fn_UI_ObjectCreate("Fn_Org_GroupOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
		Select Case sAction
				Case "Create"
						If sGrpName<>"" Then
							'Set Name
							call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Name",sGrpName)
						End If
						If sGrpDesc<>"" Then
							'Set description
							call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Description",sGrpDesc)
						End If
						'Set Security
						If sGrpSecurity<>"" Then
							Call Fn_Button_Click( "Fn_Org_GroupOperations", objGroup, "Dropdown")	
							'wait(2)
							wait(1)
							Set objSelectType=description.Create()
							objSelectType("Class Name").value = "JavaTable"					
							objSelectType("path").value = ".*LOVTreeTable.*"
							objSelectType("path").RegularExpression = True
							Set  intNoOfObjects = objGroup.ChildObjects(objSelectType)
							if intNoOfObjects.count > 0 then
								bGblFuncRetVal = Fn_SISW_UI_JavaTable_Operations("Fn_Org_GroupOperations", "SelectRow", intNoOfObjects(0) , "", "GetValueAt.getDisplayableValue", 0, sGrpSecurity, "", "", "", "")
								If bGblFuncRetVal = false Then
									Exit function
								End If
						   End If
						End If
						If sToParent<>"" Then
							'Set parent
							objGroup.JavaCheckBox("ToParent_Groups").Click 1,1,"LEFT"
							'wait(2)
							wait(1)
							'Count number of items in a List and select required.
							bReturn = objGroup.JavaList("ToParent_Groups").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objGroup.JavaList("ToParent_Groups").GetItem(iCounter))) = Trim(lcase(sToParent)) Then
									objGroup.JavaList("ToParent_Groups").Activate sToParent
									Exit For
								End If
							Next
						End If
						If bDBAPrivilage<>"" Then
							'Set DBA privilege status
							Call Fn_CheckBox_Set("Fn_Org_GroupOperations" ,objGroup,"DBAPrivilege", bDBAPrivilage)
						End If
						If sDefaultVol<>"" Then
							'Set Default Volume
							objGroup.JavaCheckBox("DefaultVolume_Groups").Click 1,1,"LEFT"
							'Count number of items in a List and select required.
							'wait(2)
							wait(1)
							bReturn = objGroup.JavaList("DefaultVolume_Groups").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objGroup.JavaList("DefaultVolume_Groups").GetItem(iCounter))) = Trim(lcase(sDefaultVol)) Then
									objGroup.JavaList("DefaultVolume_Groups").Activate sDefaultVol
									Exit For
								End If
							Next
						End If
						If sDefaultLocalVol<>"" Then
							'Set Default Local Volume
							objGroup.JavaCheckBox("DefaultLocalVolume_Groups").Click 1,1,"LEFT"
							'wait(2)
							wait(1)
							'Count number of items in a List and select required.
							bReturn = objGroup.JavaList("DefaultLocalVolume_Groups").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objGroup.JavaList("DefaultLocalVolume_Groups").GetItem(iCounter))) = Trim(lcase(sDefaultLocalVol)) Then
									objGroup.JavaList("DefaultLocalVolume_Groups").Activate sDefaultLocalVol
									Exit For
								End If
							Next
						End If						
							'Select Roles
							'Set Tab to Roles
							Call Fn_UI_JavaTab_Select("Fn_Org_GroupOperations",objGroup,"RolesAttributesTab", "Roles:")
						If sRoles<>"" Then
								bReturn = objGroup.JavaList("DefinedRoles").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sRoles, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objGroup.JavaList("DefinedRoles").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objGroup.JavaList("DefinedRoles").Select aColname(iRowData)
											'Click on Add column Button
											Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Add")
											Exit For 
										End If
									Next
								Next
						End If
						'Set Tab to ADA/ITAR Attributes
						Call Fn_UI_JavaTab_Select("Fn_Org_GroupOperations",objGroup,"RolesAttributesTab", "ADA/ITAR Attributes")
						If sAttributes<>"" Then
								'Select Attributes
								aAttributes = split(sAttributes, ":",-1,1)
								'Set Org Name
								If Trim(lcase(aAttributes(0))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Name_Group",aAttributes(0))
								End If
								'Set Org Legal Name
								If Trim(lcase(aAttributes(1))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Legal Name_Group",aAttributes(1))
								End If
								'Set Org Alt Name
								If Trim(lcase(aAttributes(2))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Alternate Name_Group",aAttributes(2))
								End If
								'Set Org Address
								If Trim(lcase(aAttributes(3))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Address_Group",aAttributes(3))
								End If
								'Set Org URL
								If Trim(lcase(aAttributes(4))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization URL_Group",aAttributes(4))
								End If
								'Set Operational status 'It is numeric field
								If Trim(lcase(aAttributes(5))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Operational Status_Group",aAttributes(5))
								End If
								'Set Org ID
								If Trim(lcase(aAttributes(6))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization ID_Group",aAttributes(6))
								End If
								'Set Org Type
								If Trim(lcase(aAttributes(7))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Type_Group",aAttributes(7))
								End If
								'Set Nationality
								If Trim(lcase(aAttributes(8))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Nationality_Group",aAttributes(8))
								End If
						End If
						'Click on create button.
						Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Create")

				Case "Delete"
						'Click on create button.
						Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Delete")
						'Click on yes button of delete dialog						
'						Call Fn_Button_Click("Fn_Org_GroupOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("DeleteConfirmation"), "Yes")
'commented code at line 697 & adde new code to suit hierarchy changes. Can be reverted to original code if required
'Shreyas 28-08-2012
						'If Window("Organization - Teamcenter_2").JavaApplet("JApplet").JavaDialog("Delete Confirmation").Exist Then
							'	Window("Organization - Teamcenter_2").JavaApplet("JApplet").JavaDialog("Delete Confirmation").JavaButton("Yes").Click
						'Else
						'	JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
					'	end if
						'If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("DeleteConfirmation").Exist Then
						If Fn_SISW_UI_Object_Operations("Fn_Org_GroupOperations","Exist",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("DeleteConfirmation"),SISW_MINLESS_TIMEOUT) Then
								JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
						end if


			    Case "Modify"
						If sGrpName<>"" Then
							'Set Name
							call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Name",sGrpName)
						End If
						If sGrpDesc<>"" Then
							'Set description
							call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Description",sGrpDesc)
						End If
						'Set Security
						If sGrpSecurity<>"" Then
							Call Fn_Button_Click( "Fn_Org_GroupOperations", objGroup, "Dropdown")							
								Set objSelectType=description.Create()
								objSelectType("Class Name").value = "JavaStaticText"					
								Set  intNoOfObjects = objGroup.ChildObjects(objSelectType)
									For  iCounter = 0 to intNoOfObjects.count-1
										   If  intNoOfObjects(iCounter).getROProperty("label") = sGrpSecurity Then
													intNoOfObjects(iCounter).Click 1,1
													Exit for
										   End If
									Next
						End If
						If sToParent<>"" Then
							'Set parent
							objGroup.JavaCheckBox("ToParent_Groups").Click 1,1,"LEFT"
							'wait(2)
							wait(1)
							'Count number of items in a List and select required.
							bReturn = objGroup.JavaList("ToParent_Groups").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objGroup.JavaList("ToParent_Groups").GetItem(iCounter))) = Trim(lcase(sToParent)) Then
									objGroup.JavaList("ToParent_Groups").Activate sToParent
									Exit For
								End If
							Next
						End If
						If bDBAPrivilage<>"" Then
							'Set DBA privilege status
							Call Fn_CheckBox_Set("Fn_Org_GroupOperations" ,objGroup,"DBAPrivilege", bDBAPrivilage)
						End If
						If sDefaultVol<>"" Then
							'Set Default Volume
							objGroup.JavaCheckBox("DefaultVolume_Groups").Click 1,1,"LEFT"
							'Count number of items in a List and select required.
							'wait(2)
							wait(1)
							bReturn = objGroup.JavaList("DefaultVolume_Groups").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objGroup.JavaList("DefaultVolume_Groups").GetItem(iCounter))) = Trim(lcase(sDefaultVol)) Then
									objGroup.JavaList("DefaultVolume_Groups").Activate sDefaultVol
									Exit For
								End If
							Next

						End If
						If sDefaultLocalVol<>"" Then
							'Set Default Local Volume
							objGroup.JavaCheckBox("DefaultLocalVolume_Groups").Click 1,1,"LEFT"
							'wait(2) 
							wait(1) 
							'Count number of items in a List and select required.
							bReturn = objGroup.JavaList("DefaultLocalVolume_Groups").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objGroup.JavaList("DefaultLocalVolume_Groups").GetItem(iCounter))) = Trim(lcase(sDefaultLocalVol)) Then
									objGroup.JavaList("DefaultLocalVolume_Groups").Activate sDefaultLocalVol
									Exit For
								End If
							Next
						End If
						'Select Roles
						'Set Tab to Roles
						Call Fn_UI_JavaTab_Select("Fn_Org_GroupOperations",objGroup,"RolesAttributesTab", "Roles:")
						If sRoles<>"" Then
									'Extract no of columns in Defined List
									iDefcount = objGroup.JavaList("DefinedRoles").GetROProperty("items count")
									'Extract no of columns in Selected List
									iSelcount = objGroup.JavaList("SelectedRoles").GetROProperty("items count")
									'Extract the index of row at which the object exist.
									aColname = split(sRoles, ":",-1,1)
									iCount = Ubound(aColname)
									For iRowData=0 to iCount
											aCols = split(aColname(iRowData), "|",-1,1)
											If Trim(lcase(aCols(0)))="add" Then
												'Adding column
												For iCounter=0 to iDefcount-1
														If Trim(lcase(objGroup.JavaList("DefinedRoles").GetItem(iCounter))) = Trim(lcase(aCols(1))) then
																objGroup.JavaList("DefinedRoles").Select aCols(1)
																'Click on Add column Button
																Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Add")
																Exit For 
														End If
												Next
											ElseIf Trim(lcase(aCols(0)))="remove" Then   
												'Removing column
												For iCounter=0 to iSelcount-1
														If Trim(lcase(objGroup.JavaList("SelectedRoles").GetItem(iCounter))) = Trim(lcase(aCols(1))) then
																objGroup.JavaList("SelectedRoles").Select aCols(1)
																'Click on Add column Button
																Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Remove")
																Exit For 
														End If
												Next
											End If			
									Next
						End If
						'Set Tab to ADA/ITAR Attributes
						Call Fn_UI_JavaTab_Select("Fn_Org_GroupOperations",objGroup,"RolesAttributesTab", "ADA/ITAR Attributes")
						If sAttributes<>"" Then
								'Select Attributes
								aAttributes = split(sAttributes, ":",-1,1)
								'Set Org Name
								If Trim(lcase(aAttributes(0))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Name_Group",aAttributes(0))
								End If
								'Set Org Legal Name
								If Trim(lcase(aAttributes(1))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Legal Name_Group",aAttributes(1))
								End If
								'Set Org Alt Name
								If Trim(lcase(aAttributes(2))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Alternate Name_Group",aAttributes(2))
								End If
								'Set Org Address
								If Trim(lcase(aAttributes(3))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Address_Group",aAttributes(3))
								End If
								'Set Org URL
								If Trim(lcase(aAttributes(4))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization URL_Group",aAttributes(4))
								End If
								'Set Operational status 'It is numeric field
								If Trim(lcase(aAttributes(5))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Operational Status_Group",aAttributes(5))
								End If
								'Set Org ID
								If Trim(lcase(aAttributes(6))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization ID_Group",aAttributes(6))
								End If
								'Set Org Type
								If Trim(lcase(aAttributes(7))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Type_Group",aAttributes(7))
								End If
								'Set Nationality
								If Trim(lcase(aAttributes(8))) <> Trim(lcase("None")) Then
									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Nationality_Group",aAttributes(8))
								End If
						End If
						'Click on modify button.
						Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Modify")	

				Case "OrgTreeModifyDetails"
'						If sGrpName<>"" Then
'							'Set Name
'							call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Name",sGrpName)
'						End If
'						If sGrpDesc<>"" Then
'							'Set description
'							call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Description",sGrpDesc)
'						End If
'						'Set Security
'						If sGrpSecurity<>"" Then
'							Call Fn_Button_Click( "Fn_Org_GroupOperations", objGroup, "Dropdown")							
'								Set objSelectType=description.Create()
'								objSelectType("Class Name").value = "JavaStaticText"					
'								Set  intNoOfObjects = objGroup.ChildObjects(objSelectType)
'									For  iCounter = 0 to intNoOfObjects.count-1
'										   If  intNoOfObjects(iCounter).getROProperty("label") = sGrpSecurity Then
'													intNoOfObjects(iCounter).Click 1,1
'													Exit for
'										   End If
'									Next
'						End If
'						If sToParent<>"" Then
'							'Set parent
'							objGroup.JavaCheckBox("ToParent_Groups").Click 1,1,"LEFT"
'							wait(2)
'							'Count number of items in a List and select required.
'							bReturn = objGroup.JavaList("ToParent_Groups").GetROProperty("items count")			    				
'							For iCounter=0 to bReturn -1
'								If Trim(lcase(objGroup.JavaList("ToParent_Groups").GetItem(iCounter))) = Trim(lcase(sToParent)) Then
'									objGroup.JavaList("ToParent_Groups").Activate sToParent
'									Exit For
'								End If
'							Next
'						End If
'						If bDBAPrivilage<>"" Then
'							'Set DBA privilege status
'							Call Fn_CheckBox_Set("Fn_Org_GroupOperations" ,objGroup,"DBAPrivilege", bDBAPrivilage)
'						End If
						If sDefaultVol<>"" Then
							'Modify 'DefaultVolume' control's index to identify it from the Org tree
							objGroup.JavaCheckBox("DefaultVolume_Groups").SetTOProperty "index", "0"
							'Set Default Volume
							objGroup.JavaCheckBox("DefaultVolume_Groups").Click 1,1,"LEFT"
							'Count number of items in a List and select required.
							wait(2)
							bReturn = objGroup.JavaList("DefaultVolume_Groups").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objGroup.JavaList("DefaultVolume_Groups").GetItem(iCounter))) = Trim(lcase(sDefaultVol)) Then
									objGroup.JavaList("DefaultVolume_Groups").Activate sDefaultVol
									Exit For
								End If
							Next
							'Reset index
							objGroup.JavaCheckBox("DefaultVolume_Groups").SetTOProperty "index", "1"
						End If
'						If sDefaultLocalVol<>"" Then
'							'Set Default Local Volume
'							objGroup.JavaCheckBox("DefaultLocalVolume_Groups").Click 1,1,"LEFT"
'							wait(2) 
'							'Count number of items in a List and select required.
'							bReturn = objGroup.JavaList("DefaultLocalVolume_Groups").GetROProperty("items count")			    				
'							For iCounter=0 to bReturn -1
'								If Trim(lcase(objGroup.JavaList("DefaultLocalVolume_Groups").GetItem(iCounter))) = Trim(lcase(sDefaultLocalVol)) Then
'									objGroup.JavaList("DefaultLocalVolume_Groups").Activate sDefaultLocalVol
'									Exit For
'								End If
'							Next
'						End If
'						'Select Roles
'						'Set Tab to Roles
'						Call Fn_UI_JavaTab_Select("Fn_Org_GroupOperations",objGroup,"RolesAttributesTab", "Roles:")
'						If sRoles<>"" Then
'									'Extract no of columns in Defined List
'									iDefcount = objGroup.JavaList("DefinedRoles").GetROProperty("items count")
'									'Extract no of columns in Selected List
'									iSelcount = objGroup.JavaList("SelectedRoles").GetROProperty("items count")
'									'Extract the index of row at which the object exist.
'									aColname = split(sRoles, ":",-1,1)
'									iCount = Ubound(aColname)
'									For iRowData=0 to iCount
'											aCols = split(aColname(iRowData), "|",-1,1)
'											If Trim(lcase(aCols(0)))="add" Then
'												'Adding column
'												For iCounter=0 to iDefcount-1
'														If Trim(lcase(objGroup.JavaList("DefinedRoles").GetItem(iCounter))) = Trim(lcase(aCols(1))) then
'																objGroup.JavaList("DefinedRoles").Select aCols(1)
'																'Click on Add column Button
'																Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Add")
'																Exit For 
'														End If
'												Next
'											ElseIf Trim(lcase(aCols(0)))="remove" Then   
'												'Removing column
'												For iCounter=0 to iSelcount-1
'														If Trim(lcase(objGroup.JavaList("SelectedRoles").GetItem(iCounter))) = Trim(lcase(aCols(1))) then
'																objGroup.JavaList("SelectedRoles").Select aCols(1)
'																'Click on Add column Button
'																Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Remove")
'																Exit For 
'														End If
'												Next
'											End If			
'									Next
'						End If
'						'Set Tab to ADA/ITAR Attributes
'						Call Fn_UI_JavaTab_Select("Fn_Org_GroupOperations",objGroup,"RolesAttributesTab", "ADA/ITAR Attributes")
'						If sAttributes<>"" Then
'								'Select Attributes
'								aAttributes = split(sAttributes, ":",-1,1)
'								'Set Org Name
'								If Trim(lcase(aAttributes(0))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Name_Group",aAttributes(0))
'								End If
'								'Set Org Legal Name
'								If Trim(lcase(aAttributes(1))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Legal Name_Group",aAttributes(1))
'								End If
'								'Set Org Alt Name
'								If Trim(lcase(aAttributes(2))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Alternate Name_Group",aAttributes(2))
'								End If
'								'Set Org Address
'								If Trim(lcase(aAttributes(3))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Address_Group",aAttributes(3))
'								End If
'								'Set Org URL
'								If Trim(lcase(aAttributes(4))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization URL_Group",aAttributes(4))
'								End If
'								'Set Operational status 'It is numeric field
'								If Trim(lcase(aAttributes(5))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Operational Status_Group",aAttributes(5))
'								End If
'								'Set Org ID
'								If Trim(lcase(aAttributes(6))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization ID_Group",aAttributes(6))
'								End If
'								'Set Org Type
'								If Trim(lcase(aAttributes(7))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Organization Type_Group",aAttributes(7))
'								End If
'								'Set Nationality
'								If Trim(lcase(aAttributes(8))) <> Trim(lcase("None")) Then
'									call Fn_Edit_Box("Fn_Org_GroupOperations",objGroup,"Nationality_Group",aAttributes(8))
'								End If
'						End If
						'Click on modify button.
						Call Fn_Button_Click("Fn_Org_GroupOperations", objGroup, "Modify")	

				Case "Verify"
						bFLag = True
						'Verify Name
						If sGrpName<>"" Then							
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Name"))) = Trim(Lcase(sGrpName)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupName value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupName value dose'nt matches with actual value")
								bFlag = False
							End If
						End If
						'Verify Description
						If sGrpDesc<>"" Then							
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Description"))) = Trim(Lcase(sGrpDesc)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupDescription value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupDescription value dose'nt matches with actual value")
								bFlag = False
							End If
						End If
						'Verify Security
						If sGrpSecurity<>"" Then							
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Security"))) = Trim(Lcase(sGrpSecurity)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupSecurity value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupSecurity value dose'nt matches with actual value")
								bFlag = False
							End If
						End If
						'Verify Parent
						If sToParent<>"" Then
							If Trim(Lcase(objGroup.JavaCheckBox("ToParent_Groups").GetROProperty("attached text"))) = Trim(Lcase(sToParent)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Parent value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Parent value dose'nt matches with actual value")
								bFlag = False
							End If
						End If
						'Verify DBA privilege status
						If bDBAPrivilage<>"" Then
							If objGroup.JavaCheckBox("DBAPrivilege").GetROProperty("value") = bDBAPrivilage Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "DBAPrivilege value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "DBAPrivilege value dose'nt matches with actual value")
								bFlag = False
							End If
						End If
 						'Verify Default Volume
						If sDefaultVol<>"" Then
							If Trim(Lcase(objGroup.JavaCheckBox("DefaultVolume_Groups").GetROProperty("attached text"))) = Trim(Lcase(sDefaultVol)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupVolume value matches with actual value")
							Else
								'To verify Default Volume infor from OrgTree
								objGroup.JavaCheckBox("DefaultVolume_Groups").SetToProperty "index", "0"
								If Trim(Lcase(objGroup.JavaCheckBox("DefaultVolume_Groups").GetROProperty("attached text"))) = Trim(Lcase(sDefaultVol)) Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupVolume value matches with actual value")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "GroupVolume value dose'nt matches with actual value")
									bFlag = False
								End If
								'Reset index
								objGroup.JavaCheckBox("DefaultVolume_Groups").SetToProperty "index", "1"
							End If
						End If
 						'Verify Default Local Volume
						If sDefaultLocalVol<>"" Then
							If Trim(Lcase(objGroup.JavaCheckBox("DefaultLocalVolume_Groups").GetROProperty("attached text"))) = Trim(Lcase(sDefaultLocalVol)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Group Default Volume value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Group Default Volume value dose'nt matches with actual value")
								bFlag = False
							End If
						End If
						'Verify Roles  													''Not yet coded
						'Verify ADA/ITAR Attributes
						If sAttributes<>"" Then
								Call Fn_UI_JavaTab_Select("Fn_Org_GroupOperations",objGroup,"RolesAttributesTab", "ADA/ITAR Attributes")
								'Select Attributes
								aAttributes = split(sAttributes, ":",-1,1)
								'Set Org Name
								If Trim(lcase(aAttributes(0))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Organization Name_Group"))) = Trim(Lcase(aAttributes(0))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Name value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Name value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
								'Set Org Legal Name
								If Trim(lcase(aAttributes(1))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Organization Legal Name_Group"))) = Trim(Lcase(aAttributes(1))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Legal Name value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Legal Name value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
								'Set Org Alt Name
								If Trim(lcase(aAttributes(2))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Organization Alternate Name_Group"))) = Trim(Lcase(aAttributes(2))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Alt Name Group value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Alt Name Group value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
								'Set Org Address
								If Trim(lcase(aAttributes(3))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Organization Address_Group"))) = Trim(Lcase(aAttributes(3))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Address value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Address value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
								'Set Org URL
								If Trim(lcase(aAttributes(4))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Organization URL_Group"))) = Trim(Lcase(aAttributes(4))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org URL value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org URL value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
								'Set Operational status 'It is numeric field
								If Trim(lcase(aAttributes(5))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Operational Status_Group"))) = Trim(Lcase(aAttributes(5))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Operational Status value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Operational Status value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
								'Set Org ID
								If Trim(lcase(aAttributes(6))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Organization ID_Group"))) = Trim(Lcase(aAttributes(6))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org ID value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org ID value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
								'Set Org Type
								If Trim(lcase(aAttributes(7))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Organization Type_Group"))) = Trim(Lcase(aAttributes(7))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Type value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Org Type value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
								'Set Nationality
								If Trim(lcase(aAttributes(8))) <> Trim(lcase("None")) Then
									If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_GroupOperations",objGroup,"Nationality_Group"))) = Trim(Lcase(aAttributes(8))) Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Nationality value matches with actual value")
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Nationality value dose'nt matches with actual value")
										bFlag = False
									End If
								End If
						End If
						If bFlag=True Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All values verified successfully")
							Fn_Org_GroupOperations = TRUE
							Set objGroup = nothing 
							Exit Function
						ElseIf bFlag = False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All values Does not match successfully")
							Fn_Org_GroupOperations = False
							Set objGroup = nothing 
							Exit Function
						End If						
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_GroupOperations function failed")
						Fn_Org_GroupOperations = FALSE
						Exit Function						
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Group [" + sGrpName + "] of function Fn_Org_GroupOperations")
	Fn_Org_GroupOperations = TRUE
	Set objGroup = nothing 
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Org_RoleOperations()
'###
'###	PREQUISITE			  :		1. DBA Pervilages needed.
'###												2. Organization Prespective is Open.
'###												3. Select the Group from Organization Tree
'###
'###    DESCRIPTION        :   Add/Remove Roles From Groups
'###
'###    PARAMETERS      :   1. sAction: Add/Remove
'###											 2.	sRoleName: ":" Seperated String(for AddExisting)
'###											3.	sRoleDesc
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           03/06/2010         1.0
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  Ketan Raje  				(Added Case : "AddAllExisting")
'###
'###    EXAMPLE          : 		Case "AddNew" : Call Fn_Org_RoleOperations("AddNew","Ketan2","desc")
'###										 Case "AddExisting " : Call Fn_Org_RoleOperations("AddExisting","designer:DBA","")
'###										 Case "AddAllExisting" : Call Fn_Org_RoleOperations("AddAllExisting","","")
'###										 Case "Remove" : Call Fn_Org_RoleOperations("Remove","","")
'###										 Case "Create" : Call Fn_Org_RoleOperations("Create","Ketan4","desccc")	
'###										 Case "Modify" :  Call Fn_Org_RoleOperations("Modify","Ketan4","desccc")	
'###										 Case "Delete" : Call Fn_Org_RoleOperations("Delete","","")								
'#############################################################################################################

Public Function Fn_Org_RoleOperations(sAction,sRoleName,sRoleDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_RoleOperations"
	Dim objRole, objRoleGrp, aColname, iCounter, bReturn, iCount, iRowData
	Set objRole = Fn_UI_ObjectCreate("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
		Select Case sAction
				Case "AddNew"
						'Click on AddRole button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRole, "AddRole")
						'Set obj for Role wizard
						
						'Swapnil : Changing the hierarchy: 0620 build : 26-JUNE-2012
						'Set objRoleGrp = Fn_UI_ObjectCreate("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Organization Role Wizard"))	
						Set objRoleGrp = Fn_UI_ObjectCreate("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Role Wizard"))
						'Select the Add new role radio button						
						Call Fn_UI_JavaRadioButton_SetON("Fn_Org_RoleOperations",objRoleGrp, "Add new role to the group")
						'Click on Next button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Next")
						'Set Role
						Call  Fn_Edit_Box("Fn_Org_RoleOperations",objRoleGrp,"Role",sRoleName)
						'Set Description
						Call  Fn_Edit_Box("Fn_Org_RoleOperations",objRoleGrp,"Description",sRoleDesc)
						'Click on finish button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Finish")
						'Click on yes button
						'If objRoleGrp.JavaButton("Yes").Exist(5) Then
						If Fn_SISW_UI_Object_Operations("Fn_Org_RoleOperations","Exist",objRoleGrp.JavaButton("Yes"),SISW_MINLESS_TIMEOUT) Then
							Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Yes")
						End If
						'Set property of the window
						'Call Fn_UI_Object_SetTOProperty("Fn_Org_RoleOperations",JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"),"text","Role(s) added")
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Role(s) added"
						'Click on OK button
						'Call Fn_Button_Click("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"), "OK")
						'If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist(5) Then
						If Fn_SISW_UI_Object_Operations("Fn_Org_RoleOperations","Exist",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog"),SISW_MICRO_TIMEOUT) Then
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
						End If
						'Click on close button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Close")

				Case "AddExisting"
						'Click on AddRole button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRole, "AddRole")
						'Set obj for Role wizard
						'Set objRoleGrp = Fn_UI_ObjectCreate("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Organization Role Wizard"))	
                         Set objRoleGrp = Fn_UI_ObjectCreate("Fn_Org_RoleOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Role Wizard"))						
						'Select the Add new role radio button						
						Call Fn_UI_JavaRadioButton_SetON("Fn_Org_RoleOperations",objRoleGrp, "Add existing role to the")
						'Click on Next button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Next")
						
						'[TC1122-2016011300-20_01_2016-VivekA-Maintenance] - Added as  per Design change ----------------------------
'						'Extract the index of row at which the object exist.
'						aColname = split(sRoleName, ":",-1,1)
'						iCount = Ubound(aColname)
'						For iRowData=0 to iCount
'								bReturn = Fn_List_Select("Fn_Org_RoleOperations", objRoleGrp, "ExistingRoles", aColname(iRowData))
'								If bReturn = false Then
'									Fn_Org_RoleOperations = false
'									Exit function
'								End If
'								'Click on Add column Button
'								Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Add")
'						Next						
						'get item count
						bReturn = objRoleGrp.JavaList("ExistingRoles").GetROProperty("items count")
						'Extract the index of row at which the object exist.
						aColname = split(sRoleName, ":",-1,1)
						iCount = Ubound(aColname)
						For iRowData=0 to iCount
							For iCounter=0 to bReturn-1
								If Trim(lcase(objRoleGrp.JavaList("ExistingRoles").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
									objRoleGrp.JavaList("ExistingRoles").Select aColname(iRowData)
									'Click on Add column Button
									Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Add")
									Exit For 
								End If
							Next
						Next
						'------------------------------------------------------------------------------------------------------------
						
						'Click on finish button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Finish")
						'If objRoleGrp.JavaButton("Yes").Exist Then
						If Fn_SISW_UI_Object_Operations("Fn_Org_RoleOperations","Exist",objRoleGrp.JavaButton("Yes"),SISW_MINLESS_TIMEOUT) Then
							'Click on yes button
							Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Yes")
							'Set property of the window
							Call Fn_UI_Object_SetTOProperty("Fn_Org_RoleOperations",JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"),"text","Role(s) added")
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Role(s) added"
							'Click on OK button
							Call Fn_Button_Click("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"), "OK")
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
						End If
						'Click on close button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Close")

				Case "AddAllExisting"
						'Click on AddRole button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRole, "AddRole")
						'Set obj for Role wizard
						Set objRoleGrp = Fn_UI_ObjectCreate("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Organization Role Wizard"))						
						'Select the Add new role radio button						
						Call Fn_UI_JavaRadioButton_SetON("Fn_Org_RoleOperations",objRoleGrp, "Add existing role to the")
						'Click on Next button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Next")
						'Set Role
						bReturn = objRoleGrp.JavaList("ExistingRoles").GetROProperty("items count")
						'Selecting All the items in the List.
						objRoleGrp.JavaList("ExistingRoles").SelectRange 0,(bReturn-1)
						'Click on Add column Button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Add")
						'Click on finish button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Finish")
						'Click on yes button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Yes")
						'Set property of the window
						'Call Fn_UI_Object_SetTOProperty("Fn_Org_RoleOperations",JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"),"text","Role(s) added")
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Role(s) added"
						'Click on OK button
						'Call Fn_Button_Click("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"), "OK")
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
						'Click on close button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRoleGrp, "Close")

				Case "Remove"
						'Click on remove button.
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRole, "Remove_Subgroup")
						JavaDialog("DeleteConfirmation").SetTOProperty "Title","Remove User Confirmation"
						'If JavaDialog("DeleteConfirmation").Exist(3) Then
						If Fn_SISW_UI_Object_Operations("Fn_Org_RoleOperations","Exist",JavaDialog("DeleteConfirmation"),SISW_MINLESS_TIMEOUT) Then
								Call Fn_Button_Click("Fn_Org_RoleOperations", JavaDialog("DeleteConfirmation"), "Yes")
					    End If
						
				Case "Create"
						'Set Role
						If sRoleName <> "" Then
							Call  Fn_Edit_Box("Fn_Org_RoleOperations",objRole,"Role",sRoleName)
						End If
						'Set Description
						If sRoleDesc <> "" Then
							Call  Fn_Edit_Box("Fn_Org_RoleOperations",objRole,"Description",sRoleDesc)
						End If
						'Click on create button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRole, "Create")
						
			    Case "Modify"
						'Set Role
						If sRoleName <> "" Then
							Call  Fn_Edit_Box("Fn_Org_RoleOperations",objRole,"Role",sRoleName)
						End If
						'Set Description
						If sRoleDesc <> "" Then
							Call  Fn_Edit_Box("Fn_Org_RoleOperations",objRole,"Description",sRoleDesc)
						End If
						'Click on modify button
						Call Fn_Button_Click("Fn_Org_RoleOperations", objRole, "Modify")

			    Case "Delete"
					  'Click on Delete button.
					  Call Fn_Button_Click("Fn_Org_RoleOperations", objRole, "Delete")
					  'Click on yes button to delete the site.
					  For iCount = 0 to 0
					   JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete Confirmation"		'Modified code to handle Msgbox in multiple hierarchy.
					   'If JavaDialog("DeleteConfirmation").Exist Then
					   If Fn_SISW_UI_Object_Operations("Fn_Org_RoleOperations","Exist",JavaDialog("DeleteConfirmation"),SISW_MINLESS_TIMEOUT) Then
							JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
							Exit For
					   End If
					   JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").SetTOProperty "title", "Delete Confirmation"
					   'If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").Exist Then
					   If Fn_SISW_UI_Object_Operations("Fn_Org_RoleOperations","Exist",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog"),SISW_MICRO_TIMEOUT) Then
							JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").JavaButton("Yes").Click
							Exit For
					   End If
					  Next
					  Call Fn_ReadyStatusSync(1)

				Case "Verify"
						iCount = 0
						iCounter = 0
						'Verify Role
						If sRoleName <> "" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_RoleOperations",objRole,"Role"))) = Trim(Lcase(sRoleName)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Description
						If sRoleDesc <> "" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_RoleOperations",objRole,"Description"))) = Trim(Lcase(sRoleDesc)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Return value
						If iCount <> iCounter Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to verify Role properties.")
							Fn_Org_RoleOperations = FALSE
							Exit Function												
						End If
				Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_RoleOperations function failed")
						Fn_Org_RoleOperations = FALSE
						Exit Function						
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Role [" + sRoleName + "] of function Fn_Org_RoleOperations")
	Fn_Org_RoleOperations = TRUE
	Set objRole = nothing 
	Set objRoleGrp = nothing 
End Function

'#########################################################################################################################################
'###    FUNCTION NAME   :    Fn_Org_SubGroupOperations()  
'###
'###    DESCRIPTION     :   Add/Remove Sub-Groups
'###
'###    PARAMETERS      :   1. sAction: New/Existing/Remove
'### 											 2. sSubGrpName:
'### 											3. sSubGrpDesc:
'### 											4. sSubGrpSecurity:
'### 											5.bDBAPrivilage:
'### 											6.sDefaultVol:
'### 											7.sDefaultLocalVol: 
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Swapna    		02-june-2010	  1.0
'###
'###    REVIWED BY      :   Harshal	   		02-June-2010	  1.0          
'###
'###    MODIFIED BY     :    Harshal/Ketan
'###
'###      MODIFIED BY     :  Swapna Ghatge    8-Dec-2011  
'###                                       :  Sanjeet Kumar       7-June-2012 
'###	
'###		DETAILS    :		Modified Cases "AddNew" & "AddExisting" : - on Build 1130  confirmation dailog not coming while adding new/existing subgroup in group.
 '###                                       Changed hierarchy for   'Organization Group Wizard'  from Window to JavaApplet.               
'###
'###    EXAMPLE         :  'Call Fn_Org_SubGroupOperations("AddNew","test1","testDesc","External","ON","volume1","")
											'Call Fn_Org_SubGroupOperations("AddExisting","TestGroup2","","","","","")
											'Call Fn_Org_SubGroupOperations("Remove","","","","","","")
											'Call Fn_Org_SubGroupOperations("SearchVerify","dba:TestGroup","","","","","")
											'Call Fn_Org_SubGroupOperations("AddAllExisting","","","","","","")
'############################################################################################################################################
Function Fn_Org_SubGroupOperations(sAction,sSubGrpName,sSubGrpDesc,sSubGrpSecurity,bDBAPrivilage,sDefaultVol,sDefaultLocalVol)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_SubGroupOperations"
		Dim objOrgGrp,objErrorDialog,objSelectType,objjApplet,objDialog, iCounter, intNoOfObjects, bReturn, objaddsubgroup, aColname, iCount
			Select Case sAction
					Case "AddExisting"

											Set objjApplet = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet")
											 Set objOrgGrp =JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Group Wizard")
                                    			'To Click on AddSub-Group Button
													Call Fn_Button_Click("Fn_Org_SubGroupOperations",objjApplet ,"AddSub-Group")
												'To Check The "Organization Group Wizard"  Dialog is present or not
											If  Fn_UI_ObjectExist("Fn_Org_SubGroupOperations",objOrgGrp)=True Then	
												 'To select the "Add existing group as" Radio Button
													Call Fn_UI_JavaRadioButton_SetON("Fn_Org_SubGroupOperations",objOrgGrp, "Add existing group as")
												'To Click on Next Button
													Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Next")
												'To select the Group from Available Groups
													Call Fn_List_Select("Fn_Org_SubGroupOperations",objOrgGrp,"Available Groups",sSubGrpName)
												'To Click on ">" Button
													Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,">")
												If  objOrgGrp.javaButton("Next").GetROProperty("enabled")=1 Then													
															'To Click on Next Button
																Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Next")
															'To Click on Yes Button
																Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Yes")		
															'To Click on OK Button
															JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Group(s) added"
															JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
												End If
													'To Click on Finish Button
												If objOrgGrp.javaButton("Finish").Exist Then
															Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Finish")
												End If
												'To Click on close Button
												Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Close")
												Fn_Org_SubGroupOperations = TRUE
											End If									

					Case "AddAllExisting"
											Set objjApplet = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet")
											 Set objOrgGrp =JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Group Wizard")
                                    			'To Click on AddSub-Group Button
													Call Fn_Button_Click("Fn_Org_SubGroupOperations",objjApplet ,"AddSub-Group")
												'To Check The "Organization Group Wizard"  Dialog is present or not
											If  Fn_UI_ObjectExist("Fn_Org_SubGroupOperations",objOrgGrp)=True Then	
												 'To select the "Add existing group as" Radio Button
													Call Fn_UI_JavaRadioButton_SetON("Fn_Org_SubGroupOperations",objOrgGrp, "Add existing group as")
												'To Click on Next Button
													Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Next")
												'Get Count of All items in the AvailableGroups List.
												bReturn = objOrgGrp.JavaList("Available Groups").GetROProperty("items count")
												'To select All Groups in the List.
													objOrgGrp.JavaList("Available Groups").SelectRange 0,(bReturn-1)
												'To Click on ">" Button
													Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,">")
												'To Click on Next Button
													Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Next")
												'To Click on Yes Button
													Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Yes")		
												'To Click on OK Button
												JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Group(s) added"
												JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
												Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Close")
												Fn_Org_SubGroupOperations = TRUE
											End If									
					Case "AddNew"
										Set objOrgGrp =JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Group Wizard")
										Set objjApplet = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet")
										'To Click on AddSub-Group Button
										Call Fn_Button_Click("Fn_Org_SubGroupOperations",objjApplet,"AddSub-Group")
									'To Check The "Organization Group Wizard"  Dialog is present or not
									If  Fn_UI_ObjectExist("Fn_Org_SubGroupOperations",objOrgGrp)=True Then
											Call Fn_UI_JavaRadioButton_SetON("Fn_Org_SubGroupOperations",objOrgGrp, "Add new group as sub-group")
											Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Next")	
											'To set Name fro group
											Call Fn_Edit_Box("Fn_Org_SubGroupOperations",objOrgGrp,"Name",sSubGrpName)	
											'To set Description
											Call Fn_Edit_Box("Fn_Org_SubGroupOperations",objOrgGrp,"Description",sSubGrpDesc)	
											'To select  security type
											If sSubGrpSecurity<>"" Then		
												Call Fn_Button_Click( "Fn_Org_SubGroupOperations", objOrgGrp, "Security")
												JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Group Wizard").JavaStaticText("Security").SetTOProperty "label",sSubGrpSecurity
												JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Group Wizard").JavaStaticText("Security").Click 1,1,"LEFT"
											End If
											'To check on DBA Privilege		
											Call Fn_CheckBox_Set("Fn_Org_SubGroupOperations", objOrgGrp,"DBA Privilege", bDBAPrivilage)
											If sDefaultVol<>"" Then
												'Set Default Volume
												objOrgGrp.JavaCheckBox("DefaultVolume").Click 1,1,"LEFT"
												'Count number of items in a List and select required.
												wait(2)

												'[commented By Shreyas & Added new code 12-07-2012]

'												bReturn = objOrgGrp.JavaList("Volumes").GetROProperty("items count")			    				
'													For iCounter=0 to bReturn -1
'														If Trim(lcase(objOrgGrp.JavaList("Volumes").GetItem(iCounter))) = Trim(lcase(sDefaultVol)) Then
'															objOrgGrp.JavaList("Volumes").Activate sDefaultVol
'															Exit For
'														End If																											
'													Next

												If objOrgGrp.JavaList("Volumes").Exist(0) Then
													bReturn = objOrgGrp.JavaList("Volumes").GetROProperty("items count")			    				
														For iCounter=0 to bReturn -1
															If Trim(lcase(objOrgGrp.JavaList("Volumes").GetItem(iCounter))) = Trim(lcase(sDefaultVol)) Then
																objOrgGrp.JavaList("Volumes").Activate sDefaultVol
																Exit For
															End If																											
														Next
												Else
													bReturn = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaList("DefaultVolume_Groups").GetROProperty("items count")			    				
														For iCounter=0 to bReturn -1
															If Trim(lcase(JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaList("DefaultVolume_Groups").GetItem(iCounter))) = Trim(lcase(sDefaultVol)) Then
																JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaList("DefaultVolume_Groups").Activate sDefaultVol
																Exit For
															End If																											
														Next
												End If
'											End If
											End If
											If sDefaultLocalVol<>"" Then
												'Set Default Local Volume
												objOrgGrp.JavaCheckBox("DefaultLocalVolume").Click 1,1,"LEFT"
												wait(2)
												'Count number of items in a List and select required.
												bReturn = objOrgGrp.JavaList("Volumes").GetROProperty("items count")			    				
												For iCounter=0 to bReturn -1
														If Trim(lcase(objOrgGrp.JavaList("Volumes").GetItem(iCounter))) = Trim(lcase(sDefaultLocalVol)) Then
															objOrgGrp.JavaList("Volumes").Activate sDefaultLocalVol
															Exit For
														End If
												Next
											End If
											If  objOrgGrp.javaButton("Next").GetROProperty("enabled")=1 Then
															'To Click on Next Button
															Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Next")
															'To Click on Yes Button
															Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Yes")
															'To Click on OK Button
															JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Group(s) added"
															JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
											End If
                                            'To Click on Finish Button
											If objOrgGrp.javaButton("Finish").Exist Then
															Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Finish")
											End If
											Call Fn_Button_Click("Fn_Org_SubGroupOperations",objOrgGrp,"Close")
											Fn_Org_SubGroupOperations = TRUE
								End If																		
                Case "Remove"
							Set objjApplet = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet")
							'To remove Selected item Organization
							Call Fn_Button_Click("Fn_Org_SubGroupOperations",objjApplet,"Remove_Subgroup")
							Fn_Org_SubGroupOperations = TRUE
				Case "SearchVerify"
							Call Fn_Button_Click("Fn_Org_SubGroupOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"),"AddSub-Group")	
							If Fn_UI_ObjectExist("Fn_Org_SubGroupOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Group Wizard")) =True Then
								Set objaddsubgroup = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization Group Wizard")
	                                'Select the "AddExisting GroupToTheGroupRole" radio button
									Call Fn_UI_JavaRadioButton_SetON("Fn_Org_SubGroupOperations", objaddsubgroup, "Add existing group as")
                                    'Click on Next button
									Call Fn_Button_Click("Fn_Org_SubGroupOperations", objaddsubgroup, "Next")
									'Search the PersonName is present in User List.
									aColname = split(sSubGrpName, ":",-1,1)
									iCount = Ubound(aColname)
										For iCounter=0 to iCount
												'Click on reset Button
												Call Fn_Button_Click("Fn_Org_SubGroupOperations", objaddsubgroup, "reset")
												'Set  Search Edit-Box.
												Call Fn_Edit_Box("Fn_Org_SubGroupOperations",objaddsubgroup,"SearchBox","")
												'Set  Search Edit-Box.
												Call Fn_Edit_Box("Fn_Org_SubGroupOperations",objaddsubgroup,"SearchBox",aColname(iCounter))
												'Click on Search Button
												Call Fn_Button_Click("Fn_Org_SubGroupOperations", objaddsubgroup, "SearchGroup")
												'Set TOProperty of Search window.
												JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog").SetTOProperty "text","Search"
												If Fn_UI_ObjectExist("Fn_Org_SubGroupOperations", JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"))=true Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iCounter)&" does not exist in the group list")
													Fn_Org_SubGroupOperations = FALSE
														Set objaddsubgroup = Nothing
													Exit Function
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iCounter)&" exist in the group list")
												End If												
										Next									
									'Click on Close button
									Call Fn_Button_Click("Fn_Org_SubGroupOperations", objaddsubgroup, "Close")
									Fn_Org_SubGroupOperations = TRUE
									Set objaddsubgroup = Nothing
							End If			
							End Select										
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Group [" + sSubGrpName + "] of function Fn_Org_SubGroupOperations")
	Set objOrgGrp = Nothing
	Set objjApplet = Nothing
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_SetupWizard()
'###
'###	PREQUISITE			  :		DBA User has successfully navigated at step1 of Setup Wizard.
'###
'###    DESCRIPTION        :   Function will able user to perform following operation using setup wizard
'###												Creating Persons
'###												Creating Roles
'###												Creating Volumes
'###												Creating Group
'###												Creating Users
'###
'###    PARAMETERS      :   1. iSteps:
'###											 2.	sinputfilePath:
'###											 3. sPrimaryDelimiter: 
'###										    4. sSecondaryDelimiter:
'###										    5. sUserID:
'###										   6.  sPersonName:
'###										   7.  sAddress:
'###										  8.  sCity:
'###										  9.  sState:
'###										10. sZipCode:
'###										11. sCountry:
'###										12. sOrganization:
'###										13. sEmpNumber:
'###										14. sintMailCode:
'###										15. sEmail:
'###										16. sTelephone:
'###										17. sGroupRoles:
'###										18. bCreateVol:
'###										19. bSameGroup:
'###										20.sSelectGroupRole
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           04/06/2010         1.0
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  Ketan Raje				07/06/2010			2.0
'###
'###    EXAMPLE          : 		Call Fn_SetupWizard("8", "D:\Ketan.txt", ",", ",", "UserID", "PersonName", "Address", "City", "state", "ZipCode", "Country", "organization", "EmpNumber", "intMailCode", "Email", "Telephone", "TestGroup1:dba/DBA:KG4/none:KG1/KR1", "No", "ON","TestGroup1")
'#############################################################################################################

Public Function Fn_SetupWizard(iStep, sinputfilePath, sPrimaryDelimiter, sSecondaryDelimiter, sUserID, sPersonName, sAddress, sCity, sState, sZipCode, sCountry, sOrganization, sEmpNumber, sintMailCode, sEmail, sTelephone, sGroupRoles, bCreateVol, bSameGroup, sSelectGroupRole)
	GBL_FAILED_FUNCTION_NAME="Fn_SetupWizard"
	Dim objSetupWizard, iCounter, bReturn, iCount, iDefcount, aColname, aCols, iRowData, bFlag, iCnt
	iCnt=0
	Set objSetupWizard = Fn_UI_ObjectCreate("Fn_SetupWizard", JavaWindow("Setup Wizard - Teamcenter").JavaWindow("SWApplet"))
				'If Fn_UI_ObjectExist("Fn_SetupWizard",objSetupWizard)=false Then
				If Fn_SISW_UI_Object_Operations("Fn_SetupWizard","Exist",objSetupWizard,SISW_MINLESS_TIMEOUT) = False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SetupWizard function failed")
						Set objSetupWizard = nothing
						Fn_SetupWizard = FALSE
						Exit Function						
				End If
				'Step : 1
				If objSetupWizard.JavaButton("Next").GetROProperty("enabled")=1 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Next button is enabled")
					iCnt = iCnt+1
					If iStep=Cstr(iCnt) Then
						Set objSetupWizard = nothing
						Fn_SetupWizard = TRUE
						Exit Function
					End If
					'Click on next button
					Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Next")
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Next button is disbaled")
					Set objSetupWizard = nothing
					Fn_SetupWizard =FALSE
					Exit Function
				End If
				'Step : 2
				'Set PrimaryDelimiter Edit box
				Call  Fn_Edit_Box("Fn_SetupWizard",objSetupWizard,"PrimaryDelimiter",sPrimaryDelimiter)
				'Set SecondaryDelimiter Edit box
				Call  Fn_Edit_Box("Fn_SetupWizard",objSetupWizard,"SecondaryDelimiter",sSecondaryDelimiter)
				'Set Input File Edit box
				Call  Fn_Edit_Box("Fn_SetupWizard",objSetupWizard,"SelectInputFile",sinputfilePath)
				'Press enter key
				'JavaWindow("Setup Wizard - Teamcenter").JavaWindow("SWApplet").JavaEdit("SelectInputFile").PressKey micEnter
				JavaWindow("Setup Wizard - Teamcenter").JavaWindow("SWApplet").JavaEdit("SelectInputFile").Activate
					iCnt = iCnt+1
					If iStep=Cstr(iCnt) Then
						Set objSetupWizard = nothing
						Fn_SetupWizard = TRUE
						Exit Function
					End If
				'Click on next button
				Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Next")
				'Step : 3
				'Select from UserID list
				If sUserID<>"" Then
					bReturn = objSetupWizard.JavaList("User Id").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("User Id").GetItem(iCount))) = Trim(lcase(sUserID)) Then							
							objSetupWizard.JavaList("User Id").Select sUserID
							Exit For
						End If
					Next
				End If
				'Select from Person Name list
				If sPersonName<>"" Then
					bReturn = objSetupWizard.JavaList("Person Name").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("Person Name").GetItem(iCount))) = Trim(lcase(sPersonName)) Then
							objSetupWizard.JavaList("Person Name").Select sPersonName
							Exit For
						End If
					Next
				End If
				'Select from Address list
				If sAddress<>"" Then
					bReturn = objSetupWizard.JavaList("Address").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("Address").GetItem(iCount))) = Trim(lcase(sAddress)) Then
							objSetupWizard.JavaList("Address").Select sAddress
							Exit For
						End If
					Next
				End If
				'Select from City list
				If sCity<>"" Then
					bReturn = objSetupWizard.JavaList("City").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("City").GetItem(iCount))) = Trim(lcase(sCity)) Then
							objSetupWizard.JavaList("City").Select sCity
							Exit For
						End If
					Next
				End If
				'Select from State list
				If sState<>"" Then
					bReturn = objSetupWizard.JavaList("State").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("State").GetItem(iCount))) = Trim(lcase(sState)) Then
							objSetupWizard.JavaList("State").Select sState
							Exit For
						End If
					Next
				End If
				'Select from Zip Code list
				If sZipCode<>"" Then
					bReturn = objSetupWizard.JavaList("Zip Code").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("Zip Code").GetItem(iCount))) = Trim(lcase(sZipCode)) Then
							objSetupWizard.JavaList("Zip Code").Select sZipCode
							Exit For
						End If
					Next
				End If
				'Select from Country list
				If sCountry<>"" Then
					bReturn = objSetupWizard.JavaList("Country").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("Country").GetItem(iCount))) = Trim(lcase(sCountry)) Then
							objSetupWizard.JavaList("Country").Select sCountry
							Exit For
						End If
					Next
				End If
				'Select from Organization list
				If sOrganization<>"" Then
					bReturn = objSetupWizard.JavaList("Organization").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("Organization").GetItem(iCount))) = Trim(lcase(sOrganization)) Then
							objSetupWizard.JavaList("Organization").Select sOrganization
							Exit For
						End If
					Next
				End If
				'Select from Employee Number list
				If sEmpNumber<>"" Then
					bReturn = objSetupWizard.JavaList("Employee Number").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("Employee Number").GetItem(iCount))) = Trim(lcase(sEmpNumber)) Then
							objSetupWizard.JavaList("Employee Number").Select sEmpNumber
							Exit For
						End If
					Next
				End If
				'Select from Internal Mail Code list
				If sintMailCode<>"" Then
					bReturn = objSetupWizard.JavaList("Internal Mail Code").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("Internal Mail Code").GetItem(iCount))) = Trim(lcase(sintMailCode)) Then
							objSetupWizard.JavaList("Internal Mail Code").Select sintMailCode
							Exit For
						End If
					Next
				End If
				'Select from EMail list
				If sEmail<>"" Then
					bReturn = objSetupWizard.JavaList("EMail").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("EMail").GetItem(iCount))) = Trim(lcase(sEmail)) Then
							objSetupWizard.JavaList("EMail").Select sEmail
							Exit For
						End If
					Next
				End If
				'Select from Telephone list
				If sTelephone<>"" Then
					bReturn = objSetupWizard.JavaList("Telephone").GetROProperty("items count")			    				
					For iCount=0 to bReturn -1
						If Trim(lcase(objSetupWizard.JavaList("Telephone").GetItem(iCount))) = Trim(lcase(sTelephone)) Then
							objSetupWizard.JavaList("Telephone").Select sTelephone
							Exit For
						End If
					Next
				End If
					iCnt = iCnt+1
					If iStep=Cstr(iCnt) Then
						Set objSetupWizard = nothing
						Fn_SetupWizard = TRUE
						Exit Function
					End If
				'Click on next button
				Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Next")
				'Step : 4
				'Select Groups and Roles
				If sGroupRoles<>"" Then
							'Extract no of columns in Defined List
							iDefcount = objSetupWizard.JavaList("ExistingGroupRole").GetROProperty("items count")
							'Extract the index of row at which the object exist.
							aColname = split(sGroupRoles, ":",-1,1)
							iCount = Ubound(aColname)
							For iRowData=0 to iCount
								bFlag = false
								For iCounter=0 to iDefcount-1
										If Trim(lcase(objSetupWizard.JavaList("ExistingGroupRole").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objSetupWizard.JavaList("ExistingGroupRole").Select aColname(iRowData)
											'Click on Add column Button
											Call Fn_Button_Click("Fn_SetupWizard", objSetupWizard, "Add")
											bFlag = true
											Exit For 
										End If
								 Next							
								 If bFlag=false Then
												aCols = split(aColname(iRowData), "/",-1,1)
												If aCols(0)<>"" Then
													'Set Group to be added.
													Call  Fn_Edit_Box("Fn_SetupWizard",objSetupWizard,"Group",aCols(0))
												End If
												If Trim(lcase(aCols(1)))<>Trim(lcase("None")) Then
													'Set Role to be added.
													Call  Fn_Edit_Box("Fn_SetupWizard",objSetupWizard,"Role",aCols(1))
												End If
												'Click on Add Group/Role button. 
												Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Add GroupRole")		
								 End If										
							Next
				End If
					iCnt = iCnt+1
					If iStep=Cstr(iCnt) Then
						Set objSetupWizard = nothing
						Fn_SetupWizard = TRUE
						Exit Function
					End If
				'Click on next button
				Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Next")
				'Step : 5	
				'Select the Create Volume Radio button.
				Call Fn_UI_JavaRadioButton_SetON("Fn_SetupWizard",objSetupWizard, bCreateVol)
					iCnt = iCnt+1
					If iStep=Cstr(iCnt) Then
						Set objSetupWizard = nothing
						Fn_SetupWizard = TRUE
						Exit Function
					End If
				'Click on next button
				Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Next")
				'Step : 7
				'Set "Same Group for all users" Check Box  value to ON or OFF.
				Call Fn_CheckBox_Set("Fn_SetupWizard", objSetupWizard, "Same Group for all Users", bSameGroup)
				'Set GroupRoles in the table
				objSetupWizard.JavaTable("Step7SelectGroupRole").ActivateRow "0"
                objSetupWizard.JavaTable("Step7SelectGroupRole").SetCellData "0","2",sSelectGroupRole
					iCnt = iCnt+1
					If iStep=Cstr(iCnt) Then
						Set objSetupWizard = nothing
						Fn_SetupWizard = TRUE
						Exit Function
					End If
				'Click on next button
				Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Next")
				'Step : 8
				'Click on "Yes" button.
				Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Yes")
				'Set property of the window
				JavaWindow("Setup Wizard - Teamcenter").Dialog("ErrorDialog").SetTOProperty "text","Setup Wizard"
				'Cilck on OK button of Setup Wizard msgbox.
				JavaWindow("Setup Wizard - Teamcenter").Dialog("ErrorDialog").WinButton("OK").Click 1,1,micLeftBtn
				'Click on Home button
				Call Fn_Button_Click("Fn_SetupWizard",objSetupWizard,"Home")
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Sucessfully completed function Fn_SetupWizard")
	Fn_SetupWizard = TRUE
	Set objSetupWizard = nothing 
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Org_UserOperations()  
'###
'###    DESCRIPTION        :   Create/Modify/Delete/AddExisting/SearchVerify.
'###	Prequisite 					:	1.Organization Prespective is Open.
'###										     2.For Modify/Delete User Should be selected
'###											3.For AddExisting and SearchVerify Group/Role should be selected in Organization tree
'###
'###    PARAMETERS      :  1.sAction:Create/Modify/Delete/AddExisting
'###  											2.sSearch:
'###  											3.sPersonName:For Search ":"Seprated Users can be passed sErrorMessage
'###  											4.sUserID:
'###  											5.sOSName
'###  											6.sPassword:
'###  											7.bClearThePassword:
'###  											8.sLastLoginTime:
'###  											9.sDefaultGroup:
'###  											10.sDefaultVolume:
'###  											11.sDefaultLocalVolume:
'###  											12.sUserStatus:
'###  											13.sChangeOwnershipto:
'###  											14.sIPClearance
'###  											15.sGovtClearence:
'###  											16.sTTCDate:
'###  											17.sGeography:
'###  											18.sNationality:
'###  											19.sGrpMemberSetting: ON:Active:OFF:OFF:ON
'###  											20.sLicensingLevel: 
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :   Swapna Ghatge           07/06/2010         1.0
'###
'###    REVIWED BY     :    Ketan Raje
'###
'###    MODIFIED BY   :  	Sushma - Added the code for sGrpMemberSetting on 21st Dec 2010
'###										Shreyas - Added code to handle Active OR Inactive All Members Dialog on 9-Aug-2011.
'###					
'###									Swapna Ghatge - Modifeid code in case "AddExisting"      7-Dec-2011.
'###									Swapna Ghatge - Modifeid code in case " Modify"				  8-Dec-2011.
'###									Pranav Shirode - Added Case "AddNew"							  21-Jan-2012
'###									Rima Patil - Modified "Delete" case 										23-Aug-2012
'###    EXAMPLE          : 		Case "Create" : Call Fn_Org_UserOperations("Create" , "" , "abc12" , "abc12" , "abc" , "abc" , "OFF" , "10:30" , "Engineering" , "volume1", "", "Active", "", "secret", "secret", "", "abc",  "Indian", "", "Viewer")
'###										 Case "Modify" : Call Fn_Org_UserOperations("Modify" , "" , "om1" , "" , "" , "" , "" , "" , "" , "volume1", "", "", "", "", "", "", "",  "", "", "")
'###										 Case "Delete" : Call Fn_Org_UserOperations("Delete" , "" , "" , "" , "" , "" , "" , "" , "" , "", "", "", "", "", "", "", "",  "", "", "")
'###										Case  "AddExisting" : Call Fn_Org_UserOperations("AddExisting" , "PersonName (userid)" , "" , "" , "" , "" , "" , "" , "" , "", "", "", "", "", "", "", "",  "", "", "")
'###										Case "SearchVerify" : Call Fn_Org_UserOperations("SearchVerify" , "" , "PersonName" , "" , "" , "" , "" , "" , "" , "", "", "", "", "", "", "", "",  "", "", "")
'###										Case "Remove" : Call Fn_Org_UserOperations("Remove" , "" , "" , "" , "" , "" , "" , "" , "" , "", "", "", "", "", "", "", "",  "", "", "")
'###										Case "Verify" : Fn_Org_UserOperations("Verify" , "" , "" , "" , "" , "" , "" , "" , "" , "", "", "", "", "super-secret", "super-secret", "No Date Set.", "",  "", "", "Author")
'###										Case "AddNew" 	:Fn_Org_UserOperations("AddNew" , "", "Test", "Test", "Test", "Test", "","", "dba:DBA", "volume1", "","","","","","","","","","Consumer")
'###							Avinash Jagale "Modified the  object hirachy in  the Add new " case [Tc91 to 10.0]
'###							
'#############################################################################################################
Public Function Fn_Org_UserOperations(sAction , sSearch , sPersonName , sUserID , sOSName , sPassword , bClearThePassword , sLastLoginTime , sDefaultGroup , sDefaultVolume, sDefaultLocalVolume, sUserStatus, sChangeOwnershipto, sIPClearance, sGovtClearence, sTTCDate, sGeography,  sNationality, sGrpMemberSetting, sLicensingLevel)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_UserOperations"
	Dim objUser,objdelete,objcreate,objdelete1,objadduser, arrGrpMemberSetting, iCount, iCounter, sGroup,objPerson, sStatus
	Dim DicCustom,objTable,objChild,objEdit,arrDate,arrDateActual
	Dim aValues, objSiteSelection, iRowData, iListItemCount, iListCount, bFlag,var
	Set DicCustom = CreateObject("Scripting.Dictionary")
	Set objUser = Fn_UI_ObjectCreate("Fn_Org_UserOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
	If vartype(sAction) = "9" Then
		Set DicCustom = sAction
		sAction = DicCustom("Action")
	End If
		Select Case sAction
				Case "Create", "Modify","ModifyExt","CreateExt"
						If sPersonName<>"" Then	
						'Set  Person Name
							call Fn_Edit_Box("Fn_Org_UserOperations",objUser,"PersonName",sPersonName)
						End If

						If sUserID<>"" Then
							'Set  User ID 
							call Fn_Edit_Box("Fn_Org_UserOperations",objUser,"UserID",sUserID)
						End If
							If sOSName<>"" Then
								'Set  OS Name 
								call Fn_Edit_Box("Fn_Org_UserOperations",objUser,"OSName",sOSName)
							End If
							'ClearThe Password
							'Commented code as Clear The Password checkbox remove from application : Build : 2013041700
'							Call Fn_CheckBox_Set("Fn_Org_UserOperations", objUser,"ClearThePassword", bClearThePassword)
							If sPassword<>"" Then
								'Set  Password
								If sAction = "CreateExt" Then
									'call Fn_SISW_UI_JavaEdit_Operations("Fn_Org_UserOperations","SetExt",objUser,"Password",sPassword)
									objUser.JavaEdit("Password").Set ""
									objUser.JavaEdit("Password").Set sPassword
								Else
								    call Fn_Edit_Box("Fn_Org_UserOperations",objUser,"Password",sPassword)
								End If
							End If
							If sLastLoginTime<>"" Then
								'Set  Last Login Time
								call Fn_Edit_Box("Fn_Org_UserOperations",objUser,"LastLoginTime",sLastLoginTime)
						End If
						  If sDefaultGroup<>"" Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Modified By Ketan on 1st Aug due to OR changes on build 20110720 on TC 9.1''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'							If sAction = "Modify"   Then
'							objUser.JavaCheckBox("DefaultGroup_User").SetTOProperty "toolkit class","com.teamcenter.rac.organization.OrgUserPanel$7"
'							Elseif sAction = "Create" Then
'							objUser.JavaCheckBox("DefaultGroup_User").SetTOProperty "toolkit class","com.teamcenter.rac.organization.UserPanel$6"
'							End If
							'Set Default Group
							objUser.JavaCheckBox("DefaultGroup_User").Click 1,1,"LEFT"
							'Count number of items in a List and select required.
							'wait(5)
							wait(2)
							bReturn = objUser.JavaList("DefaultGroup_User").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objUser.JavaList("DefaultGroup_User").GetItem(iCounter))) = Trim(lcase(sDefaultGroup)) Then
									objUser.JavaList("DefaultGroup_User").Activate sDefaultGroup
									Exit For
								End If
							Next
						End If
						  If sDefaultVolume<>"" Then
							'Set Default Volume
							objUser.JavaCheckBox("DefaultVolume_User").Click 1,1,"LEFT"
							'Count number of items in a List and select required.
							wait(2)
							bReturn = objUser.JavaList("DefaultVolume_User").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objUser.JavaList("DefaultVolume_User").GetItem(iCounter))) = Trim(lcase(sDefaultVolume)) Then
									objUser.JavaList("DefaultVolume_User").Activate sDefaultVolume
									Exit For
								End If
							Next
						End If
						If sDefaultLocalVolume<>"" Then
							'Set Default Local Volume
							objUser.JavaCheckBox("DefaultLocalVolume_User").Click 1,1,"LEFT"
							wait(2)
							'Count number of items in a List and select required.
							bReturn = objUser.JavaList("DefaultLocalVolume_User").GetROProperty("items count")			    				
							For iCounter=0 to bReturn -1
								If Trim(lcase(objUser.JavaList("DefaultLocalVolume_User").GetItem(iCounter))) = Trim(lcase(sDefaultLocalVolume)) Then
									objUser.JavaList("DefaultLocalVolume_User").Activate sDefaultLocalVolume
									Exit For
								End If
							Next
						End If
						If sUserStatus<>"" Then
								If instr(1,sUserStatus,":")>0 Then
									aUser=split(sUserStatus,":",-1,1)
									sStatus =  aUser(0)
									'To Select User	Status
								Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations",objUser, aUser(0))
								Else
								'To Select User	Status
									sStatus =  sUserStatus
									Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations",objUser, sUserStatus)		
								End If
						End If
						'Change OwnerShip to
						'objUser.JavaCheckBox("ChangeOwnershipTo").Click 1,1,"LEFT"
						If sIPClearance<>"" Then
							If sIPClearance = "none"  Then
								'To set IP Clearance
'								Call Fn_List_Select("Fn_Org_UserOperations", objUser, "IPClearance", "") 'As it is not Working
								objUser.JavaList("IPClearance").Select ""
							Else
								Call Fn_List_Select("Fn_Org_UserOperations", objUser, "IPClearance", sIPClearance)	
							End If
						End If
						If sGovtClearence<>"" Then
							If sGovtClearence = "none"  Then
								'To set Gov't Clearance
'								Call Fn_List_Select("Fn_Org_UserOperations", objUser, "Gov'tClearance:", "") 'As it is not Working
								objUser.JavaList("Gov'tClearance:").Select ""
							Else
								Call Fn_List_Select("Fn_Org_UserOperations", objUser, "Gov'tClearance:", sGovtClearence)
							End IF
						End If
						If sTTCDate<>"" Then
							If sTTCDate = "none"  Then
								'To set TTC Date
								'JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaCheckBox("Date_User").Object.setDate ""
								'[TC1123-20161122-02_12_2016-VivekA-NewDevelopment] - REG-Admin development
								objUser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text", "TTC Date:"
								If objUser.JavaEdit("CustomJavaEdit").exist Then
									objUser.JavaEdit("CustomJavaEdit").Set ""
								End If
								Call Fn_KeyBoardOperation("SendKeys","{TAB}")
								'--------------------------------------------------------
							Else
								'JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaCheckBox("Date_User").Object.setDate sTTCDate
								'[TC1123-20161122-02_12_2016-VivekA-NewDevelopment] - REG-Admin development
								objUser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text", "TTC Date:"
								If objUser.JavaEdit("CustomJavaEdit").exist Then
									objUser.JavaEdit("CustomJavaEdit").Set sTTCDate
								End If
								Call Fn_KeyBoardOperation("SendKeys","{TAB}")
								'--------------------------------------------------------
							End If
						End If
						'Set Geography
						If sGeography<>"" Then
							call Fn_Edit_Box("Fn_Org_UserOperations",objUser,"Geography",sGeography)
						End If
						If sNationality<>"" Then
							'Set  Nationality
							call Fn_Edit_Box("Fn_Org_UserOperations",objUser,"Nationality",sNationality)
						End If
						''=======================================================
						''Added By Sushma for TestCase ChangeDefaultGroupRole (MyTeamcenter) on 21st Dec 2010
						If sGrpMemberSetting<>"" Then
							 arrGrpMemberSetting = split(sGrpMemberSetting,":",-1,1)
							 If  lcase(arrGrpMemberSetting(0))<>"none" Then
									Call Fn_CheckBox_Set("Fn_Org_UserOperations" ,objUser,"Group Administrator_GMS", arrGrpMemberSetting(0))
							 End If
							 If  lcase(arrGrpMemberSetting(1))<>"none" Then
									If trim(lcase(arrGrpMemberSetting(1))) = "active" Then
										Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations",objUser, "Active_GMS")	
									Else
										Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations",objUser, "Inactive_GMS")     
									End If	
							 End If
							 If  lcase(arrGrpMemberSetting(2))<>"none" Then
									Call Fn_CheckBox_Set("Fn_Org_UserOperations" ,objUser,"DefaultRole_GMS", arrGrpMemberSetting(2)) 
							 End If
							 If  lcase(arrGrpMemberSetting(3))<>"none" Then
									Call Fn_CheckBox_Set("Fn_Org_UserOperations" ,objUser,"Externally Managed_GMS", arrGrpMemberSetting(3)) 
							 End If
						End If
						''=======================================================
						If sLicensingLevel<>"" Then
								'To Set Licensing Level	
								Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations",objUser, sLicensingLevel)		
						End If
						
						If DicCustom.count > 0 Then		' Added by Reema W - [ TC1015-2015071400-10_08_2015-VivekA-NewDevlopment ]
							'[TC1123-20161122-02_12_2016-VivekA-NewDevelopment] - REG-Admin development
							If DicCustom.Exists("AddAdditionalProperties") Then
								If DicCustom("AddAdditionalProperties")="Yes" Then
									If objUser.JavaStaticText("AddAdditionalProperties").Exist Then
										objUser.JavaStaticText("AddAdditionalProperties").Click 1,1
									End If
								ElseIf DicCustom("AddAdditionalProperties")="No" Then
									'Do Nothing
								End If
								DicCustom.Remove("AddAdditionalProperties")
							Else  '-------------------------------------------------------------------
								If objUser.JavaStaticText("AddAdditionalProperties").Exist Then
									objUser.JavaStaticText("AddAdditionalProperties").Click 1,1
								End If
							End If
							
							For Each Elem in DicCustom
								Select Case Elem
									Case "Char1", "Char1_ITAR", "Char2_ITAR","Char3_ITAR","Char4_ITAR",_
										 "Double1", "Double1_ITAR","Double2_ITAR","Double3_ITAR","Double4_ITAR","Double5_ITAR",_
										 "Integer1","Integer1_ITAR","Integer2_ITAR","Integer3_ITAR","Integer4_ITAR",_
										 "Candid String",_
										 "LongString1",_
										 "String1","String1_ITAR","String2_ITAR","String3_ITAR","String4_ITAR",_
										 "Unique Int1","Unique Int2", "License Server:","License Bundle:"
										 '"Char2","Char3","Char4", 
										 '"Double2","Double3","Double4","Double5",
										 '"Integer2","Integer3","Integer4",
										 '"String2","String3","String4",
										objUser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text", Elem
										wait 1
										If objUser.JavaEdit("CustomJavaEdit").exist Then
											objUser.JavaEdit("CustomJavaEdit").Set DicCustom(Elem)
										End If
									Case "LOV3 Ind1","LOV3 Ind1 Sub","LOV4 Ind1","LOV4 Ind1 Sub","string LOV_ITAR","sub LOV1_ITAR","LOV5 Ind1","LOV5 Ind1 Sub"
										objUser.JavaStaticText("CustomJavaStaticText").SetTOProperty "label",Elem
										wait 1
										objUser.JavaButton("CustomLOVdropdwnButton").Click	
										wait 2
										Set objTable = Description.Create()
										objTable("Class Name").value = "JavaTable"
'										objTable("tagname").value="LOVTreeTable"
										objTable("toolkit class").value="com.teamcenter.rac.common.lov.view.components.LOVTreeTable"
										set objChild = objUser.ChildObjects(objTable)
										If  objChild.count > 0 Then
											For iCounter=0 to objChild(0).GetROProperty("rows")
	 											If trim(DicCustom(Elem))=trim(objChild(0).Object.getValueAt(iCounter,0).getDisplayableValue()) Then
													objChild(0).DoubleClickCell iCounter,0
													Exit for
												End If
											Next
										End If
										Set objTable=Nothing
										Set objChild=Nothing
									Case "Boolean1_ITAR" '"Boolean1", 
										objUser.JavaStaticText("CustomJavaStaticText").SetTOProperty "label",Elem
										wait 1
										objUser.JavaRadioButton("CustomRadioButton").SetTOProperty "attached text",DicCustom(Elem)
										wait 1
										objUser.JavaRadioButton("CustomRadioButton").Set "ON"
									Case "Date1","Date1_ITAR","Date2_ITAR","Date3_ITAR"
															 '"Date2","Date3", 
										'objUser.JavaEdit("CustomJavaEdit").SetTOProperty "label",Elem
										objUser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text",Elem    ' Added by Jotiba [TC1123-2016062900_8_72016_Maintenace]
										wait 1
										objUser.JavaEdit("CustomJavaEdit").Set DicCustom(Elem)
									
									'[TC1123-20161122-02_12_2016-VivekA-NewDevelopment] - REG-Admin development
									Case "Citizenships"	
											aValues=split(DicCustom("Citizenships"),"~")
											Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_Org_UserOperations",objUser.JavaStaticText("CustomJavaStaticText"),"label", Elem+":")
											Call Fn_CheckBox_Set("Fn_Org_UserOperations", objUser, "ListCheckBox", "ON")
												For iCounter = 0 to UBound(aValues)
													Call Fn_Edit_Box("Fn_Org_UserOperations",objUser,"CheckBoxEdit",aValues(iCounter))
													Call Fn_Button_Click("Fn_Org_UserOperations", objUser,"Add")
												Next
												Call Fn_CheckBox_Set("Fn_Org_UserOperations", objUser, "ListCheckBox", "OFF")
									Case "Home Site"
										Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_Org_UserOperations",objUser.JavaList("CustomListBox"),"attached text", Elem+":")
										objUser.JavaList("CustomListBox").Select DicCustom(Elem)
										wait 1
										
									Case "DenyLoginAtSites"
										aValues=split(DicCustom("DenyLoginAtSites"),"~")
										Call Fn_Button_Click("Fn_Org_UserOperations", objUser,"SelectSites")
										Call Fn_ReadyStatusSync(1)
										Set objSiteSelection=objUser.JavaDialog("SiteSelection")
										If objSiteSelection.Exist(3) Then
											iListItemCount=objSiteSelection.JavaList("AvailableSites").GetROProperty("items count")
												For iRowData=0 to UBound(aValues)
													For iCounter = 0 to iListItemCount-1
														If Trim(lcase(objSiteSelection.JavaList("AvailableSites").GetItem(iCounter))) = Trim(lcase(aValues(iRowData))) Then
																objSiteSelection.JavaList("AvailableSites").Select aValues(iRowData)
																'Click on "+" Button
																Call Fn_Button_Click("Fn_Org_UserOperations", objSiteSelection, "AddSites")											
																Exit For 
														End If
													Next 
												Next
												Call Fn_Button_Click("Fn_Org_UserOperations", objSiteSelection, "OK")	
												Set objSiteSelection=Nothing
										End If
										'-------------------------------------------------------------------
								End Select
							Next
						End If
						'wait 3
						wait 1
						'To Click on Create Button
						If sAction="Create" OR sAction="CreateExt" Then
							If objUser.JavaButton("Create").GetROProperty("enabled") =  1 Then
								Call Fn_Button_Click("Fn_Org_UserOperations",objUser,"Create")															'Added new starts
								JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Error"
								'Added by Yogini on 19-Feb-13
								Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").SetTOProperty "title","Error"

								'If JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").Exist  Then
								If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog"),SISW_MINLESS_TIMEOUT) Then
									JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click
									'Added by Yogini on 19-Feb-13
								'ElseIf Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").Exist  Then
								ElseIf Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error"),SISW_MICRO_TIMEOUT)  Then
									If DicCustom.Exists("CreateError") Then
										If DicCustom("CreateError") =  Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").JavaEdit("DetailMsg").GetROProperty("value") Then
											Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").JavaButton("OK").Click
										Else
											Fn_Org_UserOperations = False
											Exit Function
										End If
									Else
										Window("TeamcenterWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Error").JavaButton("OK").Click
									End If
								Else
									For iCount = 0 to 0
										JavaDialog("Delete").SetTOProperty "title", "Create Person"		'Changed the hierarchy of JavaDialog
										 'If JavaDialog("Delete").Exist  Then	
										  If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",JavaDialog("Delete"),SISW_MINLESS_TIMEOUT) Then										 'Due to build changes in 2011042000
											JavaDialog("Delete").JavaButton("Yes").Click
											Exit For
										End If
										JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").SetTOProperty "title", "Set Person"
										 'If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").Exist  Then
										 If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog"),SISW_MICRO_TIMEOUT) Then	
											JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").JavaButton("Yes").Click
											Exit For
										End If
									Next
									'Code added by shrikant for check existance of JavaDialog "User already exist"
									'Swapnil: 10-JULY-2012 :Hierarchy changed.

										'If JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Error").exist(10) Then
											'JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Error").JavaButton("OK").Click
										'If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error").exist(10) Then
										If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error"),SISW_MINLESS_TIMEOUT) Then
											JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error").JavaButton("OK").Click
											Fn_Org_UserOperations = False
											Exit Function
										End If
									End If
							Else
								Fn_Org_UserOperations = False
								Exit Function
							End If	
						ElseIf sAction="Modify" or sAction="ModifyExt" Then
									Call Fn_Button_Click("Fn_Org_UserOperations",objUser,"Modify")
									If sPersonName<>"" Then
										Set objPerson=Fn_SISW_Org_GetObject("DeleteConfirmation")
'											JavaDialog("DeleteConfirmation").SetTOProperty "title", "Create Person"
'											If  JavaDialog("DeleteConfirmation").Exist Then			'Checked the existance of Dialog
'												JavaDialog("DeleteConfirmation").JavaButton("No").Click
'											End If
											If sAction="Modify" Then
												objPerson.SetTOProperty "title", "Create Person"
												'If  objPerson.Exist Then			'Checked the existance of Dialog
												If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",objPerson,SISW_MINLESS_TIMEOUT) Then
													objPerson.JavaButton("No").Click
													Set objPerson=nothing
												End If
											Else
												objPerson.SetTOProperty "title", "Create Person"
												'If  objPerson.Exist Then			'Checked the existance of Dialog
												If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",objPerson,SISW_MINLESS_TIMEOUT) Then
													objPerson.JavaButton("Yes").Click
													Set objPerson=nothing
												End If											
											End If											
									End If
									'check the exixtence of 'Inactive All Members Dialog
									If sStatus = "Inactive" Then
										JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Inactivate All Members").SetTOProperty "title","Inactivate All Members"
										'If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Inactivate All Members").Exist then
										If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Inactivate All Members"),SISW_MINLESS_TIMEOUT) Then
													JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Inactivate All Members").JavaButton("Yes").Click micLeftBtn								
										End If
									Else
										'check the exixtence of 'Active All Members Dialog
										JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Inactivate All Members").SetTOProperty "title","Activate All Members"
										'If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Inactivate All Members").Exist then
										If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Inactivate All Members"),SISW_MICRO_TIMEOUT) Then
														JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Inactivate All Members").JavaButton("Yes").Click micLeftBtn								
										End If
									End If
						End If
				Case "Delete"
'						'Click on Delete button.
'						Call Fn_Button_Click("Fn_Org_UserOperations", objUser, "Delete")
'						Call Fn_ReadyStatusSync(1)
'						If  JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Delete User").Exist(5)=True Then
'								Set objdelete1=JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Delete User")
'						Else
'                               Set objdelete1= JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Delete User")
'						End If
'                        					
'								Call Fn_Button_Click("Fn_Org_UserOperations", objdelete1 ,"Delete")
'								Call Fn_ReadyStatusSync(1)
''									JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete User Confirmation"
''									JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
'								JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete User Confirmation"			
'								JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("DeleteConfirmation").JavaButton("Yes").Click

'Added new code to function as per hierarchy changes. Can revert to original code if required
'Shreyas 28-08-2012

						Call Fn_Button_Click("Fn_Org_UserOperations", objUser, "Delete")
						 objdelete  =Fn_UI_ObjectExist("Fn_Org_UserOperations", JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Delete User"))
						If objdelete =True Then	
							Set objdelete1= JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Delete User")
								Call Fn_Button_Click("Fn_Org_UserOperations", objdelete1 ,"Delete")
'									JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete User Confirmation"
'									JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
								JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete User Confirmation"			
								JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
						Else
							Set objdelete1= JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Delete User")
								Call Fn_Button_Click("Fn_Org_UserOperations", objdelete1 ,"Delete")

								JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete User Confirmation"			
								JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
						End If
						
				Case "Remove"
						'Click on remove button.
						Call Fn_Button_Click("Fn_Org_UserOperations", objUser, "Remove_User")
						JavaDialog("DeleteConfirmation").SetTOProperty "Title","Remove User Confirmation"
						'If JavaDialog("DeleteConfirmation").Exist(3) Then
						If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist", JavaDialog("DeleteConfirmation"),SISW_MINLESS_TIMEOUT) Then
								Call Fn_Button_Click("Fn_Org_RoleOperations", JavaDialog("DeleteConfirmation"), "Yes")
					    End If
				Case "AddExisting"
							Call Fn_Button_Click("Fn_Org_UserOperations",objUser,"Add User")	
							
							If  Fn_UI_ObjectExist("Fn_Org_UserOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization User Wizard") )=True Then
							'If Fn_UI_ObjectExist("Fn_Org_UserOperations",JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Organization User Wizard")) =True Then
								'Set objadduser = JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Organization User Wizard")
								Set objadduser=JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization User Wizard")
									'Select the "AddExistingUserToTheGroupRole" radio button
									Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations", objadduser, "AddExistingUserToTheGroupRole")
									'Click on Next button
									Call Fn_Button_Click("Fn_Org_UserOperations", objadduser, "Next")
									'Click on reset Button
									Call Fn_Button_Click("Fn_Org_UserOperations", objadduser, "Reset")
									wait(1)
									'To select the Group from Available Groups
									Call Fn_SyncTCObjects()
									Call Fn_List_Select("Fn_Org_UserOperations",objadduser,"Available Users",sSearch)
									'To Click on ">" Button
									Call Fn_Button_Click("Fn_Org_UserOperations",objadduser,">")
									If  objadduser.javaButton("Next").GetROProperty("enabled")=1 Then
										'To Click on Next Button
										Call Fn_Button_Click("Fn_Org_UserOperations",objadduser,"Next")
										'To Click on Yes Button
										Call Fn_Button_Click("Fn_Org_UserOperations",objadduser,"Yes")
										'Below Code Is Commented Out Due to OK WinButton Change to JavaButton in Build Teamcenter 8 (20100707.00)
										'To Click on OK Button
'										JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog").SetTOProperty "text", "User(s) added"
'										JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog").WinButton("OK").Click 1,1,micLeftBtn
										Call Fn_Button_Click("Fn_Org_UserOperations",objadduser.JavaDialog("User(s) added"),"OK")
									End If	
									'To Click on Finish Button
									'If objadduser.javaButton("Finish").Exist Then
									If Fn_SISW_UI_Object_Operations("Fn_Org_UserOperations","Exist", objadduser.javaButton("Finish"),SISW_MINLESS_TIMEOUT) Then
										Call Fn_Button_Click("Fn_Org_UserOperations",objadduser,"Finish")
									End If
									Call Fn_Button_Click("Fn_Org_UserOperations",objadduser,"Close")
							End If	
				Case "AddNew"
						'Click on Add User Button
						Call Fn_Button_Click("Fn_Org_UserOperations",objUser,"Add User")
						Wait(2)
						If Fn_UI_ObjectExist("Fn_Org_UserOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization User Wizard")) =True Then
						Set objadduser = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization User Wizard")
							'Select the "AddExistingUserToTheGroupRole" radio button
							Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations", objadduser, "AddNewUserToTheGroupRole")
							'Click on Next button
							Call Fn_Button_Click("Fn_Org_UserOperations", objadduser, "Next")
							If sPersonName<>"" Then	
							'Set  Person Name
								call Fn_Edit_Box("Fn_Org_UserOperations",objadduser,"PersonName",sPersonName)
							End If
    						If sUserID<>"" Then
								'Set  User ID 
								call Fn_Edit_Box("Fn_Org_UserOperations",objadduser,"UserID",sUserID)
							End If
							If sOSName<>"" Then
								'Set  OS Name 
								call Fn_Edit_Box("Fn_Org_UserOperations",objadduser,"OSName",sOSName)
							End If
							If sPassword<>"" Then
								'Set  Password
								call Fn_Edit_Box("Fn_Org_UserOperations",objadduser,"Password",sPassword)
							End If
							'Set the Group
							If sDefaultGroup<>"" Then
								sGroup=split(sDefaultGroup,":",-1,1)			'=====Split the group to pass role
								objadduser.JavaCheckBox("DefaultGroup").Click 1,1,"LEFT"
								'Count number of items in a List and select required.
								wait(2)
								bReturn = objadduser.JavaList("List").GetROProperty("items count")
								For iCounter=0 to bReturn -1
									If Trim(lcase(objadduser.JavaList("List").GetItem(iCounter))) = Trim(lcase(sGroup(0))) Then
										objadduser.JavaList("List").Activate sGroup(0)
										Exit For
									End If
								Next

								'For Roles
								objadduser.JavaCheckBox("Roles").Click 1,1,"LEFT"
								'Count number of items in a List and select required.
								wait(2)
								bReturn = objadduser.JavaList("List").GetROProperty("items count")
								For iCounter=0 to bReturn -1
									If Trim(lcase(objadduser.JavaList("List").GetItem(iCounter))) = Trim(lcase(sGroup(1))) Then
										objadduser.JavaList("List").Activate sGroup(1)
										Exit For
									End If
								Next
							End If
							'Set the Volume
							If sDefaultVolume<>"" Then
								'Set Default Volume
								objadduser.JavaCheckBox("DefaultVolume").Click 1,1,"LEFT"
								'Count number of items in a List and select required.
								wait(2)
								bReturn = objadduser.JavaList("List").GetROProperty("items count")
								For iCounter=0 to bReturn -1
									If Trim(lcase(objadduser.JavaList("List").GetItem(iCounter))) = Trim(lcase(sDefaultVolume)) Then
										objadduser.JavaList("List").Activate sDefaultVolume
										Exit For
									End If
								Next
							End If
							If sDefaultLocalVolume<>"" Then
								'Set Default Local Volume
								objadduser.JavaCheckBox("DefaultLocalVolume").Click 1,1,"LEFT"
								wait(2)
								'Count number of items in a List and select required.
								bReturn = objadduser.JavaList("List").GetROProperty("items count")
								For iCounter=0 to bReturn -1
									If Trim(lcase(objadduser.JavaList("List").GetItem(iCounter))) = Trim(lcase(sDefaultLocalVolume)) Then
										objadduser.JavaList("List").Activate sDefaultLocalVolume
										Exit For
									End If
								Next
							End If
							If sLicensingLevel<>"" Then
									'To Set Licensing Level	
									Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations",objadduser, sLicensingLevel)		
							End If
							'custom properties added by Reema On 8/4/2015 [ TC1015-2015071400-10_08_2015-VivekA-NewDevlopment ]
							If DicCustom.count > 0 Then		
								Call Fn_Button_Click("Fn_Org_UserOperations", objadduser, "Next")
								
								For Each Elem in DicCustom
									Select Case Elem
										Case "Char1", "Char1_ITAR", "Char2_ITAR","Char3_ITAR","Char4_ITAR",_
											 "Candid String",_
											 "Double1", "Double1_ITAR","Double2_ITAR","Double3_ITAR","Double4_ITAR","Double5_ITAR",_
											 "Integer1", "Integer1_ITAR","Integer2_ITAR","Integer3_ITAR","Integer4_ITAR",_
											 "LongString1",_
											 "String1","String1_ITAR","String2_ITAR","String3_ITAR","String4_ITAR",_
											 "Unique Int1","Unique Int2"
											' "Char2","Char3","Char4",
											'"String2","String3","String4",
											'"Double2","Double3","Double4","Double5",
											'"Integer2","Integer3","Integer4",
											
											objadduser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text", Elem
											wait 1
											If objadduser.JavaEdit("CustomJavaEdit").exist(5) Then
												objadduser.JavaEdit("CustomJavaEdit").Set DicCustom(Elem)
											wait 2
											End If
										Case "LOV3 Ind1","LOV3 Ind1 Sub","LOV4 Ind1","LOV4 Ind1 Sub","string LOV_ITAR","sub LOV1_ITAR","LOV5 Ind1","LOV5 Ind1 Sub"
											objadduser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text", Elem
											wait 1
											If objadduser.JavaEdit("CustomJavaEdit").exist Then
												objadduser.JavaEdit("CustomJavaEdit").SetFocus()
												objadduser.JavaEdit("CustomJavaEdit").Set DicCustom(Elem)
												call Fn_KeyBoardOperation("SendKeys", "{TAB}")
											End If 
										Case "Boolean1_ITAR" '"Boolean1", 
											objadduser.JavaStaticText("CustomJavaStaticText").SetTOProperty "label",Elem
											wait 1
											objadduser.JavaRadioButton("CustomRadioButton").SetTOProperty "attached text",DicCustom(Elem)
											wait 1
											objadduser.JavaRadioButton("CustomRadioButton").Set "ON"
										Case "Date1","Date1_ITAR","Date2_ITAR","Date3_ITAR"
											 '"Date2","Date3", 
											'objadduser.JavaStaticText("CustomJavaStaticText").SetTOProperty "label",Elem
											objadduser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text",Elem    ' Added by Jotiba [TC1123-2016062900_8_72016_Maintenace]
											objadduser.JavaEdit("CustomJavaEdit").set ""
											wait 15
											objadduser.JavaEdit("CustomJavaEdit").Activate
											'objadduser.JavaCheckBox("CustomDateCheckBox").Object.SetDate(DicCustom(Elem))
											objadduser.JavaEdit("CustomJavaEdit").Set(DicCustom(Elem))
									End Select
								Next
							End If
							wait 5
							'To Click on Finish Button
							If objadduser.javaButton("Finish").Exist Then
								Call Fn_Button_Click("Fn_Org_UserOperations",objadduser,"Finish")
							End If
							wait 5
							Call Fn_Button_Click("Fn_Org_UserOperations",objadduser,"Close")
						End If
				Case "SearchVerify"
							Call Fn_Button_Click("Fn_Org_UserOperations",objUser,"Add User")	
							If Fn_UI_ObjectExist("Fn_Org_UserOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization User Wizard")) =True Then
								Set objadduser = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Organization User Wizard")
							''	Modified Hierarchy of Dialog Organization User Wizard By Dipali k 
								''Set objadduser = JavaWindow("Organization - Teamcenter").JavaWindow("OrgWindow").JavaDialog("Organization User Wizard")
									'Select the "AddExistingUserToTheGroupRole" radio button
									Call Fn_UI_JavaRadioButton_SetON("Fn_Org_UserOperations", objadduser, "AddExistingUserToTheGroupRole")
									'Click on Next button
									Call Fn_Button_Click("Fn_Org_UserOperations", objadduser, "Next")
									'Click on reset Button
									Call Fn_Button_Click("Fn_Org_UserOperations", objadduser, "Reset")
									'Search the PersonName is present in User List.
									aColname = split(sPersonName, ":",-1,1)
									iCount = Ubound(aColname)
										For iCounter=0 to iCount
												'Set  Search Edit-Box.
												Call Fn_Edit_Box("Fn_Org_UserOperations",objadduser,"SearchUsers",aColname(iRowData))
												Wait 10
												'Click on Search Button
												Call Fn_Button_Click("Fn_Org_UserOperations", objadduser, "SearchUser")
												'Set TOProperty of Search window.
												JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog").SetTOProperty "text","Search"
												If Fn_UI_ObjectExist("Fn_Org_UserOperations", JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"))=true Then
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData)&" does not exist in the users list")
													Fn_Org_UserOperations = FALSE
													Exit Function
												Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData)&" exist in the users list")
												End If
										Next
									Fn_Org_UserOperations = TRUE
									'Click on Close button
									Call Fn_Button_Click("Fn_Org_UserOperations", objadduser, "Close")
							End If
			Case "Verify"
						iCount = 0
						iCounter = 0
						'Code to handle above fields is not yet coded.
						'Verify  User status 
						If sUserStatus<>"" Then
								iCount = iCount + 1
								If cint(objUser.JavaRadioButton(sUserStatus).GetROProperty("value")) = 1 then
									iCounter = iCounter + 1
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "User Status is [" + sUserStatus + "]")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "IUser Status is not [" + sUserStatus + "]")
								End If
						End If	
						'Verify  User Member Settings.
						If sGrpMemberSetting<>"" Then
								iCount = iCount + 1
								If cint(objUser.JavaRadioButton(sGrpMemberSetting).GetROProperty("value")) = 1 then
									iCounter = iCounter + 1
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Group Member Settings  is [" + sGrpMemberSetting + "]")
								Else
									 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Group Member Settings  is not [" + sGrpMemberSetting + "]")
								End If
						End If
						'Verify OSName
						If sOSName<>"" Then
								iCount = iCount + 1
								If  Trim(objUser.JavaEdit("OSName").GetROProperty("value"))=Trim(sOSName) Then
									iCounter = iCounter + 1
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "OS Name matches matches with actual value")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "OS Name does not matches matches with actual value")
								End If
						End If

						'*Added by Nilesh on 23-Nov-12 to verify Default group
						'Verify Default Volume
						If sDefaultVolume<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objUser.JavaCheckBox("DefaultVolume_User").GetRoProperty("label")))=Trim(Lcase(sDefaultVolume)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Default Group value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Default Group value dose'nt matches with actual value")
							End If
						End If
						'*End

						'Verify IPClearance
						If sIPClearance<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objUser.JavaList("IPClearance").GetItem(objUser.JavaList("IPClearance").Object.getSelectedIndex)))=Trim(Lcase(sIPClearance)) Then
									iCounter = iCounter + 1
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "IP Clearance value matches with actual value")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "IP Clearance value dose'nt matches with actual value")
							End If
						End If
						'Verify Govt Clearance
						If sGovtClearence<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objUser.JavaList("Gov'tClearance:").GetItem(objUser.JavaList("Gov'tClearance:").Object.getSelectedIndex)))=Trim(Lcase(sGovtClearence)) Then
									iCounter = iCounter + 1
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Gov'tClearance value matches with actual value")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Gov'tClearance value dose'nt matches with actual value")
							End If
						End If
						'Verify TTC Date
						If sTTCDate<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objUser.JavaCheckBox("Date_User").GetROProperty("attached text"))) = Trim(Lcase(sTTCDate)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Date value dose'nt matches with actual value")
							End If
						End If
						'Verify  Geography
						If sGeography<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_UserOperations",objUser,"Geography"))) = Trim(Lcase(sGeography)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Geography value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Geography value dose'nt matches with actual value")
							End If
						End If
						'Verify  Nationality
						If sNationality<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_UserOperations",objUser,"Nationality"))) = Trim(Lcase(sNationality)) Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Nationality value matches with actual value")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Nationality value dose'nt matches with actual value")
							End If
						End If		
						'Verify  Licening Level
						If sLicensingLevel<>"" Then
							iCount = iCount + 1
							If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaRadioButton(sLicensingLevel).GetROProperty("value") = 1 Then
								iCounter = iCounter + 1
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sLicensingLevel &" JavaRadioButton is Set ON")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sLicensingLevel &" JavaRadioButton is Set OFF")
							End If
						End If
						'verify custom properties 
						If DicCustom.count > 0 Then		'Added by Reema W : [ TC1015-2015071400-10_08_2015-VivekA-NewDevlopment ]
							If objUser.JavaStaticText("AddAdditionalProperties").Exist Then
								objUser.JavaStaticText("AddAdditionalProperties").Click 1,1
							End If
							For Each Elem in DicCustom
								Select Case Elem
									Case "Char1", "Char1_ITAR", "Char2_ITAR","Char3_ITAR","Char4_ITAR",_
										 "Candid String",_
										 "Integer1", "Integer1_ITAR","Integer2_ITAR","Integer3_ITAR","Integer4_ITAR",_
										 "LongString1",_
										 "String1","String1_ITAR","String2_ITAR","String3_ITAR","String4_ITAR",_
										 "Unique Int1","Unique Int2", "License Server:","License Bundle:","TTC Date:"
										 '"Char2","Char3","Char4", 
										 '"Double2","Double3","Double4","Double5",
										 '"Integer2","Integer3","Integer4",
										 '"String2","String3","String4",
										iCount = iCount + 1
										objUser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text", Elem
										wait 1
										If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_UserOperations",objUser,"CustomJavaEdit"))) = Trim(Lcase(DicCustom(Elem))) Then
											iCounter = iCounter + 1
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ""+Elem+" value matches with actual value")
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ""+Elem+" value dose'nt matches with actual value")
										End If
										
									Case  "Double1", "Double1_ITAR","Double2_ITAR","Double3_ITAR","Double4_ITAR","Double5_ITAR"' Added by Jotiba as per changes - LCS-535656 -Fnd0TrimZeroes 
											iCount = iCount + 1
											objUser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text", Elem
											wait 1
											If instr(1,Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_UserOperations",objUser,"CustomJavaEdit"))),Trim(Lcase(DicCustom(Elem))))>0 Then
												iCounter = iCounter + 1
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ""+Elem+" value matches with actual value")
											ElseIf Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_UserOperations",objUser,"CustomJavaEdit"))) = Trim(Lcase(DicCustom(Elem))) Then    'Added by Rishabh to verify blank custom properties 
											iCounter = iCounter + 1
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ""+Elem+" value matches with actual value")
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ""+Elem+" value dose'nt matches with actual value")
											End If
																				
									Case "LOV3 Ind1","LOV3 Ind1 Sub","LOV4 Ind1","LOV4 Ind1 Sub","string LOV_ITAR","sub LOV1_ITAR","LOV5 Ind1","LOV5 Ind1 Sub"
										iCount = iCount + 1
										objUser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text", Elem
										If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_UserOperations",objUser,"CustomJavaEdit"))) = Trim(Lcase(DicCustom(Elem))) Then
											iCounter = iCounter + 1
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ""+Elem+" value matches with actual value")
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), ""+Elem+" value dose'nt matches with actual value")
										End If
									Case "Boolean1_ITAR" ' "Boolean1", 
										iCount = iCount + 1
										objUser.JavaStaticText("CustomJavaStaticText").SetTOProperty "label",Elem
										wait 1
										objUser.JavaRadioButton("CustomRadioButton").SetTOProperty "attached text",DicCustom(Elem)
										wait 1
										If cint(objUser.JavaRadioButton("CustomRadioButton").GetROProperty("value")) = 1 then
											iCounter = iCounter + 1
										 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Elem + "] is " +DicCustom(Elem)+"")
										Else
										 	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " [" + sGrpMemberSetting + "] is not " +DicCustom(Elem)+"")
										End If	
									Case "Date1", "Date1_ITAR","Date2_ITAR","Date3_ITAR"
										'"Date2","Date3",
										objUser.JavaEdit("CustomJavaEdit").SetTOProperty "attached text",Elem    ' Added by Jotiba [TC1123-2016062900_8_72016_Maintenace]
										wait 1
										iCount = iCount + 1
										If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_UserOperations",objUser,"CustomJavaEdit"))) = Trim(Lcase(DicCustom(Elem))) Then
											iCounter = iCounter + 1
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Elem + "] Date value matches with actual value")
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + Elem + "] Date value dose'nt matches with actual value")
										End If
									'[TC1123-20161122-02_12_2016-VivekA-NewDevelopment] - REG-Admin development
									Case "Home Site"
											iCount = iCount + 1
											If Fn_UI_Object_SetTOProperty_ExistCheck("Fn_Org_UserOperations",objUser.JavaList("CustomListBox"),"attached text", Elem+":") Then
												If Trim(lcase(objUser.JavaList("CustomListBox").GetItem(objUser.JavaList("CustomListBox").Object.getSelectedIndex)))=Trim(Lcase(DicCustom("Home Site")))  Then
													iCounter = iCounter + 1
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Home Site value matches with actual value")
												Else 
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Home Site value dose'nt matches with actual value")
												End If
											End If
									Case "Citizenships"
											iCount = iCount + 1
											aValues=split(DicCustom("Citizenships"),"~")
											If Fn_UI_Object_SetTOProperty_ExistCheck("Fn_Org_UserOperations",objUser.JavaList("CustomListBox"),"attached text", Elem+":") Then
												iListItemCount=Fn_UI_Object_GetROProperty("Fn_Org_UserOperations",objUser.JavaList("CustomListBox"), "items count")
													For iRowData = 0 to UBound(aValues)
														bFlag=False 
														For iListCount = 0 To iListItemCount-1
															If Trim(lcase(objUser.JavaList("CustomListBox").GetItem(iListCount))) = Trim(lcase(aValues(iRowData))) Then
																bFlag = True
																Exit For 
															End If 	
														Next
													Next 
													If bFlag = True Then
														iCounter = iCounter + 1
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Citizenships value matches with actual value")	
													Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Citizenships value dose'nt matches with actual value")
													End If
											End If 
									Case "Deny Login At Sites"	
											iCount = iCount + 1
											aValues=split(DicCustom("Deny Login At Sites"),"~")
											If Fn_UI_Object_SetTOProperty_ExistCheck("Fn_Org_UserOperations",objUser.JavaList("CustomListBox"),"attached text", Elem+":") Then
												iListItemCount=Fn_UI_Object_GetROProperty("Fn_Org_UserOperations",objUser.JavaList("CustomListBox"), "items count")
													For iRowData = 0 to UBound(aValues)
														bFlag=False
														For iListCount = 0 To iListItemCount-1
															If Trim(lcase(objUser.JavaList("CustomListBox").GetItem(iListCount))) = Trim(lcase(aValues(iRowData))) Then
																bFlag = True
																Exit For 
															End If 	
														Next
													Next 
													If bFlag = True Then
														iCounter = iCounter + 1
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Deny Login At Sites value matches with actual value")
													Else
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Deny Login At Site value dose'nt matches with actual value")													
													End If
											End If 	
									'----------------------------------------------------------
								End Select
							Next
						End If
						If iCount = iCounter Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All values verified successfully")
							Fn_Org_UserOperations = TRUE
							Set objUser = nothing 
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All values does not match successfully")
							Fn_Org_UserOperations = FALSE
							Set objUser = nothing 
							Exit Function
						End If	
		Case "GetIndex"
			'Index of   User
			For iCounter=0 to JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaTree("CategoryTree").GetROProperty ("items count")-1
				sCmpUser=JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaTree("CategoryTree").GetItem (iCounter)
				If instr(1,sCmpUser,sSearch)<> 0 Then
					Fn_Org_UserOperations=cint(iCounter)
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Sucessfully found the index of user"+sSearch)
					Set objUser = nothing 
					Set objdelete = nothing
					Set objdelete1 = nothing  
					Set objcreate = nothing 
					Set objadduser = nothing 
					Exit Function	
				End If
			Next
			If iCounter=JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaTree("CategoryTree").GetROProperty ("items count")-1 Then
				Fn_Org_UserOperations=False
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Could not find the index of user"+sSearch) 
			End If
		Case "VerifyDefaultGrpList"
				If sDefaultGroup<>"" Then
					aValues = split(sDefaultGroup,"~")
					objUser.JavaCheckBox("DefaultGroup_User").Click 1,1,"LEFT"
					wait(2)
					bReturn = objUser.JavaList("DefaultGroup_User").GetROProperty("items count")			    				
					For iCount = 0 to ubound(aValues)
						bFlag = False
						For iCounter=0 to bReturn -1
							If Trim(lcase(objUser.JavaList("DefaultGroup_User").GetItem(iCounter))) = Trim(lcase(aValues(iCount))) Then
								bFlag = True
								Exit For
							End If
						Next
						If bFlag = False Then
							Fn_Org_UserOperations=False
							Set objUser = nothing 
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Could not find Group"+aValues(iCount))
							Exit Function	
						Else
							Fn_Org_UserOperations=True
						End if
					Next
				    objUser.JavaCheckBox("DefaultGroup_User").Click 1,1,"LEFT"
					wait(2)
			 	End If
			Case "GetDefaultGrpListCount"
					objUser.JavaCheckBox("DefaultGroup_User").Click 1,1,"LEFT"
					wait(2)
					Fn_Org_UserOperations = objUser.JavaList("DefaultGroup_User").GetROProperty("items count")
					objUser.JavaCheckBox("DefaultGroup_User").Click 1,1,"LEFT"
					wait(2)
					Set objUser = Nothing
					Exit Function
			Case "ClearDefaultGrp"
					objUser.JavaCheckBox("DefaultGroup_User").Click 1,1,"LEFT"
					wait(2)
					set var = Description.Create()
					var("Class Name").value = "JavaButton"
					set objChild = objUser.ChildObjects(var)
					For iCount = 0 To objChild.count - 1
						If trim(objChild(iCount).GetROProperty("path")) = "JButton;JPanel;JPanel;AbstractPopupButton$6;JPanel;JLayeredPane;JRootPane;Popup$HeavyWeightWindow;WEmbeddedFrame;" Then
							objChild(iCount).Click
							Wait 2
							Set objChild = Nothing
							Set var = Nothing
							Fn_Org_UserOperations=True
						End If
					Next 
					objUser.JavaCheckBox("DefaultGroup_User").Click 1,1,"LEFT"
					wait(2)						
		Case Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_UserOperations function failed")
			Fn_Org_UserOperations = FALSE
			Exit Function
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_Org_UserOperations")
	Fn_Org_UserOperations = TRUE
	Set DicCustom = nothing
	Set objUser = nothing 
	Set objdelete = nothing
	Set objdelete1 = nothing  
	Set objcreate = nothing 
	Set objadduser = nothing 
End Function
'#################################################################################################################
'###    FUNCTION NAME   :    Fn_SetupWizard_MsgVerify(sErrMsg,sDialogTitle)
'###
'###    DESCRIPTION     :   This function used to verify the Dialog messages
'###
'###    PARAMETERS      :   sErrMsg,sDialogTitle
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY         :   AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Harshal      		04/06/2010	  1.0
'###
'###    REVIWED BY      :   Harshal		   		04/06/2010	  1.0          
'###
'###    MODIFIED BY     :
'###    EXAMPLE         :   Fn_SetupWizard_MsgVerify("Setup Wizard","Length for UserId")
'################################################################################################################
Function  Fn_SetupWizard_MsgVerify(sDialogTitle,sErrMsg)

	Dim dicErrorInfo
	Set dicErrorInfo = CreateObject("Scripting.Dictionary")
	With dicErrorInfo 
	 .Add "Title", sDialogTitle
	 .Add "Message", sErrMsg
	 .Add "Button", "OK"
	End with
	Fn_SetupWizard_MsgVerify = Fn_SISW_ErrorVerify(dicErrorInfo)

End Function

'#################################################################################################################
'###    FUNCTION NAME   :     Fn_Org_PersonOperations()  
'###
'###    DESCRIPTION     :   Create/Modify/Delete Persons
'###
'###    PARAMETERS      :  sAction:Create/Modify/Delete/AddExisting
'###                        					sName:
'###                        					sAddress:
'###                        					sCity:
'###                        					sState:
'###                        					sZipCode:
'###                        					sCountry:
'###                        					sOrganization:
'###                        					sEmpNumber:
'###                        					sIntMailCode:
'###                        					sEmail:
'###                        					sTelePhone:
'###                        					sLocale:
'###                        					sTimeZone:
'###                        					sUserImage:
'###                        
'###	Return Value     :	True/False 
'###
'###		Prequisite		:	Organization Prespective is Open.
'###
'###    Function Calls  :   Fn_WriteLogFile
'###
'###    HISTORY         :  		 AUTHOR          					  DATE        					VERSION
'###
'###    CREATED BY      :   Ketan Raje   					08/06/2010  						1.0
'###
'###    REVIWED BY      :   Harshal
'###
'###    EXAMPLE         :  Case "Create" : Call Fn_Org_PersonOperations("Create", "kstar", "SQS", "pune", "MH", "421202", "India", "sqs", "95295", "0251", "ketan@sqs.com", "9821087858", "it_IT", "Asia/Tokyo", "C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\Winter.jpg")
'###									Case "Modify" : Call Fn_Org_PersonOperations("Modify", "k2star", "SQS", "pune", "MH", "421202", "India", "sqs", "95295", "0251", "ketan@sqs.com", "9821087858", "it_IT", "Asia/Tokyo", "C:\Documents and Settings\All Users\Documents\My Pictures\Sample Pictures\Winter.jpg")	
'###									Case "Delete" : Call Fn_Org_PersonOperations("Delete", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
'################################################################################################################

Public Function Fn_Org_PersonOperations(sAction, sName, sAddress, sCity, sState, sZipCode, sCountry, sOrganization, sEmpNumber, sIntMailCode, sEmail, sTelePhone, sLocale, sTimeZone, sUserImage)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_PersonOperations"
	Dim objPerson, bReturn, iCount, iCounter
	Set objPerson = Fn_UI_ObjectCreate("Fn_Org_PersonOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
	Select Case sAction
			Case "Create","Modify"
						'Set Name
						If sName<>"" Then
							Call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"Name",sName)
						End If
						'Set Address
						 If sAddress<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"Address:",sAddress)
						End If
						'Set City
						If sCity<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"City",sCity)
						End If
						'Set State
						 If sState<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"State:",sState)
						End If
						'Set Zip Code
						 If sZipCode<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"Zip Code",sZipCode)
						End If
						'Set Country
					    If sCountry<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"Country",sCountry)
						End If
						'Set Organization
					    If sOrganization<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"Organization",sOrganization)
						End If
						'Set Employee Number
						If sEmpNumber<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"Employee Number:",sEmpNumber)
						End If
						'Set Internal Mail Code
						If sIntMailCode<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"Internal Mail Code:",sIntMailCode)
						End If
						'Set Email
						If sEmail<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"E_Mail",sEmail)
						End If
						'Set Telephone number
						If sTelePhone<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"Telephone",sTelePhone)
						End If
						'Set Local
						If sLocale<>"" Then
							If objPerson.JavaList("Locale:").Exist(2) Then
								bReturn = objPerson.JavaList("Locale:").GetROProperty("items count")			    				
								For iCount=0 to bReturn -1
									If Trim(lcase(objPerson.JavaList("Locale:").GetItem(iCount))) = Trim(lcase(sLocale)) Then							
										objPerson.JavaList("Locale:").Select sLocale
										Exit For
									End If
								Next
							Else   '[TC1123-20161122-02_12_2016-VivekA-NewDevelopment] - REG-Admin development
								objPerson.JavaEdit("CustomJavaEdit").SetTOProperty "attached text","Locale"
								If NOT objPerson.JavaEdit("CustomJavaEdit").Exist Then
									objPerson.JavaEdit("CustomJavaEdit").SetTOProperty "attached text","Locale:"
								End If
								'objPerson.JavaEdit("CustomJavaEdit").Type sLocale
								Call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"CustomJavaEdit",sLocale)
								Wait 1
								Set objShell = CreateObject("Wscript.Shell")
								objShell.SendKeys "{ENTER}"
								Set objShell = Nothing
								Wait 1
							End If  '----------------------------------------------------
						End If
						'Set Time Zone
						If sTimeZone<>"" Then
							If objPerson.JavaList("Time Zone").Exist(1) Then
								bReturn = objPerson.JavaList("Time Zone").GetROProperty("items count")			    				
								For iCount=0 to bReturn -1
									If Trim(lcase(objPerson.JavaList("Time Zone").GetItem(iCount))) = Trim(lcase(sTimeZone)) Then							
										objPerson.JavaList("Time Zone").Select sTimeZone
										Exit For
									End If
								Next
							Else  '[TC1123-20161122-02_12_2016-VivekA-NewDevelopment] - REG-Admin development
								objPerson.JavaEdit("CustomJavaEdit").SetTOProperty "attached text","Time Zone"
								'objPerson.JavaEdit("CustomJavaEdit").Type sLocale
								Call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"CustomJavaEdit",sTimeZone)
								Wait 1
								Set objShell = CreateObject("Wscript.Shell")
								objShell.SendKeys "{ENTER}"
								Set objShell = Nothing
								Wait 1
							End If  '----------------------------------------------------
						End If
						'Set User Image
						If sUserImage<>"" Then
							call Fn_Edit_Box("Fn_Org_PersonOperations",objPerson,"User Image",sUserImage)
						End If
						'To Click on Create or Modify Button
						If sAction="Create" Then			
								Call Fn_Button_Click("Fn_Org_PersonOperations",objPerson,"Create")	
						ElseIf sAction="Modify" Then			
								Call Fn_Button_Click("Fn_Org_PersonOperations",objPerson,"Modify")	
						End If
			Case "Delete"
					  'Click on Delete button.
					  Call Fn_Button_Click("Fn_Org_PersonOperations", objPerson, "Delete")
					  'Click on yes button to delete the site.
					  For iCount = 0 to 0
					   JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete Confirmation"		'Modified code to handle Msgbox in multiple hierarchy.
					   If JavaDialog("DeleteConfirmation").Exist Then
						JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
						Exit For
					   End If
					   JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").SetTOProperty "title", "Delete Confirmation"
					   If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").Exist Then
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").JavaButton("Yes").Click
						Exit For
					   End If
					  Next
					  Call Fn_ReadyStatusSync(2)
			'[TC1123-20161122-02_12_2016-VivekA-NewDevelopment] - REG-Admin development
			Case "Verify"
					iCount=0
					iCounter=0
					'Verify Name 
					If sName<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("Name").GetROProperty("value"))=Trim(sName) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Person Name matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Person Name does not matches with actual value")
						End If
					End If
					'Verify Address
					If sAddress<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("Address:").GetROProperty("value"))=Trim(sAddress) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Address matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Address does not matches with actual value")
						End If
					End If
					'Verify City
					If sCity<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("City").GetROProperty("value"))=Trim(sCity) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "City matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "City does not matches with actual value")
						End If
					End If
					'Verify State
					If sState<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("State:").GetROProperty("value"))=Trim(sState) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "State matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "State does not matches with actual value")
						End If
					End If
					'Verify Zip Code
					If sZipCode<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("Zip Code").GetROProperty("value"))=Trim(sZipCode) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Zip Code matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Zip Code does not matches with actual value")
						End If
					End If
					'Verify Country
					If sCountry<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("Country").GetROProperty("value"))=Trim(sCountry) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Country matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Country does not matches with actual value")
						End If
					End If
					'Verify Organization
					If sOrganization<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("Organization").GetROProperty("value"))=Trim(sOrganization) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Organization matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Organization does not matches with actual value")
						End If
					End If
					'Verify Employee Number
					If sEmpNumber<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("Employee Number:").GetROProperty("value"))=Trim(sEmpNumber) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Organization matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Organization does not matches with actual value")
						End If
					End If
					'Verify Internal Mail Code
					If sIntMailCode<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("Internal Mail Code:").GetROProperty("value"))=Trim(sIntMailCode) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Internal Mail Code matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Internal Mail Code does not matches with actual value")
						End If
					End If
					'Verify Email Address
					If sEmail<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("E_Mail").GetROProperty("value"))=Trim(sEmail) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Email Address matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Email Address does not matches with actual value")
						End If
					End If
					'Verify Phone Number
					If sTelePhone<>"" Then
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("Telephone").GetROProperty("value"))=Trim(sTelePhone) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Phone Number matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Phone Number does not matches with actual value")
						End If
					End If
					'Verify Locale
					If sLocale<>"" Then
						objPerson.JavaEdit("CustomJavaEdit").SetTOProperty "attached text","Locale"
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("CustomJavaEdit").GetROProperty("value"))=Trim(sLocale) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Locale matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Locale does not matches with actual value")
						End If
					End If
					
					'Verify Time Zone
					If sTimeZone<>"" Then
						objPerson.JavaEdit("CustomJavaEdit").SetTOProperty "attached text","Time Zone"
						iCount = iCount + 1
						If  Trim(objPerson.JavaEdit("CustomJavaEdit").GetROProperty("value"))=Trim(sTimeZone) Then
						iCounter = iCounter + 1
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Time Zone matches with actual value")
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Time Zone does not matches with actual value")
						End If
					End If
					If iCount = iCounter Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All values verified successfully")
						Fn_Org_PersonOperations = TRUE
						Set objUser = nothing 
						Exit Function
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "All values does not match successfully")
						Fn_Org_PersonOperations = FALSE
						Set objUser = nothing 
						Exit Function
					End If	
			'-------------------------------------------------------------------
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_PersonOperations function failed")
						Fn_Org_PersonOperations = FALSE
						Exit Function						
		End Select
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Person [" + sName + "] of function Fn_Org_PersonOperations")
		Fn_Org_PersonOperations = TRUE
Set objPerson=Nothing
End Function

'#################################################################################################################
'###    FUNCTION NAME   :     Fn_Org_VolumeOperations()  
'###
'###    DESCRIPTION     :   Create/Modify/Delete/GetDetails of Volumes
'###
'###    PARAMETERS      :  sAction:Create/Modify/Delete/AddExisting
'###											sVolumeName:
'###											sNodeName:
'###											sMachineType:
'###											sPathName: at a time only Unix or Windows is required.
'###											sFSCPathName:
'###											sIDType:
'###											sID:
'###											sFMSConfig:
'###											sAccessors:
'###                        
'###	Return Value     :	True/False/Array in Case: Get Details
'###
'###		Prequisite		:	Organization Prespective is Open.
'###
'###    Function Calls  :   Fn_WriteLogFile(), Fn_Edit_Box(), Fn_UI_JavaRadioButton_SetON(), Fn_Button_Click()
'###
'###    HISTORY         :  		 AUTHOR          					  DATE        					VERSION
'###
'###    CREATED BY      :   Ketan Raje   					09/06/2010  						1.0
'###
'###    REVIWED BY      :   Harshal
'###
'###    EXAMPLE         :  Case "Create" : Call Fn_Org_VolumeOperations("Create", "VolumeA", "pnv6s108", "Windows", "C:\apps\Siemens\VolumeA", "", "FSC", "FSC_pnv6s108_autoadmin", "", "")
'###									Case "Modify" : Call Fn_Org_VolumeOperations("Modify", "VolumeB", "pnv6s108", "Windows", "C:\apps\Siemens\VolumeB", "", "FSC", "FSC_pnv6s108_autoadmin", "", "Revoke:auto_grp1_163418")
'###									Case "Delete" : Call Fn_Org_VolumeOperations("Delete", "", "", "", "", "", "", "", "", "")
'###									Case "GetDetails" : ArrayVariable = Fn_Org_VolumeOperations("GetDetails", "", "", "", "", "", "", "", "", "")		
'################################################################################################################

Public Function Fn_Org_VolumeOperations(sAction, sVolumeName, sNodeName, sMachineType, sPathName, sFSCPathName, sIDType, sID, sFMSConfig, sAccessors)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_VolumeOperations"
	Dim objVolume, aVoldetails, aAccessors, iRowData, iCount, iCounter, bReturn
	ReDim aVoldetails(6)
	Set objVolume = Fn_UI_ObjectCreate("Fn_Org_VolumeOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
	Select Case sAction
			Case "Create","Modify"
						'Set Volume Name
						If sVolumeName<>"" Then
							Call Fn_Edit_Box("Fn_Org_VolumeOperations",objVolume,"Volume Name:",sVolumeName)
						End If
						'Set Node Name
						 If sNodeName<>"" Then
							call Fn_Edit_Box("Fn_Org_VolumeOperations",objVolume,"Node Name:",sNodeName)
						End If
						'Select MachineType Radio button
						If sMachineType<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_Org_VolumeOperations",objVolume, sMachineType)
						End If
						'Set PathName
						If sMachineType="Unix" Then
							'Set Unix Path
							If sPathName<>"" Then
								Call Fn_Edit_Box("Fn_Org_VolumeOperations",objVolume,"UNIX Path Name:",sPathName)
							End If
						Else
							'Set Windows Path
							If sPathName<>"" Then
								Call Fn_Edit_Box("Fn_Org_VolumeOperations",objVolume,"Windows Path Name:",sPathName)
							End If
						End If
						'Set FSC PathName
						 If sFSCPathName<>"" Then
                             If objVolume.JavaEdit("FSC Path Name:").Exist(5) Then
								call Fn_Edit_Box("Fn_Org_VolumeOperations",objVolume,"FSC Path Name:",sFSCPathName)
							End If
						End If
						'Set IDType Radio button
						If sIDType<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_Org_VolumeOperations",objVolume, sIDType)
						End If
						'Set ID
						 If sID<>"" Then
							call Fn_Edit_Box("Fn_Org_VolumeOperations",objVolume,"ID:",sID)
						End If
						' "In Modify Case" To Click on Grant/Revoke to add/remove Accessors
						If sAccessors<>"" Then
								bReturn = objVolume.JavaList("Accessors:").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aAccessors = split(sAccessors, ":",-1,1)
								iCount = Ubound(aAccessors)
								For iRowData=1 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objVolume.JavaList("Accessors:").GetItem(iCounter))) = Trim(lcase(aAccessors(iRowData))) then
											objVolume.JavaList("Accessors:").Select aAccessors(iRowData)
											'Click on Add Button
											Call Fn_Button_Click("Fn_Org_VolumeOperations", objVolume, aAccessors(0))
											Exit For 
										End If
									Next
								Next
						End If
						'To Click on Create or Modify Button
						If sAction="Create" Then			
								Call Fn_Button_Click("Fn_Org_VolumeOperations",objVolume,"Create")	
						ElseIf sAction="Modify" Then			
								Call Fn_Button_Click("Fn_Org_VolumeOperations",objVolume,"Modify")
								'Set TOProperty of Window
								'Call Fn_UI_Object_SetTOProperty("Fn_Org_VolumeOperations",JavaDialog("DeleteConfirmation"),"title","Move volume")
								JavaDialog("DeleteConfirmation").SetTOProperty "title","Move volume"
								If JavaDialog("DeleteConfirmation").Exist Then
									'Click on yes Button
									Call Fn_Button_Click("Fn_Org_VolumeOperations", JavaDialog("DeleteConfirmation"), "Yes")
								End If
						End If
			Case "Delete"
						'Click on create button.
						Call Fn_Button_Click("Fn_Org_VolumeOperations", objVolume, "Delete")
						'Click on yes button of delete dialog
						'Call Fn_Button_Click("Fn_Org_VolumeOperations", JavaDialog("DeleteConfirmation"), "Yes")
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").SetTOProperty "title", "Delete Confirmation"
						Call Fn_Button_Click("Fn_Org_VolumeOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog"), "Yes")
			Case "GetDetails"
						'Get  Volume Name
						aVoldetails(0) = Fn_Edit_Box_GetValue("Fn_Org_VolumeOperations",objVolume,"Volume Name:")
						'Get  Node Name
						aVoldetails(1) = Fn_Edit_Box_GetValue("Fn_Org_VolumeOperations",objVolume,"Node Name:")
						'Check Machine Type
						If Fn_UI_Object_GetROProperty("Fn_Org_VolumeOperations",objVolume.JavaRadioButton("Unix"), "value")="1" Then
							aVoldetails(2) = "Unix"
						Else
							aVoldetails(2) = "Windows"
						End If
						If aVoldetails(2)="Unix" Then
							'Get  Unix PathName
							aVoldetails(3) = Fn_Edit_Box_GetValue("Fn_Org_VolumeOperations",objVolume,"UNIX Path Name:")
						Else
							'Get  Windows PathName
							aVoldetails(3) = Fn_Edit_Box_GetValue("Fn_Org_VolumeOperations",objVolume,"Windows Path Name:")							
						End If
						'Get FSC PathName
                        If objVolume.JavaEdit("FSC Path Name:").Exist(5) Then
							aVoldetails(4) = Fn_Edit_Box_GetValue("Fn_Org_VolumeOperations",objVolume,"FSC Path Name:")							
						End If
						'Check ID Type
						If Fn_UI_Object_GetROProperty("Fn_Org_VolumeOperations",objVolume.JavaRadioButton("FSC"), "value")="1" Then
							aVoldetails(5) = "FSC"
						ElseIf Fn_UI_Object_GetROProperty("Fn_Org_VolumeOperations",objVolume.JavaRadioButton("Filestore Group"), "value")="1" Then
							aVoldetails(5) = "Filestore Group"
						Else
							aVoldetails(5) = "Load Balancer"
						End If
						'Get ID
						aVoldetails(6) = Fn_Edit_Box_GetValue("Fn_Org_VolumeOperations",objVolume,"ID:")
						Fn_Org_VolumeOperations = aVoldetails
						Set objPerson=Nothing
						Exit Function												
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_VolumeOperations function failed")
						Fn_Org_VolumeOperations = FALSE
						Set objPerson=Nothing
						Exit Function						
		End Select
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Volume [" + sVolumeName + "] of function Fn_Org_VolumeOperations")
		Fn_Org_VolumeOperations = TRUE
Set objPerson=Nothing
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Org_GroupMemberSettings(sAction,bGroupAdmin,Groupmenberstatus,bDefaultRole,bExtManaged)
'###
'###    DESCRIPTION        :   Create,Modify,Delete Groups
'###
'###    PARAMETERS      :   1. sAction: Edit / Verify
'###											 2.	bGroupAdmin:
'###											3.	Groupmenberstatus:
'###											4.	bDefaultRole:
'###											5.	bExtManaged:
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           11/06/2010         1.0
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "Edit" : Call Fn_Org_GroupMemberSettings("Edit","ON","Active","OFF","ON")
'###										 Case "Verify" :
'#############################################################################################################

Public Function Fn_Org_GroupMemberSettings(sAction,bGroupAdmin,Groupmenberstatus,bDefaultRole,bExtManaged)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_GroupMemberSettings"
	Dim objGroup
	Set objGroup = Fn_UI_ObjectCreate("Fn_Org_GroupMemberSettings", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Edit"
						'Set Group Admin Checkbox				
						If bGroupAdmin<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_GroupMemberSettings", objGroup, "Group Administrator_GMS", bGroupAdmin)
						End If
						'Select Group Member Status.
						If Groupmenberstatus<>"" Then
								If Groupmenberstatus = "Active" Then
										 'Set Active ON 
										 Call Fn_UI_JavaRadioButton_SetON("Fn_Org_GroupMemberSettings",objGroup, "Active_GMS")										 
								ElseIf Groupmenberstatus = "InActive" Then
										 'Set InActive ON
										 Call Fn_UI_JavaRadioButton_SetON("Fn_Org_GroupMemberSettings",objGroup, "Inactive_GMS")										 
								End If
						End If
						'Set Default Role Checkbox				
						If bDefaultRole<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_GroupMemberSettings", objGroup, "DefaultRole_GMS", bDefaultRole)
						End If
						'Set Externally Managed Checkbox				
						If bExtManaged<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_GroupMemberSettings", objGroup, "Externally Managed_GMS", bExtManaged)
						End If			
						'Click on Modify button
						Call Fn_Button_Click("Fn_Org_GroupMemberSettings",objGroup,"Modify")									

			Case "Verify"
		
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_GroupMemberSettings function failed")
						Fn_Org_GroupMemberSettings = FALSE
						Exit Function						
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Group [" + sGrpName + "] of function Fn_Org_GroupMemberSettings")
	Fn_Org_GroupMemberSettings = TRUE
	Set objGroup = nothing 
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Reg_EditorOperation()
'###
'###    DESCRIPTION        :   Operations in Registry editor
'###
'###    PARAMETERS      :  sAction,sFilePath,sKey,sValue,sSearchText
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Harshal
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          :  Call Fn_Reg_EditorOperation("LoadFile","C:\auto_tc\TC83\pnv6s112\2010051100\win\rac\plugins\configuration_8000.3.0\portal.properties","","","")
											'Call Fn_Reg_EditorOperation("Add","","Harshal","Hello","")
											'Call Fn_Reg_EditorOperation("VerifyEditor","","","","Harshal=Hello")
											'Call Fn_Reg_EditorOperation("Search","","Harshal","Hello","Harshal")
											'Call Fn_Reg_EditorOperation("Delete","","Harshal","Hello","Harshal")
											'Call Fn_Reg_EditorOperation1("Modify","","Harshal:Pranav","Hello:Hi","")
'#############################################################################################################
Function Fn_Reg_EditorOperation(sAction,sFilePath,sKey,sValue,sSearchText)
	GBL_FAILED_FUNCTION_NAME="Fn_Reg_EditorOperation"
' Variable Delecration
   Dim objReg,objOpen,iRowCount,iCount,sText
'Defining Objects
   Set objReg = Fn_UI_ObjectCreate("Fn_Reg_EditorOperation", JavaWindow("Registry Editor - Teamcenter").JavaWindow("RegEditApplet"))   
'Select Case Starts
Select Case sAction
'Loading the File Portal.Properties
		Case "LoadFile"
				If Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTab("RegEditTab"), "value") <>"Properties" Then					
					objReg.JavaTab("RegEditTab").Select "Properties"
				End if
				Call Fn_MenuOperation("Select","File:Open...")
				Set objOpen = Fn_UI_ObjectCreate("Fn_Reg_EditorOperation", JavaWindow("Registry Editor - Teamcenter").Dialog("OpenWindow"))
				'If objOpen.Exist(5)  Then
				If Fn_SISW_UI_Object_Operations("Fn_Reg_EditorOperation","Exist",objOpen,SISW_MICROLESS_TIMEOUT) Then
						objOpen.WinEdit("File name").Set sFilePath
						'Call Fn_UI_WinButton_Click("Fn_Reg_EditorOperation", objOpen, "Open","1","1","LEFT")
						JavaWindow("Registry Editor - Teamcenter").Dialog("OpenWindow").WinButton("Open").Click
						Fn_Reg_EditorOperation = True
				Else
						Fn_Reg_EditorOperation = False
						Exit Function
				End If
'Adding a Row To Table
	Case "Add"
			If Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTab("RegEditTab"), "value") <>"Properties" Then					
					objReg.JavaTab("RegEditTab").Select "Properties"
            End If
			Call Fn_Button_Click("Fn_Reg_EditorOperation",objReg,"Add")
			'objReg.JavaButton("Add").Click
			iRowCount = Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTable("RegistryEditorTable"), "rows")
			'objReg.JavaTable("RegistryEditorTable").GetROProperty("rows")
			For iCount = 0 to (iRowCount -1)				
				If  Fn_UI_JavaTable_GetCellData("Fn_Reg_EditorOperation", objReg, "RegistryEditorTable",iCount,"0") = "NewKey" Then
					Call Fn_UI_JavaTable_SetCellData("Fn_Reg_EditorOperation",objReg,"RegistryEditorTable",iCount,"0",sKey)
					'objReg.JavaTable("RegistryEditorTable").SetCellData iCount,0,sKey
					Call Fn_UI_JavaTable_SetCellData("Fn_Reg_EditorOperation",objReg,"RegistryEditorTable",iCount,"1",sValue)
					'objReg.JavaTable("RegistryEditorTable").SetCellData iCount,1,sValue
					Exit for
				End If
			Next
			Call Fn_RegEdit_ToolbatButtonClick("Save modified file. (Ctrl+S)")
			Fn_Reg_EditorOperation = True
'Verifying the Text in Editor Tab
	Case "VerifyEditor"
			If Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTab("RegEditTab"), "value") <>"Editor" Then
					objReg.JavaTab("RegEditTab").Select "Editor"
            End if
            sText = Fn_Edit_Box_GetValue("",objReg,"RegEditBox")
			'sText = objReg.JavaEdit("RegEditBox").Object.getText
			If  instr(1,sText,sSearchText,1)<>0 Then
					Fn_Reg_EditorOperation = True
			Else
					Fn_Reg_EditorOperation = False
			End If
'Searching the Key value in Properties tab
	Case "Search"
		If Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTab("RegEditTab"), "value") <>"Properties" Then
					objReg.JavaTab("RegEditTab").Select "Properties"
        End If
		Call Fn_Edit_Box("Fn_Reg_EditorOperation",objReg,"RegEditBox",sSearchText)
		'objReg.JavaEdit("RegEditBox").Set sSearchText
		Call Fn_Button_Click("Fn_Reg_EditorOperation",objReg,"Search")
		'objReg.JavaButton("Search").Click
		If objReg.JavaTable("RegistryEditorTable").Object.getSelectedRow<> "-1" Then
				iIndex = cint(objReg.JavaTable("RegistryEditorTable").Object.getSelectedRow())
				If  Fn_UI_JavaTable_GetCellData("Fn_Reg_EditorOperation", objReg, "RegistryEditorTable",iIndex, 0) = skey AND Fn_UI_JavaTable_GetCellData("Fn_Reg_EditorOperation", objReg, "RegistryEditorTable",iIndex, 1 ) = sValue Then
					Fn_Reg_EditorOperation =True
				Else
					Fn_Reg_EditorOperation = False
				End If
		Else
				Fn_Reg_EditorOperation = False
		End If
'Delete a Row In Properties Tab
	Case "Delete"
		If Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTab("RegEditTab"), "value") <>"Properties" Then
					objReg.JavaTab("RegEditTab").Select "Properties"
        End If		
		iRowCount = Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTable("RegistryEditorTable"), "rows")
			For iCount = 0 to (iRowCount -1)				
				If  Fn_UI_JavaTable_GetCellData("Fn_Reg_EditorOperation", objReg, "RegistryEditorTable",iCount,"0") = sKey Then
					Call Fn_UI_JavaTable_SelectRow("Fn_Reg_EditorOperation", objReg, "RegistryEditorTable",iCount)
					'objReg.JavaTable("RegistryEditorTable").SelectRow(iCount)
					Call Fn_Button_Click("Fn_Reg_EditorOperation",objReg,"Remove")
					'objReg.JavaButton("Remove").Click
					Exit for
				End If
			Next
			Call Fn_RegEdit_ToolbatButtonClick("Save modified file. (Ctrl+S)")
			Fn_Reg_EditorOperation = True
'Modifing a keyValue In Table
	Case "Modify"
		aKey = Split(sKey,":",-1)
		aValue = Split(sValue,":",-1)
			If Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTab("RegEditTab"), "value") <>"Properties" Then					
					objReg.JavaTab("RegEditTab").Select "Properties"
            End If
			iRowCount = Fn_UI_Object_GetROProperty("Fn_Reg_EditorOperation",objReg.JavaTable("RegistryEditorTable"), "rows")
			'objReg.JavaTable("RegistryEditorTable").GetROProperty("rows")
			For iCount = 0 to (iRowCount -1)				
				If  Fn_UI_JavaTable_GetCellData("Fn_Reg_EditorOperation", objReg, "RegistryEditorTable",iCount,"0") = aKey(0) Then
					Call Fn_UI_JavaTable_SetCellData("Fn_Reg_EditorOperation",objReg,"RegistryEditorTable",iCount,"0",aKey(1))
					'objReg.JavaTable("RegistryEditorTable").SetCellData iCount,0,sKey
					Call Fn_UI_JavaTable_SetCellData("Fn_Reg_EditorOperation",objReg,"RegistryEditorTable",iCount,"1",aValue(1))
					'objReg.JavaTable("RegistryEditorTable").SetCellData iCount,1,sValue
					Exit for
				End If
			Next
			Call Fn_RegEdit_ToolbatButtonClick("Save modified file. (Ctrl+S)")
			Fn_Reg_EditorOperation = True
End Select
Set objReg = Nothing
Set objOpen = Nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_RegEdit_ToolbatButtonClick()
'###
'###    DESCRIPTION        :   Toobar Button click for Registry Editor
'###
'###    PARAMETERS      :  SButtonName: ToolTip of the Button to be Click
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Harshal
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Call Fn_RegEdit_ToolbatButtonClick("Save modified file. (Ctrl+S)")
'#############################################################################################################
Public Function Fn_RegEdit_ToolbatButtonClick(sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_RegEdit_ToolbatButtonClick"
	Dim ObjDesc, ArrLists, iToolCnt, iCounter, sContents

	If JavaWindow("Registry Editor - Teamcenter").Exist(iTimeOut) Then
		'Create Toolbar object
		Set ObjDesc = Description.Create() 
		ObjDesc("to_class").Value = "JavaToolbar" 
		ObjDesc("enabled").Value = 1
	
		JavaWindow("Registry Editor - Teamcenter").Activate
		'following statement is used to change the focus to the main window
		JavaWindow("Registry Editor - Teamcenter").Type(micAlt)
	
		'Get the total of Toolbar objects
		Set ArrLists =JavaWindow("Registry Editor - Teamcenter").ChildObjects(ObjDesc)
		iToolCnt = JavaWindow("Registry Editor - Teamcenter").ChildObjects(ObjDesc).count
	
		For iCounter = 0 to iToolCnt-1
			sContents = ArrLists(iCounter).GetContent()
			If instr(sContents, sButtonName) > 0 Then
				ArrLists(iCounter).Press sButtonName
				'Reporter.ReportEvent micPass, "ToolbarButtonClick", "Successfully Clicked on Toolbat Button [" + sButtonName + "]"
				'Call Fn_WriteLogFile("Fn_ToolbatButtonClick", 3, Err.Number,"PASS: Successfully Clicked on Toolbat Button [" + sButtonName + "]")
				Fn_RegEdit_ToolbatButtonClick = TRUE
				Exit For
			End If
		Next
	
		If iCounter = iToolCnt Then
			'Reporter.ReportEvent micFail, "ToolbarButtonClick", "Failed to Click on Toolbat Button [" + sButtonName + "]"
			'Call Fn_WriteLogFile("Fn_ToolbatButtonClick", 1, Err.Number,"FAIL: Failed to Click on Toolbat Button [" + sButtonName + "]")
			Fn_RegEdit_ToolbatButtonClick = FALSE
		End If
	
		Set ObjDesc = Nothing
		Set ArrLists = Nothing
	Else
		Fn_RegEdit_ToolbatButtonClick = FALSE
		'Reporter.ReportEvent micFail, "ToolbarButtonClick","Teamcenter Default Window not Accessibles"
		'Call Fn_WriteLogFile("Fn_ToolbatButtonClick", 1, Err.Number,"FAIL: Teamcenter Default Window not Accessibles")
	End If
End Function
''****************************************************************************************************************************************************************************************
'Function Name		:				Fn_CommSup_MenuOperation()

'Description			 :		 		 Actions performed in this function are:
'													1. Select a particular Menu and Click on Hide / Show.
'													2. Get the list of Child nodes in a Parent node.

'Parameters			   :	 			1. sAction: Action to be performed.
'													2. sNodeName: Application Name.

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Command Suppression Prespective is Open.

'Examples				:				Case "Hide" : Call Fn_CommSup_MenuOperation("Show","File") / Call Fn_CommSup_MenuOperation("Hide","Edit:Copy")
'													Case "Show" : Call Fn_CommSup_MenuOperation("Show","File[HIDDEN]") / Call Fn_CommSup_MenuOperation("Show","Edit:Copy[HIDDEN]")
'													Case "GetList" : Call Fn_CommSup_MenuOperation("Getlist","Tools")
'													Case "Exist" : Call Fn_CommSup_MenuOperation("Exist","Tools")		
'													Case "GetParentMenus" : Call Fn_CommSup_MenuOperation("GetParentMenus","")
'													Case "HideByPopupMenu" : Call Fn_CommSup_MenuOperation("HideByPopupMenu","File")
'													Case "ShowByPopupMenu" : Call Fn_CommSup_MenuOperation("ShowByPopupMenu","File[HIDDEN]")  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			17/06/2010			              1.0										Created								Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   			02/05/2013			              1.1										Added case : HideByPopupMenu,ShowByPopupMenu
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   			28/05/2013			              1.2										Added case : HidePopupMenuExist,ShowPopupMenuExist			Veena G
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_CommSup_MenuOperation(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_CommSup_MenuOperation"
	Dim objJavaWindowCom, aNodeName, iIndex, iTotal, sReturn, iLen, iCounter, arr, bFlag, jCounter,sTempNodeName
	Dim ObjDesc, ArrLists, iToolCnt, sContents
	Set objJavaWindowCom = Fn_UI_ObjectCreate( "Fn_CommSup_MenuOperation",JavaWindow("DefaultWindow"))
	sNodeName  = Replace(sNodeName,"[HIDDEN]","[Hidden by Application]")
	
	Select Case sAction
		'----------------------------------------------------------------------- Show / Hide Menu options -------------------------------------------------------------------------
		Case "Hide"
			  If sNodeName <> "" Then
					If Instr(1, sNodeName, ":") = 0 Then
						'Select the node
						bGblFuncRetVal = Fn_JavaTree_NodeIndexExt("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",sNodeName & "[Hidden by Application]","","")
						If bGblFuncRetVal = -1 Then
							Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",sNodeName)	
						End If						
					Else
						aNodeName = Split(sNodeName,":",-1,1)
						'Expand the parent node
						Call Fn_UI_JavaTree_Expand("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",aNodeName(0))
						'Select the node
						bGblFuncRetVal = Fn_JavaTree_NodeIndexExt("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree", sNodeName & "[Hidden by Application]","","")
						If bGblFuncRetVal = -1 Then
							Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",sNodeName)	
						End If
					End If
			  Else
				bGblFuncRetVal = -1			
			  End If		
					
					If bGblFuncRetVal = -1 Then
						'Click Hide button
						'Call Fn_UI_JavaToolbar_Press("Fn_CommSup_MenuOperation", objJavaWindowCom, "ComSupToolBar", "&Hide")
						Set ObjDesc = Description.Create() 
						ObjDesc("Class Name").Value = "JavaToolbar" 
						ObjDesc("enabled").Value = 1
					
						'JavaWindow("MyTeamcenter").Activate
						JavaWindow("DefaultWindow").Maximize
						'following statement is used to change the focus to the main window
						JavaWindow("DefaultWindow").Type(micAlt)
					
						'Get the total of Toolbar objects
						Set ArrLists =JavaWindow("DefaultWindow").ChildObjects(ObjDesc)
						iToolCnt = JavaWindow("DefaultWindow").ChildObjects(ObjDesc).count
					
						For iCounter = 0 to iToolCnt-1
							sContents = ArrLists(iCounter).GetContent()
							If instr(sContents, "&Hide") > 0 Then
								ArrLists(iCounter).Press "&Hide"
								Exit For
							End If
						Next
					
						If iCounter = iToolCnt Then
							'Reporter.ReportEvent micFail, "ToolbarButtonClick", "Failed to Click on Toolbat Button [" + sButtonName + "]"
							'Call Fn_WriteLogFile("Fn_ToolbatButtonClick", 1, Err.Number,"FAIL: Failed to Click on Toolbat Button [" + sButtonName + "]")
							Fn_CommSup_MenuOperation = FALSE
						End If
					
						Set ObjDesc = Nothing
						Set ArrLists = Nothing
						Call  Fn_MenuOperation("Select", "File:Save")
					End if
					Fn_CommSup_MenuOperation = TRUE
					
		Case "Show"
					sNodeName  = Replace(sNodeName,"[HIDDEN]","[Hidden by Application]")
					If Instr(1, sNodeName, ":") = 0 Then
						'Select the node
						bGblFuncRetVal = Fn_JavaTree_NodeIndexExt("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree", sNodeName,"","")
						If bGblFuncRetVal <> -1 Then
							Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",sNodeName)
						End if
					Else
						aNodeName = Split(sNodeName,":",-1,1)
						'Expand the parent node
						For i = 0 To (uBound(aNodeName)-1)
							Select Case i
							Case "0"
							Call Fn_UI_JavaTree_Expand("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",aNodeName(0))
							
							Case "1"
							Call Fn_UI_JavaTree_Expand("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",aNodeName(0)+":"+aNodeName(1))
						End Select
						Next
						bGblFuncRetVal = Fn_JavaTree_NodeIndexExt("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",sNodeName,"","")
						If bGblFuncRetVal <> -1 Then
							'Select the node
							Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",sNodeName)
						End if
					End If
					'Select a ToolBar buttnon.
					If bGblFuncRetVal <> -1 Then
						'Click Show button
						'Call Fn_UI_JavaToolbar_Press("Fn_CommSup_MenuOperation", objJavaWindowCom, "ComSupToolBar", "&Show")	
						Set ObjDesc = Description.Create() 
						ObjDesc("Class Name").Value = "JavaToolbar" 
						ObjDesc("enabled").Value = 1
					
						'JavaWindow("MyTeamcenter").Activate
						JavaWindow("DefaultWindow").Maximize
						'following statement is used to change the focus to the main window
						JavaWindow("DefaultWindow").Type(micAlt)
					
						'Get the total of Toolbar objects
						Set ArrLists =JavaWindow("DefaultWindow").ChildObjects(ObjDesc)
						iToolCnt = JavaWindow("DefaultWindow").ChildObjects(ObjDesc).count
					
						For iCounter = 0 to iToolCnt-1
							sContents = ArrLists(iCounter).GetContent()
							If instr(sContents, "&Show") > 0 Then
								ArrLists(iCounter).Press "&Show"
								Exit For
							End If
						Next
					
						If iCounter = iToolCnt Then
							'Reporter.ReportEvent micFail, "ToolbarButtonClick", "Failed to Click on Toolbat Button [" + sButtonName + "]"
							'Call Fn_WriteLogFile("Fn_ToolbatButtonClick", 1, Err.Number,"FAIL: Failed to Click on Toolbat Button [" + sButtonName + "]")
							Fn_CommSup_MenuOperation = FALSE
						End If
					
						Set ObjDesc = Nothing
						Set ArrLists = Nothing
						'Click on Save button
						Call  Fn_MenuOperation("Select", "File:Save")
					End if
					Fn_CommSup_MenuOperation = TRUE

		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
					'-----------------------------------------------------------------------
		            aNodeName = Split(sNodeName,":",-1,1)
					sTempNodeName=aNodeName(0)
					For iCounter=0 to ubound(aNodeName)-1
						If iCounter<>0 Then
							 sTempNodeName=sTempNodeName+":"+aNodeName(iCounter)
						 End If
						 objJavaWindowCom.JavaTree("SuppressComandTree").Expand sTempNodeName
					Next
					sNodeName  = Replace(sNodeName,"[HIDDEN]","[Hidden by Application]")
					'-----------------------------------------------------------------------
					iTotal = objJavaWindowCom.JavaTree("SuppressComandTree").GetROProperty ("items count") 
					For iCounter = 0 to iTotal - 1						
						If Trim(lcase(objJavaWindowCom.JavaTree("SuppressComandTree").GetItem(iCounter))) = Trim(Lcase(sNodeName)) Then
							Fn_CommSup_MenuOperation = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cint(iCounter) = Cint(iTotal) Then
						Fn_CommSup_MenuOperation = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If

		Case "VerifyHidden"
					iIndex = 0
					If Trim(sNodeName) <> "" Then
						aNodeName = Split(sNodeName, ";", -1, 1)
						For iCounter = 0 To UBound(aNodeName)
							bFlag = False
							arr = Split(aNodeName(iCounter),":", -1, 1)
							For jCounter = 0 To UBound(arr)
								If jCounter = UBound(arr) Then
									Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree", aNodeName(iCounter))
									If InStr(1, arr(jCounter), "[Hidden by Application]", 1) > 0 Then
										iIndex = iIndex + 1
										bFlag = True
									End If
								Else
									If jCounter = 0 Then
										sReturn = arr(jCounter)
									Else
										sReturn = sReturn + ":"+arr(jCounter)
									End If
									If Trim(sReturn) <> "" Then
										Call Fn_UI_JavaTree_Expand("Fn_CommSup_MenuOperation",objJavaWindowCom,"SuppressComandTree",sReturn)
									End If
								End If
							Next
							If bFlag = True Then
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS : Node ["+CStr(aNodeName(iCounter))+"] Verified Successfully for StrikeThrough with Line [Hidden Marked].")
							Else
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL : Node ["+CStr(aNodeName(iCounter))+"] Verification Failed to Verify for StrikeThrough with Line [Hidden Marked].")
								Exit For
							End If
						Next
					End If
					If bFlag = True Then
						Fn_CommSup_MenuOperation = TRUE
					Else
						Fn_CommSup_MenuOperation = FALSE
					End If

		Case "Getlist"
					iCounter = 0
					iLen = Len(sNodeName)
					'Expand the parent node
					Call Fn_UI_JavaTree_Expand("Fn_CommSup_MenuOperation", JavaWindow("DefaultWindow"), "SuppressComandTree",sNodeName)
					'Get the Index of parent node
					iIndex = Fn_JavaTree_NodeIndex("Fn_CommSup_MenuOperation",JavaWindow("DefaultWindow"),"SuppressComandTree",sNodeName)
					'Get the count of nodes in the tree.
					iTotal = Fn_UI_Object_GetROProperty("Fn_CommSup_MenuOperation",JavaWindow("DefaultWindow").JavaTree("SuppressComandTree"), "count_all_items")
					For iCount=(iIndex+1) to iTotal
						sReturn = JavaWindow("DefaultWindow").JavaTree("SuppressComandTree").GetItem(iCount)
						iReturn = Len(sReturn)
						iReturn = iReturn - iLen
						If Instr(1, sReturn, sNodeName) > 0 Then
							If iCounter=0 Then
								sResult = mid(sReturn,(iLen+2),iReturn)+","
							Else
								sResult = sResult+mid(sReturn,(iLen+2),iReturn)+","
							End If
							'arr(iCounter) = mid(sReturn,(iLen+1),iReturn)
						Else
							Exit For
						End If
						iCounter = iCounter+1
					Next
					sResult = Mid(sResult,1,(Len(sResult)-1)) 
					arr = Split(sResult,",")
					Fn_CommSup_MenuOperation = arr

			Case "MultiHide","MultiShow"
					arr = Split(sNodeName,";", -1, 1)
					For jCounter = 0  To UBound(arr)
								If Instr(1, arr(jCounter), ":") = 0 Then
									'Select the node
									If jCounter = 0 Then
											Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",arr(jCounter))
									Else
											Call Fn_UI_JavaTree_ExtendSelect("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree", arr(jCounter))
									End If

								Else
									aNodeName = Split(arr(jCounter),":",-1,1)
									'Expand the parent node
									If jCounter = 0 Then
											Call Fn_UI_JavaTree_Expand("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",aNodeName(0))
											'Select the node
											Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",arr(jCounter))
									Else
											'Select the node
											Call Fn_UI_JavaTree_ExtendSelect("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",arr(jCounter))
									End If
								End If
					Next

					'Select a ToolBar buttnon.
					If sAction="MultiHide" Then
						'Click Hide button
						'Call Fn_UI_JavaToolbar_Press("Fn_CommSup_MenuOperation", objJavaWindowCom, "ComSupToolBar", "&Hide")
						Set ObjDesc = Description.Create() 
						ObjDesc("Class Name").Value = "JavaToolbar" 
						ObjDesc("enabled").Value = 1
					
						'JavaWindow("MyTeamcenter").Activate
						JavaWindow("DefaultWindow").Maximize
						'following statement is used to change the focus to the main window
						JavaWindow("DefaultWindow").Type(micAlt)
					
						'Get the total of Toolbar objects
						Set ArrLists =JavaWindow("DefaultWindow").ChildObjects(ObjDesc)
						iToolCnt = JavaWindow("DefaultWindow").ChildObjects(ObjDesc).count
					
						For iCounter = 0 to iToolCnt-1
							sContents = ArrLists(iCounter).GetContent()
							If instr(sContents, "&Hide") > 0 Then
								ArrLists(iCounter).Press "&Hide"
								Exit For
							End If
						Next
						
						
					ElseIf sAction="MultiShow" Then
						'Click Show button
						'Call Fn_UI_JavaToolbar_Press("Fn_CommSup_MenuOperation", objJavaWindowCom, "ComSupToolBar", "&Show")	
						Set ObjDesc = Description.Create() 
						ObjDesc("Class Name").Value = "JavaToolbar" 
						ObjDesc("enabled").Value = 1
					
						'JavaWindow("MyTeamcenter").Activate
						JavaWindow("DefaultWindow").Maximize
						'following statement is used to change the focus to the main window
						JavaWindow("DefaultWindow").Type(micAlt)
					
						'Get the total of Toolbar objects
						Set ArrLists =JavaWindow("DefaultWindow").ChildObjects(ObjDesc)
						iToolCnt = JavaWindow("DefaultWindow").ChildObjects(ObjDesc).count
					
						For iCounter = 0 to iToolCnt-1
							sContents = ArrLists(iCounter).GetContent()
							If instr(sContents, "&Show") > 0 Then
								ArrLists(iCounter).Press "&Show"
								Exit For
							End If
						Next
					End If
					'Click on Save button
'					Call Fn_ToolBarOperation("Click", "Save (Ctrl+S)", "" )
					Call  Fn_MenuOperation("Select", "File:Save")
					Fn_CommSup_MenuOperation = TRUE

		Case "GetParentMenus"
					If JavaWindow("DefaultWindow").JavaTree("SuppressComandTree").Exist Then
							'Get the count of nodes in the tree.
							iTotal = Fn_UI_Object_GetROProperty("Fn_CommSup_MenuOperation",JavaWindow("DefaultWindow").JavaTree("SuppressComandTree"), "items count")
							ReDim arr(iTotal-1)
							For iCount=0 to iTotal-1
								arr(iCount) = JavaWindow("DefaultWindow").JavaTree("SuppressComandTree").GetItem(iCount)								
							Next
							Fn_CommSup_MenuOperation = arr
					Else
							Fn_CommSup_MenuOperation = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "No Parent Menus exist in Suppress Command Tree.")
					End If					
		'----------------------------------------------------------------------- Show / Hide Popup Menu options -------------------------------------------------------------------------
		Case "HideByPopupMenu","ShowByPopupMenu","HidePopupMenuExist","ShowPopupMenuExist"
					If Instr(1, sNodeName, ":") = 0 Then
						'Select the node
						Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",sNodeName)
					Else
						aNodeName = Split(sNodeName,":",-1,1)
						'Expand the parent node
						Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",aNodeName(0))
						'Select the node
						Call Fn_JavaTree_Select("Fn_CommSup_MenuOperation", objJavaWindowCom, "SuppressComandTree",sNodeName)
					End If
					Call Fn_UI_JavaTree_OpenContextMenu("Fn_CommSup_MenuOperation",objJavaWindowCom,"SuppressComandTree",sNodeName)
					'Select a ToolBar buttnon.
					If sAction="HideByPopupMenu" or sAction="HidePopupMenuExist" Then
						'Click Hide button
						StrMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath("Hide")
					ElseIf sAction="ShowByPopupMenu" or sAction="ShowPopupMenuExist" Then
						'Click Show button
						StrMenu = JavaWindow("DefaultWindow").WinMenu("ContextMenu").BuildMenuPath("Show")
					End If
					If sAction="HideByPopupMenu" or sAction="ShowByPopupMenu" Then
						wait(3)
						JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select StrMenu
						Fn_CommSup_MenuOperation = true
						Call Fn_MenuOperation("Select", "File:Save")
					ElseIf sAction="HidePopupMenuExist" or sAction="ShowPopupMenuExist" Then
						Fn_CommSup_MenuOperation =JavaWindow("DefaultWindow").WinMenu("ContextMenu").GetItemProperty(StrMenu,"Exists")
						Call Fn_KeyBoardOperation("SendKeys", "{ESC}")
						wait 1
					End If

		Case Else
						Fn_CommSup_MenuOperation = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_CommSup_MenuOperation function failed")
						Set objJavaWindowCom = nothing
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_CommSup_MenuOperation")
	Set objJavaWindowCom = nothing
End Function

''*********************************************************		Function to perform selection of applications	***********************************************************************
'Function Name		:				Fn_CommSup_SelectApp()

'Description			 :		 		 Actions performed in this function are:
'													1. Select a particular Application.

'Parameters			   :	 			1. sAction: Action to be performed.
'													2. sNodeName: Application Name.

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Command Suppression Prespective is Open.

'Examples				:				Case "Select" : Call Fn_CommSup_SelectApp("Select","Web Browser")
'													Case "Exist" : Call Fn_CommSup_SelectApp("Exist","Web Browser")
'													Case "GetIndex" : Call Fn_CommSup_SelectApp("GetIndex","Web Browser")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			17/06/2010			              1.0										Created								Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_CommSup_SelectApp(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_CommSup_SelectApp"
	Dim objJavaWindowCom, iRows
	Set objJavaWindowCom = Fn_UI_ObjectCreate( "Fn_CommSup_SelectApp",JavaWindow("DefaultWindow"))
	Select Case sAction
		'----------------------------------------------------------------------- For selecting particular application -------------------------------------------------------------------------
		Case "Select"
                    'Get row count of table
					iRows = Fn_Table_GetRowCount("Fn_CommSup_SelectApp",objJavaWindowCom, "ApplicationTable")
					For iCount=0 to iRows-1
						If Trim(Lcase(Fn_UI_JavaTable_GetCellData("Fn_CommSup_SelectApp", objJavaWindowCom, "ApplicationTable",iCount,"0"))) = Trim(Lcase(sNodeName)) Then
							Call Fn_UI_JavaTable_SelectRow("Fn_CommSup_SelectApp", objJavaWindowCom, "ApplicationTable",iCount)
							objJavaWindowCom.JavaTable("ApplicationTable").ActivateRow iCount
							Fn_CommSup_SelectApp = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTable Selected")
							Exit For
						End If
					Next
					If Cstr(iCount) = iRows Then
						Fn_CommSup_SelectApp = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTable does not exist")
						Set objJavaWindowCom = Nothing
						Exit Function
					End If
		'----------------------------------------------------------------------- For Checking existance of a particular  Application name-------------------------------------------------------------------------
		Case "Exist"
                    'Get row count of table
					iRows = Fn_Table_GetRowCount("Fn_CommSup_SelectApp",objJavaWindowCom, "ApplicationTable")
					For iCount=0 to iRows-1
						If Trim(Lcase(Fn_UI_JavaTable_GetCellData("Fn_CommSup_SelectApp", objJavaWindowCom, "ApplicationTable",iCount,"0"))) = Trim(Lcase(sNodeName)) Then
							Fn_CommSup_SelectApp = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTable Exist")
							Exit For
						End If
					Next
					If Cstr(iCount) = iRows Then
						Fn_CommSup_SelectApp = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTable does not Exist")
						Set objJavaWindowCom = Nothing
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular Application name-------------------------------------------------------------------------
		Case "GetIndex"
                    'Get row count of table
					iRows = Fn_Table_GetRowCount("Fn_CommSup_SelectApp",objJavaWindowCom, "ApplicationTable")
					For iCount=0 to iRows-1
						If Trim(Lcase(Fn_UI_JavaTable_GetCellData("Fn_CommSup_SelectApp", objJavaWindowCom, "ApplicationTable",iCount,"0"))) = Trim(Lcase(sNodeName)) Then
							Fn_CommSup_SelectApp = iCount	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTable Exist.")
							Exit For
						End If
					Next
					If Cstr(iCount) = iRows Then
						Fn_CommSup_SelectApp = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTable does not Exist")
						Set objJavaWindowCom = Nothing
						Exit Function
					End If

		Case Else
						Fn_CommSup_SelectApp = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_CommSup_SelectApp function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Row [" + sNodeName + "] of JavaTable of function Fn_CommSup_SelectApp")
	Set objJavaWindowCom = Nothing
End Function

''*********************************************************		Function to perform action on OrganisationTree	***********************************************************************
'Function Name		:				Fn_CommSup_SelectOrg()

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. GetIndex	

'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Organization Prespective is Open.

'Examples				:				Case "Select" : Call Fn_CommSup_SelectOrg("Select","Validation Administration:Validation Administrator")
'													Case "Expand" : Call Fn_CommSup_SelectOrg("Expand","cvgroup1")
'													Case "Collapse" : Call Fn_CommSup_SelectOrg("Collapse","cvgroup1")
'													Case "Exist" : Call Fn_CommSup_SelectOrg("Exist","cvgroup1")
'													Case "GetIndex" : Call Fn_CommSup_SelectOrg("GetIndex","cvgroup1")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			17/06/2010			              1.0										Created								Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_CommSup_SelectOrg(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_CommSup_SelectOrg"
	Dim objJavaWindowOrg, objJavaTreeOrg, intNodeCount, intCount, sTreeItem
	Set objJavaWindowOrg = Fn_UI_ObjectCreate( "Fn_CommSup_SelectOrg",JavaWindow("DefaultWindow"))
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
                    Call Fn_JavaTree_Select("Fn_CommSup_SelectOrg", objJavaWindowOrg, "OrganisationTree",sNodeName)
					Fn_CommSup_SelectOrg = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_CommSup_SelectOrg",objJavaWindowOrg,"OrganisationTree",sNodeName)
					Fn_CommSup_SelectOrg = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_CommSup_SelectOrg", objJavaWindowOrg,"OrganisationTree",sNodeName)
					Fn_CommSup_SelectOrg = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeOrg = Fn_UI_ObjectCreate( "Fn_Org_OrganizationTreeOperations", objJavaWindowOrg.JavaTree("OrganisationTree"))
					intNodeCount = objJavaTreeOrg.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeOrg.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_CommSup_SelectOrg = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_CommSup_SelectOrg = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowOrg.JavaTree("OrganisationTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowOrg.JavaTree("OrganisationTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_CommSup_SelectOrg = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_CommSup_SelectOrg = FALSE
				End If

		Case Else
						Fn_CommSup_SelectOrg = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_CommSup_SelectOrg function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_CommSup_SelectOrg")
	Set objJavaWindowOrg = nothing
	Set objJavaTreeOrg = nothing
End Function
'*********************************************************		Function to Handle NullPointerException		***********************************************************************
'Function Name		:				Fn_NullPointerExceptionHandler() 

'Description			 :		 		 It Handle the Null pointer ecxception occurs  when we get it while opening the Authorization prespective in 511 build

'Return Value		   : 				True/False 

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'														Harshal													   18/06/2010			          1.0										Created									Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''*********************************************************'*********************************************************'**********************************************************************
Function Fn_NullPointerExceptionHandler()
		GBL_FAILED_FUNCTION_NAME="Fn_NullPointerExceptionHandler"
		Dim sMsg,nonDBAobj,iCount,bFlag
		bFlag = False
		Set nonDBAobj = JavaWindow("DefaultWindow").JavaWindow("ErrorAcessDialog")
		  For iCount=0 to 0
			 nonDBAobj.SetTOProperty "title","Multiple problems have occurred"
			 If nonDBAobj.Exist(5) Then
			   Call Fn_Button_Click("Fn_NullPointerExceptionHandler",nonDBAobj,"OK")
			   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "NullPointerException Handled Sucessfully")
			   bFlag = True
			   Exit For
			 End If
			 nonDBAobj.SetTOProperty "title","Error"
			 If nonDBAobj.Exist(5) Then
			   Call Fn_Button_Click("Fn_NullPointerExceptionHandler",nonDBAobj,"OK")			   
			   bFlag = True
			   Exit For
			 End If
		  Next
			If bFlag=True Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "NullPointerException Handled Sucessfully")
				Fn_NullPointerExceptionHandler = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "NullPointerException Handling failed")
				Fn_NullPointerExceptionHandler = False	
			End If
		Set nonDBAobj = nothing
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Auth_ShowHideApplication(sAction,sApplication)
'###
'###    DESCRIPTION        :   Show / Hide Applications
'###
'###    PARAMETERS      :   1. sAction: Show / Hide
'###											 2.	sApplication
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           22/06/2010         1.0
'###
'###    REVIWED BY     :   Harshal
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "Show" : Call Fn_Auth_ShowHideApplication("Show","Workflow Designer:Access Manager:Project")
'###										 Case "Hide" : Call Fn_Auth_ShowHideApplication("Hide","Workflow Designer:Access Manager:Project")
'###										Case "VerifyShowlist" : Call Fn_Auth_ShowHideApplication("VerifyShowlist","Workflow Designer:Access Manager:Project")
'#############################################################################################################
Public Function Fn_Auth_ShowHideApplication(sAction,sApplication)
	GBL_FAILED_FUNCTION_NAME="Fn_Auth_ShowHideApplication"
	Dim objAppln, iCounter, bReturn, aColname, iCount, iRowData
	Set objAppln = Fn_UI_ObjectCreate("Fn_Auth_ShowHideApplication", JavaWindow("Authorization  - Teamcenter").JavaWindow("AuthorizationApplet"))
	Fn_Auth_ShowHideApplication = FALSE
		Select Case sAction
				Case "Show"
						If sApplication<>"" Then
								bReturn = objAppln.JavaList("AvailableApp").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sApplication, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objAppln.JavaList("AvailableApp").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objAppln.JavaList("AvailableApp").Select aColname(iRowData)
											'Click on Add Button
											Call Fn_Button_Click("Fn_Auth_ShowHideApplication", objAppln, "Add")
											Fn_Auth_ShowHideApplication = TRUE
											Exit For 
										End If
									Next
								Next
						End If
						'Click on Save button
						Call Fn_Button_Click("Fn_Auth_ShowHideApplication", objAppln, "Save")
						
				Case "Hide"
						If sApplication<>"" Then
								bReturn = objAppln.JavaList("ShownApp").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sApplication, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objAppln.JavaList("ShownApp").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objAppln.JavaList("ShownApp").Select aColname(iRowData)
											'Click on Remove Button
											Call Fn_Button_Click("Fn_Auth_ShowHideApplication", objAppln, "Remove")
											Fn_Auth_ShowHideApplication = TRUE
											Exit For 
										End If
									Next
								Next
						End If
						'Click on Save button
						Call Fn_Button_Click("Fn_Auth_ShowHideApplication", objAppln, "Save")
			Case "VerifyShowlist"
				bReturn = objAppln.JavaList("ShownApp").GetROProperty("items count")
						If sApplication<>"" AND Cint(bReturn)<> 0 Then
                                'Extract the index of row at which the object exist.
								aColname = split(sApplication, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to Cint(bReturn)-1
										If Trim(lcase(objAppln.JavaList("ShownApp").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" Sucessfully found in Show appln List")
											Fn_Auth_ShowHideApplication = TRUE
                                            Exit For 
										Elseif iCounter = Cint(bReturn)-1 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aColname(iRowData) &" not found in Show appln List")
											Fn_Auth_ShowHideApplication = FALSE
                                            Exit Function											
										End If
									Next
								Next
						Else
								Fn_Auth_ShowHideApplication = FALSE
						End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Auth_ShowHideApplication function failed")
						Fn_Auth_ShowHideApplication = FALSE
						Exit Function						
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_Auth_ShowHideApplication")
    Set objAppln = nothing 	
End Function
''*********************************************************		Function to action perform on NavTree	***********************************************************************
'Function Name		:				Fn_Auth_OrganizationTreeOpration(sAction,sNodeName)

'Description			 :		 		 Actions performed in this function are:
'																	1. Node Select
'                                                                   2. Node Expand
'																	3. Node Collapse
'																	4. Node Exist
'																	5. Activate, Deactivate
'																	
'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') 

'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Authorization Prespective is Open.

'Examples				:				Case "Select" : Call Fn_Auth_OrganizationTreeOpration("Select","Organization:Engineering:Designer")
'													Case "Expand" : Call Fn_Auth_OrganizationTreeOpration("Expand","Organization:Engineering")
'													Case "Collapse" : Call Fn_Auth_OrganizationTreeOpration("Collapse","Organization:Engineering")
'													Case "Exist" : Call Fn_Auth_OrganizationTreeOpration("Exist","Organization:Engineering")
'													Case "GetIndex" : Call Fn_Auth_OrganizationTreeOpration("GetIndex","Organization:Engineering")
  												
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   			22/06/2010			              1.0										Created								Harshal
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_Auth_OrganizationTreeOpration(sAction,sNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_Auth_OrganizationTreeOpration"
	Dim objJavaWindowOrg, objJavaTreeOrg, intNodeCount, intCount, sTreeItem
	Set objJavaWindowOrg = Fn_UI_ObjectCreate( "Fn_Auth_OrganizationTreeOpration",JavaWindow("Authorization  - Teamcenter").JavaWindow("AuthorizationApplet"))
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"
                    Call Fn_JavaTree_Select("Fn_Auth_OrganizationTreeOpration", objJavaWindowOrg, "OrganizationTree",sNodeName)
					Fn_Auth_OrganizationTreeOpration = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_Auth_OrganizationTreeOpration",objJavaWindowOrg,"OrganizationTree",sNodeName)
					Fn_Auth_OrganizationTreeOpration = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_Auth_OrganizationTreeOpration", objJavaWindowOrg,"OrganizationTree",sNodeName)
					Fn_Auth_OrganizationTreeOpration = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreeOrg = Fn_UI_ObjectCreate( "Fn_Auth_OrganizationTreeOpration", objJavaWindowOrg.JavaTree("OrganizationTree"))
					intNodeCount = objJavaTreeOrg.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreeOrg.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_Auth_OrganizationTreeOpration = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_Auth_OrganizationTreeOpration = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowOrg.JavaTree("OrganizationTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowOrg.JavaTree("OrganizationTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_Auth_OrganizationTreeOpration = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_Auth_OrganizationTreeOpration = FALSE
				End If
		Case Else
				Fn_Auth_OrganizationTreeOpration = FALSE
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Auth_OrganizationTreeOpration function failed")
				Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_Auth_OrganizationTreeOpration")
	Set objJavaWindowOrg = nothing
	Set objJavaTreeOrg = nothing
End Function
'*********************************************************		Function to Check the button is enabled / disabled		**********************************************************************
'Function Name		:				Fn_Project_CheckButtonEnable(sReferencePath)

'Description			 :		 		 To Check the button is enabled or disabled.

'Parameters			   :	 			1) sAction: Action to be performed on the Tab
'													 2) sTabName: Tab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Project prespective should be displayed.

'Examples				:				Fn_Project_CheckButtonEnable(JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet").JavaButton("Add"))  

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					23-06-10			1.0																Harshal	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_Project_CheckButtonEnable(sReferencePath)
	GBL_FAILED_FUNCTION_NAME="Fn_Project_CheckButtonEnable"
        If Fn_UI_Object_GetROProperty("Fn_Project_CheckButtonEnable",sReferencePath, "enabled")=1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The button is enabled")
			   Fn_Project_CheckButtonEnable = True
        Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The button is disabled")
            Fn_Project_CheckButtonEnable = False
		End If
End Function
'*********************************************************		Function to select  the Tab into Project		**********************************************************************
'Function Name		:				Fn_Project_TabOperation

'Description			 :		 		 This Tab includes Definition Tab,AM Rule Tab
'													1)Click on Tab
'													2)Verify the Tab is open
'													3)Close the Tab

'Parameters			   :	 			1) sAction: Action to be performed on the Tab
'													 2) sTabName: Tab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Project prespective should be displayed.

'Examples				:				Fn_Project_TabOperation("Activate", "Summary")  

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					23-06-10			1.0																Harshal	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_Project_TabOperation(sAction, sTabName)  
	GBL_FAILED_FUNCTION_NAME="Fn_Project_TabOperation"
	Dim iTabCount, iTabIndex, sTabVal	
	Dim objJavaTab
	Set objJavaTab =	Fn_UI_ObjectCreate("Fn_Project_TabOperation", JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet").JavaTab("ConfigTab"))
    Select Case sAction
				Case "VerifyActivate"
								'Check Weather Requested tab is open or not( Activated or Not )
                              	If  sTabName = objJavaTab.GetROProperty("value")   Then							
										'Call Log file
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTabName+ "] : Tab is Activate(Open) ")
										Fn_Project_TabOperation = TRUE
								else							
									'Call Log file
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [" +sTabName+ "] : Tab is NOT Activate (Open)")	
									Fn_Project_TabOperation = FALSE
								End If
				Case "Activate" 
								If True =  Fn_Project_TabSet(sTabName) Then
                                    'Call Log file
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTabName+ "] : Tab is Activated ")
									Fn_Project_TabOperation = TRUE
								else
									'Call Log file
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: [" +sTabName+ "] : Tab is Not Available ")	
									Fn_Project_TabOperation = FALSE
								End if
				Case  "Close" 
									'Counting Number of Tabs 
									 iTabCount = objJavaTab.GetROProperty("items count")
									For iTabIndex = 0 to iTabCount-1
										objJavaTab.Select "#"&iTabIndex
										sTabVal=objJavaTab.GetROProperty("value") 
										If sTabVal = sTabName Then
											objJavaTab.CloseTab  sTabName
											'Call Log file
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: [" +sTabName+ "] : Tab is Successfully Closed ")
											Fn_Project_TabOperation = TRUE		
											Exit For
										End If
									Next     
									If iTabIndex = iTabCount Then
										'Call Log file
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), " Tab  [" +sAction+ "] is Not Available.")
										Fn_Project_TabOperation = FALSE
									End If
				Case Else
								'Call Log file
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Invalid ACTION  [" +sAction+ "] is Requested.")
								Fn_Project_TabOperation = FALSE
	End Select
	Set objJavaTab = Nothing
End Function
'*********************************************************		Function to select  the Tab into Project		***********************************************************************
'Function Name		:				Fn_Project_TabSet(StrTabName)

'Description			 :		 		 This function is used to select the required Tab.

'Parameters			   :	 			1.  StrTabName:Name of the Tab to be selected.
											
'Return Value		   : 				TRUE \ FALSE

'Pre-requisite			:		 		Project prespective should be displayed.

'Examples				:				 Fn_Project_TabSet("AM Rules")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje					23/06/2010																			Harshal
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_Project_TabSet(StrTabName)
	GBL_FAILED_FUNCTION_NAME="Fn_Project_TabSet"
	Dim objJavaWindow
	Set objJavaWindow = Fn_UI_ObjectCreate( "Fn_Project_TabSet", JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet"))
	   Select Case StrTabName
				'For selecting Definition Tab
			   Case "Definition" 				
						Call Fn_UI_JavaTab_Select("Fn_Project_TabSet", objJavaWindow, "ConfigTab", "Definition")
						Fn_Project_TabSet = TRUE				
			    'For selecting AM Rules Tab
				Case "AM Rules" 				
						Call Fn_UI_JavaTab_Select("Fn_Project_TabSet", objJavaWindow, "ConfigTab", "AM Rules")
						Fn_Project_TabSet = TRUE								
				'Error message If the above Tab is not selected
				Case Else 
						 Fn_Project_TabSet = FALSE
	   End Select
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Select Tab " & StrTabName & " succeeded")
		Set objJavaWindow = Nothing 
End Function

'''*********************************************************		Function to perform action on Project Tree	***********************************************************************
'' Commented this function as a part Redundant Function elimination By Sushma Pagare  23-Jan-2013 : Use Fn_PWCRules_TreeOpeartion .
''Function Name		:				Fn_Project_TreeOpeartion()
'
''Description			 :		 		 Actions performed in this function are:
''																	1. Node Select
''                                                                   2. Node Expand
''																	3. Node Collapse
''																	4. Node Exist
''																	5. GetIndex	
'
''Parameters			   :	 			1. sAction: Action to be performed
''													2. sNodeName: Fully qulified tree Path (delimiter as ':') 
''												  3. StrMenu: Context menu to be selected
'
''Return Value		   : 				TRUE / FALSE and Index Value in "GetIndex" case.
'
''Pre-requisite			:		 		Project Prespective is Open.
'
''Examples				:				Case "Select" : Call Fn_Project_TreeOpeartion("Select","In Project(  ) -> Projects:Has Class( Schedule ) -> Scheduling Objects","")
''													Case "PopupMenuSelect" : Call Fn_Project_TreeOpeartion("PopupMenuSelect","In Project(  ) -> Projects:Has Class( Schedule ) -> Scheduling Objects","Copy	Ctrl+C")
''													Case "PopupMenuExist" : Call Fn_Project_TreeOpeartion("PopupMenuExist","In Project(  ) -> Projects:Has Class( Schedule ) -> Scheduling Objects","Copy	Ctrl+C")
'  												
''History					 :		
''													Developer Name												Date						Rev. No.						Changes Done						Reviewer
''												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''													Ketan Raje										   			23/06/2010			              1.0										Created									Harshal
''												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'Function Fn_Project_TreeOpeartion(sAction,sNodeName, sMenu)
'	Dim objJavaWindowProj, objJavaTreeProj, intNodeCount, intCount, sTreeItem, aMenuList
'	Set objJavaWindowProj = Fn_UI_ObjectCreate( "Fn_Project_TreeOpeartion",JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet"))
'	Select Case sAction
'		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
'		Case "Select"					
'                    Call Fn_JavaTree_Select("Fn_Project_TreeOpeartion", objJavaWindowProj, "AMRuleTree",sNodeName)
'					Fn_Project_TreeOpeartion = TRUE
'		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
'		Case "Expand"
'                    Call Fn_UI_JavaTree_Expand("Fn_Project_TreeOpeartion",objJavaWindowProj,"AMRuleTree",sNodeName)
'					Fn_Project_TreeOpeartion = TRUE
'		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
'		Case "Collapse"
'                    Call Fn_UI_JavaTree_Collapse("Fn_Project_TreeOpeartion", objJavaWindowProj,"AMRuleTree",sNodeName)
'					Fn_Project_TreeOpeartion = TRUE
'		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
'		Case "Exist"
'				Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_Project_TreeOpeartion", objJavaWindowProj.JavaTree("AMRuleTree"))
'					intNodeCount = objJavaTreeProj.GetROProperty ("items count") 
'					For intCount = 0 to intNodeCount - 1
'						sTreeItem = objJavaTreeProj.GetItem(intCount)
'						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
'							Fn_Project_TreeOpeartion = TRUE	
'							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
'							Exit For
'						End If
'					Next
'					If Cstr(intCount) = intNodeCount Then
'						Fn_Project_TreeOpeartion = FALSE						
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
'						Exit Function
'					End If
'		'----------------------------------------------------------------------- For selecting popup menu of  a particular  node-------------------------------------------------------------------------
'		Case "PopupMenuSelect"
'			Set objJavaTreeProj = Fn_UI_ObjectCreate( "Fn_Project_TreeOpeartion", JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet").JavaTree("AMRuleTree"))
'					'Build the Popup menu to be selected
'					aMenuList = split(sMenu, ":",-1,1)
'					intCount = Ubound(aMenuList)
'					'Select node
'                    Call Fn_JavaTree_Select("Fn_Project_TreeOpeartion",objJavaWindowProj,"AMRuleTree",sNodeName)
'					'Open context menu
'					Call Fn_UI_JavaTree_OpenContextMenu("Fn_Project_TreeOpeartion",objJavaWindowProj,"AMRuleTree",sNodeName)
'					'Select Menu action
'					Select Case intCount
'						Case "0"
'							 sMenu = objJavaWindowProj.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
'						Case "1"
'							sMenu = objJavaWindowProj.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
'						Case "2"
'							sMenu = objJavaWindowProj.WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
'						Case Else
'							Fn_Project_TreeOpeartion = FALSE
'							Exit Function
'					End Select
'					If JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet").WinMenu("ContextMenu").Exist Then
'						JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet").WinMenu("ContextMenu").Select sMenu
'						Fn_Project_TreeOpeartion = TRUE
'					Else
'						Fn_Project_TreeOpeartion = FALSE
'					End If					
'		'----------------------------------------------------------------------- CHECK EXISTANCE OF POP-UP MENU-------------------------------------------------------------------------
'		Case "PopupMenuExist"
'				Call Fn_UI_JavaTree_OpenContextMenu("Fn_Project_TreeOpeartion",objJavaWindowProj,"AMRuleTree",sNodeName)
'				If JavaWindow("Project - Teamcenter 8").JavaApplet("ProjectApplet").WinMenu("ContextMenu").GetItemProperty (sMenu,"Exists") = True Then
'					Fn_Project_TreeOpeartion = TRUE
'				Else
'					Fn_Project_TreeOpeartion = FALSE
'			  	End If
'		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
'		Case "GetIndex"
'				bFlag = False
'				For intCount=0 to objJavaWindowProj.JavaTree("AMRuleTree").GetROProperty ("items count")-1
'					sTreeItem = objJavaWindowProj.JavaTree("AMRuleTree").GetItem (intCount)
'					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
'						Fn_Project_TreeOpeartion = intCount
'						bFlag = True
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
'						Exit For
'					End If
'				Next
'				If bFlag = False Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
'					Fn_Project_TreeOpeartion = FALSE
'				End If
'
'		Case Else
'						Fn_Project_TreeOpeartion = FALSE
'						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Project_TreeOpeartion function failed")
'						Exit Function
'	End Select
'	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_Project_TreeOpeartion")
'	Set objJavaWindowProj = nothing
'	Set objJavaTreeProj = nothing
'End Function
'#########################################################################################################################################
'###    FUNCTION NAME   :    Fn_Auth_ImportExportRule(sAction,sFileName)
'###
'###    DESCRIPTION     :   To Import / Export Rule.
'###
'###    PARAMETERS      :   1. sAction: ImportRule / ExportRule
'### 											 2. sFileName : Fully qualified File name with its Path.
'###                        
'###    Function Calls  :   Fn_WriteLogFile ()
'###
'###    HISTORY         :   		AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   Ketan    				25-June-2010	  1.0
'###
'###    REVIWED BY      :   Harshal	   			25-June-2010	  1.0          
'###
'###    MODIFIED BY     :   
'###
'###    EXAMPLE         :  
'###
'############################################################################################################################################
Function Fn_Auth_ImportExportRule(sAction,sFileName)
		GBL_FAILED_FUNCTION_NAME="Fn_Auth_ImportExportRule"
		Dim ObjImpExp
			Set ObjImpExp = Fn_UI_ObjectCreate( "Fn_Auth_ImportExportRule",JavaWindow("Authorization  - Teamcenter").JavaWindow("AuthorizationApplet"))
			'Check if the Save button is enabled
			If Fn_UI_Object_GetROProperty("Fn_Auth_ImportExportRule",ObjImpExp.JavaButton("Save"), "enabled")=1 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Save button is enabled")
				Call Fn_Button_Click("Fn_Auth_ImportExportRule", ObjImpExp, "Save")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Save button is disabled")
			End If
			Select Case sAction
					Case "Import"
								'Click on ImportRule Button
								Call Fn_Button_Click("Fn_Auth_ImportExportRule", ObjImpExp, "ImportRule")
								'Set value in File Name Editbox.
								call Fn_Edit_Box("Fn_Auth_ImportExportRule",ObjImpExp.JavaDialog("ImportExportRule"),"File name",sFileName)
								'Click on Import Rule button.
								Call Fn_Button_Click("Fn_Auth_ImportExportRule", ObjImpExp.JavaDialog("ImportExportRule"), "ImportExportbtn")
                                Call Fn_ReadyStatusSync(2)
					Case "Export"
								'Click on ExportRule Button
								Call Fn_Button_Click("Fn_Auth_ImportExportRule", ObjImpExp, "ExportRule")
								'Set value in File Name Editbox.
								call Fn_Edit_Box("Fn_Auth_ImportExportRule",ObjImpExp.JavaDialog("ImportExportRule"),"File name",sFileName)
								'Click on Export Rule button.
								Call Fn_Button_Click("Fn_Auth_ImportExportRule", ObjImpExp.JavaDialog("ImportExportRule"), "ImportExportbtn")
                                Call Fn_ReadyStatusSync(2)
					Case Else
							Fn_Auth_ImportExportRule = FALSE
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Auth_ImportExportRule function failed")
							Set ObjImpExp = Nothing
							Exit Function
			End Select										
	Fn_Auth_ImportExportRule = TRUE
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_Auth_ImportExportRule")
	Set ObjImpExp = Nothing
End Function
'#########################################################################################################################################
'###    FUNCTION NAME   :    Fn_CreateSetupWizardFile(FileLocation, TestData)
'###
'###    DESCRIPTION     :   Create text file required for Setup wizard
'###
'###    PARAMETERS      :   1. Location to Save the file
'###						            2. Testdata seprated by ,
'###                        
'###    Function Calls  :   Fn_CreateSetupWizardFile ()
'###
'###    HISTORY         :   		AUTHOR                   DATE        VERSION
'###
'###    CREATED BY      :   		Samir   				30-June-2010	  1.0
'###
'###    REVIWED BY      :   		Harshal					30-June-2010	  1.0          
'###
'###    MODIFIED BY     :   
'###
'###    EXAMPLE         :  
'###
'############################################################################################################################################
Public Function Fn_CreateSetupWizardFile(FileLocation, TestData)
	GBL_FAILED_FUNCTION_NAME="Fn_CreateSetupWizardFile"
	Dim objFSO, objFile
	Dim sBatchFldr, sFilePath

	On Error Resume Next 

    Set objFSO = CreateObject("Scripting.FileSystemObject")	

	If not (objFSO.FileExists(FileLocation)) Then
		Set objFile = objFSO.CreateTextFile(FileLocation)
	Else
		objFSO.DeleteFile sFilePath,True
		Set objFile = objFSO.CreateTextFile(FileLocation)
	End If
	If Err.Number = 0 Then
		Set objFile = objFSO.OpenTextFile(FileLocation,8)
		objFile.Write TestData
		Fn_CreateSetupWizardFile = True
	Else
		Fn_CreateSetupWizardFile = False
	End If

	Set objFSO = Nothing
	Set objLogFile = Nothing

End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Auth_ShowHideApplExt(sAction,sApplication)
'###
'###    DESCRIPTION        :   Show / Hide Applications
'###
'###    PARAMETERS      :   1. sAction: Show / Hide
'###											 2.	sApplication
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR                 DATE        VERSION
'###
'###    CREATED BY     :   Ketan Raje           05/07/2010         1.0
'###
'###    REVIWED BY     :   
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "Show" : Call Fn_Auth_ShowHideApplExt("Show","Workflow Designer:Access Manager:Project")
'###										 Case "Hide" : Call Fn_Auth_ShowHideApplExt("Hide","Workflow Designer:Access Manager:Project")
'#############################################################################################################
Public Function Fn_Auth_ShowHideApplExt(sAction,sApplication)
	GBL_FAILED_FUNCTION_NAME="Fn_Auth_ShowHideApplExt"
	Dim objAppln, iCounter, bReturn, aColname, iCount, iRowData
	Set objAppln = Fn_UI_ObjectCreate("Fn_Auth_ShowHideApplExt", JavaWindow("Authorization  - Teamcenter").JavaWindow("AuthorizationApplet"))
	Fn_Auth_ShowHideApplExt = FALSE
		Select Case sAction
				Case "Show"
						If sApplication<>"" Then
								bReturn = objAppln.JavaList("AvailableApp").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sApplication, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objAppln.JavaList("AvailableApp").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objAppln.JavaList("AvailableApp").Select aColname(iRowData)
											'Click on Add Button
											Call Fn_Button_Click("Fn_Auth_ShowHideApplExt", objAppln, "Add")
											Fn_Auth_ShowHideApplExt = TRUE
											Exit For 
										End If
									Next
								Next
						End If
				Case "Hide"
						If sApplication<>"" Then
								bReturn = objAppln.JavaList("ShownApp").GetROProperty("items count")
								'Extract the index of row at which the object exist.
								aColname = split(sApplication, ":",-1,1)
								iCount = Ubound(aColname)
								For iRowData=0 to iCount
									For iCounter=0 to bReturn-1
										If Trim(lcase(objAppln.JavaList("ShownApp").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
											objAppln.JavaList("ShownApp").Select aColname(iRowData)
											'Click on Remove Button
											Call Fn_Button_Click("Fn_Auth_ShowHideApplExt", objAppln, "Remove")
											Fn_Auth_ShowHideApplExt = TRUE
											Exit For 
										End If
									Next
								Next
						End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Auth_ShowHideApplExt function failed")
						Fn_Auth_ShowHideApplExt = FALSE
						Exit Function						
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of function Fn_Auth_ShowHideApplExt")
    Set objAppln = nothing 	
End Function

'*********************************************************		Function to set User Rate 	***********************************************************************

'Function Name		:					Fn_Org_UserRate

'Description			 :		 		  This function is used to set User Rate value

'Parameters			   :	 			1.  sRate : The default rate value.
'													2. sCurrencys : Currencys for rate.
'												   3.sUser : Full path of user name for default rate need to set.

'Return Value		   : 			True/False 

'Pre-requisite			:		 	Organization pane should be opened.

'Examples				:			Fn_Org_UserRate( "Users:AutoTest2 (autotest2)","$25.00","USD")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Rupali						 27-July-2010	          1.0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Shreyas					 30/08/2011          1.0				Removed Popup menu call			Prasanna
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Org_UserRate(sUser,sRate,sCurrencys)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_UserRate"
   On Error Resume Next
   Dim objUsrRateDlg,bReturn,iIndex,aMenuList,sName
   
   Set objUsrRateDlg = Fn_SISW_Org_GetObject("User Rate")
   If Not objUsrRateDlg.Exist(5) Then


			iIndex = Fn_JavaTree_NodeIndex("Fn_Org_CategoryTreeOperations",JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"),"CategoryTree","OrganizationListTree_ROOT:"+sUser)
			iIndex = Cint(iIndex) + 1
				Call Fn_JavaTree_Select("Fn_Org_CategoryTreeOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"), "CategoryTree","OrganizationListTree_ROOT:"+sUser)
				Wait(2)
			JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaTree("CategoryTree").Deselect sUser
				Wait(2)
				Call Fn_JavaTree_Select("Fn_Org_CategoryTreeOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"), "CategoryTree","OrganizationListTree_ROOT:"+sUser)
				Wait(2)
			sName = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaTree("CategoryTree").GetItem(iIndex)
					'Select node
				Call Fn_JavaTree_Select("Fn_Org_CategoryTreeOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"), "CategoryTree",sName)
				Wait(2)
				Call Fn_JavaTree_Select("Fn_Org_CategoryTreeOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"), "CategoryTree","OrganizationListTree_ROOT:"+sUser)
				Wait(12)
		JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaTree("CategoryTree").Click 0,0,"RIGHT"
				Wait(2)
		JavaWindow("Organization - Teamcenter").WinMenu("ContextMenu").Select "Rate"
				Wait(10)

	    Call Fn_ReadyStatusSync(2)
		If objUsrRateDlg.Exist(20) Then
			objUsrRateDlg.JavaEdit("Rate").Object.setText sRate
			objUsrRateDlg.JavaEdit("Rate").Activate
			index= objUsrRateDlg.JavaList("Currency").GetItemIndex(sCurrencys)
			objUsrRateDlg.JavaList("Currency").Object.setSelectedIndex Cint(index)
			wait(1)
			objUsrRateDlg.JavaButton("Finish").WaitProperty "enabled",1,20000
			objUsrRateDlg.JavaButton("Finish").Click micLeftBtn

		If Err.Number < 0 Then
			Fn_Org_UserRate = False
		   Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Failed to set the user rate value " + sRate + " for " + sUser)	
		   objUsrRateDlg.JavaButton("Cancel").Click micLeftBtn
		   Set objUsrRateDlg = Nothing
		   Exit Function 
		Else 
			Fn_Org_UserRate = True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully set the user rate value " + sRate + " for " + sUser )	
		End If
	Else
		Fn_Org_UserRate = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"User Rate dialog does not exist" )	
		Set objUsrRateDlg = Nothing
	End If 
	End If
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Org_SiteOperations()  
'###
'###    DESCRIPTION        :   Create/Modify/Delete/Verify.
'###	Prequisite 			:	1.Organization Prespective is Open.
'###								2.For Modify/Delete Site Should be selected
'###
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR        		DATE        	VERSION			REVIEWER
'###
'###    CREATED BY     :   Shrikant N      29-Apr-2014                1.0			
'###
'###    MODIFIED BY   :  	
'###
'###									'Set DictSiteOp = CreateObject("Scripting.Dictionary")
'###									'DictSiteOp("sAction")="Create"
'###									'DictSiteOp("sSiteName")="Site123"
'###									'DictSiteOp(sSiteID)="12345"
'###									'DictSiteOp(sLicServer)="Default Local License Server"
'###    EXAMPLE          : 		Case "Create" : Msgbox Fn_Org_SiteOperations(DictSiteOp)
'###
'#############################################################################################################
'Public Function Fn_Org_SiteOperations(sAction , sSiteName , sSiteID , sURL , sSOAURL , sTcGSURL, sGeography, bObjDicServices, bIsAHub, bHTTP, bTCXML, bIsOffline, bAllowDel)

Public Function Fn_Org_SiteOperations(DictSiteOp)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_SiteOperations"
	Dim objSite, iCount, iCounter
	Dim WshShell
	Set objSite = Fn_UI_ObjectCreate("Fn_Org_SiteOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
		Select Case DictSiteOp("sAction")
				Case "Create", "Modify"
						'Set Site Name
						If DictSiteOp("sSiteName")<>"" Then	
							objSite.JavaEdit("Name").SetTOProperty "attached text","Site Name:"
							call Fn_Edit_Box("Fn_Org_SiteOperations",objSite,"Name",DictSiteOp("sSiteName"))
						End If
						'Set Site ID 
						If DictSiteOp("sSiteID")<>"" Then
							objSite.JavaEdit("ID:").SetTOProperty "attached text","Site ID:"
							call Fn_Edit_Box("Fn_Org_SiteOperations",objSite,"ID:",DictSiteOp("sSiteID"))
						End If
						'Set  URL Name 
						If DictSiteOp("sURL")<>"" Then							
							call Fn_Edit_Box("Fn_Org_SiteOperations",objSite,"SiteNodeURL",DictSiteOp("sURL"))
						End If
						'Set SOA URL
						If DictSiteOp("sSOAURL")<>"" Then														
							call Fn_Edit_Box("Fn_Org_SiteOperations",objSite,"SOAURL",DictSiteOp("sSOAURL"))
						End If
						'Set TcGS URL
						If DictSiteOp("sTcGSURL")<>"" Then														
							call Fn_Edit_Box("Fn_Org_SiteOperations",objSite,"TcGSURL",DictSiteOp("sTcGSURL"))
						End If
						'Set License Server 
						If DictSiteOp("sLicServer")<>"" Then
							objSite.JavaEdit("ID:").SetTOProperty "attached text","License Server:"
							call Fn_Edit_Box("Fn_Org_SiteOperations",objSite,"ID:",DictSiteOp("sLicServer"))	
							Set WshShell = CreateObject("WScript.Shell")
							WshShell.SendKeys "{ENTER}"
						End If
						'Set Geography				 
						If DictSiteOp("sGeography")<>"" Then
                            call Fn_Edit_Box("Fn_Org_SiteOperations",objSite,"Geography",DictSiteOp("sGeography"))
						End If						
						'Set Provide Object Dictionary Services.
						If DictSiteOp("bObjDicServices")<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_SiteOperations", objSite,"ProvideObjectDirectory", DictSiteOp("bObjDicServices"))
						End If
						'Set Is A Hub
						If DictSiteOp("bIsAHub")<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_SiteOperations", objSite,"IsAHub", DictSiteOp("bIsAHub"))
						End If
						'Set HTTP Enabled Multi-Site
						If DictSiteOp("bHTTP")<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_SiteOperations", objSite,"HTTPEnabledMultiSite", DictSiteOp("bHTTP"))
						End If
						'Set TCXML PayLoad
						If DictSiteOp(bTCXML)<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_SiteOperations", objSite,"UsesTCXMLPayload", DictSiteOp(bTCXML))
						End If
						'Set Is Offline
						If DictSiteOp("bIsOffline")<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_SiteOperations", objSite,"IsOffline", DictSiteOp("bIsOffline"))
						End If
						
						'Is Unmanaged
						If DictSiteOp("Is Unmanaged")<>"" Then
						objSite.JavaCheckBox("ContentEnabled").SetTOProperty "attached text","Is Unmanaged"
						Call Fn_CheckBox_Set("Fn_Org_SiteOperations",objSite, "ContentEnabled",DictSiteOp("Is Unmanaged"))
						End If
						
						'Briefcase Browser
						If DictSiteOp("Briefcase Browser")<>"" Then
						objSite.JavaCheckBox("LoginEnabled").SetTOProperty "attached text","Briefcase Browser"
						Call Fn_CheckBox_Set("Fn_Org_SiteOperations",objSite, "LoginEnabled",DictSiteOp("Briefcase Browser"))
						End If
						
						'Briefcase Browser with Plugin
						If DictSiteOp("Briefcase Browser with Plugin")<>"" Then
						objSite.JavaCheckBox("ContentEnabled").SetTOProperty "attached text","Briefcase Browser with Plugin"
						Call Fn_CheckBox_Set("Fn_Org_SiteOperations",objSite, "ContentEnabled",DictSiteOp("Briefcase Browser with Plugin"))
						End If
						
						'Set Allow Deletion of replicated Objects.
						If DictSiteOp("bAllowDel")<>"" Then
							Call Fn_CheckBox_Set("Fn_Org_SiteOperations", objSite,"Allowdeletion", DictSiteOp("bAllowDel"))
						End If
						'To Click on Create OR Modify Button.
						If DictSiteOp("sAction")="Create" Then			
								Call Fn_Button_Click("Fn_Org_SiteOperations",objSite,"Create")
						ElseIf DictSiteOp("sAction")="Modify" Then			
								Call Fn_Button_Click("Fn_Org_SiteOperations",objSite,"Modify")
						End If
						Call Fn_ReadyStatusSync(2)
				Case "Delete"
					  'Click on Delete button.
					  Call Fn_Button_Click("Fn_Org_SiteOperations", objSite, "Delete")
					  'Click on yes button to delete the site.
					  For iCount = 0 to 0
					   JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete Confirmation"		'Modified code to handle Msgbox in multiple hierarchy.
					   If JavaDialog("DeleteConfirmation").Exist Then
						JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
						Exit For
					   End If
					   JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").SetTOProperty "title", "Delete Confirmation"
					   If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").Exist Then
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").JavaButton("Yes").Click
						Exit For
					   End If
					  Next
					  Call Fn_ReadyStatusSync(2)
			Case "Verify"
						iCount = 0
						iCounter = 0
						'Verify Site Name
						If DictSiteOp("sSiteName")<>"" Then
							iCount = iCount + 1
							objSite.JavaEdit("Name").SetTOProperty "attached text","Site Name:"
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_SiteOperations",objSite,"Name"))) = Trim(Lcase(DictSiteOp("sSiteName"))) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify Site ID
						If DictSiteOp("sSiteID")<>"" Then
							iCount = iCount + 1
							objSite.JavaEdit("ID:").SetTOProperty "attached text","Site ID:"
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_SiteOperations",objSite,"ID:"))) = Trim(Lcase(DictSiteOp("sSiteID"))) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify Site Node/URL
						If DictSiteOp("sURL")<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_SiteOperations",objSite,"SiteNodeURL"))) = Trim(Lcase(DictSiteOp("sURL"))) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify SOA URL
						If DictSiteOp("sSOAURL")<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_SiteOperations",objSite,"SOAURL"))) = Trim(Lcase(DictSiteOp("sSOAURL"))) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify TcGS URL
						If DictSiteOp("sTcGSURL")<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_SiteOperations",objSite,"TcGSURL"))) = Trim(Lcase(DictSiteOp("sTcGSURL"))) Then
								iCounter = iCounter + 1
							End If
						End If		
						'Verify  License Server				 
						If DictSiteOp("sLicServer")<>"" Then							
							iCount = iCount + 1
							objSite.JavaEdit("ID:").SetTOProperty "attached text","License Server:"
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_SiteOperations",objSite,"ID:"))) = Trim(Lcase(DictSiteOp("sLicServer"))) Then
								iCounter = iCounter + 1
							End If
						End If	
						'Verify  Geography				 
						If DictSiteOp("sGeography")<>"" Then							
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_SiteOperations",objSite,"Geography"))) = Trim(Lcase(DictSiteOp("sGeography"))) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify Provide Object Dictionary Services.
						If DictSiteOp("bObjDicServices")<>"" Then
							iCount = iCount + 1
							If objSite.JavaCheckBox("ProvideObjectDirectory").GetROProperty("value") = 1 and Trim(Lcase(DictSiteOp("bObjDicServices"))) = "on" Then
								iCounter = iCounter + 1
							ElseIf objSite.JavaCheckBox("ProvideObjectDirectory").GetROProperty("value") = 0 and Trim(Lcase(DictSiteOp("bObjDicServices"))) = "off" Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Is A Hub
						If DictSiteOp("bIsAHub")<>"" Then
							iCount = iCount + 1
							If objSite.JavaCheckBox("IsAHub").GetROProperty("value") = 1 and Trim(Lcase(DictSiteOp("bIsAHub"))) = "on" Then
								iCounter = iCounter + 1
							ElseIf objSite.JavaCheckBox("IsAHub").GetROProperty("value") = 0 and Trim(Lcase(DictSiteOp("bIsAHub"))) = "off" Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify HTTP Enabled Multi-Site
						If DictSiteOp("bHTTP")<>"" Then
							iCount = iCount + 1
							If objSite.JavaCheckBox("HTTPEnabledMultiSite").GetROProperty("value") = 1 and Trim(Lcase(DictSiteOp("bHTTP"))) = "on" Then
								iCounter = iCounter + 1
							ElseIf objSite.JavaCheckBox("HTTPEnabledMultiSite").GetROProperty("value") = 0 and Trim(Lcase(DictSiteOp("bHTTP"))) = "off" Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify TCXML PayLoad
						If DictSiteOp("bTCXML")<>"" Then
							iCount = iCount + 1
							If objSite.JavaCheckBox("UsesTCXMLPayload").GetROProperty("value") = 1 and Trim(Lcase(DictSiteOp("bTCXML"))) = "on" Then
								iCounter = iCounter + 1
							ElseIf objSite.JavaCheckBox("UsesTCXMLPayload").GetROProperty("value") = 0 and Trim(Lcase(DictSiteOp("bTCXML"))) = "off" Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Is Offline
						If DictSiteOp("bIsOffline")<>"" Then
							iCount = iCount + 1
							If objSite.JavaCheckBox("IsOffline").GetROProperty("value") = 1 and Trim(Lcase(DictSiteOp("bIsOffline"))) = "on" Then
								iCounter = iCounter + 1
							ElseIf objSite.JavaCheckBox("IsOffline").GetROProperty("value") = 0 and Trim(Lcase(DictSiteOp("bIsOffline"))) = "off" Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Allow Deletion of replicated Objects.
						If DictSiteOp("bAllowDel")<>"" Then
							iCount = iCount + 1
							If objSite.JavaCheckBox("Allowdeletion").GetROProperty("value") = 1 and Trim(Lcase(DictSiteOp("bAllowDel"))) = "on" Then
								iCounter = iCounter + 1
							ElseIf objSite.JavaCheckBox("Allowdeletion").GetROProperty("value") = 0 and Trim(Lcase(DictSiteOp("bAllowDel"))) = "off" Then
								iCounter = iCounter + 1
							End If
						End If
						If iCount <> iCounter Then
							Fn_Org_SiteOperations = FALSE
							Exit Function														
						End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_SiteOperations function failed")
						Fn_Org_SiteOperations = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_Org_SiteOperations")
	Fn_Org_SiteOperations = TRUE
	Set objSite = nothing 
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Org_DisciplinesOperations()  
'###
'###    DESCRIPTION        :   Create/Modify/Delete/Verify.
'###	PrequiDiscipline 					:	1.Organization Prespective is Open.
'###										     2.For Modify/Delete Discipline Should be selected
'###
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR        		DATE        	VERSION			REVIEWER
'###
'###    CREATED BY     :   Ketan Raje        12-May-2011         1.0				Harshal A.
'###
'###    MODIFIED BY   :  	
'###
'###    EXAMPLE          : 		Case "Create" : Msgbox Fn_Org_DisciplinesOperations("Create", "Dis3", "TestDis", "$9.00", "USD", "AutoTest1 (autotest1):AutoTest2 (autotest2)", "")
'###    								Case "Modify" : Msgbox Fn_Org_DisciplinesOperations("Modify", "Dis3", "TestingDis", "$7.00", "USD", "AutoTest3 (autotest3)", "AutoTest1 (autotest1):AutoTest2 (autotest2)")
'###    								Case "Verify" : Msgbox Fn_Org_DisciplinesOperations("Verify", "Dis3", "TestDis", "$9.00", "USD", "AutoTest1 (autotest1):AutoTest2 (autotest2)", "")
'#############################################################################################################
Public Function Fn_Org_DisciplinesOperations(sAction , sName, sDescription, sRate, sCurrency, sAddUsers, sRemoveUsers)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_DisciplinesOperations"
	Dim objDiscipline, iCount, iCounter, iRows, aAddUsers, arrAddUsers, aRemoveUsers, intCount
	Set objDiscipline = Fn_UI_ObjectCreate("Fn_Org_DisciplinesOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet"))
		Select Case sAction
				Case "Create", "Modify","NoButtonClick"
						'Set Discipline Name
						If sName <> "" Then	
							call Fn_Edit_Box("Fn_Org_DisciplinesOperations",objDiscipline,"Name",sName)
						End If
						'Set Discipline Description
						If sDescription <> "" Then
							call Fn_Edit_Box("Fn_Org_DisciplinesOperations",objDiscipline,"Description",sDescription)
						End If
						'Set Defalut Rate
						If sRate <> "" Then
							call Fn_Edit_Box("Fn_Org_DisciplinesOperations",objDiscipline,"DefaultRate",sRate)
						End If
						'Set Defalut Currency
						If sCurrency <> "" Then
							call Fn_Edit_Box("Fn_Org_DisciplinesOperations",objDiscipline,"DefaultCurrency",sCurrency)
						End If
						'Associate Users
						If sAddUsers <> ""  Then
							aAddUsers = Split(sAddUsers,":",-1,1)
							For iCount = 0 to Ubound(aAddUsers)
								'ReDim arrAddUsers(2)
								arrAddUsers = Split(aAddUsers(iCount),"(",-1,1)
								'Enter the name of user in search Editbox
								call Fn_Edit_Box("Fn_Org_DisciplinesOperations",objDiscipline,"Users","")
								call Fn_Edit_Box("Fn_Org_DisciplinesOperations",objDiscipline,"Users",Trim(arrAddUsers(0)))
								'Click on search button
								Call Fn_Button_Click("Fn_Org_DisciplinesOperations",objDiscipline,"Search")
								iRows = objDiscipline.JavaList("UsersList").GetROProperty("items count")
								If Cint(iRows) <> 0 Then
									'Select User from List
									objDiscipline.JavaList("UsersList").Select aAddUsers(iCount)
									'Click on AddUser button
									Call Fn_Button_Click("Fn_Org_DisciplinesOperations", objDiscipline, "Add")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aAddUsers(iCount) &" User not found in UserList")
									Fn_Org_DisciplinesOperations = False
									Set objDiscipline = nothing 
									Exit Function
								End If
							Next
						End If
						'Remove Users
						If sRemoveUsers <> ""  Then
							aRemoveUsers = Split(sRemoveUsers,":",-1,1)
							For iCount = 0 to Ubound(aRemoveUsers)
								iRows = objDiscipline.JavaList("AssociatedUsersList").GetROProperty("items count")
								If Cint(iRows) <> 0 Then
									'Select User from List
									objDiscipline.JavaList("AssociatedUsersList").Select aRemoveUsers(iCount)
									'Click on AddUser button
									Call Fn_Button_Click("Fn_Org_DisciplinesOperations", objDiscipline, "Remove")
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), aAddUsers(iCount) &" User not found in Associated UserList")
									Fn_Org_DisciplinesOperations = False
									Set objDiscipline = nothing 
									Exit Function
								End If
							Next
						End If
						wait(5)
						'To Click on Create OR Modify Button.
						If sAction="Create" Then			
								Call Fn_Button_Click("Fn_Org_DisciplinesOperations",objDiscipline,"Create")
						ElseIf sAction="Modify" Then			
								Call Fn_Button_Click("Fn_Org_DisciplinesOperations",objDiscipline,"Modify")
						End If
						Call Fn_ReadyStatusSync(2)
			Case "Verify"
						iCount = 0
						iCounter = 0
						'Verify Discipline Name
						If sName <> "" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_DisciplinesOperations",objDiscipline,"Name"))) = Trim(Lcase(sName)) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify Discipline Description.
						If sDescription <> "" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_DisciplinesOperations",objDiscipline,"Description"))) = Trim(Lcase(sDescription)) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify Default Rate
						If sRate <> "" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_DisciplinesOperations",objDiscipline,"DefaultRate"))) = Trim(Lcase(sRate)) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify Default Currency
						If sCurrency <> "" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_DisciplinesOperations",objDiscipline,"DefaultCurrency"))) = Trim(Lcase(sCurrency)) Then
								iCounter = iCounter + 1
							End If
						End If						
						'Verify associated Users
						If sAddUsers <> "" Then
							aAddUsers = Split(sAddUsers,":",-1,1)
							iRows = objDiscipline.JavaList("AssociatedUsersList").GetROProperty("items count")
							For intCounter = 0 to Ubound(aAddUsers)
								iCount = iCount + 1
								For intCount = 0 to Cint(iRows)-1
									If Trim(Lcase(objDiscipline.JavaList("AssociatedUsersList").GetItem(intCount))) = Trim(Lcase(aAddUsers(intCounter))) Then
										iCounter = iCounter + 1
										Exit For
									End If									
								Next
							Next
						End If
						If iCount <> iCounter Then
							Fn_Org_DisciplinesOperations = FALSE
							Exit Function														
						End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_DisciplinesOperations function failed")
						Fn_Org_DisciplinesOperations = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_Org_DisciplinesOperations")
	Fn_Org_DisciplinesOperations = TRUE
	Set objDiscipline = nothing 
End Function
'*********************************************************		Function to  Set Schedule Calendar	***********************************************************************

'Function Name		:					Fn_Org_CalendarOperations

'Description			 :		 		  This function is used to set Calendar

'Parameters			   :	 			1.  sAction: Action to Execute
'												3. dicCalendar: Calendar Dictionary
											
'Return Value		   : 				 True/False

'Pre-requisite			:		 		Organization window should be displayed .

'Examples				:				Call Fn_Org_CalendarOperations("Create", dicCalendar)

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Ketan Raje				13-May-2011	   1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_Org_CalendarOperations(sAction, dicCalendar)
			GBL_FAILED_FUNCTION_NAME="Fn_Org_CalendarOperations"
			Dim bReturn, objCalWin
			Dim aDays, objDailyWin, iCount
			Dim iCounter, iRow, iCol, intCount
			Dim aRowData, aCellData

			On Error Resume Next
			
			Set objCalWin = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet")
			
			Select Case sAction
			Case "Create","Modify"
					'Set Calendar Name.
					If dicCalendar("CalendarName") <> "" Then						
						Call Fn_Edit_Box("Fn_Org_CalendarOperations", objCalWin,"CalendarName",dicCalendar("CalendarName"))
					End If
					'Set Year
					If dicCalendar("sYear") <> "" Then						
						Call Fn_Edit_Box("Fn_Org_CalendarOperations", objCalWin,"Year",dicCalendar("sYear"))
					End If
					'Set Month
					If dicCalendar("sMonth") <> "" Then
						objCalWin.JavaStaticText("Month").SetToProperty "Label",dicCalendar("sMonth")
						While Not JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaStaticText("Month").Exist(2)
							'Click on right month button
							Call Fn_Button_Click("Fn_Org_CalendarOperations",objCalWin,"RightMonth")
						Wend
					End If
					'Set Date
					If dicCalendar("sDate") <> "" Then						
						objCalWin.JavaCheckBox("Date").SetToProperty "attached text",dicCalendar("sDate")
						Call Fn_CheckBox_Select("Fn_Org_CalendarOperations", objCalWin, "Date")
					End If
					'Select Time Zone
					If dicCalendar("sTimeZone") <> "" Then
						objCalWin.JavaList("TimeZoneList").Select dicCalendar("sTimeZone")
					End If		
					'Set Working Days for Schedule
					If dicCalendar("OnWeekDays") <> "" Then
						aDays = split(dicCalendar("OnWeekDays"), "~", -1, 1)
						'Check the Days Check-Boxes for the required week days
						For iCounter = 0 to Ubound(aDays)
							objCalWin.JavaCheckBox("DayCheck").SetTOProperty "attached text", aDays(iCounter)
							objCalWin.JavaCheckBox("DayCheck").Set "ON"
						Next
					End If
					'Set Non-Working Days for Schedule
					If dicCalendar("OffWeekDays") <> "" Then
						aDays = split(dicCalendar("OffWeekDays"), "~", -1, 1)
						'Check the Days Check-Boxes for the required week days
						For iCounter = 0 to Ubound(aDays)
							objCalWin.JavaCheckBox("DayCheck").SetTOProperty "attached text", aDays(iCounter)
							objCalWin.JavaCheckBox("DayCheck").Set "OFF"
						Next
					End If
					Set objDailyWin = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Daily Defaults Details")
					If dicCalendar("DayHrDetails") <> "" Then
								'Set Working Hours for All the Working Days
								For iCounter = 1 to 7
									objCalWin.JavaButton("Details...").SetTOProperty "index", iCounter
									wait(1)
									If objCalWin.JavaButton("Details...").GetROProperty("enabled") Then
										objCalWin.JavaButton("Details...").Click micLeftBtn
										'Add Daily Details
										If objDailyWin.Exist(10) Then
											'Set Specific Times Radio ON
											objDailyWin.JavaRadioButton("SpecificTimes").Set "ON"	
											'Set the Timings in the Table
											If dicCalendar("DayHrDetails") <> "" Then
												aRowData = split(dicCalendar("DayHrDetails"), ",", -1, 1)
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Data for Detail Hors not Provided Properly")
												Fn_Org_CalendarOperations = False
												objDailyWin.JavaButton("Cancel").Click micLeftBtn
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
												Next
											Next
											'Click OK on Daily Details Dialog
											objDailyWin.JavaButton("OK").Click micLeftBtn
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Find [Daily Defaults Details] Dialog")
											Fn_Org_CalendarOperations = False
											Set objCalWin = Nothing
											Set objDailyWin = Nothing
											Exit Function
										End If
									End If
								Next
					End If
					'Click on Create or Modify Button depending on case.
					If Trim(Lcase(sAction)) = "create" Then
						'Click on Create Button
						Call Fn_Button_Click("Fn_Org_CalendarOperations",objCalWin,"Create")
					ElseIf Trim(Lcase(sAction)) = "modify" Then
						'Click on Modify Button
						Call Fn_Button_Click("Fn_Org_CalendarOperations",objCalWin,"Modify")
					End If

			Case "Verify"
					'Verify Calendar Name.
					If dicCalendar("CalendarName") <> "" Then						
						iCount = iCount + 1
						If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_CalendarOperations", objCalWin,"CalendarName"))) = Trim(Lcase(dicCalendar("CalendarName"))) Then
							intCount = intCount + 1
						End If
					End If
					'Verify Year
					If dicCalendar("sYear") <> "" Then
						iCount = iCount + 1
						If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_Org_CalendarOperations", objCalWin,"Year"))) = Trim(Lcase(dicCalendar("sYear"))) Then
							intCount = intCount + 1
						End If
					End If
					'Verify Month
					If dicCalendar("sMonth") <> "" Then
						iCount = iCount + 1
						objCalWin.JavaStaticText("Month").SetToProperty "Label",dicCalendar("sMonth")
						If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaStaticText("Month").Exist(5) Then
							intCount = intCount + 1
						End If
					End If
					'Verify Date
					If dicCalendar("sDate") <> "" Then
						iCount = iCount + 1
						objCalWin.JavaCheckBox("Date").SetToProperty "attached text",dicCalendar("sDate")
						If Trim(Lcase(objCalWin.JavaCheckBox("Date").GetROProperty("background"))) = "blue" Then
							intCount = intCount + 1
						End If
					End If
					'Verify Time Zone
					If dicCalendar("sTimeZone") <> "" Then
						iCount = iCount + 1						
						If Trim(Lcase(objCalWin.JavaList("TimeZoneList").GetROProperty("value"))) = Trim(Lcase(dicCalendar("sTimeZone"))) Then
							intCount = intCount + 1
						End If
					End If		
					'Verify Working Days for Schedule
					If dicCalendar("OnWeekDays") <> "" Then
						aDays = split(dicCalendar("OnWeekDays"), "~", -1, 1)
						'Check the Days Check-Boxes for the required week days
						For iCounter = 0 to Ubound(aDays)
							iCount = iCount + 1
							objCalWin.JavaCheckBox("DayCheck").SetTOProperty "attached text", aDays(iCounter)				
							If objCalWin.JavaCheckBox("DayCheck").GetROProperty("value") = 1 Then
								intCount = intCount + 1
							End If
						Next
					End If
					'Verify Non-Working Days for Schedule
					If dicCalendar("OffWeekDays") <> "" Then
						aDays = split(dicCalendar("OffWeekDays"), "~", -1, 1)
						'Check the Days Check-Boxes for the required week days
						For iCounter = 0 to Ubound(aDays)
							iCount = iCount + 1
							objCalWin.JavaCheckBox("DayCheck").SetTOProperty "attached text", aDays(iCounter)
							If objCalWin.JavaCheckBox("DayCheck").GetROProperty("value") = 0 Then
								intCount = intCount + 1
							End If
						Next
					End If
					If intCount <> iCount Then
							Fn_Org_CalendarOperations = False
							Set objCalWin = Nothing
							Set objDailyWin = Nothing
							Exit Function
					End If
			Case "Delete"
					  'Click on Delete button.
					  Call Fn_Button_Click("Fn_Org_CalendarOperations", objCalWin, "Delete")
					  'Click on yes button to delete the site.
					  For iCount = 0 to 0
					   JavaDialog("DeleteConfirmation").SetTOProperty "title", "Delete Confirmation"		'Modified code to handle Msgbox in multiple hierarchy.
					   If JavaDialog("DeleteConfirmation").Exist Then
						JavaDialog("DeleteConfirmation").JavaButton("Yes").Click
						Exit For
					   End If
					   JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").SetTOProperty "title", "Delete Confirmation"
					   If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").Exist Then
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("MsgDialog").JavaButton("Yes").Click
						Exit For
					   End If
					  Next
					  Call Fn_ReadyStatusSync(2) 

			End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" of function Fn_Org_CalendarOperations successfully executed.")
	Fn_Org_CalendarOperations = True
	Set objCalWin = Nothing
	Set objDailyWin = Nothing
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_Org_AddDisciplinesOperations()
'###
'###	PREQUISITE		   :		1. DBA Pervilages needed.
'###									  2. Organization Prespective is Open.
'###									  3. Select the Group from Organization Tree
'###
'###    DESCRIPTION        :   Add/Remove Roles From Groups
'###
'###    PARAMETERS         :   1. sAction: Add/Remove
'###								   2. sDispName: ":" Seperated String(for AddExisting)
'###								   3. sDispDesc
'###                                         
'###    Function Calls       :   Fn_WriteLogFile() 
'###
'###	 HISTORY             :   AUTHOR        				DATE      		  VERSION
'###
'###    CREATED BY     	   :  Priyanka Bhave           23/11/2011      	   1.0
'###
'###    REVIWED BY     	   :  prasanna
'###
'###    MODIFIED BY   :  
'###
'###    EXAMPLE          : 		Case "AddNew" : Call Fn_Org_AddDisciplinesOperations("AddNew","MyDisp1","MyDisp1")
'###										 Case "AddExisting " : Call Fn_Org_AddDisciplinesOperations("AddExisting","AutoDisp1:AutoDisp2","")
'###										 Case "AddAllExisting" : Call Fn_Org_AddDisciplinesOperations("AddAllExisting","","")
'###										 Case "Remove" : Call Fn_Org_AddDisciplinesOperations("Remove","","")
'###                                         Case "Modify" :  Call Fn_Org_AddDisciplinesOperations("Modify","MyDisp123","MyDisp123")	
'#############################################################################################################

Public Function Fn_Org_AddDisciplinesOperations(sAction,sDispName,sDispDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_AddDisciplinesOperations"
	Dim objDsp, objOrgDsp, aColname, iCounter, bReturn, iCount, iRowData
	Set objOrgDsp = JavaWindow("Organization - Teamcenter").JavaWindow("JApplet")
		Select Case sAction
				Case "AddNew"
						If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("OrganizationDisciplineWizard").Exist(10) = False  Then
							'Click on AddDiscipline button
							Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objOrgDsp, "AddDiscipline")	
							Call Fn_ReadyStatusSync(2)
						End If
						'Set obj for Discipline wizard
						Set objDsp = Fn_UI_ObjectCreate("Fn_Org_AddDisciplinesOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("OrganizationDisciplineWizard"))						
						'Select the Add new discipline radio button						
						Call Fn_UI_JavaRadioButton_SetON("Fn_Org_AddDisciplinesOperations",objDsp, "Add new discipline")
						'Click on Next button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Next")
						'Set discipline
						Call  Fn_Edit_Box("Fn_Org_AddDisciplinesOperations",objDsp,"Name",sDispName)
						'Set Description
						Call  Fn_Edit_Box("Fn_Org_AddDisciplinesOperations",objDsp,"Description",sDispDesc)
						Call Fn_ReadyStatusSync(2)
						'Click on finish button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Finish")
						'Click on yes button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Yes")
						'Set property of the window
						'Call Fn_UI_Object_SetTOProperty("Fn_Org_RoleOperations",JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"),"text","Role(s) added")
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error").SetTOProperty "title","Discipline(s) added"
						'Click on OK button
						'Call Fn_Button_Click("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"), "OK")
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error").JavaButton("OK").Click micLeftBtn
						'Click on close button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Close")

				Case "AddExisting"
						If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("OrganizationDisciplineWizard").Exist(10) = False  Then
							'Click on AddDiscipline button
							Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objOrgDsp, "AddDiscipline")	
							Call Fn_ReadyStatusSync(2)
						End If
						'Set obj for Discipline wizard
						Set objDsp = Fn_UI_ObjectCreate("Fn_Org_AddDisciplinesOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("OrganizationDisciplineWizard"))						
						'Select the Add Existing Discipline radio button						
						Call Fn_UI_JavaRadioButton_SetON("Fn_Org_AddDisciplinesOperations",objDsp, "Add existing discipline")
						'Click on Next button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Next")
						'Set Discipline
						bReturn = objDsp.JavaList("ExistingDiscipline").GetROProperty("items count")
						'Extract the index of row at which the object exist.
						aColname = split(sDispName, ":",-1,1)
						iCount = Ubound(aColname)
						For iRowData=0 to iCount
							For iCounter=0 to bReturn-1
								If Trim(lcase(objDsp.JavaList("ExistingDiscipline").GetItem(iCounter))) = Trim(lcase(aColname(iRowData))) then
									objDsp.JavaList("ExistingDiscipline").Select aColname(iRowData)
									'Click on Add column Button
									Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Add")
									Exit For 
								End If
							Next
						Next		
						Call Fn_ReadyStatusSync(2)				
						'Click on finish button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Finish")
						'Click on yes button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Yes")
						'Set property of the window
						'Call Fn_UI_Object_SetTOProperty("Fn_Org_RoleOperations",JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"),"text","Role(s) added")
						'JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Discipline(s) added"
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error").SetTOProperty "title","Discipline(s) added"
						
						'Click on OK button
						'Call Fn_Button_Click("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"), "OK")
						'JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error").JavaButton("OK").Click micLeftBtn
						
						'Click on close button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Close")

				Case "AddAllExisting"
						If JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("OrganizationDisciplineWizard").Exist(10) = False  Then
							'Click on AddDiscipline button
							Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objOrgDsp, "AddDiscipline")	
							Call Fn_ReadyStatusSync(2)
						End If
						'Set obj for Discipline wizard
						Set objDsp = Fn_UI_ObjectCreate("Fn_Org_AddDisciplinesOperations", JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("OrganizationDisciplineWizard"))						
						'Select the Add existing Discipline radio button						
						Call Fn_UI_JavaRadioButton_SetON("Fn_Org_AddDisciplinesOperations",objDsp, "Add existing discipline")
						'Click on Next button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Next")
						'Set discipline
						bReturn = objDsp.JavaList("ExistingDiscipline").GetROProperty("items count")
						'Selecting All the items in the List.
						objDsp.JavaList("ExistingDiscipline").SelectRange 0,(bReturn-1)
						'Click on Add column Button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Add")
						Call Fn_ReadyStatusSync(2)
						'Click on finish button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Finish")
						'Click on yes button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Yes")
						'Set property of the window
						'Call Fn_UI_Object_SetTOProperty("Fn_Org_RoleOperations",JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"),"text","Role(s) added")
						'JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").SetTOProperty "title","Discipline(s) added"
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error").SetTOProperty "title","Discipline(s) added"
						
						'Click on OK button
						'Call Fn_Button_Click("Fn_Org_RoleOperations", JavaWindow("Organization - Teamcenter").Dialog("ErrorDialog"), "OK")
						'JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("ErrorDialog").JavaButton("OK").Click micLeftBtn
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaDialog("Error").JavaButton("OK").Click micLeftBtn
						
						'Click on close button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objDsp, "Close")

				Case "Remove"
						'Click on remove button.
						JavaWindow("Organization - Teamcenter").JavaWindow("JApplet").JavaButton("Remove").SetTOProperty "label", "Remove"
                        Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objOrgDsp, "Remove")
						Call Fn_ReadyStatusSync(2)
						
                Case "Modify"
                        'Set Discipline
						If sDispName <> "" Then
							Call Fn_Edit_Box("Fn_Org_AddDisciplinesOperations",objOrgDsp,"Name",sDispName)
						End If
						'Set Description
						If sDispDesc <> "" Then
							Call Fn_Edit_Box("Fn_Org_AddDisciplinesOperations",objOrgDsp,"Description",sDispDesc)
						End If
						'Click on modify button
						Call Fn_Button_Click("Fn_Org_AddDisciplinesOperations", objOrgDsp, "Modify")
						Call Fn_ReadyStatusSync(2)

                Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_Org_AddDisciplinesOperations function failed")
						Fn_Org_RoleOperations = FALSE
						Exit Function						
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Discipline [" + sDispName + "] of function Fn_Org_AddDisciplinesOperations")
	Fn_Org_AddDisciplinesOperations = TRUE
	Set objOrgDsp = nothing 
	Set objDsp = nothing 
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function Name		:	Fn_Org_LanguageOperations

'Description			 :	Function Used perform operations on Languages ( action like create language , modify language , delete language)

'Parameters			   :   '1.StrAction: Action Name
'										 2.StrLanguageName : Language name
'										 3.StrModifiedLangName : new name for existing language in case of modification
'										 4.StrISOLangCode : ISO Language code
'										 5.StrISOCountryCode : ISO Language country
'										 6.StrLangFileInitials : Language file initials
'										 7.StrLangDesc : Language Description
'										 8.StrDefaultPublishingFont : DefaultPublishingFont  of language
'										 9.StrDescription : General description
'										 10.bLoginEnabled : Login Enabled option
'										 11.bMetadataEnabled : Metadata Enabled option
'										 12.bContentEnabled : Content Enabled option
'
'Return Value		   : 	True or False

'Pre-requisite			:	Organization perspective should be activated

'Examples				:  	Fn_Org_LanguageOperations("Create","TestLang1","","aa, Afar","TV, Tuvalu","AT","Tuvalu Lang","Helvetica","Test language creation","off","","")
'										Fn_Org_LanguageOperations("Modify","TestLang1","TestLang1_Mod","ak, Akan","UG, Uganda","AU","Unganda lnguage","","modified Test language creation","","off","off")
'										Fn_Org_LanguageOperations("Delete","TestLang1_Mod","","","","","","","","","","")	
'
'History					 :			
'													Developer Name											Date								Rev. No.						Changes Done						Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'													Sandeep N												10-Apr-2012								1.0																						Sunny R
'													Sandeep N												14-Mar-2013								1.1					Modified case : Create						Sunny R
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Function Fn_Org_LanguageOperations(StrAction,StrLanguageName,StrModifiedLangName,StrISOLangCode,StrISOCountryCode,StrLangFileInitials,StrLangDesc,StrDefaultPublishingFont,StrDescription,bLoginEnabled,bMetadataEnabled,bContentEnabled)
	GBL_FAILED_FUNCTION_NAME="Fn_Org_LanguageOperations"
 	'Declaring variables
	Dim objApplet,bFlag,objStaticText,objAppletChild
    Dim WshShell
	Fn_Org_LanguageOperations=False
    Set WshShell = CreateObject("WScript.Shell")
	'Creating object of [ JApplet ]
	Set objApplet=JavaWindow("Organization - Teamcenter").JavaWindow("JApplet")
	Select Case StrAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Create","Modify"
			'Assigning bFlag to False
			bFlag=False
			If StrAction="Create" Then
				'Selecting [ Language ] node from Category Tree
				bFlag=Fn_Org_CategoryTreeOperations("Select","Language")
			Else
				'Expanding [ Language ] node from Category Tree
				Call Fn_Org_CategoryTreeOperations("Expand","Language")	
				'Selecting specific Language from Category Tree to modify
				bFlag=Fn_Org_CategoryTreeOperations("Select","Language:"+StrLanguageName)
			End If
			If bFlag=False Then
				Set objApplet=Nothing
				Exit function
			End If
			Call Fn_ReadyStatusSync(1)
			If StrAction="Create" Then
				'Entering Language name
				If StrLanguageName<>"" Then
					Call Fn_Edit_Box("Fn_Org_LanguageOperations",objApplet,"LanguageName", StrLanguageName)
				End If
			Else
				'Entering New Language name
				If StrModifiedLangName<>"" Then
					Call Fn_Edit_Box("Fn_Org_LanguageOperations",objApplet,"LanguageName", StrModifiedLangName)
				End If	
			End If

			If StrISOLangCode<>"" Then
				'Selecting ISO Language code
'				objApplet.JavaStaticText("LanguageFieldName").SetTOProperty "label","ISO Language Code:"
'				Call Fn_Button_Click("Fn_Org_LanguageOperations",objApplet,"LanguageDropDown")
'				wait 2
'				Set objStaticText=Description.Create()
'				objStaticText("Class Name").value="JavaStaticText"
'				objStaticText("label").value=StrISOLangCode
'				Set objAppletChild=objApplet.ChildObjects(objStaticText)
'				objAppletChild(0).Click 1,1
'				Set objStaticText=Nothing
'				Set objAppletChild=Nothing
				objApplet.JavaEdit("ISOCountryCode").Click 1,1,"LEFT"
				Call Fn_Edit_Box("Fn_Org_LanguageOperations",objApplet,"ISOLanguageCode", StrISOLangCode)
				wait 2
				WshShell.SendKeys "{ENTER}"
				wait 1
			End If
			If StrISOCountryCode<>"" Then
				'Selecting ISO Country code
'				objApplet.JavaStaticText("LanguageFieldName").SetTOProperty "label","ISO Country Code:"
'				Call Fn_Button_Click("Fn_Org_LanguageOperations",objApplet,"LanguageDropDown")
'				wait 2
'				Set objStaticText=Description.Create()
'				objStaticText("Class Name").value="JavaStaticText"
'				objStaticText("label").value=StrISOCountryCode
'				Set objAppletChild=objApplet.ChildObjects(objStaticText)
'				objAppletChild(0).Click 1,1
'				Set objStaticText=Nothing
'				Set objAppletChild=Nothing
				objApplet.JavaEdit("ISOLanguageCode").Click 1,1,"LEFT"
				Call Fn_Edit_Box("Fn_Org_LanguageOperations",objApplet,"ISOCountryCode", StrISOCountryCode)
				wait 2
				WshShell.SendKeys "{ENTER}"
				wait 1
			End If
			If StrLangFileInitials<>"" Then
				'Setting Language File Initials
				Call Fn_Edit_Box("Fn_Org_LanguageOperations",objApplet,"LanguageFileInitials", StrLangFileInitials)
			End If
			If StrLangDesc<>"" Then
				'Setting Language Description 
				Call Fn_Edit_Box("Fn_Org_LanguageOperations",objApplet,"LanguageDescription", StrLangDesc)
			End If
			If StrDefaultPublishingFont<>"" Then
				'Setting Default Publishing Font
				Call Fn_Edit_Box("Fn_Org_LanguageOperations",objApplet,"DefaultPublishingFont", StrDefaultPublishingFont)
			End If
			If StrDescription<>"" Then
				'Setting Description
				Call Fn_Edit_Box("Fn_Org_LanguageOperations",objApplet,"Description", StrDescription)
			End If
			If bLoginEnabled<>"" Then
				'Setting Login Enabled option
				Call Fn_CheckBox_Set("Fn_Org_LanguageOperations", objApplet,"LoginEnabled",bLoginEnabled)
			End If
			If bMetadataEnabled<>"" Then
				'Setting Metadata Enabled option
				Call Fn_CheckBox_Set("Fn_Org_LanguageOperations", objApplet,"MetadataEnabled",bMetadataEnabled)
			End If
			If bContentEnabled<>"" Then
				'Setting Content Enabled option
				Call Fn_CheckBox_Set("Fn_Org_LanguageOperations", objApplet,"ContentEnabled",bContentEnabled)
			End If
			If StrAction="Create" Then
				'Clicking on Create button to create new Language
				Call Fn_Button_Click("Fn_Org_LanguageOperations",objApplet,"Create")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully created Language [ " &StrLanguageName&" ] ")
			Else
				'Clicking on Modify button to Modify existing Language
				Call Fn_Button_Click("Fn_Org_LanguageOperations",objApplet,"Modify")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully modified Language [ " &StrLanguageName&" ] ")
			End If
			Call Fn_ReadyStatusSync(1)
			Fn_Org_LanguageOperations=True
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Delete"
			'Expanding [ Language ] node from Category Tree
			Call Fn_Org_CategoryTreeOperations("Expand","Language")	
			'Selecting specific Language from Category Tree to Delete
			bFlag=Fn_Org_CategoryTreeOperations("Select","Language:"+StrLanguageName)
			If bFlag=False Then
				Set objApplet=Nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Language [ " &StrLanguageName&" ] is not exist under Language node")
				Exit function
			End If
			'Clicking on Delete button to create new Language
			Call Fn_Button_Click("Fn_Org_LanguageOperations",objApplet,"Delete")
			Call Fn_ReadyStatusSync(1)
			If objApplet.JavaDialog("DeleteConfirmation").Exist(6) Then
					'Clicking on yes button to Confirm deletion of language
					Call Fn_Button_Click("Fn_Org_LanguageOperations",objApplet.JavaDialog("DeleteConfirmation"),"Yes")
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully deleted Language [ " &StrLanguageName&" ] ")
					Fn_Org_LanguageOperations=True
			End If
	End Select
    Set objApplet=Nothing
	'Releasing object of [ JApplet ]
	Set objApplet=Nothing
End Function
