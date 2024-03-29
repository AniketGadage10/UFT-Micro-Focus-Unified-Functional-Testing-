Option Explicit


'*********************************************************	Function List		***********************************************************************
'1. Fn_AudMgr_AuditLogOperation()
'2. Fn_AudMgr_AuditDefinitionOperation()
'3. Fn_AudMgr_NavTreeOperations()
'4. Fn_SISW_AudMgr_SrchResultValidations()
'*********************************************************	Function List		***********************************************************************s
'##############################################################################################################################################################################################################################################
'###
'###    FUNCTION NAME   :  Fn_AudMgr_AuditLogOperation()
'###
'###    DESCRIPTION     :  Verify the content in Audit log table. sTableHeaderName and aTableContent parameter should be identical.
'###
'###    PARAMETERS      :   sAction = "Find or Clear or Export Audit Log or Cancel"
'###                        sTableHeaderName = "UserId~EventTypeName"
'###                        aTableContent = "x_baldot~Check-out","x_baldot~Check-in"
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'###
'###	 HISTORY         :   AUTHOR                 DATE        	VERSION
'###
'###    CREATED BY      :   Mahendra Bhandarkar	  10/06/2010        1.0
'###
'###    REVIWED BY      :   Mohit Khare		 	  11/06/2010	    1.0
'###
'###    MODIFIED BY     :
'###    EXAMPLE         : Fn_AudMgr_AuditLogOperation("Find", "UserId~EventTypeName", "autotestdba~Modify, autotestdba~Modify", "")
'#######################################################################################################################################################################################################################

Public Function Fn_AudMgr_AuditLogOperation(sAction, sTableHeaderName, aTableContent, aSearchDictionary )
	GBL_FAILED_FUNCTION_NAME="Fn_AudMgr_AuditLogOperation"
	Dim objDialog, objTableDialog, iRows, iCols, bFlag, oDesc, aTableHeader, aTableContent1, aTableContent2, intCount1, intCounter, sValOpt
	Dim iRowSelect, sValOpt2, intCount2, intCounted, aTemp, intCounter1, intCount3
	
	bFlag = False
	
	If Fn_UI_ObjectExist("Fn_AudMgr_AuditLogOperation", Window("AuditManagerWindow")) = True Then
		Call Fn_Preference_Search_Operation("Modify","TC_audit_manager","ON")
		Call Fn_AudMgr_NavTreeOperations("Modify", "Dataset - Check-Out", "checked_out_date:checked_out_user")
		Call Fn_AudMgr_NavTreeOperations("Modify", "Dataset - Check-In", "checked_out_date:checked_out_user")
		Call Fn_SetPerspective("My Teamcenter")
		Wait 3
	End If
	
	If Fn_UI_ObjectExist("Fn_AudMgr_AuditLogOperation", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log")) = False Then
		Call Fn_MenuOperation("Select", "View:Audit:View Audit Logs")
		Wait 2
	End If
	
	Set objDialog = Fn_UI_ObjectCreate( "Fn_AudMgr_AuditLogOperation", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Audit Log"))
	Set objTableDialog = Fn_UI_ObjectCreate("Fn_AudMgr_AuditLogOperation", objDialog.JavaTable("LogTable"))

	Select Case sAction
	
		Case "Find"
		
			Call Fn_Button_Click("Fn_AudMgr_AuditLogOperation", objDialog, "Find")
		
			aTableHeader = Split(sTableHeaderName, "~", -1, 1)
		
			aTableContent1 = Split(aTableContent, ",", -1, 1)
		
			iRows = objTableDialog.GetROProperty("rows")
			iCols = objTableDialog.GetROProperty("cols")
		
			intCount2 = 0
			
			For intCounter = 0 To UBound(aTableContent1)
		
			aTableContent2 = Split(aTableContent1(intCounter), "~", -1, 1)
		
			If bFlag = True OR intCounter = 0 Then
		
				For intCount1 = 0 To iRows - 1
		
				intCounted = 0
			
				For intCount2 = 0 To UBound(aTableContent2)
						aTemp = Split(aTableContent2(intCount2), ":", -1, 1)
						intCounter1 = 0
						For intCount3 = 0 To UBound(aTemp)
							If InStr(1, objTableDialog.GetCellData(intCount1,aTableHeader(intCounter)), aTemp(intCount3), 1 ) > 0 Then
									intCounter1 = intCounter1 + 1
							End If
						Next
						
						If UBound(aTemp) =  intCounter1 - 1 Then
							intCounted = intCounted + 1
						End If
				Next
		
				Next
		
				If intCounted = UBound(aTableContent2)  Then
						bFlag = True
				Else
						bFlag = False
					Exit For
				End If
		
			End If
		
			Next

		Case "Clear"
		
			Call Fn_Button_Click("Fn_AudMgr_AuditLogOperation", objDialog, "Clear")
			' Case Pending to implement
		
		Case "Export Audit Log"
		
			Call Fn_Button_Click("Fn_AudMgr_AuditLogOperation", objDialog, "Export Audit Log")
			' Case Pending to implement
		
		Case "Cancel"
			objDialog.JavaButton("Clear").SetTOProperty "label", "Cancel"
			bGblFuncRetVal = Fn_SISW_UI_JavaButton_Operations("Fn_AudMgr_AuditLogOperation", "Click", objDialog,"Clear")
			If bGblFuncRetVal = False Then
				Fn_AudMgr_AuditLogOperation = False
				Exit Function
			End If
			' Case Pending to implement

End Select

If bFlag = True Then
	Fn_AudMgr_AuditLogOperation = True
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Audit Log Operation performed  successfully in Function " & sFunctionName)
Else
	Fn_AudMgr_AuditLogOperation = False
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Audit Log Operation not performed Correctly in Function " & sFunctionName)
End If

Set objDialog = Nothing
Set objTableDialog = Nothing

End Function

'##############################################################################################################################################################################################################################################
'###
'###    FUNCTION NAME   :  Fn_AudMgr_AuditDefinitionOperation(sAction, sObjType, sEventType, sLogHandler, sStorageType, sArchiveMedia, sAddObjProperties, sRemoveObjProperties)
'###
'###    DESCRIPTION     :  Perform Create, Modify, Delete, Clear or verify operation in AuditManager dailog
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'###
'###	 HISTORY         :   AUTHOR                 DATE        	VERSION
'###
'###    CREATED BY      :   Mahendra Bhandarkar	  10/06/2010        1.0
'###
'###    REVIWED BY      :   Mohit Khare		 	  11/06/2010	    1.0
'###
'###    MODIFIED BY     :	Ketan Raje			26-Apr-2011		
'###    EXAMPLE         : Case "Create" : Msgbox Fn_AudMgr_AuditDefinitionOperation("Create", "3DMarkup", "Email Sent", "CICO_audithandler", "Database", "", "lsd", "Data:IMAN_based_on")
'###    							Case "Modify" : Msgbox Fn_AudMgr_AuditDefinitionOperation("Modify", "", "", "", "", "", "lsd", "Data:IMAN_based_on")
'###    							Case "Verify" : Msgbox Fn_AudMgr_AuditDefinitionOperation("Verify", "3DMarkup", "Email Not Sent", "CICO_audithandler", "Database", "", "Data:IMAN_based_on", "lsd")
'###    							Case "Delete" : Msgbox Fn_AudMgr_AuditDefinitionOperation("Delete", "", "", "", "", "", "", "")
'###    							Case "Clear" : Msgbox Fn_AudMgr_AuditDefinitionOperation("Clear", "", "", "", "", "", "", "")
'#######################################################################################################################################################################################################################
Public Function Fn_AudMgr_AuditDefinitionOperation(sAction, sObjType, sEventType, sLogHandler, sStorageType, sArchiveMedia, sAddObjProperties, sRemoveObjProperties)
	GBL_FAILED_FUNCTION_NAME="Fn_AudMgr_AuditDefinitionOperation"

	Dim objDialog, aAddObjProperties, aRemoveObjProperties, iCount, iCounter, iRows, iCnt, intCnt,bFlag
	Fn_AudMgr_AuditDefinitionOperation = False
	Set objDialog = Fn_UI_ObjectCreate( "Fn_AudMgr_AuditDefinitionOperation", Window("AuditManagerWindow").JavaWindow("AuditManagerApplet"))
	
	Select Case sAction
	
	Case "Create", "Modify"
			'Select Object Type.
			If sObjType<>"" Then
				Call Fn_List_Select("Fn_AudMgr_AuditDefinitionOperation", objDialog, "ObjectType",sObjType)
			End If
			'Select Event Type.
			If sEventType<>"" Then
				Call Fn_List_Select("Fn_AudMgr_AuditDefinitionOperation", objDialog, "EventType",sEventType)
			End If
			'Select Log Handler.
			If sLogHandler<>"" Then
				Call Fn_List_Select("Fn_AudMgr_AuditDefinitionOperation", objDialog, "LogHandler",sLogHandler)
			End If
			'Select Storage Type.
			If sStorageType<>"" Then
				objDialog.JavaRadioButton("StorageType").SetTOProperty "attached text",sStorageType
				Call Fn_UI_JavaRadioButton_SetON("Fn_AudMgr_AuditDefinitionOperation",objDialog, "StorageType")
			End If
			'Add Object Properties to Audit Definition.
			If Trim(sAddObjProperties) <> "" Then
				aAddObjProperties = Split(sAddObjProperties,":",-1,1)
				For iCount = 0 to Ubound(aAddObjProperties)
'					bFlag=Fn_UI_ListItemExist("Fn_AudMgr_AuditDefinitionOperation", objDialog, "LoggedProperties",aAddObjProperties(iCount))
					'/*Added by Das on 24-June-2013 to verify existence of properties in logged properties list 
					iCnt = objDialog.JavaList("LoggedProperties").GetROProperty("items count")
					For iCounter = 0 To iCnt - 1
						If 	objDialog.JavaList("LoggedProperties").GetItem(iCounter) <> ""	 Then
								If  Trim(cstr(objDialog.JavaList("LoggedProperties").GetItem(iCounter)))=Trim(cstr(aAddObjProperties(iCount)))         Then
										bFlag=True
										Exit For
								Else
									bFlag=False
								End If
						End If
					Next
					'/*End
					If bFlag=False Then
						Call Fn_List_Select("Fn_AudMgr_AuditDefinitionOperation", objDialog, "ObjectProperties",aAddObjProperties(iCount))
						Call Fn_Button_Click("Fn_AudMgr_AuditDefinitionOperation", objDialog, "Add")
					End If
				Next				
			End If
			'Remove Object Properties from Audit Definition.
			If Trim(sRemoveObjProperties) <> "" Then
				aRemoveObjProperties = Split(sRemoveObjProperties,":",-1,1)
				For iCount = 0 to Ubound(aRemoveObjProperties)
					Call Fn_List_Select("Fn_AudMgr_AuditDefinitionOperation", objDialog, "LoggedProperties",aRemoveObjProperties(iCount))
					Call Fn_Button_Click("Fn_AudMgr_AuditDefinitionOperation", objDialog, "Remove")
				Next				
			End If
			If sAction = "Create" Then
				'Click on Create button to create a Audit Definition.
				Call Fn_Button_Click("Fn_AudMgr_AuditDefinitionOperation", objDialog, "Create")
			ElseIf sAction = "Modify" Then
				'Click on Modify button to create a Audit Definition.
				Call Fn_Button_Click("Fn_AudMgr_AuditDefinitionOperation", objDialog, "Modify")
			End If			
			Fn_AudMgr_AuditDefinitionOperation = True
	Case "Delete"
			'Call Fn_Button_Click("Fn_AudMgr_AuditDefinitionOperation", objDialog, "Remove")
			Call Fn_Button_Click("Fn_AudMgr_AuditDefinitionOperation", objDialog, "Delete")
			Fn_AudMgr_AuditDefinitionOperation = True
	Case "Clear"
			'Call Fn_Button_Click("Fn_AudMgr_AuditDefinitionOperation", objDialog, "Remove")
			Call Fn_Button_Click("Fn_AudMgr_AuditDefinitionOperation", objDialog, "Clear")
			Fn_AudMgr_AuditDefinitionOperation = True
	Case "Verify"	
				iCount = 0
				iCounter = 0
				'Verify value of Object Type
				If sObjType<>"" Then
					iCount = iCount + 1
					If Trim(Lcase(objDialog.JavaList("ObjectType").GetROProperty("value"))) = Trim(Lcase(sObjType)) Then
						iCounter = iCounter + 1
					End If
				End If
				'Verify value of Event Type
				If sEventType<>"" Then
					iCount = iCount + 1
					If Trim(Lcase(objDialog.JavaList("EventType").GetROProperty("value"))) = Trim(Lcase(sEventType)) Then
						iCounter = iCounter + 1
					End If
				End If
				'Verify value of Log Handler
				If sLogHandler<>"" Then
					iCount = iCount + 1
					If Trim(Lcase(objDialog.JavaList("LogHandler").GetROProperty("value"))) = Trim(Lcase(sLogHandler)) Then
						iCounter = iCounter + 1
					End If
				End If
				'Verify value of Storage Type
				If sStorageType<>"" Then
					iCount = iCount + 1
					objDialog.JavaRadioButton("StorageType").SetTOProperty "attached text",sStorageType
					If objDialog.JavaRadioButton("StorageType").GetROProperty("value") = 1 Then
						iCounter = iCounter + 1
					End If
				End If
				'Verify Object Properties
				If Trim(sAddObjProperties) <> "" Then
					aAddObjProperties = Split(sAddObjProperties,":",-1,1)
					For intCnt = 0 to Ubound(aAddObjProperties)
						iCount = iCount + 1
						iRows = objDialog.JavaList("ObjectProperties").GetROProperty("items count")
							For iCnt = 0 to iRows-1
									If Trim(Lcase(objDialog.JavaList("ObjectProperties").GetItem(iCnt))) = Trim(Lcase(aAddObjProperties(intCnt))) Then
										iCounter = iCounter + 1
										Exit For
									End If
							Next					
					Next				
				End If				
				'Verify Logged Properties
				If Trim(sRemoveObjProperties) <> "" Then
					aRemoveObjProperties = Split(sRemoveObjProperties,":",-1,1)
					For intCnt = 0 to Ubound(aRemoveObjProperties)
						iCount = iCount + 1
						iRows = objDialog.JavaList("LoggedProperties").GetROProperty("items count")
							For iCnt = 0 to iRows-1
									If Trim(Lcase(objDialog.JavaList("LoggedProperties").GetItem(iCnt))) = Trim(Lcase(aRemoveObjProperties(intCnt))) Then
										iCounter = iCounter + 1
										Exit For
									End If
							Next					
					Next				
				End If
				If iCount = iCounter Then
					Fn_AudMgr_AuditDefinitionOperation = True
				Else
					Fn_AudMgr_AuditDefinitionOperation = False
				End If
	End Select
	
	Set objDialog = Nothing

End Function
'##############################################################################################################################################################################################################################################
'###
'###    FUNCTION NAME   :  Fn_AudMgr_NavTreeOperations()
'###
'###    DESCRIPTION     :  Perform Audit Manager Navigation Tree Operations
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'###
'###	 HISTORY         :   AUTHOR                 DATE        	VERSION
'###
'###    CREATED BY      :   Mahendra Bhandarkar	  10/06/2010        1.0
'###
'###    REVIWED BY      :   Mohit Khare		 	  11/06/2010	    1.0
'###
'###    MODIFIED BY     :	Ketan Raje			  26-Apr-2011		1.0
'###    EXAMPLE         : Fn_AudMgr_NavTreeOperations("Select", "Dataset - Check-Out", "")
'###    							Case "Exist" : Msgbox Fn_AudMgr_NavTreeOperations("Exist", "Audit Definition Objects:AM_ACL - Modify", "")
'#######################################################################################################################################################################################################################
Public Function Fn_AudMgr_NavTreeOperations(sAction, sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_AudMgr_NavTreeOperations"
	Dim objDialog, NodeLists, intNodeCount, intCount

	If Fn_UI_ObjectExist("Fn_AudMgr_NavTreeOperations", Window("AuditManagerWindow")) = False Then
		Call Fn_SetPerspective("Audit Manager")
	End If

	Set objDialog = Fn_UI_ObjectCreate( "Fn_AudMgr_NavTreeOperations", Window("AuditManagerWindow").JavaWindow("AuditManagerApplet"))

	Select Case sAction
	Case "Select"
				Call Fn_JavaTree_Select("Fn_AudMgr_NavTreeOperations", objDialog, "DefObjectTree","Audit Definition Objects:"+sNodeName)
				Fn_AudMgr_NavTreeOperations = True
	Case "Expand"
				Call Fn_JavaTree_Select("Fn_AudMgr_NavTreeOperations", objDialog, "DefObjectTree","Audit Definition Objects:"+sNodeName)
				Call Fn_UI_JavaTree_Expand("Fn_AudMgr_NavTreeOperations",objDialog,"DefObjectTree","Audit Definition Objects:"+sNodeName)
				Fn_AudMgr_NavTreeOperations = True
	Case "Collapse"
				Call Fn_JavaTree_Select("Fn_AudMgr_NavTreeOperations", objDialog, "DefObjectTree","Audit Definition Objects:"+sNodeName)
				Call Fn_UI_JavaTree_Collapse("Fn_AudMgr_NavTreeOperations", objDialog,"DefObjectTree","Audit Definition Objects:"+sNodeName)
				Fn_AudMgr_NavTreeOperations = True		
	Case "Exist"
				intNodeCount = objDialog.JavaTree("DefObjectTree").GetROProperty ("items count") 
				For intCount = 0 to intNodeCount - 1
					sTreeItem = objDialog.JavaTree("DefObjectTree").GetItem(intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_AudMgr_NavTreeOperations = True
						Exit For
					End If
				Next
				If cint(intCount) = cint(intNodeCount) Then
					Fn_AudMgr_NavTreeOperations = False
				End If
	End Select	

Set objDialog = Nothing
End Function




'''''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''''''/$$$$
'''''''''/$$$$   FUNCTION NAME   :    Fn_SISW_AudMgr_SrchResultValidations(sAction,sNodeName,sInfo1,sInfo2,sInfo3)
'''''''''/$$$$
'''''''''/$$$$   DESCRIPTION        : This function Will  Perform  Validation Operations on Search Results for Audit Logs
'''''''''/$$$$
'''''''''/$$$$	PRE-REQUISITERS :  Summary Tab should Be Activated
'''''''''/$$$$
'''''''''/$$$$  PARAMETERS   : 		sAction : Valid Action Name
'''''''''/$$$$										sNodeName : Valid Node Name
'''''''''/$$$$										sInfo1 : For Future Use
'''''''''/$$$$										sInfo2:	For Future Use
'''''''''/$$$$										sInfo3:	For Future Use
'''''''''/$$$$	
'''''''''/$$$$		Return Value : 				True or False
'''''''''/$$$$
'''''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''''''/$$$$
'''''''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
'''''''''/$$$$  
'''''''''/$$$$    CREATED BY     :   SHREYAS          18/09/2012         1.0
'''''''''/$$$$
'''''''''/$$$$    REVIWED BY     :  Shreyas			18/09/2012            1.0
'''''''''/$$$$
''''''''/$$$$		Added [ PopupMenuSelect ] Case : Harshal Tanpure : 11-October-2012 : Teamcenter 10 (20120919.00)
'''''''''/$$$$
'''''''''/$$$$		How To Use :      				Example #1
''''''''/$$$$																
''''''''/$$$$							bReturn=Fn_SISW_MyTc_SrchResultValidations("Select","000118-item001___Check_In_18-Sep-2012" ,"","","")
''''''''/$$$$
''''''''/$$$$													Example #2
''''''''/$$$$							
''''''''/$$$$							bReturn=Fn_SISW_MyTc_SrchResultValidations("Select","000118-item001___Check_Out_18-Sep-2012" ,"","","")
''''''''/$$$$
''''''''/$$$$													Example #3
''''''''/$$$$							
''''''''/$$$$							bReturn=Fn_SISW_AudMgr_SrchResultValidations("MultipleOccuranceSelect","TestUser122___Create_25-Sep-2012" ,"2","","")
''''''''/$$$$
''''''''/$$$$													Example #4
''''''''/$$$$
''''''''/$$$$						bReturn = Fn_SISW_AudMgr_SrchResultValidations("PopupMenuSelect","000018-TestItem___Check_Out_11-Oct-2012","1","View Properties","")
''''''''/$$$$
''''''''/$$$$
'''''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public function Fn_SISW_AudMgr_SrchResultValidations(sAction,sNodeName,sInfo1,sInfo2,sInfo3)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_AudMgr_SrchResultValidations"
   On Error Resume Next 
   Fn_SISW_AudMgr_SrchResultValidations=false
   Dim objTree,sValue,sIndex,iCount,aSelNodes,aNewNames(),NodeCount(),ArraySize,iCnt
   Dim arrMenuList , iCounter
   Set objTree=JavaWindow("DefaultWindow").JavaTree("SearchResultTree")

Select Case sAction

				Case "Select"
					  sValue=Fn_MyTc_SrchResltTreeOperation ("GetChildrenList", "", "")
					  aNodes=split(sValue,"~",-1,1)
					For iCount=0 to uBound(aNodes)
						If instr(aNodes(iCount),sNodeName)>0 Then
							sIndex=iCount
							sNodePath="#0:#"+cstr(sIndex)
							objTree.Select sNodePath
							Fn_SISW_AudMgr_SrchResultValidations=true
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node ["+sNodeName+"]")
							Exit for
						End If
				 Next

				 If iCount-1= uBound(aNodes) Then
							Fn_SISW_AudMgr_SrchResultValidations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To Select Node ["+sNodeName+"]")
				 End If
				 ' Migrated MultipleOccuranceSelect Case from Tags/TC10_1 to Mainline by Nilesh Gadekar on 8-Jan-2013
				 Case "MultipleOccuranceSelect"
					  sValue=Fn_MyTc_SrchResltTreeOperation ("GetChildrenList", "", "")

					  aNodes=split(sValue,"~",-1,1)
					  iCnt=0

					For iCount=0 to uBound(aNodes)
						If instr(aNodes(iCount),sNodeName)>0 Then
							iCnt=iCnt+1
							ArraySize=iCnt-1
							ReDim preserve aNewNames(ArraySize)
							aNewNames(ArraySize)=iCount
						End if
					Next			

					If iCnt=0 or uBound(aNewNames)<0 Then
							Fn_SISW_AudMgr_SrchResultValidations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To Select Node ["+sNodeName+"]")
							Exit Function
					End If
					sNodePath="#0:#"+cstr(aNewNames( cint(sInfo1)-1))
					objTree.Select sNodePath
					If err.number<0 Then
							Fn_SISW_AudMgr_SrchResultValidations=False
							Exit Function
					Else
							Fn_SISW_AudMgr_SrchResultValidations=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node ["+sNodeName+"]  at occurance ["+sInfo1+"]")
					End If
				' Migrated PopupMenuSelect Case from Tags/TC10_1 to Mainline by Nilesh Gadekar on 8-Jan-2013
				' Added [ PopupMenuSelect ] Case : Harshal Tanpure : 11-October-2012 : Teamcenter 10 (20120919.00)

				Case "PopupMenuSelect"
					  sValue=Fn_MyTc_SrchResltTreeOperation ("GetChildrenList", "", "")

					  aNodes=split(sValue,"~",-1,1)
					  iCnt=0

					For iCount=0 to uBound(aNodes)
						If instr(aNodes(iCount),sNodeName)>0 Then
							iCnt=iCnt+1
							ArraySize=iCnt-1
							ReDim preserve aNewNames(ArraySize)
							aNewNames(ArraySize)=iCount
						End if
					Next			

					If iCnt=0 or uBound(aNewNames)<0 Then
							Fn_SISW_AudMgr_SrchResultValidations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To Select Node ["+sNodeName+"]")
							Exit Function
					End If
					sNodePath="#0:#"+cstr(aNewNames( cint(sInfo1)-1))
					objTree.OpenContextMenu(sNodePath)

					arrMenuList = split(sMenu, ":",-1,1)
					iCounter = Ubound(arrMenuList)

					wait(2)
					Select Case iCounter
					Case "0"								
						JavaWindow("DefaultWindow").JavaMenu("label:="&arrMenuList(0)&"","index:=0" ).Select
                        Fn_SISW_AudMgr_SrchResultValidations =True
					Case "1"								
						JavaWindow("DefaultWindow").JavaMenu("label:="&arrMenuList(0)&"","index:=0").JavaMenu("label:="&arrMenuList(1)&"","index:=0").Select
                         Fn_SISW_AudMgr_SrchResultValidations =True        						
					Case "2"
						 JavaWindow("DefaultWindow").JavaMenu("label:="&arrMenuList(0)&"","index:=0").JavaMenu("label:="&arrMenuList(1)&"","index:=0").JavaMenu("label:="&arrMenuList(2)&"","index:=0").Select
                         Fn_SISW_AudMgr_SrchResultValidations =True
					Case Else
						Fn_MyTc_SrchResltTreeOperation = FALSE
						Call Fn_WriteLogFile("Fn_MyTc_SrchResltTreeOperation", 1, Err.Number,"FAIL:Wrong Parameter for Context Menu" )
						Exit Function
					End Select

					If err.number<0 Then
							Fn_SISW_AudMgr_SrchResultValidations=False
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fail to Select Node ["+sNodeName+"]  and do RMB and select ["+sMenu+" ] ") 
							Exit Function
					Else
							Fn_SISW_AudMgr_SrchResultValidations=True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Selected Node ["+sNodeName+"]  and done RMB and selected ["+sMenu+"]")
					End If
					
		End Select

End function
