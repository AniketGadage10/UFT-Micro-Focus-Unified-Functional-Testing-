 Option Explicit
'*********************************************************	Function List		***********************************************************************
'1. Fn_SISW_Pref_GetObject
'2. Fn_SISW_Pref_PreferenceOperations
'3. Fn_SISW_Pref_Search_Operation
'4. Fn_SISW_Pref_WinPrefOperation
'5. Fn_SISW_Pref_Search_Operation_WithCategory
'6. Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog
'7. Fn_SISW_Pref_RptChngs_Operation
'8. Fn_SISW_Pref_Organization_Operation
'9. Fn_SISW_Pref_MultiValue_Operation
'10. Fn_SISW_Pref_Search_CreateOperation
'11. Fn_SISW_Pref_Search_Lock_Unlock
'12. Fn_SISW_Pref_SearchPreferenceOperations
'13. Fn_SISW_Pref_ImportFromOrganization
'*********************************************************	Function List		***********************************************************************

''****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_Pref_GetObject
'
''Description		    :  	Function to get specified Object hierarchy.

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_Pref_GetObject("IndexOptions")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Sushma Pagare		 15-July-2013		1.0				
'   -----------------------------------------------------------------------------------------------------------------------------------
Function Fn_SISW_Pref_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\Preference.xml"
	Set Fn_SISW_Pref_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function

'*********************************************************		Fn_SISW_Pref_PreferenceOperations		***********************************************************************
'
'	Function Name			:			Fn_SISW_Pref_PreferenceOperations
'
'	Description			 	:		 	This Function is used for following :-
'										1. Create 	[---Done---]
'										2. Modify 	[---Done---]
'										3. Delete 	[Not Done]
'										4. Import 	[Not Done]
'										5. Export 	[Not Done]
'										6. Lock  	Prerequisite : Index Page should be present in Edit->Options
'										7. Unlock	Prerequisite : Index Page should be present in Edit->Options
'												
'
'	Parameters				:	 		1. StrAction:  FunctionDose Example:Modify
'										2. StrPrefName: Prefrences Name  
'										3. StrDesc: Value To Be modified
'										*******Other Passed for Future Use (All these parameters will be passed as blank in double quote**********	
'										4.StrScope:
'										5.StrCategory
'										6.StrValue
'										7.StrType:
'										8.BlnMultiVal:
'										9.StrImportFilePath:
'										10. StrImportMd:
'										11.StrImportPrefOpt:
'										12. StrExportFileName:
'										******************************************************
'											
'	Return Value			: 			The String which represents the result : "True" or "False" with the reason
'
'	Pre-requisite			:		 	User should logged in to the teamcenter with DBA Privilledge
'
'	Examples				:			Call Fn_SISW_Pref_PreferenceOperations("Modify","Harshal_Prefrences_ForTest","OkTested","","","","","","","","","")
'										Call Fn_SISW_Pref_PreferenceOperations("Import","","","User","Classification","","","","E:\testfile.xml","Automatic","Skip the preference","")
'										Call Fn_SISW_Pref_PreferenceOperations("VerifyPreferenceValue","ADA_license_administration_privilege","","","","ITAR_ADMIN","","","","","","")		
'										Call Fn_SISW_Pref_PreferenceOperations("CreateNewPreferenceInstance","Pref_MyTc","","","","False","","","","","","")
'										Call Fn_SISW_Pref_PreferenceOperations("Lock","","","","","","","","","","","")
'										Call Fn_SISW_Pref_PreferenceOperations("Unlock","","","","","","","","","","","")										
'										Call Fn_SISW_Pref_PreferenceOperations("Refresh","","","","","","","","","","","")										
'	History:
'
'	Developer Name				Date			Rev. No.						Changes Done																				Reviewer			Reviewed Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Harshal					24-Mar-2010		1.0																																Santosh				24-Mar-10				
'	Swapna	Ghatge			21-May-2010		1.1			Added Create Case 																									Mohit				21-May-2010	
'	Pranav					09-July-2010				Added Case "VerifyPreferenceValue"																					Ketan				09-July2010
'	ANJALI					30-Dec-2010		1.2		 	Added Case ""VerifyPreferenceByCategory"																			SHREYAS				30-Dec-2010
'	Saurabh					31st-Dec-2010				Modified the Case "Lock" and "Unlock"																				Pritish 			31st-Dec-2010
'	Saurabh					31st-Dec-2010				Added Case VerifyPreferenceMultipleValue,Export   																	Harshal			 	27st-Jan-2011
'	Sachin					01-June-2012				Modified Case "VerifyPreferenceValue"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh				12-June-2012	2.0			Modified function according to TC10.0 UI changes
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Mohammad				15-June-2012	2.0			Added new cases "CreatePrefrenceDefinitionCollision", "VerifyMsgCreatePreferenceDialog"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh				21-June-2012	2.0			Added new case CreateNewPreferenceInstance
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pooja					28-June-2012	2.0			Added new Object hierarchy																							Koustubh
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Nilesh					16-July-2012	2.1			Modified function for Unlock functionality		
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ikhlaque				17-July-2012	2.2			Modified case VerifyPreferenceWithScope
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Madhura Puranik			07-Mar-2016		2.2			Added cases "VerifyErrMsgAfterModify", "ModifyWithoutClose", "VerifyErrorMsgAfterLock", "VerifyErrMsgAfterUnlock"	[Tc1122:2016021000:07Mar2016:AnkitN:NewDevelopment]
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh				25-Apr-2016		2.2			Added cases "Refresh"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Pref_PreferenceOperations(StrAction,StrPrefName,StrDesc,StrScope,StrCategory,StrValue,StrType,BlnMultiVal,StrImportFilePath,StrImportMd,StrImportPrefOpt,StrExportFileName)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_PreferenceOperations"
	Dim objPrefOper,objDialog,objSelectType, objSelectType1, intNoOfObjects, intNoOfObjects1, iCounter1
	Dim iItemCnt,iCnt,sItemName,bFlag,iCounter,ArrValue,bReturn,sprefsplt,iCnt2,sListvalue,iFlag,iFlag2,sListcount,bUnlock,aStrAction
	Dim location,bClose,sProtectionScope, sPerfValue
	Dim objDelete, objTcDefaultApplet, objDialog2,aStrPrefName,aInfo, dicErrorInfo,aPrefValue
	
	Fn_SISW_Pref_PreferenceOperations = False
   
	'Menu operation function called to select Option from Edit  Menu.
	If Fn_SISW_Pref_GetObject("IndexOptions").Exist(5) = False And Fn_SISW_Pref_GetObject("IndexOptions2").Exist(1) = False And Fn_SISW_Pref_GetObject("IndexOptions3").Exist(1) = False Then
		Call Fn_MenuOperation("Select","Edit:Options...")
		Call Fn_ReadyStatusSync(1) 
	End If
	If Fn_SISW_Pref_GetObject("IndexOptions").Exist(6) = True Then
		Set objPrefOper =   Fn_SISW_Pref_GetObject("IndexOptions")
	ElseIf Fn_SISW_Pref_GetObject("IndexOptions2").Exist(1) = True Then
		Set objPrefOper =   Fn_SISW_Pref_GetObject("IndexOptions2")
	ElseIf Fn_SISW_Pref_GetObject("IndexOptions3").Exist(1) = True Then
		Set objPrefOper =   Fn_SISW_Pref_GetObject("IndexOptions3")
	Else 
        Set objPrefOper = Nothing
		Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: To Find Object In Object Repository ")		
		Exit Function
	End If

	bFlag = False
	'Synchronization for Ready state
	Call Fn_ReadyStatusSync(1) 	
	' clicking on Filter / Index
	If Not Fn_SISW_UI_Object_Operations("Fn_SISW_Pref_PreferenceOperations", "Exist", Fn_SISW_Pref_GetObject("PreferenceDefinitionCollision"), SISW_MICRO_TIMEOUT) Then
		Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"Index",0,0,"LEFT")
		Call Fn_ResizeWindow("Resize","700", "800", objPrefOper)	
	End If
	'Added by Nilesh on 16 -July-2012
    bUnlock=True
	If Instr(StrAction,"~")>0Then
		aStrAction=Split(StrAction,"~",-1,1)
		If Ubound( aStrAction)<>0 Then
			StrAction=aStrAction(0)
			bUnlock=aStrAction(1)
		End If
	End If

		'Added by Nilesh 13-Sep-12 
	bClose=True
	If Instr(StrPrefName,"~")>0Then
		aStrPrefName=Split(StrPrefName,"~",-1,1)
		If Ubound( aStrPrefName)<>0 Then
			StrPrefName=aStrPrefName(0)
			bClose=aStrPrefName(1)
		End If
	End If

	Select Case StrAction
		Case "Refresh"
			bFlag = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Pref_PreferenceOperations",  "Click", objPrefOper, "Refresh")
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Failed to click on Refresh button")
				exit function
			End If
			Call Fn_ReadyStatusSync(2)
			bFlag = Fn_SISW_UI_JavaButton_Operations("Fn_SISW_Pref_PreferenceOperations",  "Click", objPrefOper, "Close")
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Failed to click on Close")
				exit function
			End If
			Fn_SISW_Pref_PreferenceOperations = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case  "Create", "CreatePrefrenceDefinitionCollision", "VerifyMsgCreatePreferenceDialog" ,"CreateCategory",  "SelectCategory", "VerifyPreference",  "VerifyPreferenceByCategory", "VerifyLock" , "Import"' non details tab
				' do nothing
		Case  "VerifyPreferenceWithScope", "DeletePreferenceWithScope", "ModifyPreferenceWithScope" , "Modify", "DeletePreference"
				'Imp Note
				objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Definition"
				If objPrefOper.JavaStaticText("BottomLink").Exist(5) = False Then
'					objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Details"
					objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"
				End If
				Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"BottomLink",10, 10,"LEFT")
				
				Fn_SISW_Pref_PreferenceOperations = false

				objPrefOper.JavaEdit("SrchPrefName").Set ""
				objPrefOper.JavaEdit("SrchPrefName").Type StrPrefName
				wait 1
				'----------------------End------------------------------------------------------------------------------------------
				
'				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"SrchPrefName",StrPrefName)
				Call Fn_ReadyStatusSync(1)
				If StrCategory <> "" Then
					Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					objSelectType("label").value = StrCategory
					Call Fn_Button_Click( "Fn_SISW_Pref_PreferenceOperations", objPrefOper,"SrchPrefCategoryBtn" )
					wait 1
					Set objDialog =objPrefOper.ChildObjects(objSelectType)
					objDialog(0).Click 5, 5, "LEFT"
				End If
				Wait(3)
				'Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Find")

				'Select Preference
				iItemCnt = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
				For iCnt = 0 to iItemCnt - 1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  StrPrefName Then
						wait 2
						objPrefOper.JavaTable("PreferencesListTable").Click 0,0,"LEFT"
						wait 1
						If trim(StrScope) <> "" Then
							If lcase(trim(objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Location"))) = lcase(trim(StrScope)) Then
								objPrefOper.JavaTable("PreferencesListTable").ClickCell iCnt,"Name"
								bFlag = True
								Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: [" + sItemName+"] Preference With "+CStr(StrScope)+" Scope Found in the Preference List  ")
								Exit for
							End If
						Else
							bFlag = True
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: [" + sItemName+"] Preference Found in the Preference List  ")
							objPrefOper.JavaTable("PreferencesListTable").ClickCell iCnt,"Name"
							exit for
						End If
					End If
				Next
				wait 3
				If bFlag = False Then
					Fn_SISW_Pref_PreferenceOperations = False
					Exit Function
				End If

		Case Else '  details tab and select pref
			'	clikcing on "Details" tab
'				objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Details"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"
	
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"BottomLink",10, 10,"LEFT")
			If StrPrefName<>"" Then  'Added by Nilesh on 16-July-2012
				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"SrchPrefName", StrPrefName)
				wait 2
				objPrefOper.JavaTable("PreferencesListTable").Click 0,0,"LEFT"
				wait 1
				'Select Preference
				iItemCnt = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
				For iCnt = 0 to iItemCnt - 1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  StrPrefName Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
						bFlag = True
						wait 2
						Exit for
					End If
				Next
				If bFlag = False Then
					Fn_SISW_Pref_PreferenceOperations = False
					Exit Function
				End If
			End If
	End Select

	Select Case StrAction
		Case "Modify", "VerifyErrMsgAfterModify", "ModifyWithoutClose"				'[Tc1122:2016021000:07Mar2016:MadhuraP:NewDevelopment] - Added Case to verify Error message after modification
			If bFlag Then
				'click on Edit Button
				If objPrefOper.JavaButton("Edit").Exist(3) Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Edit")
				End If
		
				If StrScope <> "" Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "PrefScopeDrpDwn")
					iCounter1 = 0
					If objPrefOper.JavaStaticText("ValueLocation").Exist(1) Then
						location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
						If lcase(trim(StrScope)) = location Then
							iCounter1 = 1
						End If
					End If
					wait 4
					' add code to select static text of scope
					Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					objSelectType("label").value = StrScope
					Set objDialog = objPrefOper.ChildObjects(objSelectType)
					objDialog(iCounter1).Click 5, 5, "LEFT"
				End If
		
				If Instr(StrValue,"~") > 0 Then
					aPrefValue = split(StrValue,"~",-1,1)
					For iCnt=0 to Ubound(aPrefValue)
						If Fn_UI_ListItemExist("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "PrefMultiValList",aPrefValue(iCnt)) <> True Then
							Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CurrentValues",aPrefValue(iCnt))
							Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewMultiValPrefAddValue")	
						End If
					Next
				Else
					'Set value
					objPrefOper.JavaEdit("CurrentValues").Set ""
					objPrefOper.JavaEdit("CurrentValues").Type StrValue
				End If
				Wait(3)
		
				' Click on Save Button
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Save")
				Fn_SISW_Pref_PreferenceOperations = True
			Else
				Fn_SISW_Pref_PreferenceOperations = False
			End If
			'-------------------------------------------------------------------------------------------------------------------------------------------			
			'Condition to verify Error message is correct or not .				'[Tc1122:2016021000:07Mar2016:MadhuraP:NewDevelopment]
			If StrAction = "VerifyErrMsgAfterModify" Then					
				If StrExportFileName <> "" Then
					aInfo = Split(StrExportFileName,"~")
					Set dicErrorInfo = CreateObject("Scripting.Dictionary")
					With dicErrorInfo	
						.Add "Title", aInfo(0)
						.Add "Message", aInfo(1)
						.Add "Button", aInfo(2)
					End with
					bReturn = Fn_SISW_ErrorVerify(dicErrorInfo)
					If bReturn = False Then
						Fn_SISW_Pref_PreferenceOperations = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error message is not verified.")
					End If
					Set dicErrorInfo = Nothing
				End If
			End If
			'-------------------------------------------------------------------------------------------------------------------------------------------
			'Condition to handle close button 									'[Tc1122:2016021000:07Mar2016:MadhuraP:NewDevelopment]
			If StrAction <> "ModifyWithoutClose" Then
				' Click on Close Button
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")				
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CreatePrefrenceDefinitionCollision"   
			Dim objPrefDefColl
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "CreateNewPreference")

			'Set value in Name Edit  box
			Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"Name",StrPrefName)
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "PrefScopeDrpDwn")
			Set objPrefDefColl=Fn_SISW_Pref_GetObject("PreferenceDefinitionCollision")
			Wait (5)
			If objPrefDefColl.Exist(2)  Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : The [Preference Definition Collision] Exist")
				objPrefDefColl.click 0,0
				Fn_SISW_Pref_PreferenceOperations = True
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : The [Preference Definition Collision]  Does Not Exist")
				Fn_SISW_Pref_PreferenceOperations = False
			End If
			Set objPrefDefColl=Nothing
				
		Case "VerifyMsgCreatePreferenceDialog"  
			'Declaring Variables
			Dim objPrefDefCollision, aPara, sErrMsg,sAttachedtxt
			sAttachedtxt=StrDesc
			'Creating Object of [ PreferenceDefinitionCollision ] dialog
			Set objPrefDefCollision=Fn_SISW_Pref_GetObject("PreferenceDefinitionCollision")
			'Checking existing of [ PreferenceDefinitionCollision ] dialog
			If objPrefDefCollision.Exist(6) Then
				If sAttachedtxt<>"" Then
					aPara=Split(sAttachedtxt,"~")
					'Taking current error message display on [ PreferenceDefinitionCollision ] dialog
					sErrMsg=objPrefDefCollision.JavaObject("MLabel").GetROProperty("text")
					If instr(1,sErrMsg, aPara(0)) > 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully verified error message [" + aPara(0) + "] display on [ Preference Definition Collision ] dialog")
						Fn_SISW_Pref_PreferenceOperations=true
					else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : error message [" + aPara(0) + "] not displayed on [ Preference Definition Collision ] dialog")
					End if
					If ubound(aPara)=1 Then
						objPrefDefCollision.click 0,0
						Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objPrefDefCollision,aPara(1))
					else
						objPrefDefCollision.click 0,0
						'Clicking on Switch button
						Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objPrefDefCollision, "Switch")
					End If
				else
					'Clicking on Switch button
					objPrefDefCollision.click 0,0
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objPrefDefCollision, "Switch")
					Fn_SISW_Pref_PreferenceOperations=true
				End If
			else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : [ Preference Definition Collision ] dialog not exist")
			End If
			' Click on Close Button
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
			'Releasing Object of [ PreferenceDefinitionCollision ] dialog
			Set objPrefDefCollision=nothing

		Case "Create", "CreateNewPreferenceInstance"   
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "CreateNewPreference")
			'Added by Nilesh 0n 19-Jul-2012
			Call Fn_ReadyStatusSync(1)
			If objPrefOper.JavaEdit("Name").GetRoProperty("editable")=0 And StrAction="Create" Then
					Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"Search",0,0,"LEFT")
					Call Fn_ReadyStatusSync(1)
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "CreateNewPreference")
					Call Fn_ReadyStatusSync(1)
			End If
			'End

			If StrAction <> "CreateNewPreferenceInstance" Then
				'Set value in Name Edit  box
				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"Name",StrPrefName)
			End If
			
			'Set value in Description Edit  box
			If StrScope <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "PrefScopeDrpDwn")
				iCounter1 = 0
				If objPrefOper.JavaStaticText("ValueLocation").Exist(1) Then
					location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
					If lcase(trim(StrScope)) = location Then
						iCounter1 = 1
					End If
				End If
				wait 2
				' add code to select static text of scope
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = StrScope
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(iCounter1).Click 5, 5, "LEFT"
			End If

			If StrAction <> "CreateNewPreferenceInstance" Then
				If Trim(StrDesc) = "" Then 
					StrDesc = "Description"
				End If
				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"Description", StrDesc)
			End If
			

			'Select category
			If StrAction <> "CreateNewPreferenceInstance" Then
				If Trim(StrCategory) = "" Then
					StrCategory = "General"
				End If
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "PrefCategoryDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = StrCategory
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			'Select Multiple value
			If BlnMultiVal <> "" Then
				If lcase(BlnMultiVal) = "off" OR lcase(cstr(BlnMultiVal)) = "false" Then
					BlnMultiVal = "Single"
				Else
					BlnMultiVal = "Multiple"
				End If
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "PrefMultipleDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = BlnMultiVal
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			'Select Type
			If Trim(StrType) <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewPrefTypeDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = StrType
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			If Trim(StrValue) <> "" Then
				If Lcase(StrValue) = "blank" Then
					Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CurrentValues","")
					If  objPrefOper.JavaButton("NewMultiValPrefAddValue").Exist Then
						Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewMultiValPrefAddValue")
				    End If
				ElseIf instr(StrValue,":") > 0 Then
					ArrValue = Split(StrValue,":",-1)
					For iCnt=0 to Ubound(ArrValue)
						Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CurrentValues",ArrValue(iCnt))
						wait 1
						Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewMultiValPrefAddValue")
						wait 1								
					Next		
				Else
					Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CurrentValues",StrValue)
					'Added by pritam
					If  objPrefOper.JavaButton("NewMultiValPrefAddValue").Exist Then
						Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewMultiValPrefAddValue")
				    End If
				End If
			End If
					   
			Do
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Save")
				bFlag = Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog("Preference already exists for the scope. You can modify the value(s) using the Details tab.")
				If bFlag = True Then
					If  objPrefOper.JavaButton("Cancel").Exist = True Then
						Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Cancel")
						Exit Do
					End If
				End If
				'	Loop Until JavaWindow("My Teamcenter - Teamcenter").JavaWindow("My Teamcenter").JavaDialog("Options").JavaButton("Create").GetROProperty("enabled") = "1"
			Loop Until objPrefOper.Exist = True
			wait(5)
'				Added by Nilesh on 13-Sep-2012
              If bClose<>"False" Then
				If  objPrefOper.JavaButton("Cancel").Exist = True Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Cancel")
				End If
				If  objPrefOper.JavaButton("Close").Exist = True Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
				End If
			End If	


			'Log for Success
			If bFlag = True Then				   
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Creation of New Prefrence Operation failed ")				   
				Fn_SISW_Pref_PreferenceOperations = False				   				   
			Else
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Creation of New Prefrence Operation is Done successfully ")
				Fn_SISW_Pref_PreferenceOperations = True	
			End If

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "CreateCategory"
			'Setting TO property to javaStatic text
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Category"
			'Clicking on java Static text
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"BottomLink",10, 10,"LEFT")
			'Imp Note
			'StrPrefName use As Parent Catogero in this case
			If StrPrefName<>"" Then
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = StrCategory
				Call Fn_Button_Click( "Fn_SISW_Pref_PreferenceOperations", objPrefOper ,"CreatePrefCategoryBtn" )
				wait 2
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CategoryCreateCategory",StrCategory)

			'Clicking on Create button to create new category
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Create")

			If  objPrefOper.JavaDialog("Information").Exist(5) Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper.JavaDialog("Information"), "OK")
			End If

			If  objPrefOper.JavaButton("Close").Exist Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
			End If

			If  objPrefOper.JavaButton("Cancel").Exist Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Cancel")
			End If

			Fn_SISW_Pref_PreferenceOperations = True

			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "SelectCategory"
'				Call Fn_Button_Click( "Fn_SISW_Pref_PreferenceOperations", objPrefOper,"SrchPrefCategoryBtn" )
'				bReturn = objPrefOper.JavaButton("SrchPrefCategoryBtn").GetROProperty("focused")
'				If bReturn = 0 Then
'					objPrefOper.JavaButton("SrchPrefCategoryBtn").PressKey micTab
'				End If

'				objPrefOper.JavaEdit("SrchCategoryName").Set StrCategory
'				Set objSelectType = description.Create()
'				objSelectType("Class Name").value = "JavaStaticText"					
'				objSelectType("label").value = StrCategory				
'				Set  intNoOfObjects = objPrefOper.ChildObjects(objSelectType)
''				For  iCounter = 0 to intNoOfObjects.count-1
''					 If intNoOfObjects(iCounter).getROProperty("label") = StrCategory Then
'						 wait 2
'						intNoOfObjects(0).Click 77,8,"LEFT"
''						intNoOfObjects(iCounter).Click 77,8,"LEFT"
''						intNoOfObjects(iCounter).highlight
'						bFlag = True
''						Exit for
''					 End If
''				Next
			bFlag = Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"SrchCategoryName",StrCategory)
			If bFlag = False Then
				Fn_SISW_Pref_PreferenceOperations = False
				Exit Function
			Else
				Fn_SISW_Pref_PreferenceOperations  = True
			End If

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyPreference"
			'Imp Note
			'StrPrefName use As Sub Catogero in this case
			Fn_SISW_Pref_PreferenceOperations = false
			objPrefOper.JavaEdit("SrchPrefName").Set ""
			objPrefOper.JavaEdit("SrchPrefName").Type StrPrefName
				
'				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"SrchPrefName",StrPrefName)
			Call Fn_ReadyStatusSync(1)
			If StrCategory <> "" Then
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = StrCategory
				Call Fn_Button_Click( "Fn_SISW_Pref_PreferenceOperations", objPrefOper,"SrchPrefCategoryBtn" )
				wait 1
				Set objDialog =objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			'Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Find")
			Call Fn_ReadyStatusSync(1)
			'Select Preference
			iItemCnt = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
			wait 2
			objPrefOper.JavaTable("PreferencesListTable").Click 0,0,"LEFT"
			wait 1
			For iCnt = 0 to iItemCnt - 1
				If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  StrPrefName Then
					objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
					wait 2
					Fn_SISW_Pref_PreferenceOperations = True
					Exit for
				End If
			Next

			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		 Case "VerifyPreferenceValue"
				Fn_SISW_Pref_PreferenceOperations = False
				If bFlag=True Then
					'sPerfValue = objPrefOper.JavaEdit("CurrentValues").GetROProperty ("value")
					sPerfValue=Fn_UI_Object_GetROProperty("ROProperty_Perf_Value",objPrefOper.JavaEdit("CurrentValues"),"value")
					If LCase(sPerfValue) = LCase(StrValue) Then
						Fn_SISW_Pref_PreferenceOperations = True
					End If
				End If
	
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyPreferenceMultipleValue", "VerifyPreferenceMultipleValueExt"
				iFlag=0
				iFlag2=0
				'Set the preference name in the Preference name box
				Fn_SISW_Pref_PreferenceOperations = false
				If bFlag=True Then
					If StrAction = "VerifyPreferenceMultipleValueExt" Then
						sprefsplt=Split(StrValue,"~",-1,1)
					Else
						sprefsplt=Split(StrValue,":",-1,1)
					End If
					For iCnt2=0 to UBound(sprefsplt)
						If Fn_UI_ListItemExist("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "PrefMultiValList", sprefsplt(iCnt2)) Then
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Successfully found  [ " & sprefsplt(iCnt2) & " ] in the list  ")
							Fn_SISW_Pref_PreferenceOperations = true
						Else
							Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Fail: Fail to found [ " & sprefsplt(iCnt2) & " ] in the list  ")
							Fn_SISW_Pref_PreferenceOperations = false
							Exit for
						End If
					Next
				End If
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyPreferenceWithScope"

			If StrScope<>"" Then
				If objPrefOper.JavaStaticText("ValueLocation").Exist(5) Then
					sProtectionScope=objPrefOper.JavaStaticText("ValueLocation").GetRoProperty("label")
					If instr(1,sProtectionScope,StrScope) > 0 Then
						bFlag=True
					Else
						bFlag=False
					End If
				Else
					bFlag=False
				End If
			End If
			Fn_SISW_Pref_PreferenceOperations = bFlag 
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "DeletePreferenceWithScope"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"BottomLink",10, 10,"LEFT")
			
			'Imp Note
			If bFlag Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "DeletePreference")
				Set objDelete =  Fn_SISW_Pref_GetObject("DeletePreference" )
				If objDelete.Exist =True Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objDelete, "Yes")	
					Fn_SISW_Pref_PreferenceOperations = True
				End If 
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: [" + sItemName+"] Preference With "+CStr(StrScope)+" Scope Found in the Preference List  ")
			End If
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
			Set objDelete = Nothing
                  ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		  '***** Added by Anjali *******
		Case "VerifyPreferenceByCategory"
			Fn_SISW_Pref_PreferenceOperations = False
			If StrCategory <> "" Then
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = StrCategory
				Call Fn_Button_Click( "Fn_SISW_Pref_PreferenceOperations", objPrefOper,"SrchPrefCategoryBtn" )
				wait 1
				Set objDialog =objPrefOper.ChildObjects(objSelectType)
				If objDialog(0).Exist(5) Then
					objDialog(0).Click 5, 5, "LEFT"
					Fn_SISW_Pref_PreferenceOperations = True
				End If
			End If
                ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Delete", "DeletePreference"
			' Case was not implemented before.. - Koustubh [13-Jun-2012 ]
			If bFlag Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "DeletePreference")
				Set objDelete =  Fn_SISW_Pref_GetObject("DeletePreference" )
				If objDelete.Exist =True Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objDelete, "Yes")	
					Fn_SISW_Pref_PreferenceOperations = True
				End If 
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: [" + sItemName+"] Preference With "+CStr(StrScope)+" Scope Found in the Preference List  ")
			End If
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
			Set objDelete = Nothing
                ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Import"
	'			 objPrefOper.JavaStaticText("BottomLink").SetTOProperty "label", "New"
	'			objPrefOper.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
				objPrefOper.JavaStaticText("BottomLink").SetTOProperty "label", "Import"
				objPrefOper.JavaStaticText("BottomLink").Click 10, 10,"LEFT"
				
				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"ImportFileName",StrImportFilePath)
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "BrowseFile")
				Set objDialog2 =Fn_SISW_Pref_GetObject("ImportPreferences")
				IF objDialog2.Exist = True Then
					Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objDialog2,"FileName",StrImportFilePath)
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objDialog2, "Import")
				End If
				
'				'Select Scope
'				If Trim(StrScope) <> "" Then
'				  objPrefOper.JavaRadioButton("NewPrefScope").SetTOProperty "attached text", StrScope
'				  objPrefOper.JavaRadioButton("NewPrefScope").Set "ON"
'				End If
				If StrScope <> "" Then
'					objPrefOper.JavaEdit("ToLocation").Type StrScope
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "ToLocationDrpDwn")
					wait 2
					' add code to select static text of scope
					Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					Set objDialog = objPrefOper.ChildObjects(objSelectType)
					bFlag = false
					For iCounter = 0 to objDialog.count-1
						If instr(lcase(objDialog(iCounter).GetROProperty("label")),lcase(StrScope)) > 0 Then
							wait 1
							On error resume next
							objDialog(iCounter).Click 1, 1
							bFlag = True
							Err.Number = 0
						End If
					Next
					If bFlag = False Then
						exit function
					End If
				End If
'				
'			   'Select category
'				If Trim(StrCategory) <> "" Then
'				   Fn_SISW_Pref_GetObject("IndexOptions").JavaButton("ImportCategoryDrpDwn").Click
'				   Set objSelectType=description.Create()
'				   objSelectType("Class Name").value = "JavaStaticText"
'				   objSelectType("label").value = StrCategory
'				   Set objDialog = Fn_SISW_Pref_GetObject("IndexOptions").ChildObjects(objSelectType)
'				   objDialog(0).Click 5, 5, "LEFT"
'				End If

				'Select Import Mode
				If Trim(StrImportMd) <> "" Then
					objPrefOper.JavaRadioButton("ImportMode").SetTOProperty "attached text", StrImportMd & " Import"
					objPrefOper.JavaRadioButton("ImportMode").Set "ON"
				End If
				
				If Trim(StrImportMd) <> "" AND Trim(StrImportMd) = "Automatic" Then
					If Trim(StrImportPrefOpt) <> "" Then
						objPrefOper.JavaRadioButton("OverridePrefMode").SetTOProperty "attached text", StrImportPrefOpt
						objPrefOper.JavaRadioButton("OverridePrefMode").Set "ON"
					End If
					
					bFlag = Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Import")
				Else
					bFlag = Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "DryRun")
				End If
				If objPrefOper.JavaDialog("Information").Exist(10) Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper.JavaDialog("Information"), "OK")
				End If
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
				If bFlag=True Then
					Fn_SISW_Pref_PreferenceOperations = True
				Else
					Fn_SISW_Pref_PreferenceOperations = False
					Exit Function
				End If
				   
                ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Export"
				Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"Export",0,0,"LEFT")

				' selecting From Location : Site
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "FromLocationDrpDwn")
				wait 2
				' add code to select static text of scope
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				bFlag = false
				For iCounter = 0 to objDialog.count-1
					If instr(lcase(objDialog(iCounter).GetROProperty("label")),lcase("site")) > 0 Then
						wait 1
						On error resume next
						objDialog(iCounter).Click 1, 1
						bFlag = True
						Err.Number = 0
					End If
				Next
				If bFlag = False Then
					exit function
				End If

				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"Export File Name",StrExportFileName)
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Export")

				Set objDialog2 =Fn_SISW_Pref_GetObject("ExportPreferences")
				IF objDialog2.Exist = True Then					
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objDialog2, "Yes")
				End If

				If objPrefOper.JavaDialog("Information").Exist(10) Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper.JavaDialog("Information"), "OK")
				End If				
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Pass: " + "Successfully Exported Prefrence")
				Fn_SISW_Pref_PreferenceOperations = True
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
			
			Case "Export_Extn"	'Added Case for DIPRO TC Development - By Alok D
				ArrValue = Split(strScope,"~")
				Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"Export",0,0,"LEFT")
				
				' selecting From Location : Site
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "FromLocationDrpDwn")
				wait 2
				' add code to select static text of scope
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				bFlag = false
				For iCounter = 0 to objDialog.count-1
					If instr(lcase(objDialog(iCounter).GetROProperty("label")),lcase(ArrValue(0))) > 0 Then
						wait 1
						On error resume next
						objDialog(iCounter).Click 1, 1
						bFlag = True
						Err.Number = 0
					End If
				Next
				If bFlag = False Then
					exit function
				End If
				
				'From Location Sub value
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "ExportPrefForAllCategoryDrpDwn")
				wait 2
				' add code to select static text of scope
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				bFlag = false
				For iCounter = 0 to objDialog.count-1
					If instr(lcase(objDialog(iCounter).GetROProperty("label")),lcase(ArrValue(1))) > 0 Then
						wait 1
						On error resume next
						objDialog(iCounter).Click 1, 1
						bFlag = True
						Err.Number = 0
					End If
				Next
				If bFlag = False Then
					exit function
				End If
				
				
				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"Export File Name",StrExportFileName)
				Wait 2
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Export")
				
				Set objDialog2 =Fn_SISW_Pref_GetObject("ExportPreferences")
				IF objDialog2.Exist = True Then					
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objDialog2, "Yes")
				End If
				
				If objPrefOper.JavaDialog("Information").Exist(10) Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper.JavaDialog("Information"), "OK")
				End If				
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Pass: " + "Successfully Exported Prefrence")
				Fn_SISW_Pref_PreferenceOperations = True
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")

                ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Lock","VerifyErrorMsgAfterLock"					'[Tc1122:2016021000:07Mar2016:MadhuraP:NewDevelopment] - Added Case to verify error massage on Lock
				
			Call Fn_ReadyStatusSync(1)
			If Fn_UI_ObjectExist("Fn_SISW_Pref_PreferenceOperations", objPrefOper.JavaButton("LockPreference"))=True Then
				If Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"LockPreference") = True Then
					If StrAction <> "VerifyErrorMsgAfterLock" Then
						Set objSelectType=Description.Create()
						Set objSelectType1=Description.Create()
						objSelectType("Class Name").value = "JavaDialog"
						objSelectType1("Class Name").value = "JavaButton"		
						Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")				
						Do
							bFlag = False
							Set  intNoOfObjects =  objTcDefaultApplet.ChildObjects(objSelectType)
							For iCounter = 0 to intNoOfObjects.count-1
								If intNoOfObjects(iCounter).getroproperty("tagname") = "Lock Site Preferences" Then
									bFlag = True
									Set  intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
									For iCounter1 = 0 to intNoOfObjects1.count-1
										If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
												 intNoOfObjects1(iCounter1).click
												 Fn_SISW_Pref_PreferenceOperations = True
											Exit For
										End If
									Next
									Exit For
								End If
							Next
							Wait 3
						Loop While bFlag = True
					End If
					'-----------------------------------------------------------------------------------------------------------------------------
					'Added Condition to verify error message if exist after Lock button click .
					If StrAction = "VerifyErrorMsgAfterLock" Then
						If StrExportFileName <> "" Then
							aInfo = Split(StrExportFileName,"~")
							Set dicErrorInfo = CreateObject("Scripting.Dictionary")
							With dicErrorInfo	
								.Add "Title", aInfo(0)
								.Add "Message", aInfo(1)
								.Add "Button", aInfo(2)
							End with
							bReturn = Fn_SISW_ErrorVerify(dicErrorInfo)
							If bReturn = False Then
								Fn_SISW_Pref_PreferenceOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error message is not verified.")
							Else 
								Fn_SISW_Pref_PreferenceOperations = True
							End If
							Set dicErrorInfo = Nothing
						End If
					End If
					'-----------------------------------------------------------------------------------------------------------------------------
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " + "Preference locked successfully")		
				Else
					Fn_SISW_Pref_PreferenceOperations = False
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: " + "Failed to Lock the Preference")
				End If
			Else
				Fn_SISW_Pref_PreferenceOperations = False
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: " + "Lock Button does not Exist")
			End If

			'Swapnil:08-MAR-2013 : Preference has to be unlocked before closing the Option Dialog hence commented close call. 
			'Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
                ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Unlock","VerifyErrMsgAfterUnlock"					'[Tc1122:2016021000:07Mar2016:MadhuraP:NewDevelopment] - Added Case to verify error massage on Un-Lock
				
				Call Fn_ReadyStatusSync(1)
				If	Fn_UI_ObjectExist("Fn_SISW_Pref_PreferenceOperations", objPrefOper.JavaButton("OpenLock"))=True Then
					If Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"OpenLock") = True Then
						If StrAction <> "VerifyErrMsgAfterUnlock" Then
							Set objSelectType=Description.Create()
							Set objSelectType1=Description.Create()
	
							objSelectType("Class Name").value = "JavaDialog"
							objSelectType1("Class Name").value = "JavaButton"
							Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")		
							Do
								bFlag = False
								Set  intNoOfObjects =  objTcDefaultApplet.ChildObjects(objSelectType)
								For iCounter = 0 to intNoOfObjects.count-1
									If intNoOfObjects(iCounter).getroproperty("tagname") = "Unlock Site Preferences" Then
										bFlag = True
										Set  intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
										For iCounter1 = 0 to intNoOfObjects1.count-1
											If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
													 intNoOfObjects1(iCounter1).click
													 Fn_SISW_Pref_PreferenceOperations = True
											Exit For
											End If
										Next
										Exit For
									End If
								Next
								Wait 3
							Loop While bFlag = True	
						End If
						'-----------------------------------------------------------------------------------------------------------------------------
						'Added Condition to verify error message if exist after Un-Lock button click .
						If StrAction = "VerifyErrMsgAfterUnlock" Then
							If StrExportFileName <> "" Then
								aInfo = Split(StrExportFileName,"~")
								Set dicErrorInfo = CreateObject("Scripting.Dictionary")
								With dicErrorInfo	
									.Add "Title", aInfo(0)
									.Add "Message", aInfo(1)
									.Add "Button", aInfo(2)
								End with
								bReturn = Fn_SISW_ErrorVerify(dicErrorInfo)
								If bReturn = False Then
									Fn_SISW_Pref_PreferenceOperations = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Error message is not verified.")
								Else
									Fn_SISW_Pref_PreferenceOperations = True
								End If
								Set dicErrorInfo = Nothing
							End If
						End If
						'-----------------------------------------------------------------------------------------------------------------------------
						Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " + "Preference unlocked successfully")		
					Else
						Fn_SISW_Pref_PreferenceOperations = False
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: " + "Failed to Unlock the Preference")
					End If
				else
					Fn_SISW_Pref_PreferenceOperations = False
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: " + "Unlock button does not Exist")
				End If
                ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "VerifyLock"
				If Fn_UI_ObjectExist("Fn_SISW_Pref_PreferenceOperations", objPrefOper.JavaButton("Close"))=True Then
					If Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"Close") = True Then
						Set objSelectType=Description.Create()
						Set objSelectType1=Description.Create()
						objSelectType("Class Name").value = "JavaDialog"
						objSelectType1("Class Name").value = "JavaButton"
						Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")		

						Do
							bFlag = False
							Set  intNoOfObjects =  objTcDefaultApplet.ChildObjects(objSelectType)
							For iCounter = 0 to intNoOfObjects.count-1
								If intNoOfObjects(iCounter).getroproperty("tagname") = "Unlock Site Preferences" Then
									bFlag = True
									Set  intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
									For iCounter1 = 0 to intNoOfObjects1.count-1
										If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
											 intNoOfObjects1(iCounter1).click
											 Fn_SISW_Pref_PreferenceOperations = True
										Exit For
										End If
									Next
									Exit For
								End If
							Next
							Wait 3
						Loop While bFlag = True				
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " + "Successfully Verified that Preference is Locked")
						Fn_SISW_Pref_PreferenceOperations = True		
					Else
						Fn_SISW_Pref_PreferenceOperations = False
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: " + "Failed to Verify that Preference is Locked")
					End If
				Else
					Fn_SISW_Pref_PreferenceOperations = False
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: " + "Close button does not Exist")
				End If
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
				
                ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "ModifyPreferenceWithScope"			
			Call Fn_ReadyStatusSync(1)

			If bFlag Then
				'click on Edit Button
				If objPrefOper.JavaButton("Edit").Exist(3) Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Edit")
				End If

				Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CurrentValues",StrValue)
				If objPrefOper.JavaList("PrefMultiValList").Exist(3) Then
					Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewMultiValPrefAddValue")
				End If
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Save")
				Fn_SISW_Pref_PreferenceOperations = True

			Else
				Fn_SISW_Pref_PreferenceOperations = False
			End If
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		 Case "GetPreferenceValue"
				Fn_SISW_Pref_PreferenceOperations = False
				If bFlag=True Then
					sPerfValue = objPrefOper.JavaEdit("CurrentValues").GetROProperty ("value")
					Fn_SISW_Pref_PreferenceOperations = sPerfValue
				End If
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Close")
	End Select
	'Call Fn_WriteLogFile("Fn_SISW_Pref_PreferenceOperations", 3, Err.Number ,"Sucessfully Executed Fn_SISW_Pref_PreferenceOperations")
	'Fn_SISW_Pref_PreferenceOperations = True

	If StrAction<>"Lock"  And bUnlock=True Then     'Added by Nilesh on 16-July-2012
		Set objSelectType=Description.Create()
		Set objSelectType1=Description.Create()
		objSelectType("Class Name").value = "JavaDialog"
		objSelectType1("Class Name").value = "JavaButton"
		Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")		
        
		Set  intNoOfObjects = objTcDefaultApplet.ChildObjects(objSelectType)
		For iCounter = 0 to intNoOfObjects.count-1
			If intNoOfObjects(iCounter).getroproperty("tagname") = "Unlock Site Preferences" Then
				Set  intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
				For iCounter1 = 0 to intNoOfObjects1.count-1
					If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
						intNoOfObjects1(iCounter1).click
						'Unlock preferences and Click on cancel button.
						Call Fn_SISW_Pref_PreferenceOperations("Unlock","","","","","","","","","","","")
						Exit For
					End If
				Next
				Exit For
			End If
		Next
	End If
	
	Set objPrefOper = Nothing
	Set objDialog=Nothing
	Set objSelectType=Nothing
	Set objSelectType1 = Nothing
	Set intNoOfObjects = Nothing
	Set intNoOfObjects1 = Nothing
	Set objTcDefaultApplet = Nothing
	'Exit Function
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_SISW_Pref_PreferenceOperations modified the Preference [ "&StrPrefName & " ]")

' Above code is being commented as this function returns an invalid "False" when Fn_MenuOperation function is called.
End Function

'*********************************************************		Fn_SISW_Pref_Search_Operation		'Function Name		:				Fn_SISW_Pref_Search_Operation

'Description			 :		 		 This Function is used for following :-
'																				
'													1. Modify [---Done---]
'													2.Delete [---Done---]
'													3. Remove [---Done---]
'													4. Add [-- Done--]
'													5. Lock - Prerequisite : Index Page should be present in Edit->Options
'														Example : Fn_SISW_Pref_PreferenceOperations("Lock","","","","","","","","","","","")
'													6. Unlock - Prerequisite : Index Page should be present in Edit->Options
'														Example : Fn_SISW_Pref_PreferenceOperations("Unlock","","","","","","","","","","","")
'Parameters			   :	 			
'													1.sAction,
'                                                   2.sSearchOnKeyWord,
'                                                   3.sCurrentValue)
										
'Return Value		   : 			The String which represents the result : "PASS" or "FAIL" with the reason

'Pre-requisite			:		 	User should logged in to the teamcenter with DBA Privilledge

'Examples				:			Call Fn_SISW_Pref_Search_Operation("Modify","ItemRevision.SUMMARYRENDERING","SampleItemRevStylesheet")
'												Call Fn_SISW_Pref_Search_Operation("Delete","abc","")
'
'History					:			
'	Developer Name				Date			Rev. No.				Changes Done						Reviewer				Reviewed Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Rakesh						30-Apr-2010			1.0							Harshal/									Santosh						30-Apr-2010			
'	Swapna Ghatge				21-May-2010			1.1						Added Code of Delete			Mohit										21-May-2010	
'	Pallavi Patil				25-May-2010			1.2					Added Exists Case					Mohit				
'	Sandeep N					08-Mar-2011		1.3					Added "DeleteUserPreference" Case					Sunny				
'	Sachin J.					05-June-2012		1.4					Modified case "Exists"
'	Sachin J.					05-June-2012		1.4					Modified case "Modify"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W.					12-June-2012		2.0					Modified function according to TC10.0 UI changes
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Jeevan M.					29-June-2012		2.0					Modified cases Add, Remove
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Jeevan M.					06-July-2012		2.0					Modified cases "Modify"  Verify Editable
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Pref_Search_Operation(sAction,sSearchOnKeyWord,sCurrentValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_Search_Operation"
	Dim objPrefOper,objSearch,WshShell,objdelete,objdeletedialog, bFound,objTcDefaultApplet
	Dim oDesc, iCnt, iCounter, iSubCounter, bFlag,bResult
	Dim iCountChecked, objVerifyMsg, intNoOfEditObj
	Dim  objSelectType, objSelectType1, intNoOfObjects, intNoOfObjects1, iCounter1,bClose,aAction

	Fn_SISW_Pref_Search_Operation = False
	bFlag = false
	'Code added by Nilesh to handle close option dialog condition
	bClose=True
	If  Instr(sAction,"~")>0 Then
		aAction=Split(sAction,"~",-1,1)
		If Ubound(aAction)<>0 Then
			sAction=aAction(0)
			bClose=aAction(1)
		End If
	End If
	'End 
'	Modified Function To Handle : Fn_SISW_Pref_GetObject("IndexOptions2") -  17-Jul-2012  - Pranav  Ingle

	If Fn_SISW_Pref_GetObject("IndexOptions").Exist(5) = False And Fn_SISW_Pref_GetObject("IndexOptions2").Exist(5) = False Then
		Call Fn_MenuOperation("Select","Edit:Options...")
		Call Fn_ReadyStatusSync(1) 
	End If

	If Fn_SISW_Pref_GetObject("IndexOptions").Exist(6) Then
		Set objPrefOper =   Fn_SISW_Pref_GetObject("IndexOptions")
	ElseIf Fn_SISW_Pref_GetObject("IndexOptions2").Exist(6) Then
		Set objPrefOper =   Fn_SISW_Pref_GetObject("IndexOptions2")
	Else 
        Set objPrefOper = Nothing
		Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: To Find Object In Object Repository ")		
		Exit Function
	End If

	Call Fn_ReadyStatusSync(2)
	Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Search_Operation", objPrefOper,"Search",0,0,"LEFT")
	Call Fn_ResizeWindow("Resize","700", "800", objPrefOper)

	Select Case sAction
		Case "DeleteUserPreference"
			'objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Details"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"			
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"BottomLink",10, 10,"LEFT")

			Call Fn_Edit_Box("Fn_SISW_Pref_Search_Operation",objPrefOper,"KeywordSearch",sSearchOnKeyWord)
			Wait(3)
			'Select Preference
				iCounter = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
				For iCnt = 0 to iCounter - 1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSearchOnKeyWord Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
						wait 2
						If lcase(trim(objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Location"))) = lcase(trim("user")) Then
							bFlag = True
							Exit for
						End If
					End If
				Next
			
				Case Else
		'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		'Added if condition to run the code according to the passed value to the parameter "sSearchOnKeyWord"
		'By Pallavi Patil
		'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			If  sSearchOnKeyWord <> "" Then
					Call Fn_Edit_Box("Fn_SISW_Pref_Search_Operation",objPrefOper,"KeywordSearch",sSearchOnKeyWord)
					Wait(3)
					'Select Preference
					iCounter = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
					For iCnt = 0 to iCounter - 1
						If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSearchOnKeyWord Then
							objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
							bFlag = True
							wait 2
							Exit for
						End If
					Next
					If bFlag = False Then
						If sAction = "Delete" Then 
							Fn_SISW_Pref_Search_Operation = True
						End If
						If sAction = "Exists" Then
							objPrefOper.Close
						End If
						Exit Function
					End If
			End If
	End Select	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Modify"
			'click on Edit Button
			If objPrefOper.JavaButton("Edit").Exist(3) Then
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Edit")
			End If

			'Set value
			bResult= Fn_UI_Object_GetROProperty("Fn_SISW_Pref_Search_Operation",objPrefOper.JavaEdit("CurrentValues"), "editable")
			If bResult=0  Then
				Fn_SISW_Pref_Search_Operation=False
				objPrefOper.Close()
				Exit Function
			Else
				Call Fn_Edit_Box("Fn_SISW_Pref_Search_Operation",objPrefOper,"CurrentValues",sCurrentValue)
			End If
			
			
			If objPrefOper.JavaList("PrefMultiValList").Exist(2) Then
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "NewMultiValPrefAddValue")
			End If

			' Click on Save Button
			If objPrefOper.JavaButton("Save").Exist(3) Then
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Save")
			End If

			Fn_SISW_Pref_Search_Operation = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Exists"
			If objPrefOper.JavaButton("Edit").Exist(3) Then'Added By Rima Patil
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Edit")
			End If
			'End
			If objPrefOper.JavaList("PrefMultiValList").Exist(2) Then
				Fn_SISW_Pref_Search_Operation = Fn_UI_ListItemExist("Fn_SISW_Pref_Search_Operation", objPrefOper, "PrefMultiValList",sCurrentValue)			
			Else
				If Fn_UI_Object_GetROProperty("Fn_SISW_Pref_Search_Operation",objPrefOper.JavaEdit("CurrentValues"), "value") =sCurrentValue Then
					Fn_SISW_Pref_Search_Operation = True
				Else
					Fn_SISW_Pref_Search_Operation = False
				End If
			' Click on Cancel Button
			If objPrefOper.JavaButton("Cancel").Exist(3) Then
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Cancel")
			End If
				wait 3
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Delete"
'			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Details"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Definition"  'Modified by Nilesh on 11 Jul-12
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Search_Operation", objPrefOper,"BottomLink",10, 10,"LEFT")

			Dim bReturn
			' click on delete button
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "DeletePreference")            
			Set objdelete =  Fn_SISW_Pref_GetObject("DeletePreference" )
			If objdelete.Exist =True Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objdelete, "Yes")	
				Fn_SISW_Pref_Search_Operation=True
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " +sSearchOnKeyWord+ " Preference Deleted successfully")		
			 End If  
			'variables are released
			Set objdelete = Nothing
			Set objdeletedialog = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Remove" ' To remove the selected Preference current value
			If objPrefOper.JavaButton("Edit").Exist(3) Then
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Edit")
			End If

			If Fn_UI_ListItemExist("Fn_UI_ListItemExist", objPrefOper, "PrefMultiValList",sCurrentValue) = True Then
				Call Fn_List_Select("Fn_SISW_Pref_Search_Operation", objPrefOper, "PrefMultiValList",sCurrentValue)
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "NewMultiValPrefRemoveValue") 
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Save")
			End If
			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " +sSearchOnKeyWord+ " Prefrence Current Value :["+sCurrentValue+"] Removed successfully")	
			Fn_SISW_Pref_Search_Operation = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Add" ' To Add the Requested Preference current value if not found
			If objPrefOper.JavaButton("Edit").Exist(3) Then
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Edit")
			End If
			If Fn_UI_ListItemExist("Fn_UI_ListItemExist", objPrefOper, "PrefMultiValList",sCurrentValue) = False Then
				Call Fn_Edit_Box("Fn_SISW_Pref_Search_Operation",objPrefOper,"CurrentValues",sCurrentValue)
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "NewMultiValPrefAddValue") 
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Save")
			Else
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Cancel")
			End If
			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " +sSearchOnKeyWord+ " Prefrence Current Value :["+sCurrentValue+"] Removed successfully")	
			Fn_SISW_Pref_Search_Operation = True
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Lock"				
			If Fn_Button_Click("Fn_SISW_Pref_Search_Operation",objPrefOper,"LockPreference") = True Then
				Set objSelectType=Description.Create()
				Set objSelectType1=Description.Create()
				Set objVerifyMsg = Description.Create()
				objSelectType("Class Name").value = "JavaDialog"	
				objSelectType1("Class Name").value = "JavaButton"
				objVerifyMsg("Class Name").value = "JavaEdit"
                Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")		
				
				Do
					bFlag = False
					Set  intNoOfObjects = objTcDefaultApplet.ChildObjects(objSelectType)
					For iCounter = 0 to intNoOfObjects.count-1
						If intNoOfObjects(iCounter).getroproperty("tagname") = "Lock Site Preferences" Then							
							'Getting the JavaButton child objects
							Set intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
									If sCurrentValue <> "" Then
										'Getting the JavaEdit child objects
										Set  intNoOfEditObj = intNoOfObjects(iCounter).ChildObjects(objVerifyMsg)
										'Verify Msg 
										For iCounter1 = 0 to intNoOfEditObj.count-1
											If InStr(1, intNoOfEditObj(iCounter1).getroproperty("value"), sCurrentValue) <> 0 Then
												bFlag = True
												Exit For
											End If
										Next
									Else
										bFlag = True
									End If
							'Click on OK button
							For iCounter1 = 0 to intNoOfObjects1.count-1
								If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
										 intNoOfObjects1(iCounter1).click
										 Fn_SISW_Pref_Search_Operation = True
									Exit For
								End If
							Next
							Exit For
						End If
					Next
					Wait 3
				Loop While bFlag = True
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " + "Preference locked successfully")		

			Else
				Fn_SISW_Pref_Search_Operation = False
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "Unlock"								
			If Fn_Button_Click("Fn_SISW_Pref_Search_Operation",objPrefOper,"OpenLock") = True Then
				Set objSelectType=Description.Create()
				Set objSelectType1=Description.Create()
				Set objVerifyMsg = Description.Create()
				objSelectType("Class Name").value = "JavaDialog"
				objSelectType1("Class Name").value = "JavaButton"
				objVerifyMsg("Class Name").value = "JavaEdit"
                Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")	
				Do
					bFlag = False
					Set  intNoOfObjects = objTcDefaultApplet.ChildObjects(objSelectType)
					For iCounter = 0 to intNoOfObjects.count-1
						If intNoOfObjects(iCounter).getroproperty("tagname") = "Unlock Site Preferences" Then
							'Getting the JavaButton child objects
							Set intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
							If sCurrentValue <> "" Then
								'Getting the JavaEdit child objects
								Set  intNoOfEditObj = intNoOfObjects(iCounter).ChildObjects(objVerifyMsg)
								'Verify Msg 
								For iCounter1 = 0 to intNoOfEditObj.count-1
									If InStr(1, intNoOfEditObj(iCounter1).getroproperty("value"), sCurrentValue) <> 0 Then
										bFlag = True
										Exit For
									End If
								Next
							Else
								bFlag = True
							End If
							'Click on OK button
							For iCounter1 = 0 to intNoOfObjects1.count-1
								If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
									 intNoOfObjects1(iCounter1).click
									 Fn_SISW_Pref_Search_Operation = True
									Exit For
								End If
							Next
							Exit For
						End If
					Next
					Wait 3
				Loop While bFlag = True
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " + "Preference unlocked successfully")		
			Else
				Fn_SISW_Pref_Search_Operation = False
			End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
		Case "DeleteUserPreference"   'Case To Delete User Level Preference					
			' click on delete button
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Delete")
			objdelete = Fn_UI_ObjectExist("Fn_SISW_Pref_Search_Operation",JavaDialog("Delete Preference(s)"))		
			Set objdeletedialog =  Fn_UI_ObjectCreate("Fn_SISW_Pref_Search_Operation",JavaDialog("Delete Preference(s)"))	
			'Check the Delete Preference(s) window exist
			If objdelete = True Then
				'if yes Click on Yes button
				wait(2)
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objdeletedialog, "Yes")	
				wait(2)
				'Log for success
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " +sSearchOnKeyWord+ " Preference Deleted successfully")		
			 End If  
			Fn_SISW_Pref_Search_Operation=True
		 '----------------------------------------------------
        Case "GetValue"  'Case added by Rima Patil
			If objPrefOper.JavaButton("Edit").Exist(SISW_MICRO_TIMEOUT) Then'Added By Rima Patil
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Edit")
			End If
			'End
			If objPrefOper.JavaList("PrefMultiValList").Exist(SISW_MICRO_TIMEOUT) Then
				Fn_SISW_Pref_Search_Operation = Fn_UI_ListItemExist("Fn_SISW_Pref_Search_Operation", objPrefOper, "PrefMultiValList",sCurrentValue)	
			Else
				Fn_SISW_Pref_Search_Operation = Fn_UI_Object_GetROProperty("Fn_SISW_Pref_Search_Operation",objPrefOper.JavaEdit("CurrentValues"), "value")
			End If	
			' Click on Cancel Button
			If objPrefOper.JavaButton("Cancel").Exist(SISW_MIN_TIMEOUT) Then
				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Cancel")
			End If
		End Select		
		wait SISW_MIN_TIMEOUT

		If objPrefOper.Exist(SISW_MICRO_TIMEOUT) And bClose=True Then
			'Click on cancel button
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation", objPrefOper, "Close")	
		End If
	
		Set objPrefOper = Nothing
		Set objSearch = Nothing

		Set objSelectType=Description.Create()
		Set objSelectType1=Description.Create()
		objSelectType("Class Name").value = "JavaDialog"
		objSelectType1("Class Name").value = "JavaButton"
        Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")	
		Set  intNoOfObjects = objTcDefaultApplet.ChildObjects(objSelectType)
		For iCounter = 0 to intNoOfObjects.count-1
			If intNoOfObjects(iCounter).getroproperty("tagname") = "Unlock Site Preferences" Then
				Set  intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
				For iCounter1 = 0 to intNoOfObjects1.count-1
					If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
						intNoOfObjects1(iCounter1).click
						Exit For
					End If
				Next
			Exit For
			End If
		Next
		Set  oDesc = Nothing
		Set  iCnt = Nothing
		Set  iCounter = Nothing
		Set  iSubCounter  = Nothing
		Set iCountChecked = Nothing
		Set objSelectType=Nothing
		Set objSelectType1 = Nothing
		Set objVerifyMsg = Nothing
		Set intNoOfObjects = Nothing
		Set intNoOfObjects1 = Nothing
		Set intNoOfEditObj = Nothing
		Set objTcDefaultApplet  = Nothing
End Function
'#######  EOF = Fn_SISW_Pref_Search_Operation ################################################################################################


'*********************************************************		Fn_SISW_Pref_WinPrefOperation		***********************************************************************
'Function Name		:				Fn_SISW_Pref_WinPrefOperation

'Description			 :		 		 This function will be used for performing an operation on Preference dialog

'Parameters			   :	 			1. sAction: "Verify / Modify / "
'													2. sNodeName: "General" or "Install/Update" or "Search:Favorites"
'													3. dParameter: number of parameters(send through Dictionary object) (Object Identifier/Name:Value)
'''' 												dParameter is considered as Name value pair seperated by ":" As we haven't finalized upon Dictionary usage.
'Return Value		   : 			True/False

'Pre-requisite			:		 	Nothing

'Examples				:			'msgbox Fn_SISW_Pref_WinPrefOperation("Verify","General","")
												'msgbox Fn_SISW_Pref_WinPrefOperation("Verify","Search:Results","")
												'Call Fn_SISW_Pref_WinPrefOperation("Modify","General","Minimum characters for view title:30")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Harshal		 													   10/05/2010			              1.0																					Sameer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'*********************************************************************************************************************************************************************************************
Function Fn_SISW_Pref_WinPrefOperation(sAction, sNodeName, dParameter)
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_WinPrefOperation"
		Dim iNodeCount, iCounter, sNode, aParameterValue
		Dim objPreference, objPreferenceTree
        If Fn_SISW_Pref_GetObject("Preferences@1" ).exist(3)= true Then
			Set objPreference = Fn_SISW_Pref_GetObject("Preferences@1" )
	    else	
			Set objPreference = Fn_SISW_Pref_GetObject("Preferences" )	
		End If

		' Opening Peference window if not exist
		If objPreference.Exist = False Then
				Call Fn_MenuOperation("Select","Window:Preferences")
		End If

		Select Case sAction
				Case "Verify"
						Set objPreferenceTree = Fn_UI_ObjectCreate("Fn_SISW_Pref_WinPrefOperation", objPreference.JavaTree("Tree"))

						'Expand Search node if supplied
						If InStr(1,  sNodeName, "Search") <> 0Then
								Call Fn_UI_JavaTree_Expand("Fn_SISW_Pref_WinPrefOperation", objPreference, "Tree","Search")
						End If

						' Total node count
						iNodeCount = Fn_UI_Object_GetROProperty("Fn_SISW_Pref_WinPrefOperation",objPreferenceTree, "items count")

						' Searching for the node existance
						For iCounter = 0 To iNodeCount - 1
								sNode = objPreference.JavaTree("Tree").GetItem(iCounter)
								If sNodeName = sNode Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Node verification of node "+sNodeName+" in preferences completed successfully.")
										Fn_SISW_Pref_WinPrefOperation = True
										Call Fn_Button_Click ("Fn_SISW_Pref_WinPrefOperation", objPreference, "OK")
										Exit Function
								End If
						Next
                    	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Node verification of node "+sNodeName+" in preferences failed.")
						Fn_SISW_Pref_WinPrefOperation = False

				Case "Modify"
						Select Case sNodeName
								Case "General"
										aParameterValue = Split(dParameter, ":")

										' Sets the "Show traditional style tabs" check box if provided
										If aParameterValue(0) = "Show traditional style tabs" Then
												Call Fn_CheckBox_Set("Fn_SISW_Pref_WinPrefOperation", objPreference, "Show traditional style", aParameterValue(1))
										End If

										' sets the "MinCharFor"  value
										If aParameterValue(0)="Minimum characters for view title" Then
												objPreference.JavaEdit("MinCharFor").Set ""
												objPreference.JavaEdit("MinCharFor").PressKey Trim(aParameterValue(1))
										End If

										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Modification of  node "+sNodeName+" in preferences completed successfully.")

								Case "Teamcenter"
										aParameterValue = Split(dParameter, ":")

										' Sets the "Show traditional style tabs" check box if provided
										If aParameterValue(0) = "Show traditional style tabs" Then
												Call Fn_CheckBox_Set("Fn_SISW_Pref_WinPrefOperation", objPreference, "Show traditional style", aParameterValue(1))
										End If

										' sets the "MinCharFor"  value
										 objPreference.JavaTree("Tree").Select "Teamcenter"
										 wait 1
										If aParameterValue(0)="Minimum characters for view title" Then
												objPreference.JavaEdit("MinCharFor").Set ""
												objPreference.JavaEdit("MinCharFor").PressKey Trim(aParameterValue(1))
										End If

									If err.number<0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Modification of  node "+sNodeName+" in preferences Failed")
										Fn_SISW_Pref_WinPrefOperation = False
										Exit function
									End If

										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Modification of  node "+sNodeName+" in preferences completed successfully.")
										Fn_SISW_Pref_WinPrefOperation = true
						End Select
	
		End Select

		Call Fn_Button_Click ("Fn_SISW_Pref_WinPrefOperation", objPreference, "OK")
		Call  Fn_ReadyStatusSync(3)

		Set objPreference=Nothing
		Set objPreferenceTree=Nothing
End Function

'************************************************************** End of Fn_SISW_Pref_WinPrefOperation ***********************************************************************************

'*********************************************************		Fn_SISW_Pref_Search_Operation_WithCategory		***********************************************************************
'	Function Name			:				Fn_SISW_Pref_Search_Operation_WithCategory

'	Description			 	:				This Function is used for following :-
'																				
'											1. Modify [---Done---]
'											2. Delete [---Done---]
'											3. Exists
'											4. MultipleDelete
'											5. ListExists
'
'	Parameters				:	 			1.sAction,
'											2.sSearchOnKeyWord,
'											3.sCurrentValue
'											******* Others parameter added for future use but not implemented (except StrCategory) **********	
'											4. StrPrefName: Prefrences Name  
'											5. StrDesc: Value To Be modified
'											6.StrScope:
'											7.StrCategory
'											8.StrValue
'											9.StrType:
'											10.BlnMultiVal:
'											11.StrImportFilePath:
'											12. StrImportMd:
'											13.StrImportPrefOpt:
'											14. StrExportFileName:
										
'	Return Value			:				The String which represents the result : "PASS" or "FAIL" with the reason

'	Pre-requisite			:		 		User should logged in to the teamcenter with DBA Privilledge

'	Examples				:				Call Fn_SISW_Pref_Search_Operation_WithCategory("Exists", "AE*", "", "", "", "", "Configuration.Data Management", "", "", "", "", "", "","")
'                                               
'	History					:
'			
'	Developer Name			Date			Rev. No.						Changes Done											Reviewer					Reviewed Date
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Mahendra Bhandarkar   29-May-2010			1.0																					Mohit						29-May-2010			
'	Sandeep Navghane	  08-Mar-2011			1.0																					Sunny						08-Mar-2011
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Mohammad Ansari		  14-Jun-2012			2.0			Modified function according to TC10.0 UI changes
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Pallavi Patil		  05-Oct-2012			2.0			Added case "Verify" 
'															Added Code to modify "Description" field
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Ankit Nigam		      24-Feb-2016			2.0			Added Case "ListExists"											[Tc1122:24Feb2016:2016021600:AnkitN:NewDevelopment]
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Pref_Search_Operation_WithCategory(sAction,sSearchOnKeyWord,sCurrentValue, StrPrefName,StrDesc,StrScope,StrCategory,StrValue,StrType,BlnMultiVal,StrImportFilePath,StrImportMd,StrImportPrefOpt,StrExportFileName )
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_Search_Operation_WithCategory"
    Dim objPrefOper,objSearch,WshShell,objdelete,objdeletedialog, objSelectType, objDialog
    Dim oDesc, iCnt, iCounter, iSubCounter
    Dim iCountChecked,  bReturn, bFlag
	Dim aValue, iCount,absVal,StrActual
	Fn_SISW_Pref_Search_Operation_WithCategory=False
	Set objPrefOper = Fn_SISW_Pref_GetObject("IndexOptions")
	'Menu operation function called to select Option from Edit  Menu.
	If Fn_UI_ObjectExist("Fn_SISW_Pref_Search_Operation_WithCategory",objPrefOper)=False Then
		'IF not exist then opening from menu
		Call Fn_MenuOperation("Select","Edit:Options...")
	End If
	Call Fn_ReadyStatusSync(2)
	Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper,"Search",0,0,"LEFT")
	Call Fn_ResizeWindow("Resize","700", "800", objPrefOper)
	Call Fn_Edit_Box("Fn_SISW_Pref_Search_Operation_WithCategory",objPrefOper,"KeywordSearch",sSearchOnKeyWord)
	Wait(3)

	Select Case sAction
		Case "Modify","Exists","Verify"			'=============== To Search Prefrance  Category =========================
			If Trim(StrCategory) <> "" Then
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = StrCategory
				objPrefOper.JavaButton("SrchPrefCategoryBtn").Click
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If
		Case "Delete","MultiDelete","Create","ModifyCategory"
			'========== Do Not Search Prefrance Category==========================================
		Case Else
		'======== Do Nothing==============================
	End Select

	Select Case sAction
		Case "Modify"
		
		'============= To Select  Preference=======================================
		iCounter = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
		For iCnt = 0 to iCounter - 1
			If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSearchOnKeyWord Then
				objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
				bFlag = True
				wait 2
				Exit for
			End If
		Next
		'click on Edit Button
		If objPrefOper.JavaButton("Edit").Exist(3) Then
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "Edit")
		End If
	'Set value
		If sCurrentValue <> "" Then
			If sCurrentValue = "ModifyToBlankValue" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_Search_Operation_WithCategory",objPrefOper,"CurrentValues","")
			else
			 	Call Fn_Edit_Box("Fn_SISW_Pref_Search_Operation_WithCategory",objPrefOper,"CurrentValues",sCurrentValue)
			End If
		End If
		'Set Description    Added by Pallavi Patil on 05-Oct-2012
		If StrDesc <> "" Then
			Call Fn_Edit_Box("Fn_SISW_Pref_Search_Operation_WithCategory",objPrefOper,"Description",StrDesc)
		End If
		'Added by Nilesh on 25-Sep-2012
		If StrScope <> "" Then
			If objPrefOper.JavaButton("PrefScopeDrpDwn").GetROProperty("enabled")=1 Then
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = StrScope
                absVal=objPrefOper.JavaEdit("ProtectionScope").GetROProperty("abs_y")
				objPrefOper.JavaButton("PrefScopeDrpDwn").Click
				Wait 2
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				If objDialog.Count>1 Then
					For iCounter=0 To objDialog.Count-1
						If objDialog(iCounter).GetROProperty("abs_y")<> absVal Then
							objDialog(iCounter).Click 5,5 ,"LEFT"
							Exit For
						End If
					Next
				Else
					objDialog(0).Click 5, 5, "LEFT"
				End If
			Else
				Fn_SISW_Pref_Search_Operation_WithCategory=False
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Fail: Scope field is disabled")
				Exit Function
			End If
		End If
		Wait(3)
		' Click on Save Button
		If objPrefOper.JavaButton("Save").Exist(3) Then
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "Save")
		End If
		Fn_SISW_Pref_Search_Operation_WithCategory=True
		
		Case "Exists"
    	iCnt=objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows")
		iCountChecked = 0
		sSearchOnKeyWord = REPLACE(sSearchOnKeyWord, "*", "")
		For iCounter = 0 To Cint(iCnt) - 1 '============ To get the 0th cell data of PreferencesListTable and match with Search Keyword
			oDesc=objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCounter,"Name")
			If InStr(1, oDesc, sSearchOnKeyWord, 1) > 0  Then
				iCountChecked = iCountChecked + 1
				Exit For
			End If
		Next
		If iCountChecked > 0 Then
			Fn_SISW_Pref_Search_Operation_WithCategory = True
		Else
			Fn_SISW_Pref_Search_Operation_WithCategory = False
		End If
		'=================== To Click on cancel/close  button
		'Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "Close")
		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------				
		Case "ListExists"					'[Tc1122:24Feb2016:2016021600:AnkitN:NewDevelopment] - Added Case to verify list of search result
			iCnt=objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows")
			iCountChecked = 0
			sSearchOnKeyWord = REPLACE(sSearchOnKeyWord, "*", "")
			For iCounter = 0 To Cint(iCnt) - 1 
				oDesc=objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCounter,"Name")
				If InStr(1, oDesc, sSearchOnKeyWord, 1) > 0  Then
					iCountChecked = iCountChecked + 1
				End If
			Next
			If cStr(iCountChecked) = cStr(iCnt) And cStr(iCnt) <> "0" Then
				Fn_SISW_Pref_Search_Operation_WithCategory = True
			Else
				Fn_SISW_Pref_Search_Operation_WithCategory = False
			End If	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------			
		Case "Delete"
		'============= To Select  Preference=======================================
		iCounter = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
		For iCnt = 0 to iCounter - 1
			If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSearchOnKeyWord Then
				objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
				bFlag = True
				wait 2
				Exit for
			End If
		Next
		'============= To  Delete Preference =======================================
		' click on delete button
		Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "DeletePreference")
		Set objdelete =  Fn_SISW_Pref_GetObject("DeletePreference" )
		If objdelete.Exist =True Then
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objdelete, "Yes")	
			Fn_SISW_Pref_Search_Operation_WithCategory=True
			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " +sSearchOnKeyWord+ " Preference Deleted successfully")		
		End If  
		
	Case "MultiDelete"
	'============= To Select  Preference Modified by Pritam =======================================
	bFlag = False
	aValue = Split(sCurrentValue,":",-1,1)
'			iCounter = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
	For iCount = 0 to UBound(aValue)
		iCounter = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
		For iCnt = 0 to iCounter - 1
			If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  aValue(iCount) Then
				objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
				bFlag = True
				wait 2
				Exit for
			End If
		Next
		If bFlag = True Then
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "DeletePreference")
			Set objdelete =  Fn_SISW_Pref_GetObject("DeletePreference" )
			If objdelete.Exist =True Then
					Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objdelete, "Yes")	
					wait 5
					Fn_SISW_Pref_Search_Operation_WithCategory=True
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " +sSearchOnKeyWord+ " Preference Deleted successfully")		
			End If
		Else
			  Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Fail: Failed to Delete Prefrence Operation ")
			  Fn_SISW_Pref_Search_Operation_WithCategory = False 
			  Exit Function
		End If
	Next
	'============= To  Delete Preference =======================================
	' click on delete button
'			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "DeletePreference")
'			objdelete = Fn_UI_ObjectExist("Fn_SISW_Pref_Search_Operation_WithCategory",JavaDialog("Delete Preference(s)"))		
'			Set objdeletedialog =  Fn_UI_ObjectCreate("Fn_SISW_Pref_Search_Operation_WithCategory",JavaDialog("Delete Preference(s)"))	
'			'Check the Delete Preference(s) window exist
'			If objdelete = True Then
'				'if yes Click on Yes button
'				Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objdeletedialog, "Yes")	
'				'Log for success
'				Fn_SISW_Pref_Search_Operation_WithCategory=True
'				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: " +sSearchOnKeyWord+ " Preference Deleted successfully")		
'			End If  

		Case "Create"   
			Dim bResult
			bResult=Fn_SISW_Pref_PreferenceOperations("Create",StrPrefName,StrDesc,StrScope,StrCategory,StrValue,StrType,BlnMultiVal,StrImportFilePath,StrImportMd,StrImportPrefOpt,StrExportFileName)
			If bResult=true Then
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Creation of New Prefrence Operation is Done successfully ")
				Fn_SISW_Pref_Search_Operation_WithCategory = True
			else
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Fail: Failed to create  New Prefrence Operation ")
				Fn_SISW_Pref_Search_Operation_WithCategory = False      
			End If
		
		Case "ModifyCategory"
		'============= To Select  Preference=======================================
		iCounter = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
		For iCnt = 0 to iCounter - 1
			If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSearchOnKeyWord Then
				objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
				bFlag = True
				wait 2
				Exit for
			End If
		Next
		' ==============================To Click Edit=================
		If objPrefOper.JavaButton("Edit").Exist(3) Then
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "Edit")
		End If
		'================ To Edit Category===============================
		If Trim(StrCategory) <> "" Then
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = StrCategory
			objPrefOper.JavaButton("PrefCategoryDrpDwn").Click
			Set objDialog = objPrefOper.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
		End If
		'Added by njilesh on 17-July-2012
        objPrefOper.JavaEdit("Description").Click 5,5
		Wait 1
		'End 
		'=========== To Click Save =====================
		If objPrefOper.JavaButton("Save").Exist(3) Then
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "Save")
		End If
		Fn_SISW_Pref_Search_Operation_WithCategory=True

    
	Case "Verify" 'Added by Pallavi Patil on 05-Oct-2012
		'============= To Select  Preference=======================================
		iCounter = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
		For iCnt = 0 to iCounter - 1
			If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSearchOnKeyWord Then
				objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
				bFlag = True
				wait 2
				Exit for
			End If
		Next
				If StrScope <> "" Then
					StrActual = Fn_UI_Object_GetROProperty("Fn_SISW_Pref_Search_Operation_WithCategory",objPrefOper.JavaEdit("ProtectionScope"),"value")
					If StrScope= StrActual Then
						Fn_SISW_Pref_Search_Operation_WithCategory=True
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Verification of [Protection Scope] is Done successfully ")
					Else
						Fn_SISW_Pref_Search_Operation_WithCategory=False
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Fail: Failed to To Verify [Protection scope]")
						Exit Function
					End If
				End If
				'[TC1121-2015102600-04_11_2015-VivekA-NewDevelopment] - Added by Poonam, code to verify description - [PSM New Developement]
				If StrDesc <> "" Then
					StrActual = Fn_UI_Object_GetROProperty("Fn_SISW_Pref_Search_Operation_WithCategory",objPrefOper.JavaEdit("Description"),"value")
					If Trim(StrDesc) = Trim(StrActual) Then
						Fn_SISW_Pref_Search_Operation_WithCategory=True
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Verification of [Description] is Done successfully ")
					Else
						Fn_SISW_Pref_Search_Operation_WithCategory=False
						Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"Fail: Failed to To Verify [Description]")
						Exit Function
					End If
				End If
				'------------------------------------------------------
	End Select '======= End of Select Statement======

	Select Case sAction
		Case "Modify","Exists","Verify","Delete","MultiDelete","ModifyCategory"	,"ListExists"			'=============== To Close Prefrance Dialog =========================
			'=========== To click Close ====================
			Call Fn_Button_Click("Fn_SISW_Pref_Search_Operation_WithCategory", objPrefOper, "Close")
		Case "Create" 
			'========== Do Not Close Prefrance Dialog, This activirty is already handeled in Function Call [ Fn_SISW_Pref_PreferenceOperations ]==========================================
		Case Else
		'======== Do Nothing==============================
	End Select
 
	'variables are released
	Set objdelete = Nothing
	Set objdeletedialog = Nothing
	Set objPrefOper = Nothing
	Set objSearch = Nothing
	Set  oDesc = Nothing
	Set  iCnt = Nothing
	Set  iCounter = Nothing
	Set  iSubCounter  = Nothing
	Set iCountChecked = Nothing
	Set objSelectType = Nothing
	Set objDialog = Nothing

End Function


'*********************************************************		To handle the Error Dialog of Create New Preference		*************************************************************************************
'Function Name		:				Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog

'Description			 :		 		 To handle the Error Dialog of Create New Preference

'Parameters			   :                1.sAttachedtxt: 

'Return Value		   : 				True\False

'Pre-requisite			:		 		The Error Dialog should be opened

'Examples				:                Call Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog("Preference already exists for the scope. You can modify the value(s) using the Details tab.")

'History:
'										Developer Name							Date				Rev. No.			Changes Done			Reviewer	 Review Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Dhananjay								29-May-2010			1.0																	Rizwan			29-May-10											
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'********************************************************************************************************************************************************************************************************************

Public Function Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog(sAttachedtxt) 
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog"
Dim sResult, diaCreatePref, btnOK, lblMsg, tmp, sErrorMsg, objWin, sActMsg, objOptions
'sAttachedtxt = "Preference already exists for the scope. You can modify the value(s) using the Details tab."
	Set objOptions = Fn_SISW_Pref_GetObject("Options")

	Set objWin = objOptions.JavaDialog("Create New Preference")
	'objWin.SetTOProperty "text","Create New Preference"
	If objWin.Exist = True Then
			If sAttachedtxt <> "" Then
					sActMsg = objWin.JavaEdit("PrefTextArea").GetROProperty("value")
					If instr(sActMsg, sAttachedtxt) > 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Validated Message [" + sAttachedtxt + "] on [Create New Preference Window]")
					Else
						Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog = False
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Validate Message [" + sAttachedtxt + "] on [Create New Preference]")
						Exit Function
					End If
			End If
			Call Fn_Button_Click("Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog",objWin, "OK")
	Else
	' Create Object Description of "New Preference" Dialog 
	Set diaCreatePref=description.Create()
	diaCreatePref("micclass").value="Dialog"
	diaCreatePref("regexpwndtitle").value = "Create New Preference"
	diaCreatePref("regexpwndclass").value = "#32770"

	'Description of Ok Button Object  on "Create New Preference" Dialog
	Set btnOK=description.Create()
	btnOK("micclass").value="WinButton"
	btnOK("nativeclass").value = "Button"
	btnOK("regexpwndtitle").value = "OK"

	'General Object description to search all Objects
	Set lblMsg=description.Create()
	'lblMsg("text").value = "Preference already exists for the scope. You can modify the value(s) using the Details tab."
    	If Dialog(diaCreatePref).Exist Then
			'Log the result to verify the Create new preference Dialog exist
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Create New Preference Dialog Exist.")
			'Capture All runtime objects to find message text
			Set  tmp = Dialog(diaCreatePref).ChildObjects(lblMsg)
			'Set message text to variable 
			sErrorMsg = tmp(1). getroproperty("text")  
			'compare run time message to verify  the error message
			If (sAttachedtxt = sErrorMsg ) Then
            		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verified Successfully")
           Else
                	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Error Message Verification Failed.")
					Exit Function
			End If
			' To Click "OK" Button after verification
			wait(2)
			Dialog(diaCreatePref).WinButton(btnOK).Click
			If Dialog(diaCreatePref).Exist Then
					wait(2)
					Dialog(diaCreatePref).WinButton(btnOK).Click
			End If
		Else
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Create New Preference Dialog does not Exist")
			Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog = False
			Exit Function
        End If
		End If
		If Err.Number < 0 Then
			Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog = False
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click Button [OK] on [Create New Preference] Dialog")
		End If		
	Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog = True
	Set diaCreatePref=nothing
	Set btnOK=nothing
	Set lblMsg=nothing
	Set tmp=nothing
	Set objOptions = nothing
	Set objWin = nothing
End Function

'##############################################################################################################################################
'###    FUNCTION NAME   :   Fn_SISW_Pref_RptChngs_Operation()
'###
'###    DESCRIPTION     :   Report Changes functionality through Edit>>Options dialog
'###
'###    PARAMETERS      :   sLink: This is the bottom Options link (Index/Search)
'###						sScope: Popup option (User/Role/Group/Site) afetr clicking on Customize Prefernce button
'###						sPreference: Preference to be selected from the list on the Report Changes dialog
'###						sExportPath: Local system path to export preference to
'###						bOpenOnCreate: True/False boolean tag to open exported file
'###						sOrigVal: Original value(s) of the preference - pass "~" separated values if multiple values need to check for
'###						sModVal: Modified value(s) of the preference - pass "~" separated values if multiple values need to check for (Multiple element seperated by :)
'###
'###    Function Calls  :   Fn_WriteLogFile (To report errors )
'###
'###    HISTORY         :   AUTHOR            		DATE        VERSION
'###
'###    CREATED BY      :   Mahendra Bhandarkar    	05/06/2010   1.0
'###
'###    REVIWED BY      :   Mohit Khare			    03/06/2010    1.0
'###
'###    MODIFIED BY     :   Shreyas Waichal			14-June-2012			Modified changes accordign to UI changes
'###
'###	MODIFIED BY     :   Nilesh &Pallavi 			10-Jul-12					Modified for TC10.0 Changes
'###
'###    EXAMPLE         :   Fn_SISW_Pref_RptChngs_Operation("Search", "Site", "AE_dataset_default_keep_limit", "", False, "3",  "4" )
'###############################################################################################################################################

Public Function Fn_SISW_Pref_RptChngs_Operation(sLink, sScope, sPreference, sExportPath, bOpenOnCreate, sOrigVal, sModVal )
		GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_RptChngs_Operation"
		Dim oDesc, iCounter, oList, bFlag, iCounter2, iCheckCnt, objPrefOper, oArrayList,objReportChange,objReport, objDialog2
		
		bFlag = False
		
		If Fn_SISW_Pref_GetObject("IndexOptions").Exist(5) = False and  Fn_SISW_Pref_GetObject("SearchOptions").Exist =  False and Fn_SISW_Pref_GetObject("OrganizationOptions").Exist = False Then
				Call Fn_MenuOperation("Select", "Edit:Options...")
				Call Fn_ReadyStatusSync(3)
		End If
		
		Set objReportChange=Fn_SISW_Pref_GetObject("ReportChanges")
		Set objReport=Fn_SISW_Pref_GetObject("Report")
		Set objDialog2 =Fn_SISW_Pref_GetObject("ExportPreferences")
		
		Set oDesc = Description.Create
		oDesc("Class Name").Value = "JavaButton"
		oDesc("label").Value = sScope
		
		Select Case sLink
		
		Case "Index"
		
			Set objPrefOper = Fn_SISW_Pref_GetObject("IndexOptions")
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"Index",0,0,"LEFT")
		
			objPrefOper.JavaCheckBox("Report").Click 5, 5,"LEFT"
			Set oList =	Fn_SISW_Pref_GetObject("IndexOptions").ChildObjects(oDesc)
			For iCounter = 0 To oList.Count -1
				If oList(iCounter).GetROProperty("label") =  sScope Then
					oList(iCounter).Click 
					Exit For
				End If
			Next
		
		Case "Search"
		
			 Set objPrefOper = Fn_SISW_Pref_GetObject("IndexOptions")
		'	 Set objSearch = Fn_UI_ObjectCreate("Fn_SISW_Pref_Search_Operation_WithCategory",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("SearchOptions").JavaStaticText("OptionBottomLink"))
			 Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Search_Operation_WithCategory",objPrefOper,sLink,0,0,"LEFT")
		
			objPrefOper.JavaCheckBox("Report").Click 5, 5,"LEFT"
			Wait 2
			If objReport.Exist(5)=False Then
					objPrefOper.JavaStaticText("Report").DblClick 0,1 
			End If
		
		'Set	oList =	JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("SearchOptions").ChildObjects(oDesc)
		'		For iCounter = 0 To oList.Count -1
		'			If oList(iCounter).GetROProperty("label") =  sScope Then
		'				oList(iCounter).Click 
		'				Exit For
		'			End If
		'		Next
		
			If sScope<>"" Then
				objReport.JavaButton(sScope).Click 
				Wait 2
			End If
		
		End Select
		
		If Trim(sPreference) <> "" Then
			Set oDesc = Description.Create
			oDesc("Class Name").Value = "JavaStaticText"
		
			objReportChange.JavaButton("PrefListDrpDwn").Click
		Set	oList =	objReportChange.ChildObjects(oDesc)
			For iCounter = 0 To oList.Count -1
				If oList(iCounter).GetROProperty("label") =  sPreference Then
					oList(iCounter).Click 1, 1,  "LEFT"
					Exit For
				End If
			Next
		Else
		'	Fn_SISW_Pref_GetObject("ReportChanges").JavaEdit("Preference").Type sPreference
		
		End If
		
		If Trim(sOrigVal) <> ""  Then
					iCheckCnt = 0
					oArrayList = Split(sOrigVal,":")															
					Set oDesc = Description.Create
					oDesc("Class Name").Value = "JavaStaticText"
			Set		oList = objReportChange.JavaList("OriginalValue").ChildObjects(oDesc)
						' Select the element from list  
						For iCounter=0 To oList.Count -1
							For iCounter2 = 0 To UBound(oArrayList)
								If oList(iCounter).GetROProperty("label") = oArrayList(iCounter2) then
								iCheckCnt = iCheckCnt + 1
								End If
							Next
						Next
		
						If iCheckCnt = UBound(oArrayList) Then
							bFlag = True
						End If
			'Call Fn_List_Select("Fn_SISW_Pref_RptChngs_Operation", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Report Changes"), "OriginalValue", sOrigVal)
		End If
		
		If Trim(sModVal) <> "" Then
					iCheckCnt = 0
					oArrayList = Split(sModVal,":")									
					Set oDesc = Description.Create
					oDesc("Class Name").Value = "JavaStaticText"
			Set		oList = objReportChange.JavaList("Modified Value").ChildObjects(oDesc)
						' Select the element from list  
						For iCounter=0 To oList.Count -1
							For iCounter2 = 0 To UBound(oArrayList)
								If oList(iCounter).GetROProperty("label") = oArrayList(iCounter2) then
								iCheckCnt = iCheckCnt + 1
								End If
							Next
						Next
		
						If iCheckCnt = UBound(oArrayList) Then
							bFlag = True
						End If
						
			'Call Fn_List_Select("Fn_SISW_Pref_RptChngs_Operation", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Report Changes"), "", sModVal)
		End If
		
		If Trim(sExportPath) <> "" Then
			Call Fn_Edit_Box("Fn_SISW_Pref_RptChngs_Operation",  objReportChange, "Export File Name", sExportPath )
			Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation", objReportChange, "Export")
		End If
		
		
		IF objDialog2.Exist = True Then					
			Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation", objDialog2, "Yes")
		End If		
		
		If Trim(bOpenonCreate) <> "" Then
			If Trim(bOpenOnCreate) = "True" Then
				Call Fn_CheckBox_Set("Fn_SISW_Pref_RptChngs_Operation", objReportChange, "Open on Export", "ON")
			Else
				Call Fn_CheckBox_Set("Fn_SISW_Pref_RptChngs_Operation", objReportChange, "Open on Export", "OFF")
			End If
		End If
		
		If objReportChange.JavaButton("Export").GetROProperty("enabled") = 1 Then
			Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation", objReportChange, "Export")
		End If
		
		IF objDialog2.Exist = True Then					
			Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation", objDialog2, "Yes")
		End If
		
		If Fn_UI_ObjectExist("Fn_SISW_Pref_RptChngs_Operation", objReportChange) = True Then
			Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation", objReportChange, "Cancel")
		End If
		
		Select Case sLink
				Case "Index"
		
		If Fn_UI_ObjectExist("Fn_SISW_Pref_RptChngs_Operation", objPrefOper.JavaButton("Cancel")) = True Then
					Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation", objPrefOper, "Cancel")
		Elseif Fn_UI_ObjectExist("Fn_SISW_Pref_RptChngs_Operation", objPrefOper.JavaButton("Close")) = True Then
					Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation", objPrefOper, "Close")	
		End if
		
				Case "Search"
		
		If Fn_UI_ObjectExist("Fn_SISW_Pref_RptChngs_Operation", objPrefOper.JavaButton("Cancel")) = True Then
					Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation",objPrefOper, "Cancel")
		Elseif Fn_UI_ObjectExist("Fn_SISW_Pref_RptChngs_Operation", objPrefOper.JavaButton("Close")) = True Then
					Call Fn_Button_Click("Fn_SISW_Pref_RptChngs_Operation", objPrefOper, "Close")	
		End if
		
		End Select
		
		If bFlag = True Then
			Fn_SISW_Pref_RptChngs_Operation = True
			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Report Changes added successfully")
		Else
			Fn_SISW_Pref_RptChngs_Operation = False
			Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Report Changes not added successfully")
		End If
		
		Set objPrefOper = Nothing
		Set oList  =Nothing
		Set oDesc = Nothing
		Set iCounter = Nothing
		Set bFlag  =Nothing
		Set iCounter2  =Nothing
		Set iCheckCnt = Nothing
		Set objReportChange=Nothing
		Set objReport=Nothing

End Function

'##############################################################################################################################################################################################################################################
'###
'###    FUNCTION NAME   :  Fn_SISW_Pref_Organization_Operation(sAction, sUser, sPrefName, sCategory, sType, sDescription, bMultiVal, sValues, sScope, sSrchText)
'###
'###    DESCRIPTION     :   Refer the doc attached with Fn_PreferenceOperation function. in this function also we need to code all the cases thr Organization link on the Options dialog
'###
'###    PARAMETERS      :   sAction - Create, Modify, Verify, Delete
'###                        sUser: User to be selected from org tree
'###                        sPrefName: 
'###                        sCategory:
'###                        sType:
'###                        sDescription:
'###                        bMultivAL:
'###                        sValus:
'###                        sSrchText:
'###                                         
'###    Function Calls  :  Fn_WriteLogFile()
'###
'###	 HISTORY         :   AUTHOR                 DATE        	VERSION
'###
'###    CREATED BY      :   Mahendra Bhandarkar	  27/04/2010        1.0
'###
'###    REVIWED BY      :   Mohit Khare		 	  27/04/2010	    1.0
'###
'###    MODIFIED BY     :	Ketan Raje			Modified "Delete" case on 31-Dec-2010.
'###
'###    MODIFIED BY     :	Koustubh W			Modified function as per TC10.0 UI changes 13-June-2012
'###  
'###     MODIFIED BY   :    Pritam Shikare     Modified case  " Modify" and "Verify "      on 16-Jul-2012
'###
'###	Modified By 	: Snehal Salunkhe		Added New Case "CreateGroupLevelPreference" and modified case "Verify " 	on 21-Apr-2015
'###
'###    EXAMPLE         : Fn_SISW_Pref_Organization_Operation("Create", "Organization:Engineering:Designer:Mahendra Bhandarkar (x_bhanda)", "XXXPreference_1234", "Classification", "String", "Test Description", "False", "False", "False", "")
'###					  bReturn = Fn_SISW_Pref_Organization_Operation("CreateGroupLevelPreference", "Organization:Engineering", "TC_refresh_notify","", "", "", "", "ON", "", "")
'#######################################################################################################################################################################################################################
Public Function Fn_SISW_Pref_Organization_Operation(sAction, sUser, sPrefName, sCategory, sType, sDescription, bMultiVal, sValues, sScope, sSrchText)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_Organization_Operation"
	Dim objList, oDesc, aDelList, jCounter, aUserList, iCounter, objSelectType, objDialog, bFlag, WshShell, objPrefOper, sPath, sValues1,objdelete
	Dim iItemCnt, iCnt, location, iCounter1, objPrefInstance, listCount,winPrefName

	Set objPrefOper = Fn_SISW_Pref_GetObject("IndexOptions")
	If Fn_UI_ObjectExist("Fn_SISW_Pref_Organization_Operation", objPrefOper)=False Then
		Call Fn_MenuOperation("Select","Edit:Options...")
		Call Fn_ReadyStatusSync(2)
	End If
	If Fn_UI_ObjectExist("Fn_SISW_Pref_Organization_Operation", objPrefOper)=True Then
		Set objPrefOper =   Fn_SISW_Pref_GetObject("IndexOptions")
	ElseIf Fn_SISW_Pref_GetObject("IndexOptions2").Exist = True Then
		Set objPrefOper =   Fn_SISW_Pref_GetObject("IndexOptions2")
	Else 
        Set objPrefOper = Nothing
		Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: To Find Object In Object Repository ")		
		Exit Function
	End If
	
	wait 2
	If sAction <> "CreateGroupLevelPreference" and sAction <> "CreateGroupLevelPreferenceWithRemoveValue" Then
		Set objPrefOper = Fn_UI_ObjectCreate("Fn_SISW_Pref_Organization_Operation", objPrefOper)
		Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper,"Organization",0,0,"LEFT")
		Wait 1
		If Trim(sUser) <> "" Then
			aUserList = Split(sUser, ":")
			For iCounter = 0 To UBound(aUserList)
				If iCounter = 0  Then
					sPath = aUserList(iCounter)
				Else
					sPath = sPath + ":"+aUserList(iCounter)
				End If
				Call Fn_JavaTree_Select("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Organization", sPath)
				Wait SISW_MICRO_TIMEOUT
				Call Fn_UI_JavaTree_Expand("Fn_SISW_Pref_Organization_Operation",objPrefOper,"Organization",sPath)
				Call Fn_ReadyStatusSync(1)
			Next
			Call Fn_JavaTree_Select("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Organization", sUser)
			Wait SISW_MICRO_TIMEOUT
		End If
	End If
   	

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Create"
			' Create Button is Disabled in build pnv6s166 - TC10.0 - 0606
			Fn_SISW_Pref_Organization_Operation = False    

'			Call Fn_UI_JavaStaticText_SetTOProperty("Fn_SISW_Pref_Organization_Operation", objPrefOper,"OrgBottomLink" , "label", "New")
'			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper,"OrgBottomLink",10, 10,"LEFT")                    
'
'			'Set value in Name Edit  box
'			Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"Name",sPrefName)
'
'			'Set value in Description Edit  box
'			If Trim(sDescription) <> "" Then
'				Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"Description",sDescription)
'			End If
'
'			'Select Scope
'			If Trim(sScope) <> "" Then
'				objPrefOper.JavaRadioButton("Scope").SetTOProperty "attached text", sScope
'				objPrefOper.JavaRadioButton("Scope").Set "ON"
'			End If
'
'			'Select category
'			If Trim(sCategory) <> "" Then
'				' JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("OrganizationOptions").JavaEdit("PrefDtlsCategory").Type sCategory
'				objPrefOper.JavaButton("PrefDtlsCategoryDrpDwn").Click
'				
'				Set objSelectType=Description.Create()
'				objSelectType("Class Name").value = "JavaStaticText"
'				objSelectType("label").value = sCategory
'				
'				Set objDialog = objPrefOper.ChildObjects(objSelectType)
'				objDialog(0).Click 5, 5, "LEFT"
'			End If
'
'			'Select Multiple value
'			If Trim(bMultiVal) <> "" Then
'				objPrefOper.JavaRadioButton("PreMultiValOpt").SetTOProperty "attached text", bMultiVal
'				objPrefOper.JavaRadioButton("PreMultiValOpt").Set "ON"
'			End If
'
'			'Select Type
'			If Trim(sType) <> "" Then
'				Set objSelectType=description.Create()
'				objSelectType("Class Name").value = "JavaStaticText"
'				objSelectType("label").value = sType
'				objPrefOper.JavaButton("PrefNewTypeDrpDwn").Click
'				
'				Set objDialog = objPrefOper.ChildObjects(objSelectType)
'				objDialog(0).Click 5, 5, "LEFT"
'			End If
'
'			If Trim(sValues) <> "" Then
'				If bMultiVal = "True" Then
'					sValues1 = Split(sValues, ":", -1, 1)
'					For iCounter = 0 To UBound(sValues1)
'						Call Fn_Edit_Box("", objPrefOper, "PrefMultival", sValues1(iCounter))
'						Call Fn_Button_Click("", objPrefOper, "PrefMultiValAdd")
'					Next
'				Else
'					Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"PrefValue",sValues)
'				End If
'			End If                                           
'
'			Do
'				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Create")
'			Loop Until objPrefOper.Exist = True
'			wait(5)
'
'			bFlag = Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog("Preference already exists for the scope. You can modify the value(s) using the Details tab.")
'
'			If  objPrefOper.JavaButton("Cancel").Exist = True Then
'				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Cancel")
'			End If
'
'			'Log for Success
'			If bFlag = True Then
'				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Creation of New Prefrence Operation failed ")                                        
'				Fn_SISW_Pref_Organization_Operation = False                                                               
'			Else
'				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Creation of New Prefrence Operation is Done successfully ")
'				Fn_SISW_Pref_Organization_Operation = True    
'			End If

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Modify","ModifyPrefRemoveAllValues","ModifyPrefRemoveValues"
			'	clikcing on "Details" tab
			'objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Details"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"			
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper,"BottomLink",10, 10,"LEFT")

			'Set value in Name Edit  box
			If Trim(sPrefName) <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"SrchPrefName",sPrefName)
                Call Fn_ReadyStatusSync(1)
			End If

			'Select category
			If Trim(sCategory) <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "FilterByCategoryDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sCategory
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			' Set the Scope of the prefernce to search 
			If Trim(sScope) <> "" Then
					Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "FilterByScopeDrpDwn")
					iCounter1 = 0
					If objPrefOper.JavaStaticText("ValueLocation").Exist(1) Then
						location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
						If lcase(trim(sScope)) = location Then
							iCounter1 = 1
						End If
					End If
					wait 2
					' add code to select static text of scope
					Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					objSelectType("label").value = sScope
					Set objDialog = objPrefOper.ChildObjects(objSelectType)
					objDialog(iCounter1).Click 5, 5, "LEFT"
			End If

			'Call Fn_List_Select("Fn_SISW_Pref_Organization_Operation", objPrefOper, "SrchPrefList",sSrchText)
			If Trim(sSrchText ) <> "" Then
				'Select Preference
				iItemCnt = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
				For iCnt = 0 to iItemCnt - 1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSrchText Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
						bFlag = True
						wait 2
						Exit for
					End If
				Next
				If bFlag = False Then
					Fn_SISW_Pref_Organization_Operation = False
					Exit Function
				End If
			End If

			
			'click on Edit Button
			If objPrefOper.JavaButton("Edit").Exist(3) Then
				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Edit")
			End If

			'click on Create Instance Button
			If objPrefOper.JavaButton("CreateNewPreference").Exist(3) Then
				If cInt(objPrefOper.JavaButton("CreateNewPreference").GetROProperty("enabled")) = 1  Then
					Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "CreateNewPreference")
				End If
			End If

			'Set value in Description Edit  box
			If Trim(sDescription) <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"Description",sDescription)
			End If

			'Select Multiple value
			If Trim(bMultiVal) <> "" Then
				If cBool(bMultiVal) Then
					bMultiVal = "Multiple"
				Else
					bMultiVal = "Single"
				End IF
				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "PrefMultipleDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = bMultiVal
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			'Select Type
			If Trim(sType) <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "NewPrefTypeDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sType
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			'Select Values
			If sAction="Modify" then
				If Trim(sValues) <> "" Then
					'Multivalues Set
					If instr(sValues,":") > 0 Then
						sValues1 = Split(sValues,":",-1)
						For iCounter=0 to Ubound(sValues1)
							Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"CurrentValues",sValues1(iCounter))
							wait 1
							Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "NewMultiValPrefAddValue")
							wait 1								
						Next
					ElseIf Lcase(Trim(sValues)) = "blank" Then
						'Set as Empty
						Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"CurrentValues","")
						 If  objPrefOper.JavaButton("NewMultiValPrefAddValue").Exist Then
							Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "NewMultiValPrefAddValue")
						 End If		
					Else
						'Single Value Set
						Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"CurrentValues",sValues)
	                    objPrefOper.JavaButton("Save").object.setEnabled true
						If  objPrefOper.JavaButton("NewMultiValPrefAddValue").Exist Then
							Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "NewMultiValPrefAddValue")
						End If
					End If
				End If
			End if
			
			If sAction="Modify" Then
				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Save")
				Fn_SISW_Pref_Organization_Operation = Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Close")
			End If
			
			If sAction="ModifyPrefRemoveAllValues" or sAction="ModifyPrefRemoveValues" Then
			
				If objPrefOper.JavaButton("Edit").Exist(5) Then
					Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Edit")
				End If
				
				listCount = objPrefOper.JavaList("PrefMultiValList").GetROProperty("items count")
				If Cint(listCount) = 0 Then
					bFlag = True	
				End If

				For iCounter = Cint(listCount -1) To 0 Step -1
					winPrefName = objPrefOper.JavaList("PrefMultiValList").GetItem(iCounter)
					If winPrefName <> "" Then
						Call Fn_List_Select("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList",winPrefName)
						Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefRemoveValue")
						bFlag = True	
					End If	
				Next
				
				If sAction="ModifyPrefRemoveValues" Then
					sValues1 = Split(sValues,":",-1)
					For iCounter=0 to Ubound(sValues1)
						Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"CurrentValues",sValues1(iCounter))
						wait 1
						Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "NewMultiValPrefAddValue")
						wait 1								
					Next
				ElseIf Lcase(Trim(sValues)) = "blank" Then
					'Set as Empty
					Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"CurrentValues","")
					 If  objPrefOper.JavaButton("NewMultiValPrefAddValue").Exist Then
						Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "NewMultiValPrefAddValue")
					 End If		
				Else
					'Single Value Set
					Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"CurrentValues",sValues)
                    objPrefOper.JavaButton("Save").object.setEnabled true
					If  objPrefOper.JavaButton("NewMultiValPrefAddValue").Exist Then
						Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "NewMultiValPrefAddValue")
					End If
				End If
				
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
			
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
				Fn_SISW_Pref_Organization_Operation = bFlag
			
			End if
		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Verify"
		
			Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Refresh")
			
			'	clikcing on "Details" tab
			'objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Details"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"
		
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper,"BottomLink",10, 10,"LEFT")

			'Set value in Name Edit  box
			If Trim(sPrefName) <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"SrchPrefName",sPrefName)
				Call Fn_ReadyStatusSync(1)
			End If

			'Select category
			If Trim(sCategory) <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "FilterByCategoryDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sCategory
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			' Set the Scope of the prefernce to search 
			If Trim(sScope) <> "" Then
					Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "FilterByScopeDrpDwn")
					iCounter1 = 0
					If objPrefOper.JavaStaticText("ValueLocation").Exist(1) Then
						location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
						If lcase(trim(sScope)) = location Then
							iCounter1 = 1
						End If
					End If
					wait 2
					' add code to select static text of scope
					Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					objSelectType("label").value = sScope
					Set objDialog = objPrefOper.ChildObjects(objSelectType)
					objDialog(iCounter1).Click 5, 5, "LEFT"
			End If

			'Call Fn_List_Select("Fn_SISW_Pref_Organization_Operation", objPrefOper, "SrchPrefList",sSrchText)
			If Trim(sSrchText ) <> "" Then
				'Select Preference
				iItemCnt = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
				For iCnt = 0 to iItemCnt - 1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSrchText Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
						bFlag = True
						wait 2
						Exit for
					End If
				Next
				If bFlag = False Then
					Call Fn_Button_Click("", objPrefOper, "Close")
					Fn_SISW_Pref_Organization_Operation = False
					Exit Function
				End If
			End If

			'Set value in Description Edit  box
			If Trim(sDescription) <> "" Then
				If objPrefOper.JavaEdit("Description").GetROProperty("value") = sDescription Then
					bFlag = True
				Else
					bFlag = False
				End If
			End If

			'Select Multiple value
			If Trim(sScope) <> "" Then
                                location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
				If  instr(location , lCase(sScope)) > 0 Then
					bFlag = True
				Else
					bFlag = False
				End If
			End If

			'Select category
			If Trim(sCategory) <> "" Then
				If objPrefOper.JavaEdit("CategoryCreateCategory").GetROProperty("value") = sCategory Then
					bFlag = True
				Else
					bFlag = False
				End If
			End If

			'Select Multiple value
			If Trim(bMultiVal) <> "" Then
				If cBool(bMultiVal) Then
					bMultiVal = "Multiple"
				Else
					bMultiVal = "Single"
				End If
				If objPrefOper.JavaEdit("Multiple").GetROProperty("value") =  bMultiVal Then
					bFlag = True
				Else
					bFlag = False
				End If
			End If

			'Verify Values
			If Trim(sValues) <> "" Then
				If objPrefOper.JavaList("PrefMultiValList").Exist = True Then
                                        bFlag = Fn_UI_ListItemExist("Fn_SISW_Pref_Organization_Operation", objPrefOper, "PrefMultiValList", sValues)
				Else
				
'					If objPrefOper.JavaEdit("CurrentValues").GetROProperty("value") = sValues Then
					
					If Fn_UI_Object_GetROProperty("Fn_SISW_Pref_Organization_Operation", objPrefOper.JavaEdit("CurrentValues"),"value") = sValues Then
						bFlag = True
					Else
						bFlag = False
					End If
				End If
			End If

			Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Close")
			wait(5)

			'Log for Success
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Verification on Preference failed ")                                        
				Fn_SISW_Pref_Organization_Operation = False                                                                                                   
			Else
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Verification on Preference Passed Successfully")
				Fn_SISW_Pref_Organization_Operation = True    
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Delete"
			'	clikcing on "Details" tab
			'objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Details"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"
		
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper,"BottomLink",10, 10,"LEFT")

			'Set value in Name Edit  box
			If Trim(sPrefName) <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_Organization_Operation",objPrefOper,"SrchPrefName",sPrefName)
			End If
			'Call Fn_List_Select("Fn_SISW_Pref_Organization_Operation", objPrefOper, "SrchPrefList",sSrchText)
			If Trim(sSrchText ) <> "" Then
				'Select Preference
				iItemCnt = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
				For iCnt = 0 to iItemCnt - 1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sSrchText Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
						bFlag = True
						wait 2
						Exit for
					End If
				Next
				If bFlag = False Then
					Fn_SISW_Pref_Organization_Operation = False
					Exit Function
				End If
			End If
			Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "DeletePreference")
			Set objdelete =  Fn_SISW_Pref_GetObject("DeletePreference" )
			If objdelete.Exist =True Then
				Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objdelete, "Yes")   
				bFlag = TRUE
			End If 
			Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Close")
			wait(5)
			'Log for Success
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Verification on Preference failed ")
				Fn_SISW_Pref_Organization_Operation = False
			Else
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Verification on Preference Passed Successfully")
				Fn_SISW_Pref_Organization_Operation = True    
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'TC11.3(20170509d.00)_DIPRO_Development_PoonamC_27Oct2017 : Added Case "CreateGroupLevelPreferenceWithRemoveValue"
		Case "CreateGroupLevelPreference","CreateGroupLevelPreferenceWithRemoveValue"
						
			'Select Organization bottom link
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper,"Organization",0,0,"LEFT")
			
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "Refresh")
			wait(3)
			
			'Set Preference Name
			objPrefOper.JavaEdit("SrchPrefName").Set ""
			objPrefOper.JavaEdit("SrchPrefName").Type sPrefName
			wait(3)
			
			If Trim(sPrefName ) <> "" Then
				'Select Preference
				iItemCnt = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
				For iCnt = 0 to iItemCnt - 1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sPrefName Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
						bFlag = True
						wait 2
						Exit for
					End If
				Next
				If bFlag = False Then
					Fn_SISW_Pref_Organization_Operation = False
					Exit Function
				End If
			End If
			wait 3
			
			'Select Instances 
			Wait 1
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper,"Instances",0, 0,"LEFT")
			
			'Select User / Group / Role
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper,"Organization",0,0,"LEFT")
			If Trim(sUser) <> "" Then
				aUserList = Split(sUser, ":")
				For iCounter = 0 To UBound(aUserList)
					If iCounter = 0  Then
						sPath = aUserList(iCounter)
					Else
						sPath = sPath + ":"+aUserList(iCounter)
					End If
					Call Fn_JavaTree_Select("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Organization", sPath)
					Call Fn_UI_JavaTree_Expand("Fn_SISW_Pref_Organization_Operation",objPrefOper,"Organization",sPath)
					wait SISW_MICRO_TIMEOUT
				Next
				Call Fn_JavaTree_Select("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Organization", sUser)
				wait 1
			End If
			
			'Click on 'CreateNewPreference' Button
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "CreateNewPreference")
			wait 1
			
			If sAction = "CreateGroupLevelPreferenceWithRemoveValue" Then	
					'Set New Value
					If Trim(sValues) <> "" Then	
						If instr(sValues,":") > 0 Then
							ArrValue = Split(sValues,":",-1)
							For iCnt=0 to Ubound(ArrValue)
								Call Fn_List_Select("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList",ArrValue(iCnt))
								Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefRemoveValue")
								bFlag = True							
							Next		
						Else
							Call Fn_List_Select("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList",sValues)
							Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefRemoveValue")
							bFlag = True
						End If
				    End if
					wait 2				    
			Else
					'Set New Value
					If Trim(sValues) <> "" Then
						If Lcase(sValues) = "blank" Then
							Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CurrentValues","")
							If  objPrefOper.JavaButton("NewMultiValPrefAddValue").Exist Then
								Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewMultiValPrefAddValue")
							End If
						ElseIf instr(sValues,":") > 0 Then
							ArrValue = Split(sValues,":",-1)
							For iCnt=0 to Ubound(ArrValue)
								Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CurrentValues",ArrValue(iCnt))
								wait 1
								Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewMultiValPrefAddValue")
								wait 1								
							Next		
						Else
							Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"CurrentValues",sValues)
							If  objPrefOper.JavaButton("NewMultiValPrefAddValue").Exist(5) Then
								Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "NewMultiValPrefAddValue")
							End If
						End If
					End If
					wait 2
			End If
			
'			Set objPrefInstance = Fn_SISW_GetObject("PreferenceInstance")
			If JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("PreferenceInstance").Exist(5) Then
				Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", JavaWindow("DefaultWindow").JavaWindow("WEmbeddedFrame").JavaDialog("PreferenceInstance"), "Cancel")
			End If
			wait 2
			
			'Select Save button
			Call Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Save")

			'Select Close button
			Fn_SISW_Pref_Organization_Operation = Fn_Button_Click("Fn_SISW_Pref_Organization_Operation", objPrefOper, "Close")

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 	
	End Select

	Set objList = Nothing
	Set oDesc = Nothing
	Set aDelList = Nothing
	Set jCounter = Nothing
	Set aUserList = Nothing
	Set iCounter = Nothing
	Set objSelectType = Nothing
	Set objDialog = Nothing
	Set bFlag = Nothing
	Set WshShell  = Nothing
	Set objPrefOper = Nothing
	Set sPath = Nothing
End Function

'*********************************************************		Fn_SISW_Pref_MultiValue_Operation 		***********************************************************************

'Function Name		:					Fn_SISW_Pref_MultiValue_Operation(sAction, sPrefName, sDesc, bScope, sCategory, sType, sPrefValue)

'Description			 :		 		  This function handles the error message which pops up during CUT operation.

'Parameters			   :	 			sAction: "CreateNew" or "AddValue" or "RemoveValue" or "VetrifyValues"
'                                       sPrefName: Preference Name
'                                       sDesc: Description
'                                       bScope: User or Role or Group or Site
'                                        sCategory: General or Configuration.Reports
'                                        sType: String or date or ........
'                                      sPrefValue: Preference Value (if multiple values need to send use "~" as delimiter)

'Return Value		   : 				 True/False

'Pre-requisite			:		 		

'Examples				:			     'Call Fn_SISW_Pref_MultiValue_Operation("AddValue", "", "", "", "", "", "ItemRevision")
													'Call Fn_SISW_Pref_MultiValue_Operation("AddValue","TC_ValidApprovedStatus", "", "Site", "", "", "TCM Released")
													'Call Fn_SISW_Pref_MultiValue_Operation("CreateNew", "TC_ValidApprovedStatus", "", "Site", "", "", "Released~TCM Released")
													'Call Fn_SISW_Pref_MultiValue_Operation("VerifyValues","TC_ValidApprovedStatus", "", "Site", "", "", "Released")
													'Call Fn_SISW_Pref_MultiValue_Operation("RemoveAllList", "BOMCompareVisibleModes", "", "", "", "", "")

'History:
'	Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Mahendra			4-Jun-2010   		1.0												Rizwan
'	Sachin Joshi		19-Oct-2010			1.0				Modified "Add Value Case"   	Manisha A.
'	Koustubh W			13-Apr-2011			1.0				Modified "Remove Value Case"
'	Sandeep N			21-March-2011		1.1				Added Case "RemoveValueExt"			Priyanka B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh W			13-Jun-2012			2.0				Modified function according to Teamcenter 10.0 UI changes
'	Sandeep N			10-Jul-2012			2.1				modified case : "RemoveValue","RemoveValueExt" Added code to handle numeric values 
'	Sandeep N			13-Jul-2012			2.2				modified case : "VerifyValues"  initially Set bFlag = False 
'	Koustubh W			17-Jul-2012			2.3				Modified case "VerifyValues"
'	Koustubh W			31-Jul-2012			2.3				Modified case "AddValues"
'  	Sandeep N			11-Sep-2012			2.4				Added code to handle [ Multiple ] static Text 
'	Vrushali W			22-Mar-2013			2.5				Added "Modify" Case
'	Shantan S			14-Dec-2015			2.5				Added "VerifyBlank" Case
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Vivek A				20-Aug-2015			2.7				Added Cases "VerifyMultiplePrefrencesValue", "ModifyMultiplePrefValueScope", "RemoveMultiplePrefValues" to work on multiple Preferences and single or multiple values
'						[TC1015-2015072100-20_08_2015-VivekA-NewDevelopment]
'						Use ~ for multiple preferences and $ seperated for wild card parameter  
'														 :  PrefeNames = "TcNotesAllowedType$TcNotesAllowedType_DesignReq~TcNotesAllowedType_Fnd0LogicalBlock"
'						Use # for pref having multiple values and ~ for diff pref values 
'														 :	PrefeValues = "Fnd0CustomNote#edut~Fnd0CustomNote"
'						Use ~ for diff pref Scopes
'														 :	bScope = "Site~Site"
'						Use ~ for diff pref Descriptions 
'														 :	sDesc = "abc~abc"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Pref_MultiValue_Operation(sAction, sPrefName, sDesc, bScope, sCategory, sType, sPrefValue)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_MultiValue_Operation"
	Dim objPrefOper, objSelectType, objDialog, bFlag, iCountChecked, oDesc, iCounter2, iCounter, aListArray, listCount, winPrefName, sPrefValues, sPrefValuePresent
	Dim aPrefValue,iDataCounter, iCounter1,prefVal
	Dim iItemCnt,iCnt,location,aType,sEnvironmet
	Dim sPreferences, sWildCard, sAllPrefNames, sAllPrefValues, bFlag1, sAllPrefScopes, sAllPrefDescs
	Dim DicPrefValues
	Set DicPrefValues = CreateObject("Scripting.Dictionary")
	
	bFlag = False
	Fn_SISW_Pref_MultiValue_Operation = False
	'Menu operation function called to select Option from Edit  Menu.
	If Fn_SISW_Pref_GetObject("IndexOptions").Exist =False Then
		If Fn_SISW_Pref_GetObject("IndexOptions2").Exist =False Then
			'IF not exist then opening from menu
			Call Fn_MenuOperation("Select","Edit:Options...")
			Call Fn_ReadyStatusSync(2)
		End If
	End If

	If Fn_SISW_Pref_GetObject("IndexOptions").Exist Then
		Set objPrefOper = Fn_SISW_Pref_GetObject("IndexOptions")
	ElseIf Fn_SISW_Pref_GetObject("IndexOptions2").Exist Then
		Set objPrefOper = Fn_SISW_Pref_GetObject("IndexOptions2")
	End If
	Wait(1)
	Set objPrefOper = Fn_UI_ObjectCreate("Fn_SISW_Pref_MultiValue_Operation",objPrefOper)
	Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper,"Index",0,0,"LEFT")
	
	Select Case sAction
		Case "CreateNew","VerifyMultiplePrefrencesValue","ModifyMultiplePrefValueScope","RemoveMultiplePrefValues"
			' do nothing
		Case Else
			'clikcing on "Details" tab
			'objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Details"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"
			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"BottomLink",10, 10,"LEFT")
	
			Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"SrchPrefName", sPrefName)
			Wait(3)
			'Select Preference
			iItemCnt = cInt(objPrefOper.JavaTable("PreferencesListTable").GetROProperty("rows"))
			For iCnt = 0 to iItemCnt - 1
				If sAction <> "VerifyValuesWithScope" Then
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") =  sPrefName Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
						bFlag = True
						wait 2
						Exit for
					End If
				Else
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Name") = sPrefName AND LCase(Trim(objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCnt,"Location"))) = LCase(Trim(bScope)) Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCnt
						bFlag = True
						wait 2
						Exit for
					End If
				End If
			Next
			If bFlag = False Then
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Fn_SISW_Pref_MultiValue_Operation: Failed to find preference [ " & sPrefName & " ].")
			End If
	End Select

	Select Case sAction
		Case "CreateNew"
			'Set value in Name Edit  box
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "CreateNewPreference")
			
			'Set value in Name Edit  box
			Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"Name",sPrefName)
			
        	'Select Multiple value
			If bScope <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefScopeDrpDwn")
				iCounter1 = 0
				If objPrefOper.JavaStaticText("ValueLocation").Exist(1) Then
					location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
					If lcase(trim(bScope)) = location Then
						iCounter1 = 1
					End If
				End If
				wait 2
				' add code to select static text of scope
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = bScope
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(iCounter1).Click 5, 5, "LEFT"
				wait 1
			End If

			   'Select category
			If Trim(sCategory) <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefCategoryDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sCategory
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
				wait 1
			End If

			'Select Multiple value
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultipleDrpDwn")
			wait 2
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = "Multiple"
			Set objDialog = objPrefOper.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
			Set objDialog =nothing
			wait 2
			If Fn_Edit_Box_GetValue("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"Multiple")<>"Multiple" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultipleDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = "Multiple"
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(1).Click 5, 5, "LEFT"
				Set objDialog =nothing
			End If

	'----------------------------------------------------------------------------------------------	
			'Added code to set value for Environment dropdown - By Jotiba T
			If Instr(1,sType,"~") Then
				aType=Split(sType,"~")
				If UBound(aType) > 0 Then
					sType=aType(0)
					sEnvironmet=aType(1)
				End If
			End If
			
               'Select Environmet                           
			If Trim(sEnvironmet)<>"" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefEnvironmentDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = trim(sEnvironmet)
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
				Set objDialog =nothing
				wait 2
				If Fn_Edit_Box_GetValue("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"Environment")<>trim(sEnvironmet) Then
					Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefEnvironmentDrpDwn")
					wait 2
					Set objSelectType=description.Create()
					objSelectType("Class Name").value = "JavaStaticText"
					objSelectType("label").value = trim(sEnvironmet)
					Set objDialog = objPrefOper.ChildObjects(objSelectType)
					objDialog(1).Click 5, 5, "LEFT"
					Set objDialog =nothing
				End If
			End If
	'----------------------------------------------------------------------------------------------
			'Select Type
			If Trim(sType) <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewPrefTypeDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sType
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
				wait 1
			End If

			'Set value in Description Edit  box
			If Trim(sDesc) <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"Description",sDesc)
			End If
				
			'Set multiple values.
			If Trim(sPrefValue) <> "" Then
				aPrefValue = split(sPrefValue,"~",-1,1)
				For iDataCounter=0 to Ubound(aPrefValue)
					Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"CurrentValues",aPrefValue(iDataCounter))
					Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefAddValue")					
				Next
			End If

			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
			wait(5)
			bFlag = Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog("Preference already exists for the scope. You can modify the value(s) using the Details tab.")
			If bFlag = False Then
				Fn_SISW_Pref_MultiValue_Operation = True
			Else
				Fn_SISW_Pref_MultiValue_Operation = bFlag
			End If
			
			If objPrefOper.JavaButton("Cancel").Exist = True Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Cancel")
			End If
			If objPrefOper.JavaButton("Close").Exist = True Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
			End If
			'Log for Success				   
				
			
	' - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - - 
	Case "AddValue"
		Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"Definition",10, 10,"LEFT")
		wait 1,500
		If objPrefOper.JavaButton("Edit").Exist(5) Then
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Edit")
		End If
		'Set value in Description Edit  box
		If Trim(sDesc) <> "" Then
			Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"Description",sDesc)
		End If
	   'Select Multiple value
	   'objPrefOper.JavaRadioButton("NewPrefScope").Set StrScope
		If bScope <> "" Then
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefScopeDrpDwn")
			iCounter1 = 0
			If objPrefOper.JavaStaticText("ValueLocation").Exist(1) Then
				location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
				If lcase(trim(bScope)) = location Then
					iCounter1 = 1
				End If
			End If
			wait 2
			' add code to select static text of scope
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = bScope
			Set objDialog = objPrefOper.ChildObjects(objSelectType)
			objDialog(iCounter1).Click 5, 5, "LEFT"
		End If

	   'Select category
		If Trim(sCategory) <> "" Then
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefCategoryDrpDwn")
			wait 2
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sCategory
			Set objDialog = objPrefOper.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
		End If

		'Select Type
		If Trim(sType) <> "" Then
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewPrefTypeDrpDwn")
			wait 2
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sType
			Set objDialog = objPrefOper.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
		End If	

		'Modified by Rupali 03-Sept-10 for multiple value of preference.
		'Set multiple values.
		wait(2)
		If Trim(sPrefValue) <> "" Then
			aPrefValue = split(sPrefValue,"~",-1,1)
			For iDataCounter=0 to Ubound(aPrefValue)
				If Fn_UI_ListItemExist("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList",aPrefValue(iDataCounter)) <> True Then
					Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"CurrentValues",aPrefValue(iDataCounter))
					Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefAddValue")					
				End If
			Next
		End If
		If CInt (Fn_UI_Object_GetROProperty("Fn_SISW_Pref_MultiValue_Operation",objPrefOper.JavaButton("Save"), "enabled")) = 1 Then
			Call  Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
			Call Fn_ReadyStatusSync(2)
		Else
			Call  Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Cancel")
		End If
	
		bFlag = Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
		Fn_SISW_Pref_MultiValue_Operation = bFlag
	' - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - - 
	Case "RemoveValue","RemoveValueExt"
		If objPrefOper.JavaButton("Edit").Exist(5) Then
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Edit")
		End If
		'Added Case "RemoveValueExt" to remove the Preference value having colon ( : ) in text
		'Added by sandeep Navghane
		If sAction="RemoveValueExt" Then
			sPrefValues=Split(sPrefValue,"~")
		Else
			sPrefValues=Split(sPrefValue,":")
		End If

		For iCounter=0 To UBound(sPrefValues)
			'Added code to handle numeric values - - - - - - - - - Added by sandeep 10-Jul-2012
			If isnumeric(sPrefValues(iCounter)) Then
				prefVal=Cint(sPrefValues(iCounter))
			else
				prefVal=sPrefValues(iCounter)
			End If
			If Fn_UI_ListItemExist("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList",prefVal) <> False Then
				objPrefOper.JavaList("PrefMultiValList").ExtendSelect sPrefValues(iCounter)
				wait 1
			End If
		Next

		bFlag = Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefRemoveValue")
		If CInt (Fn_UI_Object_GetROProperty("Fn_SISW_Pref_MultiValue_Operation",objPrefOper.JavaButton("Save"), "enabled")) = 1 Then
			Call  Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
		Else
			Call  Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Cancel")
		End If
		
		'Button click call added by Amol - 09-11-10
		Call  Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
		
		Fn_SISW_Pref_MultiValue_Operation = bFlag 
	' - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - - 
	'Case added by Rupali on 03-Sept-2010 for Remove all value fron Current Value Javalist of Index Option.
	Case "RemoveAllList"
		If objPrefOper.JavaButton("Edit").Exist(5) Then
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Edit")
		End If
		listCount = objPrefOper.JavaList("PrefMultiValList").GetROProperty("items count")
		If Cint(listCount) = 0 Then
			bFlag = True	
		End If
		
		For iCounter = Cint(listCount -1) To 0 Step -1
			winPrefName = objPrefOper.JavaList("PrefMultiValList").GetItem(iCounter)
			If winPrefName <> "" Then
				Call Fn_List_Select("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList",winPrefName)
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefRemoveValue")
				bFlag = True	
			End If			
		Next
		Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
	
		Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")

		Fn_SISW_Pref_MultiValue_Operation = bFlag
	' - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - -  - - - - - - - 
	Case "VerifyValues","VerifyValuesWithScope","VerifyValuesExt"
		bFlag = False
		If Trim(sPrefValue) <> "" Then
			'[TC1123(20161205c00)_NewDevelopment_PoonamC_22Mar2017:Added case "VerifyValuesExt" with using separator as "~" if there is pref value contains ":"]
			If sAction = "VerifyValuesExt" Then
				sPrefValues = Split(sPrefValue, "~")
			Else
				sPrefValues = Split(sPrefValue, ":")
			End If
			For iCounter = 0 To UBound(sPrefValues)
				iCountChecked = cInt(objPrefOper.JavaList("PrefMultiValList").GetROProperty("items count"))
				objPrefOper.JavaList("PrefMultiValList").Object.setEnabled(True)
				For iCounter2 = 0 To iCountChecked-1
					sPrefValuePresent = objPrefOper.JavaList("PrefMultiValList").GetItem(iCounter2)
					If sPrefValuePresent = sPrefValues(iCounter) Then
						bFlag = True
						Exit For
					End If			
				Next
				If bFlag = False Then
					Exit For
				End If
			Next
		End If
		objPrefOper.JavaList("PrefMultiValList").Object.setEnabled(False)
		Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
		Fn_SISW_Pref_MultiValue_Operation = bFlag
    Case "Modify"    'Added by Vrushali W on 22_Mar-2013
            objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Definition"
				If objPrefOper.JavaStaticText("BottomLink").Exist(5) = False Then
					objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Instances"
				End If
				Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper,"BottomLink",10, 10,"LEFT")

              If bFlag Then
				'click on Edit Button
				If objPrefOper.JavaButton("Edit").Exist(3) Then
					Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Edit")
				End If
		
				If bScope <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefScopeDrpDwn")
				iCounter1 = 0
				If objPrefOper.JavaStaticText("ValueLocation").Exist(1) Then
					location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
					If lcase(trim(bScope)) = location Then
						iCounter1 = 1
					End If
				End If
				wait 2
				' add code to select static text of scope
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = bScope
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(iCounter1).Click 5, 5, "LEFT"
			End If

			   'Select category
			If Trim(sCategory) <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefCategoryDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sCategory
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			'Select Multiple value
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultipleDrpDwn")
			wait 2
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = "Multiple"
			Set objDialog = objPrefOper.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
			Set objDialog =nothing
			wait 2
			If Fn_Edit_Box_GetValue("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"Multiple")<>"Multiple" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultipleDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = "Multiple"
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(1).Click 5, 5, "LEFT"
				Set objDialog =nothing
			End If

			'Select Type
			If Trim(sType) <> "" Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewPrefTypeDrpDwn")
				wait 2
				Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sType
				Set objDialog = objPrefOper.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
			End If

			'Set value in Description Edit  box
			If Trim(sDesc) <> "" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"Description",sDesc)
			End If
				
			'Set multiple values.
			If Trim(sPrefValue) <> "" Then
				aPrefValue = split(sPrefValue,"~",-1,1)
				For iDataCounter=0 to Ubound(aPrefValue)
					Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"CurrentValues",aPrefValue(iDataCounter))
					Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefAddValue")					
				Next
			End If

			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
			wait(5)
			bFlag = Fn_SISW_Pref_VerifyMsgCreatePreferenceDialog("Preference already exists for the scope. You can modify the value(s) using the Details tab.")
			If bFlag = False Then
				Fn_SISW_Pref_MultiValue_Operation = True
			Else
				Fn_SISW_Pref_MultiValue_Operation = bFlag
			End If
			
			If objPrefOper.JavaButton("Cancel").Exist = True Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Cancel")
			End If
		  End If
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
		'[TC1121-20151116b-14_12_2015-VivekA-NewDevelopment] - Added by Shantan S - to verify blank value in Preference
		Case "VerifyBlank"
			bFlag = False
			iCountChecked = cInt(objPrefOper.JavaList("PrefMultiValList").GetROProperty("items count"))		
			If iCountChecked = 0 Then
				bFlag = True
			End If			
			objPrefOper.JavaList("PrefMultiValList").Object.setEnabled(False)
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
			Fn_SISW_Pref_MultiValue_Operation = bFlag
			
		'-------Verifying multiple Preferences with single or multiple values------[TC1015-2015072100-20_08_2015-VivekA-NewDevelopment]		
		Case "VerifyMultiplePrefrencesValue"
			bFlag = False		
			sPreferences = Split(sPrefName,"$")
			sWildCard = sPreferences(0)
			sAllPrefNames = Split(sPreferences(1),"~")
			sAllPrefValues = Split(sPrefValue,"~")
			
			'Enter Wild card string (Pref Name) in edit box
			Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"SrchPrefName",sWildCard)
			
			'Get number of rows in Preferences List Table
			iItemCnt = Fn_UI_Object_GetROProperty("Fn_SISW_Pref_MultiValue_Operation",objPrefOper.JavaTable("PreferencesListTable"),"rows")
			'If no Preferences found
			If Cint(iItemCnt) = 0 Then
				Fn_SISW_Pref_MultiValue_Operation = False
				Exit Function 
			End If
			
			'For All pref one by one
			For iCnt = 0 To UBound(sAllPrefNames)
				bFlag1 = False			
				'For All rows one by one
				For iCounter = 0 To iItemCnt-1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCounter,"Name") = sAllPrefNames(iCnt) Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCounter
						objPrefOper.JavaList("PreferenceValueRendererCompone").object.setEnabled(True)
						
						aPrefValue = split(sAllPrefValues(iCnt),"#",-1,1)
						For iCounter2=0 to Ubound(aPrefValue)
							If isnumeric(aPrefValue(iCounter2)) Then
								prefVal = Cint(aPrefValue(iCounter2))
							Else
								prefVal = aPrefValue(iCounter2)
							End If
							bFlag1 = Fn_UI_ListItemExist("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList", prefVal)
							If bFlag1 = False Then
								Exit For
							End If
						Next		
						
						If bFlag1 = False Then
							Fn_SISW_Pref_MultiValue_Operation = False
							Exit Function
						Else
							Exit For
						End If
					ElseIf iCounter = iItemCnt-1 Then
						bFlag1 = False
					End If		
				Next
				
				If bFlag1 = False Then
					Fn_SISW_Pref_MultiValue_Operation = False
					Exit Function
				Else
					bFlag = True
				End If				
			Next
			
			If bFlag = True Then
				Fn_SISW_Pref_MultiValue_Operation = True
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
			Else
				Fn_SISW_Pref_MultiValue_Operation = False
			End If
			
		'-------Modifying multiple Preferences with single or multiple values------[TC1015-2015072100-20_08_2015-VivekA-NewDevelopment]
		Case "ModifyMultiplePrefValueScope"
			objPrefOper.JavaStaticText("BottomLink").setTOProperty "label", "Definition"
			bFlag = False		
			sPreferences = Split(sPrefName,"$")
			sWildCard = sPreferences(0)
			sAllPrefNames = Split(sPreferences(1),"~")
			sAllPrefValues = Split(sPrefValue,"~")
			sAllPrefScopes = Split(bScope,"~")
			sAllPrefDescs = Split(sDesc,"~")
			'Enter Wild card string (Pref Name) in edit box
			Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"SrchPrefName",sWildCard)
			Wait 1
			'Get number of rows in Preferences List Table
			iItemCnt = Fn_UI_Object_GetROProperty("Fn_SISW_Pref_MultiValue_Operation",objPrefOper.JavaTable("PreferencesListTable"),"rows")
			'If no Preferences found
			If Cint(iItemCnt) = 0 Then
				Fn_SISW_Pref_MultiValue_Operation = False
				Exit Function 
			End If
			
			'Verify all pref present
			For iCnt = 0 To UBound(sAllPrefNames)
				bFlag1 = False
				For iCounter = 0 To iItemCnt-1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCounter,"Name") = sAllPrefNames(iCnt) Then
						bFlag1 = True
						Exit For
					ElseIf iCounter = iItemCnt-1 Then
						bFlag1 = False
					End If
				Next
				If bFlag1 = False Then
					Fn_SISW_Pref_MultiValue_Operation = False
					Exit Function
				Else
					bFlag = True
				End If	
			Next
			If bFlag=False Then
				Fn_SISW_Pref_MultiValue_Operation = False
				Exit Function
			End If
			
			bFlag = False
			'For All pref one by one
			For iCnt = 0 To UBound(sAllPrefNames)
				bFlag1 = False			
				'For All rows one by one
				For iCounter = 0 To iItemCnt-1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCounter,"Name") = sAllPrefNames(iCnt) Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCounter
						
						'click on Edit Button
						If objPrefOper.JavaButton("Edit").Exist(3) Then
							Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Edit")
							Call Fn_ReadyStatusSync(1)
						Else
							Fn_SISW_Pref_MultiValue_Operation = False
							Exit Function
						End If
						
						If bScope <> "" Then
							Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefScopeDrpDwn")
							iCounter1 = 0
							If objPrefOper.JavaStaticText("ValueLocation").Exist(1) Then
								location = lcase(trim(objPrefOper.JavaStaticText("ValueLocation").getROProperty("label")))
								If lcase(trim(sAllPrefScopes(iCnt))) = location Then
									iCounter1 = 1
								End If
							End If
							wait 2
							' add code to select static text of scope
							Set objSelectType=description.Create()
							objSelectType("Class Name").value = "JavaStaticText"
							objSelectType("label").value = sAllPrefScopes(iCnt)
							Set objDialog = objPrefOper.ChildObjects(objSelectType)
							objDialog(iCounter1).Click 5, 5, "LEFT"
						End If
						'Set value in Description Edit  box
						If Trim(sDesc) <> "" Then
							Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"Description",sAllPrefDescs(iCnt))
						End If
				
						'Set multiple values.
						If sPrefValue <> "" Then
							aPrefValue = split(sAllPrefValues(iCnt),"#",-1,1)
							For iDataCounter=0 to Ubound(aPrefValue)
								Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"CurrentValues",aPrefValue(iDataCounter))
								Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefAddValue")					
							Next
						End If

						Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
						wait(5)
						
						bFlag1 = True
						Exit For
					ElseIf iCounter = iItemCnt-1 Then
						bFlag1 = False
					End If		
				Next
				
				If bFlag1 = False Then
					Fn_SISW_Pref_MultiValue_Operation = False
					Exit Function
				Else
					bFlag = True
				End If				
			Next
			
			If bFlag = True Then
				Fn_SISW_Pref_MultiValue_Operation = True
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
			Else
				Fn_SISW_Pref_MultiValue_Operation = False
			End If
			
		'-------Removing multiple Preferences with single or multiple values------[TC1015-2015072100-20_08_2015-VivekA-NewDevelopment]
		Case "RemoveMultiplePrefValues"
			bFlag = False		
			sPreferences = Split(sPrefName,"$")
			sWildCard = sPreferences(0)
			sAllPrefNames = Split(sPreferences(1),"~")
			sAllPrefValues = Split(sPrefValue,"~")
			
			'Enter Wild card string (Pref Name) in edit box
			Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"SrchPrefName",sWildCard)
			
			'Get number of rows in Preferences List Table
			iItemCnt = Fn_UI_Object_GetROProperty("Fn_SISW_Pref_MultiValue_Operation",objPrefOper.JavaTable("PreferencesListTable"),"rows")
			'If no Preferences found
			If Cint(iItemCnt) = 0 Then
				Fn_SISW_Pref_MultiValue_Operation = False
				Exit Function 
			End If
			
			'Verify all pref present
			For iCnt = 0 To UBound(sAllPrefNames)
				bFlag1 = False
				For iCounter = 0 To iItemCnt-1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCounter,"Name") = sAllPrefNames(iCnt) Then
						bFlag1 = True
						Exit For
					ElseIf iCounter = iItemCnt-1 Then
						bFlag1 = False
					End If
				Next
				If bFlag1 = False Then
					Fn_SISW_Pref_MultiValue_Operation = False
					Exit Function
				Else
					bFlag = True
				End If	
			Next
			If bFlag=False Then
				Fn_SISW_Pref_MultiValue_Operation = False
				Exit Function
			End If
			
			bFlag = False
			'For All pref one by one
			For iCnt = 0 To UBound(sAllPrefNames)
				bFlag1 = False			
				'For All rows one by one
				For iCounter = 0 To iItemCnt-1
					If objPrefOper.JavaTable("PreferencesListTable").GetCellData(iCounter,"Name") = sAllPrefNames(iCnt) Then
						objPrefOper.JavaTable("PreferencesListTable").SelectRow iCounter
						
						'click on Edit Button
						If objPrefOper.JavaButton("Edit").Exist(3) Then
							Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Edit")
						Else
							Fn_SISW_Pref_MultiValue_Operation = False
							Exit Function
						End If
				
						'Remove multiple values.
						If sPrefValue <> "" Then
							aPrefValue = split(sAllPrefValues(iCnt),"#",-1,1)
							For iDataCounter=0 to Ubound(aPrefValue)
								If isnumeric(aPrefValue(iDataCounter)) Then
									prefVal = Cint(aPrefValue(iDataCounter))
								Else
									prefVal = aPrefValue(iDataCounter)
								End If
								If Fn_UI_ListItemExist("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList", prefVal) <> False Then
									objPrefOper.JavaList("PrefMultiValList").ExtendSelect aPrefValue(iDataCounter)
									wait 1
								End If			
							Next
							bFlag1 = Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefRemoveValue")		
						End If
						If CInt(Fn_UI_Object_GetROProperty("Fn_SISW_Pref_MultiValue_Operation",objPrefOper.JavaButton("Save"), "enabled")) = 1 Then
							Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
						Else
							Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Cancel")
						End If	
						wait(5)
						Exit For
					ElseIf iCounter = iItemCnt-1 Then
						bFlag1 = False
					End If		
				Next
				
				If bFlag1 = False Then
					Fn_SISW_Pref_MultiValue_Operation = False
					Exit Function
				Else
					bFlag = True
				End If				
			Next
			
			If bFlag = True Then
				Fn_SISW_Pref_MultiValue_Operation = True
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
			Else
				Fn_SISW_Pref_MultiValue_Operation = False
			End If
		'====== [TC1123(20161205c00)_NewDevelopment_PoonamC_22Mar2017:Added case "CreateNewInstanceFrmExistingPref" to create new instance of Pref from existing pref] =====
		' Create New Instance from existing preference		
		 Case "CreateNewInstanceFrmExistingPref"
			'Set value in Name Edit  box
			Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper,"CreateNewPreference")
			
			If vartype(sPrefValue) = "9" Then
				Set DicPrefValues = sPrefValue
			End If
			
			'Remove existing multiple values.
			If DicPrefValues("RemoveExistingValues") <> "" Then
						aPrefValue = split(DicPrefValues("RemoveExistingValues"),"~")
						For iDataCounter=0 to Ubound(aPrefValue)
							If Fn_UI_ListItemExist("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "PrefMultiValList", aPrefValue(iDataCounter)) <> False Then
								objPrefOper.JavaList("PrefMultiValList").ExtendSelect aPrefValue(iDataCounter)
								wait 1
							End If			
						Next
						bFlag1 = Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefRemoveValue")		
			End If
			
			'Add New multiple values.
			If DicPrefValues("AddNewValues") <> "" Then
						aPrefValue = split(DicPrefValues("AddNewValues"),"~")
						For iDataCounter=0 to Ubound(aPrefValue)
							Call Fn_Edit_Box("Fn_SISW_Pref_MultiValue_Operation",objPrefOper,"CurrentValues",aPrefValue(iDataCounter))
							Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "NewMultiValPrefAddValue")					
						Next		
			End If			
			
			bFlag = Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Save")
			wait(5)
			If bFlag = True Then
				Fn_SISW_Pref_MultiValue_Operation = True
			Else
				Fn_SISW_Pref_MultiValue_Operation = bFlag
			End If
			
			If objPrefOper.JavaButton("Cancel").Exist = True Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Cancel")
			End If
			If objPrefOper.JavaButton("Close").Exist = True Then
				Call Fn_Button_Click("Fn_SISW_Pref_MultiValue_Operation", objPrefOper, "Close")
			End If
			
	End Select
	
	If Fn_SISW_Pref_MultiValue_Operation <> False Then
    	Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS: Fn_SISW_Pref_MultiValue_Operation: Executed successfully with action [ " & sAction & " ].")
	End If
	Set objPrefOper  =Nothing
	Set objSelectType = Nothing
	Set objDialog = Nothing
	Set bFlag = Nothing
	Set iCountChecked = Nothing
	Set oDesc = Nothing
	Set iCounter2 = Nothing
	Set iCounter = Nothing
	Set aListArray = Nothing
	Set DicPrefValues = Nothing
End Function

'#####################################################			To Create the new search Preference.		##########################################################~
'#
'# FUNCTION NAME:	Fn_SISW_Pref_Search_CreateOperation
'#
'# FUNCTION ID:			336
'#
'# MODULE: 					My Teamcenter
'# 				
'# DESCRIPTION:		To Create the new search Preference.
'#									
'#PARAMETERS   :  	sName: 				Name of the new preference
'#									  	sDescription:			Description of the new preference
'#										sScopes:		"User" / "Role" / "Group" / "Site"
'#										sCatagory:		of the new preference
'#										sType:			of the new preference
'#										bMulValues: 	"True"/"False"
'#										sValues:	
'#
'# RETURN VALUE : 		TRUE \ FALSE
'#
'#Examples	:				 Fn_SISW_Pref_Search_CreateOperation("pref1", "Test Preference", "Role", "Workflow", "Double", "True", "1")							
'#										
'#	History	:					
'#	Developer Name			Date		Version				Changes Done			Reviewer	
'#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Kavan Shah~			June@2010		1.0											Sunil Rai
'#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'#	Koustubh Watwe~		June@2012		2.0				Modified function according to TC10.0 UI Changes
'#####################################################			To Create the new search Preference.		##########################################################~
Public Function Fn_SISW_Pref_Search_CreateOperation(sName, sDescription, sScopes, sCategory, sType, bMulValues, sValues)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_Search_CreateOperation"
    Dim oOptnsWndw, objSearch, objIntNoOfObjects, objSelectType
	Dim bFound, bReturn, innerCntr, iCounter1, location, ArrValue
	Dim BlnMultiVal, StrType

	Set oOptnsWndw  = Fn_SISW_Pref_GetObject("IndexOptions")
	If oOptnsWndw.Exist(5) = False Then
		Call Fn_MenuOperation("Select","Edit:Options...")
		Call Fn_ReadyStatusSync(1)
		Set oOptnsWndw  = Fn_UI_ObjectCreate( "Fn_SISW_Pref_Search_CreateOperation", oOptnsWndw)
	End If
	Call Fn_ReadyStatusSync(2)
	Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_Search_Operation", oOptnsWndw,"Search",0,0,"LEFT")
	Call Fn_ResizeWindow("Resize","700", "800", oOptnsWndw)

	'++++++++++<<    Set Name / Description >>++++++++++
        Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", oOptnsWndw, "CreateNewPreference")

        Call Fn_Edit_Box("Fn_SISW_Pref_Search_CreateOperation",oOptnsWndw,"Name",sName)
	'++++++++++<<    Set Scopes/ Multiple Values >>++++++++++
	If sScopes <> "" Then
		Call Fn_Button_Click("Fn_SISW_Pref_Search_CreateOperation", oOptnsWndw, "PrefScopeDrpDwn")
		iCounter1 = 0
		If oOptnsWndw.JavaStaticText("ValueLocation").Exist(1) Then
			location = lcase(trim(oOptnsWndw.JavaStaticText("ValueLocation").getROProperty("label")))
			If lcase(trim(sScopes)) = location Then
				iCounter1 = 1
			End If
		End If
		wait 2
		' add code to select static text of scope
		Set objSelectType=description.Create()
		objSelectType("Class Name").value = "JavaStaticText"
		objSelectType("label").value = sScopes
		Set objIntNoOfObjects = oOptnsWndw.ChildObjects(objSelectType)
		objIntNoOfObjects(iCounter1).Click 5, 5, "LEFT"
	End If
		
'	If sScopes = "Site" Then  'Commented by Siddhi on 1-Nov-2012 
'            Call Fn_Edit_Box("Fn_SISW_Pref_Search_CreateOperation",oOptnsWndw,"Description",sDescription)
'	End If
	'Added by Siddhi on 1-Nov-2012 	
	If sDescription <>"" Then
            Call Fn_Edit_Box("Fn_SISW_Pref_Search_CreateOperation",oOptnsWndw,"Description",sDescription)
	End If
	'End
	If bMulValues <> "" Then
'		Call Fn_UI_Object_SetTOProperty("Fn_SISW_Pref_Search_CreateOperation",oOptnsWndw.JavaRadioButton("MultiValOption"),"attached text",bMulValues)
'		Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Pref_Search_CreateOperation",oOptnsWndw, "MultiValOption")
		If cBool(bMulValues) Then
			BlnMultiVal = "Multiple"
		Else
			BlnMultiVal = "Single"
		End If
                Call Fn_Button_Click("Fn_SISW_Pref_Search_CreateOperation", oOptnsWndw, "PrefMultipleDrpDwn")
		wait 2
		Set objSelectType=description.Create()
		objSelectType("Class Name").value = "JavaStaticText"
		objSelectType("label").value = BlnMultiVal
		Set objIntNoOfObjects = oOptnsWndw.ChildObjects(objSelectType)
		objIntNoOfObjects(0).Click 5, 5, "LEFT"
	End If

		'++++++++++<<    Set Category / Type  >>++++++++++
        If sCategory <> "" Then
		Call Fn_Button_Click("Fn_SISW_Pref_Search_CreateOperation", oOptnsWndw, "PrefCategoryDrpDwn")
		wait 2
		Set objSelectType = Description.Create()
		objSelectType("Class Name").value = "JavaStaticText"
		objSelectType("label").value = sCategory
		Set objIntNoOfObjects = oOptnsWndw.ChildObjects(objSelectType)
		objIntNoOfObjects(0).Click 5, 5, "LEFT"
	End If


	If sType <> "" Then
                Call Fn_Button_Click("Fn_SISW_Pref_Search_CreateOperation", oOptnsWndw, "NewPrefTypeDrpDwn")
		wait 2
		Set objSelectType=description.Create()
		objSelectType("Class Name").value = "JavaStaticText"
		objSelectType("label").value = StrType
		Set objIntNoOfObjects = oOptnsWndw.ChildObjects(objSelectType)
		objIntNoOfObjects(0).Click 5, 5, "LEFT"
	End If

		'++++++++++<<    Set  Values   >>++++++++++
	If sValues <> "" Then
		If instr(sValues,":") > 0 Then
			ArrValue = Split(sValues,":",-1)
			For iCounter1=0 to Ubound(ArrValue)
				Call Fn_Edit_Box("Fn_SISW_Pref_Search_CreateOperation",oOptnsWndw,"CurrentValues",ArrValue(iCounter1))
				wait 1
				Call Fn_Button_Click("Fn_SISW_Pref_Search_CreateOperation", oOptnsWndw, "NewMultiValPrefAddValue")
				wait 1								
			Next		
		Else
			Call Fn_Edit_Box("Fn_SISW_Pref_Search_CreateOperation",oOptnsWndw,"CurrentValues",sValues)
		End If
        End If 
		
	Call Fn_Button_Click( "Fn_SISW_Pref_Search_CreateOperation", oOptnsWndw, "Save" )

	Call Fn_Button_Click( "Fn_SISW_Pref_Search_CreateOperation", oOptnsWndw, "Close" )

	Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"PASS: Successfully Created New Search preference.")   	
	Fn_SISW_Pref_Search_CreateOperation = True

	Set oOptnsWndw  = Nothing
	Set objSelectType  = Nothing
	Set objIntNoOfObjects = Nothing
End Function 
'####### 		E.O.F~. =  Fn_SISW_Pref_Search_CreateOperation 		#########################################################################################################~
 '*********************************************************  Fn_SISW_Pref_Search_Lock_Unlock  ***********************************************************************
'Function Name  :    Fn_SISW_Pref_Search_Lock_Unlock
'Description    :      This Function is used for following :-
'                    
'             1. Create [---Done---]
'             2.Modify [---Done--}
'             3.Delete [---Done--}
'             4. Lock - Prerequisite : Index Page should be present in Edit->Options
'              Example : Fn_SISW_Pref_PreferenceOperations("Lock","","","","","","","","","","","")
'             6. Unlock - Prerequisite : Index Page should be present in Edit->Options
'              Example : Fn_SISW_Pref_PreferenceOperations("Unlock","","","","","","","","","","","")
'Parameters      :     
'             1.sAction,
'                                                   2.sSearchOnKeyWord,
'                                                   3.sCurrentValue)
'                                                   4.sDesc,
'             5.sScope,
'             6.sCategory
          
'Return Value     :    The String which represents the result : "PASS" or "FAIL" with the reason
'Pre-requisite   :    User should logged in to the teamcenter with DBA Privilledge
'Examples    :   Call Fn_SISW_Pref_Search_Lock_Unlock("Modify","ItemRevision.SUMMARYRENDERING","SampleItemRevStylesheet")
'            	 Call Fn_SISW_Pref_Search_Lock_Unlock("Delete","abc","")
'
'History     :   Developer Name     		Date   				Rev. No.    Changes Done      Reviewer        Reviewed Date
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Deepak Kumar/Mohit Khare						  1.0		    Created		    Mohit Khare
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'			Shreyas Waichal				14-June-2012		2.0		optimized and Modified functionaccording to TC10.0 UI Changes
'	 		Koustubh Watwe	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_Pref_Search_Lock_Unlock(sAction,sSearchOnKeyWord,sDesc,sScope,sCategory,sCurrentValue,BlnMultiVal )
 	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_Search_Lock_Unlock"
	Dim objDialog,objSelectType, objSelectType1, intNoOfObjects, intNoOfObjects1, iCounter1
	Dim objPrefOper,objSearch,WshShell,objdelete,objdeletedialog, bFound
	Dim oDesc, iCnt, iCounter, iSubCounter,objNewPrefOper, iCount
	Dim iCountChecked, objDefaultWindow

	Fn_SISW_Pref_Search_Lock_Unlock = False
' Dim  objSelectType, objSelectType1, intNoOfObjects, intNoOfObjects1, iCounter1,bFlag
'On Error Resume Next
	Set objNewPrefOper = Fn_SISW_Pref_GetObject("IndexOptions")
	If objNewPrefOper.Exist(5) =False Then
		Call Fn_MenuOperation("Select","Edit:Options...")
		Call Fn_ReadyStatusSync(3)
	End If
	Call Fn_ReadyStatusSync(1)
	Set objPrefOper = objNewPrefOper
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Create"  
                        Fn_SISW_Pref_Search_Lock_Unlock = Fn_SISW_Pref_PreferenceOperations("Create",sSearchOnKeyWord,sDesc,sScope,sCategory,sCurrentValue,"String",BlnMultiVal,"","","","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Modify"
			Fn_SISW_Pref_Search_Lock_Unlock = Fn_SISW_Pref_PreferenceOperations("Modify",sSearchOnKeyWord,sDesc,sScope,sCategory,sCurrentValue,"",BlnMultiVal,"","","","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Delete"
			Fn_SISW_Pref_Search_Lock_Unlock = Fn_SISW_Pref_PreferenceOperations("Delete",sSearchOnKeyWord,"","","","","","","","","","")
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Lock"

			Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"Index",0,0,"LEFT")
			If Fn_UI_ObjectExist("Fn_SISW_Pref_Search_Lock_Unlock", objPrefOper.JavaButton("LockPreference"))=True Then
				If Fn_Button_Click("Fn_SISW_Pref_Search_Lock_Unlock",objPrefOper,"LockPreference") = True Then
					Set objSelectType=Description.Create()
					Set objSelectType1=Description.Create()
					objSelectType("Class Name").value = "JavaDialog"
					objSelectType1("Class Name").value = "JavaButton"
                    Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")	
					Do
						bFlag = False
						Set  intNoOfObjects =  objTcDefaultApplet.ChildObjects(objSelectType)
						For iCounter = 0 to intNoOfObjects.count-1
							If intNoOfObjects(iCounter).getroproperty("tagname") = "Lock Site Preferences" Then
								bFlag = True
								Set  intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
								For iCounter1 = 0 to intNoOfObjects1.count-1
									If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
											 intNoOfObjects1(iCounter1).click
											 Fn_SISW_Pref_Search_Lock_Unlock = True
										Exit For
									End If
								Next
								Exit For
							End If
						Next
						Wait 3
					Loop While bFlag = True
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"PASS:Fn_SISW_Pref_Search_Lock_Unlock: " + "Preference locked successfully")		
				Else
					Fn_SISW_Pref_Search_Lock_Unlock = False
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Fn_SISW_Pref_Search_Lock_Unlock: " + "Failed to Lock the Preference")
				End If
			Else
				Fn_SISW_Pref_PreferenceOperations = False
				Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Fn_SISW_Pref_Search_Lock_Unlock:  Lock Button does not Exist")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "Unlock"
				Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper,"Index",0,0,"LEFT")
				If Fn_UI_ObjectExist("Fn_SISW_Pref_PreferenceOperations", objPrefOper.JavaButton("OpenLock"))=True Then
					call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objPrefOper,"OpenLock")
					Set objSelectType=Description.Create()
					Set objSelectType1=Description.Create()
					objSelectType("Class Name").value = "JavaDialog"
					objSelectType1("Class Name").value = "JavaButton"
                    Set objTcDefaultApplet = Fn_SISW_Pref_GetObject("TcDefaultApplet")	
					Do
						bFlag = False
						Set  intNoOfObjects =  objTcDefaultApplet.ChildObjects(objSelectType)
						For iCounter = 0 to intNoOfObjects.count-1
							If intNoOfObjects(iCounter).getroproperty("tagname") = "Unlock Site Preferences" Then
								bFlag = True
								Set  intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
								For iCounter1 = 0 to intNoOfObjects1.count-1
									If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
										intNoOfObjects1(iCounter1).click
										Fn_SISW_Pref_Search_Lock_Unlock = True
									Exit For
									End If
								Next
								Exit For
							End If
						Next
						Wait 3
					Loop While bFlag = True	
				Else
					Fn_SISW_Pref_Search_Lock_Unlock = False
					Call Fn_WriteLogFile(Environment.Value ("TestLogFile"),"FAIL: Fn_SISW_Pref_Search_Lock_Unlock : " + "Unlock button does not Exist")
				End If
	End Select
	 
	Set objSelectType=Description.Create()
	Set objSelectType1=Description.Create()
	objSelectType("Class Name").value = "JavaWindow"
	objSelectType1("Class Name").value = "JavaButton"
	Set objDefaultWindow =  Fn_SISW_Pref_GetObject("DefaultWindow")			
	Set  intNoOfObjects = objDefaultWindow.ChildObjects(objSelectType)
	For iCounter = 0 to intNoOfObjects.count-1
		If intNoOfObjects(iCounter).getroproperty("tagname") = "Unlock Site Preferences" Then
			Set  intNoOfObjects1 = intNoOfObjects(iCounter).ChildObjects(objSelectType1)
			For iCounter1 = 0 to intNoOfObjects1.count-1
				If intNoOfObjects1(iCounter1).getroproperty("attached text") = "OK" Then
					intNoOfObjects1(iCounter1).click
					Exit For
				End If
			Next
			Exit For
		End If
	Next

	Set objPrefOper = Nothing
	Set objSearch = Nothing
	Set  oDesc = Nothing
	Set  iCnt = Nothing
	Set  iCounter = Nothing
	Set  iSubCounter  = Nothing
	Set iCountChecked = Nothing
	Set objSelectType=Nothing
	Set objSelectType1 = Nothing
	Set intNoOfObjects = Nothing
	Set intNoOfObjects1 = Nothing
	Set objNewPrefOper = Nothing
	Set objDefaultWindow = Nothing
 End Function

 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@    Function Name		:	Fn_SISW_Pref_SearchPreferenceOperations
'@@
'@@    Description				 :	Function Used to perform operations on Search Preferences
'@@
'@@    Parameters			   :	1.StrAction: Action Name
'@@											  2.dicSearchPreferences: Search Prefernces value Dictionary
'@@
'@@    Return Value		   	   : 	True Or False
'@@
'@@    Pre-requisite			:	Simple Search Tab Should be open							
'@@
'@@    Examples					:	dicSearchPreferences("CaseSensitive")="ON"
'@@											 dicSearchPreferences("LatestDatasetVersion")="OFF"
'@@											 dicSearchPreferences("WildcardOption")="SQL Style"
'@@											 dicSearchPreferences("DelimitingCharacter")=","
'@@											 dicSearchPreferences("DefaultBOType")="ANDList"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("Search",dicSearchPreferences)
'@@											 dicSearchPreferences("FavoriteBOTypeAction")="Add"
'@@											 dicSearchPreferences("FavoriteBOType")="Folder"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("FavoriteBOType",dicSearchPreferences)
'@@											 dicSearchPreferences("FavoriteBOTypeAction")="Remove"
'@@											 dicSearchPreferences("FavoriteBOType")="Folder"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("FavoriteBOType",dicSearchPreferences)
'@@											 dicSearchPreferences("FavoriteBOTypeAction")="Up"
'@@											 dicSearchPreferences("FavoriteBOType")="Dataset"
'@@											 dicSearchPreferences("ShiftingCount")="2"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("FavoriteBOType",dicSearchPreferences)
'@@											 dicSearchPreferences("FavoriteBOTypeAction")="Down"
'@@											 dicSearchPreferences("FavoriteBOType")="Dataset"
'@@											 dicSearchPreferences("ShiftingCount")="2"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("FavoriteBOType",dicSearchPreferences)
'@@											 dicSearchPreferences("FavoriteBOTypeAction")="GetFavoriteBOTypeCount"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("FavoriteBOType",dicSearchPreferences)
'@@											 dicSearchPreferences("CaseSensitive")="Yes"
'@@											 dicSearchPreferences("LatestDatasetVersion")="Yes"
'@@											 dicSearchPreferences("WildcardOption")="SQL Style"
'@@											 dicSearchPreferences("DelimitingCharacter")="Yes"
'@@											 dicSearchPreferences("EscapeCharacter")="Yes"
'@@											 dicSearchPreferences("DefaultBOType")="Yes"
'@@											 dicSearchPreferences("SearchLocale")="Yes"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("GetSearchPrefCurrentValues",dicSearchPreferences)
'@@											 dicSearchPreferences("DefaultBOType")="Yes"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("GetSearchPrefCurrentValues",dicSearchPreferences)
'@@											 dicSearchPreferences("FavoriteBOTypeAction")="GetAllFavoriteBOTypes"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("FavoriteBOType",dicSearchPreferences)
'@@											 dicSearchPreferences("FavoriteBOTypeAction")="Remove"
'@@											 dicSearchPreferences("FavoriteBOType")="Folder~ANDList"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("FavoriteBOType",dicSearchPreferences)
'@@											 dicSearchPreferences("LoadingPageSize")="10"
'@@											 dicSearchPreferences("OpenSearchResultLimit")="20"
'@@											 dicSearchPreferences("LoadAllLimit")="500"
'@@											 Call Fn_SISW_Pref_SearchPreferenceOperations("Results",dicSearchPreferences)
'@@
'@@	   History					 	:	
'@@													Developer Name												Date						Rev. No.						Changes Done								  Reviewer
'@@------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'@@												Sandeep Navghane									08-Aug-2011						1.0																								Sunny Ruparel
'@@												Sandeep Navghane									09-Aug-2011						1.1								Added Case "FavoriteBOType"			 Sunny Ruparel
'@@												Sandeep Navghane									12-Aug-2011						1.2								Added Case "GetSearchPrefCurrentValues"			 Sunny Ruparel
'@@												Sandeep Navghane									12-Aug-2011						1.3								Added Case "GetAllFavoriteBOTypes"			 Sunny Ruparel
'@@												Sandeep Navghane									17-Aug-2011						1.4								Added Case "Results"			 				Sunny Ruparel
'@@												Sandeep Navghane									20-Dec-2012						1.5								Modified case : Search  according to 10.1 design changes
'																																																			[ Preferences ] dialog is remove and all operations put under [ Options ] dialog
'@@												Sanjeet K.											2-Jan-2013						1.5								Modified function according to Tc10.1
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Function Fn_SISW_Pref_SearchPreferenceOperations(StrAction,dicSearchPreferences)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_SearchPreferenceOperations"
    Dim objPrefDialog, objSelectBusiObjType, objOptions
    Dim iRowCount,iCounter,StrCurrFavBOType,bFlag,iShift,iCount,StrCurrVal,bReturn
	Dim StrCrrBOType,arrFavoriteBOType

	Set objPrefDialog=Fn_SISW_Pref_GetObject("IndexOptions")
	Set  objSelectBusiObjType  = Fn_SISW_Pref_GetObject("SelectBusinessObjectType")
	Set objOptions = Fn_SISW_Pref_GetObject("Options")

	Fn_SISW_Pref_SearchPreferenceOperations=False
	bFlag=False
	If Not objPrefDialog.Exist(6) Then
		wait(3)
		Call Fn_ToolbarButtonClick_Ext(2,"View Menu")
		wait(3)
'		JavaWindow("DefaultWindow").JavaMenu("label:=Preferences...","index:=0").Select
		JavaWindow("DefaultWindow").WinMenu("ContextMenu").Select "Options..."
		wait(3)
	End If
	Select Case StrAction
		Case "FavoriteBOType"
			objOptions.JavaTree("OptionsTree").Select "Options:Search:Favorite Business Object Types"
			wait 1
			Select Case dicSearchPreferences("FavoriteBOTypeAction")
				Case "Add"
					arrFavoriteBOType=Split(dicSearchPreferences("FavoriteBOType"),"~")
					For iCounter=0 To UBound(arrFavoriteBOType)
						Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "Add")
						wait 6
						Call Fn_Edit_Box("Fn_SISW_Pref_SearchPreferenceOperations",objSelectBusiObjType,"ChooseBOType",arrFavoriteBOType(iCounter))
						Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objSelectBusiObjType, "OK")
						wait 1
					Next
				Case "Remove"
                    iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaTable("PreferencesListTable"), "rows")
					arrFavoriteBOType=Split(dicSearchPreferences("FavoriteBOType"),"~")
					For iCount=0 To UBound(arrFavoriteBOType)
						For iCounter=0 To iRowCount-1
                            StrCurrFavBOType=objPrefDialog.JavaTable("PreferencesListTable").GetCellData(iCounter,"Business Object Type")
							If StrCurrFavBOType=arrFavoriteBOType(iCount) Then
								objPrefDialog.JavaTable("PreferencesListTable").SelectRow iCounter
                                  wait 1
								Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "Remove")
								bFlag=True
								wait 1
								Exit For
							End If
						Next
					Next
					If bFlag=False Then
						Set objPrefDialog= Nothing
						Set  objSelectBusiObjType  = Nothing
						Set objOptions = Nothing
						Exit Function
					End If
				Case "Up"
'					iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaTable("FavoritesTable"), "rows")
					iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaTable("PreferencesListTable"), "rows")
					For iCounter=0 To iRowCount-1
'						StrCurrFavBOType=objPrefDialog.JavaTable("FavoritesTable").GetCellData(iCounter,"Business Object Type")
						StrCurrFavBOType=objPrefDialog.JavaTable("PreferencesListTable").GetCellData(iCounter,"Business Object Type")
						If StrCurrFavBOType=dicSearchPreferences("FavoriteBOType") Then
							objPrefDialog.JavaTable("PreferencesListTable").SelectRow iCounter
'							objPrefDialog.JavaTable("FavoritesTable").SelectCell iCounter,"Business Object Type"
							wait 1
							If dicSearchPreferences("ShiftingCount")="" Then
								iShift=1
							Else
								iShift=dicSearchPreferences("ShiftingCount")
							End If
							For iCount=1 To CInt(iShift)
								Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "UpFavBOType")
							Next
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=False Then
						Set objPrefDialog= Nothing
						Set  objSelectBusiObjType  = Nothing
						Set objOptions = Nothing
						Exit Function
					End If

				Case "Down"
'					iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaTable("FavoritesTable"), "rows")
                    iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaTable("PreferencesListTable"), "rows")
					For iCounter=0 To iRowCount-1
'						StrCurrFavBOType=objPrefDialog.JavaTable("FavoritesTable").GetCellData(iCounter,"Business Object Type")
						StrCurrFavBOType=objPrefDialog.JavaTable("PreferencesListTable").GetCellData(iCounter,"Business Object Type")
						If StrCurrFavBOType=dicSearchPreferences("FavoriteBOType") Then
                            objPrefDialog.JavaTable("PreferencesListTable").SelectRow iCounter
'							objPrefDialog.JavaTable("FavoritesTable").SelectCell iCounter,"Business Object Type"
							wait 1
							If dicSearchPreferences("ShiftingCount")="" Then
								iShift=1
							Else
								iShift=dicSearchPreferences("ShiftingCount")
							End If
							For iCount=1 To CInt(iShift)
								Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "DownFavBOType")
							Next
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=False Then
						Set objPrefDialog= Nothing
						Set  objSelectBusiObjType  = Nothing
						Set objOptions = Nothing
						Exit Function
					End If
				Case "GetFavoriteBOTypeCount"
'					iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaTable("FavoritesTable"), "rows")
					iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaTable("PreferencesListTable"), "rows")
					Fn_SISW_Pref_SearchPreferenceOperations=iRowCount-1
					Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "OK")
					Set objPrefDialog= Nothing
					Set  objSelectBusiObjType  = Nothing
					Set objOptions = Nothing
					Exit Function
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			Case "GetAllFavoriteBOTypes"
						bReturn=""
						iRowCount=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaTable("PreferencesListTable"), "rows")
						For iCounter=0 To iRowCount-1
							StrCrrBOType=objPrefDialog.JavaTable("PreferencesListTable").GetCellData(iCounter,0)
							If bReturn<>"" Then
								bReturn=bReturn+":"+StrCrrBOType
							Else
								bReturn=StrCrrBOType
							End If
						Next
						If bReturn<>"" Then
							Fn_SISW_Pref_SearchPreferenceOperations=bReturn
						End If
						Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "Close")
						Set objPrefDialog= Nothing
						Set  objSelectBusiObjType  = Nothing
						Set objOptions = Nothing
						Exit Function
			End Select
			Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "OK")
			Fn_SISW_Pref_SearchPreferenceOperations=True

		Case "Search"
			objPrefDialog.JavaTree("OptionsTree").Select "Options:Search:General"
			wait 3
			If dicSearchPreferences("CaseSensitive")<>"" Then
				wait 1
				Call Fn_CheckBox_Set("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog,"CaseSensitive",dicSearchPreferences("CaseSensitive"))
			End If
			If dicSearchPreferences("LatestDatasetVersion")<>"" Then
				wait 1
				Call Fn_CheckBox_Set("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog,"LatestDatasetVersions",dicSearchPreferences("LatestDatasetVersion"))
			End If
			If dicSearchPreferences("SearchClassification")<>"" Then
				wait 1
				Call Fn_CheckBox_Set("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog,"SearchClassification",dicSearchPreferences("SearchClassification"))
			End If
			If dicSearchPreferences("EnableHierarchicalTypeSearch")<>"" Then
				wait 1
				Call Fn_CheckBox_Set("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog,"EnableHierarchicalType",dicSearchPreferences("EnableHierarchicalTypeSearch"))
			End If
			If dicSearchPreferences("WildcardOption")<>"" Then
				wait 1
				objPrefDialog.JavaRadioButton("WildcardOption").SetTOProperty "label",dicSearchPreferences("WildcardOption")
				wait 2
				Call Fn_UI_JavaRadioButton_SetON("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog, "WildcardOption")
			End If
			If dicSearchPreferences("DelimitingCharacter")<>"" Then
				wait 1
				Call Fn_Edit_Box("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog,"DelimitingChar",dicSearchPreferences("DelimitingCharacter"))
			End If
			If dicSearchPreferences("EscapeCharacter")<>"" Then
				wait 1
				Call Fn_Edit_Box("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog,"EscapeChar",dicSearchPreferences("EscapeCharacter"))
			End If
			If dicSearchPreferences("DefaultBOType")<>"" Then
				wait 1
				Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "ChangeDefaultBusinessObjectType")
				wait 6
				Call Fn_Edit_Box("Fn_SISW_Pref_SearchPreferenceOperations",objSelectBusiObjType,"ChooseBOType",dicSearchPreferences("DefaultBOType"))
				wait 1
				Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objSelectBusiObjType, "OK")
				wait 1
			End If
			If dicSearchPreferences("SearchLocale")<>"" Then
				wait 1
				Call Fn_List_Select("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "SearchLocale",dicSearchPreferences("SearchLocale"))
			End If
			wait 1
			Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "OK")
			Fn_SISW_Pref_SearchPreferenceOperations=True

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetSearchPrefCurrentValues"
			 bReturn=""
'			JavaWindow("DefaultWindow").JavaWindow("Preferences").JavaTree("Tree").Select "Teamcenter:Search"
			objPrefDialog.JavaTree("OptionsTree").Select "Options:Search"
			wait 1
			If dicSearchPreferences("CaseSensitive")<>"" Then
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaCheckBox("CaseSensitive"), "value")
				bReturn=StrCurrVal
			End If
			If dicSearchPreferences("LatestDatasetVersion")<>"" Then
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaCheckBox("LatestDatasetVersions"), "value")
				If bReturn<>"" Then
					bReturn=bReturn+":"+StrCurrVal
				Else
					bReturn=StrCurrVal
				End If
			End If
			If dicSearchPreferences("SearchClassification")<>"" Then
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaCheckBox("SearchClassification"), "value")
				If bReturn<>"" Then
					bReturn=bReturn+":"+StrCurrVal
				Else
					bReturn=StrCurrVal
				End If
			End If
			If dicSearchPreferences("EnableHierarchicalTypeSearch")<>"" Then
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaCheckBox("EnableHierarchicalType"), "value")
				If bReturn<>"" Then
					bReturn=bReturn+":"+StrCurrVal
				Else
					bReturn=StrCurrVal
				End If
			End If
			If dicSearchPreferences("WildcardOption")<>"" Then
				objPrefDialog.JavaRadioButton("WildcardOption").SetTOProperty "label",dicSearchPreferences("WildcardOption")
				wait 2
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaRadioButton("WildcardOption"),"value")
				If bReturn<>"" Then
					bReturn=bReturn+":"+StrCurrVal
				Else
					bReturn=StrCurrVal
				End If
			End If
			If dicSearchPreferences("DelimitingCharacter")<>"" Then		
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaEdit("DelimitingChar"),"value")
				If bReturn<>"" Then
					bReturn=bReturn+":"+StrCurrVal
				Else
					bReturn=StrCurrVal
				End If
			End If
			If dicSearchPreferences("EscapeCharacter")<>"" Then
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaEdit("EscapeChar"),"value")
				If bReturn<>"" Then
					bReturn=bReturn+":"+StrCurrVal
				Else
					bReturn=StrCurrVal
				End If
			End If
			If dicSearchPreferences("DefaultBOType")<>"" Then
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaEdit("DefaultBOType"),"value")
				If bReturn<>"" Then
					bReturn=bReturn+":"+StrCurrVal
				Else
					bReturn=StrCurrVal
				End If
			End If
			If dicSearchPreferences("SearchLocale")<>"" Then
				StrCurrVal=Fn_UI_Object_GetROProperty("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog.JavaList("SearchLocale"),"value")
				If bReturn<>"" Then
					bReturn=bReturn+":"+StrCurrVal
				Else
					bReturn=StrCurrVal
				End If
			End If
			If bReturn<>"" Then
				Fn_SISW_Pref_SearchPreferenceOperations=bReturn
			End If

			If objPrefDialog.JavaButton("Cancel").Exist(5)=True Then
				Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "Cancel")
			Else
				Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "Close")
			End If           
		' - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - -- - - - - - - - - - - -- - - - - -- - - - - - - - - - - -
		Case "Results"
			objPrefDialog.JavaTree("OptionsTree").Select "Options:Search:Results"
'			JavaWindow("DefaultWindow").JavaWindow("Preferences").JavaTree("Tree").Select "Teamcenter:Search:Results"
			wait 2
			 If dicSearchPreferences("LoadingPageSize")<>"" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog,"LoadingPageSize",dicSearchPreferences("LoadingPageSize"))
			 End If
			 If dicSearchPreferences("OpenSearchResultLimit")<>"" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog,"OpenSrchResultLimit",dicSearchPreferences("OpenSearchResultLimit"))
			 End If
			 If dicSearchPreferences("LoadAllLimit")<>"" Then
				Call Fn_Edit_Box("Fn_SISW_Pref_SearchPreferenceOperations",objPrefDialog,"LoadAllLimit",dicSearchPreferences("LoadAllLimit"))
			 End If
			Call Fn_Button_Click("Fn_SISW_Pref_SearchPreferenceOperations", objPrefDialog, "OK")
			Fn_SISW_Pref_SearchPreferenceOperations=True
	End Select

	Set objPrefDialog= Nothing
	Set  objSelectBusiObjType  = Nothing
	Set objOptions = Nothing

End Function


'*****************************************************************************************************************************************************************************
'''''   Function Name 			:			Fn_SISW_Pref_ImportFromOrganization()																																		'

'''''	Description					 :			To import the preference from the Organisation pane in the Options dialog											 								

'''''	Parameters				   :			sAction :                   Pass Empty string ""																																
'''''														 sUser :                    Node in the Organisation tree eg. "Organization:dba"													  
'''''                                                        sFileName :          Filename to be imported eg. "D:\mainline\TestData\My_Teamcenter\perfScope.xml"
'''''														 toLoc :                    Location to import eg. "Group:dba"   NOTE: pass "" when the organization  tree is selected
'''''                                                       strMode :                Mode of Import  eg. "Automatic"  
'''''														strActCnflt :            Action to be taken when conflicted  eg. "Merge the preference values in the XML file with the values in the database."
'''''                                                       bOpenLog :          Pass boolean value eg. TRUE
'''''														sInfo1 &  sInfo2 : For future use
'''''
'''''	Return Value			 : 				True\False
'''''
'''''	PreRequisits			 :			   TeamCenter Should be launched
'''''
'''''
'''''	History
'''''- - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - - - - - - - 
'											Developer                    Date							Revision								ChangesDone                             Reveiwer
'''''- - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - - - - - - - 
'										Pritam Shikare				13-July-2012					1.0									   Newly Developed					   Nilesh G.
'
'''''- - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - -  - - - - - - - - - - - - - - - -- - - - - - - - - - 
'******************************************************************************************************************************************************************************
Public Function Fn_SISW_Pref_ImportFromOrganization(sAction,sUser,sFileName, toLoc, strMode, strActCnflt, bOpenLog, sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_SISW_Pref_ImportFromOrganization"
	On Error Resume Next
	Fn_SISW_Pref_ImportFromOrganization = False

	Dim objPrefOper, objSelectType, objDialog, objDialog2
	Dim  aUserList, iCounter, sPath, bFlag

	'Set the object of the Options Dialog box
	Set objPrefOper = Fn_SISW_Pref_GetObject("IndexOptions")

	'Open the Options dialog by Edit>>Options Menu
	If Fn_UI_ObjectExist("Fn_SISW_Pref_ImportFromOrganization", objPrefOper)=False Then
		Call Fn_MenuOperation("Select","Edit:Options...")
		Call Fn_ReadyStatusSync(3)
	End If

	'Click on the Organisation link
	Call Fn_UI_JavaStaticText_Click("Fn_SISW_Pref_ImportFromOrganization", objPrefOper,"Organization",0,0,"LEFT")

	'Select the node from the Organization tree
	If Trim(sUser) <> "" Then
		aUserList = Split(sUser, ":")
		For iCounter = 0 To UBound(aUserList)
			If iCounter = 0  Then
				sPath = aUserList(iCounter)
			Else
				sPath = sPath + ":"+aUserList(iCounter)
			End If
			Call Fn_JavaTree_Select("Fn_SISW_Pref_ImportFromOrganization", objPrefOper, "Organization", sPath)
			wait 2
			Call Fn_UI_JavaTree_Expand("Fn_SISW_Pref_ImportFromOrganization",objPrefOper,"Organization",sPath)
			wait 2
		Next
		Call Fn_JavaTree_Select("Fn_SISW_Pref_ImportFromOrganization", objPrefOper, "Organization", sUser)
		wait 2
	End If

	'Click on the Import Link
	objPrefOper.JavaStaticText("BottomLink").SetTOProperty "label", "Import"
	objPrefOper.JavaStaticText("BottomLink").Click 10, 10,"LEFT"

	'	Enter the Filename with the full Path		
'	Call Fn_Edit_Box("Fn_SISW_Pref_ImportFromOrganization",objPrefOper,"ImportFileName",sFileName)
'	Call Fn_Button_Click("Fn_SISW_Pref_ImportFromOrganization", objPrefOper, "BrowseFile")
'
'		Set objDialog2 =Fn_SISW_Pref_GetObject("ImportPreferences")
'		IF objDialog2.Exist = True Then
'			Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objDialog2,"FileName",sFileName)
'			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objDialog2, "Import")
'		End If
	If Environment.Value("ProductName") = sUFTProductName Then
		Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations", objPrefOper, "BrowseFile")
		Set objDialog2 =Fn_SISW_Pref_GetObject("ImportPreferences")
	    IF objDialog2.Exist = True Then
		    Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objDialog2,"FileName",sFileName)
	        Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objDialog2, "Import")
	    End If
	Else
		Call Fn_Edit_Box("Fn_SISW_Pref_ImportFromOrganization",objPrefOper,"ImportFileName",sFileName)
		Call Fn_Button_Click("Fn_SISW_Pref_ImportFromOrganization", objPrefOper, "BrowseFile")
		Set objDialog2 =Fn_SISW_Pref_GetObject("ImportPreferences")
		IF objDialog2.Exist = True Then
			Call Fn_Edit_Box("Fn_SISW_Pref_PreferenceOperations",objDialog2,"FileName",sFileName)
			Call Fn_Button_Click("Fn_SISW_Pref_PreferenceOperations",objDialog2, "Import")
		End If
	End If
	'select the location to which the file is to be imported
	If toLoc <> "" Then
			Call Fn_Button_Click("Fn_SISW_Pref_ImportFromOrganization", objPrefOper, "ToLocationDrpDwn")
			wait 2
			' add code to select static text of scope
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			Set objDialog = objPrefOper.ChildObjects(objSelectType)
			bFlag = false
			For iCounter = 0 to objDialog.count-1
				If instr(lcase(objDialog(iCounter).GetROProperty("label")),lcase(toLoc)) > 0 Then
					wait 1
					On error resume next
					objDialog(iCounter).Click 1, 1
					bFlag = True
					Err.Number = 0
				End If
			Next
			If bFlag = False Then
				Fn_SISW_Pref_ImportFromOrganization = False
				Exit function
			End If
	End If


		'Select Import Mode
	If Trim(strMode) <> "" Then
		objPrefOper.JavaRadioButton("ImportMode").SetTOProperty "attached text", strMode & " Import"
		objPrefOper.JavaRadioButton("ImportMode").Set "ON"
	End If

	'Set Open the log after Import 
	If  bOpenLog <> "" Then
		If bOpenLog = True Then
			objPrefOper.JavaCheckbox("Open on Export").SetTOProperty "attached text","Open the log report after Import"
			objPrefOper.JavaCheckbox("Open on Export").Set "ON"
		Else
			objPrefOper.JavaCheckbox("Open on Export").Set "OFF"
		End If
	End If

	'Set the action to be taken in case of conflict
	If Trim(strMode) <> "" AND Trim(strMode) = "Automatic" Then
		If Trim(strActCnflt) <> "" Then
			objPrefOper.JavaRadioButton("OverridePrefMode").SetTOProperty "attached text", strActCnflt
			objPrefOper.JavaRadioButton("OverridePrefMode").Set "ON"
		End If
		bFlag = Fn_Button_Click("Fn_SISW_Pref_ImportFromOrganization", objPrefOper, "Import")
	Else
		bFlag = Fn_Button_Click("Fn_SISW_Pref_ImportFromOrganization", objPrefOper, "DryRun")
	End If

	'Hadle the information dialog	
	If objPrefOper.JavaDialog("Information").Exist(10) Then
		Call Fn_Button_Click("Fn_SISW_Pref_ImportFromOrganization", objPrefOper.JavaDialog("Information"), "OK")
	End If
	Call Fn_Button_Click("Fn_SISW_Pref_ImportFromOrganization", objPrefOper, "Close")

	'Return the Boolean result
	If bFlag=True Then
		Fn_SISW_Pref_ImportFromOrganization = True
	Else
		Fn_SISW_Pref_ImportFromOrganization = False
		Exit Function
	End If

End Function
