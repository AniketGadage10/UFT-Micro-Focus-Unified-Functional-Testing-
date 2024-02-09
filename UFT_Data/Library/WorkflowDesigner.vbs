'Function List
'Fn_SISW_WorkflowDesigner_GetObject()
'1. Fn_WrkflwDegnr_TemplateFilterApply()
'2. Fn_WrkflwDegnr_SetAvailStage()
'3. Fn_WrkflwDegnr_SetEditStage()
'4. Fn_WrkflwDegnr_HandlersOperations()
'5. Fn_WrkflwDegnr_ErrorMessageVerify()
'6. Fn_WrkflwDegnr_ProcessTemplateTree_Operations()
'7. Fn_WrkflwDegnr_Attributes()
'8. Fn_WrkflwDegnr_TempleteDetails()
'9. Fn_WrkflwDegnr_ApplyTemplateChanges()
'10. Fn_WrkflwDegnr_NamedACLCreate()
'11. Fn_WrkflwDegnr_NamedACLAssign()
'12. Fn_WrkflwDegnr_NewRootTemplateOperations()
'13.Fn_WrkflwDegnr_EditTemplateFilter_Operations
'14.Fn_WrkflwDegnr_ExportWorkflowTemplates()
'15.Fn_WrkflwDegnr_ImportWorkflowTemplates()

'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_WorkflowDesigner_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_WorkflowDesigner_GetObject("NewRootTemplate")

'History:
'	Developer Name			Date			Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Nilesh Gadekar		 26-June-2012		1.0				
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 25-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_WorkflowDesigner_GetObject(sObjectName)
	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\WorkflowDesigner.xml"
	Set Fn_SISW_WorkflowDesigner_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)
End Function


'*********************************************************  Function do Operation on WorkList Tree *********************************************************************

'Function Name		:				Fn_WrkflwDegnr_TemplateFilterApply

'Description		:		 		Sets the template filter

'Parameters			   :	 		1. sGroup: Group name to be selected
'									2. sType: Object type to be selected
'									3. aTemplate: aArray of templates to be selected

'Return Value		   : 			True/False

'Pre-requisite			:		 	user is logged in to Wrokflow Designer module

'Examples				:			Fn_WrkflwDegnr_TemplateFilterApply("dba", "XMLAuditLog1", Array("TCM Release Process","TimeSheetApproval","Process"))

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Mahendra Bhandarkar		07-Sep-2010	       1.0
'										Ashok kakade			08-June-2012	       1.0			Modified Hierarchy of ProcessTemplateFilter Dialog
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WrkflwDegnr_TemplateFilterApply(sGroup, sType, aTemplate)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_TemplateFilterApply"
	Dim objDialog, bReturn, iCounter, WshShell
    Dim sTemplateType,intNoOfObjects,i

	Set objDialog = Fn_SISW_WorkflowDesigner_GetObject("ProcessTemplateFilter")

	If objDialog.Exist(5) = False Then 
		bReturn = Fn_MenuOperation("Select", "Edit:Template Filter")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Edit --> Template Filter ") 		
			Fn_WrkflwDegnr_TemplateFilterApply = False
			Set objDialog = Nothing
			Exit Function					
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Edit --> Template Filter")
			Call Fn_ReadyStatusSync(2)
		End If
	End If

	If objDialog.Exist(5) Then

		If Trim(sGroup) <> "" Then
			objDialog.JavaList("GroupName").Select sGroup
			Wait(2)
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sGroup)+"]") 		
				Fn_WrkflwDegnr_TemplateFilterApply = False
				objDialog.JavaButton("Cancel").Click micLeftBtn
				Set objDialog = Nothing
				Exit Function	
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sGroup)+"]")
				Call Fn_ReadyStatusSync(2)
			End If
		End If

		If Trim(sType) <> "" Then
			'Added by Nilesh
			objDialog.JavaEdit("ObjectType").Set sType
			objDialog.Type micReturn

			'Modified by Omkar  for build  (20111219.00)  ... Date 28 Dec 2011
'			objDialog.JavaButton("TypeDropDown").Click micLeftBtn
'			Wait(2)
'			Set sTemplateType=Description.Create()
'			sTemplateType("Class Name").value = "JavaStaticText"
'	
'			Set  intNoOfObjects = objDialog.ChildObjects(sTemplateType)
'			  For i = 0 to intNoOfObjects.count-1
'				   If  intNoOfObjects(i).getROProperty("label") = sType Then
'							intNoOfObjects(i).Click 1,1							
'							Exit for
'				   End If
'			  Next

			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Object Type ["+CStr(sType)+"] ")
				Fn_WrkflwDegnr_TemplateFilterApply = False
				objDialog.JavaButton("Cancel").Click micLeftBtn
				Set objDialog = Nothing
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Object Type ["+CStr(sType)+"] ")
				Call Fn_ReadyStatusSync(2)
			End If
		End If

'		Set WshShell = CreateObject("WScript.Shell")
'		Wait(3)
'		WshShell.SendKeys "{ENTER}"
'		Wait(2)
'		Set WshShell = Nothing

		If IsArray(aTemplate) Then
			For iCounter = 0 To UBound(aTemplate)
				If Trim(aTemplate(iCounter)) <> "" Then
					If Fn_UI_ListItemExist("Fn_WrkflwDegnr_TemplateFilterApply", objDialog, "DefinedProcTemp",aTemplate(iCounter)) = True Then
						objDialog.JavaList("DefinedProcTemp").Select Trim(aTemplate(iCounter))
						Wait(2)
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Defined Process Template ["+CStr(aTemplate(iCounter))+"] ")
							Fn_WrkflwDegnr_TemplateFilterApply = False
							objDialog.JavaButton("Cancel").Click micLeftBtn
							Set objDialog = Nothing
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Defined Process Template ["+CStr(aTemplate(iCounter))+"] ")
							Call Fn_ReadyStatusSync(2)
						End If
						objDialog.JavaButton("Left").Click micLeftBtn
						If Err.Number < 0 Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Left Button ") 		
							Fn_WrkflwDegnr_TemplateFilterApply = False
							objDialog.JavaButton("Cancel").Click micLeftBtn
							Set objDialog = Nothing
							Exit Function						
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Left Button")
							Call Fn_ReadyStatusSync(2)
						End If
					Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : List Item ["+CStr(aTemplate(iCounter))+"] Does Not Exist in Defined Process Template")
							Fn_WrkflwDegnr_TemplateFilterApply = False
							objDialog.JavaButton("Cancel").Click micLeftBtn
							Set objDialog = Nothing
							Exit Function
					End If
				End If
			Next
		End If

		objDialog.JavaButton("OK").Click micLeftBtn

		If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on OK Button ") 		
			Fn_WrkflwDegnr_TemplateFilterApply = False
			objDialog.JavaButton("Cancel").Click micLeftBtn
			Set objDialog = Nothing
			Exit Function		
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on OK Button")
			Call Fn_ReadyStatusSync(2)
		End If

		Fn_WrkflwDegnr_TemplateFilterApply = True

	Else

		Fn_WrkflwDegnr_TemplateFilterApply = False
		Set objDialog = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WrkflwDegnr_TemplateFilterApply function failed")
	End If
	
	Set objDialog = Nothing

End Function


'*********************************************************  Function do Operation on Set Available Stage *********************************************************************

'Function Name		:				Fn_WrkflwDegnr_SetAvailStage

'Description		:		 		Sets workflow temaplate stage to available

'Parameters			   :	 		1. sButtonName: Button to be clicked

'Return Value		   : 			True/False

'Pre-requisite			:		 	User is logged into Workflow Designer module and template is set to Edit mode already

'Examples				:			Fn_WrkflwDegnr_SetAvailStage("Cancel")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Mahendra Bhandarkar		08-Sep-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WrkflwDegnr_SetAvailStage(sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_SetAvailStage"
	On Error Resume Next
	Dim objDialogHead, objDialog, objPopup

	Set objDialogHead = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame")
	Set objDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaEdit("ProcessTemplate")

	If objDialog.Exist(5) Then

		If objDialogHead.JavaCheckBox("SetStagetoAvailable").Exist(5) Then
				objDialogHead.JavaCheckBox("SetStagetoAvailable").Set "ON"

				If Err.Number < 0 Then
					Fn_WrkflwDegnr_SetAvailStage = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Set Checked to Set Stage to Available CheckBox." )	
					Set objDialogHead = Nothing
					Set objDialog = Nothing
					Exit Function 
				Else
					Set objDialogHead = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Checked to Set Stage to Available CheckBox.")	
				End If

			  Set objPopUp = JavaDialog("Offline?")
				  objPopUp.SetTOProperty "title", "Stage Change"

			  If objPopUp.Exist(5) Then
					If Trim(sButtonName) = "Yes" OR Trim(sButtonName) = "No" OR Trim(sButtonName) = "Cancel" Then
							objPopUp.Activate
						If objPopUp.JavaButton(sButtonName).Exist(5) Then
							objPopUp.JavaButton(sButtonName).Click micLeftBtn
							If Err.Number < 0 Then
								Fn_WrkflwDegnr_SetAvailStage = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on " + sButtonName + " Button." )	
								Set objPopUp = Nothing
								Set objDialog = Nothing
								Exit Function 
							Else
								Fn_WrkflwDegnr_SetAvailStage = True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on " + sButtonName + " Button.")	
							End If
						Else
							Fn_WrkflwDegnr_SetAvailStage = False
							Set objDialog = Nothing
							Set objPopUp = Nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Button "+CStr(sButtonName)+" Does Not Exist in Fn_WrkflwDegnr_SetAvailStage function ")
						End If
					Else
							Fn_WrkflwDegnr_SetAvailStage = False
							Set objDialog = Nothing
							Set objPopUp = Nothing
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Button "+CStr(sButtonName)+" Does Not Exist, It Should be (Yes/No/Cancel) in Fn_WrkflwDegnr_SetAvailStage function ")
					End If
			  Else
				Fn_WrkflwDegnr_SetAvailStage = False
				Set objDialog = Nothing
				Set objPopUp = Nothing
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Popup Stage Changed Dialog Box Does Not Exist in Fn_WrkflwDegnr_SetAvailStage function ")
			  End If
		Else
			Fn_WrkflwDegnr_SetAvailStage = False
			Set objDialog = Nothing
			Set objDialogHead = Nothing
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Set Stage to Available CheckBox Does Not Exist in Fn_WrkflwDegnr_SetAvailStage function ")
		End If
	Else
		Fn_WrkflwDegnr_SetAvailStage = False
		Set objDialog = Nothing
		Set objDialogHead = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Process Template Dialog Box Does Not Exist in Fn_WrkflwDegnr_SetAvailStage function ")
	End If

	Set objDialog = Nothing

End Function

'*********************************************************  Function do Operation on Template to Put it in Edit Mode *********************************************************************

'Function Name		:				Fn_WrkflwDegnr_SetEditStage

'Description		:		 		Put the specified template into Edit mode

'Parameters			   :	 		1. sTemplateName: Name of the template to be selected to Edit 
'									2. sButtonName:  Button to be clicked

'Return Value		   : 			True/False

'Pre-requisite			:		 	User is logged into Workflow Designer module

'Examples				:			Fn_WrkflwDegnr_SetEditStage("AutoAssignDoReview","Cancel")

'History:
'										Developer Name						Date							Rev. No.			Changes Done																			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Mahendra Bhandarkar		08-Sep-2010	      			 1.0
'										Shreyas									07-11-2011			  		  1.1					Modified Code to Select template name									Prasanna
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WrkflwDegnr_SetEditStage(sTemplateName, sButtonName)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_SetEditStage"
	 On Error Resume Next
	 Dim objDialog, objPopUp, bReturn,WshShell
	 Dim sTemName,i,sTemplateType,intNoOfObjects
	
	 Set objDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame")

	 If objDialog.Exist(5) Then
		objDialog.Object.setFocusable(True)
		If Trim(sTemplateName) <> "" Then
'			objDialog.JavaEdit("ProcessTemplate").Type(sTemplateName)

			'Set Process Template
			objDialog.JavaButton("ProcTempBtn").Click
			wait(3)
			Set sTemplateType=Description.Create()
			sTemplateType("Class Name").value = "JavaStaticText"
			sTemplateType("label").value = Trim(sTemplateName)
	
			Set  intNoOfObjects = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").ChildObjects(sTemplateType)
			  'For i = 0 to intNoOfObjects.count-1
				   'If  intNoOfObjects(i).getROProperty("label") = sTemplateName Then
				   If  intNoOfObjects.count > 0 Then
						intNoOfObjects(0).Click 1,1
						wait(1)
                        bFlag = True
						'Exit for
				   End If
			  'Next

			Wait(2)
			If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Template ["+CStr(sTemplateName)+"] ")
				Fn_WrkflwDegnr_SetEditStage = False
				Set objDialog = Nothing
				Exit Function
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Template ["+CStr(sTemplateName)+"] ")
				Call Fn_ReadyStatusSync(2)
			End If
		End If
'		objDialog.JavaEdit("ProcessTemplate").Activate
'		Set WshShell = CreateObject("WScript.Shell")
'		Wait(3)
'		WshShell.SendKeys "{ENTER}"
'		Set WshShell = Nothing

		'Verify Selected Tempalte against Expected
		sTemName = objDialog.JavaEdit("ProcessTemplate").GetROProperty("value")
		If trim(sTemName) <> trim(sTemplateName) Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Template ["+CStr(sTemplateName)+"] Not Exists")
			Fn_WrkflwDegnr_SetEditStage = False
			Set objDialog = Nothing
			Exit Function
		End If
	
		bReturn = Fn_ToolbatButtonClick("Edit Mode")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click the Edit Mode Toolbar")
			Fn_WrkflwDegnr_SetEditStage = False
			Exit Function						
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Edit Mode Toolbar ")
			Call Fn_ReadyStatusSync(2)
		End If
	
	  Set objPopUp = JavaDialog("Offline?")
		objPopUp.SetTOProperty "title", "Offline?"
		'Added by Nilesh on 19-Jun-12 for OR change on Build TC10_0606
		If objPopUp.Exist(5)=False Then
			Set objPopUp = JavaWindow("WorkflowDesignerWindow").JavaWindow("Offline?")
		End If

					  If objPopUp.Exist(5) Then
											If Trim(sButtonName) = "Yes" OR Trim(sButtonName) = "No" OR Trim(sButtonName) = "Cancel" Then
																	objPopUp.Activate
																	If objPopUp.JavaButton(sButtonName).Exist(5) Then
																						objPopUp.JavaButton(sButtonName).Click micLeftBtn
																						If Err.Number < 0 Then
																							Fn_WrkflwDegnr_SetEditStage = False
																							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Click on " + sButtonName + " Button." )	
																							Set objPopUp = Nothing
																							Exit Function 
																						Else
																							Fn_WrkflwDegnr_SetEditStage = True
																							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Clicked on " + sButtonName + " Button.")	
																						End If
																	Else
																						Fn_WrkflwDegnr_SetEditStage = False
																						Set objDialog = Nothing
																						Set objPopUp = Nothing
																						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Button "+CStr(sButtonName)+" Does Not Exist in Fn_WrkflwDegnr_SetEditStage function ")
																	End If
												Else
																	Fn_WrkflwDegnr_SetEditStage = False
																	Set objDialog = Nothing
																	Set objPopUp = Nothing
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Button "+CStr(sButtonName)+" Does Not Exist, It Should be (Yes/No/Cancel) in Fn_WrkflwDegnr_SetEditStage function ")
												End If
'					  Else 
'												Fn_WrkflwDegnr_SetEditStage = False
'												Set objDialog = Nothing
'												Set objPopUp = Nothing
'												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Popup Offline? Dialog Box Does Not Exist in Fn_WrkflwDegnr_SetEditStage function ")
					  End If

	 Else
					Fn_WrkflwDegnr_SetEditStage = False
					Set objDialog = Nothing
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Process Template Dialog Box Does Not Exist in Fn_WrkflwDegnr_SetEditStage function ")
	 End If

	Set objDialog = Nothing
	Set objPopUp = Nothing

End Function

'*********************************************************  Function do Operation on Handler Dialog ***************************************************************************************************************

'Function Name		:		Fn_WrkflwDegnr_HandlersOperations

'Description			:		 Function Use to perform operations on Handler Dialog eg. Create ,Delete Handler etc.

'Parameters			   :	 		1. strAction: Action Name eg-"Create" , "Exist" , "Delete"
'												 2. strNodeName: Handler Node Name
'												 3. strTaskAction: Task Action Name
'												 4. strActnHandler: Action Handler Name
'												 5. strArguments: Arguments Name (Tilda separated ~ )
'												 6. strValues: Arguments Values (Tilda separated ~ )
'												 7. errDialogName: Error Dialog Name (This Part is not yet Handle in function Error Comes if Action Handler is Not selected and try to create Handler then it comes)
'												 8. errMsg

'Return Value		   : 			True/False

'Pre-requisite			:		 	Should be present on Wrokflow Designer perspective

'Examples				:			Fn_WrkflwDegnr_HandlersOperations("Create","","Complete","EPM-set-property","props~values~from_att_type~to_att_type","CMSpecialInstruction~PROP::object_desc~REFERENCE~TARGET","","")
'												Fn_WrkflwDegnr_HandlersOperations("Exist","Test1:Complete:EPM-set-property","","","","","","")
'												Fn_WrkflwDegnr_HandlersOperations("Delete","Test1:Complete:EPM-set-property","","","","","","")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N				05-Oct-2010	       1.0																  Sunny R
'										Mahendra B				22-Oct-2010	      1.1																  Prasanna B
'										Shreyas						07-11-2011		  1.2																	Prasanna
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwDegnr_HandlersOperations(strAction,strNodeName,strTaskAction,strActnHandler,strArguments,strValues,errDialogName,errMsg)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_HandlersOperations"
   'Variable Declaration
   Dim ObjHandlerDialog,bReturn,arrArg,arrValues,iCount,strHandlerName,iItemCnt, arrAction
   Dim i,intNoOfObjects,sTemplateType
'Added by Nilesh on 20th June-2012
'creating object of Handlers
	Set ObjHandlerDialog=JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers")
	'invoking Handler Dialog
	Call Fn_CheckBox_Set("Fn_WrkflwDegnr_HandlersOperations", JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame"), "TskHandlerBtn", "ON")
	wait(3)
	'Added by Nilesh on 20th June 2012
	If  ObjHandlerDialog.Exist(5)=False Then
		JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaStaticText("HandlersText").DblClick 0,0
	End If
	
	Fn_WrkflwDegnr_HandlersOperations=False

	arrAction = Split(strAction, ":", -1, 1)
	If IsArray(arrAction) = True and UBound(arrAction) = 1 Then
		strAction = arrAction(0)
	End If

	Select Case strAction

		Case "Create"  'Case to create new handler and code added for clicking on handler type
			If IsArray(arrAction) = True and UBound(arrAction) = 1 Then
				If Trim(arrAction(1)) = "Rule" Then
					bReturn =  Fn_CheckBox_Set("", ObjHandlerDialog, "RuleHandler", "ON")
					If  bReturn=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : To Click on ["+arrAction(1)+"] CheckBox.")
						Set ObjHandlerDialog=Nothing
						Exit Function
					End If
				ElseIf  Trim(arrAction(1)) = "Action" Then
					bReturn =  Fn_CheckBox_Set("", ObjHandlerDialog, "ActionHandler", "ON")
					If  bReturn=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : To Click on ["+arrAction(1)+"] CheckBox.")
						Set ObjHandlerDialog=Nothing
						Exit Function
					End If
				End If
			End If
			If strTaskAction<>"" Then
				'Selecting Task Action
				bReturn= Fn_List_Select("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"TaskActionList",strTaskAction)
				If  bReturn=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : ["+strTaskAction+"] Task is not present in Task Action List")
					Set ObjHandlerDialog=Nothing
					Exit Function
				End If
			End if




			If strActnHandler<>"" Then							'
							'Selecting Action Handler
							If Ubound(arrAction) < 1 Then
	'										bReturn= Fn_List_Select("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"ActionHandlerList",strActnHandler)
					'Set handler
								JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers").JavaButton("ProcTempBtn").Click
						
								Set sTemplateType=Description.Create()
								sTemplateType("Class Name").value = "JavaStaticText"
						
								Set  intNoOfObjects = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers").ChildObjects(sTemplateType)
								  For i = 0 to intNoOfObjects.count-1
									   If  intNoOfObjects(i).getROProperty("label") = strActnHandler Then
												intNoOfObjects(i).Click 1,1
												bFlag = True
												Exit for
									   End If
								  Next


							Else
										If Trim(arrAction(1)) = "Action"  Then
'												bReturn= Fn_List_Select("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"ActionHandlerList",strActnHandler)
					'Set handler
								JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers").JavaButton("ProcTempBtn").Click
						
								Set sTemplateType=Description.Create()
								sTemplateType("Class Name").value = "JavaStaticText"
						
								Set  intNoOfObjects = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers").ChildObjects(sTemplateType)
								  For i = 0 to intNoOfObjects.count-1
									   If  intNoOfObjects(i).getROProperty("label") = strActnHandler Then
												intNoOfObjects(i).Click 1,1
												bFlag = True
												Exit for
									   End If
								  Next

										else
'												bReturn= Fn_List_Select("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"RuleHandlerList",strActnHandler)
					'Set handler
								JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers").JavaButton("ProcTempBtn").Click
						
								Set sTemplateType=Description.Create()
								sTemplateType("Class Name").value = "JavaStaticText"
						
								Set  intNoOfObjects = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers").ChildObjects(sTemplateType)
								  For i = 0 to intNoOfObjects.count-1
									   If  intNoOfObjects(i).getROProperty("label") = strActnHandler Then
												intNoOfObjects(i).Click 1,1
												bFlag = True
												Exit for
									   End If
								  Next


										End If		
																		
							 End if 
							If  bReturn=False Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : ["+strActnHandler+"] Handler is not present in ["+arrAction(1)+"] Handler List")
									Set ObjHandlerDialog=Nothing
									Exit Function
							End If
			End If
			arrValues=Split(strValues,"~")
			arrArg=Split(strArguments,"~")
			For iCount=0 To Ubound(arrArg)-1
				'Adding rows in ActionHandlerTable
				Call Fn_Button_Click("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"Add")
			Next
			For iCount=0 To Ubound(arrArg)
				'Setting Arguments in ActionHandlerTable
				Call Fn_UI_JavaTable_SetCellData("Fn_WrkflwDegnr_HandlersOperations",ObjHandlerDialog,"ActionHandlerTable",iCount,0,arrArg(iCount))
			Next
			For iCount=0 To Ubound(arrValues)
				'Setting Argument Values in ActionHandlerTable
				Call Fn_UI_JavaTable_SetCellData("Fn_WrkflwDegnr_HandlersOperations",ObjHandlerDialog,"ActionHandlerTable",iCount,1,arrValues(iCount))
			Next
			'Clicking on create button
			Call Fn_Button_Click("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"Create")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Create ["+strActnHandler+"] Handler")
			Fn_WrkflwDegnr_HandlersOperations=True

		Case "Exist" 'Case to check Existance of Handler in Handler Tree
			'Taking Item Counts from Handler Tree
			iItemCnt=Fn_UI_Object_GetROProperty("Fn_WrkflwDegnr_HandlersOperations",ObjHandlerDialog.JavaTree("TaskHandlerTree"), "items count")
			For iCount=0 To iItemCnt-1
				'Taking Item Name from Tree
				strHandlerName=ObjHandlerDialog.JavaTree("TaskHandlerTree").GetItem(iCount)
				If strHandlerName=strNodeName Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : ["+strNodeName+"] Handler is present in Handler Tree")
					Fn_WrkflwDegnr_HandlersOperations=True
					Exit For
				End If
			Next

		Case "Delete" 'Case to Delete Handler
			'Selecting Handler from Handler Tree
            Call Fn_JavaTree_Select("Fn_WrkflwDegnr_HandlersOperations",ObjHandlerDialog, "TaskHandlerTree",strNodeName)
			'Clicking on Delete button to delete Handler
			Call Fn_Button_Click("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"Delete")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully select ["+strNodeName+"] Node in Handler Tree")
			Fn_WrkflwDegnr_HandlersOperations=True

		Case "Modify"  'Case to modify handler
			If Trim(strNodeName) <> "" Then
				bReturn = Fn_JavaTree_Select("Fn_WrkflwDegnr_HandlersOperations",ObjHandlerDialog, "TaskHandlerTree",strNodeName)
				If  bReturn=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : ["+strTaskAction+"] Task is not present in Task Action List")
					Set ObjHandlerDialog=Nothing
					Exit Function
				End If
			End If
			If strTaskAction <> "" Then
				'Checking existance of Task action
				bReturn=Fn_UI_ListItemExist("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"TaskActionList",strTaskAction)
				If  bReturn=False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : ["+strTaskAction+"] Task is not present in Task Action List")
					Set ObjHandlerDialog=Nothing
					Exit Function
				End If
				'Selecting Task Action
				Call Fn_List_Select("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"TaskActionList",strTaskAction)
			End If
			If strActnHandler<>"" Then
				'Checking existance of Action Handler
'				bReturn=Fn_UI_ListItemExist("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"ActionHandlerList",strActnHandler)
'				If  bReturn=False Then
'					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : ["+strActnHandler+"] Handler is not present in Action Handler List")
'					Set ObjHandlerDialog=Nothing
'					Exit Function
'				End If

				'Selecting Action Handler
				'Call Fn_List_Select("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"ActionHandlerList",strActnHandler)
				Call Fn_Edit_Box("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"ActionHandlerEdit",strActnHandler)
				JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers").JavaEdit("ActionHandlerEdit").Activate
			End If
			arrValues=Split(strValues,"~")
			arrArg=Split(strArguments,"~")
			If ObjHandlerDialog.GetROProperty("rows") = 0 Then
					Call Fn_Button_Click("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"Add")			
			End If
			For iCount=0 To Ubound(arrArg)-1
				'Adding rows in ActionHandlerTable
				Call Fn_Button_Click("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"Add")
			Next
			For iCount=0 To Ubound(arrArg)
				'Setting Arguments in ActionHandlerTable
				Call Fn_UI_JavaTable_SetCellData("Fn_WrkflwDegnr_HandlersOperations",ObjHandlerDialog,"ActionHandlerTable",iCount,0,arrArg(iCount))
			Next
			For iCount=0 To Ubound(arrValues)
				'Setting Argument Values in ActionHandlerTable
				Call Fn_UI_JavaTable_SetCellData("Fn_WrkflwDegnr_HandlersOperations",ObjHandlerDialog,"ActionHandlerTable",iCount,1,arrValues(iCount))
			Next
			'Clicking on create button
			Call Fn_Button_Click("Fn_WrkflwDegnr_HandlersOperations", ObjHandlerDialog,"Modify")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Modified ["+strActnHandler+"] Handler")
			Fn_WrkflwDegnr_HandlersOperations=True

	End Select
	wait(3)
	JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Handlers").Close
	Set ObjHandlerDialog=Nothing
End Function

'-----------------------------------------------------------------------Function Use to Handle Error Dialog which Appears in WorkFlowDesigner Perspective-------------------------------------------------------------------------

'Function Name		:		Fn_WrkflwDegnr_ErrorMessageVerify

'Description			:		 Function Use to Handle Error Dialog which Appears in WorkFlowDesigner Perspective

'Parameters			   :	 		1. strDialogName: Error dialog Name
'												 2. strErrorMsg: Error Message
'												 3. strButton: Button Name

'Return Value		   : 			True/False

'Pre-requisite			:		 	Should be present on Wrokflow Designer perspective

'Examples				:			Fn_WrkflwDegnr_ErrorMessageVerify("Set to Available Stage Template Dialog","","Close")

'History:
'										Developer Name			Date			Rev. No.			Changes Done			Reviewer	
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N				07-Oct-2010	       1.0																  Sunny R
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwDegnr_ErrorMessageVerify(strDialogName,strErrorMsg,strButton)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_ErrorMessageVerify"
	GBL_EXPECTED_MESSAGE=strErrorMsg
   'Varaible Declaration
	Dim strMsg
	'Function Initially returns False
	Fn_WrkflwDegnr_ErrorMessageVerify=False
	'Setting text Property of ErrorDialog
	 JavaWindow("DefaultWindow").Dialog("ErrorDialog").SetTOProperty "text",strDialogName
	 'Checking Existance of ErrorDialog
	 If Fn_UI_ObjectExist("Fn_WrkflwDegnr_ErrorMessageVerify",  JavaWindow("DefaultWindow").Dialog("ErrorDialog"))=True Then
		 'Taking Error message currenty present on Error Dialog
		strMsg=JavaWindow("DefaultWindow").Dialog("ErrorDialog").Static("ErrText").GetROProperty("text")
		wait(2)
		'Checking Error Message is come which is expected
		If strMsg=strErrorMsg Then
			'Setting text Property of OK button
			JavaWindow("DefaultWindow").Dialog("ErrorDialog").WinButton("OK").SetTOProperty "text",strButton
			wait(2)
			'Clicking button which is pass by user to close Error dialog
			JavaWindow("DefaultWindow").Dialog("ErrorDialog").WinButton("OK").Click
			'Function returns True
			Fn_WrkflwDegnr_ErrorMessageVerify=True
		Else
			GBL_ACTUAL_MESSAGE=strMsg
		End If
	 End If
End Function

'*********************************************************  Function do Operation on Workflow Designer Process Template Tree *********************************************************************
'Function Name		:					Fn_WrkflwDegnr_ProcessTemplateTree_Operations
'
'Description			 :		 		    Action  performed :-
'														1. Node Select																	
'														2. Node Expand
'														3. Node Collapse																	
'														4.Exist
'
'Parameters			   :	 			1. sAction: Action to be performed
'													2. sNodeName: Fully qulified tree Path (delimiter as ':') [multiple node are separated by "," ] 
' 												 	3. sMenu : ContextMenu should be selected
'  
'Return Value		   : 			 True/False
'
'Pre-requisite			:		 	 Process Template tree should be displayed.
'
'Examples				:			 Fn_WrkflwDegnr_ProcessTemplateTree_Operations("Select","AutoRevRev1:New Review Task 1:select-signoff-team", "")
'												Fn_WrkflwDegnr_ProcessTemplateTree_Operations("PopupMenuSelect","Test Object Properties_1:Define Properties", "View Properties	Alt+Enter")
'History:
'												Developer Name				Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										     Mahendra Bhandarkar	22-Aug-2010	       1.0
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WrkflwDegnr_ProcessTemplateTree_Operations(sAction,sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_ProcessTemplateTree_Operations"
   On Error Resume Next

   Dim arrNodeList,iItemCount,iCounter,sTreeItem,arrNode,iOuterCount,aMenuList
   Dim objProcessTree, objContext, intCount, StrMenu
   Dim iNodeCounter, sExpnadNode

   Set objProcessTree =  JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaTree("ProcessTemplateTree")

	If InStr(1, sAction,":", 1) > 0 Then
		arrNode = Split(sAction, ":", -1, 1)
		sAction = arrNode(0)
		StrMenu = arrNode(1)
	End If

   If objProcessTree.Exist(5) Then
	        Select Case sAction
						Case "Select"                   		                            
										objProcessTree.Select sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwDegnr_ProcessTemplateTree_Operations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to select Node  [" + sNodeName + "] of Process Template Tree." )	
												Set objProcessTree = Nothing
												Exit Function 
										Else
												Fn_WrkflwDegnr_ProcessTemplateTree_Operations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully selected  Node [" + sNodeName + "] of Process Template Tree.")	
										End If  

						Case  "Expand"									
									objProcessTree.Expand sNodeName
										If Err.Number < 0 Then
												Fn_WrkflwDegnr_ProcessTemplateTree_Operations = False
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to expand node [ " + sNodeName + "] of Process Template Tree." )	
												Set objProcessTree = Nothing
												Exit Function 
										Else
												Fn_WrkflwDegnr_ProcessTemplateTree_Operations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully expanded node [" + sNodeName  + "] of Process Template Tree.")	
										End If

						Case "ExpandSelect"

										arrNode1 = Split(sNodeName, ":", -1, 1)
										For iNodeCounter = 0 To UBound(arrNode1)
											  If iNodeCounter = 0 Then
												   sExpnadNode = arrNode1(iNodeCounter)
											  Else
												   sExpnadNode = sExpnadNode+":"+arrNode1(iNodeCounter)
											  End If
											If iNodeCounter <>  UBound(arrNode1) Then
												objProcessTree.Expand sExpnadNode
												If Err.Number < 0 Then
														Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Node ["+sExpnadNode+"] in Process Tree.")
														Fn_WrkflwDegnr_ProcessTemplateTree_Operations = False
														Set objProcessTree = Nothing
														Exit Function
												Else
													  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expand Node ["+sExpnadNode+"] in Process Tree.")
													  Call Fn_ReadyStatusSync(2)
													   Wait(1)
												End If
											End If
										Next
			
										'Select the Node    							
										objProcessTree.Select sExpnadNode
										If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node ["+sExpnadNode+"] in Process Tree.")
												Fn_WrkflwDegnr_ProcessTemplateTree_Operations = False
												Set objProcessTree = Nothing
												Exit Function
										Else
											  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Node ["+sExpnadNode+"] in Process Tree.")
											  Call Fn_ReadyStatusSync(2)
											   Wait(1)
											   Fn_WrkflwDegnr_ProcessTemplateTree_Operations = True
												Set objProcessTree = Nothing
										End If
						
						Case  "Collapse"										
										objProcessTree.Collapse sNodeName
										If Err.Number < 0 Then
											Fn_WrkflwDegnr_ProcessTemplateTree_Operations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Collapse node [" + sNodeName + "] of Process Template Tree." )	
											Set objProcessTree = Nothing
											Exit Function 
										Else
											Fn_WrkflwDegnr_ProcessTemplateTree_Operations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Collapsed node [" + sNodeName  + "] of Process Template Tree.")	
										End If

						Case "Exist"
			                            iItemCount = objProcessTree.GetROProperty( "items count")										
										For iCounter=0 To (iItemCount-1)
											sTreeItem = objProcessTree.GetItem(iCounter)
											If Trim (LCase(sTreeItem)) = Trim(LCase(sNodeName)) Then
												Fn_WrkflwDegnr_ProcessTemplateTree_Operations = True
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully found node [" + sNodeName + "] of Process Template Tree." )	
												Exit For
											End If
										Next 	
							
										If  Cint(iCounter) = Cint (iItemCount) Then
											Fn_WrkflwDegnr_ProcessTemplateTree_Operations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to  Find node [" + sNodeName + "] of Process Template Tree." )	
											Set objProcessTree = Nothing
											Exit Function 
										End If

					Case "PopupMenuSelect"
										aMenuList = split(sMenu, ":",-1,1)
										iCounter = Ubound(aMenuList)
										objProcessTree.Select sNodeName
										wait 2
										objProcessTree.OpenContextMenu sNodeName
										wait 2
										'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
										' Changed Hierarchy  of ContextMenu . Changed By : Harshal Tanpure , Date : 07-April-2011

										Select Case iCounter
											Case "0"
													 sMenu = JavaWindow("WorkflowDesignerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0))
											Case "1"
													sMenu = JavaWindow("WorkflowDesignerWindow").WinMenu("ContextMenu").BuildMenuPath(aMenuList(0),aMenuList(1))
											Case Else
													Fn_WrkflwDegnr_ProcessTemplateTree_Operations = FALSE
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Wrong Parameter for Popup menu Select [" + sMenu + "]")	
												   Exit Function
										End Select
										Wait 5
										JavaWindow("WorkflowDesignerWindow").WinMenu("ContextMenu").Select sMenu
										If Err.Number < 0 Then
											Fn_WrkflwDegnr_ProcessTemplateTree_Operations = False
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to select menu [" + sMenu + "] of MyworkList Tree." )	
											Set objProcessTree = Nothing
											Exit Function 
										Else
											Fn_WrkflwDegnr_ProcessTemplateTree_Operations = True
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully selected menu [" + sMenu  + "] of MyworkList Tree.")	
										End If

										''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						Case Else
										Fn_WrkflwDegnr_ProcessTemplateTree_Operations = False
			End Select
   End if

Set objContext = Nothing
Set objProcessTree = Nothing

End Function


'*********************************************************  Function perform the worklist process view attributes operation *********************************************************************
'Function Name  :   Fn_WrkflwDegnr_Attributes
'
'Description    :        Workflow Viewer Attributes Operation
' 
'Parameters      :     sAction: 							Add/Remove/Modify/Verify
'           				 dicProcessViewAttributes: 	 Refer DictionaryDeclaration.vbs for the defination & keys included
' 
'Return Value     :   True/False
'
'Examples    :      
'								dicProcessViewAttributes.RemoveAll								
'								dicProcessViewAttributes.Add("ProcessTree") = "AutoRevFailPath:New Review Task 1:select-signoff-team"	---- Mandatory Field							
'								dicProcessViewAttributes.Add("State") = "Started"
'								dicProcessViewAttributes.Add("ResParty") = "AutoTestDBA (autotestdba)"
'								
'           					Call Fn_WrkflwDegnr_Attributes("Verify", dicProcessViewAttributes)
' 
'History:          Developer Name   Date    						Rev. No.   Changes Done   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           			Prasanna    		25-Oct-2010   1.0               
'           			Mahendra    		28-Oct-2010   1.0               
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Public Function Fn_WrkflwDegnr_Attributes(sAction, dicProcessViewAttributes)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_Attributes"
	 On Error Resume Next
	 Dim objAttributeDialog
	 Dim dicCount , dicKeys , dicItems
	 Dim iCounter, bReturn
	 Dim arrNodeList, iNodeCounter, arrNode, sExpnadNode
	 Dim iCounter1, sActionSelect
	 Dim objSelectType, intNoOfObjects, objDialog, iCounter2, sListHierarchy, arrQuorum
	 Dim arrValues,jCounter,iElemCount,kCounter,strValue

	 dicCount  = dicProcessViewAttributes.Count
	 dicItems = dicProcessViewAttributes.Items
	 dicKeys = dicProcessViewAttributes.Keys

	Set objDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame")
	Set objAttributeDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes")

				If dicProcessViewAttributes.Exists("ProcessTree") Then
							If dicProcessViewAttributes.Exists("ProcessTree") = ""  Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Process Tree Node Can't be Blank.")
										Fn_WrkflwDegnr_Attributes = False
										Set objDialog = Nothing
										Set objAttributeDialog = Nothing
										Exit Function
							 End If
							arrNode = Split(dicProcessViewAttributes.Item("ProcessTree"), ":", -1, 1)
							For iNodeCounter = 0 To UBound(arrNode)
							  If iNodeCounter = 0 Then
								   sExpnadNode = arrNode(iNodeCounter)
							  Else
								   sExpnadNode = sExpnadNode+":"+arrNode(iNodeCounter)
							  End If
							If iNodeCounter <>  UBound(arrNode) Then
								objDialog.JavaTree("ProcessTemplateTree").Expand sExpnadNode
								If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Expand Node ["+sExpnadNode+"] in Process Tree.")
										Fn_WrkflwDegnr_Attributes = False
										Set objDialog = Nothing
										Set objAttributeDialog = Nothing
										Exit Function
								Else
									  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Expand Node ["+sExpnadNode+"] in Process Tree.")
									  Call Fn_ReadyStatusSync(1)
									   Wait(1)
								End If
							End If
							Next

							'Select the Node    							
							bReturn = Fn_WrkflwDegnr_ProcessTemplateTree_Operations("Select",sExpnadNode,"")
							If bReturn = false Then
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Node ["+sExpnadNode+"] in Process Tree.")
										Fn_WrkflwDegnr_Attributes = False
										Set objDialog = Nothing
										Set objAttributeDialog = Nothing
										Exit Function
							Else                      
								  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Node ["+sExpnadNode+"] in Process Tree.")
								  Call Fn_ReadyStatusSync(1)
								   Wait(1)
							End If
		 			
							' Select the Attributes Dialog
							If  objAttributeDialog.Exist(2) = False  Then
									objDialog.JavaCheckBox("TskAttributeBtn").Set "ON"
									 If Err.Number < 0 Then
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Checkbox.")
										Fn_WrkflwDegnr_Attributes = False
										Set objDialog = Nothing
										Set objAttributeDialog = Nothing
										Exit Function
									 Else                      
										   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Checkbox.")
										   Call Fn_ReadyStatusSync(2)
											Wait(2)
									 End If

									'Click on Attributes text    Added by Nilesh for Build change on TC10_0606
									If objAttributeDialog.Exist(5)=False Then
											objDialog.JavaStaticText("Attributes").DblClick 1, 1, "LEFT"
											If Err.Number < 0 Then
													  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Text.")
													Fn_WrkflwDegnr_Attributes = False
													Set objDialog = Nothing
													Set objAttributeDialog = Nothing
													Exit Function
											Else                      
													  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Text.")
													  Wait(2)					
											End If
									End If
									
							End If
				End If

	If objAttributeDialog.Exist(2) = True Then

				'Activate the Attributes Dialog
				  objAttributeDialog.Activate
				  If Err.Number < 0 Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Attribute Dialog Does not exist.")
						Fn_WrkflwDegnr_Attributes = False
						Set objDialog = Nothing
						Set objAttributeDialog = Nothing
						Exit Function
				  Else
						wait(3)	
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully verified the existence of Attributes Dialog.")
				  End If


				   Select Case sAction
				   Case "Verify"
						For iCounter = 0 to dicCount - 1
							  If  dicItems(iCounter) <> "" Then
								 Select Case dicKeys(iCounter)
								 Case "State"
										   Set objComp =  objAttributeDialog.JavaObject("State").Object.getComponent(0)
										   If objComp.getText() = dicItems(iCounter) Then
													 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : State Selection ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
										   Else
													 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the State Selected Text ["+dicItems(iCounter)+"] in Attributes Dialog. ")
													Fn_WrkflwDegnr_Attributes = False
													Set objDialog = Nothing
													Set objAttributeDialog = Nothing
													Exit Function
										   End If
										   Set objComp = Nothing
						
								 Case "ResParty"
						
										   Set objComp =  objAttributeDialog.JavaObject("ResponsibleParty").Object.getComponent(0)
										   If objComp.getText() = dicItems(iCounter) Then
													 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Responsible Party Selection ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
										   Else
													 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the Responsible Party Selected Text ["+dicItems(iCounter)+"] in Attributes Dialog. ")
													Fn_WrkflwDegnr_Attributes = False
													Set objDialog = Nothing
													Set objAttributeDialog = Nothing
													Exit Function
										   End If
										   Set objComp = Nothing
						
								 Case "NameACL"
				
								 Case "SignOffsQuorum"
				
								 Case "DueDate"
				
								 Case "Duration"
				
										   Set objComp =  objAttributeDialog.JavaEdit("Duration")
										   If objComp.GetROProperty("text") = dicItems(iCounter) Then
													 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Duration Text ["+dicItems(iCounter)+"] verified in Attributes Dialog.")
										   Else
													 Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to verify the Duration Text ["+dicItems(iCounter)+"] in Attributes Dialog.")
													Fn_WrkflwDegnr_Attributes = False
													Set objDialog = Nothing
													Set objAttributeDialog = Nothing
													Exit Function
										   End If
										   Set objComp = Nothing
				
								Case "ReleaseStatus"
											bReturn = false
											arrValues = split(dicItems(iCounter),",",-1,1)
											For jCounter = 0 to UBound(arrvalues)
													   Set objComp =  objAttributeDialog.JavaList("Release Status")
													   iElemCount = objComp.GetROProperty("items count")
													   For kCounter=0 To iElemCount-1
																	If 	objComp.GetItem(kCounter) <> empty	 Then
																			If  Trim(cstr(objComp.GetItem(kCounter)))=Trim(cstr(arrValues(jCounter)))         Then
																					bReturn=True
																					Exit For
																			End If
																	End If
														Next
											Next
											If bReturn=true  Then
													' log result 
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Release Status ["+dicItems(iCounter)+"] Verified Successfully in Attributes Dialog.")
													Fn_WrkflwDegnr_Attributes = True
											Else
													'Report error when item not present in the list
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Release Status ["+dicItems(iCounter)+"] Verification Failed in Attributes Dialog.")
													Fn_WrkflwDegnr_Attributes = False
													Set objDialog = Nothing
													Set objAttributeDialog = Nothing
													Exit Function
											End If
								 End Select
								 End if
							  Next
				
					   Case "Add"
										For iCounter = 0 to dicCount - 1
												If  dicItems(iCounter) <> "" Then
														Select Case dicKeys(iCounter)	
																	Case "ReleaseStatus"
																				If dicItems(iCounter) = "unset" Then
																				 objAttributeDialog.JavaList("Release Status").Select ""
																				Else
																				 objAttributeDialog.JavaList("Release Status").Select dicItems(iCounter)
																				End If
																				If Err.Number < 0 Then
																						'Report error when item not present in the list
																						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fail to Set Release Status Field to ["+dicItems(iCounter)+"] in Attributes Dialog.")
																						Fn_WrkflwDegnr_Attributes = False
																						Set objDialog = Nothing
																						Set objAttributeDialog = Nothing
																						Exit Function
																				Else											
																						' log result 
																						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Release Status Field to ["+dicItems(iCounter)+"] in Attributes Dialog.")
																						Fn_WrkflwDegnr_Attributes = True
																				End If
														  End Select
												  End If
										   Next
				
						Case "VerifyDisplayValue"
											For iCounter = 0 to dicCount - 1
													If  dicItems(iCounter) <> "" Then
															Select Case dicKeys(iCounter)	
																	Case "ReleaseStatus"
																			   Set objComp =  objAttributeDialog.JavaList("Release Status")
																				strValue = objComp.GetROProperty("value")
																				If trim(strValue) = trim(dicItems(iCounter)) Then									
																						' log result 
																						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verified Release Status Field to ["+dicItems(iCounter)+"] in Attributes Dialog.")
																						Fn_WrkflwDegnr_Attributes = True
																				Else
																						'Report error when item not present in the list
																						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Fail to Verify Release Status Field to ["+dicItems(iCounter)+"] in Attributes Dialog.")
																						Fn_WrkflwDegnr_Attributes = False
																						Set objDialog = Nothing
																						Set objAttributeDialog = Nothing
																						Exit Function
																				End If
															End Select
													End If
											 Next

							Case "Close"
											 objAttributeDialog.Close
											 If Err.Number < 0 Then
												   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Close Attributes Dialog.")
													Fn_WrkflwDegnr_Attributes = False
													Set objDialog = Nothing
													Set objAttributeDialog = Nothing
													Exit Function
											 Else                      
												   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Closed Attributes Dialog.")
												   Call Fn_ReadyStatusSync(2)
													Wait(2)
											 End If
					End Select
	End If
	Fn_WrkflwDegnr_Attributes = True
 End Function


 '*********************************************************  Function perform the worklist Designer - Adding Descrition/Instruction for the ProcessName *********************************************************************
'Function Name  :   Fn_WrkflwDegnr_TempleteDetails
'
'Description    :        Workflow Designer : Adding Descrition/Instruction for the ProcessName
' 
'Parameters      :      sAction: Modify
'           			sProcessName: Name of the template to be selected 
'						sProcessTreeNode: Process Template Tree Node Path
'						sName: Name of Template to be changed or verified
'						sDesc: Description wants to add for Process Template
'						bEdit: "Yes" - When template to be made Available 
'						sOther: "" (for Future Parameter, right now keep it blank)
						
' 
'Return Value     :   True/False
'
'Examples    :     Fn_WrkflwDegnr_TempleteDetails("Modify","Template Name","","","New Description","Yes","")
' 
'History:          Developer Name   Date    						Rev. No.   Changes Done   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           			Prasanna    		03-Nov-2010   1.0               
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

 Public Function Fn_WrkflwDegnr_TempleteDetails(sAction,sProcessName,sProcessTreeNode,sName,sDesc,bEdit,sOther)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_TempleteDetails"
	On Error Resume Next
    Dim bReturn

    Select Case sAction
		Case "Modify"
				'Set the Process Name & Set the Edit Mode
				If sProcessName <> "" Then
	                    bReturn=Fn_WrkflwDegnr_SetEditStage(sProcessName,"Yes")
						If bReturn=false Then
									Fn_WrkflwDegnr_TempleteDetails = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Template to Edit Mode." )   									
									Exit Function 
						Else
									Fn_WrkflwDegnr_TempleteDetails = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Template to Edit Mode.")	
						End If
				End If

				'Select the ProcessTree Node
				If sProcessTreeNode <> "" Then
	                    bReturn=Fn_WrkflwDegnr_ProcessTemplateTree_Operations("Select",sProcessTreeNode, "")
						If bReturn=false Then
									Fn_WrkflwDegnr_TempleteDetails = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Process Tree Node [ "+sProcessTreeNode+"]")   									
									Exit Function 
							Else
									Fn_WrkflwDegnr_TempleteDetails = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : uccessfully Select Process Tree Node [ "+sProcessTreeNode+"]")	
							End If
				End If

				'Change the Name of Template
				If sName <> "" Then
						JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaEdit("WfName").Set sName
						If Err.Number < 0 Then
									Fn_WrkflwDegnr_TempleteDetails = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Name ["+sName+"] for the Template")   									
									Exit Function 
						Else
									Fn_WrkflwDegnr_TempleteDetails = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Name ["+sName+"] for the Template")	
						End If
				End If

				'Change the Description of Template
				If sDesc <> "" Then
						JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaEdit("WfDescription").Set sDesc
						If Err.Number < 0 Then
									Fn_WrkflwDegnr_TempleteDetails = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Name ["+sDesc+"] for the Template")   									
									Exit Function 
						Else
									Fn_WrkflwDegnr_TempleteDetails = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Name ["+sDesc+"] for the Template")	
						End If
				End If

				'Make Template Available if  bEdit = Yes
				If bEdit <> "" Then
	                    bReturn = Fn_WrkflwDegnr_SetAvailStage(bEdit)
							If bReturn = false Then
									Fn_WrkflwDegnr_TempleteDetails = False
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Template to Available Mode." )   									
									Exit Function 
						Else
									Fn_WrkflwDegnr_TempleteDetails = True
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Template to Available Mode.")	
						End If
				End If
	End Select

End Function

'*********************************************************  Function for Apply Template Changes *********************************************************************
'Function Name  :     Fn_WrkflwDegnr_ApplyTemplateChanges(bApplyTempChange, bUpdateProcess, sBtnName)
'
'Description    :        Apply Template Changes
' 
'Parameters      :      bApplyTempChange: Apply Template Changes to all workflow Processes
'           			bUpdateProcess: Update Process
'						sBtnName: Button to be Clicked					
' 
'Return Value     :   True/False
'
'Examples    :     Fn_WrkflwDegnr_ApplyTemplateChanges(true, false, "OK")
' 
'History:          Developer Name   					Date    		Rev. No.   Changes Done   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           			Mahendra Bhandarkar    		12-Nov-2010   	1.0              
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WrkflwDegnr_ApplyTemplateChanges(bApplyTempChange, bUpdateProcess, sBtnName)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_ApplyTemplateChanges"
	Dim objDialog, sTextMsg, sTextVal
	Set objDialog = Window("TeamcenterWin").Dialog("ApplyTemplateChanges")
	If objDialog.Exist(2) = True Then
			objDialog.Activate
			If bApplyTempChange <> "" Then
				If bApplyTempChange = True Then
					sTextMsg = "Checked"
					sTextVal = "ON"
				ElseIf bApplyTempChange = False Then
					sTextMsg = "Unchecked"
					sTextVal = "OFF"
				End If
		
				objDialog.WinCheckBox("Applytemplatechanges").Set sTextVal
				 If Err.Number < 0 Then
						Fn_WrkflwDegnr_ApplyTemplateChanges = False
						objDialog.WinButton("Cancel").Click micLeftBtn
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to ["+sTextMsg+"] to Apply Changes to all Active Workflow Processes." )   									
						Exit Function
				Else
						Fn_WrkflwDegnr_ApplyTemplateChanges = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully ["+sTextMsg+"] to Apply Changes to all Active Workflow Processes..")
				End If
			End If
		
			If bUpdateProcess <> "" Then
				If bUpdateProcess = True Then
					sTextMsg = "Checked"
					sTextVal = "ON"
				ElseIf bUpdateProcess = False Then
					sTextMsg = "Unchecked"
					sTextVal = "OFF"
				End If
				objDialog.WinCheckBox("Updateprocesses").Set sTextVal
				If Err.Number < 0 Then
						Fn_WrkflwDegnr_ApplyTemplateChanges = False
						objDialog.WinButton("Cancel").Click micLeftBtn
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to ["+sTextMsg+"] to Update Process in background." )   									
						Exit Function
				Else
						Fn_WrkflwDegnr_ApplyTemplateChanges = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully ["+sTextMsg+"] to Update Process in background.")
				End If
			End If
		
			If sBtnName <> "" Then
				objDialog.WinButton(sBtnName).Click micLeftBtn
				If Err.Number < 0 Then
						Fn_WrkflwDegnr_ApplyTemplateChanges = False
						objDialog.WinButton("Cancel").Click micLeftBtn
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on WinButton ["+sBtnName+"]." )   									
						Exit Function
				Else
						Fn_WrkflwDegnr_ApplyTemplateChanges = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked to WinButton ["+sBtnName+"].")
				End If
			End If
	Else
					Fn_WrkflwDegnr_ApplyTemplateChanges = False
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Apply Template Changes dialog does not exist." )
	End If
	Set objDialog = Nothing
End Function


'*********************************************************  Function for Apply Template Changes *********************************************************************
'Function Name  :     Fn_WrkflwDegnr_NamedACLCreate(sTempName, sTempNode, sACLName, arrACLDetails)
'
'Description    :        Creates Named ACL for tempalte
' 
'Parameters      :      

'Return Value     :   True/False
'
'Examples    :     
' 
'History:          Developer Name   					Date    		Rev. No.   Changes Done   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           			Vallari S				    		1-Jan-2011		   	1.0              
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'      					Ashwini P				    		26-Mar-2014		   	1.1              
'						Shweta Rathod						05-Apr-2017			1.2	 									
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwDegnr_NamedACLCreate(sTempName, sTempNode, sACLName, arrACLDetails)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_NamedACLCreate"
	Dim bReturn
	Dim objDialog, objAttributeDialog
	Dim iRows, iCols
	Dim iRowCnt, iColCnt
	Dim sVal

	Fn_WrkflwDegnr_NamedACLCreate = True
	iRows = UBound(arrACLDetails, 1)
	iCols = UBound(arrACLDetails, 2)
	
	'Select the Template and put it in Edit mode
	bReturn = Fn_WrkflwDegnr_SetEditStage(sTempName, "Yes")
	If bReturn = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Template ["+CStr(sTempName)+"] to Edit Mode")
		Fn_WrkflwDegnr_NamedACLCreate = False
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Template ["+CStr(sTempName)+"] to Edit Mode")
		Call Fn_ReadyStatusSync(3)
	End If

   'Expand and Select the Process Tree Node
   bReturn = Fn_WrkflwDegnr_ProcessTemplateTree_Operations("ExpandSelect", sTempNode, "")
	If bReturn = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Process Template Node ["+CStr(sTempNode)+"]")
		Fn_WrkflwDegnr_NamedACLCreate = False
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Process Template Node ["+CStr(sTempNode)+"]")
		Call Fn_ReadyStatusSync(3)
	End If

	Set objDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame")
	Set objAttributeDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes")
	Set objACLDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Named ACL")
	'Click Attribute button to invoke Attribute panel
	objDialog.JavaCheckBox("TskAttributeBtn").Set "ON"
	 If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Checkbox.")
		Fn_WrkflwDegnr_NamedACLCreate = False
		Set objAttributeDialog = Nothing
		Set objDialog = Nothing
		Exit Function
	 Else                      
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Checkbox.")
	   Call Fn_ReadyStatusSync(3)
		Wait(2)
	 End If

	'Click on Attributes text
	If objAttributeDialog.exist(5) = false then              '[TC1123-20170327b-05_APR_2017-ShwetaR-NewDevelopment] - added condition to check existence of dialog for performing double click
		objDialog.JavaStaticText("Attributes").DblClick 1, 1, "LEFT"
		If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Text.")
				Fn_WrkflwDegnr_NamedACLCreate = False
				Set objAttributeDialog = Nothing
				Set objDialog = Nothing
				Exit Function
		Else                      
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Text.")
			  Call Fn_ReadyStatusSync(3)
			  Wait(5)					
		End If
	End if
	
	'Click on ACL button
'	objAttributeDialog.JavaCheckBox("NamedACL").Click 5,5,"LEFT"  'Commented b y Nilesh on 27-Jul-2012
	objAttributeDialog.JavaCheckBox("NamedACL").Set "ON"
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on NamedACL Button.")
			Fn_WrkflwDegnr_NamedACLCreate = False
			objAttributeDialog.Close
			Set objAttributeDialog = Nothing
			Set objDialog = Nothing
			Exit Function
	Else                      
		  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on NamedACL Button.")
		  Call Fn_ReadyStatusSync(3)
		  Wait(2)					
	End If

	'Click on Named ACL text
	'Added Nilesh on 27-Jul-2012
	If objACLDialog.Exist(5)=False Then
		objDialog.JavaStaticText("NamedACL").DblClick 2,2,"LEFT"
		If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Named ACL Text.")
				Fn_WrkflwDegnr_NamedACLCreate = False
				objAttributeDialog.Close
				Set objAttributeDialog = Nothing
				Set objDialog = Nothing
				Exit Function
		Else                      
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Named ACL Text.")
			  Wait(2)					
		End If
	End If

	'Set the ACL Name & Click on Create button
	objACLDialog.JavaCheckBox("WorkflowNamedACL").Set "ON"
	wait 1
	objACLDialog.JavaEdit("ACLName").Set ""
	objACLDialog.JavaEdit("ACLName").Type sACLName
	wait 1
'--------------------End----------------------------------------------------------------------------------------------------------	
	'objACLDialog.JavaButton("CreateACL").Click micLeftBtn     
	objACLDialog.JavaButton("CreateACL").Object.doClick(5)  '[TC1123-20170327b-05_APR_2017-ShwetaR-NewDevelopment] - added as .click method is not supporting to the build 0327b 
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Create New ACL [" + sACLName + "] Name")
			Fn_WrkflwDegnr_NamedACLCreate = False
			objACLDialog.Close
			objAttributeDialog.Close
			Set objACLDialog = Nothing
			Set objAttributeDialog = Nothing
			Set objDialog = Nothing
			Exit Function
	Else                      
		  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Created New ACL [" + sACLName + "] Name")
		  Wait(2)					
	End If

	For iRowCnt = 0 to iRows -1
		'objACLDialog.JavaButton("AddRow").Click micLeftBtn
		objACLDialog.JavaButton("AddRow").Object.doClick(5)  '[TC1123-20170327b-05_APR_2017-ShwetaR-NewDevelopment] - added as .click method is not supporting to the build 0327b 
		For iColCnt = 0 to iCols - 1
			If Lcase(arrACLDetails(iRowCnt, iColCnt)) = "yes" Then
				sVal = "yes"
			Elseif Lcase(arrACLDetails(iRowCnt, iColCnt)) = "no" Then
				sVal = "no"
			Else
				sVal = arrACLDetails(iRowCnt, iColCnt)
			End If

			If sVal <> "" Then
	
					If iColCnt <> 1 Then
						objACLDialog.JavaTable("ACLTable").DoubleClickCell iRowCnt, iColCnt, "LEFT", "NONE"
						wait(1)
						objACLDialog.JavaList("ACLTblList").Select sVal
						If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set ACL Deatil Value [" + cstr(arrACLDetails(iRowCnt, iColCnt)) + "] in Row [" + cstr(iRowCnt) + "] & Column [" +cstr(iColCnt) + "]")
								Fn_WrkflwDegnr_NamedACLCreate = False
								objACLDialog.Close
								objAttributeDialog.Close
								Set objACLDialog = Nothing
								Set objAttributeDialog = Nothing
								Set objDialog = Nothing
								Exit Function					
						End If
					Else
						objACLDialog.JavaTable("ACLTable").DoubleClickCell iRowCnt, iColCnt, "LEFT", "NONE"
						wait(2)
						objDialog.JavaDialog("SelectAccessor").Activate
						objDialog.JavaDialog("SelectAccessor").JavaList("Accessors").Select cstr(arrACLDetails(iRowCnt, iColCnt))
						objDialog.JavaDialog("SelectAccessor").JavaButton("OK").Click micLeftBtn
						If Err.Number < 0 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select ACL Accessor [" + cstr(arrACLDetails(iRowCnt, iColCnt)) + "]")
								Fn_WrkflwDegnr_NamedACLCreate = False
								objDialog.JavaDialog("SelectAccessor").JavaButton("Cancel").Click micLeftBtn
								objACLDialog.Close
								objAttributeDialog.Close
								Set objACLDialog = Nothing
								Set objAttributeDialog = Nothing
								Set objDialog = Nothing
								Exit Function					
						End If
					End If
			End If
		Next

	Next

	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Created Named ACL [" + sACLName + "]")

	'Save the newly created ACL
	'objACLDialog.JavaButton("SaveACL").Click micLeftBtn
	objACLDialog.JavaButton("SaveACL").Object.doClick(5)   '[TC1123-20170327b-05_APR_2017-ShwetaR-NewDevelopment] - added as .click method is not supporting to the build 0327b 
	Call Fn_ReadyStatusSync(3)
	'Click Assign button
	'objACLDialog.JavaButton("Assign").Click micLeftBtn
	objACLDialog.JavaButton("Assign").Object.doClick(5) '[TC1123-20170327b-05_APR_2017-ShwetaR-NewDevelopment] - added as .click method is not supporting to the build 0327b 
	Call Fn_ReadyStatusSync(3)
	wait(5)
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Named ACL [" + sACLName + "] to the Workflow Task")
			Fn_WrkflwDegnr_NamedACLCreate = False
			objACLDialog.Close
			objAttributeDialog.Close
			Set objACLDialog = Nothing
			Set objAttributeDialog = Nothing
			Set objDialog = Nothing
			Exit Function					
	End If

	objACLDialog.Close
	objAttributeDialog.JavaCheckBox("NamedACL").Click 5,5,"LEFT"
	objAttributeDialog.JavaCheckBox("NamedACL").WaitProperty "label", sACLName, 100000

	If JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes").JavaCheckBox("NamedACL").GetROProperty("label") <> sACLName Then
		Fn_WrkflwDegnr_NamedACLCreate = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Named ACL [" + sACLName + "] to the Workflow Task")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Assigned Named ACL [" + sACLName + "] to the Workflow Task")
	End If

	'Added By Rima 
	If 	objACLDialog.Exist Then
		objACLDialog.Close
	End If
	objAttributeDialog.Close
	Set objACLDialog = Nothing
	Set objAttributeDialog = Nothing
	Set objDialog = Nothing

	'Set Workflow Template to Available state
	bReturn = Fn_WrkflwDegnr_SetAvailStage("Yes")
	If bReturn = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Template ["+CStr(sTempName)+"] to Available Mode")
		Fn_WrkflwDegnr_NamedACLCreate = False
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Template ["+CStr(sTempName)+"] to Available Mode")
		Call Fn_ReadyStatusSync(3)
	End If

End Function


'*********************************************************  Function for Apply Template Changes *********************************************************************
'Function Name  :     Fn_WrkflwDegnr_NamedACLAssign(sTempName, sTempNode, sACLName)
'
'Description    :        Assigns Named ACL for tempalte
' 
'Parameters      :      

'Return Value     :   True/False
'
'Examples    :     
' 
'History:          Developer Name   					Date    		Rev. No.   Changes Done   Reviewer 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'           			Vallari S				    		3-Jan-2011		   	1.0              
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwDegnr_NamedACLAssign(sTempName, sTempNode, sACLName)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_NamedACLAssign"
	Dim bReturn
	Dim objDialog, objAttributeDialog
	Dim iItem

	Fn_WrkflwDegnr_NamedACLAssign = True
	
	'Select the Template and put it in Edit mode
	bReturn = Fn_WrkflwDegnr_SetEditStage(sTempName, "Yes")
	If bReturn = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Template ["+CStr(sTempName)+"] to Edit Mode")
		Fn_WrkflwDegnr_NamedACLAssign = False
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Template ["+CStr(sTempName)+"] to Edit Mode")
		Call Fn_ReadyStatusSync(3)
	End If

   'Expand and Select the Process Tree Node
   bReturn = Fn_WrkflwDegnr_ProcessTemplateTree_Operations("ExpandSelect", sTempNode, "")
	If bReturn = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Process Template Node ["+CStr(sTempNode)+"]")
		Fn_WrkflwDegnr_NamedACLAssign = False
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Process Template Node ["+CStr(sTempNode)+"]")
		Call Fn_ReadyStatusSync(3)
	End If

	Set objDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame")
	Set objAttributeDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Attributes")

	'Click Attribute button to invoke Attribute panel
	objDialog.JavaCheckBox("TskAttributeBtn").Set "ON"
	 If Err.Number < 0 Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Checkbox.")
		Fn_WrkflwDegnr_NamedACLAssign = False
		Set objAttributeDialog = Nothing
		Set objDialog = Nothing
		Exit Function
	 Else                      
	   Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Checkbox.")
	   Call Fn_ReadyStatusSync(3)
		Wait(2)
	 End If

	If objAttributeDialog.exist(5) = false then              '[TC1123-20170327b-05_APR_2017-ShwetaR-NewDevelopment] - added condition to check the existence of dialg
	'Click on Attributes text
	objDialog.JavaStaticText("Attributes").DblClick 1, 1, "LEFT"
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Attributes Text.")
			Fn_WrkflwDegnr_NamedACLAssign = False
			Set objAttributeDialog = Nothing
			Set objDialog = Nothing
			Exit Function
	Else                      
		  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Attributes Text.")
		  Wait(2)					
	End If
	End if
	'Click on ACL button
       objAttributeDialog.JavaCheckBox("NamedACL").Set "ON"
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on NamedACL Button.")
			Fn_WrkflwDegnr_NamedACLAssign = False
			objAttributeDialog.Close
			Set objAttributeDialog = Nothing
			Set objDialog = Nothing
			Exit Function
	Else                      
		  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on NamedACL Button.")
		  Call Fn_ReadyStatusSync(3)
		  Wait(2)					
	End If

	Set objACLDialog = JavaWindow("WorkflowDesignerWindow").JavaWindow("WEmbeddedFrame").JavaDialog("Named ACL")
	If  objACLDialog.Exist(5)=False Then 'Added by Nilesh 29-Aug-2012
		'Click on Named ACL text
		objDialog.JavaStaticText("NamedACL").DblClick 2,2,"LEFT"
		If Err.Number < 0 Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Named ACL Text.")
				Fn_WrkflwDegnr_NamedACLAssign = False
				objAttributeDialog.Close
				Set objAttributeDialog = Nothing
				Set objDialog = Nothing
				Exit Function
		Else                      
			  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Named ACL Text.")
			  Wait(2)					
		End If
	End If

	

	'Select Named ACL
	objACLDialog.JavaCheckBox("WorkflowNamedACL").Set "ON"
	iItem = objACLDialog.JavaList("ACLSelect").GetItemIndex(sACLName)
	objACLDialog.JavaList("ACLSelect").Object.setSelectedIndex iItem,True
'	objACLDialog.JavaList("ACLSelect").Select sACLName
'	objACLDialog.JavaEdit("ACLName").Activate
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Named ACL [" + sACLName + "]")
			Fn_WrkflwDegnr_NamedACLAssign = False
			objACLDialog.Close
			objAttributeDialog.Close
			Set objACLDialog = Nothing
			Set objAttributeDialog = Nothing
			Set objDialog = Nothing
			Exit Function
	Else                      
		  Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Named ACL [" + sACLName + "]")
		  Wait(2)					
	End If

	'Click Assign button
	'objACLDialog.JavaButton("Assign").Click micLeftBtn
	objACLDialog.JavaButton("Assign").Object.doClick(5)         '[TC1123-20170327b-05_APR_2017-ShwetaR-NewDevelopment] - added as .click method is not supporting to the build 0327b 
	wait(2)
	If Err.Number < 0 Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Named ACL [" + sACLName + "] to the Workflow Task")
			Fn_WrkflwDegnr_NamedACLAssign = False
			objACLDialog.Close
			objAttributeDialog.Close
			Set objACLDialog = Nothing
			Set objAttributeDialog = Nothing
			Set objDialog = Nothing
			Exit Function					
	End If

	objACLDialog.Close
'	objAttributeDialog.JavaCheckBox("NamedACL").Click 5,5,"LEFT"
	If sACLName = "" Then
		sACLName = "namedacl_16"
	End If
	objAttributeDialog.JavaCheckBox("NamedACL").WaitProperty "label", sACLName, 100000

	If objAttributeDialog.JavaCheckBox("NamedACL").GetROProperty("label") <> sACLName Then
		Fn_WrkflwDegnr_NamedACLAssign = False
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Assign Named ACL [" + sACLName + "] to the Workflow Task")
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Assigned Named ACL [" + sACLName + "] to the Workflow Task")	
	End If

	objAttributeDialog.Close
	Set objACLDialog = Nothing
	Set objAttributeDialog = Nothing
	Set objDialog = Nothing

	'Set Workflow Template to Available state
	bReturn = Fn_WrkflwDegnr_SetAvailStage("Yes")
	If bReturn = False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Set Process Template ["+CStr(sTempName)+"] to Available Mode")
		Fn_WrkflwDegnr_NamedACLAssign = False
		Exit Function
	Else
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Set Process Template ["+CStr(sTempName)+"] to Available Mode")
		Call Fn_ReadyStatusSync(3)
	End If

End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name  : Fn_WrkflwDegnr_NewRootTemplateOperations
'Description    : Function Used to Create New Root Template
'Return Value     :  True Or False
'Examples    :  Case "Set" : Fn_WrkflwDegnr_NewRootTemplateOperations("Set", "TestRoot1", "AutoDelgReview", "Process", "OK")
             
'History      :   
'             Developer Name            Date      				Rev. No.      Changes Done      Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'             Ketan Raje.                 05/01/2011              1.0                   						Harshal
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_WrkflwDegnr_NewRootTemplateOperations(sAction, sNewRoot, sBasedRoot, sTemplateType, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_NewRootTemplateOperations"
   'Declaring Variables
    Dim aButtons, iCount, objSelectType, ObjChangeWnd, intNoOfObjects
 Fn_WrkflwDegnr_NewRootTemplateOperations=False
' Set ObjChangeWnd=Fn_UI_ObjectCreate("Fn_WrkflwDegnr_NewRootTemplateOperations",JavaWindow("WorkflowDesignerWindow").JavaWindow("WrkflwDegnrWin").JavaDialog("NewRootTemplate"))
'Added by Pallavi Patil on 26 Jun 2012
	Set ObjChangeWnd= Fn_SISW_WorkflowDesigner_GetObject("NewRootTemplate")
  Select Case sAction
		Case "Set"
				'Set New Root Template Name
				If sNewRoot<>"" Then
					Call Fn_Edit_Box("Fn_WrkflwDegnr_NewRootTemplateOperations",ObjChangeWnd,"NewRootTemplateName",sNewRoot)
				End If
				'Set Based On Root Template
				If sBasedRoot<>"" Then
					'Click on Based On Root Template DropDown.
					Call Fn_Button_Click("Fn_WrkflwDegnr_NewRootTemplateOperations", ObjChangeWnd, "BasedOnRootTemplate")
				   Set objSelectType=description.Create()
				   objSelectType("Class Name").value = "JavaStaticText"
				   objSelectType("label").value = sBasedRoot
				   Set  intNoOfObjects = ObjChangeWnd.ChildObjects(objSelectType)
					  intNoOfObjects(0).Click 1,1
					  wait(3)
					'Call Fn_Edit_Box("Fn_WrkflwDegnr_NewRootTemplateOperations",ObjChangeWnd,"NewRootTemplateName",sNewRoot)
				End If
				'Set Template Type
				If sTemplateType<>"" Then
					If Trim(Lcase(sTemplateType)) = "process" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_WrkflwDegnr_NewRootTemplateOperations",ObjChangeWnd, "Process")
					ElseIf Trim(Lcase(sTemplateType)) = "task" Then
						Call Fn_UI_JavaRadioButton_SetON("Fn_WrkflwDegnr_NewRootTemplateOperations",ObjChangeWnd, "Task")
					End If
				End If
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sNewRoot &"New Root Template set successfully")
				Fn_WrkflwDegnr_NewRootTemplateOperations = TRUE
		Case Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_WrkflwDegnr_NewRootTemplateOperations function failed")
				Fn_WrkflwDegnr_NewRootTemplateOperations = FALSE
  End Select
 'Click on Buttons
 If sButtons<>"" Then
   aButtons = split(sButtons, ":",-1,1)
   For iCount=0 to Ubound(aButtons)
    'Click on Add Button
    Call Fn_Button_Click("Fn_WrkflwDegnr_NewRootTemplateOperations", ObjChangeWnd, aButtons(iCount))
   Next
 End If
 Set ObjChangeWnd = Nothing
 Set objSelectType = Nothing
 Set intNoOfObjects = Nothing
End Function

'*********************************************************  Function do Operation on Edit Template Filter apply dialog *********************************************************************

'Function Name		:			Fn_WrkflwDegnr_EditTemplateFilter_Operations

'Description		:		 		Sets the template filter

'Parameters			   :	 		1. sAction: Action to be Performed
'												 2. sGroup: Group name to be selected
'												3. sType: Object type to be selected
'												4. AssProcessTemplate: Assigned Process Templatessss
'												4. DefinedProcessTemplate: Defined Process Template
'												4. sButon: aArray of templates to be selected

'Return Value		   : 			True/False

'Pre-requisite			:		 	user is logged in to Wrokflow Designer module

'Examples				:			 Fn_WrkflwDegnr_EditTemplateFilter_Operations("Remove", "dba", "XMLAuditLog", Array("AutoCondDoWrkFlw"),"", "Cancel")
'												 Fn_WrkflwDegnr_EditTemplateFilter_Operations("VerifyAssigned", "dba", "XMLAuditLog", Array("AutoCondDoWrkFlw"),"", "Cancel")
'
'History:
'										Developer Name						Date			Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Mahendra Bhandarkar		26-Nov-2010	       1.0																Prasanna
'										Nilesh Gadekar					4-July-2012	       1.1						Added "Clear" case
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_WrkflwDegnr_EditTemplateFilter_Operations(sAction, sGroup, sType, AssProcessTemplate, DefinedProcessTemplate, sButton)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_EditTemplateFilter_Operations"
	Dim objDialog, bReturn, iCounter, WshShell

	Set objDialog = Fn_SISW_WorkflowDesigner_GetObject("ProcessTemplateFilter")

	If objDialog.Exist(5) = False Then 
		bReturn = Fn_MenuOperation("Select", "Edit:Template Filter")
		If bReturn = False Then
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Open Edit --> Template Filter ") 		
			Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
			Set objDialog = Nothing
			Exit Function					
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Opened Edit --> Template Filter")
			Call Fn_ReadyStatusSync(2)
		End If
	End If

	If objDialog.Exist(5) Then

			Select Case sAction

				'To Add the teamplate filter
				Case "Add"
		
							bReturn = Fn_WrkflwDegnr_TemplateFilterApply(sGroup, sType, DefinedProcessTemplate)
							If bReturn = False Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Apply Template Filter.") 		
								Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
								Set objDialog = Nothing
								Exit Function	
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Applied Template Filter.")
								Call Fn_ReadyStatusSync(2)
							End If
				'To Remove the teamplate filter
				Case "Remove"
		
							If Trim(sGroup) <> "" Then
								objDialog.JavaList("GroupName").Select sGroup
								Wait(2)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sGroup)+"]") 		
									Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function	
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sGroup)+"]")
									Call Fn_ReadyStatusSync(2)
								End If
							End If
			
							If Trim(sType) <> "" Then
								objDialog.JavaEdit("ObjectType").Type sType
								Wait(2)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Object Type ["+CStr(sType)+"] ")
									Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Object Type ["+CStr(sType)+"] ")
									Call Fn_ReadyStatusSync(2)
								End If
							End If
							Set WshShell = CreateObject("WScript.Shell")
							Wait(3)
							WshShell.SendKeys "{ENTER}"
							Wait(2)
							Set WshShell = Nothing			
							If IsArray(AssProcessTemplate) Then
								For iCounter = 0 To UBound(AssProcessTemplate)
									If Trim(AssProcessTemplate(iCounter)) <> "" Then
										If Fn_UI_ListItemExist("Fn_WrkflwDegnr_EditTemplateFilter_Operations", objDialog, "AssignedProcTemp",AssProcessTemplate(iCounter)) = True Then
											'Select Template Name from Assigned Process Template
											objDialog.JavaList("AssignedProcTemp").Select Trim(AssProcessTemplate(iCounter))
											Wait(2)
											If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Assigned Process Template ["+CStr(AssProcessTemplate(iCounter))+"] ")
												Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
												objDialog.JavaButton("Cancel").Click micLeftBtn
												Set objDialog = Nothing
												Exit Function
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Assigned Process Template ["+CStr(AssProcessTemplate(iCounter))+"] ")
												Call Fn_ReadyStatusSync(2)
											End If
											'Click on Send Right Button
											objDialog.JavaButton("Right").Click micLeftBtn
											If Err.Number < 0 Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Left Button ") 		
												Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
												objDialog.JavaButton("Cancel").Click micLeftBtn
												Set objDialog = Nothing
												Exit Function						
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Left Button")
												Call Fn_ReadyStatusSync(2)
											End If
										Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : List Item ["+CStr(AssProcessTemplate(iCounter))+"] Does Not Exist in Defined Process Template")
												Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
												objDialog.JavaButton("Cancel").Click micLeftBtn
												Set objDialog = Nothing
												Exit Function
										End If
									End If
								Next
							End If
							If Trim(sButton) <> "" Then
								objDialog.JavaButton(sButton).Click micLeftBtn
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] Button ") 		
									Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function		
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on ["+sButton+"] Button")
									Call Fn_ReadyStatusSync(2)
								End If
							End If
		
				Case "VerifyAssigned"
		
							If Trim(sGroup) <> "" Then
								objDialog.JavaList("GroupName").Select sGroup
								Wait(2)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sGroup)+"]") 		
									Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function	
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sGroup)+"]")
									Call Fn_ReadyStatusSync(2)
								End If
							End If
					
							If Trim(sType) <> "" Then
								objDialog.JavaEdit("ObjectType").Type sType
								Wait(2)
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Object Type ["+CStr(sType)+"] ")
									Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Object Type ["+CStr(sType)+"] ")
									Call Fn_ReadyStatusSync(2)
								End If
							End If
							Set WshShell = CreateObject("WScript.Shell")
							Wait(3)
							WshShell.SendKeys "{ENTER}"
							Wait(2)
							Set WshShell = Nothing			
							If IsArray(AssProcessTemplate) Then
								For iCounter = 0 To UBound(AssProcessTemplate)
									If Trim(AssProcessTemplate(iCounter)) <> "" Then
										'Verify Whether Assigned Template Contains the template or not
										bReturn = Fn_UI_ListItemExist("Fn_WrkflwDegnr_EditTemplateFilter_Operations", objDialog, "AssignedProcTemp",AssProcessTemplate(iCounter))
											If bReturn = False Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Verify Assigned Process Template ["+CStr(AssProcessTemplate(iCounter))+"] ")
												Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
												objDialog.JavaButton("Cancel").Click micLeftBtn
												Set objDialog = Nothing
												Exit Function
											Else
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Verify Assigned Process Template ["+CStr(AssProcessTemplate(iCounter))+"] ")
												Call Fn_ReadyStatusSync(2)
											End If
									End If
								Next
							End If
					
							If Trim(sButton) <> "" Then
									objDialog.JavaButton(sButton).Click micLeftBtn
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] Button ") 		
										Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
										objDialog.JavaButton("Cancel").Click micLeftBtn
										Set objDialog = Nothing
										Exit Function		
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on ["+sButton+"] Button")
										Call Fn_ReadyStatusSync(2)
									End If
							End If
			Case "Clear"
								If Trim(sGroup) <> "" Then
										objDialog.JavaList("GroupName").Select sGroup
										Wait(2)
										If Err.Number < 0 Then
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Group ["+CStr(sGroup)+"]") 		
											Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
											objDialog.JavaButton("Cancel").Click micLeftBtn
											Set objDialog = Nothing
											Exit Function	
										Else
											Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Selected Group ["+CStr(sGroup)+"]")
											Call Fn_ReadyStatusSync(2)
										End If
								End If
			
								If Trim(sType) <> "" Then
									objDialog.JavaEdit("ObjectType").Set sType
									Wait(2)
									objDialog.JavaEdit("ObjectType").Type MicReturn
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Select Object Type ["+CStr(sType)+"] ")
										Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
										objDialog.JavaButton("Cancel").Click micLeftBtn
										Set objDialog = Nothing
										Exit Function
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Select Object Type ["+CStr(sType)+"] ")
										Call Fn_ReadyStatusSync(2)
									End If
								End If
								objDialog.JavaButton("Clear").Click micLeftBtn
								Wait 2
								objDialog.JavaButton("Apply").Click micLeftBtn
								Wait 2
								If Err.Number < 0 Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on Clear Button ") 		
									Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
									objDialog.JavaButton("Cancel").Click micLeftBtn
									Set objDialog = Nothing
									Exit Function		
								Else
									Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on Clear Button")
									Call Fn_ReadyStatusSync(2)
								End If

								If Trim(sButton) <> "" Then
									objDialog.JavaButton(sButton).Click micLeftBtn
									If Err.Number < 0 Then
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL : Failed to Click on ["+sButton+"] Button ") 		
										Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
										objDialog.JavaButton("Cancel").Click micLeftBtn
										Set objDialog = Nothing
										Exit Function		
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS : Successfully Clicked on ["+sButton+"] Button")
										Call Fn_ReadyStatusSync(2)
									End If
							End If

			End Select

		Fn_WrkflwDegnr_EditTemplateFilter_Operations = True

	Else

		Fn_WrkflwDegnr_EditTemplateFilter_Operations = False
		Set objDialog = Nothing
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "FAIL: Fn_WrkflwDegnr_EditTemplateFilter_Operations function failed")
	End If

	Set objDialog = Nothing

End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Fn_WrkflwDegnr_ExportWorkflowTemplates(sAction, sDirectory, sFileName, sLanguages, sAllTemplates, sContOnError, sButtons, bViewLog)
'###
'###    DESCRIPTION        :   Set / Verify Export Workflow Templates 
'###	Prequisite 				:	Workflow Designer Prespective should be Open.
'###
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Edit_Box(), Fn_CheckBox_Set, Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    		21/11/2010         1.0
'###                                           Pritam Shikare                  03/July/2012      2.0     Changed the Object Hierarchy of  JavaDialog("ExportWorkflowTemplates")
'###
'###    EXAMPLE          : 		Case "Set" : Msgbox Fn_WrkflwDegnr_ExportWorkflowTemplates("Set", "D:\Mainline", "Pranav", "", "AutoDoReview", "ON", "OK", "Yes")
'#############################################################################################################
Public Function Fn_WrkflwDegnr_ExportWorkflowTemplates(sAction, sDirectory, sFileName, sLanguages, sAllTemplates, sContOnError, sButtons, bViewLog)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_ExportWorkflowTemplates"
	Dim objWrkflw, iReturn, iCount, aAllTemplates, iCnt, aButtons, objExpComp
	Fn_WrkflwDegnr_ExportWorkflowTemplates = False	
	Set objWrkflw = Fn_SISW_WorkflowDesigner_GetObject("ExportWorkflowTemplates")
   'Checking Existance of Export Workflow Template window
   If Fn_UI_ObjectExist("Fn_PLM_ExportObjects",objWrkflw)=False Then
	   'Opening Export Workflow Template window
		Call Fn_MenuOperation("Select","Tools:Export")
   End If
		Select Case sAction
			Case "Set"
						'Set Value for Export Directory
						If sDirectory<>"" Then
							Call Fn_Edit_Box("Fn_WrkflwDegnr_ExportWorkflowTemplates",objWrkflw,"ExportDirectory",sDirectory)
						End If
						'Set Value File Name
						If sFileName<>"" Then
							Call Fn_Edit_Box("Fn_WrkflwDegnr_ExportWorkflowTemplates",objWrkflw,"FileName",sFileName)
						End If
						'Code for Languages part is to be coded as required.
						If sLanguages<>"" Then
							
						End If
						'Transfer Templates from All Templates list to Defined Templates list
						If sAllTemplates<>"" Then
							aAllTemplates = Split(sAllTemplates,"|",-1,1)
							iReturn = Cint(Fn_UI_Object_GetROProperty("Fn_WrkflwDegnr_ExportWorkflowTemplates",objWrkflw.JavaList("AllTemplates"), "items count"))
							For iCount = 0 to Ubound(aAllTemplates)								
								For iCnt = 0 to iReturn-1
									If Trim(Lcase(objWrkflw.JavaList("AllTemplates").GetItem(iCnt))) = Trim(Lcase(aAllTemplates(iCount))) Then
										'Select item from Defined tools list
										objWrkflw.JavaList("AllTemplates").Select iCnt
										'Click on AddColumn button
										Call Fn_Button_Click("Fn_WrkflwDegnr_ExportWorkflowTemplates", objWrkflw, "Add")
										Exit For
									End If
								Next
							Next
						End If
						'Set Continue on Error Checkbox
						If sContOnError<>"" Then
							Call Fn_CheckBox_Set("Fn_WrkflwDegnr_ExportWorkflowTemplates", objWrkflw, "ContinueOnError", sContOnError)
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_WrkflwDegnr_ExportWorkflowTemplates")
						Fn_WrkflwDegnr_ExportWorkflowTemplates = TRUE						
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Fn_WrkflwDegnr_ExportWorkflowTemplates function failed")
						Fn_WrkflwDegnr_ExportWorkflowTemplates = FALSE
						Set objWrkflw = nothing
						Exit Function											
		End Select
		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)				
				For iCount=0 to Ubound(aButtons)
					Call Fn_Button_Click("Fn_WrkflwDegnr_ExportWorkflowTemplates", objWrkflw, aButtons(iCount))
				Next
		End If
		'View Lof for details
		If bViewLog<>"" Then
			'Changed by Pritam Shikare ......     Setbthe object using Fn_SISW_WorkflowDesigner_GetObject function
			Set objExpComp = Fn_SISW_WorkflowDesigner_GetObject("Export Completed")
			Call Fn_Button_Click("Fn_WrkflwDegnr_ExportWorkflowTemplates", objExpComp, bViewLog)
		End If
	Set objWrkflw = nothing
	Set objExpComp = Nothing
End Function
'#########################################################################################################
'###    FUNCTION NAME   :   Fn_WrkflwDegnr_ImportWorkflowTemplates(sAction, sImportFile, sContOnError, sOverwrite, sButtons, bViewLog)
'###
'###    DESCRIPTION        :   Set / Verify Import Workflow Templates 
'###	Prequisite 				:	Workflow Designer Prespective should be Open.
'###
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Edit_Box(), Fn_CheckBox_Set, Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    		21/11/2010         1.0
'###                                           Pritam Shikare                  03/July/2012      2.0     Changed the Object Hierarchy of  JavaDialog("ImportWorkflowTemplates")
'###
'###    EXAMPLE          : 		Case "Set" : Msgbox Fn_WrkflwDegnr_ImportWorkflowTemplates("Set", "D:\Mainline\asdf.xml", "ON", "ON", "OK", "Yes")
'#############################################################################################################
Public Function Fn_WrkflwDegnr_ImportWorkflowTemplates(sAction, sImportFile, sContOnError, sOverwrite, sButtons, bViewLog)
	GBL_FAILED_FUNCTION_NAME="Fn_WrkflwDegnr_ImportWorkflowTemplates"
	Dim objWrkflw, iReturn, iCount, iCnt, aButtons, objImpComp
	Fn_WrkflwDegnr_ImportWorkflowTemplates = False	
	Set objWrkflw = Fn_SISW_WorkflowDesigner_GetObject("ImportWorkflowTemplates")
   'Checking Existance of Export Workflow Template window
   If Fn_UI_ObjectExist("Fn_PLM_ExportObjects",objWrkflw)=False Then
	   'Opening Export Workflow Template window
		Call Fn_MenuOperation("Select","Tools:Import")
   End If
		Select Case sAction
			Case "Set"
						'Click on Browse button.
						Call Fn_Button_Click("Fn_WrkflwDegnr_ImportWorkflowTemplates", objWrkflw, "Browse")
						'Set Value for Import Directory.
						If sImportFile<>"" Then						
							Call Fn_Edit_Box("Fn_WrkflwDegnr_ImportWorkflowTemplates", JavaDialog("SelectDirectory"),"FileName",sImportFile)
						End If
						'Click on Select button.
						Call Fn_Button_Click("Fn_WrkflwDegnr_ImportWorkflowTemplates", JavaDialog("SelectDirectory"), "Select")
						'Set Continue on Error Checkbox
						If sContOnError<>"" Then
							Call Fn_CheckBox_Set("Fn_WrkflwDegnr_ImportWorkflowTemplates", objWrkflw, "ContinueOnError", sContOnError)
						End If
						'Set Overwrite Duplicate Templates Checkbox
						If sOverwrite<>"" Then
							Call Fn_CheckBox_Set("Fn_WrkflwDegnr_ImportWorkflowTemplates", objWrkflw, "OverwriteDuplicateTemplates", sOverwrite)
						End If
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_WrkflwDegnr_ImportWorkflowTemplates")
						Fn_WrkflwDegnr_ImportWorkflowTemplates = TRUE						
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Fn_WrkflwDegnr_ImportWorkflowTemplates function failed")
						Fn_WrkflwDegnr_ImportWorkflowTemplates = FALSE
						Set objWrkflw = nothing
						Exit Function											
		End Select
		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)				
				For iCount=0 to Ubound(aButtons)
					Call Fn_Button_Click("Fn_WrkflwDegnr_ImportWorkflowTemplates", objWrkflw, aButtons(iCount))
				Next
		End If
		'View Lof for details
		If bViewLog<>"" Then
'			JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed").SetTOProperty "title","Import Completed"
'			Call Fn_Button_Click("Fn_PLM_ImportObjects", JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed"), bViewLog)
			'Changed by Pritam Shikare ......     Setbthe object using Fn_SISW_WorkflowDesigner_GetObject function
			Wait(15)
			Set objImpComp = Fn_SISW_WorkflowDesigner_GetObject("Import Completed")
			If Fn_UI_ObjectExist("Fn_PLM_ExportObjects",objImpComp)=False Then
				Wait(15)
			End if
			Call Fn_Button_Click("Fn_PLM_ImportObjects",objImpComp, bViewLog)
		End If
	Set objWrkflw = nothing
	Set objImpComp = Nothing
End Function
