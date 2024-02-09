Option Explicit
'Public  iTimeOut
iTimeOut=180

'--------------------------'Global variables for Teamcenter Perspective Names------------------------------------------------------------
Public GBL_PERSPECTIVE_REQUIREMENTMANAGER
GBL_PERSPECTIVE_REQUIREMENTMANAGER = "Requirement Manager"
'--------------------------'Global variables for Teamcenter Perspective Names------------------------------------------------------------

'*********************************************************	Function List		***********************************************************************
'1. Fn_ReqMgr_RMTabPanelOperation(strAction,strPanelName,strMenuName)
'2. Fn_ReqMgr_RequirmentSpecBasicCreate(sSpecType,sConfItem,sSpecID,sSpecRevID,sSpecName,sSpecDesc,sSpecUOM)
'3. Fn_ReqMgr_RequirmentBasicCreate(sReqType,sConfItem,sReqID,sReqRevID,sReqName,sReqDesc,sReqUOM)
'4. Fn_ReqMgr_ParagraphBasicCreate(sParaType,sConfItem,sParaID,sParaRevID,sParaName,sParaDesc,sParaUOM)
'5. Fn_OptionsReqMgrSettingsSet(sTraceLinkMode,sApplyTemplates,sReqDrivenDesgnValidation,sBrwseClearText,sQuickCreatePanel,sUserLevel,sKeywordsToImport)
'6.	Fn_ReqMgr_RMTable_RowIndex(StrNodeName)	 
'7.	Fn_ReqMgr_RMTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
'8. Fn_TraceabilityReportOperations(sAction,sNodeName,sNewName,sColName,sCellValue)
'9. Fn_ReqMgr_DataPanelTraceLinkOpeartions(sAction,sNodeName,sNewName,sColName,sColValue)
'10.Fn_ReqMgr_ErrorMessageVerify(sDilogName,sErrorMessage)
'11.Fn_ReqMgr_ImportReqSpec(sFileName,sSpecType,sDescription,sOption,sKeywords,sSubType)
'12.Fn_ReqMgr_MRUListOperations(strAction,strButtonName)
'13.Fn_ReqMgr_RequirmentSpecDetailsCreate(sSpecType,sConfItem,sSpecID,sSpecRevID,sSpecName,sSpecDesc,sSpecUOM,aReqSpecInfo,aReqSpecRevInfo,sProjectNames,aDefineOpt)
'14.Fn_ReqMgr_RequirmentDetailsCreate(sReqType,sConfItem,sReqID,sReqRevID,sReqName,sReqDesc,sReqUOM,aAddReqInfo,aAddReqRevInfo,sAttachAction,aWorkflowInfo,sProjectNames,aDefineOpt)
'15.Fn_ReqMgr_ParagraphDetailsCreate(sParaType,sConfItem,sParaID,sParaRevID,sParaName,sParaDesc,sParaUOM,aAddParaInfo,aAddParaRevInfo,sAttachAction,aWorkflowInfo,sProjectNames,aDefineOpt)
'16.Fn_ReqMgr_DataPanalPropertiesOperations(sAction,sPropertyName,sPropertyValue)
'17.Fn_ReqMgr_LowerRMTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
'18.Fn_TraceabilityReportColumnOperations(strTableName,strAction,sColName,sNewColName,DsplNameOpt)
'19.Fn_ReqMgr_OpenSpecByNameOperations(strAction,strSearchName,strCellValue)
'20.Fn_ReqMgr_StaticTextOperations(strAction,strStaticText)
'21.Fn_ReqMgr_SaveAsItemRevision(sItemID,sItemRevID,sItemName,sItemDesc,sItemUOM)
'22.Fn_ReqMgr_QuickPanelOperation(sName,sType,sChildOpt)
'23.Fn_ReqMgr_CustomizeIWantTo(strAction,strEntryNode)
'24.Fn_ReqMgr_BOMCompare(strMode,strReportOpt)
'25.Fn_ReqMgr_ReqMessageVerify(sDialogName,sErrorMessage)
'26.Fn_ReqMgr_AttachmentTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
'27.Fn_ReqMgr_CustomeNoteBasicCreate(sNoteType,sConfItem,sNoteID,sNoteRevID,sNoteName,sNoteDesc,sNoteUOM)
'28.Fn_ReqMgr_AllocationsTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
'29.Fn_ReqMgr_AllocationMapBasicCreate(sMapType,sConfItem,sMapID,sMapRevID,sMapName,sMapDesc,sMapUOM)
'30.Fn_ReqMgr_CreateAllocation(strName,strReason,strType)
'31.Fn_ReqMgr_MSWordTabOperations(strAction,strValue,strParameterName)
'32.Fn_ReqMgr_ParamatricValueOperation(strAction,strValue,strNoteText)
'33.Fn_ReqMgr_DetailTableOperation(sAction, sObjectName, sColumnName, sExpectedValue,sPopUpMenu)
'34.Fn_ReqMgr_ApplyColumnConfiguration(strAction,strConfigName,arrAvailableProp,bShowIntPropName,strConfigDesc)
'35.Fn_ReqMgr_DetailTableSort(strSortBy,strThenBy1,strThenBy2)
'36.Fn_ReqMgr_CollaborationTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
'37.Fn_ReqMgr_DetailsTableFilterManagement(strAction,strConditionName,strColName,strOperator,strColValue,strLogicalType)
'38.Fn_ReqMgr_MSWordTabOperationsExt(strAction,strValue,strParameterName)
'39.Fn_ReqMgr_SaveColumnConfiguration()
'*********************************************************	Function List		***********************************************************************


'*********************************************************		Function for Tabs in Requirement Manager ***********************************************************************

'Function Name		:			Fn_ReqMgr_RMTabPanelOperation

'Description		:			This function is used to Activate, Verify Activate, PopupMenuSelect(Close Panel, split panel) for RMTabs.

'Parameters			:			1.	strAction:
'								2.	strPanelName:
'								3.	strMenuName:"Close Panel" or "Split Panel"
											
'Return Value		:			True/False

'Pre-requisite		:			Requirement Manager window should be displayed .

'Examples			:			
		'Call Fn_ReqMgr_RMTabPanelOperation("Activate","(000772-MyTc002)","")
		'Call Fn_ReqMgr_RMTabPanelOperation("VerifyActivate","(REQ-000023-MyTc002)","")
		'Call Fn_ReqMgr_RMTabPanelOperation("PopupMenuSelect","(REQ-000023-MyTc002)","Split Panel")

'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Tushar B				31-May-2010			1.0											Sandeep N
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_RMTabPanelOperation(strAction,strPanelName,strMenuName)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_RMTabPanelOperation"
	on Error Resume Next
	Dim objName
	Fn_ReqMgr_RMTabPanelOperation=False	
	JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName").SetTOProperty "label", strPanelName

	Select Case strAction		
		Case "Activate"			'('"Activate","(000772-MyTc002)","")
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName").Click 5,5,"LEFT"
			Fn_ReqMgr_RMTabPanelOperation=True

		Case "VerifyActivate"	'("VerifyActivate","(REQ-000023-MyTc002)","")
			Set objName = JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName").Object			
			Fn_ReqMgr_RMTabPanelOperation=objName.selected()	
			If Fn_ReqMgr_RMTabPanelOperation="true" then
				Fn_ReqMgr_RMTabPanelOperation=True
			Else
				Fn_ReqMgr_RMTabPanelOperation=False
			End If
			Set objName = Nothing

		Case "PopupMenuSelect"	'("PopupMenuSelect","(REQ-000023-MyTc002)","Split Panel")
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName").Click 5,5,"RIGHT"
			wait 2
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaMenu("PanelMenu").SetTOProperty "label", strMenuName
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaMenu("PanelMenu").Select
			Fn_ReqMgr_RMTabPanelOperation=True
		Case Else

			Fn_ReqMgr_RMTabPanelOperation=False
	End Select

End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'*********************************************************		Function to create basic Item		***********************************************************************
'Function Name		:				Fn_ReqMgr_RequirmentSpecBasicCreate

'Description			 :		 		 Creats an Requirement Specification with basic information

'Parameters			   :	 			1.sSpecType: Type of the item.(Requirement Specification)
'													 2.sConfItem: True or False
'													 2.sSpecID: ID of the Specification it should be unique.
'													3.sSpecRevID:Revision ID of the Specification.
'													4.sSpecName:Name of Specification.
'													5.sSpecDesc: Description of the Specification.
'													6:sSpecUOM: Unit of measure of Specification. ( not handling this part)

'Return Value		   : 				Specification Id  / Revision Id

'Pre-requisite			:		 		should be logged in & present on Requirement Manager perspective

'Examples				:				 Fn_ReqMgr_RequirmentSpecBasicCreate("RequirementSpec","OFF","1213132","A","my","","")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   02/05/2010			              1.0										Created						Tushar B
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ReqMgr_RequirmentSpecBasicCreate(sSpecType,sConfItem,sSpecID,sSpecRevID,sSpecName,sSpecDesc,sSpecUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_RequirmentSpecBasicCreate"
	on Error Resume Next
	Dim sSpecificationId, sRevId
	Dim objDialogNewSpec,objSelectType,objDialog

	If Fn_UI_ObjectExist("Fn_ReqMgr_RequirmentSpecBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirementSpec"))=False Then
         Call Fn_MenuOperation("Select","File:New:Requirements Spec...")
	End If
	
	'Check the existence of "NewRequirementSpec" window
	Set objDialogNewSpec=Fn_UI_ObjectCreate("Fn_ReqMgr_RequirmentSpecBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirementSpec"))
		'Select Item Type
		Call Fn_List_Select("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec,"RequirementSpecType",sSpecType)
		'checked Configuration RequirementSpec or not
		If sConfItem <> "" Then
		 Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec,"Configuration Item",sConfItem)
		End If
		'Click on "Next" button
		 Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec,"Next")
		
		If sSpecID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec,"RequirementSpecID", sSpecID)
		End If
		
		If sSpecRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecBasicCreate",objDialogNewSpec,"RevID", sSpecRevID)
		End If
		
		If  sSpecID = "" or sSpecRevID = "" Then
			'click on assign button
			  Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec, "Assign")
		End If
		
		'Extract Creation data
		sSpecificationId =Fn_Edit_Box_GetValue("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec,"RequirementSpecID")
		sRevId = Fn_Edit_Box_GetValue("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec,"RevID")
		
		'Set RequirementSpec name
		 Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec,"SpecName",sSpecName)
		'Set description
		Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec,"Description",sSpecDesc)
		'Set UOM
			If sSpecUOM <> "" Then
				 Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sSpecUOM
				objDialogNewSpec.JavaButton("UnitOfMeasureDrpDwn").Click
				Set objDialog =objDialogNewSpec.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
		   End If
		
		 wait(2)
			objDialogNewSpec.JavaButton("Finish").WaitProperty "enabled", 1, 20000
        			
			Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec, "Finish") 
			Fn_ReqMgr_RequirmentSpecBasicCreate = sSpecificationId & "-" & sRevId
			Call Fn_ReadyStatusSync(1)

			If Fn_UI_ObjectExist("Fn_ReqMgr_RequirmentSpecBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirementSpec"))=True Then		
					Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecBasicCreate", objDialogNewSpec, "Close")
			End If
		
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Specification of ID [" + CStr(sItemId) + "]")
		Set objDialogNewSpec=Nothing
		Set objSelectType=Nothing
		Set objDialog=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'*********************************************************		Function to create basic Item		***********************************************************************
'Function Name		:				Fn_ReqMgr_RequirmentBasicCreate

'Description			 :		 		 Creats an Requirement  with basic information

'Parameters			   :	 			1.sReqType: Type of the item.(Requirement)
'													 2.sConfItem: True or False
'													 2.sReqID: ID of the Requirement it should be unique.
'													3.sReqRevID:Revision ID of the Requirement.
'													4.sReqName:Name of Requirement.
'													5.sReqDesc: Description of the Requirement.
'													6:sReqUOM: Unit of measure of Requirement. ( not handling this part)

'Return Value		   : 				Specification Id  - Revision Id

'Pre-requisite			:		 		should be logged in & present on Requirement Manager perspective

'Examples				:				 Fn_ReqMgr_RequirmentBasicCreate("Requirement","OFF","1213132","A","my","","")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   02/05/2010			              1.0										Created							Tushar B
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_RequirmentBasicCreate(sReqType,sConfItem,sReqID,sReqRevID,sReqName,sReqDesc,sReqUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_RequirmentBasicCreate"
	on Error Resume Next
	Dim sRequirementID, sRevId
	Dim objDialogNewReq,objSelectType,objDialog
	'Select menu [File -> New -> Requirement...]
		If Fn_UI_ObjectExist("Fn_ReqMgr_RequirmentBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirement"))=False Then
			 Call Fn_MenuOperation("Select","File:New:Requirement...")
		End If
		
	'Check the existence of "NewRequirement" window
		Set objDialogNewReq=Fn_UI_ObjectCreate("Fn_ReqMgr_RequirmentBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirement"))
			'Select Item Type
		Call Fn_List_Select("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq,"RequirementType",sReqType)
		'checked Configuration Requirement or not
		If sConfItem <> "" Then
		 Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq,"Configuration Item",sConfItem)
		End If
		'Click on "Next" button
		 Call Fn_Button_Click("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq,"Next")

		If sReqID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq,"ReqID", sReqID)
		End If
	
		If sReqRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_ReqMgr_RequirmentBasicCreate",objDialogNewReq,"ReqRevID", sReqRevID)
		End If

		If  sReqID = "" or sReqRevID = "" Then
			'click on assign button
			  Call Fn_Button_Click("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq, "Assign")
		End If
	
		'Extract Creation data
		sRequirementID =Fn_Edit_Box_GetValue("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq,"ReqID")
		sRevId = Fn_Edit_Box_GetValue("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq,"ReqRevID")
		
		'Set Requirement name
		 Call Fn_Edit_Box("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq,"ReqName",sReqName)
		'Set description
		Call Fn_Edit_Box("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq,"Description",sReqDesc)
		'Set UOM
		If sReqUOM <> "" Then
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sReqUOM
			objDialogNewReq.JavaButton("UOMDrpDwn").Click
			Set objDialog =objDialogNewReq.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
	   End If

		wait(2)
		objDialogNewReq.JavaButton("Finish").WaitProperty "enabled", 1, 20000
		
		'Click on "Finish" button
	
		Call Fn_Button_Click("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq, "Finish") 
		Fn_ReqMgr_RequirmentBasicCreate = sRequirementID & "-" & sRevId
		Call Fn_ReadyStatusSync(1)

		'Click on Close button
		Call Fn_Button_Click("Fn_ReqMgr_RequirmentBasicCreate", objDialogNewReq, "Close")
						
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Requirement of ID [" + CStr(sItemId) + "]")
	Set objDialogNewReq=Nothing
	Set objSelectType=Nothing
	Set objDialog=Nothing
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'*********************************************************		Function to create basic Item		***********************************************************************
'Function Name		:				Fn_ReqMgr_ParagraphBasicCreate

'Description			 :		 		 Creats an Requirement  with basic information

'Parameters			   :	 			1.sParaType: Type of the item.(Requirement)
'													 2.sConfItem: True or False
'													 2.sParaID: ID of the Requirement it should be unique.
'													3.sParaRevID:Revision ID of the Requirement.
'													4.sParaName:Name of Requirement.
'													5.sParaDesc: Description of the Requirement.
'													6:sParaUOM: Unit of measure of Requirement. ( not handling this part)

'Return Value		   : 				Specification Id  /  Revision Id

'Pre-requisite			:		 		should be logged in & present on Requirement Manager perspective

'Examples				:				 Fn_ReqMgr_ParagraphBasicCreate("Paragraph","OFF","1213132","A","my","","")

'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   02/05/2010			              1.0										Created						Tushar
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_ParagraphBasicCreate(sParaType,sConfItem,sParaID,sParaRevID,sParaName,sParaDesc,sParaUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_ParagraphBasicCreate"
	on Error Resume Next
	Dim sParagraphID, sRevId
	Dim objDialogNewPara,objSelectType,objDialog
	'Select menu [File -> New -> Paragraph...]
			If Fn_UI_ObjectExist("Fn_ReqMgr_ParagraphBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewParagraph"))=False Then
				 Call Fn_MenuOperation("Select","File:New:Paragraph...")
			End If
	
			'Check the existence of "NewParagraph" window
			Set objDialogNewPara=Fn_UI_ObjectCreate("Fn_ReqMgr_ParagraphBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewParagraph"))
			'Select Paragraph Type
            Call Fn_List_Select("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara,"ParagraphType",sParaType)
			'checked Configuration Paragraph or not
			If sConfItem <> "" Then
             Call Fn_CheckBox_Set("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara,"Configuration Item",sConfItem)
			End If
			'Click on "Next" button
             Call Fn_Button_Click("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara,"Next")

			If sParaID <> "" Then
				'Set  Item Id
                 Call Fn_Edit_Box("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara,"ParaID", sParaID)
			End If
		
			If sParaRevID <> "" Then
				'Set Revision ID
                Call Fn_Edit_Box("Fn_ReqMgr_ParagraphBasicCreate",objDialogNewPara,"ParaRevID", sParaRevID)
			End If
	
			If  sParaID = "" or sParaRevID = "" Then
				'click on assign button
                  Call Fn_Button_Click("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara, "Assign")
			End If
		
			'Extract Creation data
			sParagraphID =Fn_Edit_Box_GetValue("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara,"ParaID")
            sRevId = Fn_Edit_Box_GetValue("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara,"ParaRevID")
			
			'Set Paragraph name
             Call Fn_Edit_Box("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara,"ParaName",sParaName)
			'Set description
            Call Fn_Edit_Box("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara,"Description",sParaDesc)
			'Set UOM
			If sParaUOM <> "" Then
				 Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sParaUOM
				objDialogNewPara.JavaButton("UOMDrpDwn").Click
				Set objDialog =objDialogNewPara.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
		   End If

			wait(2)
			objDialogNewPara.JavaButton("Finish").WaitProperty "enabled", 1, 20000
			
			'Click on "Finish" button
			
                Call Fn_Button_Click("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara, "Finish") 
				Fn_ReqMgr_ParagraphBasicCreate = sParagraphID & "-" & sRevId
				Call Fn_ReadyStatusSync(1)

                'Click on Close button
				Call Fn_Button_Click("Fn_ReqMgr_ParagraphBasicCreate", objDialogNewPara, "Close")
							
            Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Paragraph of ID [" + CStr(sItemId) + "]")
		Set objDialogNewPara=Nothing
		Set objSelectType=Nothing
		Set objDialog=Nothing
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'*********************************************************		Function to create basic Item		***********************************************************************
'Function Name		:				Fn_OptionsReqMgrSettingsSet

'Description			 :		 		 RequirementManager setting from Edit->Options->Requirement Manager

'Parameters			   :	 			1.sTraceLinkMode: "ON"  or "OFF"
'													 2.sApplyTemplates: "ON"  or "OFF"
'													 3.sReqDrivenDesgnValidation: "ON"  or "OFF"
'													4.sBrwseClearText:"ON"  or "OFF"
'													5.sQuickCreatePanel:"ON"  or "OFF"
'													6.sUserLevel: "Basic" or "Intermediate"  or "Advanced"
'													7:sKeywordsToImport: Text (Keyword to Import)

'Return Value		   : 				True or False

'Pre-requisite			:		 		should be logged in  into Teamcenter

'Examples				:				 Fn_OptionsReqMgrSettingsSet("On","OFF","On","On","On","Basic" ,"Any Text")
'													Fn_OptionsReqMgrSettingsSet("On","On","Off","Off","On","Intermediate" ,"Test Text")
'													Fn_OptionsReqMgrSettingsSet("","On","","Off","On","Advanced" ,"Test Text")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   02/05/2010			              1.0										Created						Tushar B
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_OptionsReqMgrSettingsSet(sTraceLinkMode,sApplyTemplates,sReqDrivenDesgnValidation,sBrwseClearText,sQuickCreatePanel,sUserLevel,sKeywordsToImport)
GBL_FAILED_FUNCTION_NAME="Fn_OptionsReqMgrSettingsSet"
on Error Resume Next
Dim objDialogOption, strSource, strVerify
	'Checking Option Window is exist on Not
	If Fn_UI_ObjectExist("Fn_OptionsReqMgrSettingsSet",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options"))=False Then
		'Invoking Option Winidow
		Call Fn_MenuOperation("Select", "Edit:Options...")
	End If
		'Setting object of Option Dialog
		Set objDialogOption=Fn_UI_ObjectCreate("Fn_OptionsReqMgrSettingsSet",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Options"))
		'Clicking on "Options" Static Text
        Call Fn_UI_JavaStaticText_Click("Fn_OptionsReqMgrSettingsSet", objDialogOption, "Options",0,0,"LEFT")

		'Selecting  "Options:Requirements Management" From "OptionsTree" Tree
		Call Fn_JavaTree_Select("Fn_OptionsReqMgrSettingsSet",objDialogOption, "OptionsTree","Options:Requirements Management")
		'Checking  sTraceLinkMode parameter is blank or not
		If sTraceLinkMode<>"" Then
			'Setting values to TraceLinkMode Check Box
			Call Fn_CheckBox_Set("Fn_OptionsReqMgrSettingsSet", objDialogOption, "TraceLinkMode", sTraceLinkMode)
		End If
		'Checking  sApplyTemplates parameter is blank or not
		If sApplyTemplates<>"" Then
			'Setting values to ApplyTemplates Check Box
			Call Fn_CheckBox_Set("Fn_OptionsReqMgrSettingsSet", objDialogOption, "ApplyTemplates", sApplyTemplates)
		End If
		'Checking  sReqDrivenDesgnValidation parameter is blank or not
		If sReqDrivenDesgnValidation<>"" Then
			'Setting values to ReqDrivenDesgnValidation Check Box
			Call Fn_CheckBox_Set("Fn_OptionsReqMgrSettingsSet", objDialogOption, "ReqDrivenDesgnValidation", sReqDrivenDesgnValidation)
		End If
		'Checking  sBrwseClearText parameter is blank or not
		If sBrwseClearText<>"" Then
			'Setting values to BrwseClearText Check Box
			Call Fn_CheckBox_Set("Fn_OptionsReqMgrSettingsSet", objDialogOption, "BrwseClearText", sBrwseClearText)
		End If
		'Checking  sQuickCreatePanel parameter is blank or not
		If sQuickCreatePanel<>"" Then
			'Setting values to QuickCreatePanel Check Box
			Call Fn_CheckBox_Set("Fn_OptionsReqMgrSettingsSet", objDialogOption, "QuickCreatePanel", sQuickCreatePanel)
		End If
		'Checking  sUserLevel parameter is blank or not
		If sUserLevel<>"" Then
			'Setting values to UserLevel Radio Button
			objDialogOption.JavaRadioButton("UserLevel").SetTOProperty "attached text",sUserLevel
			objDialogOption.JavaRadioButton("UserLevel").Set "ON"
		End If
		'Checking  sKeywordsToImport  parameter is blank or not
'		If sKeywordsToImport<>"" Then
'			'Setting values to KeywordsToImport  Edit Box
'			Call Fn_Edit_Box(sFunctionName,objDialogOption,"KeywordsToImport",sKeywordsToImport)
'		End If
'Added by Tushar for verify case since found 1 testcases only
		If sKeywordsToImport<>"" Then
			'To verify Case
			'the value send ~ in the start
			If inStr(1,sKeywordsToImport,"~")>0 Then
				strSource=objDialogOption.JavaEdit("KeywordsToImport").GetROProperty("value")
				strVerify=split(sKeywordsToImport,"~")(1)
				 If  inStr(1,strSource,strVerify)>0 then
				 Else
					'Clicking on "OK" To save
					Call Fn_Button_Click("Fn_OptionsReqMgrSettingsSet",objDialogOption, "OK")

 					Fn_OptionsReqMgrSettingsSet=False
					Set objDialogOption=Nothing
					Exit Function

				 End if
								
			Else
				'Setting values to KeywordsToImport  Edit Box
				Call Fn_Edit_Box("Fn_OptionsReqMgrSettingsSet",objDialogOption,"KeywordsToImport",sKeywordsToImport)
			End If
		End If


		'Clicking on "OK" To save
		Call Fn_Button_Click("Fn_OptionsReqMgrSettingsSet",objDialogOption, "OK")
		Fn_OptionsReqMgrSettingsSet=True
		
Set objDialogOption=Nothing
End Function 



'*********************************************************		Function to Get ReqMgr Table Node Index in Requirement Manager		***********************************************************************

'Function Name		:				Fn_ReqMgr_RMTable_RowIndex

'Description			 :		 		This function is used to get the ReqMgr Table Node Index.

'Parameters			   :	 			1.  StrNodeName:Name of the Node to retrieve Index for.
											
'Return Value		   : 				Node index

'Pre-requisite			:				Requirement Manager window should be displayed .

'Examples				:				Fn_ReqMgr_RMTable_RowIndex(" 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3")


'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Tushar						2-June-2010		1.0											Sandeep
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ReqMgr_RMTable_RowIndex(StrNodeName)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_RMTable_RowIndex"
	On Error Resume Next
	Dim IntRows ,StrNodePath, IntCounter, ObjTable, StrIndex

	Fn_ReqMgr_RMTable_RowIndex="FAIL"
	'Verify that ReqMgr Table is displayed
	If JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").Exist(iTimeOut) Then

		'Get the No. of rows present in the ReqMgr Table
		IntRows = JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").GetROProperty("rows")
		Set ObjTable = JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").Object

		'Format the Inout as per Table Default Nodes
		StrNodeName = Replace(StrNodeName, ":", ", ")

		'Get the Row No. of required Node
		For IntCounter = 0 to IntRows -1
			StrNodePath = ObjTable.getPathForRow(IntCounter).toString
			StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
			StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
			'msgbox StrNodeName
			    
			' objRMTable.setSelectionPaths(strNodeName+",").tostring()
			'msgbox objRMTable.getSelectedRow
	
			If Trim(StrNodePath) = Trim(StrNodeName) Then
				'Reporter.ReportEvent  micPass, "ReqMgrTable", "Row Index of [" + StrNodeName +"] Node is [" + IntCounter + "]"
				'Call Fn_WriteLogFile("Fn_ReqMgr_RMTable_RowIndex", 3, Err.Number,"PASS: Row Index of [" + StrNodeName +"] Node is [" + IntCounter + "]")
				StrIndex = Cstr(IntCounter)
				Fn_ReqMgr_RMTable_RowIndex = StrIndex
				Exit For
			End If
		Next
		If IntCounter = IntRows Then
			'Call Fn_WriteLogFile("Fn_ReqMgr_RMTable_RowIndex", 1, Err.Number,"FAIL: Failed to Get Row Index of Node [" + StrNodeName +"]")
			Fn_ReqMgr_RMTable_RowIndex = "FAIL:Node Not Found"
		End If

		'Release the Table object
	   set ObjTable = Nothing

	Else
		'RMTable not displayed in Requirement Manager!
		Fn_ReqMgr_RMTable_RowIndex="FAIL"
	End If
End Function




'*********************************************************		Function to Get ReqMgr Table Node operation in Requirement Manager		***********************************************************************

'Function Name		:				Fn_ReqMgr_RMTableNodeOpeations

'Description			 :		 		This function is used to get the ReqMgr Table Node operation.

'Parameters			   :				1.	strAction = "Select"
'										2. StrNodeName:Name of the Node. 
'										3. strColName
'										4. strColValue
'										5. strPopupMenu
				
											
'Return Value		   : 				True/ False/TableNodeIndex

'Pre-requisite			:				Requirement Manager window should be displayed .


'Examples				:		Fn_ReqMgr_RMTableNodeOpeations("Select"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
'								Fn_ReqMgr_RMTableNodeOpeations("VerifyNode"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")
'								Fn_ReqMgr_RMTableNodeOpeations("Expand"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View)","","","")
'								Fn_ReqMgr_RMTableNodeOpeations("getNodeIndex"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")
'								Fn_ReqMgr_RMTableNodeOpeations("PopupMenuSelect","","","","Access...")
'								Fn_ReqMgr_RMTableNodeOpeations("PopupMenuSelect","","","","Trace Link:Start Trace Link")
'								Fn_ReqMgr_RMTableNodeOpeations("MultiSelect","REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View):000575/A;1-P23~REQ-000049/A;1-Req1 (View):REQ-000148/A;1-Req2","","","")
'								Fn_ReqMgr_RMTableNodeOpeations("VerifyColValue"," REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View)","Item Type","Requirement","")
'								Fn_ReqMgr_RMTableNodeOpeations("EditColValue"," REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View)","Find No.1","50","")
''Fn_ReqMgr_RMTableNodeOpeations(Select, VerifyNode,VerifyColValue, Expand, getNodeIndex, rightClick/PopupMenuSelect, Multiselect,EditColValue) 

								'Fn_ReqMgr_RMTableNodeOpeations("GetCellData",1,0,"","")
								'IMP Note related to "GetCellData" Case ONLY
								'(IMP NOTE)' "strNodeName" - This parameter is use as Row number in this Case--->pass Integer Value's(eg:0 or 1 or 2 or 3. . . . . .. . )
								'strColName - This parameter is use as column nuber in this case---->pass Integer Value's(eg:0 or 1 or 2 or 3. . . . . .. . ) or Column Name
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Tushar						2-June-2010		1.0																			Sandeep								
'										Sandeep					8-June-2010		1.0					Add  "GetCellData"  Case	  Archana
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Public Function Fn_ReqMgr_RMTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_RMTableNodeOpeations"
	on Error Resume Next
    Dim iRowNo, sMenu, iNodeNo, iColNo, iStart,strName
	Fn_ReqMgr_RMTableNodeOpeations=False
	'Verify ReqMgr Table
	If JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").Exist(iTimeOut) then    	
		Select Case StrAction

			Case "Select"		'("Select"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","","","")
				iRowNo = Fn_ReqMgr_RMTable_RowIndex(strNodeName)
				If isNumeric(iRowNo) Then
					JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SelectRow iRowNo
					Fn_ReqMgr_RMTableNodeOpeations=True
				End if

			Case "VerifyNode"		'("Verify"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")		
        			'Verify Node Exist
					iRowNo = Fn_ReqMgr_RMTable_RowIndex(strNodeName)
					If isNumeric(iRowNo) then
						Fn_ReqMgr_RMTableNodeOpeations=True
					End if
        
			Case "getNodeIndex"	'("getNodeIndex"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View)","","","")
				iRowNo = Fn_ReqMgr_RMTable_RowIndex(strNodeName)
				If isNumeric(iRowNo) then
					Fn_ReqMgr_RMTableNodeOpeations=iRowNo
				End if

			Case "Expand"	'("Expand"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View)","","","")
				iRowNo = Fn_ReqMgr_RMTable_RowIndex(strNodeName)
				If isNumeric(iRowNo) then
					JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SelectRow iRowNo
					call Fn_menuOperation("Select","View:Expand")

					Fn_ReqMgr_RMTableNodeOpeations=True
				End if

			Case "PopupMenuSelect"	'("PopupMenuSelect","","","","Trace Link:Start Trace Link")
				'Pre-requisite = Row should be selected
				strPopupMenu=Replace(strPopupMenu,":",";")
				iRowNo=JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").Object.getSelectedRow()
				If isNumeric(iRowNo) then
					'JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SelectRow iRowNo
					JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").ClickCell iRowNo,"Item Structure","RIGHT" 
					wait 1
					sMenu = JavaWindow("RequirementsManager").WinMenu("ContextMenu").BuildMenuPath(strPopupMenu)
					wait 1
					JavaWindow("RequirementsManager").WinMenu("ContextMenu").Select sMenu
					Fn_ReqMgr_RMTableNodeOpeations=True
				End if

			Case "MultiSelect"		'("MultiSelect","REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View):000575/A;1-P23~REQ-000049/A;1-Req1 (View):REQ-000148/A;1-Req2","","","")

				strNodeName=split(strNodeName,"~") 
				For iNodeNo=0 to Ubound(strNodeName)
					iRowNo = Fn_ReqMgr_RMTable_RowIndex(strNodeName(iNodeNo))
					If isNumeric(iRowNo) Then
						If iNodeNo=0 Then
							JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SelectRow iRowNo
							Fn_ReqMgr_RMTableNodeOpeations=True
						Else
							JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").ExtendRow "#"&iRowNo
							Fn_ReqMgr_RMTableNodeOpeations=True
						End If					
					End if
				Next
				
			Case "VerifyColValue"	'("VerifyColValue"," REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View)","Item Type","Requirement","")
				iRowNo = Fn_ReqMgr_RMTable_RowIndex(strNodeName)
            	If isNumeric(iRowNo) then
					'Get column Rows
					iColNo=JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").GetROProperty("cols")

					For iStart=0 to iColNo-1
						''Verify the Column name is similar to required column name
						If JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").GetColumnName(iStart)=strColName Then
							'Verify the Column value is similar to required value
							If JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").GetCellData(iRowNo,iStart)=strColValue then
								Fn_ReqMgr_RMTableNodeOpeations=True
							End if
							Exit For
						End If
					Next
				Else
					Fn_ReqMgr_RMTableNodeOpeations=False
				End if

			Case "EditColValue"		'("EditColValue"," REQ-000049/A;1-Req1 (View):000455/A;1-P1 (View)","Find No.1","50","")	

				iRowNo = Fn_ReqMgr_RMTable_RowIndex(strNodeName)
            	If isNumeric(iRowNo) then

					
					'Get column Rows
					iColNo=JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").GetROProperty("cols")

					For iStart=0 to iColNo-1
						''Verify the Column name is similar to required column name
						If JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").GetColumnName(iStart)=strColName Then
							'Verify the Column value is similar to required value
							JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SetCellData iRowNo,iStart,strColValue
							Fn_ReqMgr_RMTableNodeOpeations=True
							Exit For
						End If
					Next
				Else
					Fn_ReqMgr_RMTableNodeOpeations=False
				End if

				Case "GetCellData" '("GetCellData",1,0,"","")
					
					
					'(IMP NOTE)' "strNodeName" - This parameter is use as Row number in this Case
					'strColName - This parameter is use as column nuber in this case
					
						JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SelectRow strNodeName
						strName=JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").GetCellData(strNodeName,strColName)
						
					If Err.number < 0 Then
                		Fn_ReqMgr_RMTableNodeOpeations=False
					Else
						Fn_ReqMgr_RMTableNodeOpeations = MId(strName,instr(1,strName,":")+1 , Len(strName))
					End If

			Case "PopupMenuExist"		
						strPopupMenu=Replace(strPopupMenu,":",";")
						iRowNo=JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").Object.getSelectedRow()
						If isNumeric(iRowNo) then
							JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").ClickCell iRowNo,"Item Structure","RIGHT" 
							wait 1
							sMenu = JavaWindow("RequirementsManager").WinMenu("ContextMenu").BuildMenuPath(strPopupMenu)
							If JavaWindow("RequirementsManager").WinMenu("ContextMenu").GetItemProperty (sMenu,"Exists") = True Then
									Fn_ReqMgr_RMTableNodeOpeations = TRUE
							Else
									Fn_ReqMgr_RMTableNodeOpeations = FALSE
							End If
						End If
		Case "DoubleClickCell"		'("DoubleClickCell"," 000452/A;1-ReqSpec (View):REQ-000047/A;1-Req2 (View):000454/A;1-P2 (View):REQ-000048/A;1-Req3","Has Attached Notes","","")
				iRowNo = Fn_ReqMgr_RMTable_RowIndex(strNodeName)
				If isNumeric(iRowNo) Then
					JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").SelectRow iRowNo
					Fn_ReqMgr_RMTableNodeOpeations=True
				End if
					JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("RMTable").DoubleClickCell iRowNo,strColName
		End Select
	Else
		'RMTable not displayed in Requirement Manager!
		Fn_ReqMgr_RMTableNodeOpeations=False
	End if

End Function



'*********************************************************		Function to Report Operation	***********************************************************************
'Function Name		:				Fn_TraceabilityReportOperations

'Description			 :		 		 Perform Operations on "Traceability Report" Dialog

'Parameters			   :	 			1.sAction: DefiningTable:Properties (First is Table Name Compulsory)
'													 2.sNodeName: Node on which we have to perform operation
'													 3.sNewName:New Name in Property
'													4.sColName:Column Name
'													5.sCellValue:Cell Value  												

'Return Value		   : 				True or False

'Pre-requisite			:		 	Must be Selected Trace Link Node

'Examples				:			'Fn_TraceabilityReportOperations("DefiningTable:Properties","000494-Test3:Test4->Test3","NewName3","","") -->This will change the Name property  and press OK on TraceabilityReport
'												'Fn_TraceabilityReportOperations("DefiningTable:Expand","000494-Test3:Test4->Test3","","","") --->This Case Expand the tree node but it will not  press OK on TraceabilityReport
'												'Fn_TraceabilityReportOperations("DefiningTable:Select","000494-Test3:Test4->Test3:000495-Test4","","","") --->This Case Select the tree node but it will not  press OK on TraceabilityReport
'												'Fn_TraceabilityReportOperations("DefiningTable:Verify","000494-Test3:Test4->Test3:000495-Test4","","","") ---->This Case Verify the tree node but it will not  press OK on TraceabilityReport
'												'Fn_TraceabilityReportOperations("DefiningTable:Properties","000494-Test3:Test4->Test3","Change Name","","") 
'												'Fn_TraceabilityReportOperations("DefiningTable:Go To Object","Change Name","","","") 
'												'Fn_TraceabilityReportOperations("DefiningTable:Delete Trace Link","Change Name","","","") 
'												'FN_TraceabilityReportOperations("DefiningTable:CellVerify","","","Relation Type","Trace Link") 
'												'Fn_TraceabilityReportOperations("DefiningTable:Refresh Report","","","","")--->This will refresh the TraceabilityReport  window and close the TraceabilityReport
'												'Fn_TraceabilityReportOperations("DefiningTable:DescriptionProperties","000494-Test3:Test4->Test3","New Description","","") 
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   03/05/2010			              1.0										Created						Tushar B
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_TraceabilityReportOperations(sAction,sNodeName,sNewName,sColName,sCellValue)
GBL_FAILED_FUNCTION_NAME="Fn_TraceabilityReportOperations"
on Error Resume Next
'Declaring All Varaibles
Dim aAction,sTableName,iRows,iCounter,sNodePath,sIndex,bFlag, ArrLists, iToolCnt,  sContents,sCellData
'Declaring All Object's
Dim objJavaDialogReport,ObjDesc
'Spliting sAction To retriewe Table name
aAction=Split(sAction,":")
sTableName=aAction(0)
'Setting bFlag
bFlag=False

'Creating Object "Traceability Report" Dialog
Set objJavaDialogReport=Fn_UI_ObjectCreate("Fn_TraceabilityReportOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report"))
Set ObjDesc = Description.Create() 
	ObjDesc("to_class").Value = "JavaToolbar" 
	ObjDesc("enabled").Value = 1
	
    	'Get the total of Toolbar objects
		Set ArrLists =objJavaDialogReport.ChildObjects(ObjDesc)
			iToolCnt = objJavaDialogReport.ChildObjects(ObjDesc).count
'Checking "Show Trace Link" button present or not
		For iCounter = 0 to iToolCnt-1
			sContents = ArrLists(iCounter).GetContent()
			
			If instr(sContents, "Show Trace Link") > 0 Then
				'Clicking "Show Trace Link" button
				ArrLists(iCounter).Press "Show Trace Link"
				bFlag=True
                Exit For
			End If
		Next
If sNodeName<>"" Then
	'Identifying Table
	Select Case sTableName
		   Case "ComplyingTable"
				'Checking Existance of "ComplyingTable" Table
				If Fn_UI_ObjectExist("Fn_TraceabilityReportOperations",objJavaDialogReport.JavaTable("ComplyingTable"))=True Then
					'Retriwing No of rows
					iRows = Fn_Table_GetRowCount("Fn_TraceabilityReportOperations",objJavaDialogReport,"ComplyingTable")
		
	'					'objJavaDialogReport.JavaTable("ComplyingTable").SelectRow 0
						For iCounter = 0 to iRows -1
							objJavaDialogReport.JavaTable("ComplyingTable").SelectRow iCounter
							sNodePath=objJavaDialogReport.JavaTable("ComplyingTable").GetCellData(iCounter,0)
								'Checking "sNodeName" present in table or not
								If Trim(sNodePath) = Trim(sNodeName) Then
										sIndex = Cstr(iCounter)
										bFlag=True
										Exit For
								End If
						Next
								If iCounter = iRows Then
									Fn_TraceabilityReportOperations =False
									Exit Function
								End If
						End If
			 
				Case "DefiningTable"
						'Checking Existance of "DefiningTable" Table
					If Fn_UI_ObjectExist("Fn_TraceabilityReportOperations",objJavaDialogReport.JavaTable("DefiningTable"))=True Then
						iRows = Fn_Table_GetRowCount("Fn_TraceabilityReportOperations",objJavaDialogReport,"DefiningTable")
	'					'objJavaDialogReport.JavaTable("ComplyingTable").SelectRow 0
						For iCounter = 0 to iRows -1
							objJavaDialogReport.JavaTable("DefiningTable").SelectRow iCounter
							sNodePath=objJavaDialogReport.JavaTable("DefiningTable").GetCellData(iCounter,0)
							'Checking "sNodeName" present in table or not
								If Trim(sNodePath) = Trim(sNodeName) Then
										sIndex = Cstr(iCounter)
										bFlag=True
										Exit For
									End If
						Next
								If iCounter = iRows Then
									Fn_TraceabilityReportOperations =False
									Exit Function
								End If
					End if
				End Select 
End If
Select Case aAction(1)
	'To Expand Tree Node of table
	Case "Expand"
		If sTableName="ComplyingTable" Then
			objJavaDialogReport.JavaTable("ComplyingTable").SelectRow sIndex
            objJavaDialogReport.JavaTable("ComplyingTable").DoubleClickCell sIndex,0
			Fn_TraceabilityReportOperations =True
			Exit Function
		Else
			objJavaDialogReport.JavaTable("DefiningTable").SelectRow sIndex
            objJavaDialogReport.JavaTable("DefiningTable").DoubleClickCell sIndex,0
			Fn_TraceabilityReportOperations =True
			Exit Function
		End If
	'To Select Tree Node of table		
	Case "Select"
		If sTableName="ComplyingTable" Then
        	objJavaDialogReport.JavaTable("ComplyingTable").SelectRow sIndex
            Fn_TraceabilityReportOperations =True
			Exit Function
		Else
			objJavaDialogReport.JavaTable("DefiningTable").SelectRow sIndex
            Fn_TraceabilityReportOperations =True
			Exit Function
		End If
	'Verifying Node is present or not	
	Case "Verify"
            If bFlag=True Then
				Fn_TraceabilityReportOperations =True
				Exit Function
			Else
				Fn_TraceabilityReportOperations =False
				Exit Function
			End If
	'To modify properties	
	Case "Properties"
			If Fn_UI_ObjectExist("Fn_TraceabilityReportOperations",objJavaDialogReport)=True Then
				'Pressing "Properties" button to change "Name" property
				For iCounter = 0 to iToolCnt-1
					sContents = ArrLists(iCounter).GetContent()
						If instr(sContents, "Properties") > 0 Then
							ArrLists(iCounter).Press "Properties"
							wait(5)
							'Changing the "Name"
				 			Call Fn_Edit_Box("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"),"Name",sNewName)
							Call Fn_Button_Click("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"),"OK")
							Fn_TraceabilityReportOperations = True
                			Exit For
						End If
				Next
	
				If iCounter = iToolCnt Then
	    			Fn_TraceabilityReportOperations = FALSE
				End If
			Else
				Fn_TraceabilityReportOperations = FALSE
			End If
			'Refrefing the report
   			For iCounter = 0 to iToolCnt-1
				sContents = ArrLists(iCounter).GetContent()
					If instr(sContents, "Refresh Report") > 0 Then
						ArrLists(iCounter).Press "Refresh Report"
						bFlag=True
                		Exit For
					End If
			Next
	
	'To modify Description property
	Case "DescriptionProperties"
			If Fn_UI_ObjectExist("Fn_TraceabilityReportOperations",objJavaDialogReport)=True Then
				'Pressing "Properties" button to change "Name" property
				For iCounter = 0 to iToolCnt-1
					sContents = ArrLists(iCounter).GetContent()
						If instr(sContents, "Properties") > 0 Then
							ArrLists(iCounter).Press "Properties"
							wait(5)
							JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").JavaStaticText("EmptyProperties").SetTOProperty "label","Hide empty properties..."
							wait(5)
							If Fn_UI_ObjectExist("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").JavaStaticText("EmptyProperties"))=False Then
								JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").JavaStaticText("EmptyProperties").SetTOProperty "label","Show empty properties..."
								JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties").JavaStaticText("EmptyProperties").Click 1,1
								wait(5)
							End If

							'Changing the "Name"
				 			Call Fn_Edit_Box("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"),"Description",sNewName)
							Call Fn_Button_Click("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Properties"),"OK")
							Fn_TraceabilityReportOperations = True
                			Exit For
						End If
				Next
	
				If iCounter = iToolCnt Then
	    			Fn_TraceabilityReportOperations = FALSE
				End If
			Else
				Fn_TraceabilityReportOperations = FALSE
			End If
			'Refrefing the report
   			For iCounter = 0 to iToolCnt-1
				sContents = ArrLists(iCounter).GetContent()
					If instr(sContents, "Refresh Report") > 0 Then
						ArrLists(iCounter).Press "Refresh Report"
						bFlag=True
                		Exit For
					End If
			Next

	'To Delete Trace Link		
	Case "Delete Trace Link"
		'Pressing "Delete Trace Link" button to Delete Trace Link
		For iCounter = 0 to iToolCnt-1
			sContents = ArrLists(iCounter).GetContent()
            If instr(sContents, "Delete Trace Link") > 0 Then
					ArrLists(iCounter).Press "Delete Trace Link"
					bFlag=True
					wait(2)
					If Fn_UI_ObjectExist("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Refresh Window")) Then
						Call Fn_Button_Click("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Refresh Window"),"Yes")
                   End If
                Exit For
			End If
		Next
			If  bFlag=True Then
				Fn_TraceabilityReportOperations=True
			Else
				Fn_TraceabilityReportOperations=False
			End If
	'To go Object		
	Case "Go To Object"
		'Pressing "Go To Object" button to Delete Trace Link
		For iCounter = 0 to iToolCnt-1
			sContents = ArrLists(iCounter).GetContent()
            	If instr(sContents, "Go To Object") > 0 Then
					ArrLists(iCounter).Press "Go To Object"
					bFlag=True
                	Exit For
				End If
		Next
			If  bFlag=True Then
				Fn_TraceabilityReportOperations=True
			Else
				Fn_TraceabilityReportOperations=False
			End If
	  Case "CellVerify"
		 If sTableName="ComplyingTable" Then
				iRows=Fn_Table_GetRowCount("Fn_TraceabilityReportOperations",objJavaDialogReport,sTableName)
						For iCounter=0 to iRows-1
                            sCellData=objJavaDialogReport.JavaTable("ComplyingTable").GetCellData(iCounter,sColName)
								If sCellData=sCellValue Then
									bReturn=True
									Exit For
							   End If
						Next
		  Else
				iRows=Fn_Table_GetRowCount("Fn_TraceabilityReportOperations",objJavaDialogReport,sTableName)
					For iCounter=0 to iRows-1
                           sCellData=objJavaDialogReport.JavaTable("DefiningTable").GetCellData(iCounter,sColName)
							If sCellData=sCellValue Then
									bReturn=True
									Exit For
						   End If
					Next
		  End If
			
			If bReturn=True Then
				Fn_TraceabilityReportOperations=True
			Else
				Fn_TraceabilityReportOperations=False
			End If
		Case "Refresh Report"
			'Refrefing the report
   			For iCounter = 0 to iToolCnt-1
				sContents = ArrLists(iCounter).GetContent()
					If instr(sContents, "Refresh Report") > 0 Then
						ArrLists(iCounter).Press "Refresh Report"
						bReturn=True
                  		Exit For
					End If
			Next
			If bReturn=True Then
				Fn_TraceabilityReportOperations=True
			Else
				Fn_TraceabilityReportOperations=False
			End If
End Select
	'Clicking on "OK" Button of "Report" Dialog
	Call Fn_Button_Click("Fn_TraceabilityReportOperations",objJavaDialogReport,"Ok")
	
Set ObjDesc = Nothing
Set ArrLists = Nothing
Set objJavaDialogReport=Nothing
End Function



'
''*********************************************************		Function to Perform ReqMgr Panel operation in Requirement Manager		***********************************************************************
'
''Function Name		:				Fn_ReqMgr_DataPanelTraceLinkOpeartions
'
''Description			 :		 		This function is used to get the ReqMgr Table Node Index.
'
''Parameters			   :				1.	sAction = "Select"
''												2.   sNodeName:Name of the Node. 
''												3. sNewName
''												4. sColName
''												5. sColValue
'			
'											
''Return Value		   : 				True/ False
'
''Pre-requisite			:				Requirement Manager window should be displayed .
'
'
''Examples				:		   Fn_ReqMgr_DataPanelTraceLinkOpeartions("ComplyingTable:Delete Trace Link","REQ-000148/A;1-Req2:Req2->Req3","","","")
'										Fn_ReqMgr_DataPanelTraceLinkOpeartions("ComplyingTable:NodeVerify","REQ-000148/A;1-Req2:Req2->Req3","","","")
'										Fn_ReqMgr_DataPanelTraceLinkOpeartions("ComplyingTable:Select","REQ-000148/A;1-Req2:Req2->Req3","","","")
'										Fn_ReqMgr_DataPanelTraceLinkOpeartions("ComplyingTable:Expand","REQ-000148/A;1-Req2:Req2->Req3","","","")
'										Fn_ReqMgr_DataPanelTraceLinkOpeartions("ComplyingTable:VerifyCellValue","REQ-000148/A;1-Req2:Req2->Req3","","Relation Type","Trace Link")
''History:
''										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
''										Tushar						4-June-2010		1.0															Sandeep
''-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
Public Function Fn_ReqMgr_DataPanelTraceLinkOpeartions(sAction,sNodeName,sNewName,sColName,sColValue)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_DataPanelTraceLinkOpeartions"
	'On error resume next
   Dim bReturn,strTable,strOperation,objTable, iStart, iRowNo, iColNo, sColHeader, bSelectFlag

   Fn_ReqMgr_DataPanelTraceLinkOpeartions=False
   bSelectFlag=False
   JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName").SetTOProperty "label", "Trace Link"

	'Open the datapanel window if not exist
	If Not JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName").Exist(iTimeOut) Then
		call Fn_menuOperation("Select","View:Show Data Panel")
		wait 30
	End If

	'Verify the TraceLink tab is Selected
	bReturn = Fn_ReqMgr_RMTabPanelOperation("VerifyActivate","Trace Link","")
	If bReturn=False Then
		bReturn=Fn_ReqMgr_RMTabPanelOperation("Activate","Trace Link","")
		wait 30
	End If

	JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").SetTOProperty "label","Hide Trace Link"

	'Set Button label = "Hide Trace Link" by clicking on it if its not   
	If  not JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").Exist(iTimeOut) Then
		JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").SetTOProperty "label","Show Trace Link"
		JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").Click
	End If

	strTable=split(sAction,":")
	'Set the table to perform the action
	If  strTable(0)="DefiningTable" Then
		Set objTable= JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("DefiningObject")		
		sColHeader="Defining Object"
	Else
		Set objTable= JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("ComplyingObject")
		sColHeader="Complying Object"
	End If

	'get the Rows in Table
	iRowNo=objTable.GetROProperty("rows")
	'Select the Row for Operation
	For iStart=0 to iRowNo-1
		objTable.SelectRow iStart
		If  sNodeName=objTable.GetCellData(iStart,sColHeader) then
			bSelectFlag=True
			Exit for
		End if
	Next
			
	'Select Case for opearation	
	strOperation=strTable(1)


	Select Case strOperation
		Case "NodeVerify", "Select", "Expand"		'("ComplyingTable:NodeVerify","REQ-000148/A;1-Req2:Req2->Req3","","","")
			If strOperation="Expand" Then
				objTable.DoubleClickCell iStart,sColHeader
			End If

			If bSelectFlag=True Then
				Fn_ReqMgr_DataPanelTraceLinkOpeartions=True
			End If					
				
		Case "VerifyCellValue"
		'Verify Cell Value								
			If  strOperation="VerifyCellValue" and sColName= "Relation Type" and sColValue="Trace Link" and bSelectFlag=True then		
				Fn_ReqMgr_DataPanelTraceLinkOpeartions=True
			End If

		Case "Go To Object"
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").SetTOProperty "label","Go To Object"
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").Click
			wait 30
'			msgbox inStr(1,JavaWindow("DefaultWindow").GetROProperty("label"),"My",1)
			If inStr(1,JavaWindow("DefaultWindow").GetROProperty("label"),"My",1)>0 and bSelectFlag=True Then
				Fn_ReqMgr_DataPanelTraceLinkOpeartions=True
			End If

		Case "Refresh Report"
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").SetTOProperty "label","Refresh Report"
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").Click
			wait 30
'			msgbox inStr(1,JavaWindow("DefaultWindow").GetROProperty("label"),"My",1)
			If bSelectFlag=True Then
				Fn_ReqMgr_DataPanelTraceLinkOpeartions=True
			End If

		Case "Properties"	'New 
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").SetTOProperty "label","Properties"
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").Click
			wait 30
			'Need to modify this case as per requirement
			If bSelectFlag=True Then
				Fn_ReqMgr_DataPanelTraceLinkOpeartions=True
			End If

		Case "Delete Trace Link"	'("ComplyingTable:Delete Trace Link","REQ-000148/A;1-Req2:Req2->Req3","","","")
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").SetTOProperty "label","Delete Trace Link"
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("DtPanelButton").Click
			

			If Fn_UI_ObjectExist("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Refresh Window")) Then
				Call Fn_Button_Click("Fn_TraceabilityReportOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Refresh Window"),"Yes")
			End If


			If bSelectFlag=True Then
				Fn_ReqMgr_DataPanelTraceLinkOpeartions=True
			End If

		
	End Select

End Function

'-------------------------------------------------------------------------Function for Open Text Error Dialog ------------------------------------------------------------------------------------------
'Function Name		:			Fn_ReqMgr_ErrorMessageVerify

'Description		:			This Function is used to handle Error Dialog

'Parameters			:			1.	sDilogName:Error Dialog Box Name
'								2.	sErrorMessage:Expected Error Message

'Return Value		:			True/False

'Pre-requisite		:			Error Dialog should be displayed .

'Examples			:			
										'Fn_ReqMgr_ErrorMessageVerify("Create Trace Link failed","Many")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N				04-June-2010		1.0										                       Tushar B	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ReqMgr_ErrorMessageVerify(sDilogName,sErrorMessage)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_ErrorMessageVerify"
	GBL_EXPECTED_MESSAGE=sErrorMessage
	On Error Resume Next
	Dim sErrMsg,objStaticText,objErrorDialog,iCnt
	JavaWindow("RequirementsManager").JavaWindow("Open Text Error").SetTOProperty "title",sDilogName
	'Checking Error Dialog Exist or not 
	If  Fn_UI_ObjectExist("Fn_ReqMgr_ErrorMessageVerify",JavaWindow("RequirementsManager").JavaWindow("Open Text Error"))=True  Then
			'Getting the label of  ErrMsg in sErrMsg
        	Set objStaticText=Description.Create()
			objStaticText("Class Name").value="JavaStaticText"
			'Taking Child object of present Error Dialog box
			Set objErrorDialog=JavaWindow("RequirementsManager").JavaWindow("Open Text Error").ChildObjects(objStaticText)
			
			For iCnt=0 to objErrorDialog.count-1
				'Checking Error message 
                sErrMsg=objErrorDialog(iCnt).getROProperty("label")
				If Instr(1,Lcase(sErrMsg),Lcase(sErrorMessage))>0 Then
					'Clicking on ok button
					Call Fn_Button_Click("Fn_ReqMgr_ErrorMessageVerify",JavaWindow("RequirementsManager").JavaWindow("Open Text Error"),"OK")
					Fn_ReqMgr_ErrorMessageVerify=True
					Exit For
				Else
					GBL_ACTUAL_MESSAGE=sErrMsg
					Fn_ReqMgr_ErrorMessageVerify=False
				End If
			Next
	Else
		Fn_ReqMgr_ErrorMessageVerify=False
	End If
	Set objStaticText=Nothing
	Set objErrorDialog=Nothing
End Function

'*********************************************************		Function for Importing Requirement Specification ***********************************************************************
'Function Name		:			Fn_ReqMgr_ImportReqSpec

'Description		:			This function is used to Import Specification

'Parameters			:			1.	sFileName:sFileName is complete path of file with file name and its extension
'											 2.	sSpecType:Specification Type
'											3.	sDescription:Desription Of Impoerted specification
'											4.sOption: "Import "  OR "Keyword"
'											5.sKeywords:
'											6.sSubType:Specification Sub Type

'											Imp Note:sOption is Either  "Import "  OR "Keyword"
'Return Value		:		True/False

'Pre-requisite		:		Requirement Manager window should be displayed .

'Examples			:				
									'Call Fn_ReqMgr_ImportReqSpec("D:\mainline\TestData\Requierment Data File.docx","RequirementSpec","Test","Import","","Requirement")
									'Call Fn_ReqMgr_ImportReqSpec("D:\mainline\TestData\Requierment Data File.docx","RequirementSpec","Test","Keyword","Program","Requirement")
									'Call Fn_ReqMgr_ImportReqSpec("D:\mainline\TestData\Req_RM037\Requierment_Data_File.docx","RequirementSpec","","Keyword","~Test","")-To Verify value of Keiwords edit box
'History:					
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N				7-June-2010			1.0																	Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'sFileName is complete path of file with file name and its extension
Public Function Fn_ReqMgr_ImportReqSpec(sFileName,sSpecType,sDescription,sOption,sKeywords,sSubType)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_ImportReqSpec"
	On Error Resume Next
	'Declaring Variables 
	Dim iItemCnt,iCnt,bFlag,strSource,strVerify
	'Declaring Objects
	Dim objImportDialog,objImportSpecDialog
	
	Fn_ReqMgr_ImportReqSpec=False
	bFlag=False
	'Setting Object of  "Import Spec" Dilog
	Set objImportDialog=JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Import Spec")
	'Checking "Import Spec" Dialog Is exist or Not 
		If Fn_UI_ObjectExist("Fn_ReqMgr_ImportReqSpec",objImportDialog)=False Then
			'Calling Menuoperation to Open Import Spec Dialog
			Call Fn_MenuOperation("Select","File:Import Spec...")
		End If
	'Creating object of  "Import Spec" 
	Set objImportSpecDialog=Fn_UI_ObjectCreate("Fn_ReqMgr_ImportReqSpec",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Import Spec"))
		'Checking File Name is pass or not
		'File Name is Compulsory Parameter if not pass then function will exit
		If sFileName<>"" Then
            Call Fn_Edit_Box("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog,"FileName",sFileName)
		Else
			Set objImportDialog=Nothing
			Set objImportSpecDialog=Nothing
			Exit Function
		End If
		'Checking Specification Type  pass or not
		If sSpecType<>"" Then
			'Retriwing Items Count from "SpecType" Java List
			iItemCnt= Fn_UI_Object_GetROProperty("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog.JavaList("SpecType"),"items count")
			For iCnt=0 To iItemCnt-1
				If  objImportSpecDialog.JavaList("SpecType").GetItem(iCnt)=sSpecType Then
					objImportSpecDialog.JavaList("SpecType").Select(sSpecType)
					'If item is present in list then select it and changing bFlag to True
					bFlag=True
					Exit For
				End If
			Next
		End If
	'If Wrong Item Name is pass then function will exit and return False
	If bFlag=False Then
		Exit Function
	End If
	
	If sDescription<>"" Then
		'Setting Description
		Call Fn_Edit_Box("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog,"Description",sDescription)
	End If
	'Clicking On Next Button to go on Next Window
	Call Fn_Button_Click("Fn_ReqMgr_ImportReqSpec", objImportSpecDialog, "Next")
	'selection Import Option
	If Ucase(sOption)="IMPORT" Then
		'Setting "ImportAsSingleSubtype" Radion Button To "ON"
		objImportSpecDialog.JavaRadioButton("ImportAsSingleSubtype").Set "ON"
		bFlag=False
		If sSubType<>"" Then
			'Selecting Item From "ImportSubType" List
			iItemCnt= Fn_UI_Object_GetROProperty("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog.JavaList("ImportSubType"),"items count")
			For iCnt=0 To iItemCnt-1
				If objImportSpecDialog.JavaList("ImportSubType").GetItem(iCnt)=sSubType Then
					objImportSpecDialog.JavaList("ImportSubType").Select(sSubType)
					bFlag=True
					Exit For
				End If
			Next
		
				If bFlag=False Then
					Call Fn_Button_Click("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog, "Close")
					Set objImportDialog=Nothing
					Set objImportSpecDialog=Nothing
					Exit Function
				End If
		End If
	End If
	'selection Keyword Option
	If Ucase(sOption)="KEYWORD" Then
		'Setting "ImportAsSingleSubtype" Radion Button To "ON"
		objImportSpecDialog.JavaRadioButton("UseKwdsForImport").Set "ON"
		bFlag=False
		If sKeywords<>"" Then
				If inStr(1,sKeywords,"~")>0 Then
				strSource=objImportSpecDialog.JavaEdit("Keywords").GetROProperty("value")
				strVerify=split(sKeywords,"~")(1)
				 If  inStr(1,strSource,strVerify)>0 then
					 Fn_ReqMgr_ImportReqSpec=True
					 Call Fn_Button_Click("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog, "Close")
                    Set objImportSpecDialog=Nothing
					Set objImportDialog=Nothing
					Exit Function
				 Else
					'Clicking on "OK" To save
					Call Fn_Button_Click("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog, "Close")

				 End if
 					Fn_ReqMgr_ImportReqSpec=False
					Set objImportSpecDialog=Nothing
					Set objImportDialog=Nothing
					Exit Function

			
			 Else
			'Setting KeyWords
			Call Fn_Edit_Box("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog,"Keywords",sKeywords)
			End If
		End If
		'Selecting Item From "KwdSubType" List
		If  sSubType<>"" Then
			iItemCnt= Fn_UI_Object_GetROProperty("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog.JavaList("KwdSubType"),"items count")
			For iCnt=0 To iItemCnt-1
				If objImportSpecDialog.JavaList("KwdSubType").GetItem(iCnt)=sSubType Then
					objImportSpecDialog.JavaList("KwdSubType").Select(sSubType)
					bFlag=True
					Exit For
				End If
			Next

			If bFlag=False Then
				Call Fn_Button_Click("Fn_ReqMgr_ImportReqSpec",objImportSpecDialog, "Close")
				Set objImportDialog=Nothing
				Set objImportSpecDialog=Nothing
				Exit Function
			End If
		End If
	End If
	'Clicking On finish Button
	Call Fn_Button_Click("Fn_ReqMgr_ImportReqSpec", objImportSpecDialog, "Finish")
	Fn_ReqMgr_ImportReqSpec=True
	Set objImportDialog=Nothing
	Set objImportSpecDialog=Nothing
End Function

'*********************************************************		Function to preform operation on MRU List	**************************************************************
'Function Name		:				Fn_ReqMgr_MRUListOperations

'Description			 :		 		Perform operations on MRU List	

'Parameters			   :	 			1.strAction: Action to perform(Select or Exist)
'													 2.strButtonName: Button Name which appear in MRU List (Item ID-Item Name)

'Return Value		   : 				True Or False

'Pre-requisite			:		 		Should be logged in & present on Requirement Manager perspective

'Examples				:				 Fn_ReqMgr_MRUListOperations("Select","001173-test3")
'													Fn_ReqMgr_MRUListOperations("Exist","001174-test4")
'													
'													IMP NOTE :-- To use this function First clear the cache and relaunch the Teamcenter
'													Fn_ReUserTcSession(True,True, Environment.Value("TcUser1"))
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   08/06/2010			          1.0										Created								Archana D
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_MRUListOperations(strAction,strButtonName)
GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_MRUListOperations"
'Declaring Object
Dim objRMWindowApplet
    Fn_ReqMgr_MRUListOperations=False
	'Checking Existance of  "RMWindowApplet" Applet  : Its Prerequisite
	If Fn_UI_ObjectExist("Fn_ReqMgr_MRUListOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))=False Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"RMWindowApplet applet window is Not present")
		Exit Function
	End If
	'Creating Object of  "RMWindowApplet" Applet
	Set objRMWindowApplet=Fn_UI_ObjectCreate("Fn_ReqMgr_MRUListOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))
	wait(40)
	'Clicking "MostRecentlyUsedButton" Button To Activate MRU List
	Call Fn_CheckBox_Set("Fn_ReqMgr_MRUListOperations",objRMWindowApplet, "MostRecentlyUsedButton",  "ON" )
	'Changing "label" property of  "MRUListButton" to strButtonName
	objRMWindowApplet.JavaButton("MRUListButton").SetTOProperty "label",strButtonName
	'Selecting Case
	Select Case strAction
		'To select list Item
		Case "Select"
			'To selecting Item From MRU List
			Call Fn_Button_Click("Fn_ReqMgr_MRUListOperations",objRMWindowApplet,"MRUListButton")
		'To Check Existance of Item in MRU List
		Case "Exist"
			'Checking Item Is exist or Not
			If JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("MRUListButton").Exist(30)=True Then
				wait(20)
				'Closing the MRU List
				JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaObject("CloseMRUList").Click 1,1
			Else
				wait(20)
				'Closing the MRU List
				JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaObject("CloseMRUList").Click 1,1
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"),strButtonName +"Button is not Exist in MRU List")
				'Setting Object to Nothing
				Set objRMWindowApplet=Nothing
				'If item is not exist then function will exit and return False
				Exit Function
			End If
	End Select
	'Function Returning True
	Fn_ReqMgr_MRUListOperations=True
	'Setting Object to Nothing
	Set objRMWindowApplet=Nothing
End Function

'*********************************************************		Function to create  RequirmentSpec With details	***********************************************************************
'Function Name		:				Fn_ReqMgr_RequirmentSpecDetailsCreate

'Description			 :		 		 Creats an Requirement Specification with Detail information

'Parameters			   :	 			1.sSpecType: Type of the item.(Requirement Specification)
'													 2.sConfItem: True or False
'													 2.sSpecID: ID of the Specification it should be unique.
'													3.sSpecRevID:Revision ID of the Specification.
'													4.sSpecName:Name of Specification.
'													5.sSpecDesc: Description of the Specification.
'													6:sSpecUOM: Unit of measure of Specification. ( not handling this part)
'													7.aReqSpecInfo:Additional information of Requirement Specification
'													8.aReqSpecRevInfo:Additional information of Requirement Specification Revision
'													9.sProjectNames: Project Names to select
'													10 aDefineOpt:To define options
'													Imp Note: All the parameters which are start with "a"  letter are  Arrays so pass the array
'													
'Return Value		   : 				Specification Id  / Revision Id

'Pre-requisite			:		 		should be logged in & present on Requirement Manager perspective

'Examples				:				aReqSpec=Array("","","","")
'													aReqSpecRev=Array("","","","","","","","","","")
'													aDefineOpts=Array("OFF","","","")
'													Fn_ReqMgr_RequirmentSpecDetailsCreate("RequirementSpec","","","","TestItem2","","","","","",aDefineOpts)
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   09/06/2010			              1.0										Created							Archan D
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_RequirmentSpecDetailsCreate(sSpecType,sConfItem,sSpecID,sSpecRevID,sSpecName,sSpecDesc,sSpecUOM,aReqSpecInfo,aReqSpecRevInfo,sProjectNames,aDefineOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_RequirmentSpecDetailsCreate"
	on Error Resume Next
	Dim sSpecificationId, sRevId,aListItem,iItemCnt,iCnt
	Dim objDialogNewSpec,objSelectType,objDialog
	Fn_ReqMgr_RequirmentSpecDetailsCreate=False
	
	If Fn_UI_ObjectExist("Fn_ReqMgr_RequirmentSpecDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirementSpec"))=False Then
         Call Fn_MenuOperation("Select","File:New:Requirements Spec...")
	End If
	
	'Check the existence of "NewRequirementSpec" window
	Set objDialogNewSpec=Fn_UI_ObjectCreate("Fn_ReqMgr_RequirmentSpecDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirementSpec"))
		'Select Item Type
		Call Fn_List_Select("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"RequirementSpecType",sSpecType)
		'checked Configuration RequirementSpec or not
		If sConfItem <> "" Then
		 Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"Configuration Item",sConfItem)
		End If
		'Click on "Next" button
		 Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"Next")
		
		If sSpecID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"RequirementSpecID", sSpecID)
		End If
		
		If sSpecRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate",objDialogNewSpec,"RevID", sSpecRevID)
		End If
		
		If  sSpecID = "" or sSpecRevID = "" Then
			'click on assign button
			  Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec, "Assign")
		End If
		
		'Extract Creation data
		sSpecificationId =Fn_Edit_Box_GetValue("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"RequirementSpecID")
		sRevId = Fn_Edit_Box_GetValue("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"RevID")
		
		'Set RequirementSpec name
		 Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"SpecName",sSpecName)
		'Set description
		Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"Description",sSpecDesc)
		'Set UOM
			If sSpecUOM <> "" Then
				 Set objSelectType=description.Create()
				objSelectType("Class Name").value = "JavaStaticText"
				objSelectType("label").value = sSpecUOM
				objDialogNewSpec.JavaButton("UnitOfMeasureDrpDwn").Click
				Set objDialog =objDialogNewSpec.ChildObjects(objSelectType)
				objDialog(0).Click 5, 5, "LEFT"
		   End If

		If IsArray(aReqSpecInfo)=True Then
			'Changing label  of Stpes to "Enter Additional Requirement Spec Information"
			objDialogNewSpec.JavaStaticText("Stpes").SetTOProperty "label","Enter Additional Requirement Spec Information"
			'Clicking "Enter Additional Requirement Spec Information" Static text
			objDialogNewSpec.JavaStaticText("Stpes").Click 5,5
			wait(5)
            If aReqSpecInfo(0)<>"" Then
				'Setting the Titte
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"Title",aReqSpecInfo(0))
			End If
			If aReqSpecInfo(1)<>"" Then
				'Setting the Author
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"Author",aReqSpecInfo(1))
			End If
			If aReqSpecInfo(2)<>"" Then
				'Setting the Subject
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"Subject",aReqSpecInfo(2))
			End If
			If aReqSpecInfo(3)<>"" Then
				'Selecting Item From list
                objDialogNewSpec.JavaList("KeywordsList").Select aReqSpecInfo(3)'Functionality is not ready to use in Application
			End If
		End If
		
		If IsArray(aReqSpecRevInfo)=True Then
			'Changing label  of Stpes to "Enter Additional Requirement Spec Revision Information"
			objDialogNewSpec.JavaStaticText("Stpes").SetTOProperty "label","Enter Additional Requirement Spec Revision Information"
			'Clicking "Enter Additional Requirement Spec Revision Information" Static Text
			objDialogNewSpec.JavaStaticText("Stpes").Click 5,5
			wait(5)
			If aReqSpecRevInfo(0)<>"" Then
				'Setting Document Author
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"DocumentAuthor",aReqSpecRevInfo(0))
			End If
			If aReqSpecRevInfo(1)<>"" Then
				'Setting Document Subject
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"DocumentSubject",aReqSpecRevInfo(1))
			End If
			If aReqSpecRevInfo(2)<>"" Then
				'Setting Document Title
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"DocumentTitle",aReqSpecRevInfo(2))
			End If
			If aReqSpecRevInfo(3)<>"" Then
				'Setting Project ID
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"ProjectID",aReqSpecRevInfo(3))
			End If
			If aReqSpecRevInfo(4)<>"" Then
				'Setting Revision Previos ID
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"ParaRevPreviousID",aReqSpecRevInfo(4))
			End If
			If aReqSpecRevInfo(5)<>"" Then
				'Setting Serial Number
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"SerialNumber",aReqSpecRevInfo(5))
			End If
			If aReqSpecRevInfo(6)<>"" Then
				'Setting Item Comment
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"ItemComment",aReqSpecRevInfo(6))
			End If
			If aReqSpecRevInfo(7)<>"" Then
				'Setting UserData 1
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"UserData1",aReqSpecRevInfo(7))
			End If
			If aReqSpecRevInfo(8)<>"" Then
				'Setting User Data 2
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"UserData2",aReqSpecRevInfo(8))
			End If
			If aReqSpecRevInfo(9)<>"" Then
				'Setting User Data 2
                Call Fn_Edit_Box("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec,"UserData3",aReqSpecRevInfo(9))
			End If
		End If

		If sProjectNames<>"" Then
			'Changing label  of Stpes to "Assign to Project"
			objDialogNewSpec.JavaStaticText("Stpes").SetTOProperty "label","Assign to Project"
			'Clicking "Assign to Project" Static Text
			objDialogNewSpec.JavaStaticText("Stpes").Click 5,5
			wait(5)
			'Selecting Project Names and Shifting to Selected Project List
			iItemCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_RequirmentSpecDetailsCreate",objDialogNewSpec.JavaList("ProjectForSelect"),"items count")			
			aListItem=Split(sProjectNames,":")
			objDialogNewSpec.JavaList("ProjectForSelect").Select aListItem(0)
				For iCnt=1 to Ubound(aListItem)
					objDialogNewSpec.JavaList("ProjectForSelect").ExtendSelect aListItem(iCnt)
				Next
			Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec, "AddProject")
		End If

        If IsArray(aDefineOpt)=True Then
			'Changing label  of Stpes to "Define Options"
			objDialogNewSpec.JavaStaticText("Stpes").SetTOProperty "label","Define Options"
			'Clicking "Define Options"
        	objDialogNewSpec.JavaStaticText("Stpes").Click 5,5
        	wait(5)
			If aDefineOpt(0)<>"" Then
				'Setting  "ShowAsNwRt" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec, "ShowAsNwRt",aDefineOpt(0))
			End If
			If aDefineOpt(1)<>"" Then
				'Setting  "UsItIdentifierAs" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec, "UsItIdentifierAs",aDefineOpt(1))
			End If
			If aDefineOpt(2)<>"" Then
				'Setting  "UsRevIdentifier" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec, "UsRevIdentifier",aDefineOpt(2))
			End If
			If aDefineOpt(3)<>"" Then
				'Setting  "ChkOutItmRevOnCr" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec, "ChkOutItmRevOnCr",aDefineOpt(3))
			End If
		End If

			wait(2)
			objDialogNewSpec.JavaButton("Finish").WaitProperty "enabled", 1, 20000
        	'Clicking on finish button
			Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec, "Finish") 
            Call Fn_ReadyStatusSync(1)

			If Fn_UI_ObjectExist("Fn_ReqMgr_RequirmentSpecDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirementSpec"))=True Then		
					Call Fn_Button_Click("Fn_ReqMgr_RequirmentSpecDetailsCreate", objDialogNewSpec, "Close")
			End If
		Fn_ReqMgr_RequirmentSpecDetailsCreate = sSpecificationId & "-" & sRevId
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Specification of ID [" + CStr(sSpecificationId) + "]")
		Set objDialogNewSpec=Nothing
		Set objSelectType=Nothing
		Set objDialog=Nothing
		Set aListItem=Nothing
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'*********************************************************		Function to create  Requirment With details	***********************************************************************
'Function Name		:				Fn_ReqMgr_RequirmentDetailsCreate

'Description			 :		 		 Creats an Requirement  with Detail information

'Parameters			   :	 			1.sReqType: Type of the item.(Requirement Specification)
'													 2.sConfItem: True or False
'													3.sReqID: ID of the Specification it should be unique.
'													4.sReqRevID:Revision ID of the Specification.
'													5.sReqName:Name of Specification.
'													6.sReqDesc: Description of the Specification.
'													7:sReqUOM: Unit of measure of Specification. ( not handling this part)
'													8.aAddReqInfo:Additional information of Requirement 
'													9.aAddReqRevInfo:Additional information of Requirement  Revision
'													10.sAttachAction: Attache ment Action (Browse or Remove)--SeparateWith '  *   ' 
'													11.aWorkflowInfo:Additional information of Workflow
'													12.sProjectNames:Project Names to select
'													13 aDefineOpt:To define options
'													Imp Note: All the parameters which are start with "a"  letter are  Arrays so pass the array
'													
'Return Value		   : 				Requirment Id  / Revision Id

'Pre-requisite			:		 		should be logged in & present on Requirement Manager perspective

'Examples				:				arrWorkflowInfo=Array("CMII WA","")
'													aDefineOpts=Array("OFF","","","")
'													Fn_ReqMgr_RequirmentDetailsCreate("Requirement","","","","TestReq","","","","","Browse*D:\mainline\Scripting_Check_List.doc",arrWorkflowInfo,"",aDefineOpts)
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   09/06/2010			              1.0										Created							Archan D
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ReqMgr_RequirmentDetailsCreate(sReqType,sConfItem,sReqID,sReqRevID,sReqName,sReqDesc,sReqUOM,aAddReqInfo,aAddReqRevInfo,sAttachAction,aWorkflowInfo,sProjectNames,aDefineOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_RequirmentDetailsCreate"
	on Error Resume Next
	Dim sRequirementID, sRevId,iCnt,aAttachAction, iItemCnt,aListItem
	Dim objDialogNewReq,objSelectType,objDialog,objStaticText,objNewRequirementChild
	Fn_ReqMgr_RequirmentDetailsCreate=False
	'Select menu [File -> New -> Requirement...]
		If Fn_UI_ObjectExist("Fn_ReqMgr_RequirmentDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirement"))=False Then
			 Call Fn_MenuOperation("Select","File:New:Requirement...")
		End If
		
	'Check the existence of "NewRequirement" window
		Set objDialogNewReq=Fn_UI_ObjectCreate("Fn_ReqMgr_RequirmentDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewRequirement"))
			'Select Item Type
		Call Fn_List_Select("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"RequirementType",sReqType)
		'checked Configuration Requirement or not
		If sConfItem <> "" Then
		 Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"Configuration Item",sConfItem)
		End If
		'Click on "Next" button
		 Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"Next")

		If sReqID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ReqID", sReqID)
		End If
	
		If sReqRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate",objDialogNewReq,"ReqRevID", sReqRevID)
		End If

		If  sReqID = "" or sReqRevID = "" Then
			'click on assign button
			  Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "Assign")
		End If
	
		'Extract Creation data
		sRequirementID =Fn_Edit_Box_GetValue("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ReqID")
		sRevId = Fn_Edit_Box_GetValue("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ReqRevID")
		
		'Set Requirement name
		 Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ReqName",sReqName)
		'Set description
		Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"Description",sReqDesc)
		'Set UOM
		If sReqUOM <> "" Then
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sReqUOM
			objDialogNewReq.JavaButton("UOMDrpDwn").Click
			Set objDialog =objDialogNewReq.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
	   End If

		If IsArray(aAddReqInfo)=True Then
			'Changing "Stpes" Label to "Enter Additional Requirement Information"
			objDialogNewReq.JavaStaticText("Stpes").SetTOProperty "label","Enter Additional Requirement Information"
			'Clicking on "Enter Additional Requirement Information" static text
			objDialogNewReq.JavaStaticText("Stpes").Click 5,5
			wait(3)
			If aAddReqInfo(0)<>"" Then
				'Setting "ProjectID"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ProjectID",aAddReqInfo(0))
			End If
			If aAddReqInfo(1)<>"" Then
				'Setting "PreviousID"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"PreviousID",aAddReqInfo(1))
			End If
			If aAddReqInfo(2)<>"" Then
				'Setting "SerialNumber"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"SerialNumber",aAddReqInfo(2))
			End If
			If aAddReqInfo(3)<>"" Then
				'Setting "ItemComment"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ItemComment",aAddReqInfo(3))
			End If
			If aAddReqInfo(4)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"UserData1",aAddReqInfo(4))
			End If
			If aAddReqInfo(5)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"UserData2",aAddReqInfo(5))
			End If
			If aAddReqInfo(6)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"UserData3",aAddReqInfo(6))
			End If
		End If
		
		If IsArray(aAddReqRevInfo)=True Then
			'Changing "Stpes" Label to "Enter Additional Requirement  Revision Information"
			objDialogNewReq.JavaStaticText("Stpes").SetTOProperty "label","Enter Additional Requirement Revision Information"
			'Clicking on "Enter Additional Requirement  Revision Information" static text
			objDialogNewReq.JavaStaticText("Stpes").Click 5,5
			wait(3)
			If aAddReqRevInfo(0)<>"" Then
				'Setting "ProjectID"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ProjectID",aAddReqRevInfo(0))
				
			End If
			If aAddReqRevInfo(1)<>"" Then
				'Setting "ReqRevPreviousID"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ReqRevPreviousID",aAddReqRevInfo(1))
			End If
			If aAddReqRevInfo(2)<>"" Then
				'Setting "SerialNumber"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"SerialNumber",aAddReqRevInfo(2))
			End If
			If aAddReqRevInfo(3)<>"" Then
				'Setting "ItemComment"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"ItemComment",aAddReqRevInfo(3))
			End If
			If aAddReqRevInfo(4)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"UserData1",aAddReqRevInfo(4))
			End If
			If aAddReqRevInfo(5)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"UserData2",aAddReqRevInfo(5))
			End If
			If aAddReqRevInfo(6)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq,"UserData3",aAddReqRevInfo(6))
			End If
		End If

		If sAttachAction<>"" Then
			'Changing "Stpes" Label to "Enter Additional Requirement Information"
			objDialogNewReq.JavaStaticText("Stpes").SetTOProperty "label","Enter Attach Files Information"
			'Clicking on "Enter Additional Requirement Information" static text
			objDialogNewReq.JavaStaticText("Stpes").Click 5,5
			'Spliting "sAttachAction"
			aAttachAction=Split(sAttachAction,"*")
			'Taking External files
			If aAttachAction(0)="Browse" Then
				For iCnt=1 To Ubound(aAttachAction)
						Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "Browse...")
						Call Fn_Edit_Box("Fn_ReqMgr_RequirmentDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Open"),"FileName",aAttachAction(iCnt))
						Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Open"), "Open")
				Next
			End If
			''Removing Files from List
			If aAttachAction(0)="Remove" Then
				If aAttachAction(1)<>"" Then
					objDialogNewReq.JavaList("AttachFileReq").Select aAttachAction(1)
					For iCnt=2 To Ubound(aAttachAction)
                        objDialogNewReq.JavaList("AttachFileReq").ExtendSelect aAttachAction(iCnt)
					Next
				End If
				Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "Remove")
			End If
		End If

		'Defining Workflow information
		If IsArray(aWorkflowInfo)=True Then
			'Changing "Stpes" Label to "Define Workflow Information"
			objDialogNewReq.JavaStaticText("Stpes").SetTOProperty "label","Define Workflow Information"
			'Clicking on "Define Workflow Information" static text
			objDialogNewReq.JavaStaticText("Stpes").Click 5,5

			If aWorkflowInfo(0)<>"" Then
				'Selecting Item from Process Template List
				Set objStaticText=Description.Create()
				objStaticText("Class Name").value="JavaStaticText"
				objStaticText("label").value=aWorkflowInfo(0)
				Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "ProcessTemplateDrpDwn")
				Set objNewRequirementChild=objDialogNewReq.ChildObjects(objStaticText)
						objNewRequirementChild(0).Click 5,5
				
			End If

			Set objStaticText=Nothing
			Set objNewRequirementChild=Nothing

			If aWorkflowInfo(1)<>"" Then
				'Selecting Item from Process Assignment List
				Set objStaticText=Description.Create()
				objStaticText("Class Name").value="JavaStaticText"
				objStaticText("label").value=aWorkflowInfo(1)
				Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "ProcessAssignmentDrpDwn")
				Set objNewRequirementChild=objDialogNewReq.ChildObjects(objStaticText)
					objNewRequirementChild(0).Click 5,5
			End If

			Set objStaticText=Nothing
			Set objNewRequirementChild=Nothing

		End If

		If sProjectNames<>"" Then
			'Changing label  of Stpes to "Assign to Project"
			objDialogNewReq.JavaStaticText("Stpes").SetTOProperty "label","Assign to Project"
			'Clicking "Assign to Project" Static Text
			objDialogNewReq.JavaStaticText("Stpes").Click 5,5
			wait(5)
			'Selecting Project Names and Shifting to Selected Project List
			iItemCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_RequirmentDetailsCreate",objDialogNewReq.JavaList("ProjectForSelect"),"items count")			
			aListItem=Split(sProjectNames,":")
			objDialogNewReq.JavaList("ProjectForSelect").Select aListItem(0)
				For iCnt=1 to Ubound(aListItem)
					objDialogNewReq.JavaList("ProjectForSelect").ExtendSelect aListItem(iCnt)
				Next
			Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "AddProject")
		End If
		
        If IsArray(aDefineOpt)=True Then
			'Changing label  of Stpes to "Define Options"
			objDialogNewReq.JavaStaticText("Stpes").SetTOProperty "label","Define Options"
			'Clicking "Define Options"
        	objDialogNewReq.JavaStaticText("Stpes").Click 5,5
        	wait(5)
			If aDefineOpt(0)<>"" Then
				'Setting  "ShowAsNwRt" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "ShowAsNwRt",aDefineOpt(0))
			End If
			If aDefineOpt(1)<>"" Then
				'Setting  "UsItIdentifierAs" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "UsItIdentifierAs",aDefineOpt(1))
			End If
			If aDefineOpt(2)<>"" Then
				'Setting  "UsRevIdentifier" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "UsRevIdentifier",aDefineOpt(2))
			End If
			If aDefineOpt(3)<>"" Then
				'Setting  "UsRevIdentifier" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "ChkOutItmRevOnCr",aDefineOpt(3))
			End If
		End If

		wait(2)
		objDialogNewReq.JavaButton("Finish").WaitProperty "enabled", 1, 20000
        'Click on "Finish" button
    	Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "Finish") 
		Fn_ReqMgr_RequirmentDetailsCreate = sRequirementID & "-" & sRevId
		Call Fn_ReadyStatusSync(1)

		'Click on Close button
		Call Fn_Button_Click("Fn_ReqMgr_RequirmentDetailsCreate", objDialogNewReq, "Close")
						
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Requirement of ID [" + CStr(sRequirementID) +"-"+ sReqName + "]")
	Set objDialogNewReq=Nothing
	Set objSelectType=Nothing
	Set objDialog=Nothing
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'*********************************************************		Function to create  Paragraph With details	***********************************************************************
'Function Name		:				Fn_ReqMgr_ParagraphDetailsCreate

'Description			 :		 		 Creats an Paragraph  with Detail information

'Parameters			   :	 			1.sParaType: Type of the item.(Paragraph Specification)
'													 2.sConfItem: True or False
'													3.sParaID: ID of the Specification it should be unique.
'													4.sParaRevID:Revision ID of the Specification.
'													5.sParaName:Name of Specification.
'													6.sParaDesc: Description of the Specification.
'													7:sParaUOM: Unit of measure of Specification. ( not handling this part)
'													8.aAddReqInfo:Additional information of Paragraph 
'													9.aAddReqParaInfo:Additional information of Paragraph  Revision
'													10.sAttachAction: Attache ment Action (Browse or Remove)--SeparateWith '  *   ' 
'													11.aWorkflowInfo:Additional information of Workflow
'													12.sProjectNames:Project Names to select
'													13 aDefineOpt:To define options
'													Imp Note: All the parameters which are start with "a"  letter are  Arrays so pass the array
'													
'Return Value		   : 				Requirment Id  / Revision Id

'Pre-requisite			:		 		should be logged in & present on Paragraph Manager perspective

'Examples				:				arrWorkflowInfo=Array("CMII WA","")
'													aDefineOpts=Array("OFF","","","")
'													Fn_ReqMgr_ParagraphDetailsCreate("Paragraph","","","","TestReq","","","","","Browse*D:\mainline\Scripting_Check_List.doc",arrWorkflowInfo,"",aDefineOpts)
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   09/06/2010			              1.0										Created							Archan D
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Fn_ReqMgr_ParagraphDetailsCreate(sParaType,sConfItem,sParaID,sParaRevID,sParaName,sParaDesc,sParaUOM,aAddParaInfo,aAddParaRevInfo,sAttachAction,aWorkflowInfo,sProjectNames,aDefineOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_ParagraphDetailsCreate"
	on Error Resume Next
	Dim sParagraphID, sRevId,iCnt,aAttachAction, iItemCnt,aListItem
	Dim objDialogNewPara,objSelectType,objDialog,objStaticText,objNewParagraphChild
		
	Fn_ReqMgr_ParagraphDetailsCreate=False
	'Select menu [File -> New -> Paragraph...]
		If Fn_UI_ObjectExist("Fn_ReqMgr_ParagraphDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewParagraph"))=False Then
			 Call Fn_MenuOperation("Select","File:New:Paragraph...")
		End If
		
	'Check the existence of "NewParagraph" window
		Set objDialogNewPara=Fn_UI_ObjectCreate("Fn_ReqMgr_ParagraphDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewParagraph"))
			'Select Item Type
		Call Fn_List_Select("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ParagraphType",sParaType)
		'checked Configuration Paragraph or not
		If sConfItem <> "" Then
		 Call Fn_CheckBox_Set("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"Configuration Item",sConfItem)
		End If
		'Click on "Next" button
		 Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"Next")

		If sParaID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ParaID", sParaID)
		End If
	
		If sParaRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate",objDialogNewPara,"ParaRevID", sParaRevID)
		End If

		If  sParaID = "" or sParaRevID = "" Then
			'click on assign button
			  Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "Assign")
		End If
	
		'Extract Creation data
		sParagraphID =Fn_Edit_Box_GetValue("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ParaID")
		sRevId = Fn_Edit_Box_GetValue("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ParaRevID")
		
		'Set Paragraph name
		 Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ParaName",sParaName)
		'Set description
		Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"Description",sParaDesc)
		'Set UOM
		If sParaUOM <> "" Then
			Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sParaUOM
			objDialogNewPara.JavaButton("UOMDrpDwn").Click
			Set objDialog =objDialogNewPara.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
	   End If

		If IsArray(aAddParaInfo)=True Then
			'Changing "Stpes" Label to "Enter Additional Paragraph Information"
			objDialogNewPara.JavaStaticText("Stpes").SetTOProperty "label","Enter Additional Paragraph Information"
			'Clicking on "Enter Additional Paragraph Information" static text
			objDialogNewPara.JavaStaticText("Stpes").Click 5,5
			wait(5)
			If aAddParaInfo(0)<>"" Then
				'Setting "ProjectID"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ProjectID",aAddParaInfo(0))
			End If
			If aAddParaInfo(1)<>"" Then
				'Setting "PreviousID"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"PreviousID",aAddParaInfo(1))
			End If
			If aAddParaInfo(2)<>"" Then
				'Setting "SerialNumber"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"SerialNumber",aAddParaInfo(2))
			End If
			If aAddParaInfo(3)<>"" Then
				'Setting "ItemComment"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ItemComment",aAddParaInfo(3))
			End If
			If aAddParaInfo(4)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"UserData1",aAddParaInfo(4))
			End If
			If aAddParaInfo(5)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"UserData2",aAddParaInfo(5))
			End If
			If aAddParaInfo(6)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"UserData3",aAddParaInfo(6))
			End If
		End If
		
		If IsArray(aAddParaRevInfo)=True Then
			'Changing "Stpes" Label to "Enter Additional Paragraph  Revision Information"
			objDialogNewPara.JavaStaticText("Stpes").SetTOProperty "label","Enter Additional Paragraph Revision Information"
			'Clicking on "Enter Additional Paragraph  Revision Information" static text
			objDialogNewPara.JavaStaticText("Stpes").Click 5,5
			wait(5)
			If aAddParaRevInfo(0)<>"" Then
				'Setting "ProjectID"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ProjectID",aAddParaRevInfo(0))
				
			End If
			If aAddParaRevInfo(1)<>"" Then
				'Setting "ParaRevPreviousID"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ParaRevPreviousID",aAddParaRevInfo(1))
			End If
			If aAddParaRevInfo(2)<>"" Then
				'Setting "SerialNumber"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"SerialNumber",aAddParaRevInfo(2))
			End If
			If aAddParaRevInfo(3)<>"" Then
				'Setting "ItemComment"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"ItemComment",aAddParaRevInfo(3))
			End If
			If aAddParaRevInfo(4)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"UserData1",aAddParaRevInfo(4))
			End If
			If aAddParaRevInfo(5)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"UserData2",aAddParaRevInfo(5))
			End If
			If aAddParaRevInfo(6)<>"" Then
				'Setting "UserData1"
				Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara,"UserData3",aAddParaRevInfo(6))
			End If
		End If

		If sAttachAction<>"" Then
			'Changing "Stpes" Label to "Enter Additional Paragraph Information"
			objDialogNewPara.JavaStaticText("Stpes").SetTOProperty "label","Enter Attach Files Information"
			'Clicking on "Enter Additional Paragraph Information" static text
			objDialogNewPara.JavaStaticText("Stpes").Click 5,5
			'Spliting "sAttachAction"
			wait(5)
			aAttachAction=Split(sAttachAction,"*")
			'Taking External files
			If aAttachAction(0)="Browse" Then
				For iCnt=1 To Ubound(aAttachAction)
						Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "Browse...")
						Call Fn_Edit_Box("Fn_ReqMgr_ParagraphDetailsCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Open"),"FileName",aAttachAction(iCnt))
						Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Open"), "Open")
				Next
			End If
			''Removing Files from List
			If aAttachAction(0)="Remove" Then
				If aAttachAction(1)<>"" Then
					objDialogNewPara.JavaList("AttachFileReq").Select aAttachAction(1)
					For iCnt=2 To Ubound(aAttachAction)
                        objDialogNewPara.JavaList("AttachFileReq").ExtendSelect aAttachAction(iCnt)
					Next
				End If
				Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "Remove")
			End If
		End If

		'Defining Workflow information
		If IsArray(aWorkflowInfo)=True Then
			'Changing "Stpes" Label to "Define Workflow Information"
			objDialogNewPara.JavaStaticText("Stpes").SetTOProperty "label","Define Workflow Information"
			'Clicking on "Define Workflow Information" static text
			objDialogNewPara.JavaStaticText("Stpes").Click 5,5
			wait(5)
			If aWorkflowInfo(0)<>"" Then
				'Selecting Item from Process Template List
				Set objStaticText=Description.Create()
				objStaticText("Class Name").value="JavaStaticText"
				objStaticText("label").value=aWorkflowInfo(0)
				Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "ProcessTemplateDrpDwn")
				Set objNewParagraphChild=objDialogNewPara.ChildObjects(objStaticText)
						objNewParagraphChild(0).Click 5,5
				
			End If

			Set objStaticText=Nothing
			Set objNewParagraphChild=Nothing

			If aWorkflowInfo(1)<>"" Then
				'Selecting Item from Process Assignment List
				Set objStaticText=Description.Create()
				objStaticText("Class Name").value="JavaStaticText"
				objStaticText("label").value=aWorkflowInfo(1)
				Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "ProcessAssignmentDrpDwn")
				Set objNewParagraphChild=objDialogNewPara.ChildObjects(objStaticText)
					objNewParagraphChild(0).Click 5,5
			End If

			Set objStaticText=Nothing
			Set objNewParagraphChild=Nothing

		End If

		If sProjectNames<>"" Then
			'Changing label  of Stpes to "Assign to Project"
			objDialogNewPara.JavaStaticText("Stpes").SetTOProperty "label","Assign to Project"
			'Clicking "Assign to Project" Static Text
			objDialogNewPara.JavaStaticText("Stpes").Click 5,5
			wait(5)
			'Selecting Project Names and Shifting to Selected Project List
			iItemCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_ParagraphDetailsCreate",objDialogNewPara.JavaList("ProjectForSelect"),"items count")			
			aListItem=Split(sProjectNames,":")
			objDialogNewPara.JavaList("ProjectForSelect").Select aListItem(0)
				For iCnt=1 to Ubound(aListItem)
					objDialogNewPara.JavaList("ProjectForSelect").ExtendSelect aListItem(iCnt)
				Next
			Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "AddProject")
		End If
		
        If IsArray(aDefineOpt)=True Then
			'Changing label  of Stpes to "Define Options"
			objDialogNewPara.JavaStaticText("Stpes").SetTOProperty "label","Define Options"
			'Clicking "Define Options"
        	objDialogNewPara.JavaStaticText("Stpes").Click 5,5
        	wait(5)
			If aDefineOpt(0)<>"" Then
				'Setting  "ShowAsNwRt" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "ShowAsNwRt",aDefineOpt(0))
			End If
			If aDefineOpt(1)<>"" Then
				'Setting  "UsItIdentifierAs" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "UsItIdentifierAs",aDefineOpt(1))
			End If
			If aDefineOpt(2)<>"" Then
				'Setting  "UsRevIdentifier" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "UsRevIdentifier",aDefineOpt(2))
			End If
			If aDefineOpt(3)<>"" Then
				'Setting  "UsRevIdentifier" to ON or OFF
                Call Fn_CheckBox_Set("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "ChkOutItmRevOnCr",aDefineOpt(3))
			End If
		End If

		wait(2)
		objDialogNewPara.JavaButton("Finish").WaitProperty "enabled", 1, 20000
        'Click on "Finish" button
    	Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "Finish") 
		Fn_ReqMgr_ParagraphDetailsCreate = sParagraphID & "-" & sRevId
		Call Fn_ReadyStatusSync(1)

		'Click on Close button
		Call Fn_Button_Click("Fn_ReqMgr_ParagraphDetailsCreate", objDialogNewPara, "Close")
						
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Paragraph of ID [" + CStr(sParagraphID) +"-"+ sParaName + "]")
	Set objDialogNewPara=Nothing
	Set objSelectType=Nothing
	Set objDialog=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------Function to Handle DataPanel Tab-------------------------------------------------------------------------------
'Function Name		:				Fn_ReqMgr_DataPanalPropertiesOperations

'Description		:		 		The function is used to perform operations on the property in DataPanel Tab


'Parameters			:	 			1.  sAction		- Action To Perform	
'												2.  sPropertyName	- Property Name
'												3.  sPropertyValue	- Property Value

'Return Value		: 				True/False

'Pre-requisite		:		 		NA.
													'Note : If we need to check "TraceLink" is blank that time function will return False
													'(Means when TraceLink List does not contain Item that time function will return False)

'Examples			:				Call Fn_ReqMgr_DataPanalPropertiesOperations("Verify", "TraceLink", "002529/A;1-Item2")-For multiple property value check we are using ~ Seperator
'												Call Fn_ReqMgr_DataPanalPropertiesOperations("Verify", "HasTraceLink", "False")
'												Call Fn_ReqMgr_DataPanalPropertiesOperations("Verify", "HasTraceLink", "Y")
'												Call Fn_ReqMgr_DataPanalPropertiesOperations("Verify", "CheckOut", "")
'												Call Fn_ReqMgr_DataPanalPropertiesOperations("EditCheckOutSave", "Revision", "NewRevisionName")
'												Call Fn_ReqMgr_DataPanalPropertiesOperations("EditCheckOutSave", "Name", "NewName")
'												Call Fn_ReqMgr_DataPanalPropertiesOperations("EditCheckOutVerify", "Checked-Out","Y")
'History			:				Developer Name			Date			Rev. No.			Changes Done			Reviewer		Modified By		
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									 		Sandeep N				15-June-10			1.0																Tushar B			
'																				24-June-10			1.1																Tushar B		Sandeep N
'											Added New Case "EditCheckOutSave"
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Public Function Fn_ReqMgr_DataPanalPropertiesOperations(sAction,sPropertyName,sPropertyValue)
GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_DataPanalPropertiesOperations"
Dim bFlag,iItemCnt,sRValue,aPropValue,iCounter,iCnt,sChkOutValue,ChkValue
Dim objRMWindowApplet
	'Setting Label to Properties 
	JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName").SetTOProperty "label","Properties"
	'Checking Properties Panel is Exist or Not
	 If Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName"))=False Then
		'Selecting Data Panel
		Call Fn_MenuOperation("Select","View:Show Data Panel")
		wait(5)
	 End If
		JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaStaticText("SpecTabName").Click 1,1
		wait(5)
	  'Creating Object of RMWindowApplet 
	  Set objRMWindowApplet=Fn_UI_ObjectCreate("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))
	  'Selecting Case
	  Select Case sAction
		Case "Verify"
			 'Setting Label to All
			  objRMWindowApplet.JavaStaticText("DtPnlPropBottomLink").SetTOProperty "label","All"
			  'Clicking on All 
			  objRMWindowApplet.JavaStaticText("DtPnlPropBottomLink").Click 1,1
			  wait(5)
			 'Checking All properties are show on panel or Not
			 objRMWindowApplet.JavaStaticText("ShowEmptyProp").SetTOProperty "label","Show empty properties..."
			 wait(5)
			If Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet.JavaStaticText("ShowEmptyProp"))=False  Then
				  'Clicking on ShowEmptyProperties
				   objRMWindowApplet.JavaStaticText("ShowEmptyProp").SetTOProperty "label","Hide empty properties..."
				   objRMWindowApplet.JavaStaticText("ShowEmptyProp").Click 1,1
				   wait(5)
			End If
				Select Case sPropertyName
						Case "HasTraceLink"
							'Taking Value of HasTraceLink Radio button
							objRMWindowApplet.JavaStaticText("ShowEmptyProp").SetTOProperty "label","Has Tracelink:"
							wait(5)
							objRMWindowApplet.JavaStaticText("ShowEmptyProp").SetTOProperty "label","Show empty properties..."
							If Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet.JavaStaticText("ShowEmptyProp"))=True  Then
								'sRValue=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet.JavaRadioButton("HasTraceLink"),"value")
								'Matching Value With sPropertyValue parameter
								'If Cstr(Cbool(sRValue))=sPropertyValue Then
								Fn_ReqMgr_DataPanalPropertiesOperations=True 
								'End If
							Else
								If sPropertyValue="False" Then
									Fn_ReqMgr_DataPanalPropertiesOperations=True
								Else
									Fn_ReqMgr_DataPanalPropertiesOperations=False         
								End If
							End If	
						Case "TraceLink"
								bFlag=False
								If  Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet.JavaList("TraceLink"))=True Then
									 'Taking Item Count of "TraceLink" List
									 iItemCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet.JavaList("TraceLink"),"items count")
									 'Spliting PropertyValue To check multiple values
									   aPropValue=Split(sPropertyValue,"~")
										For iCounter=0 To Ubound(aPropValue)
												For iCnt=0 To iItemCnt-1
														If objJavaApplet.JavaList("TraceLink").GetItem(iCnt)=aPropValue(iCounter) Then
															Fn_ReqMgr_DataPanalPropertiesOperations=True
															bFlag=True
															Exit For
														End If
												Next
													If bFlag=False Then
														Fn_ReqMgr_DataPanalPropertiesOperations=False
														Exit For
													End If
										Next
								Else
									Fn_ReqMgr_DataPanalPropertiesOperations=False
								End If
                             Case "CheckOut"
								'Checking Existance of PropChecked-Out Edit Box
								If	Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet.JavaEdit("PropChecked-Out"))=True Then
									sChkOutValue=Fn_Edit_Box_GetValue("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet,"PropChecked-Out")
											If Trim(sChkOutValue)=Trim(sPropertyValue) Then
												Fn_ReqMgr_DataPanalPropertiesOperations=True
											Else
												Fn_ReqMgr_DataPanalPropertiesOperations=False
										   End If
								Else
											If sPropertyValue="" Then
												Fn_ReqMgr_DataPanalPropertiesOperations=True
											Else
												Fn_ReqMgr_DataPanalPropertiesOperations=False
											End If
								End If
				End Select
		Case "EditCheckOutSave"
				Fn_ReqMgr_DataPanalPropertiesOperations=False
			 'Setting Label to All
			  objRMWindowApplet.JavaStaticText("DtPnlPropBottomLink").SetTOProperty "label","General"
			  'Clicking on All 
			  objRMWindowApplet.JavaStaticText("DtPnlPropBottomLink").Click 1,1
			  wait(10)
				If Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet.JavaButton("CheckOtAndEdit"))=True Then
					Call Fn_Button_Click("Fn_ReqMgr_DataPanalPropertiesOperations", objRMWindowApplet,"CheckOtAndEdit")
					wait(10)
					 If Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Check-Out"))=True Then
						Call Fn_Button_Click("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Check-Out"),"Yes")
					Else
						Exit Function
					End If
				End If


            	Select Case sPropertyName
					'Modifing Name Property
						Case "Name"
							If sPropertyValue<>"" Then
                                Call Fn_Edit_Box("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"PropName",sPropertyValue)
								JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("CheckOtAndEdit").SetTOProperty "label","Save"
								wait(20)
								Call Fn_Button_Click("Fn_ReqMgr_DataPanalPropertiesOperations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"CheckOtAndEdit")
								Fn_ReqMgr_DataPanalPropertiesOperations=True
							End If
						'Modifing Revision Property
						 Case "Revision"
							If sPropertyValue<>"" Then
								JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaEdit("PropName").SetTOProperty "attached text","Revision:"
								wait(5)
								Call Fn_Edit_Box("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"PropName",sPropertyValue)
                                JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaButton("CheckOtAndEdit").SetTOProperty "label","Save"
								wait(20)
								Call Fn_Button_Click("Fn_ReqMgr_DataPanalPropertiesOperations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"CheckOtAndEdit")
								Fn_ReqMgr_DataPanalPropertiesOperations=True
							End If
					End Select
			Case "EditCheckOutVerify"
				Fn_ReqMgr_DataPanalPropertiesOperations=False
			 'Setting Label to All
			  objRMWindowApplet.JavaStaticText("DtPnlPropBottomLink").SetTOProperty "label","General"
			  'Clicking on All 
			  objRMWindowApplet.JavaStaticText("DtPnlPropBottomLink").Click 1,1
			  wait(5)
				If Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",objRMWindowApplet.JavaButton("CheckOtAndEdit"))=True Then
					Call Fn_Button_Click("Fn_ReqMgr_DataPanalPropertiesOperations", objRMWindowApplet,"CheckOtAndEdit")
					wait(5)
					 If Fn_UI_ObjectExist("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Check-Out"))=True Then
						Call Fn_Button_Click("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Check-Out"),"Yes")
						wait(5)
					Else
						Exit Function
					End If
				End If

            	Select Case sPropertyName
					'Modifing Name Property
					Case "Checked-Out"
                        			
						'Setting Label to All
						  objRMWindowApplet.JavaStaticText("DtPnlPropBottomLink").SetTOProperty "label","All"
						  'Clicking on All 
						  objRMWindowApplet.JavaStaticText("DtPnlPropBottomLink").Click 1,1
						  wait(10)


							JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaEdit("PropName").SetTOProperty "attached text","Checked-Out:"
							wait(10)
                            ChkValue=Fn_Edit_Box_GetValue("Fn_ReqMgr_DataPanalPropertiesOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"PropName")
							If sPropertyValue=ChkValue Then
								Fn_ReqMgr_DataPanalPropertiesOperations=True
							End If						
				End Select

	End Select
Set objRMWindowApplet=Nothing
End Function

'______________________________________________________________________________________________________________________________________________


'*********************************************************		Function to perform operation on Lower ReqMgr Table in Requirement Manager		***********************************************************************

'Function Name		:				Fn_ReqMgr_LowerRMTableNodeOpeations

'Description			 :		 		This function is used to Perform operations on  the Lower ReqMgr Table 

'Parameters			   :				1.	strAction = e.g-"Select"
'													2. StrNodeName:Name of the Node. 
'												3. strColName
'												4. strColValue
'												5. strPopupMenu
				
											
'Return Value		   : 				True/ False

'Pre-requisite			:				Lower RM Table window should be displayed .

'Examples				:		Fn_ReqMgr_LowerRMTableNodeOpeations("VerifyTable","000221A;1-Test","","","")-Pass First Cell Data for this Case
'											Fn_ReqMgr_LowerRMTableNodeOpeations("VerifyNode","000138/A;1-test (View):REQ-000025/A;1-Test","","","")
'											Fn_ReqMgr_LowerRMTableNodeOpeations("Select","000138/A;1-test (View):REQ-000025/A;1-Test","","","")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep					25-June-2010		1.0																	Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
Public Function Fn_ReqMgr_LowerRMTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_LowerRMTableNodeOpeations"
    Dim iRowNo,IntRows,IntCounter,StrIndex,StrNodePath
	Dim ObjTable,ObjTableChilds,ObjLwrTable
    Fn_ReqMgr_LowerRMTableNodeOpeations=False
	
	Set ObjTable=Description.Create()
			ObjTable("Class Name").value="JavaTable"
	Set ObjTableChilds=JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").ChildObjects(ObjTable)
	'Checking Lower RMT Table is Fist or on
	If ObjTableChilds.count <= 1 Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Lower RMT Table Is Not Exist")
		Exit Function
	End If
	If strNodeName<>"" Then
			'Get the No. of rows present in the ReqMgr Table
			IntRows = ObjTableChilds(1).GetROProperty("rows")
			Set ObjLwrTable =ObjTableChilds(1).Object
			'Format the Inout as per Table Default Nodes
			StrNodeName = Replace(StrNodeName, ":", ", ")
	
			'Get the Row No. of required Node
			For IntCounter = 0 to IntRows -1
				StrNodePath = ObjLwrTable.getPathForRow(IntCounter).toString
				StrNodePath = Right(StrNodePath, (Len(StrNodePath)-Instr(1, StrNodePath, ",", 1)))					
				StrNodePath = Left(StrNodePath, Len(StrNodePath)-1)
			
				If Trim(StrNodePath) = Trim(StrNodeName) Then
					StrIndex = Cstr(IntCounter)
					Exit For
				End If
			Next
			If Cint(IntCounter) = Cint(IntRows) Then
                StrIndex = "FAIL:Node Not Found"
			End If
	End If
   
	Select Case strAction
		'Verifying Table is Exist or Not 
		Case "VerifyTable"
				If Instr(1,ObjTableChilds(1).GetCellData(0,0),strNodeName)>0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"Successfully Verified Lower RMT Table Is Exist")
					Fn_ReqMgr_LowerRMTableNodeOpeations=True
				End If
	   Case "VerifyNode"
				iRowNo =StrIndex
					If isNumeric(iRowNo) then
						Fn_ReqMgr_LowerRMTableNodeOpeations=True
					End if
	   Case "Select"
				iRowNo =StrIndex
					If isNumeric(iRowNo) then
						ObjTableChilds(1).SelectRow iRowNo
						Fn_ReqMgr_LowerRMTableNodeOpeations=True
					End if
	End Select
	Set ObjLwrTable = Nothing
	Set ObjTableChilds=Nothing
	Set ObjTable=Nothing
End Function
'______________________________________________________________________________________________________________________________________________
'*********************************************************		Function to perform operations on TraceabilityReport Table Column*************************************
'Function Name		:			Fn_TraceabilityReportColumnOperations

'Description			 :		 	  Function to perform operations on TraceabilityReport Table Column

'Parameters			   :	 			1.strTableName: Table Name
															'|-DefiningTable
															'|-ComplyingTable
'													 2.strAction: Action Name
															'|-InsertColumn
															'|-RemoveColumn
'													 3.sColName: Column Name on which have to right click
															'|-Group ID
															'|-Release Status
															'|-Type
'													 4.sNewColName:New column Name which have to add in report
															'|-2D Spanshot
															'|-Release Status
'													 5.DsplNameOpt: Displayable Name option
'															'|-"ON" Or "OFF" OR ""
															
'Return Value		   : 			True Or False

'Pre-requisite			:		 	Object Should be Slected on which have to perform the operations

'Examples				:			Fn_TraceabilityReportColumnOperations("DefiningTable","RemoveColumn","Group ID","","")	
												'Fn_TraceabilityReportColumnOperations("DefiningTable","InsertColumn","Type","Release Status","On")
												'Fn_TraceabilityReportColumnOperations("DefiningTable","ColumnExist","Release Status","","")
'History					 :			
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'													Sandeep N										   				07/07/2010			           1.0																						Tushar B
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_TraceabilityReportColumnOperations(strTableName,strAction,sColName,sNewColName,DsplNameOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_TraceabilityReportColumnOperations"
   'Declaring Variables
	Dim ObjReportDialog
	Dim iCount,iCnt,strColumnName,bFlag
	'Function Returning False
	Fn_TraceabilityReportColumnOperations=False
	'Setting bFlag to False
	bFlag=False
	'Checking Existance of Traceability Report...
    If Fn_UI_ObjectExist("Fn_TraceabilityReportColumnOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report"))=False Then
		'Invoking Traceability Report
   	   Call Fn_MenuOperation("Select","Tools:Trace Link:Traceability Report")
   End If
   'Creating object  of Traceability Report...
   Set ObjReportDialog=Fn_UI_ObjectCreate("Fn_TraceabilityReportColumnOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report"))
   'Cases for Table
   Select Case strTableName
	 	Case "DefiningTable"
            'Taking total count of columns
			iCount=Fn_UI_Object_GetROProperty("Fn_TraceabilityReportColumnOperations",ObjReportDialog.JavaTable("DefiningTable"),"cols")
			For iCnt=0 To iCount-1
				'Taking column name
				strColumnName=ObjReportDialog.JavaTable("DefiningTable").GetColumnName(iCnt)
				'Checking user pass correct column name
				If strColumnName=sColName Then
					If strAction<>"ColumnExist" Then
					'Right clicking on column sColName
					ObjReportDialog.JavaTable("DefiningTable").SelectColumnHeader sColName,"RIGHT"
					End If
					'Setting bFlag to False
					bFlag=True
					Exit For

				 End If
			Next
		Case "ComplyingTable"
			'Taking total count of columns
			iCount=Fn_UI_Object_GetROProperty("Fn_TraceabilityReportColumnOperations",ObjReportDialog.JavaTable("ComplyingTable"),"cols")
			'Taking column name			
			For iCnt=0 To iCount-1
				strColumnName=ObjReportDialog.JavaTable("ComplyingTable").GetColumnName(iCnt)
				'Checking user pass correct column name
				If strColumnName=sColName Then
					If strAction<>"ColumnExist" Then
					ObjReportDialog.JavaTable("ComplyingTable").SelectColumnHeader sColName,"RIGHT"
					'Setting bFlag to False
					End If
					bFlag=True
					Exit For
				 End If
			Next
		Case Else
			Exit Function
   End Select
   Select Case strAction
		Case  "RemoveColumn"
				If  bFlag=True Then
					'Selecting Remove this column menu
					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report").JavaMenu("label:=Remove this column").Select
					'Clicking on yes button to confirm remove of column
					Call Fn_Button_Click("Fn_TraceabilityReportColumnOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Remove Column"),"Yes")
                	Fn_TraceabilityReportColumnOperations=True
				Else
					Exit Function
			   End If
		Case  "InsertColumn"
				If  bFlag=True Then
					'Selecting Insert column\(s\) menu
					JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Traceability Report").JavaMenu("label:=Insert column\(s\) \.\.\.").Select
					If DsplNameOpt<>"" Then
						'Setting Desplayable Name option
						Call Fn_CheckBox_Set("Fn_TraceabilityReportColumnOperations", JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Change Columns"), "UseDisplayableName", DsplNameOpt)
					End If
						'Selecting Column Name from column list
						JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Change Columns").JavaList("ListAvailableCols").Select sNewColName
						'Adding Column
						Call Fn_Button_Click("Fn_TraceabilityReportColumnOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Change Columns"),"Add")
						'Applying Changes
						Call Fn_Button_Click("Fn_TraceabilityReportColumnOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Change Columns"),"Apply")
						'Closing window
						Call Fn_Button_Click("Fn_TraceabilityReportColumnOperations",JavaWindow("DefaultWindow").JavaWindow("TcDefaultApplet").JavaDialog("Change Columns"),"Cancel")
                        Fn_TraceabilityReportColumnOperations=True
				End If
		 Case  "ColumnExist"
						If bFlag=True Then
							Fn_TraceabilityReportColumnOperations=True
						 End If
		Case  Else
					Exit Function
				
	End Select
	'Closing TraceabilityReport Report
	Call Fn_Button_Click("Fn_TraceabilityReportColumnOperations",ObjReportDialog,"Ok")
	Set ObjReportDialog=Nothing
End Function
'______________________________________________________________________________________________________________________________________________

'*********************************************************		Function to preform operation on Open Specification by Name List***************************************
'Function Name		:				Fn_ReqMgr_OpenSpecByNameOperations

'Description			 :		 		Perform operations on operation on Open Specification by Name Table

'Parameters			   :	 			1.strAction: Action to perform(Select or Exist)
'													 2.strSearchName: Name for search
													'3.strCellValue:DataTable Cell Value
'Return Value		   : 				True Or False

'Pre-requisite			:		 		Should be logged in & present on Requirement Manager perspective

'Examples				:				 Fn_ReqMgr_OpenSpecByNameOperations("DoubleClickCell","Rew","00255-Rewa")
'													
'													IMP NOTE :-- To use this function First clear the cache and relaunch the Teamcenter
'													Fn_ReUserTcSession(True,True, Environment.Value("TcUser1"))
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'												------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   09/07/2010			          1.0										Created								Tushar B
'												-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_OpenSpecByNameOperations(strAction,strSearchName,strCellValue)
GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_OpenSpecByNameOperations"
'Declaring Object
Dim objRMWindowApplet,bFlag,iRowCount,sCellData,iCounter
	bFlag=False
    Fn_ReqMgr_OpenSpecByNameOperations=False
	'Checking Existance of  "RMWindowApplet" Applet  : Its Prerequisite
	If Fn_UI_ObjectExist("Fn_ReqMgr_OpenSpecByNameOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))=False Then
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"RMWindowApplet applet window is Not present")
		Exit Function
	End If
	'Creating Object of  "RMWindowApplet" Applet
	Set objRMWindowApplet=Fn_UI_ObjectCreate("Fn_ReqMgr_OpenSpecByNameOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))
    Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ReqMgr_OpenSpecByNameOperations",objRMWindowApplet.JavaCheckBox("MostRecentlyUsedButton"),"attached text","openstructurebyname_16")
	'Clicking "OpenStructureByName" Button OpenStructureByName Table
	Call Fn_CheckBox_Set("Fn_ReqMgr_OpenSpecByNameOperations",objRMWindowApplet, "MostRecentlyUsedButton",  "ON" )
	'Setting Name for search
    Call Fn_Edit_Box("Fn_ReqMgr_OpenSpecByNameOperations",objRMWindowApplet,"openSpecbyNameEdit",strSearchName)
	'Clicking on Find button to find object
	Call Fn_Button_Click("Fn_ReqMgr_OpenSpecByNameOperations",objRMWindowApplet,"openSpecbyNameFind")
	'If user put wrong search then below code handle that scenario and exit the function and function will return False
	'Checking Existance of error dailog
	JavaWindow("RequirementsManager").JavaWindow("Open Text Error").SetTOProperty "title","Nothing found!"
	wait(10)
    If  Fn_UI_ObjectExist("Fn_ReqMgr_OpenSpecByNameOperations",JavaWindow("RequirementsManager").JavaWindow("Open Text Error"))=True Then
		'Handling Error dialog
		Call Fn_ReqMgr_ErrorMessageVerify("Nothing found!","No objects found")
		JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaObject("CloseMRUList").Click 1,1
		Exit Function
	End If
	'Taking count of rows present in openSpecbyNameTable
	iRowCount=Fn_UI_Object_GetROProperty("Fn_ReqMgr_OpenSpecByNameOperations",objRMWindowApplet.JavaTable("openSpecbyNameTable"),"rows")
    
	For iCounter=0 To iRowCount-1
		'Taking cell Data from table
        sCellData=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_OpenSpecByNameOperations",objRMWindowApplet,"openSpecbyNameTable",iCounter,0)
		'Verifying data 
		If strCellValue=sCellData Then
			bFlag=True
			Exit For
		End If
	Next
	If bFlag=False Then
		Exit Function
	End If
	'Selecting Case
	Select Case strAction
		'To select list Item
		Case "DoubleClickCell"
			'Selecting a Row
			JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("openSpecbyNameTable").SelectRow(iCounter)
			'Double Clicking on Cell
            Call Fn_UI_JavaTable_DoubleClickCell("Fn_ReqMgr_OpenSpecByNameOperations", objRMWindowApplet, "openSpecbyNameTable",iCounter,0,"","")
	End Select
	'Function Returning True
	Fn_ReqMgr_OpenSpecByNameOperations=True
	'Setting Object to Nothing
	Set objRMWindowApplet=Nothing
End Function

'-------------------------------------------------------------------------Function to perform operations on Static Text----------------------------------------------------------------------------------
'Function Name		:			Fn_ReqMgr_StaticTextOperations

'Description		:			This Function is used to perform operations on static text

'Parameters			:			1.	strAction:Action to perform
'											  2.	strStaticText:Static text on which perform oprshn

'Return Value		:			True/False

'Pre-requisite		:			Static Text should be displayed

'Examples			:			
										'Fn_ReqMgr_StaticTextOperations("Exist","Import a specification")
										'Fn_ReqMgr_StaticTextOperations("Click","Import a specification")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N				09-Jully-2010		1.0										                       Tushar B	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_StaticTextOperations(strAction,strStaticText)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_StaticTextOperations"
   'Declaring the variables
    Dim ObjRMWindowApplet
	'Setting function to false
	Fn_ReqMgr_StaticTextOperations=False
	'Checking existance of RMWindowApplet
	If Fn_UI_ObjectExist("Fn_ReqMgr_StaticTextOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))=False Then
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"RMWindowApplet applet window is Not present")
		Exit Function
	End If
	'Creating object of RMWindowApplet
	Set ObjRMWindowApplet=Fn_UI_ObjectCreate("Fn_ReqMgr_StaticTextOperations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))
	'Changing Label of ImportASpecification
	Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ReqMgr_StaticTextOperations",ObjRMWindowApplet.JavaStaticText("ImportASpecification"),"Label",strStaticText)

	Select Case strAction
		Case "Exist"
			'checking existance static text
			Call Fn_Java_StaticText_Exist("Fn_ReqMgr_StaticTextOperations",ObjRMWindowApplet,"ImportASpecification")
	Case "Click"
		Call Fn_UI_JavaStaticText_Click("Fn_ReqMgr_StaticTextOperations", ObjRMWindowApplet,"ImportASpecification", 1, 1,"LEFT")
	Case Else
			Exit Function
	End Select
	Fn_ReqMgr_StaticTextOperations=True
	Set ObjRMWindowApplet=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------Function to Save as Item Revision -------------------------------------------------------------------------------------------
'Function Name		:				Fn_ReqMgr_SaveAsItemRevision

'Description			 :		 		 This function Save As the Item Revision of Existing item

'Parameters			   :	 			1.sItemID:Item id
'													 2.sItemRevID: Item Revision Id
'													3.sItemName:Item Name
'													4.sItemDesc:Description of the Item
'													5.sItemUOM:

'Return Value		   : 			Item Id  / Revision Id

'Pre-requisite			:		 Should be logged in & present on Requirement Manager perspective and Req spec or Req Or Para should be selected of which have to save as

'Examples				:				Fn_ReqMgr_SaveAsItemRevision("","","","","")
'													Fn_ReqMgr_SaveAsItemRevision("000248","A","ItemName","Revision","")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   12/07/2010			              1.0										Created						Tushar B
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_SaveAsItemRevision(sItemID,sItemRevID,sItemName,sItemDesc,sItemUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_SaveAsItemRevision"
    Dim strItemID, strItemRevID
	Dim ObjSaveAsDialog,objSelectType,objDialog
	Fn_ReqMgr_SaveAsItemRevision=False
	'Checking Existance of Save As Dialog
	If Fn_UI_ObjectExist("Fn_ReqMgr_SaveAsItemRevision",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("SaveAs"))=False Then
		'Opening  Save As Dialog
         Call Fn_MenuOperation("Select","File:Save As...:Item\(Revision\)...")
	End If
    'Check the existence of  Save As Dialog window
	Set ObjSaveAsDialog=Fn_UI_ObjectCreate("Fn_ReqMgr_SaveAsItemRevision",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("SaveAs"))
	If sItemID <> "" Then
		'Setting  Item Id
        Call Fn_Edit_Box("Fn_ReqMgr_SaveAsItemRevision", ObjSaveAsDialog,"ItemID", sItemID)
	End If	
	If sItemRevID <> "" Then
		'Setting Item Revision ID
        Call Fn_Edit_Box("Fn_ReqMgr_SaveAsItemRevision",ObjSaveAsDialog,"Revision", sItemRevID)
	End If	
	If  sItemID = "" or sItemRevID = "" Then
		'click on assign button
		  Call Fn_Button_Click("Fn_ReqMgr_SaveAsItemRevision", ObjSaveAsDialog, "Assign")	  
	End If	
	'Extract Creation data
	strItemID =Fn_Edit_Box_GetValue("Fn_ReqMgr_SaveAsItemRevision", ObjSaveAsDialog,"ItemID")
	strItemRevID = Fn_Edit_Box_GetValue("Fn_ReqMgr_SaveAsItemRevision",ObjSaveAsDialog,"Revision")
	If  sItemName<>"" Then
		'Set Item name
		 Call Fn_Edit_Box("Fn_ReqMgr_SaveAsItemRevision", ObjSaveAsDialog,"Name",sItemName)
	End If
	'Set description
	Call Fn_Edit_Box("Fn_ReqMgr_SaveAsItemRevision", ObjSaveAsDialog,"Description",sItemDesc)
	'Set UOM
		If sItemUOM <> "" Then
			 Set objSelectType=description.Create()
			objSelectType("Class Name").value = "JavaStaticText"
			objSelectType("label").value = sItemUOM
			ObjSaveAsDialog.JavaButton("UOM").Click
			Set objDialog =ObjSaveAsDialog.ChildObjects(objSelectType)
			objDialog(0).Click 5, 5, "LEFT"
	   End If
	   'Click on Finish button
       Call Fn_Button_Click("Fn_ReqMgr_SaveAsItemRevision", ObjSaveAsDialog, "Finish")
		Fn_ReqMgr_SaveAsItemRevision = strItemID & "-" & strItemRevID
		'Checking Existance of Save As Dialog
		If Fn_UI_ObjectExist("Fn_ReqMgr_SaveAsItemRevision",ObjSaveAsDialog)=True Then
				'Closing Save As Dialog
				Call Fn_Button_Click("Fn_ReqMgr_SaveAsItemRevision",ObjSaveAsDialog, "Close")
		End If
		
        Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created Specification of ID [" + CStr(strItemId) + "]")
		Set ObjSaveAsDialog=Nothing
		Set objSelectType=Nothing
		Set objDialog=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------Function to create Para Or Req from quick panel--------------------------------------------------------------------
'Function Name		:			Fn_ReqMgr_QuickPanelOperation

'Description			 :		 This function is used to create Requirement or Paragraph Through Quick Panel

'Parameters			   :	 		1.sName:Name of Requirement or Paragraph (Mandetory Parameter)
'												 2.sType: Type(Mandetory Parameter)
													'1.Requirement
													'2.Paragraph
'												 3.sChildOpt:"ON" OR "OFF"
'												

'Return Value		   : 	True/False

'Pre-requisite			:		 Should be logged in & present on Requirement Manager perspective and Quick panel has to display

'Examples				:		Fn_ReqMgr_QuickPanelOperation("Test","Paragraph","OFF")
'											Fn_ReqMgr_QuickPanelOperation("Test","Requirement","ON")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   12/07/2010			              1.0										Created						Tushar B
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_QuickPanelOperation(sName,sType,sChildOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_QuickPanelOperation"
   Dim ObjRMWindowApplet
	Fn_ReqMgr_QuickPanelOperation=False
	If Fn_UI_ObjectExist("Fn_ReqMgr_QuickPanelOperation",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))=False Then
        Exit Function
	End If
	Set ObjRMWindowApplet=Fn_UI_ObjectCreate("Fn_ReqMgr_QuickPanelOperation",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"))
	'Setting Name to Req Or Para
	Call Fn_Edit_Box("Fn_ReqMgr_QuickPanelOperation",ObjRMWindowApplet,"QuickCreatePnlName",sName)
	'Selecting Type
	Call Fn_List_Select("Fn_ReqMgr_QuickPanelOperation",ObjRMWindowApplet,"QuickCreatePnlType",sType)
	'JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaList("QuickCreatePnlType").Select sType
    If  sChildOpt <> ""Then
		'Setting Child option "ON" Or "OFF"
		Call Fn_CheckBox_Set("Fn_ReqMgr_QuickPanelOperation",ObjRMWindowApplet,"QuickCreatePnlChild",sChildOpt)
	End If
    'Clicking on create button to create Para Or Req
	Call Fn_Button_Click("Fn_ReqMgr_QuickPanelOperation",ObjRMWindowApplet,"QuickCreatePnlCreate")
	Fn_ReqMgr_QuickPanelOperation=True
	Set ObjRMWindowApplet=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------------Function to Customize Menu's--------------------------------------------------------------------
'Function Name		:			Fn_ReqMgr_CustomizeIWantTo

'Description			 :		 This function is used to Customize Menu's

'Parameters			   :	 		1.strAction:Action Name
													'1."Add"
													'2."Remove"
													'3.VerifyEntries
'												 2.strEntryNode: Node Name
													'1.In "Add" Case Node Name is From "Available Entries" Tree (:) Seperated
													'2.In "Remove" And "VerifyEntries" Case Node Name is From "Selected Entries" Table

											'Compulsory
'											IMP Note : To Use this function Clear the cache first and Launch New Application	
											'Fn_ReUserTcSession(True, True, Environment.Value("TcUser1"))
'Return Value		   : 	True/False

'Pre-requisite			:	Should be logged in & present on Requirement Manager perspective

											'IMP Note : To Use this function Clear the cache first and Launch New Application	
											'Fn_ReUserTcSession(True, True, Environment.Value("TcUser1"))

'Examples				:		Fn_ReqMgr_CustomizeIWantTo("Add","View:Show Data Panel")
'											Fn_ReqMgr_CustomizeIWantTo("Remove","Show Data Panel")
'											Fn_ReqMgr_CustomizeIWantTo("VerifyEntries","Show Data Panel")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   13/07/2010			              1.0										Created						Tushar B
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_ReqMgr_CustomizeIWantTo(strAction,strEntryNode)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_CustomizeIWantTo"
   'Declaring Variables
	Dim ObjReqMgrIWantTo,bFlag,sEntryNames,iCount,iItemCount,iCnt,sCellData
	'Initially Setting function to False
	Fn_ReqMgr_CustomizeIWantTo=False
  
	'Checking Existance "ReqMgrIWantTo" Window
	If Fn_UI_ObjectExist("Fn_ReqMgr_CustomizeIWantTo",JavaWindow("RequirementsManager").JavaWindow("ReqMgrIWantTo"))=False Then
		'Changing Label of "IWantTo" static text to "History"
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ReqMgr_CustomizeIWantTo",JavaWindow("RequirementsManager").JavaStaticText("IWantTo..."),"label","History")
		'Clicking on "History"
		Call Fn_UI_JavaStaticText_Click("Fn_ReqMgr_CustomizeIWantTo", JavaWindow("RequirementsManager"),"IWantTo...",1,1, "")


		'Changing Label of "IWantTo" static text to "Open Items"
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ReqMgr_CustomizeIWantTo",JavaWindow("RequirementsManager").JavaStaticText("IWantTo..."),"label","Open Items")
		'Clicking on "Open Items"
		Call Fn_UI_JavaStaticText_Click("Fn_ReqMgr_CustomizeIWantTo", JavaWindow("RequirementsManager"),"IWantTo...",1,1, "")
		
		'Changing Label of "IWantTo" static text to "Favorites"
		Call Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ReqMgr_CustomizeIWantTo",JavaWindow("RequirementsManager").JavaStaticText("IWantTo..."),"label","Favorites")
		'Clicking on "Favorites"
		Call Fn_UI_JavaStaticText_Click("Fn_ReqMgr_CustomizeIWantTo", JavaWindow("RequirementsManager"),"IWantTo...",1,1, "")
		'Clicking on Customize Toolbar button to Open "ReqMgrIWantTo" Window
		Call Fn_ToolbarButtonClick_Ext(2,"Customize")
	End If
	'Creating Object "ReqMgrIWantTo" Window
	Set ObjReqMgrIWantTo=Fn_UI_ObjectCreate("Fn_ReqMgr_CustomizeIWantTo",JavaWindow("RequirementsManager").JavaWindow("ReqMgrIWantTo"))
   Select Case strAction
    Case "Add" 'Case to Add Entries
		'Selecting Item from Available Entries Tree
		Call Fn_JavaTree_Select("Fn_ReqMgr_CustomizeIWantTo", ObjReqMgrIWantTo,"Available Entries",strEntryNode)
		'Click on Plus button to Add Entry
        Call Fn_Button_Click("Fn_ReqMgr_CustomizeIWantTo", ObjReqMgrIWantTo, "Plus")
	Case "Remove"  'Case to Remove Entries
		'Spliting Node Name
		sEntryNames=Split(strEntryNode,":")
		For iCount=0 To Ubound(sEntryNames)
			bFlag=False
			'Taking Total item Count of Selected Entries Table
            iItemCount=Fn_UI_Object_GetROProperty("Fn_ReqMgr_CustomizeIWantTo",ObjReqMgrIWantTo.JavaTable("Selected Entries"),"rows")
			For iCnt=0 To iItemCount-1
				'Taking Data from Selected Entries Table
				sCellData=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_CustomizeIWantTo",ObjReqMgrIWantTo, "Selected Entries",iCnt,0)
				'Checking Entry with table data
				If  sEntryNames(iCount)=sCellData Then
					'Selecting Data from table
                    Call Fn_UI_JavaTable_SelectCell("Fn_ReqMgr_CustomizeIWantTo", ObjReqMgrIWantTo, "Selected Entries",iCnt,0)
					'Clicking on Remove button to remove Entry
					Call Fn_Button_Click("Fn_ReqMgr_CustomizeIWantTo", ObjReqMgrIWantTo, "Minus")
					bFlag=True
					Exit For					
				End If
			Next
				If bFlag=False Then
					'If data is not present in table then exit the function
					Call Fn_Button_Click("Fn_ReqMgr_CustomizeIWantTo", ObjReqMgrIWantTo, "OK")
					Set ObjReqMgrIWantTo=Nothing
					Exit Function
				End If
		Next
	
	Case "VerifyEntries"
		sEntryNames=Split(strEntryNode,":")
		For iCount=0 To Ubound(sEntryNames)
			bFlag=False
			'Taking Total item Count of Selected Entries Table
            iItemCount=Fn_UI_Object_GetROProperty("Fn_ReqMgr_CustomizeIWantTo",ObjReqMgrIWantTo.JavaTable("Selected Entries"),"rows")
			For iCnt=0 To iItemCount-1
				'Taking Data from Selected Entries Table
				sCellData=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_CustomizeIWantTo",ObjReqMgrIWantTo, "Selected Entries",iCnt,0)
				If  sEntryNames(iCount)=sCellData Then
                    bFlag=True
					Exit For					
				End If
			Next
		Next
		If bFlag=False Then
			'If data is not present in table then exit the function
			Call Fn_Button_Click("Fn_ReqMgr_CustomizeIWantTo", ObjReqMgrIWantTo, "OK")
			Set ObjReqMgrIWantTo=Nothing
			Exit Function
		End If
   End Select
   'Setting Function to True
   Fn_ReqMgr_CustomizeIWantTo=True
   'Clicking on Apply button 
	Call Fn_Button_Click("Fn_ReqMgr_CustomizeIWantTo", ObjReqMgrIWantTo, "Apply")
   'Clicking on OK button 
	Call Fn_Button_Click("Fn_ReqMgr_CustomizeIWantTo", ObjReqMgrIWantTo, "OK")
  'Setting object to Nothing
  Set ObjReqMgrIWantTo=Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------Function to set BOM Compare mode--------------------------------------------------------------------------------------------
'Function Name		:			Fn_ReqMgr_BOMCompare

'Description			 :		 	 Function to set BOM Compare mode

'Parameters			   :	 		1.strMode: Mode Type
'												 2.strReportOpt: Report Option 
														'eg."ON" or "OFF" or ""
'						
'Return Value		   : 			True Or False

'Pre-requisite			:		 	 Should be logged in & present on Requirement Manager perspective Panel have to be split

'Examples				:			Fn_ReqMgr_BOMCompare("Single level (with find no)","ON")
'												Fn_ReqMgr_BOMCompare("Multi Level (with find no)","OFF")
'												Fn_ReqMgr_BOMCompare("Lowest Level","")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									  15 /07/2010			              1.0										Created						    Tushar B
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_BOMCompare(strMode,strReportOpt)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_BOMCompare"
    Dim ObjBOMCompareWindow
	Fn_ReqMgr_BOMCompare=False
	'Checking Existance Of ReqMgrBOMCompare Window
    If Fn_UI_ObjectExist("Fn_ReqMgr_BOMCompare",JavaWindow("RequirementsManager").JavaWindow("ReqMgrBOMCompare"))=False Then
		'Select menu [File ->Revise...]
		Call Fn_MenuOperation("Select","Tools:BOM Compare...")
	End If	
	'Check the existence of ReqMgrBOMCompare Window
	Set ObjBOMCompareWindow=Fn_UI_ObjectCreate("Fn_ReqMgr_BOMCompare",JavaWindow("RequirementsManager").JavaWindow("ReqMgrBOMCompare"))
	'Selecting Mode From Mode List
    Call Fn_List_Select("Fn_ReqMgr_BOMCompare", ObjBOMCompareWindow, "Mode",strMode)
	If strReportOpt<>"" Then
		'Selecting Report Option
		Call Fn_CheckBox_Set("Fn_ReqMgr_BOMCompare", ObjBOMCompareWindow, "Report", strReportOpt)
	End If
	'Clicking OK button
    Call Fn_Button_Click("Fn_ReqMgr_BOMCompare", ObjBOMCompareWindow, "OK")
	'Function return True
	Fn_ReqMgr_BOMCompare=True
	'Click on Cancel button to Close ReqMgrBOMCompare Window
	Call Fn_Button_Click("Fn_ReqMgr_BOMCompare", ObjBOMCompareWindow, "Cancel")
	'Setting Object nothing
	Set ObjBOMCompareWindow=Nothing
End Function
'_____________________________________________________________________________________________________________________________________

'-------------------------------------------------------------------------Function for Error Dialog which appears on Import Spec Dialog--------------------------------------------------------------------------------------
'Function Name		:			Fn_ReqMgr_ReqMessageVerify

'Description		:			This Function is used to handle Error Dialog which appears on Import Spec Dialog

'Parameters			:			1.	sDilogName:Error Dialog Box Name
'								2.	sErrorMessage:Expected Error Message

'Return Value		:			True/False

'Pre-requisite		:			Error Dialog should be displayed .
'								Note:- This function not check the error message (Need Improvements)
'Examples			:			
										'Fn_ReqMgr_ReqMessageVerify("Enter keywords","Invalid")
'History:
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N				16-Jully-2010		1.0										    Tushar B	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_ReqMessageVerify(sDialogName,sErrorMessage)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_ReqMessageVerify"
   'Variable Diclaration
	Dim sErrMsg,objStaticText,objErrorDialog,iCnt,bFlag
	Fn_ReqMgr_ReqMessageVerify=False
	'Changind dialog name and checking existance of dialog
	bFlag=Fn_UI_Object_SetTOProperty_ExistCheck("Fn_ReqMgr_ReqMessageVerify",Dialog("Enterkeywords"),"text",sDialogName)	
	'Checking Error Dialog Exist or not 
	If  bFlag=True  Then
'			'Getting the label of  ErrMsg in sErrMsg
'        	Set objStaticText=Description.Create()
'			objStaticText("to_class").value="JavaStaticText"
'			'Taking Child object of present Error Dialog box
'			Set objErrorDialog=Dialog("Enterkeywords").ChildObjects(objStaticText)
'			For iCnt=0 to objErrorDialog.count-1
'				'Checking Error message 
'                sErrMsg=objErrorDialog(iCnt).getROProperty("label")
'				If Instr(1,Lcase(sErrMsg),Lcase(sErrorMessage))>0 Then
'					'Clicking on ok button
					Call Fn_UI_WinButton_Click("Fn_ReqMgr_ReqMessageVerify", Dialog("Enterkeywords"), "OK","","","")
					
					For iCnt=0 To 4
						If Fn_UI_ObjectExist("Fn_ReqMgr_ImportReqSpec",Dialog("Enterkeywords"))=True Then
							Call Fn_UI_WinButton_Click("Fn_ReqMgr_ReqMessageVerify", Dialog("Enterkeywords"), "OK","","","")
						 Else
   							bFlag=False
                        End If
						If bFlag=False Then
							Exit For
							Exit Function
						End If
					Next
					Fn_ReqMgr_ReqMessageVerify=True
'					Exit For
'				Else
'					Fn_ReqMgr_ReqMessageVerify=False
'				End If
'			Next
	Else
		Fn_ReqMgr_ReqMessageVerify=False
	End If
	'Checking "Import Spec" Dialog Is exist or Not 
		If Fn_UI_ObjectExist("Fn_ReqMgr_ImportReqSpec",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Import Spec"))=True Then
			'Closing "Import Spec" Dialog
			Call Fn_Button_Click("Fn_ReqMgr_ReqMessageVerify",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("Import Spec"),"Close")
		End If
	'Set objStaticText=Nothing
	'Set objErrorDialog=Nothing
End Function

'-------------------------------------------------------------------------Function To Perform Operations on Attachments Table------------------------------------------------------------------
'Function Name		:	Fn_ReqMgr_AttachmentTableNodeOpeations

'Description			:	This Function is used to Perform operations on Attachments Table

'Parameters			:			1.	strAction:Action Name 			Eg -> "Select"
'											 2.	 strNodeName:Node name on which have to perform operation
											'3.	strColName:Column Name
											'4.strColValue:Column Value
											'5.strPopupMenu:PopUp Menu

'Return Value		:			True/False

'Pre-requisite		:			Data Panel Should be Open
'											
'Examples			:		Fn_ReqMgr_AttachmentTableNodeOpeations("VerifyNode","REQ-000048/A;1-Test:CSTMNOTE-000000/A;1-Note","", "","")
										'Fn_ReqMgr_AttachmentTableNodeOpeations("Select","REQ-019554/A;1-test:CSTMNOTE-000004/A;1-test","", "","")
										'Fn_ReqMgr_AttachmentTableNodeOpeations("Expand","REQ-000048/A;1-Test:CSTMNOTE-000000/A;1-Note","", "","")
										'Fn_ReqMgr_AttachmentTableNodeOpeations("NodeDoubleClick","REQ-000048/A;1-Test:Test","", "","")
										'Fn_ReqMgr_AttachmentTableNodeOpeations("GetCellData","000427/A;1-Test:CSTMNOTE-000021/A;1-Test","Relation", "","")->Full Path of note need and Column Name
'History:	
'										Developer Name			Date				Rev. No.			Changes Done			Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'										Sandeep N				20-Jully-2010		1.0										   							 Tushar B	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_AttachmentTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_AttachmentTableNodeOpeations"
   'Declaring Variable
    Dim ObjAttachmentsTable,ObjTable
    Dim IntRows ,StrNodePath, IntCounter,StrIndex,iRowNo
	'Function Return False
	Fn_ReqMgr_AttachmentTableNodeOpeations=False
	'Verify Attachement Table
    If Fn_UI_ObjectExist("Fn_ReqMgr_AttachmentTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("AttachmentsTable"))=False Then
		'Activating Attachement Tab
        Call Fn_ReqMgr_RMTabPanelOperation("Activate","Attachments","")
	End If
	'Creating Object of Attachment Table
	Set ObjAttachmentsTable=Fn_UI_ObjectCreate("Fn_ReqMgr_AttachmentTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("AttachmentsTable"))
	'Get the No. of rows present in the Attachements Table
	IntRows =Fn_UI_Object_GetROProperty("Fn_ReqMgr_AttachmentTableNodeOpeations",ObjAttachmentsTable,"rows")
	'Creating Object of Attachments Table
	Set ObjTable =ObjAttachmentsTable.Object
	If strNodeName<>"" Then
			'Get the Row No. of required Node
			For IntCounter = 0 to IntRows -1
                Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_AttachmentTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"AttachmentsTable",IntCounter)
				StrNodePath = Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_AttachmentTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"AttachmentsTable",IntCounter,0)
				StrNodePath =Mid(StrNodePath,Instr(1,StrNodePath,":")+1,Len(StrNodePath))
											
                If Trim(StrNodePath) = Trim(StrNodeName) Then
                	StrIndex = Cstr(IntCounter)
                    Exit For
				End If
			Next
			If Cint(IntCounter) = Cint(IntRows) Then
        		StrIndex = "FAIL:Node Not Found"
                Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "&  Trim(StrNodeName) &" Invalid Node Name") 
			End If
	End If
   	Select Case StrAction
			Case "Select"		'Fn_ReqMgr_AttachmentTableNodeOpeations("Select","REQ-019554/A;1-test:CSTMNOTE-000004/A;1-test","", "","")
				iRowNo = StrIndex
				If isNumeric(iRowNo) Then
                    Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_AttachmentTableNodeOpeations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"AttachmentsTable",iRowNo)
					Fn_ReqMgr_AttachmentTableNodeOpeations=True
				End if
		   Case "VerifyNode"		'Fn_ReqMgr_AttachmentTableNodeOpeations("VerifyNode","REQ-000048/A;1-Test:CSTMNOTE-000000/A;1-Note","", "","")
            		iRowNo = StrIndex
					If isNumeric(iRowNo) then
						Fn_ReqMgr_AttachmentTableNodeOpeations=True
					End if
			Case "Expand"	'Fn_ReqMgr_AttachmentTableNodeOpeations("Expand","REQ-000048/A;1-Test:CSTMNOTE-000000/A;1-Note","", "","")
				iRowNo =StrIndex
				If isNumeric(iRowNo) then
                    Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_AttachmentTableNodeOpeations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"AttachmentsTable",iRowNo)
                    Call Fn_MenuOperation("Select", "View:Expand")
					Fn_ReqMgr_AttachmentTableNodeOpeations=True
				End if
			Case "NodeDoubleClick"	'Fn_ReqMgr_AttachmentTableNodeOpeations("NodeDoubleClick","REQ-000048/A;1-Test:Test","", "","")
				iRowNo = StrIndex
				If isNumeric(iRowNo) Then
                    Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_AttachmentTableNodeOpeations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"AttachmentsTable",iRowNo)
					Call Fn_UI_JavaTable_DoubleClickCell("Fn_ReqMgr_AttachmentTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"AttachmentsTable",iRowNo,0,"","")
					Fn_ReqMgr_AttachmentTableNodeOpeations=True
				End if
			Case "GetCellData"		'Fn_ReqMgr_AttachmentTableNodeOpeations("GetCellData","000427/A;1-Test:CSTMNOTE-000021/A;1-Test","Relation", "","")
					iRowNo = StrIndex
					If isNumeric(iRowNo) Then
						Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_AttachmentTableNodeOpeations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"), "AttachmentsTable",iRowNo)
						wait(5)
						Fn_ReqMgr_AttachmentTableNodeOpeations=JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("AttachmentsTable").GetCellData(iRowNo,strColName)
					End If
	End Select
	'Release the Table object
	 Set ObjTable = Nothing
	 Set ObjAttachmentsTable=Nothing
End Function

'------------------------------------------------------------------------------Function to CreateBasic  Custome Note------------------------------------------------------------------------------------------------------
'Function Name		:			Fn_ReqMgr_CustomeNoteBasicCreate

'Description			 :		 	Function Creates Custome Note with Basic Information

'Parameters			   :	 		1.sNoteType: Type of the item.(Custome Note)
'												 2.sConfItem: "True" or "False"
'												 3.sNoteID: ID of the Custome Note it should be unique.
'												4.sNoteRevID:Revision ID of the Custoem Note.
'												5.sNoteName:Name of Custome Note.
'												6.sNoteDesc: Description of the Custome Note.
'												7:sNoteUOM: Unit of measure of Custome Note. ( not handling this part)

'Return Value		   : 				Note Id  - Revision Id

'Pre-requisite			:		 		should be logged in & present on Requirement Manager perspective Object should be selected on which have to create Custome Note

'Examples				:				 Fn_ReqMgr_CustomeNoteBasicCreate("Custom Requirement","","","","Test Note","","")
													'Fn_ReqMgr_CustomeNoteBasicCreate("Custom Requirement","OFF","","","Test Note","Test Description","")
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   20/07/2010			              1.0										Created							Tushar B
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_CustomeNoteBasicCreate(sNoteType,sConfItem,sNoteID,sNoteRevID,sNoteName,sNoteDesc,sNoteUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_CustomeNoteBasicCreate"
	'Variable Declaration
	Dim strNoteID, sRevId,bFlag
	Dim ObjCustomeNote,objSelectType
		Fn_ReqMgr_CustomeNoteBasicCreate=False
		'Verifying Existance of "NewCustomNote" window
		If Fn_UI_ObjectExist("Fn_ReqMgr_CustomeNoteBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewCustomNote"))=False Then
			'Select menu [File -> New -> Custome Note]
			 Call Fn_MenuOperation("Select","File:New:Custom Note")
		End If
		'Check the existence of "NewCustomNote" window
		Set ObjCustomeNote=Fn_UI_ObjectCreate("Fn_ReqMgr_CustomeNoteBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewCustomNote"))
		'Checking Existance of Note Type
		bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote,"EnterAdditionalNote",sNoteType)
		If bFlag=True Then
			'Select Note Type
            Call Fn_List_Select("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote,"EnterAdditionalNote",sNoteType)
		Else
	        Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& sNoteType &" Is Invalid note type")
			Set ObjCustomeNote=Nothing
			Exit Function
		End If
		'checked Configuration Requirement or not
		If sConfItem <> "" Then
            Call Fn_CheckBox_Set("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote,"ConfigurationItem",sConfItem)
		End If
		'Click on "Next" button
		Call Fn_Button_Click("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote,"Next")
		
		If sNoteID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote,"NoteID", sNoteID)
		End If
	
		If sNoteRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_ReqMgr_CustomeNoteBasicCreate",ObjCustomeNote,"Revision", sNoteRevID)
		End If
		
		If  sNoteID = "" or sNoteRevID = "" Then
			'click on assign button
			  Call Fn_Button_Click("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote, "Assign")
		End If
		
		'Extract Creation data
		strNoteID =Fn_Edit_Box_GetValue("Fn_ReqMgr_CustomeNoteBasicCreate",ObjCustomeNote,"NoteID")
		sRevId = Fn_Edit_Box_GetValue("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote,"Revision")		
		'Set Requirement name
		 Call Fn_Edit_Box("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote,"Name",sNoteName)
		'Set description
		Call Fn_Edit_Box("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote,"Description",sNoteDesc)
		'Click on "Finish" button
	
		Call Fn_Button_Click("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote, "Finish") 
		Fn_ReqMgr_CustomeNoteBasicCreate = strNoteID & "-" & sRevId
		wait(10)
		ObjCustomeNote.JavaButton("Close").WaitProperty "enabled", 1, 20000
        'Click on Close button
		Call Fn_Button_Click("Fn_ReqMgr_CustomeNoteBasicCreate", ObjCustomeNote, "Close")
						
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Custome Note of ID [" + CStr(sNoteID) + "]")
		'Release object of Custome Note window
		Set ObjCustomeNote=Nothing
End Function
'-------------------------------------------------------------------------Function To Perform Operations on Attachments Table------------------------------------------------------------------
'Function Name		:	Fn_ReqMgr_AttachmentTableNodeOpeations

'Description			:	This Function is used to Perform operations on Attachments Table

'Parameters			:			1.	strAction:Action Name 			Eg -> "Select"
'											 2.	 strNodeName:Node name on which have to perform operation
											'3.	strColName:Column Name
											'4.strColValue:Column Value
											'5.strPopupMenu:PopUp Menu

'Return Value		:			True/False

'Pre-requisite		:		Allocations Table Should be displayed  -> Call Fn_MenuOperation("Select","Edit:Toggle In Allocation Context Mode")
'											
'Examples			:		'Fn_ReqMgr_AllocationsTableNodeOpeations("Select","REQ-019554/A;1-test:Allocation1","", "","")
										
'History:
'										Developer Name					Date						Rev. No.			Changes Done			Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane			21/07/2010			            1.0						Created							Tushar B
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_AllocationsTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_AllocationsTableNodeOpeations"
   'Declaring Variable
    Dim ObjAllocationsTable,ObjTable
    Dim IntRows ,StrNodePath, IntCounter,StrIndex,iRowNo
	'Function Return False
	Fn_ReqMgr_AllocationsTableNodeOpeations=False
	'Verify Allocation Table
    If Fn_UI_ObjectExist("Fn_ReqMgr_AllocationsTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("Allocations"))=False Then
		Exit Function
		'Activating Allocation Tab
        'Call Fn_MenuOperation("Select","Edit:Toggle In Allocation Context Mode")
	End If
	'Creating Object of Allocation Table
	Set ObjAllocationsTable=Fn_UI_ObjectCreate("Fn_ReqMgr_AllocationsTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("Allocations"))
	'Get the No. of rows present in the Allocation Table
	IntRows =Fn_UI_Object_GetROProperty("Fn_ReqMgr_AllocationsTableNodeOpeations",ObjAllocationsTable,"rows")
	'Creating Object of Attachments Table
	Set ObjTable =ObjAllocationsTable.Object
	If strNodeName<>"" Then
			'Get the Row No. of required Node
			For IntCounter = 0 to IntRows -1
                Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_AllocationsTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"Allocations",IntCounter)
				StrNodePath = Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_AllocationsTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"Allocations",IntCounter,0)
				StrNodePath =Mid(StrNodePath,Instr(1,StrNodePath,":")+1,Len(StrNodePath))
											
                If Trim(StrNodePath) = Trim(StrNodeName) Then
                	StrIndex = Cstr(IntCounter)
                    Exit For
				End If
			Next
			If Cint(IntCounter) = Cint(IntRows) Then
        		StrIndex = "FAIL:Node Not Found"
                Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "&  Trim(StrNodeName) &" Invalid Node Name") 
			End If
	End If
   	Select Case StrAction
			Case "Select"		'Fn_ReqMgr_AllocationsTableNodeOpeations("Select","REQ-019554/A;1-test:Allocation1","", "","")
				iRowNo = StrIndex
				If isNumeric(iRowNo) Then
                    Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_AllocationsTableNodeOpeations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"Allocations",iRowNo)
					Fn_ReqMgr_AllocationsTableNodeOpeations=True
				End if
	End Select
	'Release the Table object
	 Set ObjTable = Nothing
	 Set ObjAllocationsTable=Nothing
End Function

'------------------------------------------------------------------------------Function to Create Basic  Allocation Map------------------------------------------------------------------------------------------------------
'Function Name		:			Fn_ReqMgr_AllocationMapBasicCreate

'Description			 :		 	Function Creates Allocation Map with Basic Information

'Parameters			   :	 		1.sMapType: Type of the item.(AllocationMap)
'												 2.sConfItem: "True" or "False"
'												 3.sMapID: ID of the Allocation Map it should be unique.
'												4.sMapRevID:Revision ID of the Allocation Map.
'												5.sMapName:Name of Allocation Map.
'												6.sMapDesc: Description of the Allocation Map..
'												7:sMapUOM: Unit of measure of Allocation Map.. ( not handling this part)

'Return Value		   : 			Allocation Map  Id  - Revision Id

'Pre-requisite			:		 	Should be logged in & present on Requirement Manager perspective & 2 BOM structure should be present

'Examples				:			Fn_ReqMgr_AllocationMapBasicCreate("AllocationMap","OFF","","","Test Map","Test Description","")
													
'History					 :		
'													Developer Name												Date						Rev. No.						Changes Done						Reviewer
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'												Sandeep Navghane									   21/07/2010			              1.0										Created							Tushar B
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_AllocationMapBasicCreate(sMapType,sConfItem,sMapID,sMapRevID,sMapName,sMapDesc,sMapUOM)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_AllocationMapBasicCreate"
	'Variable Declaration
	Dim strAlocMapID, sRevId,bFlag
	Dim ObjAllocationMap,objSelectType
		Fn_ReqMgr_AllocationMapBasicCreate=False
		'Verifying Existance of "NewAllocationMap" window
		If Fn_UI_ObjectExist("Fn_ReqMgr_AllocationMapBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewAllocationMap"))=False Then
			'Opening NewAllocationMap Dialog
			Call Fn_Button_Click("Fn_ReqMgr_AllocationMapBasicCreate",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"createAllocContext")
		End If
		'Check the existence of "NewCustomMap" window
		Set ObjAllocationMap=Fn_UI_ObjectCreate("Fn_ReqMgr_AllocationMapBasicCreate",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewAllocationMap"))
		'Checking Existance of Map Type
		bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap,"AllocationMapList",sMapType)
		If bFlag=True Then	
			'Select Map Type
            Call Fn_List_Select("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap,"AllocationMapList",sMapType)
		Else
	        Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "& sMapType &" Is Invalid Map type")
			Set ObjAllocationMap=Nothing
			Exit Function
		End If
		'checked Configuration Requirement or not
		If sConfItem <> "" Then	
            Call Fn_CheckBox_Set("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap,"ConfigurationItem",sConfItem)
		End If
		'Click on "Next" button
		Call Fn_Button_Click("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap,"Next")
		
		If sMapID <> "" Then
			'Set  Item Id
			 Call Fn_Edit_Box("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap,"AllocationMapId", sMapID)
		End If
		
		If sMapRevID <> "" Then
			'Set Revision ID
			Call Fn_Edit_Box("Fn_ReqMgr_AllocationMapBasicCreate",ObjAllocationMap,"AllocRevision", sMapRevID)
		End If
		
		If  sMapID = "" or sMapRevID = "" Then
			'click on assign button
			  Call Fn_Button_Click("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap, "Assign")
		End If
		
		'Extract Creation data
		strAlocMapID =Fn_Edit_Box_GetValue("Fn_ReqMgr_AllocationMapBasicCreate",ObjAllocationMap,"AllocationMapId")
		sRevId = Fn_Edit_Box_GetValue("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap,"AllocRevision")		
		'Set Requirement name
		 Call Fn_Edit_Box("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap,"AllocName",sMapName)
		'Set description
		Call Fn_Edit_Box("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap,"Description",sMapDesc)
		'Click on "Finish" button
		Call Fn_Button_Click("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap, "Finish") 
		Fn_ReqMgr_AllocationMapBasicCreate = strAlocMapID & "-" & sRevId
        'Click on Close button
		Call Fn_Button_Click("Fn_ReqMgr_AllocationMapBasicCreate", ObjAllocationMap, "Close")
						
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Created an Allocation Map of ID [" + CStr(sMapID) + "]")
		'Release object of Custome Map window
		Set ObjAllocationMap=Nothing
End Function

'-------------------------------------------------------------------------This Function is used to Create New Allocation-----------------------------------------------------------------------------
'Function Name		:	Fn_ReqMgr_CreateAllocation

'Description			:	This Function is used to Create New Allocation

'Parameters			:			1.	strName:Allocation Name
'											 2.	 strReason:Allocation Reason
											'3.	strType:

'Return Value		:	True/False

'Pre-requisite		:	New Allocations Dialog should be displayed     -> Fn_MenuOperation("Select","Tools:Allocate To...")
'											
'Examples			:	Fn_ReqMgr_CreateAllocation("Allocation1","Test Reason","")
'									Fn_ReqMgr_CreateAllocation("Allocation2","","")
'									Fn_ReqMgr_CreateAllocation("","Test Reason","")
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   21/07/2010			              1.0										Created							Tushar B
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_CreateAllocation(strName,strReason,strType)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_CreateAllocation"
   'Function Returns False
	Fn_ReqMgr_CreateAllocation=False
	'Verifying "NewAllocation" dialogs Existance
	If Fn_UI_ObjectExist("Fn_ReqMgr_CreateAllocation",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewAllocation"))=True Then
		If strName<>"" Then
			'Setting Name to Allocation
			Call Fn_Edit_Box("Fn_ReqMgr_CreateAllocation",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewAllocation"),"Name",strName)
		End If
		If strReason<>"" Then
			'Setting Reason To Allocation
			Call Fn_Edit_Box("Fn_ReqMgr_CreateAllocation",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewAllocation"),"Reason",strReason)
		End If
		'Click On OK button
		Call Fn_Button_Click("Fn_ReqMgr_CreateAllocation",JavaWindow("RequirementsManager").JavaWindow("RMWindow").JavaDialog("NewAllocation"), "OK")
		'Function Returns True
		Fn_ReqMgr_CreateAllocation=True
	End If
End Function

'-------------------------------------------------------------------------This Function is used to to perform operation on MS Word tab---------------------------------------------------------------
'Function Name		:	Fn_ReqMgr_MSWordTabOperations

'Description			:	This Function is used to to perform operation on MS Word tab

'Parameters			:			1.	strAction:Action Name
'											 2.	 strValue:Value to set in text Box Or to verify the value
											'3.	strParameterName:Parameter Name

'Return Value		:	True/False

'Pre-requisite		:	Object should be selected
'											
'Examples			:	Fn_ReqMgr_MSWordTabOperations("SetValue","[parameter1:value1,value2,value3..]"+VbCrlf+"[parameter2: value1,value2,value3..]","")
'									Fn_ReqMgr_MSWordTabOperations("VerifyValue","value1","parameter1")
'							Case "VerifyInStr"	 : Fn_ReqMgr_MSWordTabOperations("VerifyInStr","[humidity:9,8,7]","")	'Case Added By Ketan Raje on 15-Nov-2010
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   28/07/2010			              1.0										Created							Tushar B
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_MSWordTabOperations(strAction,strValue,strParameterName)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_MSWordTabOperations"
	'Variable declaration
   Dim bReturn,bFlag,iRowCnt,sCellData,sValueName,iCount
   'Function Return False
   Fn_ReqMgr_MSWordTabOperations=False
   bReturn=False
   bFlag=False
   'Verifying MS Word tab is activated or not
   bFlag=Fn_MyTc_TabOperation("VerifyActivate", "MS Word")
   If bFlag=False Then
	   'Activating MS Word tab
	   Call Fn_SetView("Teamcenter:MS Word")
   End If
	Select Case strAction
		'"SetValue" this action set the value 
		Case "SetValue"		'Fn_ReqMgr_MSWordTabOperations("SetValue","[parameter1:value1,value2,value3..]"+VbCrlf+"[parameter2: value1,value2,value3..]","")
				'Setting Value in text box
                Call Fn_Edit_Box("Fn_ReqMgr_MSWordTabOperations",JavaWindow("MyTeamcenter"),"MSWordTabText",strValue)
				'Saving the changes
                Call Fn_ToolbatButtonClick("Save")
				Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully Set  value"& strValue)   
				'Function Return True
				Fn_ReqMgr_MSWordTabOperations=True
		 'VerifyValue Case to verify paameter values
		Case "VerifyValue"	'Fn_ReqMgr_MSWordTabOperations("VerifyValue","value1","parameter1")
				'Taking no of rows from ParametricValues table
				iRowCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_MSWordTabOperations",JavaWindow("MyTeamcenter").JavaTable("ParametricValues"),"rows")
				For iCount=0 To iRowCnt-1
						'Taking data from ParametricValues table
						sCellData=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_MSWordTabOperations",JavaWindow("MyTeamcenter"), "ParametricValues",iCount,0)
						If strParameterName=sCellData Then
								sValueName=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_MSWordTabOperations",JavaWindow("MyTeamcenter"), "ParametricValues",iCount,1)
								If strValue=sValueName Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully verify  value"& strValue )   
									bReturn=True
									Exit For
								End If
						End If
				Next
				If bReturn=False Then
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:"& strParameterName & "Parameter is not present")   
					Exit Function
				Else
					  Fn_ReqMgr_MSWordTabOperations=True
				End If
		Case "VerifyInStr"		
				'Getting Value in MSWord text box
				sValueName = Fn_Edit_Box_GetValue("Fn_ReqMgr_MSWordTabOperations",JavaWindow("MyTeamcenter"),"MSWordTabText")
				If Instr(1, sValueName, strValue)<>0 Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:"& strValue &"Successfully found")
					Fn_ReqMgr_MSWordTabOperations=True
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:"& strValue &"Not Found.")
					Fn_ReqMgr_MSWordTabOperations=False
				End If		
		Case "SetValueWithoutSave"		'Fn_ReqMgr_MSWordTabOperations("SetValue","[parameter1:value1,value2,value3..]"+VbCrlf+"[parameter2: value1,value2,value3..]","")
				'Setting Value in text box
                Call Fn_Edit_Box("Fn_ReqMgr_MSWordTabOperations",JavaWindow("MyTeamcenter"),"MSWordTabText",strValue)
				'Saving the changes
     			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully Set  value"& strValue)   
				'Function Return True
				Fn_ReqMgr_MSWordTabOperations=True						
	End Select
	'Activating summary tab
	'Call Fn_MyTc_TabOperation("Activate", "Summary")
End Function
'-------------------------------------------------------------------------This Function is used to to perform operation on Input Parametric Values Window-----------------------------------------
'Function Name		:	Fn_ReqMgr_ParamatricValueOperation

'Description			:	This Function is used to to perform operation on Input Parametric Values Window

'Parameters			:			1.	strAction:Action Name
'											 2.	 strValue:Value to select from Table
											'3.	 strNoteText:

'Return Value		:	True/False

'Pre-requisite		:	 Object should be selected
'											
'Examples			:	Fn_ReqMgr_ParamatricValueOperation("SetParametricValue","value1:value2","")
'									
'									
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   28/07/2010			              1.0										Created							Tushar B
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_ParamatricValueOperation(strAction,strValue,strNoteText)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_ParamatricValueOperation"
   'Declaring variables
   Dim iCounter,strVal
   Dim objSelectType,objDialog
	'Function Return False
    Fn_ReqMgr_ParamatricValueOperation=False
	'Verifying existance of Input Parametric Values window
	If Fn_UI_ObjectExist("Fn_ReqMgr_ParamatricValueOperation",JavaWindow("MyTeamcenter").JavaWindow("Input Parametric Values"))=False Then
		'Invoking Input Parametric Values window
		Call Fn_MenuOperation("Select","Edit:Attach Requirements/Notes:Parametric Requirement")
	End If
	
	Select Case strAction
			'SetParametricValue Case set values
			Case "SetParametricValue"	'Fn_ReqMgr_ParamatricValueOperation("SetParametricValue","value1:value2","")
                    strVal=Split(strValue,":")
                    For iCounter=0 To Ubound(strVal)
						If strVal(iCounter)<>"" Then
								Set objSelectType=description.Create()
								objSelectType("Class Name").value = "JavaList"
								Set objDialog =JavaWindow("MyTeamcenter").JavaWindow("Input Parametric Values").JavaTable("Table").ChildObjects(objSelectType)
                                objDialog(iCounter).Select strVal(iCounter)
								Fn_ReqMgr_ParamatricValueOperation=True
								Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully Set  value"& strVal(iCounter))
						End If
					Next				
		End Select
	 'Click on OK button
	Call Fn_Button_Click("Fn_ReqMgr_ParamatricValueOperation", JavaWindow("MyTeamcenter").JavaWindow("Input Parametric Values"), "OK")
	'Releasing objects
	Set objSelectType=Nothing
	Set objDialog =Nothing
End Function
'-------------------------------------------------------------------------This Function is used to to perform operation on IDetails Table In RM----------------------------------------------------
'Function Name		:	Fn_ReqMgr_DetailTableOperation

'Description			:	This Function is used to to perform operation on IDetails Table In RM

'Parameters			:	   1.) sAction: Action string to navigate to appropriate case
'									    2.) sObjectName: Name of the object under Details Table
'  										3.) sColumnName: Name of the column under Details Table
'										4.) sExpectedValue: Expected value of the object property under Details Table

'Return Value		:	True/False/ColumnCount

'Pre-requisite		:	 Should be present on RM perspective And If details table is open then Good
'											
'Examples			:	Fn_ReqMgr_DetailTableOperation("ColumnCount", "", "", "","")
'									Fn_ReqMgr_DetailTableOperation("Rowmultiselect","REQ-000301/A;1-Req1:REQ-000302/A;1-Req2", "", "","")
'									Fn_ReqMgr_DetailTableOperation("PopUpMenuSelect","", "", "","Apply Column Configuration...")
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   28/07/2010			              1.0										Created							Tushar B
'Added GetcellData Case  Sukhada                                                               02/08/10                                                                      Modified                           Tushar B
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_DetailTableOperation(sAction, sObjectName, sColumnName, sExpectedValue,sPopUpMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_DetailTableOperation"
	Dim bReturn,oCounter,aObjList,intItemCount,sText,iCounter,iRowCnt,sObjName,sCellValue
	Dim ObjDetailsTable, strName
	Fn_ReqMgr_DetailTableOperation=False
	'Verifying Existance of Details Table
	If Fn_UI_ObjectExist("Fn_ReqMgr_DetailTableOperation",JavaWindow("RequirementsManager").JavaTable("RMDetailTable"))=False Then
		'Invoking Details Table
		Call Fn_SetView("Teamcenter:Details")
	End If
	Set ObjDetailsTable=Fn_UI_ObjectCreate("Fn_ReqMgr_DetailTableOperation",JavaWindow("RequirementsManager").JavaTable("RMDetailTable"))
	Select Case sAction
		 Case "ColumnCount"		'Return Number of columns currently present in Details Table
   				'Returning Number of columns present in Details Table
				Fn_ReqMgr_DetailTableOperation =Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailTableOperation",ObjDetailsTable,"cols")			
		 Case "Rowmultiselect"
				'Split the string where " : " exist
				aObjList = Split(sObjectName,":")
				intItemCount =ubound(aObjList)
				'Count number of rows of Table
				bReturn=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailTableOperation",ObjDetailsTable,"rows")
				'Extract the index of row at which the object exist.
				For oCounter=0 to intItemCount
						For iCounter=0 to bReturn-1
						sText = ObjDetailsTable.GetCellData(iCounter,"Object")						
						If IsNumeric(aObjList(oCounter)) Then
							 If cstr(sText) = cstr(cint(aObjList(oCounter)))  Then
								 ObjDetailsTable.ClickCell iCounter, "Object","LEFT","CONTROL"
								 Exit for
							End If
						ElseIf cstr(sText) = cstr(aObjList(oCounter))  Then
								 ObjDetailsTable.ClickCell iCounter, "Object","LEFT","CONTROL"
								 Fn_ReqMgr_DetailTableOperation=True
								 Exit for
						End If									
						Next
				Next
		   Case "PopUpMenuSelect"
					Call Fn_ToolbatButtonClick("View Menu")
					JavaWindow("RequirementsManager").WinMenu("ContextMenu").Select sPopUpMenu
					 Fn_ReqMgr_DetailTableOperation=True

			Case "GetCellData" '("GetCellData",1,0,"",")
					
						JavaWindow("RequirementsManager").JavaTable("RMDetailTable").SelectRow sObjectName
						strName=JavaWindow("RequirementsManager").JavaTable("RMDetailTable").GetCellData(sObjectName,sColumnName)
						
					If Err.number < 0 Then
                		Fn_ReqMgr_DetailTableOperation=False
					Else
						Fn_ReqMgr_DetailTableOperation =MId(strName,instr(1,strName,":")+1 , Len(strName))
					End If

			Case "GetIndex" '("GetCellData",1,0,"","")
					iRowCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailTableOperation",ObjDetailsTable,"rows")
					For iCounter=0 To iRowCnt-1
						sObjName=ObjDetailsTable.GetCellData(iCounter,"Object")
						If sObjName=sObjectName Then
							
                            	Fn_ReqMgr_DetailTableOperation=iCounter
								Exit For
							
						End If
					Next

		 Case "VerifyCell"
					iRowCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailTableOperation",ObjDetailsTable,"rows")
					For iCounter=0 To iRowCnt-1
						sObjName=ObjDetailsTable.GetCellData(iCounter,"Object")
						If sObjName=sObjectName Then
							sCellValue=ObjDetailsTable.GetCellData(iCounter,sColumnName)
							If sCellValue=sExpectedValue Then
								Fn_ReqMgr_DetailTableOperation=True
								Exit For
							End If
						End If
					Next
		End Select
        		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "PASS: Fn_MyTc_DetailTableContentOperation passed with case "&sAction&" on Object "&sObjectName)	
				Set ObjDetailsTable=Nothing 
End Function
'-------------------------------------------------------------------------This Function is used to to perform operation on Apply Column Application menu--------------------------------------------------------
'Function Name		:	Fn_ReqMgr_ApplyColumnConfiguration

'Description			:	This Function is used to to perform operation on Apply Column Application menu

'Parameters			:	   1.) sAction: Action string to navigate to appropriate case
'									    2.) strConfigName: Name of Configuration (It should be unique in ConfigurationSaveAs and Add case )
'  										3.) arrAvailableProp: Avaiable Propeties array
'										4.) bShowIntPropName: Show Internal names of Properties option
'										5.)strConfigDesc:Description of Configuration

'Return Value		:	True/False

'Pre-requisite		:	Should be details table present
'											
'Examples			:	Fn_ReqMgr_ApplyColumnConfiguration("ConfigurationSaveAs","Demo4","","","")
'									Fn_ReqMgr_ApplyColumnConfiguration("ColumnAdd","Demo5",columnName,"","")
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   28/07/2010			              1.0										Created							Tushar B
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_ApplyColumnConfiguration(strAction,strConfigName,arrAvailableProp,bShowIntPropName,strConfigDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_ApplyColumnConfiguration"
   'Declaring variables
   Dim bFlag,iCounter,iRowCount,iCnt,sDsplColName,iAvlRowCount,intCount,avlColName
   'Declaring Object
   bFlag=False
   Dim ObjColumnWnd,ObjColumnMngmntWnd,ObjColumnConfigWnd
   Fn_ReqMgr_ApplyColumnConfiguration=False
   'verifying existance of ApplyColumnConfiguration window
	If Fn_UI_ObjectExist("Fn_ReqMgr_ApplyColumnConfiguration",JavaWindow("RequirementsManager").JavaWindow("ApplyColumnConfiguration"))=False Then
		'Invoking ApplyColumnConfiguration window
		Call Fn_ReqMgr_DetailTableOperation("PopUpMenuSelect","", "", "","Apply Column Configuration...")
	End If
	'Creating objects 
	Set ObjColumnWnd=Fn_UI_ObjectCreate("Fn_ReqMgr_ApplyColumnConfiguration",JavaWindow("RequirementsManager").JavaWindow("ApplyColumnConfiguration"))
	Set ObjColumnMngmntWnd=JavaWindow("RequirementsManager").JavaWindow("ApplyColumnConfiguration").JavaWindow("Column Management")
	Set ObjColumnConfigWnd=JavaWindow("RequirementsManager").JavaWindow("ApplyColumnConfiguration").JavaWindow("Column Management").JavaWindow("SaveColumnConfiguration")
	Select Case strAction
		Case "ConfigurationSaveAs" 'This Case use to create same as default configuration
			Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd,"PlusButton")
			'Verifying existance Column Management window
			If Fn_UI_ObjectExist("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnMngmntWnd)=False Then
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to Invoke Column Management Window")	
				Set ObjColumnWnd=Nothing
				Set ObjColumnMngmntWnd=Nothing
				Set ObjColumnConfigWnd=Nothing
				Exit Function
			End If
			'Clicking on save button to open SaveColumnConfiguration
			Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Save")
			'Setting name of configuration 
            Call Fn_UI_EditBox_Type("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnConfigWnd,"Name",strConfigName)
			If strConfigDesc<>"" Then
				'Setting Description of configuration 
				Call Fn_Edit_Box("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnConfigWnd,"Description",strConfigDesc)
			End If
			'Clicking on save button to create configuration
			Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnConfigWnd,"Save")
			'Closing the window
			Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Close")
			'Checking existance of Configuration in ColumnConfigurations list
            bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd, "ColumnConfigurations",strConfigName)
			If bFlag=True Then
				'Selecting Confugarion in ColumnConfigurations list
				Call Fn_List_Select("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd, "ColumnConfigurations",strConfigName)
				'Applying the changes
				Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd,"Apply")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to create confugaration of name" & strConfigName)	
				Set ObjColumnWnd=Nothing
				Set ObjColumnMngmntWnd=Nothing
				Set ObjColumnConfigWnd=Nothing
				Exit Function
			End If
	Case "ColumnAdd"
			Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd,"PlusButton")
			'Verifying existance Column Management window
			arrAvailableProp=split(arrAvailableProp,":")
			'If  IsArray(arrAvailableProp) Then
				For iCounter=0 To Ubound(arrAvailableProp)
					 If arrAvailableProp(iCounter)<>"" Then
						 iRowCount=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailTableOperation",ObjColumnMngmntWnd.JavaTable("DisplayedColumns"),"rows")
						 For iCnt=0 To iRowCount-1
							 bFlag=False
							sDsplColName=ObjColumnMngmntWnd.JavaTable("DisplayedColumns").GetCellData(iCnt,"")
								If sDsplColName=arrAvailableProp(iCounter) Then
									bFlag=True
									Exit For
								End If
						Next
						If bFlag=False Then
							iAvlRowCount=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailTableOperation",ObjColumnMngmntWnd.JavaTable("AvailableProperties"),"rows")
							'Selecting Confugarion in ColumnConfigurations list
							For intCount=0 To iAvlRowCount-1
									avlColName=ObjColumnMngmntWnd.JavaTable("AvailableProperties").GetCellData(intCount,"Property")
									If avlColName=arrAvailableProp(iCounter) Then
										ObjColumnMngmntWnd.JavaTable("AvailableProperties").SelectCell intCount,"Property"
										Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnMngmntWnd,"AddItem")
										Exit For
									Else
'										Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to create confugaration of name" & strConfigName)	
'										Set ObjColumnWnd=Nothing
'										Set ObjColumnMngmntWnd=Nothing
'										Set ObjColumnConfigWnd=Nothing
'										Exit Function
									End If
							Next                     					
						End If
					 End If
				Next
			'End If
			'Clicking on save button to open SaveColumnConfiguration
			Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Save")
			'Setting name of configuration 
            Call Fn_UI_EditBox_Type("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnConfigWnd,"Name",strConfigName)
			If strConfigDesc<>"" Then
				'Setting Description of configuration 
				Call Fn_Edit_Box("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnConfigWnd,"Description",strConfigDesc)
			End If
			'Clicking on save button to create configuration
			Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnConfigWnd,"Save")
			'Closing the window
			Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnMngmntWnd,"Close")
			'Checking existance of Configuration in ColumnConfigurations list
            bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd, "ColumnConfigurations",strConfigName)
			If bFlag=True Then
				'Selecting Confugarion in ColumnConfigurations list
				Call Fn_List_Select("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd, "ColumnConfigurations",strConfigName)
				'Applying the changes
				Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd,"Apply")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to create confugaration of name" & strConfigName)	
				Set ObjColumnWnd=Nothing
				Set ObjColumnMngmntWnd=Nothing
				Set ObjColumnConfigWnd=Nothing
				Exit Function
			End If
	 Case "Apply"
			bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd, "ColumnConfigurations",strConfigName)
			If bFlag=True Then
				'Selecting Confugarion in ColumnConfigurations list
				Call Fn_List_Select("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd, "ColumnConfigurations",strConfigName)
				'Applying the changes
				Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd,"Apply")
			Else
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail:Failed to create confugaration of name" & strConfigName)	
				Set ObjColumnWnd=Nothing
				Set ObjColumnMngmntWnd=Nothing
				Set ObjColumnConfigWnd=Nothing
				Exit Function
			End If
	End Select
	Fn_ReqMgr_ApplyColumnConfiguration=True
	Call Fn_Button_Click("Fn_ReqMgr_ApplyColumnConfiguration",ObjColumnWnd,"Close")
	'Releasing all objects
	Set ObjColumnWnd=Nothing
	Set ObjColumnMngmntWnd=Nothing
	Set ObjColumnConfigWnd=Nothing
End Function
'-------------------------------------------------------------------------This Function is used to Sort Details Table Containt-----------------------------------------------------------------------------------------------
'Function Name		:	Fn_ReqMgr_DetailTableSort

'Description			:	This Function is used to Sort Details Table Containt

'Parameters			:	   1.) strSortBy: First Sorting criteria (Name Of Column And Criteria )
'																  Eg."Type:Asc"
'									    2.) strThenBy1: Second Sorting criteria (Name Of Column And Criteria )
'																  Eg."Object:Desc"
'  										3.) strThenBy2: Third Sorting criteria (Name Of Column And Criteria )
'																  Eg."Description"

'Return Value		:	True/False

'Pre-requisite		:	Details Table Should be Present
'											
'Examples			:	Fn_ReqMgr_DetailTableSort("Object:Asc","Type:Desc","Group ID:")
'									
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   29/07/2010			              1.0										Created							Tushar B
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_DetailTableSort(strSortBy,strThenBy1,strThenBy2)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_DetailTableSort"
   'Decalring Variables
   Dim ObjSortWnd,bFlag,sSortCriteria,sSortCriteria1,sSortCriteria2
   'Setting False to bFlag And Function
   bFlag=False
   Fn_ReqMgr_DetailTableSort=False
   'Verifying Existance of Sort Window
   If Fn_UI_ObjectExist("Fn_ReqMgr_DetailTableSort",JavaWindow("RequirementsManager").JavaWindow("Sort"))=False Then
	   'Invoking Sort Window
		Call Fn_ReqMgr_DetailTableOperation("PopUpMenuSelect","", "", "","Sort...")
   End If
	'Ceating object of Sort window
	Set ObjSortWnd=Fn_UI_ObjectCreate("Fn_ReqMgr_DetailTableSort",JavaWindow("RequirementsManager").JavaWindow("Sort"))
	 'Clearing Previous sort criteria
	Call Fn_Button_Click("Fn_ReqMgr_DetailTableSort", ObjSortWnd,"Clear")
	'Spliting first sort parameter
	sSortCriteria=Split(strSortBy,":")
	'Apllying first sort criteria
	If sSortCriteria(0)<>"" Then
		'Checking existance of Item in Sort item list
		bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_DetailTableSort", ObjSortWnd, "SortBy",sSortCriteria(0))
		If bFlag=False Then
            Call Fn_Button_Click("Fn_ReqMgr_DetailTableSort", ObjSortWnd,"Cancel")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Wrong Search Criteria"& sSortCriteria(0) &"Pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
		'Selecting item
		Call Fn_List_Select("Fn_ReqMgr_DetailTableSort", ObjSortWnd, "SortBy",sSortCriteria(0))
	End If
	'Selecting Sort type Ascending Or Descending
	If sSortCriteria(1)<>"" Then
		If sSortCriteria(1)="Asc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_ReqMgr_DetailTableSort",ObjSortWnd,"SortByAsc")
		ElseIf	sSortCriteria(1)="Desc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_ReqMgr_DetailTableSort",ObjSortWnd,"SortByDesc")
		Else
			Call Fn_Button_Click("Fn_ReqMgr_DetailTableSort", ObjSortWnd,"Cancel")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Invalid criteria pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
	End If
	'Spliting Second sort parameter
    sSortCriteria1=Split(strThenBy1,":")
	'Apllying Second sort criteria
	If sSortCriteria1(0)<>"" Then
		'Checking existance of Item in Sort item list
		bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_DetailTableSort", ObjSortWnd, "ThenBy1",sSortCriteria1(0))
		If bFlag=False Then
            Call Fn_Button_Click("Fn_ReqMgr_DetailTableSort", ObjSortWnd,"Cancel")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Wrong Search Criteria"& sSortCriteria1(0) &"Pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
		'Selecting item
		Call Fn_List_Select("Fn_ReqMgr_DetailTableSort", ObjSortWnd, "ThenBy1",sSortCriteria1(0))
	End If
	'Selecting Sort type Ascending Or Descending
	If sSortCriteria1(1)<>"" Then
		If sSortCriteria1(1)="Asc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_ReqMgr_DetailTableSort",ObjSortWnd,"ThenBy1Asc")
		ElseIf	sSortCriteria1(1)="Desc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_ReqMgr_DetailTableSort",ObjSortWnd,"ThenBy1Desc")
		Else
			Call Fn_Button_Click("Fn_ReqMgr_DetailTableSort", ObjSortWnd,"Cancel")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Invalid criteria pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
	End If
	'Spliting Third sort parameter
	sSortCriteria2=Split(strThenBy2,":")
	'Apllying Third sort criteria
	If sSortCriteria2(0)<>"" Then
		'Checking existance of Item in Sort item list
		bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_DetailTableSort", ObjSortWnd, "ThenBy2",sSortCriteria2(0))
		If bFlag=False Then
            Call Fn_Button_Click("Fn_ReqMgr_DetailTableSort", ObjSortWnd,"Cancel")
            Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Wrong Search Criteria"& sSortCriteria2(0) &"Pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
		'Selecting item
		Call Fn_List_Select("Fn_ReqMgr_DetailTableSort", ObjSortWnd, "ThenBy2",sSortCriteria2(0))
	End If
	'Selecting Sort type Ascending Or Descending
	If sSortCriteria2(1)<>"" Then
		If sSortCriteria2(1)="Asc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_ReqMgr_DetailTableSort",ObjSortWnd,"ThenBy2Asc")
		ElseIf	sSortCriteria2(1)="Desc" Then
			Call Fn_UI_JavaRadioButton_SetON("Fn_ReqMgr_DetailTableSort",ObjSortWnd,"ThenBy2Desc")
		Else
			Call Fn_Button_Click("Fn_ReqMgr_DetailTableSort", ObjSortWnd,"Cancel")
			Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: Invalid criteria pass")
			Set ObjSortWnd=Nothing
			Exit Function
		End If
	End If
	'Function Return True
	Fn_ReqMgr_DetailTableSort=True
	Call Fn_Button_Click("Fn_ReqMgr_DetailTableSort", ObjSortWnd,"OK")
	'Releasing Sort window object
	Set ObjSortWnd=Nothing
End Function

'-------------------------------------------------------------------------Function To Perform Operations on Collaboration Table--------------------------------------------------------------
'Function Name		:	Fn_ReqMgr_CollaborationTableNodeOpeations

'Description			:	This Function is used to Perform operations on Collaboration Table

'Parameters			:			1.	strAction:Action Name 			Eg -> "Select" and "Expand"
'											 2.	 strNodeName:Node name on which have to perform operation
											'3.	strColName:Column Name
											'4.strColValue:Column Value
											'5.strPopupMenu:PopUp Menu

'Return Value		:			True/False

'Pre-requisite		:		Collaboration Table Should be displayed  -> Call Fn_MenuOperation("Select","View:Show Collaboration Panel")
'											
'Examples			:		'Fn_ReqMgr_CollaborationTableNodeOpeations("Expand","000652-ReqSpec_RM22:000652/A;1-ReqSpec_RM22","", "","")
										'Fn_ReqMgr_CollaborationTableNodeOpeations("Select","000652-ReqSpec_RM22:000652/A;1-ReqSpec_RM22:View","", "","")
'History:
'										Developer Name					Date						Rev. No.			Changes Done			Reviewer	
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane			29/07/2010			            1.0						Created							Tushar B
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_CollaborationTableNodeOpeations(strAction,strNodeName,strColName, strColValue, strPopupMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_CollaborationTableNodeOpeations"
   'Declaring Variable
    Dim ObjCollabTable,ObjTable
    Dim IntRows ,StrNodePath, IntCounter,StrIndex,iRowNo
	'Function Return False
	Fn_ReqMgr_CollaborationTableNodeOpeations=False
	'Verify Allocation Table
    If Fn_UI_ObjectExist("Fn_ReqMgr_CollaborationTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("CollaborationTable"))=False Then
		'Invoking Collaboration Table
        Call Fn_MenuOperation("Select","View:Show Collaboration Panel")
	End If
	'Creating Object of Allocation Table
	Set ObjCollabTable=Fn_UI_ObjectCreate("Fn_ReqMgr_CollaborationTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet").JavaTable("CollaborationTable"))
	'Get the No. of rows present in the Allocation Table
	IntRows =Fn_UI_Object_GetROProperty("Fn_ReqMgr_CollaborationTableNodeOpeations",ObjCollabTable,"rows")
	'Creating Object of Attachments Table
	Set ObjTable =ObjCollabTable.Object
	If strNodeName<>"" Then
			'Get the Row No. of required Node
			For IntCounter = 0 to IntRows -1
                Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_CollaborationTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"CollaborationTable",IntCounter)
				StrNodePath = Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_CollaborationTableNodeOpeations",JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"CollaborationTable",IntCounter,0)
				StrNodePath =Mid(StrNodePath,Instr(1,StrNodePath,":")+1,Len(StrNodePath))
											
                If Trim(StrNodePath) = Trim(StrNodeName) Then
                	StrIndex = Cstr(IntCounter)
                    Exit For
				End If
			Next
			If Cint(IntCounter) = Cint(IntRows) Then
        		StrIndex = "FAIL:Node Not Found"
                Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail: "&  Trim(StrNodeName) &" Invalid Node Name") 
			End If
	End If
   	Select Case StrAction
			Case "Select"		
				iRowNo = StrIndex
				If isNumeric(iRowNo) Then
                    Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_CollaborationTableNodeOpeations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"CollaborationTable",iRowNo)
					Fn_ReqMgr_CollaborationTableNodeOpeations=True
				End if
			 Case "Expand"		
				iRowNo = StrIndex
				If isNumeric(iRowNo) Then
                    Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_CollaborationTableNodeOpeations", JavaWindow("RequirementsManager").JavaApplet("RMWindowApplet"),"CollaborationTable",iRowNo)
					Call Fn_MenuOperation("Select","View:Expand")
					Fn_ReqMgr_CollaborationTableNodeOpeations=True
				End if
	End Select
	'Release the Table object
	 Set ObjTable = Nothing
	 Set ObjCollabTable=Nothing
End Function
'-------------------------------------------------------------------------This Function is used to to Filter content of Details Table-------------------------------------------------------------------------------------
'Function Name		:	Fn_ReqMgr_DetailsTableFilterManagement

'Description			:	This Function is used to to Filter content of Details Table

'Parameters			:	   1.) sAction: Action Name
'									    2.) strConditionName: Condition Name ("Object == REQ-001140/A;1-Req")
'  										3.) strColName: Column Name to set condition (Object,Type,Group ID.................)
'										4.) strOperator: Operator to Set Condition (==,=,!=,<>.......................)
'										5.) strColValue:Column Velue to Set Condition (REQ-001140/A;1-Req..............)
'										6.)strLogicalType:Logical Type (And or OR....)

'Return Value		:	True/False

'Pre-requisite		:	Should be details table present
'											
'Examples			:	Fn_ReqMgr_DetailsTableFilterManagement("AddCondition","Type!=Requirement Revision","Type","!=","Requirement Revision","And")
'									
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   30/07/2010			              1.0										Created							Tushar B
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_DetailsTableFilterManagement(strAction,strConditionName,strColName,strOperator,strColValue,strLogicalType)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_DetailsTableFilterManagement"
   'Veriable Declaration
	Dim bFlag,iRowCnt,iCounter,strCellData
	Dim ObjAutoFilterWnd
	bFlag=False
	Fn_ReqMgr_DetailsTableFilterManagement=False
   'verifying existance of AutoFilter window
	If Fn_UI_ObjectExist("Fn_ReqMgr_DetailsTableFilterManagement",JavaWindow("RequirementsManager").JavaWindow("AutoFilter"))=False Then
		'Invoking AutoFilter window
		Call Fn_ToolbatButtonClick("Filter Management")
	End If
	'Creating objects 
	Set ObjAutoFilterWnd=Fn_UI_ObjectCreate("Fn_ReqMgr_DetailsTableFilterManagement",JavaWindow("RequirementsManager").JavaWindow("AutoFilter"))
	
	Select Case strAction
		'Case to Add New Condiation
		Case "AddCondition" 
					'Clicking plus button
		            Call Fn_Button_Click("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd, "PlusButton")
					'Setting Column Name Condition
					If strColName<>"" Then
						bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"ColumnList",strColName)
						If bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fali: "& strColName &"Column is not present in List" ) 
							Set ObjAutoFilterWnd=Nothing
							Exit Function
						End If
							Call Fn_List_Select("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"ColumnList",strColName)
					End If
					bFlag=False
					'Setting Operator Condiation
					If strOperator<>"" Then
						bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"OperatorList",strOperator)
						If bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fali: "& strOperator &"Column is not present in List" ) 
							Set ObjAutoFilterWnd=Nothing
							Exit Function
						End If
							Call Fn_List_Select("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"OperatorList",strOperator)
					End If
					bFlag=False
					'Setting Column Value Condition
					If strColValue<>"" Then
						bFlag=Fn_UI_ListItemExist("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"ObjectNameList",strColValue)
						If bFlag=False Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fali: "& strColValue &"Column is not present in List" ) 
							Set ObjAutoFilterWnd=Nothing
							Exit Function
						End If
							Call Fn_List_Select("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"ObjectNameList",strColValue)
					End If
					'Clicking Plus button to add condition into Table
					Call Fn_Button_Click("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"), "plusButton")
					bFlag=False
                    iRowCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaWindow("FilterConditionEditor").JavaTable("Table"),"rows")
					For iCounter=0 To iRowCnt-1
							strCellData=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_DetailsTableFilterManagement",ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"), "Table",iCounter,0)
							If Trim(Replace(strCellData," ",""))=Trim(Replace(strConditionName," ","")) Then
								iRowCnt=iCounter
								bFlag=True
								Exit For
							End If
					Next
					If bFlag=True Then
                        Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_DetailsTableFilterManagement",ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"), "Table",iRowCnt)
					Else
						Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fali: Invalid parameter pass" ) 
						Set ObjAutoFilterWnd=Nothing
						Exit Function
					End If
					'Clicking ok to Apply condition
					Call Fn_Button_Click("Fn_ReqMgr_DetailsTableFilterManagement",ObjAutoFilterWnd.JavaWindow("FilterConditionEditor"),"OK")

					 iRowCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd.JavaTable("ConditionTable"),"rows")
					For iCounter=0 To iRowCnt-1
							strCellData=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_DetailsTableFilterManagement",ObjAutoFilterWnd,"ConditionTable",iCounter,0)
							If Trim(Replace(strCellData," ",""))=Trim(Replace(strConditionName," ","")) Then
                                JavaWindow("RequirementsManager").JavaWindow("AutoFilter").JavaTable("ConditionTable").ActivateCell iCounter,0
								Call Fn_UI_JavaTable_SelectRow("Fn_ReqMgr_DetailsTableFilterManagement",ObjAutoFilterWnd, "ConditionTable",iCounter)
								Exit For
							End If
					Next
	End Select
	Fn_ReqMgr_DetailsTableFilterManagement=True
	'Closing Window
	Call Fn_Button_Click("Fn_ReqMgr_DetailsTableFilterManagement", ObjAutoFilterWnd, "Close")
	'Releasing object Auto Filter Window
	Set ObjAutoFilterWnd=Nothing
End Function

'-------------------------------------------------------------------------This Function is used to to perform operation on MS Word tab---------------------------------------------------------------
'Function Name		:	Fn_ReqMgr_MSWordTabOperationsExt

'Description			:	This Function is used to to perform operation on MS Word tab

'Parameters			:			1.	strAction:Action Name
'											 2.	 strValue:Value to set in text Box Or to verify the value
											'3.	strParameterName:Parameter Name

'Return Value		:	True/False

'Pre-requisite		:	Object should be selected
'											
'Examples			:	Fn_ReqMgr_MSWordTabOperationsExt("VerifyValue","value1","parameter1")
'									
'									
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   15/11/2010			              1.0										Created							Tushar B
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_MSWordTabOperationsExt(strAction,strValue,strParameterName)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_MSWordTabOperationsExt"
	'Variable declaration
   Dim bReturn,bFlag,iRowCnt,sCellData,sValueName,iCount
   'Function Return False
   Fn_ReqMgr_MSWordTabOperationsExt=False
   bReturn=False
	   'Activating MS Word tab
	   Call Fn_SetView("Teamcenter:MS Word")
	Select Case strAction
		Case "VerifyValue"	'Fn_ReqMgr_MSWordTabOperationsExt("VerifyValue","value1","parameter1")
				'Taking no of rows from ParametricValues table
				iRowCnt=Fn_UI_Object_GetROProperty("Fn_ReqMgr_MSWordTabOperationsExt",JavaWindow("RequirementsManager").JavaTable("ParametricValues"),"rows")
				For iCount=0 To iRowCnt-1
						'Taking data from ParametricValues table
						sCellData=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_MSWordTabOperationsExt",JavaWindow("RequirementsManager"), "ParametricValues",iCount,0)
						If strParameterName=sCellData Then
								sValueName=Fn_UI_JavaTable_GetCellData("Fn_ReqMgr_MSWordTabOperationsExt",JavaWindow("RequirementsManager"), "ParametricValues",iCount,1)
								If strValue=sValueName Then
									Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Pass:Successfully verify  value"& strValue )   
									bReturn=True
									Exit For
								End If
						End If
				Next

				If bReturn=False Then
                    Call Fn_WriteLogFile(Environment.Value("TestLogFile") ,"Fail:"& strParameterName & "Parameter is not present")   
					Exit Function
				Else
					  Fn_ReqMgr_MSWordTabOperationsExt=True
				End If
	End Select
	wait(2)
	JavaWindow("RequirementsManager").JavaTab("MSWord").CloseTab "MS Word"
End Function


'-------------------------------------------------------------------------This Function is used to to perform operation on Save Column Application menu--------------------------------------------------------
'Function Name		:	Fn_ReqMgr_SaveColumnConfiguration

'Description			:	This Function is used to to perform operation on Save Column Application menu

'Parameters			:	   1.) strConfigName: Name of Configuration (It should be unique in ConfigurationSaveAs and Add case )
'										2.)strConfigDesc:Description of Configuration

'Return Value		:	True/False

'Pre-requisite		:	Should be details table present
'											
'Examples			:	Fn_ReqMgr_SaveColumnConfiguration("TestConfig","Test Configuration")
'								
'History:
'									Developer Name												Date								Rev. No.						Changes Done				Reviewer	
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'									Sandeep Navghane									   19/08/2010			              1.0										Created							Tushar B
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_ReqMgr_SaveColumnConfiguration(strConfigName,strConfigDesc)
	GBL_FAILED_FUNCTION_NAME="Fn_ReqMgr_SaveColumnConfiguration"
  'Declaring Variables
   Dim ObjColumnConfigWnd
   Fn_ReqMgr_SaveColumnConfiguration=False
   'verifying existance of ApplyColumnConfiguration window
	If Fn_UI_ObjectExist("Fn_ReqMgr_SaveColumnConfiguration",JavaWindow("RequirementsManager").JavaWindow("SaveColumnConfiguration"))=False Then
		'Invoking ApplyColumnConfiguration window
		Call Fn_ReqMgr_DetailTableOperation("PopUpMenuSelect","", "", "","Save Column Configuration...")
	End If
	'Creating objects 	
	Set ObjColumnConfigWnd=JavaWindow("RequirementsManager").JavaWindow("SaveColumnConfiguration")
	'Setting Configuration Name
    Call Fn_UI_EditBox_Type("Fn_ReqMgr_SaveColumnConfiguration",ObjColumnConfigWnd,"Name",strConfigName)
	If strConfigDesc<>"" Then
		'Setting Description of configuration 
		Call Fn_Edit_Box("Fn_ReqMgr_SaveColumnConfiguration",ObjColumnConfigWnd,"Description",strConfigDesc)
	End If
	'Clicking on save button to create configuration
	Call Fn_Button_Click("Fn_ReqMgr_SaveColumnConfiguration",ObjColumnConfigWnd,"Save")
    Fn_ReqMgr_SaveColumnConfiguration=True
    'Releasing all objects
	Set ObjColumnConfigWnd=Nothing
End Function
