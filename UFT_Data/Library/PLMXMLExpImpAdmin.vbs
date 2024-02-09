Option Explicit
'---------------------------------'Global variables for Teamcenter Perspective Names-----------------------------------------------------
Public GBL_PERSPECTIVE_PLM_XML_EXPORT_IMPORT_ADMINISTRATION
GBL_PERSPECTIVE_PLM_XML_EXPORT_IMPORT_ADMINISTRATION = "PLM XML Export Import Administration"
'---------------------------------'Global variables for Teamcenter Perspective Names-----------------------------------------------------
'=======================================================================================================================================================
' Function List
'-------------------------------------------------------------------------------------------------------------------------------------------------------
'0) 	Fn_SISW_PLMXML_GetObject()
'1.)	Fn_PLM_ImportExportModes_TreeOpeartion()
'2.)	Fn_PLM_PLMXMLList_TreeOpeartion()
'3.)	Fn_PLM_ClosureRuleOperations()
'4.)	Fn_PLM_TransferModeOperations()
'5.)	Fn_PLM_ExportObjects()
'6.)	Fn_PLM_StringViewer_Operations()
'7.)	Fn_PLM_ActionRule()
'8.)	Fn_PLM_FilterRuleOperations()
'9.)	Fn_PLM_PropertySetOperations()
'10.)	Fn_PLM_TransferOptionSetOperations()
'11.)	Fn_PLM_ImportObjects()
'12.)   Fn_PLM_PropertySetOperations_Ext()
'13.)   Fn_PLM_FilterRuleOperations_Ext()
'14.)   Fn_PLM_TransferOptionSetOperations_Ext()
'15.)   Fn_PLM_ClosureRuleOperations_Ext()
'16.)   Fn_PLM_RenameItemsFromXMLFile()
'17.)   Fn_PLM_VerifyValueInTag()
'=======================================================================================================================================================
'****************************************    Function to get Object hierarchy ***************************************
'
''Function Name		 	:	Fn_SISW_PLMXML_GetObject
'
''Description		    :  	Function to get Object hierarchy

''Parameters		    :	1. sObjectName : Object Handle name
								
''Return Value		    :  	Object \ Nothing
'
''Examples		     	:	Fn_SISW_PLMXML_GetObject("Export Completed")

'History:
'	Developer Name			Date				Rev. No.		Reviewer		Changes Done	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Pranav Ingle		 		11-Jul-2012			1.0			
'Shreyas                            20-08-2012		  1.1	
'-----------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------
'	Anurag Khera		 26-Sept-2013		2.0								Externalized object hierarchies
'-----------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_SISW_PLMXML_GetObject(sObjectName)

	Dim sObjectXMLPath
	sObjectXMLPath = Fn_GetEnvValue("User", "AutomationDir") & "\TestData\AutomationXML\ObjectXML\PLMXMLExpImpAdmin.xml"
	Set Fn_SISW_PLMXML_GetObject = Fn_SISW_Setup_GetObjectFromXML(sObjectXMLPath, sObjectName)

End Function

''*********************************************************		Function to perform action on Project Tree	***********************************************************************
'Function Name		:				Fn_PLM_ImportExportModes_TreeOpeartion()

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

'Examples				:				
											'Msgbox Fn_PLM_ImportExportModes_TreeOpeartion("Select","PLM XML Import Export Modes:ConfiguredRequirementDataExport", "")
											'Msgbox Fn_PLM_ImportExportModes_TreeOpeartion("Expand","PLM XML Import Export Modes:ConfiguredRequirementDataExport", "")
											'Msgbox Fn_PLM_ImportExportModes_TreeOpeartion("Collapse","PLM XML Import Export Modes:ConfiguredRequirementDataExport", "")
											'Msgbox Fn_PLM_ImportExportModes_TreeOpeartion("Exist","PLM XML Import Export Modes:ConfiguredRequirementDataExport", "")
											'Msgbox Fn_PLM_ImportExportModes_TreeOpeartion("GetIndex","PLM XML Import Export Modes:ConfiguredRequirementDataExport", "")
  												
'History					 :		
' 											Developer Name												Date						Rev. No.						Changes Done						Reviewer
'										------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Ketan Raje													23-11-10						1.0																				Harshal	
'											Pranav S														  28-May-12									Added Case: "Double Click"						Sunny
'										------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Function Fn_PLM_ImportExportModes_TreeOpeartion(sAction,sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_ImportExportModes_TreeOpeartion"
	Dim objJavaWindowPLM, objJavaTreePLM, intNodeCount, intCount, sTreeItem, aMenuList
	Set objJavaWindowPLM = Fn_UI_ObjectCreate( "Fn_PLM_ImportExportModes_TreeOpeartion",JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))

	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"					
                    Call Fn_JavaTree_Select("Fn_PLM_ImportExportModes_TreeOpeartion", objJavaWindowPLM, "PIETree",sNodeName)
					Fn_PLM_ImportExportModes_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_PLM_ImportExportModes_TreeOpeartion",objJavaWindowPLM,"PIETree",sNodeName)
					Fn_PLM_ImportExportModes_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_PLM_ImportExportModes_TreeOpeartion", objJavaWindowPLM,"PIETree",sNodeName)
					Fn_PLM_ImportExportModes_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreePLM = Fn_UI_ObjectCreate( "Fn_PLM_ImportExportModes_TreeOpeartion", objJavaWindowPLM.JavaTree("PIETree"))
					intNodeCount = objJavaTreePLM.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreePLM.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_PLM_ImportExportModes_TreeOpeartion = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_PLM_ImportExportModes_TreeOpeartion = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowPLM.JavaTree("PIETree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowPLM.JavaTree("PIETree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_PLM_ImportExportModes_TreeOpeartion = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_PLM_ImportExportModes_TreeOpeartion = FALSE
				End If
		' ----------------------------------------------------------------------- Double Click on a particular node-------------------------------------------------------------------------
		Case "DoubleClick"
				objJavaWindowPLM.JavaTree("PIETree").Activate sNodeName
				Call Fn_ReadyStatusSync(1)
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully DoubleClick Node [" + sNodeName + "] of PLM_ImportExportModes_Tree")
				Fn_PLM_ImportExportModes_TreeOpeartion = TRUE
		Case Else
						Fn_PLM_ImportExportModes_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_ImportExportModes_TreeOpeartion function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_PLM_ImportExportModes_TreeOpeartion")
	Set objJavaWindowPLM = nothing
	Set objJavaTreePLM = nothing
End Function

''*********************************************************		Function to perform action on Project Tree	***********************************************************************
'Function Name		:				Fn_PLM_PLMXMLList_TreeOpeartion()

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

'Examples				:			  'Case "Select" : Msgbox Fn_PLM_PLMXMLList_TreeOpeartion("Select","ClosureRule", "")											
											'Case "Expand" : Msgbox Fn_PLM_PLMXMLList_TreeOpeartion("Expand","ClosureRule", "")
											'Case "Collapse" : Msgbox Fn_PLM_PLMXMLList_TreeOpeartion("Collapse","ClosureRule", "")
											'Case "Exist" : Msgbox Fn_PLM_PLMXMLList_TreeOpeartion("Exist","ClosureRule:AccountabilityAll", "")
											'Case "GetIndex" : Msgbox Fn_PLM_PLMXMLList_TreeOpeartion("GetIndex","ClosureRule:AccountabilityAll", "")
  												
'History					 :		
' 											Developer Name												Date						Rev. No.						Changes Done						Reviewer
'										------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'											Ketan Raje													23-11-10						1.0																				Harshal	
'										------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function Fn_PLM_PLMXMLList_TreeOpeartion(sAction,sNodeName, sMenu)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_PLMXMLList_TreeOpeartion"
	Dim objJavaWindowPLM, objJavaTreePLM, intNodeCount, intCount, sTreeItem, aMenuList
	Set objJavaWindowPLM = Fn_UI_ObjectCreate( "Fn_PLM_PLMXMLList_TreeOpeartion",JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
	sNodeName = "#0:"+sNodeName
	Select Case sAction
		'----------------------------------------------------------------------- For selecting single node -------------------------------------------------------------------------
		Case "Select"					
                    Call Fn_JavaTree_Select("Fn_PLM_PLMXMLList_TreeOpeartion", objJavaWindowPLM, "PIEListTree",sNodeName)
					Fn_PLM_PLMXMLList_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For expanding a particular  node-------------------------------------------------------------------------
		Case "Expand"
                    Call Fn_UI_JavaTree_Expand("Fn_PLM_PLMXMLList_TreeOpeartion",objJavaWindowPLM,"PIEListTree",sNodeName)
					Fn_PLM_PLMXMLList_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For collapssing a particular  node-------------------------------------------------------------------------
		Case "Collapse"
                    Call Fn_UI_JavaTree_Collapse("Fn_PLM_PLMXMLList_TreeOpeartion", objJavaWindowPLM,"PIEListTree",sNodeName)
					Fn_PLM_PLMXMLList_TreeOpeartion = TRUE
		'----------------------------------------------------------------------- For Checking existance of a particular  node-------------------------------------------------------------------------
		Case "Exist"
				Set objJavaTreePLM = Fn_UI_ObjectCreate( "Fn_PLM_PLMXMLList_TreeOpeartion", objJavaWindowPLM.JavaTree("PIEListTree"))
					intNodeCount = objJavaTreePLM.GetROProperty ("items count") 
					For intCount = 0 to intNodeCount - 1
						sTreeItem = objJavaTreePLM.GetItem(intCount)
						If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
							Fn_PLM_PLMXMLList_TreeOpeartion = TRUE	
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree exist")
							Exit For
						End If
					Next
					If Cstr(intCount) = intNodeCount Then
						Fn_PLM_PLMXMLList_TreeOpeartion = FALSE						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "[" + sNodeName + "] of JavaTree does not exist")
						Exit Function
					End If
		'----------------------------------------------------------------------- Get Index value of a particular node-------------------------------------------------------------------------
		Case "GetIndex"
				bFlag = False
				For intCount=0 to objJavaWindowPLM.JavaTree("PIEListTree").GetROProperty ("items count")-1
					sTreeItem = objJavaWindowPLM.JavaTree("PIEListTree").GetItem (intCount)
					If Trim(lcase(sTreeItem)) = Trim(Lcase(sNodeName)) Then
						Fn_PLM_PLMXMLList_TreeOpeartion = intCount
						bFlag = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The index of the given node is "&intCount)
						Exit For
					End If
				Next
				If bFlag = False Then
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The given node does not exist")
					Fn_PLM_PLMXMLList_TreeOpeartion = FALSE
				End If

		Case Else
						Fn_PLM_PLMXMLList_TreeOpeartion = FALSE
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_PLMXMLList_TreeOpeartion function failed")
						Exit Function
	End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed on Node [" + sNodeName + "] of JavaTree of function Fn_PLM_PLMXMLList_TreeOpeartion")
	Set objJavaWindowPLM = nothing
	Set objJavaTreePLM = nothing
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_PLM_ClosureRuleOperations(sAction, sRuleName, sDescription, sScope, sSchema)
'###
'###    DESCRIPTION        :   Add / Modify / Delete / Exist
'###	Prequisite 				:	PLM XML Export Import Administration Prespective is Open.
'###
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Edit_Box(), Fn_UI_JavaRadioButton_SetON(), Fn_Button_Click
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    	24/11/2010         1.0
'###
'###    REVIWED BY     :    Harshal
'###
'###    EXAMPLE          : 		Case "Create" : Msgbox Fn_PLM_ClosureRuleOperations("Create", "HarshalTest", "Testing", "Export", "PLMXML")
'###									Case "Modify" : Msgbox Fn_PLM_ClosureRuleOperations("Modify", "", "Test", "Import", "PLMXML")
'###									Case "Delete" : Msgbox Fn_PLM_ClosureRuleOperations("Delete", "", "", "", "")
'###									Case "AddRow" : Msgbox Fn_PLM_ClosureRuleOperations("AddRow", "Primary Object Class Type:Primary Object:Secondary Object Class Type:Secondary Object:Relation Type:Related Property Or Object:Action Type:Conditional Clause", "TYPE:abc:TYPE:def:RELATIONP2S:xyz:PROCESS:no", "", "")
'###									Case "ModifyRow" : Msgbox Fn_PLM_ClosureRuleOperations("ModifyRow", "Action Type", "PROCESS+TRAVERSE", "", "")
'#############################################################################################################
Public Function Fn_PLM_ClosureRuleOperations(sAction, sRuleName, sDescription, sScope, sSchema)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_ClosureRuleOperations"
	Dim objClosureRule, iReturn, aColumns, aColsData, iCols, iRows, iRowCnt, iFlag, intCount, iCount, iCounter
	Fn_PLM_ClosureRuleOperations = False
	Set objClosureRule = Fn_UI_ObjectCreate("Fn_PLM_ClosureRuleOperations", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Rule Name
						If sRuleName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_ClosureRuleOperations",objClosureRule,"TraversalRuleName",sRuleName)
						End If
						'Set Value for Description
						If sDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_ClosureRuleOperations",objClosureRule,"Description",sDescription)
						End If
						'Set the Scope of Traversal.
						If sScope<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_PLM_ClosureRuleOperations",objClosureRule, sScope)
						End If
						'Set Output Schema format
						If sSchema<>"" Then
							iReturn = objClosureRule.JavaList("SchemaFormat").GetItemIndex(sSchema)
							objClosureRule.JavaList("SchemaFormat").Object.setSelectedIndex iReturn
						End If
						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations", objClosureRule, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations", objClosureRule, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations", objClosureRule, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "AddRow","ModifyRow"
					iFlag = 0
					If sAction="AddRow" Then
						'Click on Add button.
						Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations", objClosureRule, "AddColumn")
					End If
							aColumns = Split(sRuleName,":",-1,1)
							aColsData = Split(sDescription,":",-1,1)
							iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
							iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
										'Logic for Setting Data
										If aColumns(iCounter)="Primary Object Class Type" OR aColumns(iCounter)="Secondary Object Class Type" OR aColumns(iCounter)="Relation Type" Then										
											'Select value from List
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
										ElseIf aColumns(iCounter)="Action Type" Then
'											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").Set ""
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaEdit("ClosureTableEdit").Set aColsData(iCounter)
										Else
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),""
											'Set value in Cell.
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
										End If
										Exit For
									End If									
								Next
							Next
							'Click on Modify button.
							Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations", objClosureRule, "Modify")
							If iFlag=Ubound(aColumns)+1 Then
								Fn_PLM_ClosureRuleOperations = True
							Else
								Fn_PLM_ClosureRuleOperations = False
							End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_ClosureRuleOperations function failed")
						Fn_PLM_ClosureRuleOperations = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_ClosureRuleOperations")
	Fn_PLM_ClosureRuleOperations = TRUE
	Set objClosureRule = nothing 
End Function

'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_PLM_TransferModeOperations(sAction, sTMName, sTMContext, sTMDescription, sTMTransferType, sTMSchema, sTMSupport, sTMClosureRule, sTMFilterRule, sTMPropertySet, sTMRevisionRule, sTMActionList)
'###
'###    DESCRIPTION        :   Add / Modify / Delete / Exist
'###	Prequisite 				:	PLM XML Export Import Administration Prespective is Open.
'###
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Edit_Box(), Fn_UI_JavaRadioButton_SetON(), Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    	26/11/2010         1.0
'###
'###    REVIWED BY     :    Harshal
'###
'###    EXAMPLE          : 		Case "Create" : Msgbox Fn_PLM_TransferModeOperations("Create", "TestTM", "", "", "", "PLMXML", "", "HarshalTest", "", "", "", "test123456|TestActionRule")
'###									Case "Modify" : Msgbox Fn_PLM_TransferModeOperations("Modify", "TestTM", "", "", "", "PLMXML", "", "ExportActivities", "", "", "", "test123456|TestActionRule")
'###									Case "Delete" : Msgbox Fn_PLM_TransferModeOperations("Delete", "", "", "", "", "", "", "", "", "", "", "")
'###									Case "Verify" : Msgbox Fn_PLM_TransferModeOperations("Verify", "ConfiguredDataImportDefault", "DEFAULT_PIE_CONTEXT_STRING", "", "Import", "PLMXML", "ON", "ConfiguredDataImportDefault", "ConfiguredDataImportDefault", "ConfiguredDataImportDefault", "", "DefinedTools|RDV_IMPORT_TR_VEH_POST_ACTION")
'#############################################################################################################
Public Function Fn_PLM_TransferModeOperations(sAction, sTMName, sTMContext, sTMDescription, sTMTransferType, sTMSchema, sTMSupport, sTMClosureRule, sTMFilterRule, sTMPropertySet, sTMRevisionRule, sTMActionList)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_TransferModeOperations"
	Dim objTransferMode, iReturn, iFlag, intCount, iCount, iCounter, aTMActionList, aTMSepActionList, iCnt
	Fn_PLM_TransferModeOperations = False
	Set objTransferMode = Fn_UI_ObjectCreate("Fn_PLM_TransferModeOperations", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Transfer Mode Name
						If sTMName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_TransferModeOperations",objTransferMode,"TMName",sTMName)
						End If
						'Set Value for Context
						If sTMContext<>"" Then
							Call Fn_Edit_Box("Fn_PLM_TransferModeOperations",objTransferMode,"TMContext",sTMContext)
						End If
						'Set Value for Description
						If sTMDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_TransferModeOperations",objTransferMode,"Description",sTMDescription)
						End If
						'Set the Type of Transfer
						If sTMTransferType<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_PLM_TransferModeOperations",objTransferMode, sTMTransferType)
						End If
						'Set Output Schema format
						If sTMSchema<>"" Then
							iReturn = objTransferMode.JavaList("SchemaFormat").GetItemIndex(sTMSchema)
							objTransferMode.JavaList("SchemaFormat").Object.setSelectedIndex iReturn
						End If
						'Set the Support Incremental
						If sTMSupport<>"" Then
							Call Fn_CheckBox_Set("Fn_PLM_TransferModeOperations", objTransferMode, "TMSupportIncremental", sTMSupport)
						End If
						'Set the Closure Rule
						If sTMClosureRule<>"" Then
							Call Fn_List_Select("Fn_PLM_TransferModeOperations", objTransferMode, "TMClosureRuleList",sTMClosureRule)
						End If
						'Set the Filter Rule
						If sTMFilterRule<>"" Then
							Call Fn_List_Select("Fn_PLM_TransferModeOperations", objTransferMode, "TMFilterRuleList",sTMFilterRule)
						End If
						'Set the Property Set
						If sTMPropertySet<>"" Then
							Call Fn_List_Select("Fn_PLM_TransferModeOperations", objTransferMode, "TMPropertySetList",sTMPropertySet)
						End If
						'Set the Revision Rule
						If sTMRevisionRule<>"" Then
							Call Fn_List_Select("Fn_PLM_TransferModeOperations", objTransferMode, "TMRevisionRuleList",sTMRevisionRule)
						End If
						'Transfer tools from Defined list to selected list
						If sTMActionList<>"" Then
							aTMActionList = Split(sTMActionList,"|",-1,1)
							iReturn = Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaList("TMDefinedToolsList"), "items count"))
							For iCount = 0 to Ubound(aTMActionList)								
								For iCnt = 0 to iReturn-1
									If Trim(Lcase(objTransferMode.JavaList("TMDefinedToolsList").GetItem(iCnt))) = Trim(Lcase(aTMActionList(iCount))) Then
										'Select item from Defined tools list
										objTransferMode.JavaList("TMDefinedToolsList").Select iCnt
										'Click on AddColumn button
										Call Fn_Button_Click("Fn_PLM_TransferModeOperations", objTransferMode, "AddColumn")
										Exit For
									End If
								Next
							Next
						End If
						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_TransferModeOperations", objTransferMode, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_TransferModeOperations", objTransferMode, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_TransferModeOperations", objTransferMode, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_TransferModeOperations", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "Verify"
						iCount = 0
						iCounter = 0
						'Verify TM Name
						If sTMName<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferModeOperations",objTransferMode,"TMName"))) = Trim(Lcase(sTMName)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify TM Context
						If sTMContext<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferModeOperations",objTransferMode,"TMContext"))) = Trim(Lcase(sTMContext)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify TM Description
						If sTMDescription<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferModeOperations",objTransferMode,"Description"))) = Trim(Lcase(sTMDescription)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Type of Transfer
						If sTMTransferType<>"" Then
							iCount = iCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaRadioButton(sTMTransferType), "value")) = 1 Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Output Scheme Format
						If sTMSchema<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("SchemaFormat").GetItem(objTransferMode.JavaList("SchemaFormat").Object.getSelectedIndex))) = Trim(Lcase(sTMSchema)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Support Increment
						If sTMSupport<>"" Then
							iCount = iCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaCheckBox("TMSupportIncremental"), "value")) = 1 Then
								iCounter = iCounter + 1
							End If							
						End If
						'Verify Closure Rule
						If sTMClosureRule<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("TMClosureRuleList").GetItem(objTransferMode.JavaList("TMClosureRuleList").Object.getSelectedIndex))) = Trim(Lcase(sTMClosureRule)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Filter Rule
						If sTMFilterRule<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("TMFilterRuleList").GetItem(objTransferMode.JavaList("TMFilterRuleList").Object.getSelectedIndex))) = Trim(Lcase(sTMFilterRule)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Property Set
						If sTMPropertySet<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("TMPropertySetList").GetItem(objTransferMode.JavaList("TMPropertySetList").Object.getSelectedIndex))) = Trim(Lcase(sTMPropertySet)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Revision Rule
						If sTMRevisionRule<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("TMRevisionRuleList").GetItem(objTransferMode.JavaList("TMRevisionRuleList").Object.getSelectedIndex))) = Trim(Lcase(sTMRevisionRule)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Action List
						If sTMActionList<>"" Then
								aTMActionList = Split(sTMActionList,",",-1,1)
								For intCount = 0 to Ubound(aTMActionList)
									iCount = iCount + 1
									aTMSepActionList = Split(aTMActionList(intCount),"|",-1,1)
									If Trim(Lcase(aTMSepActionList(0))) = "definedtools" Then								
										iReturn = Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaList("TMDefinedToolsList"), "items count"))
										For iCnt = 0 to iReturn-1
											If Trim(Lcase(JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("TMDefinedToolsList").GetItem(iCnt))) = Trim(Lcase(aTMSepActionList(1))) Then
												iCounter = iCounter + 1
												Exit For
											End If
										Next
									ElseIf Trim(Lcase(aTMSepActionList(0))) = "selectedtools" Then
										iReturn = Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaList("TMSelectedToolsList"), "items count"))
										For iCnt = 0 to iReturn-1
											If Trim(Lcase(JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("TMSelectedToolsList").GetItem(iCnt))) = Trim(Lcase(aTMSepActionList(1))) Then
												iCounter = iCounter + 1
												Exit For
											End If
										Next
									End If
								Next
						End If
							'The part for Action List is to be coded as required
							If iCount=iCounter Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : All Given values match actual values")
								Fn_PLM_TransferModeOperations = TRUE
								Set objTransferMode = nothing 
								Exit Function
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : All Given values does not match actual values")
								Fn_PLM_TransferModeOperations = FALSE
								Set objTransferMode = nothing 
								Exit Function
						End If					


Case "VerifyBlank"
						iCount = 0
						iCounter = 0
						'Verify TM Name
						If sTMName<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferModeOperations",objTransferMode,"TMName"))) = Trim(Lcase(sTMName)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify TM Context
						If sTMContext<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferModeOperations",objTransferMode,"TMContext"))) = Trim(Lcase(sTMContext)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify TM Description
						If sTMDescription<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferModeOperations",objTransferMode,"Description"))) = Trim(Lcase(sTMDescription)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Type of Transfer
						If sTMTransferType<>"" Then
							iCount = iCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaRadioButton(sTMTransferType), "value")) = 1 Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Output Scheme Format
						If sTMSchema<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("SchemaFormat").GetItem(objTransferMode.JavaList("SchemaFormat").Object.getSelectedIndex))) = Trim(Lcase(sTMSchema)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Support Increment
						If sTMSupport<>"" Then
							iCount = iCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaCheckBox("TMSupportIncremental"), "value")) = 0 Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Closure Rule
						If sTMClosureRule<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("TMClosureRuleList").GetItem(objTransferMode.JavaList("TMClosureRuleList").Object.getSelectedIndex))) = Trim(Lcase(sTMClosureRule)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Filter Rule
						If sTMFilterRule<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("TMFilterRuleList").GetItem(objTransferMode.JavaList("TMFilterRuleList").Object.getSelectedIndex))) = Trim(Lcase(sTMFilterRule)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Property Set
						If sTMPropertySet<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("TMPropertySetList").GetItem(objTransferMode.JavaList("TMPropertySetList").Object.getSelectedIndex))) = Trim(Lcase(sTMPropertySet)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Revision Rule
						If sTMRevisionRule<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTransferMode.JavaList("TMRevisionRuleList").GetItem(objTransferMode.JavaList("TMRevisionRuleList").Object.getSelectedIndex))) = Trim(Lcase(sTMRevisionRule)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Action List
						If sTMActionList<>"" Then
								aTMActionList = Split(sTMActionList,",",-1,1)
								For intCount = 0 to Ubound(aTMActionList)
									iCount = iCount + 1
									aTMSepActionList = Split(aTMActionList(intCount),"|",-1,1)
									If Trim(Lcase(aTMSepActionList(0))) = "definedtools" Then								
										iReturn = Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaList("TMDefinedToolsList"), "items count"))
											If  iReturn = 0 Then
												iCounter = iCounter + 1
												Exit For
											End If
									End If
									iCount = iCount + 1
									If Trim(Lcase(aTMSepActionList(1))) = "selectedtools" Then								
										iReturn = Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferModeOperations",objTransferMode.JavaList("TMSelectedToolsList"), "items count"))
											If  iReturn = 0 Then
												iCounter = iCounter + 1
												Exit For
											End If
									End If
								Next
						End If
							'The part for Action List is to be coded as required
							If iCount=iCounter Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : All Given values match actual values")
								Fn_PLM_TransferModeOperations = TRUE
								Set objTransferMode = nothing 
								Exit Function
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : All Given values does not match actual values")
								Fn_PLM_TransferModeOperations = FALSE
								Set objTransferMode = nothing 
								Exit Function
						End If
					
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Fn_PLM_TransferModeOperations function failed")
						Fn_PLM_TransferModeOperations = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_TransferModeOperations")
	Fn_PLM_TransferModeOperations = TRUE
	Set objTransferMode = nothing 
End Function
'----------------------------------------------------------------------------------------------Function for PLMXML Export of a file.----------------------------------------------------------------------------------------------------------------------------
'Function Name		  :	  Fn_PLM_ExportObjects

'Description			 :	 Function to Export Object to Word

'Parameters			   :	sAction, sExportType, sExportDirectory, sExportFilename, sTransferMode, sRevisionRule, sLanguages, bExportInBack, bOpenFile, sObjectList, aGlobalDictionary, sButtons

'Return Value		   : 	True Or False

'Pre-requisite			:	Object should be selected which have to export

'Examples				:	Msgbox Fn_PLM_ExportObjects("SetPLMXMLExport", "PLMXML", "C:\auto_tc\TC9\pnv6s169\20101220\win\rac", "Ketan.xml", "ConfiguredDataExportDefault", "Latest Working", "", "OFF", "OFF", "", "Yes", "", "OK")

'History			   :					Developer Name												Date						Rev. No.	
'---------------------------------------------------------------------------------------------------------------------------------------------
'											Ketan Raje										   			18/01/2011			           1.0		
'  										    Sushma Pagare								   		15/09/2011			           1.0		   Added Case "SetToPLMXMLExport" for menu Tools->Export-> To PLMXML
'  										    Ashok kakade								   		20/06/2012			           1.0		   Modified hierachy of Dialog Export Completed 
'											Vrushali Wani										05/12/2012								added case 'VerifyAndSetToPLMXMLExport' and implemented code for ObjectList	
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PLM_ExportObjects(sAction, sExportType, sExportDirectory, sExportFilename, sTransferMode, sRevisionRule, sLanguages, bExportInBack, bOpenFile, sObjectList, bViewLog, aGlobalDictionary, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_ExportObjects"
   'Declaring Variables
   Dim ObjExport, aButtons, iCount,objFSO,objExportcomp
   Dim WshShl,Shell,sDirPath,sXmlPath,sValue,bReturn
   Dim iRows, sObjListValue, bFlag
   'Setting bReturn to False
   bReturn=False
   'Initially Function returns False
   Fn_PLM_ExportObjects=False

   Select Case sAction
	 	Case "SetPLMXMLExport",   "SetToPLMXMLExport",  "VerifyAndSetToPLMXMLExport"
					If  sAction = "SetToPLMXMLExport" OR sAction = "VerifyAndSetToPLMXMLExport" Then          '' Case added for menu Tools->Export-> To PLMXML
								'Checking Existence of PLMXML Export... Window
							   If Fn_UI_ObjectExist("Fn_PLMXML_Export",JavaWindow("PLMXML-TeamCenter").JavaWindow("PLMXMLExport"))=False Then
								   'Opening PLMXML Export... Window
								   	If Fn_MenuOperation("Exist","Tools:Export:To PLMXML") Then
										Call Fn_MenuOperation("Select","Tools:Export:To PLMXML")
									Else
										Call Fn_KeyBoardOperation("SendKeys","{ESC}")
										Call Fn_KeyBoardOperation("SendKeys","{ESC}")
										Call Fn_MenuOperation("WinMenuSelect","Tools:Export:To PLMXML...")
									End If
							   End If							
							   'Creating object of PLMXML Export... Window	
							   Set ObjExport= Fn_UI_ObjectCreate("Fn_PLMXML_Export",JavaWindow("PLMXML-TeamCenter").JavaWindow("PLMXMLExport"))						   

					ElseIf sAction = "SetPLMXMLExport" Then          ''Case for Tools -> Export ->Objects   
							'Checking Existance of ExportToWord Window
							   If Fn_UI_ObjectExist("Fn_PLM_ExportObjects",JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export"))=False Then
								   'Opening ExportToWord Window
									Call Fn_MenuOperation("Select","Tools:Export:Objects...")
							   End If
							  For iCount=0 to 0
								 JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export").SetTOProperty "title","Export ..."
								 If JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export").Exist Then
									 bReturn = True
									 Exit For
								 End If
								 JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export").SetTOProperty "title","PLM XML Export ..."
								 If JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export").Exist Then
									Exit For
								 End If
							  Next  
							   'Creating object of ExportToWord Window	
							Set ObjExport= Fn_UI_ObjectCreate("Fn_PLM_ExportObjects",JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export"))
							If bReturn=True Then
								'Select the Export Type
								ObjExport.JavaCheckBox("ExportType").SetTOProperty "attached text",sExportType
								If sExportType<>"" Then
									Call Fn_CheckBox_Select("Fn_PLM_ExportObjects", ObjExport, "ExportType")
								End If
								'SET property for Export dialog
								ObjExport.SetTOProperty "title","PLM XML Export ..."
							End If
					End If
					'Set Value for Export Directory.
					If sExportDirectory<>"" Then
						If instr(1,sExportDirectory,"PIE")>1 Then
							Set objFSO = CreateObject("Scripting.FileSystemObject")
								If not objFSO.FolderExists(Environment.Value("BatchFldName")+"\PIE") then
									objFSO.CreateFolder(Environment.Value("BatchFldName")&"\PIE")
									wait 5
								End if
								Set objFSO=nothing
						End If
						'Call Fn_Edit_Box("Fn_PLM_ExportObjects",ObjExport,"ExportDirectory",sExportDirectory)
						ObjExport.JavaEdit("ExportDirectory").Set ""
						ObjExport.JavaEdit("ExportDirectory").Set sExportDirectory
					End If
					'Set Value for Export FileName.
					If sExportFilename<>"" Then
						Call Fn_Edit_Box("Fn_PLM_ExportObjects",ObjExport,"ExportFilename",sExportFilename)
					End If
					'Set the Transfer Mode
					If sTransferMode<>"" Then
						Call Fn_List_Select("Fn_PLM_ExportObjects", ObjExport, "TransferMode",sTransferMode)
					End If
					'Set the Revision Rule
					If sRevisionRule<>"" Then
						Call Fn_List_Select("Fn_PLM_ExportObjects", ObjExport, "RevisionRule",sRevisionRule)
					End If
					'Code for Languages
					If sLanguages<>"" Then
					End If
					'Set perform export in background checkbox.
					If bExportInBack<>"" Then
						Call Fn_CheckBox_Set("Fn_PLM_ExportObjects", ObjExport, "PerformExportInBackground", bExportInBack)
					End If
					'Set Open PLMXML File checkbox.
					If bOpenFile<>"" Then
						Call Fn_CheckBox_Set("Fn_PLM_ExportObjects", ObjExport, "OpenPLMXMLFile", bOpenFile)
					End If

					If 	sAction = "VerifyAndSetToPLMXMLExport" Then
					bFlag = False
						iRows = CInt(ObjExport.JavaTable("ObjectTable").GetROProperty("rows"))
						For iCount = 0 to iRows-1
							sObjListValue = ObjExport.JavaTable("ObjectTable").GetCellData(iCount, "Object")		
							If sObjListValue = sObjectList Then
								bFlag = True
								Exit For
							End If
						Next
						If  bFlag = True Then
							Fn_PLM_ExportObjects = True
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully exceuted "& sAction &" case of Fn_PLM_ExportObjects")
						End If
					Else
						Fn_PLM_ExportObjects = True
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully exceuted "& sAction &" case of Fn_PLM_ExportObjects")
					End If

		Case Else 
                Fn_PLM_ExportObjects = False			 				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Fn_PLM_ExportObjects failed due to Invalid arguments")
		End Select

		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)		
				For iCount=0 to Ubound(aButtons)
					Call Fn_Button_Click("Fn_PLM_ExportObjects", ObjExport, aButtons(iCount))
				Next
		End If
		Call Fn_ReadyStatusSync(3)'Added By nilesh

'Check for the Export Complete Message

Set WshShl = CreateObject("WScript.Shell")
	Set Shell = WshShl.Environment("User")
	sDirPath = Shell("AutomationDir")	
	sXmlPath=sDirPath+"\TestData\AutomationXML\WebConfig\PIE_Messages.xml"
	bReturn=Fn_GetXMLNodeValue(sXmlPath, "ExportCompleteMessage")
'Added By Ashok kakade
'Modified Code by Shreyas
	IF Window("PLMXML-TeamCenterWindow").JavaApplet("JApplet").JavaDialog("Export Completed").Exist(10) = True Then
		Set objExportcomp = Window("PLMXML-TeamCenterWindow").JavaApplet("JApplet").JavaDialog("Export Completed")
		sValue = objExportcomp.JavaObject("Message").Object.text
	ElseIF JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed").Exist(1) = True Then 
		Set objExportcomp = JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed")
		sValue = objExportcomp.JavaObject("Message").Object.text
	ElseIF JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaDialog("Export Completed").Exist(1) = True Then
		Set objExportcomp = JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaDialog("Export Completed")
		sValue = objExportcomp.JavaObject("Message").Object.text
	Else
		Fn_PLM_ExportObjects = False
		Exit Function 
	End IF

If lCase(trim(cstr(sValue)))=Lcase(Trim(cstr(bReturn))) Then
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Verified That the Export Is complete without Any Errors")
	Fn_PLM_ExportObjects = True
Else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed To Verify That the Export Is complete without Any Errors")
	Fn_PLM_ExportObjects = False
	Call Fn_Button_Click("Fn_PLM_ExportObjects",objExportcomp, "No")
	Exit Function
End If

		'View Lof for details
		If bViewLog<>"" Then
			'If JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Export Completed").Exist Then
				Do
				Loop Until  objExportcomp.Exist = True
				Call Fn_Button_Click("Fn_PLM_ExportObjects", objExportcomp, bViewLog)
			'End If
		End If


	Set objExport  = Fn_SISW_PLMXML_GetObject("Export")
		IF objExport.Exist(1) = True Then
		        Call Fn_Button_Click("Fn_PLM_ExportObjects", objExport, aGlobalDictionary("Button"))
	   End If
		'Setting Object to Nothing
		Set ObjExport=Nothing
End Function

'*********************************************************                        Function to  perform operations on String Viewer Dialog            ***********************************************************************
'Function Name    :     Fn_PLM_StringViewer_Operations

'Description        :     Performs operations like verify and retrieve  data from String Viewer Dialogbox

'Parameters       :      1. sAction : Actions to perform (Verify or get)
'                               2. sFormatType :  Format of Type
'                              3. sValue  :  Value to be verified.

'Return Value    :    TRUE \ FALSE
'                            String data in case of 'Get' /  False
 
'Pre-requisite    :   String Viewer dialog box should be open

'Examples       :   Fn_PLM_StringViewer_Operations("Verify","Html","Error: None", "Close")
'                         Fn_PLM_StringViewer_Operations("Verify","Text","Error: None", "Close")    
'                        Fn_PLM_StringViewer_Operations("Get","Text","","Close")
'History         :                     
'  						Developer Name         Date                        Rev. No.     Reviewer
'            ---------------------------------------------------------------------------------------------------------------------------
' 							Pranav S.            19/01/2010                     1.0            Ketan
'            ---------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PLM_StringViewer_Operations(sAction, sFormatType, sValue, sButtons)
		GBL_FAILED_FUNCTION_NAME="Fn_PLM_StringViewer_Operations"
		Dim StringViewerboxObj, childObjects, var, sStringValue, aButtons, iCount
		Fn_PLM_StringViewer_Operations = False
		Set StringViewerboxObj = JavaDialog("String Viewer")
		Select Case sAction
				Case "Verify", "Get"
						' setting Type checkbox
						If sFormatType <> "" Then
							If sFormatType = "Text" Then
								'StringViewerboxObj.JavaCheckBox("Text").Set("ON")
								Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_PLM_StringViewer_Operations", "Set", StringViewerboxObj.JavaCheckBox("Text"), "", "ON")
							ElseIf sFormatType = "Html" Then
								'StringViewerboxObj.JavaCheckBox("HTML").Set("ON")
								Call Fn_SISW_UI_JavaCheckBox_Operations("Fn_PLM_StringViewer_Operations", "Set", StringViewerboxObj.JavaCheckBox("HTML"), "", "ON")
							End If
						End If						
						' Getting text from the preview editbox
						Set var=description.Create()
						var("Class Name").value = "JavaEdit"
						Set  childObjects= StringViewerboxObj.ChildObjects(var)
						sStringValue = "" + childObjects(0).getROProperty("value")
				Select Case sAction
						Case "Verify"
						If sValue <> "" Then
							'Kavan:Instr Call Modified For Textual Comparision instead of binary comparision.
							If Instr(1,lcase(sStringValue), lcase(sValue),1) <> 0 Then
								Fn_PLM_StringViewer_Operations = True
							End If
						End If                                                    						
						' get text from Editbox
						Case "Get"
						Fn_PLM_StringViewer_Operations = sStringValue						
				End Select
		End Select
		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)				
				For iCount=0 to Ubound(aButtons)
					Call Fn_Button_Click("Fn_PLM_StringViewer_Operations", StringViewerboxObj, aButtons(iCount))
				Next
		End If
		Set StringViewerboxObj =  nothing
		Set childObjects =  nothing
End Function
'#############################################################################################################
'###    FUNCTION NAME   :   Fn_PLM_ActionRule(sAction, sARName, sARDescription, sARScope, sARSchema, sARLocation, sARHandler)
'###
'###    DESCRIPTION        :   Add / Modify / Delete / Exist
'###	Prequisite 				:	PLM XML Export Import Administration Prespective is Open.
'###
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Edit_Box(), Fn_UI_JavaRadioButton_SetON(), Fn_Button_Click()
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                  	19/01/2010         		1.0
'###
'###
'###    EXAMPLE          : 		Case "Create" : Msgbox Fn_PLM_ActionRule("Create", "ketan", "Testing", "Import", "PLMXML", "During Action", "")
'###									Case "Modify" : Msgbox Fn_PLM_ActionRule("Modify", "ketan", "Test", "Export", "PLMXML", "Pre Action", "")
'###									Case "Delete" : Msgbox Fn_PLM_ActionRule("Delete", "", "", "", "", "", "")
'###									Case "Verify" : Msgbox Fn_PLM_ActionRule("Verify", "ketan", "Testing", "Import", "PLMXML", "During Action", "")
'#############################################################################################################
Public Function Fn_PLM_ActionRule(sAction, sARName, sARDescription, sARScope, sARSchema, sARLocation, sARHandler)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_ActionRule"
	Dim objActionRule, iCount, iCounter
	Fn_PLM_ActionRule = False
	Set objActionRule = Fn_UI_ObjectCreate("Fn_PLM_ActionRule", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Action Rule Name
						If sARName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_ActionRule",objActionRule,"ARName",sARName)
						End If
						'Set Value for Description
						If sARDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_ActionRule",objActionRule,"Description",sARDescription)
						End If
						'Set the Type of Transfer
						If sARScope<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_PLM_ActionRule",objActionRule, sARScope)
						End If
						'Set Output Schema format
						If sARSchema<>"" Then
							Call Fn_List_Select("Fn_PLM_ActionRule", objActionRule, "SchemaFormat",sARSchema)
						End If
						'Set the Support Incremental
						If sARLocation<>"" Then
							Call Fn_List_Select("Fn_PLM_ActionRule", objActionRule, "ARLocationList",sARLocation)
						End If
						'Set the Closure Rule
						If sARHandler<>"" Then
							Call Fn_List_Select("Fn_PLM_ActionRule", objActionRule, "ARHandlerList",sARHandler)
						End If
						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_ActionRule", objActionRule, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_ActionRule", objActionRule, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_ActionRule", objActionRule, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_ActionRule", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "Verify"
						iCount = 0
						iCounter = 0
						'Verify AR Name
						If sARName<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_ActionRule",objActionRule,"ARName"))) = Trim(Lcase(sARName)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify TM Description
						If sARDescription<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_ActionRule",objActionRule,"Description"))) = Trim(Lcase(sARDescription)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Scope of Action
						If sARScope<>"" Then
							iCount = iCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_ActionRule",objActionRule.JavaRadioButton(sARScope), "value")) = 1 Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Output Scheme Format
						If sARSchema<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objActionRule.JavaList("SchemaFormat").GetItem(objActionRule.JavaList("SchemaFormat").Object.getSelectedIndex))) = Trim(Lcase(sARSchema)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify AR Location
						If sARLocation<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objActionRule.JavaList("ARLocationList").GetItem(objActionRule.JavaList("ARLocationList").Object.getSelectedIndex))) = Trim(Lcase(sARLocation)) Then
								iCounter = iCounter + 1
							End If							
						End If
						'Verify Closure Rule
						If sARHandler<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objActionRule.JavaList("ARHandlerList").GetItem(objActionRule.JavaList("ARHandlerList").Object.getSelectedIndex))) = Trim(Lcase(sARHandler)) Then
								iCounter = iCounter + 1
							End If
						End If
							If iCount=iCounter Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : All Given values match actual values")
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : All Given values does not match actual values")
								Fn_PLM_ActionRule = FALSE
								Set objActionRule = nothing 
								Exit Function
						End If						
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : Fn_PLM_ActionRule function failed")
						Fn_PLM_ActionRule = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_ActionRule")
	Fn_PLM_ActionRule = TRUE
	Set objActionRule = nothing 
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_PLM_FilterRuleOperations(sAction, sFRName, sDescription, sScope, sSchema)
'###
'###    DESCRIPTION        :   Add / Modify / Delete / AddRow / ModifyRow operations on Filter Rule
'###	Prequisite 				:	PLM XML Export Import Administration Prespective is Open.
'###
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Edit_Box(), Fn_UI_JavaRadioButton_SetON(), Fn_Button_Click
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    	19/01/2010         1.0
'###
'###    EXAMPLE          : 		Case "Create" : Msgbox Fn_PLM_FilterRuleOperations("Create", "KetanTest", "Testing", "Export", "PLMXML")
'###									Case "Modify" : Msgbox Fn_PLM_FilterRuleOperations("Modify", "", "Test", "Import", "PLMXML")
'###									Case "Delete" : Msgbox Fn_PLM_FilterRuleOperations("Delete", "", "", "", "")
'###									Case "AddRow" : Msgbox Fn_PLM_FilterRuleOperations("AddRow", "Object Class Type:Object Name:Filter Rule Name", "TYPE:xyz:MySampleFilter", "", "")
'###									Case "ModifyRow" : Msgbox Fn_PLM_FilterRuleOperations("ModifyRow", "Object Class Type:Object Name:Filter Rule Name", "CLASS:abc:MySampleFilter", "", "")
'#############################################################################################################
Public Function Fn_PLM_FilterRuleOperations(sAction, sFRName, sDescription, sScope, sSchema)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_FilterRuleOperations"
	Dim objClosureRule, iReturn, aColumns, aColsData, iCols, iRows, iFlag, intCount, iCounter
	Fn_PLM_FilterRuleOperations = False
	Set objClosureRule = Fn_UI_ObjectCreate("Fn_PLM_FilterRuleOperations", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Rule Name
						If sFRName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_FilterRuleOperations",objClosureRule,"FRName",sFRName)
						End If
						'Set Value for Description
						If sDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_FilterRuleOperations",objClosureRule,"Description",sDescription)
						End If
						'Set the Scope of Filter.
						If sScope<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_PLM_FilterRuleOperations",objClosureRule, sScope)
						End If
						'Set Output Schema format
						If sSchema<>"" Then
							iReturn = objClosureRule.JavaList("SchemaFormat").GetItemIndex(sSchema)
							objClosureRule.JavaList("SchemaFormat").Object.setSelectedIndex iReturn
						End If
						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_FilterRuleOperations", objClosureRule, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_FilterRuleOperations", objClosureRule, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_FilterRuleOperations", objClosureRule, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_FilterRuleOperations", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "AddRow","ModifyRow"
					iFlag = 0
					If sAction="AddRow" Then
						'Click on Add button.
						Call Fn_Button_Click("Fn_PLM_FilterRuleOperations", objClosureRule, "AddColumn")
					End If
							aColumns = Split(sFRName,":",-1,1)
							aColsData = Split(sDescription,":",-1,1)
							iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
							iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
										'Logic for Setting Data
										If aColumns(iCounter)="Object Class Type" OR aColumns(iCounter)="Filter Rule Name" Then										
											'Select value from List
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
										Else
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),""
											'Set value in Cell.
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
										End If
										Exit For
									End If									
								Next
							Next
							'Click on Modify button.
							Call Fn_Button_Click("Fn_PLM_FilterRuleOperations", objClosureRule, "Modify")
							If iFlag=Ubound(aColumns)+1 Then
								Fn_PLM_FilterRuleOperations = True
							Else
								Fn_PLM_FilterRuleOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" of Fn_PLM_FilterRuleOperations failed")
								Set objClosureRule = nothing
								Exit Function
							End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_FilterRuleOperations function failed")
						Fn_PLM_FilterRuleOperations = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_FilterRuleOperations")
	Fn_PLM_FilterRuleOperations = TRUE
	Set objClosureRule = nothing 
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_PLM_PropertySetOperations(sAction, sPSName, sDescription, sScope)
'###
'###    DESCRIPTION        :   Add / Modify / Delete / AddRow / ModifyRow operations on Property Set
'###	Prequisite 				:	PLM XML Export Import Administration Prespective is Open.
'###
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Edit_Box(), Fn_UI_JavaRadioButton_SetON(), Fn_Button_Click
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    	19/01/2010         1.0
'###
'###    EXAMPLE          : 		Case "Create" : Msgbox Fn_PLM_PropertySetOperations("Create", "KetanTest", "Testing", "Export")
'###									Case "Modify" : Msgbox Fn_PLM_PropertySetOperations("Modify", "", "Test", "Import")
'###									Case "Delete" : Msgbox Fn_PLM_PropertySetOperations("Delete", "", "", "")
'###									Case "AddRow" : Msgbox Fn_PLM_PropertySetOperations("AddRow", "Primary Object Class Type:Primary Object:Relation Type:Related Property Or Object:Property Action Type", "CLASS:ItemRevision:PROPERTY:yes:DO", "")
'###									Case "ModifyRow" : Msgbox Fn_PLM_PropertySetOperations("ModifyRow", "Related Property Or Object", "no", "")
'#############################################################################################################
Public Function Fn_PLM_PropertySetOperations(sAction, sPSName, sDescription, sScope)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_PropertySetOperations"
	Dim objClosureRule, iReturn, aColumns, aColsData, iCols, iRows, iRowCnt, iFlag, intCount, iCount, iCounter
	Fn_PLM_PropertySetOperations = False
	Set objClosureRule = Fn_UI_ObjectCreate("Fn_PLM_PropertySetOperations", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Rule Name
						If sPSName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_PropertySetOperations",objClosureRule,"PSName",sPSName)
						End If
						'Set Value for Description
						If sDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_PropertySetOperations",objClosureRule,"Description",sDescription)
						End If
						'Set the Scope of Filter.
						If sScope<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_PLM_PropertySetOperations",objClosureRule, sScope)
						End If
						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "AddRow","ModifyRow"
					iFlag = 0
					If sAction="AddRow" Then
						'Click on Add button.
						Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "AddColumn")
					End If
							aColumns = Split(sPSName,":",-1,1)
							aColsData = Split(sDescription,":",-1,1)
							'iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
							'iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							 iCols =  Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations",objClosureRule.JavaTable("ClosureRuleTable"),"cols")
							 iRows =Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations",objClosureRule.JavaTable("ClosureRuleTable"),"rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
										'Logic for Setting Data
										If aColumns(iCounter)="Primary Object Class Type" OR aColumns(iCounter)="Relation Type" OR aColumns(iCounter)="Property Action Type" Then										
											'Select value from List
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
										Else
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),""
											'Set value in Cell.
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
										End If
										Exit For
									End If									
								Next
							Next
							'Click on Modify button.
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "Modify")
							If iFlag=Ubound(aColumns)+1 Then
								Fn_PLM_PropertySetOperations = True
							Else
								Fn_PLM_PropertySetOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" of Fn_PLM_PropertySetOperations failed")
								Set objClosureRule = nothing
								Exit Function
							End If
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_PropertySetOperations function failed")
						Fn_PLM_PropertySetOperations = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_PropertySetOperations")
	Fn_PLM_PropertySetOperations = TRUE
	Set objClosureRule = nothing 
End Function
'#########################################################################################################
'###
'###    FUNCTION NAME   :   Fn_PLM_TransferOptionSetOperations(sAction, sTOSName, sDescription, bRemoteSite, sRemoteSiteID, sTransferMode)
'###
'###    DESCRIPTION        :   Add / Modify / Delete / AddRow / ModifyRow operations on Transfer Option Set
'###	Prequisite 				:	PLM XML Export Import Administration Prespective is Open.
'###
'###    Function Calls       :   Fn_WriteLogFile(), Fn_UI_ObjectCreate(), Fn_Edit_Box(), Fn_UI_JavaRadioButton_SetON(), Fn_Button_Click
'###
'###	 HISTORY             :   AUTHOR                 			DATE        	VERSION
'###
'###    CREATED BY     :    Ketan Raje                    	20/01/2010         1.0
'###
'###    EXAMPLE          : 		Case "Create" : Msgbox Fn_PLM_TransferOptionSetOperations("Create", "Test111", "Testing", "OFF", "", "qwerty123")
'###									Case "Modify" : Msgbox Fn_PLM_TransferOptionSetOperations("Modify", "", "Test", "OFF", "", "qwerty123")
'###									Case "Delete" : Msgbox Fn_PLM_TransferOptionSetOperations("Delete", "", "", "", "", "")
'###									Case "AddRow" : Msgbox Fn_PLM_TransferOptionSetOperations("AddRow", "Option:Display Name:Default Value:Description:Group Name:Read Only", "abc:def:True:hij:klm:ON", "", "", "")
'###									Case "ModifyRow" : Msgbox Fn_PLM_TransferOptionSetOperations("ModifyRow", "Default Value:Read Only", "False:OFF", "", "", "")
'###									Case "Verify" : Msgbox Fn_PLM_TransferOptionSetOperations("Verify", "Test111", "Testing", "OFF", "", "PLMXMLAdminDataExport")
'#############################################################################################################
Public Function Fn_PLM_TransferOptionSetOperations(sAction, sTOSName, sDescription, bRemoteSite, sRemoteSiteID, sTransferMode)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_TransferOptionSetOperations"
	Dim objTOS, iReturn, aColumns, aColsData, iCols, iRows, iFlag, intCount, iCounter
	Fn_PLM_TransferOptionSetOperations = False
	Set objTOS = Fn_UI_ObjectCreate("Fn_PLM_TransferOptionSetOperations", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Transfer Option Set
						If sTOSName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_TransferOptionSetOperations",objTOS,"TOSName",sTOSName)
						End If
						'Set Value for Description
						If sDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_TransferOptionSetOperations",objTOS,"Description",sDescription)
						End If
						'Set Remote Site Option true / false
						If bRemoteSite<>"" Then
							Call Fn_CheckBox_Set("Fn_PLM_TransferOptionSetOperations", objTOS, "RemoteSite", bRemoteSite)
						End If
						'Set Remote Site ID
						If sRemoteSiteID<>"" Then
							Call Fn_List_Select("Fn_PLM_TransferOptionSetOperations", objTOS, "TOSRemoteSiteID",sRemoteSiteID)
						End If
						'Set Transfer Mode
						If sTransferMode<>"" Then
							Call Fn_List_Select("Fn_PLM_TransferOptionSetOperations", objTOS, "TOSTransferMode",sTransferMode)
						End If
						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations", objTOS, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations", objTOS, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations", objTOS, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "AddRow","ModifyRow"
					iFlag = 0
					If sAction="AddRow" Then
						'Click on Add button.
						Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations", objTOS, "AddColumn")
					End If
							aColumns = Split(sTOSName,":",-1,1)
							aColsData = Split(sDescription,":",-1,1)
							iCols = objTOS.JavaTable("ClosureRuleTable").GetROProperty("cols")
							iRows = objTOS.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objTOS.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										'Logic for Setting Data
										If aColumns(iCounter)="Default Value" Then
											objTOS.JavaTable("ClosureRuleTable").SetCellData iRows,"Default Value",aColsData(iCounter)
										ElseIf aColumns(iCounter) = "Read Only" Then
											If aColsData(iCounter) = "ON" Then
												Call Fn_UI_JavaTable_ClickCell("Fn_PLM_TransferOptionSetOperations",objTOS,"ClosureRuleTable",iRows, aColumns(iCounter))											
											End If											
										Else
											'Set value TableEditbox
											objTOS.JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
											Call Fn_Edit_Box("Fn_PLM_TransferOptionSetOperations",objTOS,"TableEditbox",aColsData(iCounter))
										End If
										Exit For
									End If									
								Next
							Next
							'Click on Modify button.
							Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations", objTOS, "Modify")
							If iFlag=Ubound(aColumns)+1 Then
								Fn_PLM_TransferOptionSetOperations = True
							Else
								Fn_PLM_TransferOptionSetOperations = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" of Fn_PLM_TransferOptionSetOperations failed")
								Set objTOS = nothing
								Exit Function
							End If
			Case "Verify"
						intCount = 0
						iCounter = 0
						'Verify TOS Name
						If sTOSName<>"" Then
							intCount = intCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferOptionSetOperations",objTOS,"TOSName"))) = Trim(Lcase(sTOSName)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify TOS Description
						If sDescription<>"" Then
							intCount = intCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferOptionSetOperations",objTOS,"Description"))) = Trim(Lcase(sDescription)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Remote Site Checkbox
						If bRemoteSite<>"" Then
							intCount = intCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferOptionSetOperations",objTOS.JavaCheckBox("RemoteSite"), "value")) = 1 and Trim(Lcase(bRemoteSite)) = "on" Then
								iCounter = iCounter + 1
							ElseIf Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferOptionSetOperations",objTOS.JavaCheckBox("RemoteSite"), "value")) = 0 and Trim(Lcase(bRemoteSite)) = "off" Then
								iCounter = iCounter + 1
							End If							
						End If
						'Verify Remote Site ID
						If sRemoteSiteID<>"" Then
							intCount = intCount + 1
							If Trim(Lcase(objTOS.JavaList("TOSRemoteSiteID").GetItem(objTOS.JavaList("TOSRemoteSiteID").Object.getSelectedIndex))) = Trim(Lcase(sRemoteSiteID)) Then
								iCounter = iCounter + 1
							End If
						End If
						'Verify Transfer Mode
						If sTransferMode<>"" Then
							intCount = intCount + 1
							If Trim(Lcase(objTOS.JavaList("TOSTransferMode").GetItem(objTOS.JavaList("TOSTransferMode").Object.getSelectedIndex))) = Trim(Lcase(sTransferMode)) Then
								iCounter = iCounter + 1
							End If
						End If
							If intCount=iCounter Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : All Given values match actual values")
								Fn_PLM_TransferOptionSetOperations = TRUE
								Set objTOS = nothing 
								Exit Function
							Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : All Given values does not match actual values")
								Fn_PLM_TransferOptionSetOperations = FALSE
								Set objTOS = nothing 
								Exit Function
						End If						
			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_TransferOptionSetOperations function failed")
						Fn_PLM_TransferOptionSetOperations = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_TransferOptionSetOperations")
	Fn_PLM_TransferOptionSetOperations = TRUE
	Set objTOS = nothing 
End Function
'----------------------------------------------------------------------------------------------Function for PLMXML Import of a file.----------------------------------------------------------------------------------------------------------------------------
'Function Name		  :	  Fn_PLM_ImportObjects

'Description			 :	 Function to Import Objects

'Parameters			   :	sAction, sImportType, sImportObject, sTransferMode, bViewLog, aGlobalDictionary, sButtons

'Return Value		   : 	True Or False

'Pre-requisite			:	Object which we have to Import should be selected.

'Examples				:	Call Fn_PLM_ImportObjects("SetPLMXMLImport", "PLMXML", "D:\Mainline\Ketan.xml", "ConfiguredDataImportDefault", "Yes", "", "OK")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Updated By Ketan on 25-Feb-2011 to handle IC Context while Import.'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'								dicIncrementalChange("sICSelectionType") = "SelectICContext" 
'								dicIncrementalChange("sICName") = "ICNew1"
'								dicIncrementalChange("sICId") = "CN0016"
'								Environment.Value("TestLogFile") = "D:\Log.txt"
'								Call Fn_PLM_ImportObjects("SetPLMXMLImport", "PLMXML", "D:\Mainline\qwerty.xml", "ConfiguredDataImportDefault", "Yes", "", "OK")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'History					 :					Developer Name												Date						Rev. No.	
'---------------------------------------------------------------------------------------------------------------------------------------------
'													Ketan Raje										   	21/01/2011			           1.0		
'													Sushma Pagare					   		     	    15/09/2011			           1.0      Added Case "SetFromPLMXML"	for Menu Tools-> Import -> From PLMXML	 
'													Koustubh W					   		     	        20/09/2011 						Modified code to handle Import Completed dialog
'													Pranav Ingle					   		     	    13/06/2012 						Modified Import Completed Path
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PLM_ImportObjects(sAction, sImportType, sImportObject, sTransferMode, bViewLog, aGlobalDictionary, sButtons)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_ImportObjects"
   'Declaring Variables
   Dim ObjImport,  aButtons, iCount, bReturn, sItem, objTable, iRowCount
   Dim WshShell,WshSysEnv,sDrivePath, fso, sLogFilePath, sToolsImport
   'Setting bReturn to False
   bReturn=False
			Set fso = CreateObject("Scripting.FileSystemObject")
			'Check the exsitance of file.log
			sLogFilePath = sImportObject+".log"
			If  fso.FileExists(sLogFilePath) Then
					fso.DeleteFile sLogFilePath
			End IF
            Set fso = Nothing
   'Initially Function returns False
   Fn_PLM_ImportObjects=False


   Select Case sAction

	Case "SetFromPLMXML" ,  "SetPLMXMLImport"    	
		''  SetFromPLMXML : Added Case for Menu Tools-> Import -> From PLMXML	 
        '' SetPLMXMLImport : Case for Menu Tools-> Import -> Objects
			If  sAction = "SetFromPLMXML"  Then
	
					 'Checking Existance of PLMXML Import Window
					 	sToolsImport=Fn_GetXMLNodeValue(Fn_LogUtil_GetXMLPath("RAC_Menu"), "ToolsImportFromPLMXML")
					   If Fn_UI_ObjectExist("Fn_PLM_ImportObjects",JavaWindow("PLMXML-TeamCenter").JavaWindow("PLMXMLImport"))=False Then
						   'Opening PLMXML Import  Window
							Call Fn_MenuOperation("Select",sToolsImport)
							Call Fn_ReadyStatusSync(1)
					   End If
					
					   'Creating object of PLMXML Import  Window	
					   Set ObjImport = Fn_UI_ObjectCreate("Fn_PLM_ImportObjects",JavaWindow("PLMXML-TeamCenter").JavaWindow("PLMXMLImport"))
					   'Set Value for Import File.
					  ' Call Fn_Edit_Box("Fn_PLM_ImportObjects",ObjImport ,"ImportingXMLFile",sImportObject)
					  	ObjImport.JavaEdit("ImportingXMLFile").Set sImportObject
					    Wait 1
					 	'Set the Transfer Mode Name
						If sTransferMode<>"" Then
							Call Fn_List_Select("Fn_PLM_ImportObjects", ObjImport, "TransferMode",sTransferMode)
						End If
                        
			ElseIf sAction =  "SetPLMXMLImport"  Then

					   'Checking Existance of ImportToWord Dialog
					   If Fn_UI_ObjectExist("Fn_PLM_ImportObjects",JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Import"))=False Then
						   'Opening ImportObjects Window
							Call Fn_MenuOperation("Select","Tools:Import:Objects...")
					   End If
					   'Creating object of ImportToWord Dialog	
					   Set ObjImport= Fn_UI_ObjectCreate("Fn_PLM_ImportObjects",JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Import"))

						'Select the Import Type
						ObjImport.JavaCheckBox("ImportType").SetTOProperty "attached text",sImportType
						If sImportType<>"" Then
							Call Fn_CheckBox_Select("Fn_PLM_ImportObjects", ObjImport, "ImportType")
						End If
						'SET property for Import dialog
						ObjImport.SetTOProperty "title","PLM XML Import ..."
						'Click on Browse button.
						Call Fn_Button_Click("Fn_PLM_ImportObjects", ObjImport, "Browse")
					'Set Value for Import Directory.
					If sImportObject<>"" Then
						ObjImport.JavaDialog("SelectObject").JavaEdit("FileName").SetTOProperty "attached text","File name:"
						If ObjImport.JavaDialog("SelectObject").JavaEdit("FileName").Exist Then
							Call Fn_Edit_Box("Fn_PLM_ImportObjects",ObjImport.JavaDialog("SelectObject"),"FileName",sImportObject)
						Else
							ObjImport.JavaDialog("SelectObject").JavaEdit("FileName").SetTOProperty "attached text","Directory name:"
							If ObjImport.JavaDialog("SelectObject").JavaEdit("FileName").Exist Then
								Call Fn_Edit_Box("Fn_PLM_ImportObjects",ObjImport.JavaDialog("SelectObject"),"FileName",sImportObject)
							End If
						End If						
					End If
					'Click on Select button.
					Call Fn_Button_Click("Fn_PLM_ImportObjects", ObjImport.JavaDialog("SelectObject"), "Select")	
					'Set the Transfer Mode Name
					If sTransferMode<>"" Then
						Call Fn_List_Select("Fn_PLM_ImportObjects", ObjImport, "TransferMode",sTransferMode)
					End If
				End If    '' If sAction = "SetFromPLMXML"  or  "SetPLMXMLImport" 

				'Selecting IC Context
				'Added By Ketan on 28-Feb-2011 to Load DictionaryDeclaration.vbs	Due to UI changes in Build: 20110119
				  Set WshShell = CreateObject("WScript.Shell")
				  Set WshSysEnv = WshShell.Environment("User")
				  If WshSysEnv("AutomationDir") <> "" Then
				   sDrivePath = WshSysEnv("AutomationDir")
				  Else
				   sDrivePath = "D:\mainline"
				  End If
				  Set WshShell = Nothing
				  Set WshSysEnv = Nothing 				
				  ExecuteFile sDrivePath + "\Library\DictionaryDeclaration.vbs"
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				If dicIncrementalChange("sICSelectionType") <> "" Then						
					Select Case dicIncrementalChange("sICSelectionType")							
						Case "SelectICContext"
								sItem = cstr(dicIncrementalChange("sICId")) & "-" & dicIncrementalChange("sICName")
								'Click on Select IC Context button
								Call Fn_CheckBox_Set("Fn_PLM_ImportObjects", ObjImport, "SelectICContext", "ON")
								' typeing value in Name edit box
									If dicIncrementalChange("sICName")<> ""  Then
											ObjImport.JavaWindow("SelectICContext").JavaEdit("IC_Name").Set ""
											Call Fn_UI_EditBox_Type("Fn_PLM_ImportObjects",ObjImport.JavaWindow("SelectICContext"),"IC_Name",dicIncrementalChange("sICName"))
									End If
									' typeing value in Id edit box
									If dicIncrementalChange("sICId")<> ""  Then
											ObjImport.JavaWindow("SelectICContext").JavaEdit("IC_ID").Set ""
											Call Fn_UI_EditBox_Type("Fn_PLM_ImportObjects",ObjImport.JavaWindow("SelectICContext"),"IC_ID",dicIncrementalChange("sICId"))
									End If
									wait(5)
									' clicking on Find button
									Call Fn_Button_Click("Fn_PLM_ImportObjects", ObjImport.JavaWindow("SelectICContext"), "IC_Find")
									wait(5)
									Set objTable = ObjImport.JavaWindow("SelectICContext").JavaTable("IC_ICTable")
			
									iRowCount = cint(objTable.GetROProperty("rows"))
									For iCount = 0 to iRowCount -1
												If objTable.GetCellData(iCount,"Object") = sItem  Then
													objTable.SelectRow iCount
													Exit for
												End If
									Next
									iCount = cint(objTable.GetROProperty("SelectedRow"))
									If iCount <> -1 Then
										objTable.DoubleClickCell iCount,"Object","LEFT"
											Call Fn_ReadyStatusSync(5)
									Else
										Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL: Function [ Fn_PLM_ImportObjects ] Case [ " & sAction  & " ] No Item is selected.")
										 Fn_PLM_ImportObjects = False
										Exit function
									End If
									Fn_PLM_ImportObjects = True								
					End Select
				End If
					Fn_PLM_ImportObjects = True
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass: Successfully exceuted "& sAction &" case of Fn_PLM_ImportObjects")
		Case Else 
                Fn_PLM_ImportObjects = False			 				
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail: Fn_PLM_ImportObjects failed due to Invalid arguments")
		End Select
		'Click on Buttons
		If sButtons<>"" Then
				aButtons = split(sButtons, ":",-1,1)				
				For iCount=0 to Ubound(aButtons)
					Call Fn_Button_Click("Fn_PLM_ImportObjects", ObjImport, aButtons(iCount))
				Next
		End If
		'View Lof for details
		If bViewLog<>"" Then
            If JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Import Completed").Exist(iTimeOut) then
				Call Fn_Button_Click("Fn_PLM_ImportObjects", JavaWindow("PLMXML-TeamCenter").JavaWindow("TcDefaultApplet").JavaDialog("Import Completed"), bViewLog)			
			End If
		End If
		'Setting Object to Nothing
		Set ObjImport=Nothing
End Function

''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
''''''/$$$$
''''''/$$$$   FUNCTION NAME   : Fn_PLM_PropertySetOperations_Ext(sAction, sPSName, sDescription, sScope,aProperties,aValues,sInfo1,sInfo2)
''''''/$$$$
''''''/$$$$   DESCRIPTION        :  This function is an extention of the Existing Function to en
''''''/$$$$
''''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
''''''/$$$$										sPSName : Valid Property set name
''''''/$$$$										sDescription : Valid Property set Description
''''''/$$$$										sScope : Valid Scope
''''''/$$$$										aProperties : Valid array of properties
''''''/$$$$										aValues : Valid array of property values
''''''/$$$$										sInfo1 : For Future Use
''''''/$$$$										sInfo2 : For Future Use
''''''/$$$$	
''''''/$$$$		Return Value : 				True or False
''''''/$$$$
''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
''''''/$$$$
''''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
''''''/$$$$  
''''''/$$$$    CREATED BY     :   SHREYAS          23/12/2011         1.0
''''''/$$$$
''''''/$$$$    REVIWED BY     :  Shreyas			23/12/2011            1.0
''''''/$$$$
''''''/$$$$		How To Use :    bReturn=Fn_PLM_PropertySetOperations_Ext("Create", "Test", "Test", "Export",i,y,"","")
''''''/$$$$										bReturn=Fn_PLM_PropertySetOperations_Ext("VerifyDetails", "Test", "Test", "Export",i,y,"","")
''''''/$$$$							
''''''/$$$$	
''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'
'

Public Function Fn_PLM_PropertySetOperations_Ext(sAction, sPSName, sDescription, sScope,aProperties,aValues,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_PropertySetOperations_Ext"
	Dim objClosureRule, iReturn, aColumns, aColsData, iCols, iRows, iRowCnt, iFlag, intCount, iCount, iCounter
	Dim bFlag
	Fn_PLM_PropertySetOperations_Ext = False
	Set objClosureRule = Fn_UI_ObjectCreate("Fn_PLM_PropertySetOperations", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Rule Name
						If sPSName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_PropertySetOperations",objClosureRule,"PSName",sPSName)
						End If
						'Set Value for Description
						If sDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_PropertySetOperations",objClosureRule,"Description",sDescription)
						End If
						'Set the Scope of Filter.
						If sScope<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_PLM_PropertySetOperations_Ext",objClosureRule, sScope)
						End If

						Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "AddColumn")
								aColumns = Split(aProperties,":",-1,1)
								aColsData = Split(aValues,":",-1,1)
								'iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
								'iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
								iCols =Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations",objClosureRule.JavaTable("ClosureRuleTable"), "cols")
								iRows =Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations",objClosureRule.JavaTable("ClosureRuleTable"),"rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
										'Logic for Setting Data
										If aColumns(iCounter)="Primary Object Class Type" OR aColumns(iCounter)="Relation Type" OR aColumns(iCounter)="Property Action Type" Then										
											'Select value from List
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
										Else
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),""
											'Set value in Cell.
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
										End If
										Exit For
									End If									
								Next
							Next

						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations_Ext", objClosureRule, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations_Ext", objClosureRule, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_PropertySetOperations_Ext", objClosureRule, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations_Ext", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "AddRow","ModifyRow"
					iFlag = 0
					If sAction="AddRow" Then
						'Click on Add button.
						Call Fn_Button_Click("Fn_PLM_PropertySetOperations_Ext", objClosureRule, "AddColumn")
					End If

							aColumns = Split(aProperties,":",-1,1)
							aColsData = Split(aValues,":",-1,1)
							'iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
							'iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							iCols =Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations",objClosureRule.JavaTable("ClosureRuleTable"), "cols")
							iRows =Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations",objClosureRule.JavaTable("ClosureRuleTable"),"rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
										'Logic for Setting Data
										If aColumns(iCounter)="Primary Object Class Type" OR aColumns(iCounter)="Relation Type" OR aColumns(iCounter)="Property Action Type" Then										
											'Select value from List
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
										Else
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),""
											'Set value in Cell.
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
										End If
										Exit For
									End If									
								Next
							Next
							'Click on Modify button.
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "Modify")
							If iFlag=Ubound(aColumns)+1 Then
								Fn_PLM_PropertySetOperations_Ext = True
							Else
								Fn_PLM_PropertySetOperations_Ext = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" of Fn_PLM_PropertySetOperations failed")
								Set objClosureRule = nothing
								Exit Function
							End If

			Case "VerifyDetails","VerifyDetailsExt"

'						Verify PS Name
						If sPSName<>"" Then

							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_PropertySetOperations_Ext",objClosureRule,"PSName"))) = Trim(Lcase(sPSName)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sPSName+"]")
							Fn_PLM_PropertySetOperations_Ext = TRUE
		    					Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sPSName+"]")
								Fn_PLM_PropertySetOperations_Ext = FALSE
								Exit Function			
						End If
					End If

	'					Verify PS Description
						If sDescription<>"" Then

							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_PropertySetOperations_Ext",objClosureRule,"Description"))) = Trim(Lcase(sDescription)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sDescription+"]")
								Fn_PLM_PropertySetOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sDescription+"]")
								Fn_PLM_PropertySetOperations_Ext = FALSE
								Exit Function			
							End If
						End If

		'				Verify Scope
						If sScope<>"" Then
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations_Ext",objClosureRule.JavaRadioButton(sScope), "value")) = 1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sScope+"]")
								Fn_PLM_PropertySetOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sScope+"]")
								Fn_PLM_PropertySetOperations_Ext = FALSE
								Exit Function			
							End If
						End If
						
						'----[TC1123(20161205c00)_PoonamC_NewDevelopment_PWC_16Feb2017:Added New Case-to verify Row Within whole Table] -------------------
						If sAction = "VerifyDetailsExt" Then

								If aProperties<>"" and aValues<>"" Then
												aColumns = Split(aProperties,":",-1,1)
												aColsData = Split(aValues,":",-1,1)
												iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
												iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
											For iCounter=0 to Ubound(aColumns)
												For intCount=0 to iCols-1
													If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
														  bFlag = False
														  For iCount=0 to iRows
														  		 sValue=objClosureRule.JavaTable("ClosureRuleTable").GetCellData(iCount,aColumns(iCounter))
														  		 If Trim(lcase(sValue))=Trim(lcase(aColsData(iCounter))) Then
																		bFlag = True
																		Exit For	
																  End If	
														  Next
														If bFlag = True Then
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
																Fn_PLM_PropertySetOperations_Ext = TRUE
														Else
																	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
																	Fn_PLM_PropertySetOperations_Ext = FALSE
																	Exit Function			
														End If
														Exit For
													End If									
												Next
											Next
							    End If

						Else		
						
							      If aProperties<>"" and aValues<>"" Then
												aColumns = Split(aProperties,":",-1,1)
												aColsData = Split(aValues,":",-1,1)
												iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
												iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
											For iCounter=0 to Ubound(aColumns)
												For intCount=0 to iCols-1
													If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
			
												sValue=objClosureRule.JavaTable("ClosureRuleTable").GetCellData(iRows,aColumns(iCounter))
													If Trim(lcase(sValue))=Trim(lcase(aColsData(iCounter))) Then
															Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
															Fn_PLM_PropertySetOperations_Ext = TRUE
													Else
																Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
																Fn_PLM_PropertySetOperations_Ext = FALSE
																Exit Function			
													End If
														Exit For
													End If									
												Next
											Next
							    End If
					 End If	    

			Case "VerifyBlank"

						iCount = 0
						iCounter = 0
						'Verify PS Name
						If sPSName<>"" Then
							iCount = iCount + 1
                                   If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_PropertySetOperations_Ext",objClosureRule,"PSName"))) = Trim(Lcase(sPSName)) Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify PS Description
						If sDescription<>"" Then
							iCount = iCount + 1
                                   If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_PropertySetOperations_Ext",objClosureRule,"Description"))) = Trim(Lcase(sDescription)) Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify Scope
						If sScope<>"" Then
							iCount = iCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_PropertySetOperations_Ext",objClosureRule.JavaRadioButton(sScope), "value")) = 1 Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify Property Set Table
						If aProperties<>"" and aValues<>"" Then
										iCount = iCount + 1
										aColumns = Split(aProperties,":",-1,1)
										iRows=JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").GetROProperty("rows")
										iCols=JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").GetROProperty("cols")
										If iRows=0 Then
											iCounter = iCounter + 1
										End If
						End If
						
						If iCount=iCounter Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : All Given values match actual values")
							Fn_PLM_PropertySetOperations_Ext = TRUE
							Set objClosureRule = nothing 
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : All Given values does not match actual values")
							Fn_PLM_PropertySetOperations_Ext = FALSE
							Set objClosureRule = nothing 
							Exit Function
						End If

			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_PropertySetOperations_Ext function failed")
						Fn_PLM_PropertySetOperations_Ext = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_PropertySetOperations_Ext")
	Fn_PLM_PropertySetOperations_Ext = TRUE
	Set objClosureRule = nothing 
End Function




'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''/$$$$
'''''/$$$$   FUNCTION NAME   : Fn_PLM_FilterRuleOperations_Ext(sAction, sFRName, sDescription, sScope, sSchema,aProperties,aValues,sInfo1,sInfo2)
'''''/$$$$
'''''/$$$$   DESCRIPTION        :  This function is an extention of the Existing Function Fn_PLM_FilterRuleOperations
'''''/$$$$
'''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
'''''/$$$$										sFRName : Valid Filter to be set name
'''''/$$$$										sDescription : Valid Filter Description
'''''/$$$$										sScope : Valid Scope for filter
'''''/$$$$										sSchema : Valid Output Schema for filter
'''''/$$$$										aProperties : Valid array of properties
'''''/$$$$										aValues : Valid array of property values
'''''/$$$$										sInfo1 : For Future Use
'''''/$$$$										sInfo2 : For Future Use
'''''/$$$$	
'''''/$$$$		Return Value : 				True or False
'''''/$$$$
'''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''/$$$$
'''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
'''''/$$$$
'''''/$$$$    CREATED BY     :   SHREYAS          26/12/2011         1.0
'''''/$$$$
'''''/$$$$    REVIWED BY     :  Shreyas			26/12/2011            1.0
'''''/$$$$
'''''/$$$$		How To Use :   bReturn=Fn_PLM_FilterRuleOperations_Ext("Create", "Shreyas", "Test Filter", "Export", "PLMXML","Object Class Type:Object Name:Filter Rule Name","TYPE:xyz:MySampleFilter","","")
'''''/$$$$									bReturn=Fn_PLM_FilterRuleOperations_Ext("VerifyDetails", "Shreyas", "Test Filter", "Export", "PLMXML","Object Class Type:Object Name:Filter Rule Name","TYPE:xyz:MySampleFilter","","")
'''''/$$$$							
'''''/$$$$	
'''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_PLM_FilterRuleOperations_Ext(sAction, sFRName, sDescription, sScope, sSchema,aProperties,aValues,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_FilterRuleOperations_Ext"
	Dim objClosureRule, iReturn, aColumns, aColsData, iCols, iRows, iFlag, intCount, iCounter
	Fn_PLM_FilterRuleOperations_Ext = False
	Set objClosureRule = Fn_UI_ObjectCreate("Fn_PLM_FilterRuleOperations_Ext", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Rule Name
						If sFRName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_FilterRuleOperations_Ext",objClosureRule,"FRName",sFRName)
						End If
						'Set Value for Description
						If sDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_FilterRuleOperations_Ext",objClosureRule,"Description",sDescription)
						End If
						'Set the Scope of Filter.
						If sScope<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_PLM_FilterRuleOperations_Ext",objClosureRule, sScope)
						End If
						'Set Output Schema format
						If sSchema<>"" Then
							iReturn = objClosureRule.JavaList("SchemaFormat").GetItemIndex(sSchema)
							objClosureRule.JavaList("SchemaFormat").Object.setSelectedIndex iReturn
						End If

						'add a table row if required

				If aProperties<>"" and aValues<>"" Then
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "AddColumn")
									aColumns = Split(aProperties,":",-1,1)
									aColsData = Split(aValues,":",-1,1)
									iCols = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaTable("ClosureRuleTable"), "cols")
									iRows = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaTable("ClosureRuleTable"), "rows") - 1
'									iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
'									iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
								For iCounter=0 to Ubound(aColumns)
									For intCount=0 to iCols-1
										If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
											iFlag = iFlag + 1
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
											'Logic for Setting Data
											If aColumns(iCounter)="Primary Object Class Type" OR aColumns(iCounter)="Relation Type" OR aColumns(iCounter)="Property Action Type" Then										
												'Select value from List
												JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
											Else
												'Set value in Cell.
												JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
											End If
											Exit For
										End If									
									Next
								Next
				End If

						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_FilterRuleOperations_Ext", objClosureRule, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_FilterRuleOperations_Ext", objClosureRule, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_FilterRuleOperations_Ext", objClosureRule, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_FilterRuleOperations_Ext", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "AddRow","ModifyRow"
					iFlag = 0
					If sAction="AddRow" Then
						'Click on Add button.
						Call Fn_Button_Click("Fn_PLM_FilterRuleOperations_Ext", objClosureRule, "AddColumn")
					End If
							aColumns = Split(sFRName,":",-1,1)
							aColsData = Split(sDescription,":",-1,1)
							iCols = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaTable("ClosureRuleTable"), "cols")
							iRows = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaTable("ClosureRuleTable"), "rows") - 1
							'iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
							'iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
										'Logic for Setting Data
										If aColumns(iCounter)="Object Class Type" OR aColumns(iCounter)="Filter Rule Name" Then										
											'Select value from List
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
										Else
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),""
											'Set value in Cell.
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
										End If
										Exit For
									End If									
								Next
							Next
							'Click on Modify button.
							Call Fn_Button_Click("Fn_PLM_FilterRuleOperations_Ext", objClosureRule, "Modify")
							If iFlag=Ubound(aColumns)+1 Then
								Fn_PLM_FilterRuleOperations_Ext = True
							Else
								Fn_PLM_FilterRuleOperations_Ext = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" of Fn_PLM_FilterRuleOperations_Ext failed")
								Set objClosureRule = nothing
								Exit Function
							End If

			Case "VerifyDetails"

						iCount = 0
						iCounter = 0
						'Verify FR Name
						If sFRName<>"" Then

							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_FilterRuleOperations_Ext",objClosureRule,"FRName"))) = Trim(Lcase(sFRName)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sFRName+"]")
							Fn_PLM_FilterRuleOperations_Ext = TRUE
		    					Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sFRName+"]")
								Fn_PLM_FilterRuleOperations_Ext = FALSE
								Exit Function			
						End If
					End If

						'Verify FR Description
						If sDescription<>"" Then

							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_FilterRuleOperations_Ext",objClosureRule,"Description"))) = Trim(Lcase(sDescription)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sDescription+"]")
								Fn_PLM_FilterRuleOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sDescription+"]")
								Fn_PLM_FilterRuleOperations_Ext = FALSE
								Exit Function			
							End If
						End If

						'Verify Scope
						If sScope<>"" Then
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_FilterRuleOperations_Ext",objClosureRule.JavaRadioButton(sScope), "value")) = 1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sScope+"]")
								Fn_PLM_FilterRuleOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sScope+"]")
								Fn_PLM_FilterRuleOperations_Ext = FALSE
								Exit Function			
							End If
						End If

						'Verify Output Scheme Format
						If sSchema<>"" Then

							If Trim(Lcase(objClosureRule.JavaList("SchemaFormat").GetItem(objClosureRule.JavaList("SchemaFormat").Object.getSelectedIndex))) = Trim(Lcase(sSchema)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sSchema+"]")
								Fn_PLM_FilterRuleOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sSchema+"]")
								Fn_PLM_FilterRuleOperations_Ext = FALSE
								Exit Function			
							End If
						End If
					
					If aProperties<>"" and aValues<>"" Then
'							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "AddColumn")
									aColumns = Split(aProperties,":",-1,1)
									aColsData = Split(aValues,":",-1,1)
									iCols = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaTable("ClosureRuleTable"), "cols")
									iRows = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaTable("ClosureRuleTable"), "rows") - 1
									'iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
									'iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
								For iCounter=0 to Ubound(aColumns)
									For intCount=0 to iCols-1
										If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then

									sValue=objClosureRule.JavaTable("ClosureRuleTable").GetCellData(iRows,aColumns(iCounter))
										If Trim(lcase(sValue))=Trim(lcase(aColsData(iCounter))) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
												Fn_PLM_FilterRuleOperations_Ext = TRUE
										Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
													Fn_PLM_FilterRuleOperations_Ext = FALSE
													Exit Function			
										End If
											Exit For
										End If									
									Next
								Next
				End If

			Case "VerifyBlank"

						iCount = 0
						iCounter = 0
						'Verify FR Name
						If sFRName<>"" Then
							iCount = iCount + 1
                                   If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_FilterRuleOperations_Ext",objClosureRule,"FRName"))) = Trim(Lcase(sFRName)) Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify FR Description
						If sDescription<>"" Then
							iCount = iCount + 1
                                   If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_FilterRuleOperations_Ext",objClosureRule,"Description"))) = Trim(Lcase(sDescription)) Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify Scope
						If sScope<>"" Then
							iCount = iCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_FilterRuleOperations_Ext",objClosureRule.JavaRadioButton(sScope), "value")) = 1 Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify Output Scheme Format
						If sSchema<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objClosureRule.JavaList("SchemaFormat").GetItem(objClosureRule.JavaList("SchemaFormat").Object.getSelectedIndex))) = Trim(Lcase(sSchema)) Then
								iCounter = iCounter + 1		
							End If
						End If

						'Verify Filter Rule Table
						If aProperties<>"" and aValues<>"" Then
										iCount = iCount + 1
										aColumns = Split(aProperties,":",-1,1)
										iRows=JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").GetROProperty("rows")
										iCols=JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").GetROProperty("cols")
										If iRows=0 Then
											iCounter = iCounter + 1
										End If
						End If
						
						If iCount=iCounter Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : All Given values match actual values")
							Fn_PLM_FilterRuleOperations_Ext = TRUE
							Set objClosureRule = nothing 
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : All Given values does not match actual values")
							Fn_PLM_FilterRuleOperations_Ext = FALSE
							Set objClosureRule = nothing 
							Exit Function
						End If

			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_FilterRuleOperations_Ext function failed")
						Fn_PLM_FilterRuleOperations_Ext = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_FilterRuleOperations_Ext")
	Fn_PLM_FilterRuleOperations_Ext = TRUE
	Set objClosureRule = nothing 
End Function



'''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''''/$$$$
'''''''/$$$$   FUNCTION NAME   : Fn_PLM_TransferOptionSetOperations_Ext(sAction, sTOSName, sDescription, bRemoteSite, sRemoteSiteID, sTransferMode,aProperties,aValues,sInfo1,sInfo2)
'''''''/$$$$
'''''''/$$$$   DESCRIPTION        :  This function is an extention of the Existing Function to en
'''''''/$$$$
'''''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
'''''''/$$$$										sTOSName : Valid TransferOption set name
'''''''/$$$$										sDescription : Valid TransferOption set Description
'''''''/$$$$										bRemoteSite : Valid Value For Remote Site
'''''''/$$$$										sRemoteSiteID : Valid ID For Remote Site
'''''''/$$$$										aProperties : Valid array of properties
'''''''/$$$$										aValues : Valid array of property values
'''''''/$$$$										sInfo1 : For Future Use
'''''''/$$$$										sInfo2 : For Future Use
'''''''/$$$$	
'''''''/$$$$		Return Value : 				True or False
'''''''/$$$$
'''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''''/$$$$
'''''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
'''''''/$$$$  
'''''''/$$$$    CREATED BY     :   SHREYAS          23/12/2011         1.0
'''''''/$$$$
'''''''/$$$$    REVIWED BY     :  Shreyas			23/12/2011            1.0
'''''''/$$$$
'''''''/$$$$		How To Use :    bReturn=Fn_PLM_TransferOptionSetOperations_Ext("Verify", "Test111", "Testing", "OFF", "", "PLMXMLAdminDataExport","Option:Display Name:Default Value:Description:Group Name:Read Only", "abc:def:True:hij:klm:1","","")
'''''''/$$$$										bReturn=Fn_PLM_PropertySetOperations_Ext("Create", "Test111", "Testing", "OFF", "", "PLMXMLAdminDataExport","Option:Display Name:Default Value:Description:Group Name:Read Only", "abc:def:True:hij:klm:1","","")
'''''''/$$$$							
'''''''/$$$$	
'''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$



Public Function Fn_PLM_TransferOptionSetOperations_Ext(sAction, sTOSName, sDescription, bRemoteSite, sRemoteSiteID, sTransferMode,aProperties,aValues,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_TransferOptionSetOperations_Ext"
	Dim objTOS, iReturn, aColumns, aColsData, iCols, iRows, iFlag, intCount, iCounter,bFlag
	Fn_PLM_TransferOptionSetOperations_Ext = False
	Set objTOS = Fn_UI_ObjectCreate("Fn_PLM_TransferOptionSetOperations_Ext", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Transfer Option Set
						If sTOSName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_TransferOptionSetOperations_Ext",objTOS,"TOSName",sTOSName)
						End If
						'Set Value for Description
						If sDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_TransferOptionSetOperations_Ext",objTOS,"Description",sDescription)
						End If
						'Set Remote Site Option true / false
						If bRemoteSite<>"" Then
							Call Fn_CheckBox_Set("Fn_PLM_TransferOptionSetOperations_Ext", objTOS, "RemoteSite", bRemoteSite)
						End If
						'Set Remote Site ID
						If sRemoteSiteID<>"" Then
							Call Fn_List_Select("Fn_PLM_TransferOptionSetOperations_Ext", objTOS, "TOSRemoteSiteID",sRemoteSiteID)
						End If
						'Set Transfer Mode
						If sTransferMode<>"" Then
							Call Fn_List_Select("Fn_PLM_TransferOptionSetOperations_Ext", objTOS, "TOSTransferMode",sTransferMode)
						End If
''						'add a table row if required
''
				If aProperties<>"" and aValues<>"" Then
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objTOS, "AddColumn")
									aColumns = Split(aProperties,":",-1,1)
									aColsData = Split(aValues,":",-1,1)
									iCols = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objTOS.JavaTable("ClosureRuleTable"), "cols")
									iRows = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objTOS.JavaTable("ClosureRuleTable"), "rows") - 1
									'iCols = objTOS.JavaTable("ClosureRuleTable").GetROProperty("cols")
									'iRows = objTOS.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
								For iCounter=0 to Ubound(aColumns)
									For intCount=0 to iCols-1
										If Trim(Lcase(objTOS.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
											iFlag = iFlag + 1
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
											'Logic for Setting Data
											If aColumns(iCounter)="Option" Then										
												'Select value from List
												JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
												wait 1,200
												Call Fn_KeyBoardOperation("SendKeys","{ENTER}")
											Elseif aColumns(iCounter)="Default Value" Then
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("TOSDefaultValue").Select aColsData(iCounter)
											Else
												'Set value in Cell.
												JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
											End If
											Exit For
										End If									
									Next
								Next
				End If

						If sAction="Create" Then
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations_Ext", objTOS, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations_Ext", objTOS, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations_Ext", objTOS, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations_Ext", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "AddRow","ModifyRow"
					iFlag = 0
					If sAction="AddRow" Then
						'Click on Add button.
						Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations_Ext", objTOS, "AddColumn")
					End If
							aColumns = Split(sTOSName,":",-1,1)
							aColsData = Split(sDescription,":",-1,1)
							iCols = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objTOS.JavaTable("ClosureRuleTable"), "cols")
							iRows = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objTOS.JavaTable("ClosureRuleTable"), "rows") - 1
							'iCols = objTOS.JavaTable("ClosureRuleTable").GetROProperty("cols")
							'iRows = objTOS.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objTOS.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										'Logic for Setting Data
										If aColumns(iCounter)="Default Value" Then
											objTOS.JavaTable("ClosureRuleTable").SetCellData iRows,"Default Value",aColsData(iCounter)
										ElseIf aColumns(iCounter) = "Read Only" Then
											If aColsData(iCounter) = "ON" Then
												Call Fn_UI_JavaTable_ClickCell("Fn_PLM_TransferOptionSetOperations_Ext",objTOS,"ClosureRuleTable",iRows, aColumns(iCounter))											
											End If											
										Else
											'Set value TableEditbox
											objTOS.JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
											Call Fn_Edit_Box("Fn_PLM_TransferOptionSetOperations_Ext",objTOS,"TableEditbox",aColsData(iCounter))
										End If
										Exit For
									End If									
								Next
							Next
							'Click on Modify button.
							Call Fn_Button_Click("Fn_PLM_TransferOptionSetOperations_Ext", objTOS, "Modify")
							If iFlag=Ubound(aColumns)+1 Then
								Fn_PLM_TransferOptionSetOperations_Ext = True
							Else
								Fn_PLM_TransferOptionSetOperations_Ext = False
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" of Fn_PLM_TransferOptionSetOperations_Ext failed")
								Set objTOS = nothing
								Exit Function
							End If

			Case "Verify"

						'Verify TOS Name
						If sTOSName<>"" Then
							intCount = intCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferOptionSetOperations_Ext",objTOS,"TOSName"))) = Trim(Lcase(sTOSName)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sTOSName+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sTOSName+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = FALSE
								Exit Function			
							End If
						End If
						'Verify TOS Description
						If sDescription<>"" Then
							intCount = intCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferOptionSetOperations_Ext",objTOS,"Description"))) = Trim(Lcase(sDescription)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sDescription+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sDescription+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = FALSE
								Exit Function			
							End If
						End If
						'Verify Remote Site Checkbox
						If bRemoteSite<>"" Then
							bFlag=False
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferOptionSetOperations_Ext",objTOS.JavaCheckBox("RemoteSite"), "value")) = 1 and Trim(Lcase(bRemoteSite)) = "on" Then
								bFlag=True
							ElseIf Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferOptionSetOperations_Ext",objTOS.JavaCheckBox("RemoteSite"), "value")) = 0 and Trim(Lcase(bRemoteSite)) = "off" Then
								bFlag=True
							End If						
							If bFlag=true Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+bRemoteSite+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+bRemoteSite+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = FALSE
								Exit Function		
							End If
						End If
						'Verify Remote Site ID
						If sRemoteSiteID<>"" Then

							If Trim(Lcase(objTOS.JavaList("TOSRemoteSiteID").GetItem(objTOS.JavaList("TOSRemoteSiteID").Object.getSelectedIndex))) = Trim(Lcase(sRemoteSiteID)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sRemoteSiteID+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sRemoteSiteID+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = FALSE
								Exit Function		
							End If
						End If
						'Verify Transfer Mode
						If sTransferMode<>"" Then

							If Trim(Lcase(objTOS.JavaList("TOSTransferMode").GetItem(objTOS.JavaList("TOSTransferMode").Object.getSelectedIndex))) = Trim(Lcase(sTransferMode)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sTransferMode+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sTransferMode+"]")
								Fn_PLM_TransferOptionSetOperations_Ext = FALSE
								Exit Function
							End If
						End If

					If aProperties<>"" and aValues<>"" Then
									aColumns = Split(aProperties,":",-1,1)
									aColsData = Split(aValues,":",-1,1)
									iCols = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objTOS.JavaTable("ClosureRuleTable"), "cols")
									iRows = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objTOS.JavaTable("ClosureRuleTable"), "rows") - 1
									'iCols = objTOS.JavaTable("ClosureRuleTable").GetROProperty("cols")
									'iRows = objTOS.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
								For iCounter=0 to Ubound(aColumns)
									For intCount=0 to iCols-1
										If Trim(Lcase(objTOS.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then

									sValue=objTOS.JavaTable("ClosureRuleTable").GetCellData(iRows,aColumns(iCounter))
										If Trim(lcase(sValue))=Trim(lcase(aColsData(iCounter))) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
												Fn_PLM_TransferOptionSetOperations_Ext = TRUE
										Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
													Fn_PLM_TransferOptionSetOperations_Ext = FALSE
													Exit Function			
										End If
											Exit For
										End If									
									Next
								Next
				End If

			Case "VerifyBlank"

						iCount = 0
						iCounter = 0

						'Verify TOS Name
						If sTOSName<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferOptionSetOperations_Ext",objTOS,"TOSName"))) = Trim(Lcase(sTOSName)) Then
								iCounter = iCounter + 1	
							End If
						End If

						'Verify TOS Description
						If sDescription<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_TransferOptionSetOperations_Ext",objTOS,"Description"))) = Trim(Lcase(sDescription)) Then
								iCounter = iCounter + 1	
							End If
						End If

						'Verify Remote Site Checkbox
						If bRemoteSite<>"" Then
							iCount = iCount + 1
							bFlag=False
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferOptionSetOperations_Ext",objTOS.JavaCheckBox("RemoteSite"), "value")) = 1 and Trim(Lcase(bRemoteSite)) = "on" Then
								bFlag=True
							ElseIf Cint(Fn_UI_Object_GetROProperty("Fn_PLM_TransferOptionSetOperations_Ext",objTOS.JavaCheckBox("RemoteSite"), "value")) = 0 and Trim(Lcase(bRemoteSite)) = "off" Then
								bFlag=True
							End If						
							If bFlag=true Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify Remote Site ID
						If sRemoteSiteID<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTOS.JavaList("TOSRemoteSiteID").GetItem(objTOS.JavaList("TOSRemoteSiteID").Object.getSelectedIndex))) = Trim(Lcase(sRemoteSiteID)) Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify Transfer Mode
						If sTransferMode<>"" Then
							iCount = iCount + 1
							If Trim(Lcase(objTOS.JavaList("TOSTransferMode").GetItem(objTOS.JavaList("TOSTransferMode").Object.getSelectedIndex))) = Trim(Lcase(sTransferMode)) Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify the TOS table
						If aProperties<>"" and aValues<>"" Then
										iCount = iCount + 1
										aColumns = Split(aProperties,":",-1,1)
										iRows=JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").GetROProperty("rows")
										iCols=JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").GetROProperty("cols")
										If iRows=0 Then
											iCounter = iCounter + 1
										End If
						End If
						
						If iCount=iCounter Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : All Given values match actual values")
							Fn_PLM_TransferOptionSetOperations_Ext = TRUE
							Set objTOS = nothing 
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : All Given values does not match actual values")
							Fn_PLM_TransferOptionSetOperations_Ext = FALSE
							Set objTOS = nothing 
							Exit Function
						End If

			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_TransferOptionSetOperations_Ext function failed")
						Fn_PLM_TransferOptionSetOperations_Ext = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_TransferOptionSetOperations_Ext")
	Fn_PLM_TransferOptionSetOperations_Ext = TRUE
	Set objTOS = nothing 
End Function


'''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''''/$$$$
'''''''/$$$$   FUNCTION NAME   :  Fn_PLM_ClosureRuleOperations_Ext(sAction, sRuleName, sDescription, sScope, sSchema,aProperties,aValues,sInfo1,sInfo2)
'''''''/$$$$
'''''''/$$$$   DESCRIPTION        :  This function is an extention of the Existing Function  Fn_PLM_ClosureRuleOperations
'''''''/$$$$
'''''''/$$$$  PARAMETERS   : 		sAction : Action to be performed
'''''''/$$$$										sRuleName : Valid Closure Rule Name 
'''''''/$$$$										sDescription : Valid Closure RuleDescription
'''''''/$$$$										sScope : Valid Scope
'''''''/$$$$										sSchema :  Valid Output Schema
'''''''/$$$$										aProperties : Valid array of properties
'''''''/$$$$										aValues : Valid array of property values
'''''''/$$$$										sInfo1 : For Future Use
'''''''/$$$$										sInfo2 : For Future Use
'''''''/$$$$	
'''''''/$$$$		Return Value : 				True or False
'''''''/$$$$
'''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''''/$$$$
'''''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
'''''''/$$$$  
'''''''/$$$$    CREATED BY     :   SHREYAS          27/12/2011         1.0
'''''''/$$$$
'''''''/$$$$    REVIWED BY     :  Shreyas			27/12/2011            1.0
'''''''/$$$$
'''''''/$$$$		How To Use :    bReturn=Fn_PLM_ClosureRuleOperations_Ext("Create", "Shreyas_Waichal", "Testing", "Export", "PLMXML","Primary Object Class Type:Primary Object:Secondary Object Class Type:Secondary Object:Relation Type:Related Property Or Object:Action Type","CLASS:Item:CLASS:Dataset:PROPERTY:IMNA_specifications:PROCESS+TRAVERSE","","")
'''''''/$$$$										bReturn=Fn_PLM_ClosureRuleOperations_Ext("VerifyDetails", "Shreyas_Waichal", "Testing", "Export", "PLMXML",i,y,"","")
'''''''/$$$$							
'''''''/$$$$	
'''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Public Function Fn_PLM_ClosureRuleOperations_Ext(sAction, sRuleName, sDescription, sScope, sSchema,aProperties,aValues,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_ClosureRuleOperations_Ext"
	Dim objClosureRule, iReturn, aColumns, aColsData, iCols, iRows, iRowCnt, iFlag, intCount, iCount, iCounter,sX,sY,sX1,sY1,i,bFlag,DeviceReplay,sTemplateType,intNoOfObjects
	Fn_PLM_ClosureRuleOperations_Ext = False
	Set objClosureRule = Fn_UI_ObjectCreate("Fn_PLM_ClosureRuleOperations_Ext", JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet"))
		Select Case sAction
			Case "Create","Modify"
						'Set Value for Rule Name
						If sRuleName<>"" Then
							Call Fn_Edit_Box("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule,"TraversalRuleName",sRuleName)
						End If
						'Set Value for Description
						If sDescription<>"" Then
							Call Fn_Edit_Box("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule,"Description",sDescription)
						End If
						'Set the Scope of Traversal.
						If sScope<>"" Then
							Call Fn_UI_JavaRadioButton_SetON("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule, sScope)
						End If
						'Set Output Schema format
						If sSchema<>"" Then
							iReturn = objClosureRule.JavaList("SchemaFormat").GetItemIndex(sSchema)
							objClosureRule.JavaList("SchemaFormat").Object.setSelectedIndex iReturn
						End If
						If aProperties<>"" and aValues<>""  Then
							Call Fn_Button_Click("Fn_PLM_PropertySetOperations", objClosureRule, "AddColumn")
								aColumns = Split(aProperties,":",-1,1)
								aColsData = Split(aValues,":",-1,1)
								iCols = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaTable("ClosureRuleTable"), "cols")
								iRows = Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaTable("ClosureRuleTable"), "rows") - 1
								'iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
								'iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
										'Logic for Setting Data
										If aColumns(iCounter)="Primary Object Class Type" OR aColumns(iCounter)="Relation Type"Then										
											'Select value from List
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
										Elseif aColumns(iCounter)="Action Type" then
											Set DeviceReplay = CreateObject("Mercury.DeviceReplay")
											Const VK_CONTROL = 29
										If instr(1,aColsData(iCounter),"+")>0 Then
													objClosureRule.JavaTable("ClosureRuleTable").Object.editCellAt iRows,6
													aValues=split(aColsData(iCounter),"+",-1,1)
													
													If aValues(0)="PROCESS" AND aValues(1)="TRAVERSE" Then
													   aValues(0) = "PROCESS+TRAVERSE"
													   aValues(1)=" "
													End If
													
													objClosureRule.JavaButton("multipledropdown_16").Click micLeftBtn
													wait 2
													Set sTemplateType=Description.Create()
													sTemplateType("Class Name").value = "JavaStaticText"
													Set  intNoOfObjects = objClosureRule.ChildObjects(sTemplateType)
														For i = 0 to intNoOfObjects.count-1
																		   If  intNoOfObjects(i).getROProperty("label") = aValues(0) Then
																				wait(1)
																				sX=	intNoOfObjects(i).GetROProperty ("abs_x")
																				sY=intNoOfObjects(i).GetROProperty ("abs_y")
																				DeviceReplay.MouseClick sX,sY,LEFT_MOUSE_BUTTON
																			End if
													
																			If  intNoOfObjects(i).getROProperty("label") = aValues(1) Then
																					wait(1)
																					DeviceReplay.KeyDown VK_CONTROL
																					objClosureRule.JavaTable("ClosureRuleTable").Object.editCellAt iRows,6
																					objClosureRule.JavaButton("multipledropdown_16").Click micLeftBtn
																					sX1=	intNoOfObjects(i).GetROProperty ("abs_x")
																					sY1=intNoOfObjects(i).GetROProperty ("abs_y")
																					DeviceReplay.MouseClick sX1,sY1,LEFT_MOUSE_BUTTON
																					DeviceReplay.KeyUp VK_CONTROL
																					Exit for
																		   End If
														Next
										Else
																	objClosureRule.JavaTable("ClosureRuleTable").Object.editCellAt iRows,6
																	objClosureRule.JavaButton("multipledropdown_16").Click micLeftBtn
																	wait 2
																	Set sTemplateType=Description.Create()
																	sTemplateType("Class Name").value = "JavaStaticText"
																	Set  intNoOfObjects = objClosureRule.ChildObjects(sTemplateType)
																	 For i = 0 to intNoOfObjects.count-1
																		   If  intNoOfObjects(i).getROProperty("label") = aColsData(iCounter) Then
																				sX=	intNoOfObjects(i).GetROProperty ("abs_x")
																				sY=intNoOfObjects(i).GetROProperty ("abs_y")
																				DeviceReplay.MouseClick sX,sY,LEFT_MOUSE_BUTTON
																				Exit for
																			End if		
																	Next 
										End If
										Else
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").ClickCell iRows,aColumns(iCounter)
											'Set value in Cell.
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
										End If
										Exit For
									End If									
								Next
							Next

					End If
							objClosureRule.Click 0,0,"LEFT"
						If sAction="Create" Then
							wait 2
							'Click on Add button
							Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations_Ext", objClosureRule, "Create")
						ElseIf sAction="Modify" Then
							'Click on Modify button
							Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations_Ext", objClosureRule, "Modify")
						End If
			Case "Delete"
						'Click on Delete button
						Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations_Ext", objClosureRule, "Delete")
						If JavaDialog("Delete Confirmation").Exist Then
							'Click on yes button
							Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations_Ext", JavaDialog("Delete Confirmation"), "Yes")
						End If
			Case "AddRow","ModifyRow"
					iFlag = 0
					If sAction="AddRow" Then
						'Click on Add button.
						Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations_Ext", objClosureRule, "AddColumn")
					End If
							aColumns = Split(aProperties,":",-1,1)
							aColsData = Split(aValues,":",-1,1)
							iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
							iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
							For iCounter=0 to Ubound(aColumns)
								For intCount=0 to iCols-1
									If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then
										iFlag = iFlag + 1
										JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").DoubleClickCell iRows,aColumns(iCounter),"LEFT","NONE"
										'Logic for Setting Data
										If aColumns(iCounter)="Primary Object Class Type" OR aColumns(iCounter)="Secondary Object Class Type" OR aColumns(iCounter)="Relation Type" Then										
											'Select value from List
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaList("ClosureTableList").Select aColsData(iCounter)
										ElseIf aColumns(iCounter)="Action Type" Then
'											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").Set ""
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaEdit("ClosureTableEdit").Set aColsData(iCounter)
										Else
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),""
											'Set value in Cell.
											JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").SetCellData iRows,aColumns(iCounter),aColsData(iCounter)
										End If
										Exit For
									End If									
								Next
							Next
							'Click on Modify button.
							Call Fn_Button_Click("Fn_PLM_ClosureRuleOperations_Ext", objClosureRule, "Modify")
							If iFlag=Ubound(aColumns)+1 Then
								Fn_PLM_ClosureRuleOperations_Ext = True
							Else
								Fn_PLM_ClosureRuleOperations_Ext = False
							End If

	Case "VerifyDetails"

						'Verify CR Name
						If sFRName<>"" Then

							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule,"TraversalRuleName"))) = Trim(Lcase(sRuleName)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sRuleName+"]")
							Fn_PLM_ClosureRuleOperations_Ext = TRUE
		    					Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sRuleName+"]")
								Fn_PLM_ClosureRuleOperations_Ext = FALSE
								Exit Function			
						End If
					End If

						'Verify CR Description
						If sDescription<>"" Then

							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule,"Description"))) = Trim(Lcase(sDescription)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sDescription+"]")
								Fn_PLM_ClosureRuleOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sDescription+"]")
								Fn_PLM_ClosureRuleOperations_Ext = FALSE
								Exit Function			
							End If
						End If

						'Verify Scope
						If sScope<>"" Then
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaRadioButton(sScope), "value")) = 1 Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sScope+"]")
								Fn_PLM_ClosureRuleOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sScope+"]")
								Fn_PLM_ClosureRuleOperations_Ext = FALSE
								Exit Function			
							End If
						End If

						'Verify Output Scheme Format
						If sSchema<>"" Then

							If Trim(Lcase(objClosureRule.JavaList("SchemaFormat").GetItem(objClosureRule.JavaList("SchemaFormat").Object.getSelectedIndex))) = Trim(Lcase(sSchema)) Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+sSchema+"]")
								Fn_PLM_ClosureRuleOperations_Ext = TRUE
		    				Else
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+sSchema+"]")
								Fn_PLM_ClosureRuleOperations_Ext = FALSE
								Exit Function			
							End If
						End If
					
					If aProperties<>"" and aValues<>"" Then
									aColumns = Split(aProperties,":",-1,1)
									aColsData = Split(aValues,":",-1,1)
									iCols = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("cols")
									iRows = objClosureRule.JavaTable("ClosureRuleTable").GetROProperty("rows") - 1
								For iCounter=0 to Ubound(aColumns)
									For intCount=0 to iCols-1
										If Trim(Lcase(objClosureRule.JavaTable("ClosureRuleTable").GetColumnName(intCount)))= Trim(Lcase(aColumns(iCounter))) Then

									sValue=objClosureRule.JavaTable("ClosureRuleTable").GetCellData(iRows,aColumns(iCounter))
										If Trim(lcase(sValue))=Trim(lcase(aColsData(iCounter))) Then
												Call Fn_WriteLogFile(Environment.Value("TestLogFile"),  "Sucessfully Verified the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
												Fn_PLM_ClosureRuleOperations_Ext = TRUE
										Else
													Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Failed to Verify the Value ["+aColsData(iCounter)+"] from Column ["+aColumns(iCounter)+"]")
													Fn_PLM_ClosureRuleOperations_Ext = FALSE
													Exit Function			
										End If
											Exit For
										End If									
									Next
								Next
				End If

		Case "VerifyBlank"
						iCount = 0
						iCounter = 0
						'Verify CR Name
						If sRuleName<>"" Then
							iCount=iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule,"TraversalRuleName"))) = Trim(Lcase(sRuleName)) Then
								iCounter = iCounter + 1	
							End If
						End If

						'Verify CR Description
						If sDescription<>"" Then
							iCount=iCount + 1
							If Trim(Lcase(Fn_Edit_Box_GetValue("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule,"Description"))) = Trim(Lcase(sDescription)) Then
									iCounter = iCounter + 1			
							End If
						End If

						'Verify Scope
						If sScope<>"" Then
							iCount=iCount + 1
							If Cint(Fn_UI_Object_GetROProperty("Fn_PLM_ClosureRuleOperations_Ext",objClosureRule.JavaRadioButton(sScope), "value")) = 1 Then
								iCounter = iCounter + 1
							End If
						End If

						'Verify Output Scheme Format
						If sSchema<>"" Then
							iCount=iCount + 1
							If Trim(Lcase(objClosureRule.JavaList("SchemaFormat").GetItem(objClosureRule.JavaList("SchemaFormat").Object.getSelectedIndex))) = Trim(Lcase(sSchema)) Then
								iCounter = iCounter + 1	
							End If
						End If

						'Verify Closure Rule Table
						If aProperties<>"" and aValues<>"" Then
										iCount = iCount + 1
										aColumns = Split(aProperties,":",-1,1)
										iRows=JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").GetROProperty("rows")
										iCols=JavaWindow("PLMXML-TeamCenter").JavaWindow("JApplet").JavaTable("ClosureRuleTable").GetROProperty("cols")
                                                  If iRows=0 Then
											iCounter = iCounter + 1
										End If
						End If
						
						If iCount=iCounter  Then
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Pass : All Given values match actual values")
							Fn_PLM_ClosureRuleOperations_Ext = TRUE
							Set objClosureRule = nothing 
							Exit Function
						Else
							Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fail : All Given values does not match actual values")
							Fn_PLM_ClosureRuleOperations_Ext = FALSE
							Set objClosureRule = nothing 
							Exit Function
						End If

			Case Else						
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Fn_PLM_ClosureRuleOperations_Ext function failed")
						Fn_PLM_ClosureRuleOperations_Ext = FALSE
						Exit Function											
		End Select
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), sAction &" Sucessfully completed of the function Fn_PLM_ClosureRuleOperations_Ext")
	Fn_PLM_ClosureRuleOperations_Ext = TRUE
	Set objClosureRule = nothing 
End Function
'''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''''/$$$$
'''''''/$$$$   	FUNCTION NAME   	 :  Fn_PLM_RenameItemsFromXMLFile(sFolder,sFilePath,sNewFileName,aReplaceVal,sInfo1,sInfo2)
'''''''/$$$$
'''''''/$$$$   	DESCRIPTION        	 :  This function Will replace specified components in an XML file
'''''''/$$$$
'''''''/$$$$  	PARAMETERS   		 : 	sFolder : Valid Path of the Folder
'''''''/$$$$							sFileName : Valid  Name of the XML File
'''''''/$$$$							sNewFileName : New Name of the XML File (Should be same as the Folder name containing .prt files)
'''''''/$$$$							aReplaceVal : Array of values to be replaced
'''''''/$$$$							sInfo1 : For Future Use
'''''''/$$$$							sInfo2 : For Future Use
'''''''/$$$$	
'''''''/$$$$	Return Value 		 :	 True or False
'''''''/$$$$
'''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''''/$$$$
'''''''/$$$$	How To Use 			 :    bReturn=Fn_PLM_RenameItemsFromXMLFile( "C:\Documents and Settings\Administrator\Desktop\Valve_KB","KBvalve_assembly_A_1-KBvalve_assembly_Original.xml","KBvalve_assembly_A_1-KBvalve_assembly.xml","_12345","","")
'''''''/$$$$						 :    bReturn=Fn_PLM_RenameItemsFromXMLFile( "C:\Documents and Settings\Administrator\Desktop\Valve_KB","KBvalve_assembly_A_1-KBvalve_assembly_Original.xml","KBvalve_assembly_A_1-KBvalve_assembly.xml","_123~No","","") - "No" to remove '_' from Id  
'''''''/$$$$  
'''''''/$$$$	HISTORY           	 :
'''''''/$$$$  
'				Developer Name			Date				Rev. No.			Changes Done									Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Shreyas					17-Jan-2012			1.0					Created											Shreyas
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'				Koustubh Watwe			21-Jan-2012			1.0					Added code to remove "_" while replacing Ids   	Rupali
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public function Fn_PLM_RenameItemsFromXMLFile(sFolderPath,sFileName,sNewFileName,aReplaceVal,sInfo1,sInfo2)

GBL_FAILED_FUNCTION_NAME="Fn_PLM_RenameItemsFromXMLFile"
Fn_PLM_RenameItemsFromXMLFile=false
Dim objFSO,objFile,xmlDoc,iRanNo,iCount
Dim objContent,bFlag,sContent,sValue, sReplaceStr,sFolderName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile( sFolderPath+"\"+sFileName)

bFlag=False
'Set xmlDoc =  CreateObject("Microsoft.XMLDOM")
'xmlDoc.Load(sFileName)
'Generate a random Number
iRanNo=Fn_RandNoGenerate()
'
''Create a duplicate Copy of the original file
'xmlDoc.Save(sFolderPath+"\"+sNewFileName)
objFSO.CopyFile sFolderPath+"\"+sFileName,sFolderPath+"\"+sNewFileName
'
'wait 3
		'check if the File is Created
		If not objFSO.FileExists(sFolderPath+"\"+sNewFileName) then
			Fn_PLM_RenameItemsFromXMLFile=false
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The file  [" + sNewFileName + "] was not created")
			Exit function
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The file  [" + sNewFileName + "] was Successfully created at location ["+sFolderPath+"\"+sNewFileName+"]")
		End if

'now replace the values
'		   Set objFile = objFSO.GetFile(sFolderPath &"\" & sNewFileName)
'			For iCount=0 to uBound(aReplaceVal)
'				   Set objContent = objFile.OpenAsTextStream(1,-2)
'					sContent = objContent.ReadAll
'					sValue=replace(sContent,aReplaceVal(iCount),aReplaceVal(iCount)+"_"+cstr(iRanNo))
'					Set objContent = objFile.OpenAsTextStream(2,-2)
'					objContent.Write(sValue)
'					bFlag=True
'			Next

'now replace the values
					Set objFile = objFSO.GetFile(sFolderPath &"\" & sNewFileName)
					Set objContent = objFile.OpenAsTextStream(1,-2)
					sContent = objContent.ReadAll
					Set objContent =nothing
					' code added by Koustubh
					If IsArray(aReplaceVal) = False Then
						aReplaceVal = split(aReplaceVal,"~")
						sReplaceStr = "_"+cstr(iRanNo)
						If uBound(aReplaceVal) = 1 then
							If lcase(aReplaceVal(1)) = "no" Then
								sReplaceStr = cstr(iRanNo)
							End IF 
						End IF
						sValue=replace(sContent,aReplaceVal(0), sReplaceStr)
					End If
					'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
					Set objContent = objFile.OpenAsTextStream(2,-2)
					objContent.Write(sValue)
					bFlag=True

			If bFlag=true Then
				Fn_PLM_RenameItemsFromXMLFile=iRanNo
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The function  [ Fn_PLM_RenameItemsFromXMLFile ] Completed Successsfully")
			Else
				Fn_PLM_RenameItemsFromXMLFile=false
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The function  [ Fn_PLM_RenameItemsFromXMLFile ] Failed")
				Exit function	
			End If

'Delete the .svn folder from the Folder to enable successful import
sFolderName=split(sNewFileName,".",-1,1)
If objFSO.FolderExists (sFolderPath+"\"+sFolderName(0)) then
	If objFSO.FolderExists(sFolderPath+"\"+sFolderName(0)+"\.svn") Then
		objFSO.DeleteFolder sFolderPath+"\"+sFolderName(0)+"\.svn",True
		Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deleted the [.svn] folder")
	End If
Else
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Foler Containing Model Files Does Not Exist")
End IF 
	
Set objFSO = nothing
Set objFile = nothing
Set xmlDoc =  nothing
Set objContent =nothing
End Function




'''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'''''''/$$$$
'''''''/$$$$   FUNCTION NAME   :  Fn_PLM_VerifyValueInTag(sXMLPath,sTag,sTagValue,sInfo1,sInfo2)
'''''''/$$$$
'''''''/$$$$   DESCRIPTION        :  This function Will  verify a value under the specified Tag
'''''''/$$$$
'''''''/$$$$  PARAMETERS   : 		sXMLPath : Valid Path of the XML File
'''''''/$$$$										sTag : Valid  Name of the XML tag
'''''''/$$$$										sTagValue : Valid Value to be verified under a tag
'''''''/$$$$										sInfo1 : For Future Use
'''''''/$$$$										sInfo2 : For Future Use
'''''''/$$$$	
'''''''/$$$$		Return Value : 				True or False
'''''''/$$$$
'''''''/$$$$    Function Calls       :   Fn_WriteLogFile()
'''''''/$$$$
'''''''/$$$$		HISTORY           :   		AUTHOR                 DATE        VERSION
'''''''/$$$$  
'''''''/$$$$    CREATED BY     :   SHREYAS          23/01/2012         1.0
'''''''/$$$$
'''''''/$$$$    REVIWED BY     :  Shreyas			23/01/2012           1.0
'''''''/$$$$
'''''''/$$$$		How To Use :    bReturn=Fn_PLM_VerifyValueInTag("C:\Documents and Settings\Administrator\Desktop\sonal\000077-new.xml","UserData","object_desc","","")
'''''''/$$$$
'''''''/$$$$	
'''''''$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Function Fn_PLM_VerifyValueInTag(sXMLPath,sTag,sTagValue,sInfo1,sInfo2)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_VerifyValueInTag"
   On error resume next
Dim xmlObj,sCount,i,j,bFlag,aValues
Fn_PLM_VerifyValueInTag=false
Set xmlObj = XMLUtil.CreateXML() 

xmlObj.LoadFile(sXMLPath) 
Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Loaded the file  ["+sXMLPath+"]" )
wait 5
sCount= xmlObj.GetRootElement().ChildElements().Count
j=0
bFlag=False
For iCounter=1 to sCount
sValue= xmlObj.GetRootElement().ChildElements().Item(iCounter)
If instr(1,sValue,sTag)>0 then
	aValues=split(sValue,vbnewline,-1,1)
For i=0 to ubound(aValues) 
			j=i+1
			If instr(1,aValues(j+1),sTagValue)>0 Then
							If aValues(j+1)="" Then
								Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Function  [Fn_PLM_VerifyValueInTag] Failed Because The Specified Value ["+sTagValue+"] was not Found")
								bFlag=False
								Exit for
							End If
				bFlag=True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Value ["+sTagValue+"] exists under the Tag ["+sTag+"]" )
				Exit for
			End If
Next
End If

If bFlag=True then 
	Fn_PLM_VerifyValueInTag=true
	Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Function  [Fn_PLM_VerifyValueInTag] PASSED.")
	Exit for
End If
Next

Set xmlObj =nothing
err.clear
End function

'*********************************************************	 Function for Delete .svn folder  ***********************************************************************
'Function Name			:		Fn_PLM_Delete_SVN(sFolder)

'Description			    :	 	    Function for Delete .svn folder

'Parameters			   :                   1.sFolder : Folder path that contain .svn

'Return Value		   	   : 		   True / False

'Pre-requisite			    :		     Folder should be available to remve [ .svn ] folder

'Examples				    :		   Fn_PLM_Delete_SVN("d:\test")
'Examples				    :		   Fn_PLM_Delete_SVN("d:\test~s") -  attach ~s to delete .svn from sub folders

'History:
'	Developer Name			Date		   Rev. No.			Changes Done			Reviewer	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Naveen		        23-Feb-2012			1.0			   
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Koustubh			01-Aug-2012			1.0			   Added code to delete .SVN folder from sub folders.
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fn_PLM_Delete_SVN(sFolder)
	GBL_FAILED_FUNCTION_NAME="Fn_PLM_Delete_SVN"
	Dim aFolderName, sSubFolders, f, f1
	Fn_PLM_Delete_SVN = false
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	IF instr(sFolder,"~") > 0 Then
		aFolderName = split(sFolder,"~")
		If objFSO.FolderExists (aFolderName(0)) then
			Set f = objFSO.GetFolder(aFolderName(0))
			Set sSubFolders = f.SubFolders
			For Each f1 in sSubFolders
				If lcase(f1.name) <> ".svn" Then
					Fn_PLM_Delete_SVN = Fn_PLM_Delete_SVN(aFolderName(0) & "\" & f1.name & "~s")
				End IF
			Next
			objFSO.DeleteFolder aFolderName(0) & "\.svn",True
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deleted the [.svn] folder under [ " & aFolderName(0) & " ]")
			Fn_PLM_Delete_SVN = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Foler Containing .svn folder ["& aFolderName(0) &"] Does Not Exist")
			Fn_PLM_Delete_SVN = False
		End IF
	Else
		If objFSO.FolderExists (sFolder) then
			If objFSO.FolderExists(sFolder+"\.svn") Then
				objFSO.DeleteFolder sFolder+"\.svn",True
				Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "Successfully Deleted the [.svn] folder")
			End If
			Fn_PLM_Delete_SVN = True
		Else
			Call Fn_WriteLogFile(Environment.Value("TestLogFile"), "The Foler Containing .svn folder ["&sFolder&"] Does Not Exist")
			Fn_PLM_Delete_SVN = False
		End IF
	End IF
	Set objFSO = nothing
End Function
